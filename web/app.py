#!/usr/bin/env python3
"""Password-protected web UI for Amazon settlement PDF processing.

Data handling:
- Uploaded PDFs are written only to a temporary directory for the current request.
- Generated Excel output is returned directly as a download.
- Temporary files are deleted after the response is sent.
- No database is used.
"""
from __future__ import annotations

import os
import shutil
import sys
import tempfile
from decimal import Decimal
from pathlib import Path
from typing import List

from flask import Flask, after_this_request, redirect, render_template_string, request, send_file, session, url_for
from werkzeug.utils import secure_filename

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / ".agents" / "skills" / "amazon-settlement-xlsx" / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

from pdf_to_table import default_mapping_path, extract_all, write_xlsx  # noqa: E402
from run import parse_manual_rates  # noqa: E402

APP_PASSWORD = os.environ.get("APP_PASSWORD", "Fang123")
SECRET_KEY = os.environ.get("SECRET_KEY", "change-this-secret-key-before-public-use")
MAX_UPLOAD_MB = int(os.environ.get("MAX_UPLOAD_MB", "80"))

app = Flask(__name__)
app.secret_key = SECRET_KEY
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024

LOGIN_HTML = """
<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Amazon Settlement 数据表工具</title>
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; background:#f6f7fb; margin:0; color:#111827; }
    .box { max-width:420px; margin:12vh auto; background:white; border-radius:16px; padding:32px; box-shadow:0 12px 40px rgba(15,23,42,.08); }
    h1 { font-size:22px; margin:0 0 8px; }
    p { color:#6b7280; line-height:1.6; }
    input { width:100%; box-sizing:border-box; padding:12px 14px; border:1px solid #d1d5db; border-radius:10px; font-size:16px; }
    button { width:100%; margin-top:16px; padding:12px 14px; border:0; border-radius:10px; background:#2563eb; color:white; font-size:16px; cursor:pointer; }
    .error { color:#b91c1c; background:#fee2e2; padding:10px 12px; border-radius:10px; }
  </style>
</head>
<body>
  <div class="box">
    <h1>Amazon Settlement 数据表工具</h1>
    <p>输入密码后上传 PDF，系统会生成 Excel 数据表。</p>
    {% if error %}<div class="error">{{ error }}</div>{% endif %}
    <form method="post">
      <input type="password" name="password" placeholder="请输入密码" autocomplete="current-password" autofocus>
      <button type="submit">进入工具</button>
    </form>
  </div>
</body>
</html>
"""

INDEX_HTML = """
<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Amazon Settlement 数据表工具</title>
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; background:#f6f7fb; margin:0; color:#111827; }
    .wrap { max-width:820px; margin:48px auto; padding:0 20px; }
    .card { background:white; border-radius:18px; padding:28px; box-shadow:0 12px 40px rgba(15,23,42,.08); }
    h1 { font-size:26px; margin:0 0 8px; }
    p { color:#6b7280; line-height:1.7; }
    label { display:block; margin-top:18px; font-weight:600; }
    input[type=file], input[type=text] { margin-top:8px; width:100%; box-sizing:border-box; padding:12px; border:1px solid #d1d5db; border-radius:10px; background:white; }
    button { margin-top:20px; padding:12px 18px; border:0; border-radius:10px; background:#2563eb; color:white; font-size:16px; cursor:pointer; }
    .secondary { background:#6b7280; text-decoration:none; display:inline-block; margin-left:8px; padding:12px 18px; border-radius:10px; color:white; }
    .notice { background:#eff6ff; border:1px solid #bfdbfe; color:#1e3a8a; padding:12px 14px; border-radius:12px; margin:18px 0; }
    .error { background:#fee2e2; border:1px solid #fecaca; color:#991b1b; padding:12px 14px; border-radius:12px; margin:18px 0; white-space:pre-wrap; }
    .small { font-size:13px; color:#6b7280; }
    ul { color:#374151; line-height:1.7; }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="card">
      <h1>Amazon Settlement PDF → Excel 数据表</h1>
      <p>上传一个或多个 Amazon Settlement PDF，系统会自动提取 income、expenses、transfer、tax、币种、汇率，并生成人民币金额。</p>
      <div class="notice">文件只在本次请求的临时目录中处理，生成下载后自动删除；网站不使用数据库保存业务文件。</div>
      {% if error %}<div class="error">{{ error }}</div>{% endif %}
      <form method="post" action="{{ url_for('process') }}" enctype="multipart/form-data">
        <label>PDF 文件</label>
        <input type="file" name="pdfs" accept="application/pdf,.pdf" multiple required>
        <p class="small">可一次上传多个 PDF。文件名建议保留站点和月份，例如 CRZX-DE_Standard-2026-04.pdf。</p>

        <label>手动汇率，可选</label>
        <input type="text" name="manual_rates" placeholder="例如：EUR=7.916 或 EUR:2026-04-01=7.916；多个用逗号分隔">
        <p class="small">默认自动取 ChinaMoney 人民币汇率中间价。取不到时可用手动汇率。</p>

        <button type="submit">生成 Excel 数据表</button>
        <a class="secondary" href="{{ url_for('logout') }}">退出</a>
      </form>

      <h3>输出字段</h3>
      <ul>
        <li>原币：销售额、销售税、销售额含税、销售费用、账单回款额</li>
        <li>人民币：销售额、销售税、销售额含税、销售费用、账单回款额</li>
        <li>复核：PDF 原始 expenses、PDF 原始 transfer、状态、备注</li>
      </ul>
    </div>
  </div>
</body>
</html>
"""


def is_allowed_pdf(filename: str) -> bool:
    return filename.lower().endswith(".pdf")


def login_required() -> bool:
    return bool(session.get("authenticated"))


@app.get("/")
def index():
    if not login_required():
        return redirect(url_for("login"))
    return render_template_string(INDEX_HTML)


@app.route("/login", methods=["GET", "POST"])
def login():
    error = ""
    if request.method == "POST":
        password = request.form.get("password", "")
        if password == APP_PASSWORD:
            session["authenticated"] = True
            return redirect(url_for("index"))
        error = "密码错误。"
    return render_template_string(LOGIN_HTML, error=error)


@app.get("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/process")
def process():
    if not login_required():
        return redirect(url_for("login"))

    files = request.files.getlist("pdfs")
    if not files:
        return render_template_string(INDEX_HTML, error="请上传至少一个 PDF 文件。")

    tmp_dir = Path(tempfile.mkdtemp(prefix="amazon_settlement_web_"))
    try:
        pdf_dir = tmp_dir / "pdfs"
        pdf_dir.mkdir(parents=True, exist_ok=True)
        pdf_paths: List[Path] = []
        for file in files:
            original = file.filename or ""
            if not original or not is_allowed_pdf(original):
                raise ValueError(f"只支持 PDF 文件：{original}")
            safe_name = secure_filename(original) or "upload.pdf"
            target = pdf_dir / safe_name
            # Avoid overwriting files with identical names in the same upload batch.
            counter = 1
            while target.exists():
                target = pdf_dir / f"{target.stem}_{counter}{target.suffix}"
                counter += 1
            file.save(target)
            pdf_paths.append(target)

        manual_input = request.form.get("manual_rates", "").strip()
        manual_items = [item.strip() for item in manual_input.replace("；", ",").split(",") if item.strip()]
        manual_rates = parse_manual_rates(manual_items)

        output_path = tmp_dir / "amazon_settlement_table.xlsx"
        rate_cache = tmp_dir / "chinamoney_rates.csv"
        rows = extract_all(
            pdf_paths,
            default_mapping_path(None),
            rate_cache,
            manual_rates,
            offline=False,
            strict=False,
        )
        write_xlsx(rows, output_path)

        @after_this_request
        def cleanup(response):
            try:
                shutil.rmtree(tmp_dir, ignore_errors=True)
            finally:
                return response

        return send_file(
            output_path,
            as_attachment=True,
            download_name="amazon_settlement_table.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as exc:
        shutil.rmtree(tmp_dir, ignore_errors=True)
        return render_template_string(INDEX_HTML, error=str(exc)), 400


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8000")))
