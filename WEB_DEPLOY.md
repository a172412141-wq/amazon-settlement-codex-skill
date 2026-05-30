# Web tool deployment

This repository includes a password-protected Flask web UI.

## Local run

```bash
python -m pip install -r web/requirements.txt
APP_PASSWORD='Fang123' SECRET_KEY='replace-with-a-long-random-string' python web/app.py
```

Open:

```text
http://localhost:8000
```

## Production environment variables

Set these on the hosting platform:

```text
APP_PASSWORD=Fang123
SECRET_KEY=<use a long random value>
MAX_UPLOAD_MB=80
```

Do not hard-code production secrets in Git.

## Render deployment

1. Create a new Web Service on Render.
2. Connect this private GitHub repository.
3. Build command:

```bash
pip install -r web/requirements.txt
```

4. Start command:

```bash
gunicorn web.app:app --bind 0.0.0.0:$PORT --workers 1 --timeout 180
```

5. Add the environment variables listed above.
6. Deploy.

## Railway deployment

1. Create a new Railway project from this GitHub repository.
2. Add the same environment variables.
3. Use the `Procfile` start command automatically, or set start command manually:

```bash
gunicorn web.app:app --bind 0.0.0.0:$PORT --workers 1 --timeout 180
```

## Data handling

- Uploaded PDF files are saved only to a temporary request directory.
- The generated XLSX is streamed back as a download.
- Temporary files are deleted after the response.
- No database is used.
- Do not add persistent upload folders unless you intentionally want file storage.

## Password note

The requested password is `Fang123`. It is implemented as the default fallback and should also be set through `APP_PASSWORD` on the hosting platform.

For real company use, replace it with a longer password before sharing outside a small internal group.
