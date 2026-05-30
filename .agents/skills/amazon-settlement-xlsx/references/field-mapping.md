# Field Mapping Notes

## PDF fields

Use the top summary table first. Common labels:

- income: `Income`, `Einnahmen`, `Ingresos`, `Entrate`, `Revenus`, `Recettes`, `Inkomsten`, `Przychody`, `売上`, `収入`
- expenses: `Expenses`, `Ausgaben`, `Gastos`, `Spese`, `Dépenses`, `Kosten`, `Wydatki`, `費用`
- tax: `Tax`, `Steuer`, `Impuesto`, `Imposte`, `Taxes`, `Belasting`, `Podatek`, `税`
- transfer: `Transfers`, `Transfer to bank account`, `Übertragungen`, `Überweisungen auf Bankkonto`, `Transferencias`, `Trasferimenti`, `Virements`, `Overboekingen`, `Przelewy`, `振込`

## Workbook cells in the provided sample

- G = income in original currency
- H = tax in original currency
- I = formula, income including tax
- J = expenses as a positive cost
- K = bank transfer amount as a positive receipt
- L = currency
- M = CNY exchange rate

## Sign handling

Amazon reports normally show transfers and expenses as negative. The sample workbook formula subtracts sales expenses from sales amounts, so the sales expense column expects a positive cost. The helper script writes absolute values for expenses and transfers, while keeping the raw signed values in the audit CSV.
