# Contributing

Thanks for your interest in improving the RMA Management System!

## Ground rules
- The project is intentionally a **single-file** app (`rma_app.py`). Please keep it that way.
- Never commit real company data, credentials, serial numbers or customer information.
- Keep the SQL Server-specific rules in mind: `TOP N` (not `LIMIT`), `LEFT()` (not `substr()`), wrap reserved words such as `[status]` in brackets, and repeat expressions in `GROUP BY` instead of using aliases.

## How to contribute
1. Fork the repo and create a branch: `git checkout -b feature/my-change`
2. Make your change and test it against a SQL Server instance (the sample data in `sql/` is handy).
3. Commit with a clear message (e.g. `Fix: TAT calculation for cancelled RMAs`).
4. Open a pull request describing what changed and why. Screenshots help for UI changes.

## Reporting bugs / requesting features
Use the issue templates. For bugs include: OS, Python version, ODBC driver version, SQL Server version and the relevant part of `rma_app.log` (remove any sensitive data first).
