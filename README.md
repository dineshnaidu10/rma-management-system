# RMA Management System

A single-file **Python / Tkinter** desktop application for managing **RMA (Return Material Authorization)** repairs in **SMT / PCBA electronics manufacturing**, backed by **Microsoft SQL Server**.

> Version 1.0.0 · MIT License · Windows is the primary target (Linux works with limitations, see [Platform notes](#platform-notes))

<!-- Add screenshots to docs/screenshots/ and uncomment:
![Dashboard](docs/screenshots/dashboard.png)
![RMA records](docs/screenshots/rma_list.png)
-->

## Features

- **RMA lifecycle tracking** – Open → In-Repair → Pending Shipment → Closed, plus On-Hold, Cancelled and Scrapped
- **Auto-numbered RMAs** – `RMA-YYYYMM-0001` style sequence
- **Repair entries** – repair mode, location / ref-des, part number, quantity, technician, date, notes
- **Pre- and post-inspection** – Visual / ICT / FCT results
- **Failure classification** – searchable failure-code library
- **Repeat-offender detection** – serial numbers received more than once are highlighted
- **TAT & aging dashboard** – warning / overdue highlighting (10 / 20 day thresholds)
- **Bulk shipment update** – scan serial numbers or auto-populate by filters, then close in one go
- **Board history** – timeline of everything that happened to a serial number
- **Audit log** – every save, bulk-ship and re-open (with mandatory reason) is recorded
- **Cost estimate** – calculated from a standard-cost spreadsheet
- **Excel export** – detail report (configurable columns) and summary report (by status / project / category / model / failure code / month / repeat offenders)
- **Local attachments** – photos and files stored per RMA on your machine
- **Excel-driven configuration** – projects, models, users, failure codes, costs and report columns are plain `.xlsx` files
- **No login screen** – the operating-system username is used automatically

## Requirements

| Item | Details |
|---|---|
| Python | 3.9 or newer (with Tkinter) |
| Database | Microsoft SQL Server (any edition, including Express) |
| ODBC driver | [Microsoft ODBC Driver 17 for SQL Server](https://learn.microsoft.com/sql/connect/odbc/download-odbc-driver-for-sql-server) (recommended) |
| Python packages | `openpyxl`, `pyodbc` (required); `tkcalendar`, `pillow` (optional) |

## Quick start

### 1. Get the code
```bash
git clone https://github.com/dineshnaidu10/rma-management-system.git
cd rma-management-system
```

### 2. Install dependencies
```bash
python -m venv .venv
.venv\Scripts\activate          # Windows  (Linux/macOS: source .venv/bin/activate)
pip install -r requirements.txt
```

### 3. Create the database
Open `sql/schema.sql` in SQL Server Management Studio (or `sqlcmd`) and run it. It creates the `RMA_DB` database and the single `rma` table.

Want demo data to try the app? Then also run `sql/sample_data.sql` (six dummy RMAs, including a repeat-offender and an overdue example). Delete them later with `DELETE FROM dbo.rma;`.

### 4. Configure the connection
```bash
copy db_config.ini.example db_config.ini      # Linux/macOS: cp
```
Edit `db_config.ini`:

```ini
[database]
server = YOUR_SERVER\INSTANCE
database = RMA_DB
driver = ODBC Driver 17 for SQL Server
trusted_connection = yes
```

For SQL authentication set `trusted_connection = no` and add `username` / `password` lines.

> `db_config.ini` is git-ignored. Never commit credentials.

### 5. (Optional) Load sample configuration
The app creates empty template files in `config/` on first run. To start with example projects, users, failure codes and costs, copy the sample files:

```bash
mkdir config
copy samples\config\*.xlsx config\          # Linux/macOS: cp samples/config/*.xlsx config/
```

### 6. Run
```bash
python rma_app.py
```

## Configuration files (`config/`)

All files are plain Excel workbooks. **Row 1 is a header; data starts on row 2.** Use the first worksheet. Press **Reload Config** in the Config tab after editing.

| File | Columns used | Purpose |
|---|---|---|
| `models_revision.xlsx` | A: Project · B: Model · C: Latest Revision | Project / model dropdowns and latest-revision lookup |
| `users.xlsx` | A: Username | "Received By" and "Technician" dropdowns |
| `failure_codes.xlsx` | A: Code · B: Description | Failure-code library (shown as `Code – Description`) |
| `standard_cost.xlsx` | A: Part Number · C: Unit Cost | Cost estimate = unit cost × quantity of repair entries (matched by part number) |
| `report_columns.xlsx` | A: Column Key | Columns exported in the detail report, in order. Remove a row to drop a column. (The "Include" column is currently informational only.) |

Valid report column keys: `rma_number, ncr_number, project, rma_category, model, revision, latest_revision, serial_number, rework_notice, received_date, received_by, failure_desc, pre_visual, pre_ict, pre_fct, post_visual, post_ict, post_fct, failure_code, shipment_date, po_number, shipment_remarks, status, cost_estimate, created_at, updated_at, created_by, updated_by`

## Folder layout

```
rma-management-system/
├── rma_app.py               # the entire application
├── db_config.ini            # your connection settings (you create this, git-ignored)
├── db_config.ini.example
├── requirements.txt
├── sql/
│   ├── schema.sql           # database + table + indexes
│   └── sample_data.sql      # optional demo rows
├── samples/config/          # example config workbooks
├── config/                  # live config (auto-created, git-ignored)
├── attachments/             # local attachments (auto-created, git-ignored)
└── rma_app.log              # runtime log (git-ignored)
```

## How it works

- **One flat table.** Repair entries, the audit log and attachment paths are stored as JSON text inside `NVARCHAR(MAX)` columns of the `rma` table.
- **One connection per call**, so several workstations can use the same database safely.
- **Attachments are local.** They are copied to `attachments/<rma_number>/` on the machine that added them and are **not** uploaded to SQL Server. Other workstations will not see them unless `attachments/` is on a shared drive.

### Status rules
- Setting a **shipment date** automatically sets the status to **Closed**.
- **Closed** requires a shipment date. For boards that never ship, use **Mark as Scrapped**.
- **Closed** / **Scrapped** records are locked until you press **Re-open RMA**, which requires a reason, sets the status to *In-Repair*, clears the shipment date and writes an audit entry.
- In the New RMA form, sections 05–10 (repairs, post-inspection, shipment, notes, attachments) appear once the RMA has been saved and re-opened for editing.

### TAT colours
Defaults are set at the top of `rma_app.py` (`TAT_WARN_DAYS = 10`, `TAT_DANGER_DAYS = 20`). Open RMAs older than the danger threshold are shown in red.

## Packaging as an executable (optional)

The repository ships only the `.py` file. If you want a standalone Windows build:

```bash
pip install pyinstaller
pyinstaller --noconfirm --windowed --onedir --name RMA rma_app.py
```

Notes:
- Prefer **`--onedir`**. With `--onefile` the app resolves its folder to a temporary directory, so `config/`, `db_config.ini` and `attachments/` would not persist between runs.
- Place `db_config.ini` and your `config/` folder next to the generated `RMA.exe` (inside `dist/RMA/`).
- The target machine needs the same ODBC driver installed.

## Troubleshooting

| Symptom | Fix |
|---|---|
| "db_config.ini not found" | The app wrote a template next to `rma_app.py`; edit it and restart. |
| "Cannot connect to SQL Server" | Check server name / instance, firewall, that the ODBC driver name in the ini matches an installed driver (`odbcad32.exe` → Drivers tab), and your permissions. |
| Certificate / SSL error with **ODBC Driver 18** | Driver 18 encrypts connections by default. Use Driver 17, or install a trusted certificate on the server. |
| "Invalid object name 'rma'" | Run `sql/schema.sql` against the database named in `db_config.ini`. |
| Date picker looks plain | Install the optional package: `pip install tkcalendar`. |
| No image thumbnails | Install the optional package: `pip install pillow`. |
| `openpyxl`/`pyodbc` missing | `pip install -r requirements.txt`. |

## Platform notes

Developed and used on Windows. On Linux you need `unixodbc` plus the Microsoft ODBC driver, and `python3-tk`. Opening folders / files uses `xdg-open`. Windows authentication (`trusted_connection = yes`) is not available on Linux; use SQL authentication.

## Limitations

- **No authentication or roles.** The app trusts the OS username; use SQL Server permissions to control access.
- Any user can delete RMA records (the deletion is logged to `rma_app.log`).
- Attachments are local-only (see above).
- Dates are stored as `YYYY-MM-DD` text.

## Contributing

Issues and pull requests are welcome. See [CONTRIBUTING.md](CONTRIBUTING.md). Please never include real customer or company data in issues or PRs.

## License

[MIT](LICENSE) © 2026 Dinesh Naidu

## Disclaimer

Provided as-is, without warranty. Test with sample data before using it with production records, and back up your database.
