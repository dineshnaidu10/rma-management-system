# Changelog

All notable changes to this project are documented here.
Format based on [Keep a Changelog](https://keepachangelog.com/), versioning follows [SemVer](https://semver.org/).

## [1.0.0] - 2026-10-10
### Added
- First public open-source release.
- SQL Server backend via `pyodbc` (single flat `rma` table).
- `sql/schema.sql`, `sql/sample_data.sql`, sample config workbooks in `samples/config/`.
### Changed
- Replaced SQLite with SQL Server.
- Repair entries, audit log and attachment paths are stored as JSON in `NVARCHAR(MAX)` columns.
- No login screen: the OS username is used automatically.
- BOM location linking removed; location / ref-des is free text.
- Attachments are stored locally (`attachments/<rma_number>/`), not on the server.
- Removed the "Recent RMA Activity" dashboard section.
