import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import os
import sys
import json
import logging
import configparser
import getpass
import threading
from datetime import datetime, date, timedelta
from pathlib import Path
from typing import Optional

# ── Third-party ────────────────────────────────────────────────────────────────
try:
    from tkcalendar import DateEntry
    HAS_CAL = True
except ImportError:
    HAS_CAL = False

try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    HAS_XLSX = True
except ImportError:
    HAS_XLSX = False

try:
    import pyodbc
    HAS_PYODBC = True
except ImportError:
    HAS_PYODBC = False

# ══════════════════════════════════════════════════════════════════════════════
#  CONSTANTS & PATHS
# ══════════════════════════════════════════════════════════════════════════════

APP_TITLE   = "RMA Management System  v2.0"
APP_VERSION = "2.0.0"
BASE_DIR    = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
CONFIG_DIR  = BASE_DIR / "config"
BOM_DIR     = BASE_DIR / "bom"
DB_CONFIG   = BASE_DIR / "db_config.ini"
LOG_PATH    = BASE_DIR / "rma_app.log"
LOCAL_ATTACHMENTS_DIR = BASE_DIR / "attachments"

RMA_TYPES    = ["In-Warranty", "Non-Warranty", "Refurbishing", "Investigation"]
STATUS_LIST  = ["Open", "In-Repair", "Pending Shipment", "Closed", "On-Hold", "Cancelled", "Scrapped"]
STATUS_MANUAL= ["Open", "In-Repair", "Pending Shipment", "On-Hold", "Cancelled"]
REPAIR_MODES = [
    "Component Replacement", "Re-Soldering", "Cleaning", "Firmware Update",
    "Mechanical Repair", "PCB Repair", "BGA Rework", "No Fault Found", "Scrap", "Other",
]
INSP_RESULTS = ["Pass", "Fail", "N/A", "Pending"]

STATUS_COLOR = {
    "Open":             "#2196F3",
    "In-Repair":        "#FF9800",
    "Pending Shipment": "#9C27B0",
    "Closed":           "#4CAF50",
    "On-Hold":          "#F44336",
    "Cancelled":        "#9E9E9E",
    "Scrapped":         "#8B0000",
}

C_BG       = "#1E2228"
C_PANEL    = "#252B34"
C_CARD     = "#2C3440"
C_ACCENT   = "#00B4D8"
C_ACCENT2  = "#0077B6"
C_SUCCESS  = "#4CAF50"
C_WARN     = "#FF9800"
C_DANGER   = "#F44336"
C_TEXT     = "#E8EDF2"
C_MUTED    = "#8A95A0"
C_BORDER   = "#3A4250"
C_HEADER   = "#141820"
C_ROW_ALT  = "#303844"
C_SEL      = "#003F5C"
C_ENTRY_BG = "#161B22"
C_ENTRY_FG = "#E8EDF2"

FONT_MAIN  = ("Segoe UI", 9)
FONT_BOLD  = ("Segoe UI", 9, "bold")
FONT_TITLE = ("Segoe UI", 12, "bold")
FONT_SMALL = ("Segoe UI", 8)
FONT_MONO  = ("Courier New", 9)

TAT_WARN_DAYS   = 10
TAT_DANGER_DAYS = 20

C_REPEAT_BG       = "#3D3500"
C_REPEAT_FG       = "#FFD600"
REPEAT_XLSX_COLOR = "FFD600"

# ══════════════════════════════════════════════════════════════════════════════
#  LOGGING
# ══════════════════════════════════════════════════════════════════════════════
logging.basicConfig(
    filename=str(LOG_PATH),
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════════════════════
#  DB CONFIG LOADER
# ══════════════════════════════════════════════════════════════════════════════

def _load_db_config() -> dict:
    cfg = configparser.ConfigParser()
    if not DB_CONFIG.exists():
        cfg["database"] = {
            "server":             "YOUR_SERVER\\INSTANCE",
            "database":           "RMA_DB",
            "driver":             "ODBC Driver 17 for SQL Server",
            "trusted_connection": "yes",
            "; username":         "rma_user",
            "; password":         "your_password",
        }
        with open(DB_CONFIG, "w") as f:
            cfg.write(f)
        raise FileNotFoundError(
            f"db_config.ini not found.\n"
            f"A template has been created at:\n{DB_CONFIG}\n\n"
            "Edit it with your SQL Server connection details and restart."
        )
    cfg.read(DB_CONFIG)
    return dict(cfg["database"])


def _build_conn_str(c: dict) -> str:
    trusted = c.get("trusted_connection", "yes").strip().lower()
    if trusted == "yes":
        return (
            f"DRIVER={{{c['driver']}}};"
            f"SERVER={c['server']};"
            f"DATABASE={c['database']};"
            "Trusted_Connection=yes;"
        )
    return (
        f"DRIVER={{{c['driver']}}};"
        f"SERVER={c['server']};"
        f"DATABASE={c['database']};"
        f"UID={c['username']};"
        f"PWD={c['password']};"
    )


# ══════════════════════════════════════════════════════════════════════════════
#  DATABASE LAYER  —  SQL Server via pyodbc, single flat table
# ══════════════════════════════════════════════════════════════════════════════

class DB:
    """
    SQL Server wrapper.
    All child data (repair_entries, audit_log, attachment_paths) lives as
    JSON strings inside NVARCHAR(MAX) columns of the single `rma` table.
    A fresh connection is used per call so concurrent workstations are safe.
    """

    def __init__(self):
        cfg_dict = _load_db_config()
        self._conn_str = _build_conn_str(cfg_dict)
        self._test_connection()

    def _test_connection(self):
        try:
            with pyodbc.connect(self._conn_str, timeout=8) as cn:
                cn.execute("SELECT 1")
            log.info("SQL Server connection verified.")
        except Exception as exc:
            raise ConnectionError(
                f"Cannot connect to SQL Server.\n"
                f"Check db_config.ini.\n\nError: {exc}"
            )

    def _cn(self):
        return pyodbc.connect(self._conn_str, autocommit=True, timeout=30)

    # ── Generic helpers ───────────────────────────────────────────────────────
    def execute(self, sql: str, params=()):
        with self._cn() as cn:
            cn.execute(sql, params)

    def fetchall(self, sql: str, params=()) -> list:
        with self._cn() as cn:
            cur  = cn.execute(sql, params)
            cols = [c[0] for c in cur.description]
            return [dict(zip(cols, row)) for row in cur.fetchall()]

    def fetchone(self, sql: str, params=()):
        with self._cn() as cn:
            cur  = cn.execute(sql, params)
            cols = [c[0] for c in cur.description]
            row  = cur.fetchone()
            return dict(zip(cols, row)) if row else None

    # ── JSON helpers ──────────────────────────────────────────────────────────
    @staticmethod
    def json_load(val) -> list:
        if not val:
            return []
        try:
            return json.loads(val)
        except Exception:
            return []

    @staticmethod
    def json_dump(obj) -> str:
        return json.dumps(obj, ensure_ascii=False)

    # ── Repeat-offender ───────────────────────────────────────────────────────
    def repeat_offender_serials(self) -> set:
        rows = self.fetchall(
            "SELECT serial_number FROM rma "
            "GROUP BY serial_number HAVING COUNT(*) > 1"
        )
        return {r["serial_number"] for r in rows}

    # ── RMA number sequencing ─────────────────────────────────────────────────
    def next_rma_number(self) -> str:
        today = datetime.now().strftime("%Y%m")
        row   = self.fetchone(
            "SELECT TOP 1 rma_number FROM rma "
            "WHERE rma_number LIKE ? ORDER BY id DESC",
            (f"RMA-{today}-%",),
        )
        seq = int(row["rma_number"].split("-")[-1]) + 1 if row else 1
        return f"RMA-{today}-{seq:04d}"


# ══════════════════════════════════════════════════════════════════════════════
#  CONFIG MANAGER  (Excel master files)
# ══════════════════════════════════════════════════════════════════════════════

class ConfigManager:
    def __init__(self):
        self.projects: dict      = {}
        self.model_revisions: dict = {}
        self.users: list         = []
        self.failure_codes: list = []
        self.standard_costs: dict= {}
        self.report_columns: list= []
        self.errors: list        = []
        self.reload()

    def reload(self):
        self.errors.clear()
        self._load_models_revision()
        self._load_users()
        self._load_failure_codes()
        self._load_standard_costs()
        self._load_report_columns()

    def get_models(self, project: str) -> list:
        return sorted(self.projects.get(project, []))

    def get_projects(self) -> list:
        return sorted(self.projects.keys())

    def get_latest_revision(self, model: str) -> str:
        return self.model_revisions.get(model, "N/A")

    # ── Loaders ───────────────────────────────────────────────────────────────
    def _load_models_revision(self):
        path = CONFIG_DIR / "models_revision.xlsx"
        if not path.exists():
            self._create_template_models_revision(path)
            self.errors.append(f"Created template: {path.name} – please populate it.")
            return
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb.active
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row[0]:
                    continue
                project = str(row[0]).strip()
                model   = str(row[1]).strip() if row[1] else ""
                rev     = str(row[2]).strip() if row[2] else "N/A"
                if project not in self.projects:
                    self.projects[project] = []
                if model and model not in self.projects[project]:
                    self.projects[project].append(model)
                if model:
                    self.model_revisions[model] = rev
        except Exception as exc:
            self.errors.append(f"models_revision.xlsx: {exc}")

    def _load_users(self):
        path = CONFIG_DIR / "users.xlsx"
        if not path.exists():
            self._create_template_users(path)
            self.errors.append(f"Created template: {path.name}")
            return
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb.active
            self.users = [
                str(row[0]).strip()
                for row in ws.iter_rows(min_row=2, values_only=True)
                if row[0]
            ]
        except Exception as exc:
            self.errors.append(f"users.xlsx: {exc}")

    def _load_failure_codes(self):
        path = CONFIG_DIR / "failure_codes.xlsx"
        if not path.exists():
            self._create_template_failure_codes(path)
            self.errors.append(f"Created template: {path.name}")
            return
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb.active
            self.failure_codes = [
                f"{row[0]} – {row[1]}" if row[1] else str(row[0])
                for row in ws.iter_rows(min_row=2, values_only=True)
                if row[0]
            ]
        except Exception as exc:
            self.errors.append(f"failure_codes.xlsx: {exc}")

    def _load_standard_costs(self):
        path = CONFIG_DIR / "standard_cost.xlsx"
        if not path.exists():
            self._create_template_standard_cost(path)
            self.errors.append(f"Created template: {path.name}")
            return
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb.active
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0] and row[2]:
                    self.standard_costs[str(row[0]).strip()] = float(row[2])
        except Exception as exc:
            self.errors.append(f"standard_cost.xlsx: {exc}")

    def _load_report_columns(self):
        path = CONFIG_DIR / "report_columns.xlsx"
        default_cols = [
            "rma_number", "ncr_number", "project", "rma_category", "model",
            "revision", "latest_revision", "serial_number", "rework_notice",
            "received_date", "received_by", "failure_desc",
            "pre_visual", "pre_ict", "pre_fct",
            "post_visual", "post_ict", "post_fct",
            "failure_code", "shipment_date", "po_number",
            "shipment_remarks", "status", "cost_estimate",
        ]
        if not path.exists():
            self._create_template_report_columns(path, default_cols)
            self.errors.append(f"Created template: {path.name}")
            self.report_columns = default_cols
            return
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb.active
            self.report_columns = [
                str(row[0]).strip()
                for row in ws.iter_rows(min_row=2, values_only=True)
                if row[0]
            ]
        except Exception as exc:
            self.errors.append(f"report_columns.xlsx: {exc}")
            self.report_columns = default_cols

    # ── Template creators ─────────────────────────────────────────────────────
    def _create_template_models_revision(self, path: Path):
        path.parent.mkdir(parents=True, exist_ok=True)
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Models"
        ws.append(["Project", "Model", "Latest Revision"])
        ws.append(["PROJECT_A", "MODEL_X1", "Rev3"])
        ws.append(["PROJECT_B", "MODEL_Y1", "Rev2"])
        wb.save(path)

    def _create_template_users(self, path: Path):
        path.parent.mkdir(parents=True, exist_ok=True)
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Users"
        ws.append(["Username", "Full Name", "Role"])
        ws.append(["jsmith", "John Smith", "QA Engineer"])
        wb.save(path)

    def _create_template_failure_codes(self, path: Path):
        path.parent.mkdir(parents=True, exist_ok=True)
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Failure Codes"
        ws.append(["Code", "Description", "Category"])
        for c in [
            ("FC-001","Open Circuit","Solder"), ("FC-002","Short Circuit","Solder"),
            ("FC-003","Missing Component","Assembly"), ("FC-004","Wrong Component","Assembly"),
            ("FC-005","Damaged Component","Mechanical"), ("FC-006","BGA Delamination","Solder"),
            ("FC-007","Contamination","Process"), ("FC-008","ESD Damage","Handling"),
            ("FC-009","Firmware Issue","Software"), ("FC-010","No Fault Found","NFF"),
        ]:
            ws.append(c)
        wb.save(path)

    def _create_template_standard_cost(self, path: Path):
        path.parent.mkdir(parents=True, exist_ok=True)
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Standard Cost"
        ws.append(["Part Number", "Description", "Unit Cost (USD)", "Currency"])
        wb.save(path)

    def _create_template_report_columns(self, path: Path, cols: list):
        path.parent.mkdir(parents=True, exist_ok=True)
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Report Columns"
        ws.append(["Column Key", "Display Name", "Include (Yes/No)"])
        labels = {
            "rma_number":"RMA #","ncr_number":"NCR #","project":"Project",
            "rma_category":"RMA Category","model":"Model","revision":"Revision",
            "latest_revision":"Latest Revision","serial_number":"Serial Number",
            "rework_notice":"Rework Notice #","received_date":"Received Date",
            "received_by":"Received By","failure_desc":"Failure Description",
            "pre_visual":"Pre Visual","pre_ict":"Pre ICT","pre_fct":"Pre FCT",
            "post_visual":"Post Visual","post_ict":"Post ICT","post_fct":"Post FCT",
            "failure_code":"Failure Code","shipment_date":"Shipment Date",
            "po_number":"PO #","shipment_remarks":"Shipment Remarks",
            "status":"Status","cost_estimate":"Cost Estimate (USD)",
        }
        for col in cols:
            ws.append([col, labels.get(col, col), "Yes"])
        wb.save(path)


# ══════════════════════════════════════════════════════════════════════════════
#  HELPER WIDGETS
# ══════════════════════════════════════════════════════════════════════════════

class DatePickerEntry(tk.Frame):
    def __init__(self, parent, initial: str = "", allow_blank: bool = False, **kw):
        super().__init__(parent, bg=C_PANEL,
                         highlightthickness=1, highlightbackground=C_BORDER)
        self._allow_blank = allow_blank
        self._is_blank    = allow_blank and not initial
        self._var = tk.StringVar(
            value="" if self._is_blank
            else (initial or datetime.now().strftime("%Y-%m-%d"))
        )
        if self._is_blank:
            self._build_blank_state()
        else:
            self._build_date_widget()

    def _build_blank_state(self):
        for w in self.winfo_children():
            w.destroy()
        self._placeholder = tk.Label(
            self, text="— Not shipped —", bg=C_ENTRY_BG, fg=C_MUTED,
            font=FONT_MAIN, relief="flat", padx=8, pady=2, cursor="hand2",
        )
        self._placeholder.pack(side="left")
        self._placeholder.bind("<Button-1>", lambda e: self._set_today_and_build())
        tk.Button(self, text="📅  Set Date", bg=C_PANEL, fg=C_ACCENT,
                  relief="flat", cursor="hand2", font=FONT_SMALL,
                  command=self._set_today_and_build).pack(side="left", padx=(4, 0))

    def _set_today_and_build(self):
        self._is_blank = False
        self._var.set(datetime.now().strftime("%Y-%m-%d"))
        self._build_date_widget()

    def _build_date_widget(self):
        for w in self.winfo_children():
            w.destroy()
        if HAS_CAL:
            self._entry = DateEntry(
                self, textvariable=self._var, date_pattern="yyyy-mm-dd",
                background=C_ACCENT2, foreground=C_TEXT,
                headersbackground=C_HEADER, headersforeground=C_TEXT,
                normalbackground=C_CARD, normalforeground=C_TEXT,
                weekendbackground=C_CARD, weekendforeground=C_ACCENT,
                selectbackground=C_ACCENT, selectforeground=C_TEXT,
                font=FONT_MAIN, width=12,
            )
        else:
            self._entry = tk.Entry(
                self, textvariable=self._var,
                bg=C_ENTRY_BG, fg=C_ENTRY_FG, insertbackground=C_ACCENT,
                relief="flat", font=FONT_MAIN, width=14,
                highlightthickness=1, highlightbackground=C_BORDER,
                highlightcolor=C_ACCENT,
            )
        self._entry.pack(side="left")
        if not HAS_CAL:
            tk.Button(self, text="📅", bg=C_PANEL, fg=C_ACCENT,
                      relief="flat", cursor="hand2", font=FONT_SMALL,
                      command=self._popup).pack(side="left", padx=(2, 0))
        if self._allow_blank:
            tk.Button(self, text="✕", bg=C_PANEL, fg=C_MUTED,
                      relief="flat", cursor="hand2", font=FONT_SMALL,
                      command=self._clear).pack(side="left", padx=(2, 0))

    def _clear(self):
        self._is_blank = True
        self._var.set("")
        self._build_blank_state()

    def _popup(self):
        top = tk.Toplevel(self)
        top.title("Pick Date")
        top.configure(bg=C_PANEL)
        top.resizable(False, False)
        try:
            current = datetime.strptime(self._var.get(), "%Y-%m-%d")
        except ValueError:
            current = datetime.now()
        cal_frame = tk.Frame(top, bg=C_PANEL)
        cal_frame.pack(padx=10, pady=10)
        self._cal_year  = current.year
        self._cal_month = current.month

        def refresh():
            for w in cal_frame.winfo_children():
                w.destroy()
            import calendar
            cal = calendar.monthcalendar(self._cal_year, self._cal_month)
            hdr = datetime(self._cal_year, self._cal_month, 1).strftime("%B %Y")
            nav = tk.Frame(cal_frame, bg=C_PANEL)
            nav.grid(row=0, column=0, columnspan=7, pady=(0, 4))
            tk.Button(nav, text="◀", bg=C_PANEL, fg=C_ACCENT, relief="flat",
                      command=lambda: prev_m()).pack(side="left")
            tk.Label(nav, text=hdr, bg=C_PANEL, fg=C_TEXT,
                     font=FONT_BOLD, width=14).pack(side="left")
            tk.Button(nav, text="▶", bg=C_PANEL, fg=C_ACCENT, relief="flat",
                      command=lambda: next_m()).pack(side="left")
            for ci, d in enumerate(["Mo","Tu","We","Th","Fr","Sa","Su"]):
                tk.Label(cal_frame, text=d, bg=C_PANEL, fg=C_MUTED,
                         font=FONT_SMALL, width=3).grid(row=1, column=ci)
            for ri, week in enumerate(cal):
                for ci, day in enumerate(week):
                    if day == 0:
                        tk.Label(cal_frame, text="", bg=C_PANEL,
                                 width=3).grid(row=ri+2, column=ci)
                    else:
                        def pick(d=day):
                            self._var.set(
                                datetime(self._cal_year, self._cal_month, d
                                         ).strftime("%Y-%m-%d"))
                            top.destroy()
                        bg = C_ACCENT if (day == current.day and
                             self._cal_month == current.month and
                             self._cal_year == current.year) else C_CARD
                        tk.Button(cal_frame, text=str(day), bg=bg, fg=C_TEXT,
                                  relief="flat", width=3, font=FONT_SMALL,
                                  command=pick).grid(row=ri+2, column=ci,
                                                     padx=1, pady=1)
        def prev_m():
            self._cal_month -= 1
            if self._cal_month < 1:
                self._cal_month = 12; self._cal_year -= 1
            refresh()
        def next_m():
            self._cal_month += 1
            if self._cal_month > 12:
                self._cal_month = 1; self._cal_year += 1
            refresh()
        refresh()

    def get(self) -> str:
        if self._allow_blank and self._is_blank:
            return ""
        return self._var.get().strip()

    def set(self, val: str):
        val = (val or "").strip()
        if not val or val == "N/A":
            if self._allow_blank:
                self._clear()
            else:
                self._var.set(datetime.now().strftime("%Y-%m-%d"))
            return
        if self._allow_blank and self._is_blank:
            self._is_blank = False
            self._build_date_widget()
        self._var.set(val)


class SearchableCombo(tk.Frame):
    def __init__(self, parent, values=(), width=20, **kw):
        super().__init__(parent, bg=C_PANEL)
        self._all_values = list(values)
        self._var  = tk.StringVar()
        self._entry = tk.Entry(
            self, textvariable=self._var,
            bg=C_ENTRY_BG, fg=C_ENTRY_FG, insertbackground=C_ACCENT,
            relief="flat", font=FONT_MAIN, width=width,
            highlightthickness=1, highlightbackground=C_BORDER,
            highlightcolor=C_ACCENT,
        )
        self._entry.pack(side="left", fill="x", expand=True)
        self._list_win = None
        self._var.trace_add("write", self._on_type)
        self._entry.bind("<FocusOut>", lambda e: self.after(150, self._close_if_unfocused))
        self._entry.bind("<Return>",   lambda e: self._close_list())
        self._entry.bind("<Escape>",   lambda e: self._close_list())
        self._entry.bind("<Down>",     self._move_down)

    def _close_if_unfocused(self):
        if not self._list_win:
            return
        try:
            focused = self.focus_get()
        except Exception:
            focused = None
        if focused and str(focused).startswith(str(self._list_win)):
            return
        self._close_list()

    def _on_type(self, *_):
        term = self._var.get().lower()
        self._show_list([v for v in self._all_values if term in v.lower()])

    def _show_list(self, items):
        self._close_list()
        if not items:
            return
        x = self._entry.winfo_rootx()
        y = self._entry.winfo_rooty() + self._entry.winfo_height()
        self._list_win = tk.Toplevel(self)
        self._list_win.wm_overrideredirect(True)
        self._list_win.geometry(f"+{x}+{y}")
        self._list_win.configure(bg=C_BORDER)
        inner = tk.Frame(self._list_win, bg=C_BORDER)
        inner.pack(padx=1, pady=1)
        visible = min(12, len(items))
        lb = tk.Listbox(inner, bg=C_CARD, fg=C_TEXT, selectbackground=C_ACCENT,
                        selectforeground=C_TEXT, relief="flat", font=FONT_MAIN,
                        height=visible, width=max(self._entry["width"], 24),
                        activestyle="none")
        vsb = ttk.Scrollbar(inner, orient="vertical", command=lb.yview)
        lb.configure(yscrollcommand=vsb.set)
        lb.pack(side="left", fill="both")
        vsb.pack(side="left", fill="y")
        for item in items:
            lb.insert("end", item)
        lb.bind("<<ListboxSelect>>", lambda e: self._pick(lb))
        lb.bind("<Return>", lambda e: self._pick(lb))
        lb.bind("<MouseWheel>",
                lambda e: lb.yview_scroll(-1*(e.delta//120), "units"))
        def _on_fo(e):
            try:
                nf = e.widget.focus_get()
            except Exception:
                nf = None
            if nf and str(nf).startswith(str(self._list_win)):
                return
            self.after(150, self._close_list)
        self._list_win.bind("<FocusOut>", _on_fo)
        lb.bind("<FocusOut>", _on_fo)

    def _pick(self, lb):
        sel = lb.curselection()
        if sel:
            self._var.set(lb.get(sel[0]))
        self._close_list()

    def _close_list(self):
        if self._list_win:
            try:
                self._list_win.destroy()
            except Exception:
                pass
            self._list_win = None

    def _move_down(self, event):
        if self._list_win:
            for w in self._list_win.winfo_children():
                if isinstance(w, tk.Frame):
                    for child in w.winfo_children():
                        if isinstance(child, tk.Listbox):
                            child.focus_set(); return
                if isinstance(w, tk.Listbox):
                    w.focus_set(); return

    def get(self) -> str:
        return self._var.get().strip()

    def set(self, val: str):
        self._var.set(val)

    def configure_values(self, values):
        self._all_values = list(values)


def styled_label(parent, text, font=FONT_MAIN, fg=C_MUTED, bg=C_PANEL, **kw):
    return tk.Label(parent, text=text, font=font, fg=fg, bg=bg, **kw)

def styled_entry(parent, textvariable=None, width=20, **kw):
    kw.setdefault("font", FONT_MAIN)
    return tk.Entry(parent, textvariable=textvariable, width=width,
                    bg=C_ENTRY_BG, fg=C_ENTRY_FG, insertbackground=C_ACCENT,
                    relief="flat",
                    highlightthickness=1, highlightbackground=C_BORDER,
                    highlightcolor=C_ACCENT, **kw)

def styled_combo(parent, values=(), textvariable=None, width=18, **kw):
    style = ttk.Style()
    style.configure("Dark.TCombobox",
                    fieldbackground=C_ENTRY_BG, background=C_ENTRY_BG,
                    foreground=C_ENTRY_FG, selectbackground=C_ACCENT,
                    selectforeground=C_HEADER, arrowcolor=C_ACCENT,
                    bordercolor=C_BORDER, lightcolor=C_BORDER, darkcolor=C_BORDER)
    style.map("Dark.TCombobox",
              fieldbackground=[("readonly", C_ENTRY_BG), ("disabled", C_PANEL)],
              foreground=[("readonly", C_ENTRY_FG), ("disabled", C_MUTED)])
    return ttk.Combobox(parent, values=values, textvariable=textvariable,
                        width=width, style="Dark.TCombobox",
                        font=FONT_MAIN, state="readonly", **kw)

def styled_button(parent, text, command=None, style="primary", **kw):
    colors = {
        "primary": (C_ACCENT,   C_HEADER),
        "success": (C_SUCCESS,  C_HEADER),
        "danger":  (C_DANGER,   C_HEADER),
        "muted":   (C_CARD,     C_TEXT),
        "warn":    (C_WARN,     C_HEADER),
    }
    bg, fg = colors.get(style, (C_ACCENT, C_HEADER))
    btn = tk.Button(parent, text=text, command=command,
                    bg=bg, fg=fg, activebackground=C_ACCENT2,
                    activeforeground=C_TEXT, relief="flat",
                    font=FONT_BOLD, padx=10, pady=4, cursor="hand2", **kw)
    btn.bind("<Enter>", lambda e: btn.config(bg=C_ACCENT2))
    btn.bind("<Leave>", lambda e: btn.config(bg=bg))
    return btn

def status_badge(parent, status: str) -> tk.Label:
    color = STATUS_COLOR.get(status, C_MUTED)
    return tk.Label(parent, text=f"  {status}  ",
                    bg=color, fg=C_HEADER, font=FONT_BOLD,
                    relief="flat", padx=4, pady=2)


# ══════════════════════════════════════════════════════════════════════════════
#  REPAIR ENTRIES  (BOM linking removed — free-text location)
# ══════════════════════════════════════════════════════════════════════════════

class RepairEntriesTable(tk.Frame):
    COLS = ["Repair Mode", "Location / Ref Des", "Part Number",
            "Description", "Qty", "Technician", "Date", "Notes"]

    def __init__(self, parent, cfg: ConfigManager, current_user: str, **kw):
        super().__init__(parent, bg=C_PANEL, **kw)
        self.cfg  = cfg
        self.user = current_user
        self._rows: list = []

        hdr = tk.Frame(self, bg=C_PANEL)
        hdr.pack(fill="x", pady=(0, 4))
        styled_label(hdr, "🔧  Repair Entries", font=FONT_BOLD, fg=C_ACCENT).pack(side="left")
        styled_button(hdr, "＋ Add Row", command=self.add_row,
                      style="primary").pack(side="right")

        hdr_row = tk.Frame(self, bg=C_HEADER)
        hdr_row.pack(fill="x")
        widths = [16, 16, 16, 18, 4, 12, 12, 18, 4]
        for i, (col, w) in enumerate(zip(self.COLS + [""], widths)):
            tk.Label(hdr_row, text=col, bg=C_HEADER, fg=C_MUTED,
                     font=FONT_SMALL, width=w, anchor="w").grid(
                row=0, column=i, padx=2, pady=2, sticky="w")

        outer = tk.Frame(self, bg=C_PANEL)
        outer.pack(fill="both", expand=True)
        self._canvas = tk.Canvas(outer, bg=C_PANEL, highlightthickness=0, height=140)
        sb = ttk.Scrollbar(outer, orient="vertical", command=self._canvas.yview)
        self._body = tk.Frame(self._canvas, bg=C_PANEL)
        self._body.bind("<Configure>",
                        lambda e: self._canvas.configure(
                            scrollregion=self._canvas.bbox("all")))
        self._canvas.create_window((0, 0), window=self._body, anchor="nw")
        self._canvas.configure(yscrollcommand=sb.set)
        self._canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

    def add_row(self, data: dict = None):
        idx    = len(self._rows)
        row_bg = C_CARD if idx % 2 == 0 else C_ROW_ALT
        fr = tk.Frame(self._body, bg=row_bg)
        fr.pack(fill="x", pady=1)

        v_mode  = tk.StringVar(value=data.get("repair_mode", "")       if data else "")
        v_loc   = tk.StringVar(value=data.get("repair_location", "")   if data else "")
        v_pn    = tk.StringVar(value=data.get("part_number", "")       if data else "")
        v_desc  = tk.StringVar(value=data.get("part_description", "")  if data else "")
        v_qty   = tk.StringVar(value=str(data.get("quantity", 1))      if data else "1")
        v_tech  = tk.StringVar(value=data.get("technician", self.user) if data else self.user)
        v_notes = tk.StringVar(value=data.get("repair_notes", "")      if data else "")

        styled_combo(fr, values=REPAIR_MODES, textvariable=v_mode, width=16
                     ).grid(row=0, column=0, padx=2, pady=1)
        styled_entry(fr, textvariable=v_loc, width=18
                     ).grid(row=0, column=1, padx=2, pady=1)
        styled_entry(fr, textvariable=v_pn, width=16
                     ).grid(row=0, column=2, padx=2, pady=1)
        styled_entry(fr, textvariable=v_desc, width=18
                     ).grid(row=0, column=3, padx=2, pady=1)
        styled_entry(fr, textvariable=v_qty, width=4
                     ).grid(row=0, column=4, padx=2, pady=1)
        styled_combo(fr, values=[""] + self.cfg.users, textvariable=v_tech, width=12
                     ).grid(row=0, column=5, padx=2, pady=1)

        date_fr = DatePickerEntry(fr, initial=data.get("repair_date", "") if data else "")
        date_fr.grid(row=0, column=6, padx=2, pady=1)

        styled_entry(fr, textvariable=v_notes, width=18
                     ).grid(row=0, column=7, padx=2, pady=1)

        def remove():
            fr.destroy()
            self._rows = [r for r in self._rows if r["frame"] is not fr]

        tk.Button(fr, text="✕", bg=row_bg, fg=C_DANGER,
                  relief="flat", font=FONT_BOLD, cursor="hand2",
                  command=remove).grid(row=0, column=8, padx=2)

        self._rows.append({
            "frame": fr, "date": date_fr,
            "v_mode": v_mode, "v_loc": v_loc, "v_pn": v_pn,
            "v_desc": v_desc, "v_qty": v_qty, "v_tech": v_tech,
            "v_notes": v_notes,
        })

    def get_entries(self) -> list:
        entries = []
        for r in self._rows:
            if not r["frame"].winfo_exists():
                continue
            entries.append({
                "repair_mode":      r["v_mode"].get() or "N/A",
                "repair_location":  r["v_loc"].get()  or "N/A",
                "part_number":      r["v_pn"].get()   or "N/A",
                "part_description": r["v_desc"].get() or "N/A",
                "quantity":         int(r["v_qty"].get() or 1),
                "technician":       r["v_tech"].get() or "N/A",
                "repair_date":      r["date"].get(),
                "repair_notes":     r["v_notes"].get() or "N/A",
            })
        return entries

    def clear(self):
        for r in self._rows:
            try:
                r["frame"].destroy()
            except Exception:
                pass
        self._rows.clear()

    def load_entries(self, entries: list):
        self.clear()
        for e in entries:
            self.add_row(data=e)

    # Kept for API compat — no BOM
    def set_bom(self, bom_rows: list):
        pass


# ══════════════════════════════════════════════════════════════════════════════
#  ATTACHMENT PANEL  (local only — nothing stored on SQL Server)
# ══════════════════════════════════════════════════════════════════════════════

class AttachmentPanel(tk.Frame):
    THUMB_SIZE = (80, 80)

    def __init__(self, parent, db: DB, current_user: str, **kw):
        super().__init__(parent, bg=C_PANEL, **kw)
        self.db          = db
        self.user        = current_user
        self.rma_id: Optional[int]  = None
        self.rma_number: Optional[str] = None
        self._paths: list = []
        self._thumbs      = []

        tb = tk.Frame(self, bg=C_PANEL)
        tb.pack(fill="x", pady=(0, 4))
        styled_label(tb, "📎  Attachments (local)", font=FONT_BOLD,
                     fg=C_ACCENT).pack(side="left")
        styled_button(tb, "📂 Open Folder", command=self._open_folder,
                      style="muted").pack(side="right", padx=(0, 4))
        styled_button(tb, "＋ Add Files", command=self._pick_files,
                      style="primary").pack(side="right")

        outer = tk.Frame(self, bg=C_PANEL,
                         highlightthickness=1, highlightbackground=C_BORDER)
        outer.pack(fill="x")
        self._canvas = tk.Canvas(outer, bg=C_PANEL, highlightthickness=0, height=110)
        hsb = ttk.Scrollbar(outer, orient="horizontal", command=self._canvas.xview)
        self._thumb_frame = tk.Frame(self._canvas, bg=C_PANEL)
        self._thumb_frame.bind(
            "<Configure>",
            lambda e: self._canvas.configure(scrollregion=self._canvas.bbox("all")))
        self._canvas.create_window((0, 0), window=self._thumb_frame, anchor="nw")
        self._canvas.configure(xscrollcommand=hsb.set)
        self._canvas.pack(fill="x")
        hsb.pack(fill="x")

    # ── Public ────────────────────────────────────────────────────────────────
    def load_rma(self, rma_id: int, rma_number: str, attachment_paths_json: str):
        self.rma_id     = rma_id
        self.rma_number = rma_number
        self._paths     = DB.json_load(attachment_paths_json)
        self._refresh_thumbs()

    def get_paths_json(self) -> str:
        return DB.json_dump(self._paths)

    # ── Folder ────────────────────────────────────────────────────────────────
    def _att_dir(self) -> Path:
        name = self.rma_number or str(self.rma_id or "unsaved")
        d    = LOCAL_ATTACHMENTS_DIR / name
        d.mkdir(parents=True, exist_ok=True)
        return d

    def _open_folder(self):
        d = self._att_dir()
        try:
            if os.name == "nt":
                os.startfile(str(d))
            else:
                os.system(f"xdg-open '{d}'")
        except Exception as exc:
            messagebox.showerror("Open Error", str(exc), parent=self)

    # ── Pick & copy ───────────────────────────────────────────────────────────
    def _pick_files(self):
        if self.rma_id is None:
            messagebox.showinfo("Save First",
                                "Please save the RMA before adding attachments.",
                                parent=self)
            return
        paths = filedialog.askopenfilenames(
            title="Select Files / Photos",
            filetypes=[
                ("Images & common",
                 "*.png *.jpg *.jpeg *.bmp *.gif *.tif *.tiff "
                 "*.pdf *.docx *.xlsx *.txt"),
                ("All files", "*.*"),
            ],
            parent=self,
        )
        if not paths:
            return
        import shutil
        dest_dir = self._att_dir()
        for src_path in paths:
            src  = Path(src_path)
            ts   = datetime.now().strftime("%H%M%S")
            dest = dest_dir / f"{ts}_{src.name}"
            try:
                shutil.copy2(src, dest)
                self._paths.append(str(dest))
            except Exception as exc:
                messagebox.showerror("Copy Error",
                                     f"Could not copy {src.name}:\n{exc}",
                                     parent=self)
        self._refresh_thumbs()

    # ── Thumbnails ────────────────────────────────────────────────────────────
    def _refresh_thumbs(self):
        for w in self._thumb_frame.winfo_children():
            w.destroy()
        self._thumbs.clear()
        if not self._paths:
            tk.Label(self._thumb_frame, text="No attachments yet.",
                     bg=C_PANEL, fg=C_MUTED, font=FONT_SMALL,
                     padx=12, pady=10).pack(side="left")
            return
        for fp_str in list(self._paths):
            fp   = Path(fp_str)
            cell = tk.Frame(self._thumb_frame, bg=C_CARD, padx=4, pady=4)
            cell.pack(side="left", padx=4, pady=4)
            thumb_shown = False
            if HAS_PIL and fp.suffix.lower() in (".png",".jpg",".jpeg",
                                                   ".bmp",".gif",".tif",".tiff"):
                try:
                    img = Image.open(fp)
                    img.thumbnail(self.THUMB_SIZE, Image.LANCZOS)
                    photo = ImageTk.PhotoImage(img)
                    self._thumbs.append(photo)
                    lbl = tk.Label(cell, image=photo, bg=C_CARD, cursor="hand2")
                    lbl.pack()
                    lbl.bind("<Button-1>", lambda e, p=fp: self._open_file(p))
                    thumb_shown = True
                except Exception:
                    pass
            if not thumb_shown:
                ext = fp.suffix.upper() or "FILE"
                icon_lbl = tk.Label(cell, text=f"📄\n{ext}", bg=C_CARD,
                                    fg=C_MUTED, font=FONT_SMALL,
                                    width=8, height=4, cursor="hand2")
                icon_lbl.pack()
                icon_lbl.bind("<Button-1>", lambda e, p=fp: self._open_file(p))
            name_txt = fp.name
            if len(name_txt) > 14:
                name_txt = name_txt[:11] + "…"
            tk.Label(cell, text=name_txt, bg=C_CARD, fg=C_MUTED,
                     font=FONT_SMALL, width=10, anchor="center").pack()
            tk.Button(cell, text="✕", bg=C_CARD, fg=C_DANGER,
                      relief="flat", font=FONT_SMALL, cursor="hand2",
                      command=lambda p=fp_str: self._remove(p)).pack()

    def _open_file(self, path: Path):
        try:
            if os.name == "nt":
                os.startfile(str(path))
            else:
                os.system(f"xdg-open '{path}'")
        except Exception as exc:
            messagebox.showerror("Open Error", str(exc), parent=self)

    def _remove(self, fp_str: str):
        fp = Path(fp_str)
        if not messagebox.askyesno("Remove Attachment",
                                   f"Remove {fp.name}?", parent=self):
            return
        self._paths = [p for p in self._paths if p != fp_str]
        try:
            fp.unlink(missing_ok=True)
        except Exception:
            pass
        self._refresh_thumbs()


# ══════════════════════════════════════════════════════════════════════════════
#  RMA FORM
# ══════════════════════════════════════════════════════════════════════════════

class RMAForm(tk.Toplevel):
    def __init__(self, parent, db: DB, cfg: ConfigManager,
                 current_user: str, rma_id: int = None, on_save=None,
                 clone_data: dict = None):
        super().__init__(parent)
        self.db      = db
        self.cfg     = cfg
        self.user    = current_user
        self.rma_id  = rma_id
        self.on_save = on_save
        self._is_edit = rma_id is not None
        is_clone = clone_data is not None and not self._is_edit

        self.title("New RMA Entry  [Cloned]" if is_clone
                   else ("Edit RMA" if self._is_edit else "New RMA Entry"))
        self.configure(bg=C_BG)
        self.resizable(True, True)
        self.geometry("1020x820")
        self.grab_set()

        self._build_ui()
        if self._is_edit:
            self._load_data()
        elif is_clone:
            self._apply_clone(clone_data)

    # ── Build UI ──────────────────────────────────────────────────────────────
    def _build_ui(self):
        hdr = tk.Frame(self, bg=C_HEADER, pady=8)
        hdr.pack(fill="x")
        title = "✏  Edit RMA" if self._is_edit else "➕  New RMA Entry"
        tk.Label(hdr, text=title, bg=C_HEADER, fg=C_TEXT,
                 font=FONT_TITLE, padx=16).pack(side="left")

        foot = tk.Frame(self, bg=C_HEADER, pady=8)
        foot.pack(fill="x", side="bottom")
        styled_button(foot, "💾  Save RMA", command=self._save,
                      style="primary").pack(side="right", padx=(4, 12))
        styled_button(foot, "✕  Cancel", command=self.destroy,
                      style="muted").pack(side="right", padx=4)

        outer  = tk.Frame(self, bg=C_BG)
        outer.pack(fill="both", expand=True)
        canvas = tk.Canvas(outer, bg=C_BG, highlightthickness=0)
        vsb    = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        self._scroll_frame = tk.Frame(canvas, bg=C_BG)
        self._scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self._scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(-1*(e.delta//120), "units"))
        body = self._scroll_frame

        # ── Section 01: RMA Identity ──────────────────────────────────────────
        self._section(body, "01  RMA IDENTITY")
        g1 = self._grid_frame(body)

        self.v_rma_num = tk.StringVar(value=self.db.next_rma_number())
        self._field(g1, 0, 0, "RMA #",
                    styled_entry(g1, textvariable=self.v_rma_num, width=18))

        self.v_ncr = tk.StringVar()
        self._field(g1, 0, 2, "NCR #",
                    styled_entry(g1, textvariable=self.v_ncr, width=18))

        self.v_category = tk.StringVar()
        self._field(g1, 1, 0, "RMA Category",
                    styled_combo(g1, values=RMA_TYPES,
                                 textvariable=self.v_category, width=18))

        self.v_status = tk.StringVar(value="Open")
        self._status_combo = styled_combo(g1, values=STATUS_MANUAL,
                                          textvariable=self.v_status, width=18)
        self._field(g1, 1, 2, "Status", self._status_combo)

        scrap_frame = tk.Frame(g1, bg=C_PANEL)
        scrap_frame.grid(row=3, column=2, sticky="w", padx=(0, 12), pady=(0, 4))
        self._scrap_btn = tk.Button(
            scrap_frame, text="🗑  Mark as Scrapped",
            bg="#5A1010", fg="#FF8A80",
            activebackground="#8B0000", activeforeground=C_TEXT,
            relief="flat", font=FONT_SMALL, padx=8, pady=3,
            cursor="hand2", command=self._mark_scrapped,
        )
        self._scrap_btn.pack(side="left")
        tk.Label(scrap_frame, text=" ← closes record without shipment",
                 bg=C_PANEL, fg=C_MUTED, font=FONT_SMALL).pack(side="left")

        self._reopen_btn = tk.Button(
            scrap_frame, text="↩  Re-open RMA",
            bg="#1A3A1A", fg="#80FF80",
            activebackground="#2D6A2D", activeforeground=C_TEXT,
            relief="flat", font=FONT_SMALL, padx=8, pady=3,
            cursor="hand2", command=self._reopen_rma,
        )
        self._reopen_hint = tk.Label(
            scrap_frame, text=" ← undo close / scrap",
            bg=C_PANEL, fg=C_MUTED, font=FONT_SMALL,
        )

        self.v_rework = tk.StringVar()
        self._field(g1, 2, 0, "Rework Notice #",
                    styled_entry(g1, textvariable=self.v_rework, width=18))

        # ── Section 02: Board Identity ────────────────────────────────────────
        self._section(body, "02  BOARD IDENTITY")
        g2 = self._grid_frame(body)

        self.v_project = tk.StringVar()
        proj_cb = styled_combo(g2, values=self.cfg.get_projects(),
                               textvariable=self.v_project, width=22)
        self._field(g2, 0, 0, "Project", proj_cb)

        self.v_model = tk.StringVar()
        self.model_cb = styled_combo(g2, values=[],
                                     textvariable=self.v_model, width=22)
        self._field(g2, 0, 2, "Model", self.model_cb)

        self.v_revision = tk.StringVar()
        self._field(g2, 1, 0, "Revision",
                    styled_entry(g2, textvariable=self.v_revision, width=12))

        self.v_latest_rev = tk.StringVar()
        self._field(g2, 1, 2, "Latest Revision",
                    styled_entry(g2, textvariable=self.v_latest_rev, width=12))

        self.v_serial = tk.StringVar()
        self._field(g2, 2, 0, "Serial Number",
                    styled_entry(g2, textvariable=self.v_serial, width=22))

        def on_project_change(*_):
            models = self.cfg.get_models(self.v_project.get())
            self.model_cb["values"] = models
            if models:
                self.v_model.set(models[0])

        def on_model_change(*_):
            lr = self.cfg.get_latest_revision(self.v_model.get())
            self.v_latest_rev.set(lr)
            self.v_revision.set(lr)

        self.v_project.trace_add("write", on_project_change)
        self.v_model.trace_add("write", on_model_change)

        # ── Section 03: Receiving ─────────────────────────────────────────────
        self._section(body, "03  RECEIVING")
        g3 = self._grid_frame(body)

        self.received_date = DatePickerEntry(
            g3, initial=datetime.now().strftime("%Y-%m-%d"))
        self._field(g3, 0, 0, "Received Date", self.received_date)

        self.v_received_by = tk.StringVar(value=self.user)
        self._field(g3, 0, 2, "Received By",
                    styled_combo(g3, values=self.cfg.users,
                                 textvariable=self.v_received_by, width=20))

        g3_txt = tk.Frame(body, bg=C_PANEL, padx=12, pady=4)
        g3_txt.pack(fill="x")
        styled_label(g3_txt, "Failure Description", fg=C_MUTED).pack(anchor="w")
        self.failure_desc = tk.Text(
            g3_txt, height=3, bg=C_ENTRY_BG, fg=C_ENTRY_FG,
            insertbackground=C_ACCENT, relief="flat", font=FONT_MAIN, wrap="word",
            highlightthickness=1, highlightbackground=C_BORDER,
            highlightcolor=C_ACCENT,
        )
        self.failure_desc.pack(fill="x", pady=(2, 0))

        # ── Section 04: Pre-Inspection ────────────────────────────────────────
        self._section(body, "04  PRE-INSPECTION (Receiving)")
        g4 = self._grid_frame(body)

        self.v_pre_visual = tk.StringVar(value="Pending")
        self._field(g4, 0, 0, "Visual Inspection",
                    styled_combo(g4, values=INSP_RESULTS,
                                 textvariable=self.v_pre_visual, width=14))
        self.v_pre_ict = tk.StringVar(value="Pending")
        self._field(g4, 0, 2, "ICT Test",
                    styled_combo(g4, values=INSP_RESULTS,
                                 textvariable=self.v_pre_ict, width=14))
        self.v_pre_fct = tk.StringVar(value="Pending")
        self._field(g4, 0, 4, "FCT Test",
                    styled_combo(g4, values=INSP_RESULTS,
                                 textvariable=self.v_pre_fct, width=14))

        g4r = tk.Frame(body, bg=C_PANEL, padx=12, pady=4)
        g4r.pack(fill="x")
        styled_label(g4r, "Pre-Inspection Defect Remarks", fg=C_MUTED).pack(anchor="w")
        self.pre_inspection_remarks = tk.Text(
            g4r, height=3, bg=C_ENTRY_BG, fg=C_ENTRY_FG,
            insertbackground=C_ACCENT, relief="flat", font=FONT_MAIN, wrap="word",
            highlightthickness=1, highlightbackground=C_BORDER,
            highlightcolor=C_ACCENT,
        )
        self.pre_inspection_remarks.pack(fill="x", pady=(2, 0))

        # ── Sections 05–09: visible only in Edit mode ─────────────────────────
        self._late_frames: list = []

        def _late_section(title: str):
            sep = tk.Frame(body, bg=C_PANEL)
            sep.pack(fill="x", padx=8, pady=(10, 2))
            tk.Label(sep, text=title, bg=C_PANEL, fg=C_ACCENT,
                     font=FONT_BOLD).pack(side="left")
            tk.Frame(sep, bg=C_BORDER, height=1).pack(
                side="left", fill="x", expand=True, padx=8, pady=4)
            self._late_frames.append(sep)

        def _late_grid() -> tk.Frame:
            fr = tk.Frame(body, bg=C_PANEL, padx=12, pady=4)
            fr.pack(fill="x")
            self._late_frames.append(fr)
            return fr

        def _late_text() -> tk.Frame:
            fr = tk.Frame(body, bg=C_PANEL, padx=12, pady=4)
            fr.pack(fill="x")
            self._late_frames.append(fr)
            return fr

        _late_section("05  REPAIR ENTRIES")
        repair_outer = tk.Frame(body, bg=C_PANEL, padx=12, pady=4)
        repair_outer.pack(fill="x")
        self._late_frames.append(repair_outer)
        self.repair_table = RepairEntriesTable(repair_outer, self.cfg, self.user)
        self.repair_table.pack(fill="both", expand=True)

        _late_section("06  FAILURE CLASSIFICATION")
        g6 = _late_grid()
        self._fc_widget = SearchableCombo(g6, values=self.cfg.failure_codes, width=32)
        g6.columnconfigure(1, weight=1)
        self._field(g6, 0, 0, "Failure Code", self._fc_widget)

        _late_section("07  POST-INSPECTION (After Repair)")
        g7 = _late_grid()
        self.v_post_visual = tk.StringVar(value="N/A")
        self._field(g7, 0, 0, "Visual Inspection",
                    styled_combo(g7, values=INSP_RESULTS,
                                 textvariable=self.v_post_visual, width=14))
        self.v_post_ict = tk.StringVar(value="N/A")
        self._field(g7, 0, 2, "ICT Test",
                    styled_combo(g7, values=INSP_RESULTS,
                                 textvariable=self.v_post_ict, width=14))
        self.v_post_fct = tk.StringVar(value="N/A")
        self._field(g7, 0, 4, "FCT Test",
                    styled_combo(g7, values=INSP_RESULTS,
                                 textvariable=self.v_post_fct, width=14))

        _late_section("08  SHIPMENT")
        g8 = _late_grid()
        self.shipment_date = DatePickerEntry(g8, allow_blank=True)
        self._field(g8, 0, 0, "Shipment Date", self.shipment_date)
        self.v_po = tk.StringVar()
        self._field(g8, 0, 2, "PO #",
                    styled_entry(g8, textvariable=self.v_po, width=18))
        g8t = _late_text()
        styled_label(g8t, "Shipment Remarks", fg=C_MUTED).pack(anchor="w")
        self.shipment_remarks = tk.Text(
            g8t, height=2, bg=C_ENTRY_BG, fg=C_ENTRY_FG,
            insertbackground=C_ACCENT, relief="flat", font=FONT_MAIN, wrap="word",
            highlightthickness=1, highlightbackground=C_BORDER,
            highlightcolor=C_ACCENT,
        )
        self.shipment_remarks.pack(fill="x", pady=(2, 0))

        _late_section("09  INTERNAL NOTES")
        g9t = _late_text()
        self.notes_text = tk.Text(
            g9t, height=3, bg=C_ENTRY_BG, fg=C_ENTRY_FG,
            insertbackground=C_ACCENT, relief="flat", font=FONT_MAIN, wrap="word",
            highlightthickness=1, highlightbackground=C_BORDER,
            highlightcolor=C_ACCENT,
        )
        self.notes_text.pack(fill="x")

        _late_section("10  ATTACHMENTS / PHOTOS")
        att_outer = tk.Frame(body, bg=C_PANEL, padx=12, pady=4)
        att_outer.pack(fill="x")
        self._late_frames.append(att_outer)
        self._attach_panel = AttachmentPanel(att_outer, self.db, self.user)
        self._attach_panel.pack(fill="x")

        if not self._is_edit:
            for fr in self._late_frames:
                fr.pack_forget()
            notice = tk.Frame(body, bg=C_PANEL, padx=12, pady=8)
            notice.pack(fill="x")
            tk.Label(notice,
                     text="ℹ  Sections 05–10 (Repair, Post-Inspection, Shipment, "
                          "Notes, Attachments) appear after the RMA is saved and "
                          "reopened for editing.",
                     bg=C_PANEL, fg=C_MUTED, font=FONT_SMALL,
                     wraplength=700, justify="left").pack(anchor="w")

    # ── Helpers ───────────────────────────────────────────────────────────────
    def _section(self, parent, title: str):
        fr = tk.Frame(parent, bg=C_PANEL)
        fr.pack(fill="x", padx=8, pady=(10, 2))
        tk.Label(fr, text=title, bg=C_PANEL, fg=C_ACCENT,
                 font=FONT_BOLD).pack(side="left")
        tk.Frame(fr, bg=C_BORDER, height=1).pack(
            side="left", fill="x", expand=True, padx=8, pady=4)

    def _grid_frame(self, parent) -> tk.Frame:
        fr = tk.Frame(parent, bg=C_PANEL, padx=12, pady=4)
        fr.pack(fill="x")
        return fr

    def _field(self, parent, row, col, label, widget):
        styled_label(parent, label, fg=C_MUTED).grid(
            row=row*2, column=col, sticky="w", padx=(0, 4), pady=(4, 0))
        widget.grid(row=row*2+1, column=col, sticky="w",
                    padx=(0, 12), pady=(0, 4))

    # ── Load existing data ────────────────────────────────────────────────────
    def _load_data(self):
        row = self.db.fetchone("SELECT * FROM rma WHERE id=?", (self.rma_id,))
        if not row:
            return
        self.v_rma_num.set(row["rma_number"])
        self.v_ncr.set(row["ncr_number"])
        self.v_category.set(row["rma_category"])
        loaded_status = row["status"]
        self.v_status.set(loaded_status)
        if loaded_status in ("Closed", "Scrapped"):
            vals = STATUS_MANUAL + ([loaded_status]
                                    if loaded_status not in STATUS_MANUAL else [])
            self._status_combo["values"] = vals
            self._status_combo.configure(state="disabled")
            self._scrap_btn.pack_forget()
            self._reopen_btn.pack(side="left")
            self._reopen_hint.pack(side="left")
        self.v_rework.set(row["rework_notice"])
        self.v_project.set(row["project"])
        self.model_cb["values"] = self.cfg.get_models(row["project"])
        self.v_model.set(row["model"])
        self.v_revision.set(row["revision"])
        self.v_latest_rev.set(row["latest_revision"])
        self.v_serial.set(row["serial_number"])
        self.received_date.set(row["received_date"])
        self.v_received_by.set(row["received_by"])
        self.failure_desc.insert("1.0", row["failure_desc"])
        self.v_pre_visual.set(row["pre_visual"])
        self.v_pre_ict.set(row["pre_ict"])
        self.v_pre_fct.set(row["pre_fct"])
        self.v_post_visual.set(row["post_visual"])
        self.v_post_ict.set(row["post_ict"])
        self.v_post_fct.set(row["post_fct"])
        self._fc_widget.set(row["failure_code"])
        self.shipment_date.set(row["shipment_date"])
        self.v_po.set(row["po_number"])
        self.shipment_remarks.insert("1.0", row["shipment_remarks"])
        self.notes_text.insert("1.0", row["notes"])
        remarks_val = row.get("pre_inspection_remarks") or ""
        if remarks_val and remarks_val not in ("N/A", ""):
            self.pre_inspection_remarks.insert("1.0", remarks_val)
        self.repair_table.load_entries(
            DB.json_load(row.get("repair_entries") or ""))
        self._attach_panel.load_rma(
            self.rma_id, row["rma_number"],
            row.get("attachment_paths") or "")

    # ── Clone ─────────────────────────────────────────────────────────────────
    def _apply_clone(self, data: dict):
        proj = data.get("project", "")
        if proj:
            self.v_project.set(proj)
            self.model_cb["values"] = self.cfg.get_models(proj)
        model = data.get("model", "")
        if model:
            self.v_model.set(model)
        if data.get("revision"):
            self.v_revision.set(data["revision"])
        if data.get("rma_category"):
            self.v_category.set(data["rma_category"])
        fc = data.get("failure_code", "")
        if fc and fc != "N/A":
            self._fc_widget.set(fc)
        rwork = data.get("rework_notice", "")
        if rwork and rwork != "N/A":
            self.v_rework.set(rwork)
        self.v_rma_num.set(self.db.next_rma_number())
        self.v_serial.set("")
        fdesc = data.get("failure_desc", "")
        if fdesc and fdesc not in ("N/A", ""):
            self.failure_desc.delete("1.0", "end")
            self.failure_desc.insert("1.0", fdesc)

    # ── Scrap ─────────────────────────────────────────────────────────────────
    def _mark_scrapped(self):
        if self.v_status.get() == "Scrapped":
            messagebox.showinfo("Already Scrapped",
                                "This RMA is already marked as Scrapped.",
                                parent=self)
            return
        if not messagebox.askyesno(
            "Mark as Scrapped",
            "Mark this RMA as Scrapped?\n\n"
            "This is a terminal action — the record will be closed.\n"
            "A shipment date is NOT required.",
            icon="warning", parent=self,
        ):
            return
        self.v_status.set("Scrapped")
        self._save()

    # ── Re-open ───────────────────────────────────────────────────────────────
    def _reopen_rma(self):
        current = self.v_status.get()
        if current not in ("Closed", "Scrapped"):
            messagebox.showinfo("Not Closed",
                                "This RMA is not Closed or Scrapped.",
                                parent=self)
            return
        reason = simpledialog.askstring(
            "Re-open RMA",
            f"This RMA is currently '{current}'.\n\n"
            "Provide a reason for re-opening (required for audit trail):",
            parent=self,
        )
        if not reason or not reason.strip():
            messagebox.showwarning("Reason Required",
                                   "A reason is required to re-open a closed RMA.",
                                   parent=self)
            return
        reason = reason.strip()
        if not messagebox.askyesno(
            "Confirm Re-open",
            f"Re-open RMA '{self.v_rma_num.get()}'?\n\n"
            f"Current status : {current}\n"
            f"New status     : In-Repair\n"
            f"Reason         : {reason}\n\n"
            "The shipment date will be cleared.\n"
            "Press Save after reviewing to commit the change.",
            icon="warning", parent=self,
        ):
            return

        self.v_status.set("In-Repair")
        self.shipment_date.set("N/A")
        self._status_combo["values"] = STATUS_MANUAL
        self._status_combo.configure(state="readonly")
        self._reopen_btn.pack_forget()
        self._reopen_hint.pack_forget()
        self._scrap_btn.configure(state="normal", text="🗑  Mark as Scrapped")
        self._scrap_btn.pack(side="left")

        if self.rma_id:
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            existing = self.db.fetchone(
                "SELECT audit_log FROM rma WHERE id=?", (self.rma_id,))
            old_log = DB.json_load(
                existing.get("audit_log") if existing else "") or []
            old_log.append({
                "timestamp":  now,
                "user":       self.user,
                "action":     "REOPEN",
                "old_status": current,
                "new_status": "In-Repair",
                "details":    (f"RMA {self.v_rma_num.get()} re-opened from "
                               f"'{current}' by {self.user}. Reason: {reason}"),
            })
            self.db.execute(
                "UPDATE rma SET [audit_log]=? WHERE id=?",
                (DB.json_dump(old_log), self.rma_id),
            )
            log.info(f"REOPEN RMA id={self.rma_id} from '{current}' "
                     f"by {self.user}. Reason: {reason}")

        messagebox.showinfo(
            "RMA Re-opened",
            "Status set to 'In-Repair'.\n\n"
            "Review any fields as needed, then press  Save  to commit.",
            parent=self,
        )

    # ── Save ──────────────────────────────────────────────────────────────────
    def _save(self):
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        def g(var) -> str:
            v = var.get().strip()
            return v if v else "N/A"

        def gt(widget) -> str:
            v = widget.get("1.0", "end-1c").strip()
            return v if v else "N/A"

        required = {
            "RMA #":         self.v_rma_num.get().strip(),
            "RMA Category":  self.v_category.get().strip(),
            "Project":       self.v_project.get().strip(),
            "Model":         self.v_model.get().strip(),
            "Serial Number": self.v_serial.get().strip(),
            "Received By":   self.v_received_by.get().strip(),
        }
        for field, val in required.items():
            if not val:
                messagebox.showwarning("Missing Field",
                                       f"'{field}' is required.", parent=self)
                return

        repair_entries = self.repair_table.get_entries()
        cost = sum(
            self.cfg.standard_costs.get(e["part_number"], 0.0) * e["quantity"]
            for e in repair_entries
        )

        shipment_date_val = self.shipment_date.get()
        current_status    = g(self.v_status)

        if current_status == "Closed" and (
                not shipment_date_val or shipment_date_val == "N/A"):
            messagebox.showwarning(
                "Cannot Close Without Shipment",
                "Status 'Closed' requires a Shipment Date.\n\n"
                "Set a shipment date, or use '🗑 Mark as Scrapped'.",
                parent=self,
            )
            return

        if shipment_date_val and shipment_date_val not in ("N/A", ""):
            if current_status not in ("Closed", "Cancelled", "Scrapped"):
                current_status = "Closed"
                self.v_status.set("Closed")
        else:
            if current_status == "Closed":
                current_status = "Open"
                self.v_status.set("Open")

        repair_json = DB.json_dump(repair_entries)
        attach_json = self._attach_panel.get_paths_json() if self._is_edit else "[]"

        action = "UPDATE" if self._is_edit else "INSERT"
        rma_num_str = self.v_rma_num.get().strip()
        if self._is_edit:
            existing = self.db.fetchone(
                "SELECT audit_log FROM rma WHERE id=?", (self.rma_id,))
            old_log = DB.json_load(
                existing.get("audit_log") if existing else "") or []
        else:
            old_log = []
        old_log.append({
            "timestamp":  now,
            "user":       self.user,
            "action":     action,
            "new_status": current_status,
            "details":    f"RMA {rma_num_str} saved by {self.user}",
        })
        audit_json = DB.json_dump(old_log)

        data = {
            "rma_number":             self.v_rma_num.get().strip(),
            "ncr_number":             g(self.v_ncr),
            "project":                g(self.v_project),
            "rma_category":           g(self.v_category),
            "model":                  g(self.v_model),
            "revision":               g(self.v_revision),
            "latest_revision":        g(self.v_latest_rev),
            "serial_number":          g(self.v_serial),
            "rework_notice":          g(self.v_rework),
            "received_date":          self.received_date.get(),
            "received_by":            g(self.v_received_by),
            "failure_desc":           gt(self.failure_desc),
            "pre_visual":             g(self.v_pre_visual),
            "pre_ict":                g(self.v_pre_ict),
            "pre_fct":                g(self.v_pre_fct),
            "post_visual":            g(self.v_post_visual),
            "post_ict":               g(self.v_post_ict),
            "post_fct":               g(self.v_post_fct),
            "failure_code":           self._fc_widget.get() or "N/A",
            "shipment_date":          shipment_date_val if shipment_date_val else "N/A",
            "po_number":              g(self.v_po),
            "shipment_remarks":       gt(self.shipment_remarks),
            "status":                 current_status,
            "notes":                  gt(self.notes_text),
            "cost_estimate":          cost,
            "pre_inspection_remarks": gt(self.pre_inspection_remarks),
            "repair_entries":         repair_json,
            "audit_log":              audit_json,
            "attachment_paths":       attach_json,
            "updated_at":             now,
            "updated_by":             self.user,
        }

        if self._is_edit:
            # Wrap every column in [brackets] — guards against reserved words
            set_clause = ", ".join(f"[{k}]=?" for k in data)
            vals       = list(data.values()) + [self.rma_id]
            self.db.execute(f"UPDATE rma SET {set_clause} WHERE id=?", vals)
            rma_id = self.rma_id
        else:
            data["created_at"] = now
            data["created_by"] = self.user
            cols = ", ".join(f"[{k}]" for k in data.keys())
            ph   = ", ".join("?" for _ in data)
            self.db.execute(f"INSERT INTO rma ({cols}) VALUES ({ph})",
                            list(data.values()))
            new = self.db.fetchone(
                "SELECT TOP 1 id FROM rma WHERE rma_number=?",
                (data["rma_number"],))
            rma_id = new["id"] if new else None

        if rma_id and not self._is_edit:
            self._attach_panel.rma_id     = rma_id
            self._attach_panel.rma_number = data["rma_number"]

        log.info(f"{action} RMA id={rma_id} rma_number={data['rma_number']} "
                 f"by {self.user}")
        if self.on_save:
            self.on_save()
        self.destroy()


# ══════════════════════════════════════════════════════════════════════════════
#  BOARD HISTORY VIEWER
# ══════════════════════════════════════════════════════════════════════════════

class BoardHistoryViewer(tk.Toplevel):
    def __init__(self, parent, db: DB, serial_number: str):
        super().__init__(parent)
        self.title(f"Board History – {serial_number}")
        self.configure(bg=C_BG)
        self.geometry("900x680")
        self.db = db

        hdr = tk.Frame(self, bg=C_HEADER, pady=8)
        hdr.pack(fill="x")
        tk.Label(hdr, text=f"📋  Full History: {serial_number}",
                 bg=C_HEADER, fg=C_TEXT, font=FONT_TITLE, padx=16).pack(side="left")

        outer = tk.Frame(self, bg=C_BG)
        outer.pack(fill="both", expand=True, padx=16, pady=12)

        rmas = db.fetchall(
            "SELECT * FROM rma WHERE serial_number=? ORDER BY received_date, id",
            (serial_number,))

        if not rmas:
            tk.Label(outer,
                     text="No RMA records found for this serial number.",
                     bg=C_BG, fg=C_MUTED, font=FONT_MAIN).pack(pady=40)
            return

        canvas = tk.Canvas(outer, bg=C_BG, highlightthickness=0)
        vsb    = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        body   = tk.Frame(canvas, bg=C_BG)
        body.bind("<Configure>",
                  lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=body, anchor="nw")
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        for idx, rma in enumerate(rmas):
            self._render_card(body, rma, idx)

    def _render_card(self, parent, rma, idx: int):
        color  = STATUS_COLOR.get(rma["status"], C_MUTED)
        card   = tk.Frame(parent, bg=C_CARD, pady=0)
        card.pack(fill="x", pady=6)
        stripe = tk.Frame(card, bg=color, width=6)
        stripe.pack(side="left", fill="y")
        body = tk.Frame(card, bg=C_CARD, padx=12, pady=10)
        body.pack(side="left", fill="both", expand=True)

        r1 = tk.Frame(body, bg=C_CARD)
        r1.pack(fill="x")
        tk.Label(r1, text=f"RMA #{rma['rma_number']}",
                 bg=C_CARD, fg=C_TEXT, font=FONT_BOLD).pack(side="left")
        status_badge(r1, rma["status"]).pack(side="left", padx=8)
        tk.Label(r1, text=rma["rma_category"],
                 bg=C_CARD, fg=C_ACCENT, font=FONT_SMALL).pack(side="left")
        tk.Label(r1,
                 text=f"Received: {rma['received_date']}  by  {rma['received_by']}",
                 bg=C_CARD, fg=C_MUTED, font=FONT_SMALL).pack(side="right")

        r2 = tk.Frame(body, bg=C_CARD)
        r2.pack(fill="x", pady=(4, 0))
        tk.Label(r2, text=f"Model: {rma['model']}  Rev: {rma['revision']}",
                 bg=C_CARD, fg=C_TEXT, font=FONT_MAIN).pack(side="left")
        if rma["failure_desc"] != "N/A":
            tk.Label(r2, text=f"│  {rma['failure_desc'][:80]}",
                     bg=C_CARD, fg=C_MUTED, font=FONT_SMALL).pack(
                         side="left", padx=8)

        repairs = DB.json_load(rma.get("repair_entries") or "")

        def _activity_row(pf, icon, text, date_str, fg=C_TEXT):
            row = tk.Frame(pf, bg=C_CARD)
            row.pack(fill="x", pady=1)
            tk.Label(row, text=date_str if date_str else "—",
                     bg=C_HEADER, fg=C_ACCENT, font=FONT_SMALL,
                     width=11, relief="flat", padx=3).pack(side="left", padx=(0, 6))
            tk.Label(row, text=f"{icon}  {text}",
                     bg=C_CARD, fg=fg, font=FONT_SMALL,
                     anchor="w").pack(side="left", fill="x")

        act_frame = tk.Frame(body, bg=C_CARD)
        act_frame.pack(fill="x", pady=(6, 0))
        tk.Label(act_frame, text="Activities",
                 bg=C_CARD, fg=C_MUTED, font=FONT_SMALL).pack(
                     anchor="w", pady=(0, 2))

        _activity_row(act_frame, "📥",
                      f"Received by {rma['received_by']}",
                      rma["received_date"])

        pre_results = []
        for k, lbl in [("pre_visual","Visual"),("pre_ict","ICT"),("pre_fct","FCT")]:
            if rma[k] not in ("N/A", ""):
                pre_results.append(f"{lbl}: {rma[k]}")
        if pre_results:
            _activity_row(act_frame, "🔬",
                          "Pre-Inspection  │  " + "  │  ".join(pre_results),
                          rma["received_date"])

        for rep in repairs:
            rep_date = rep.get("repair_date") or "N/A"
            tech = rep.get("technician", "")
            tech_info = f"  (by {tech})" if tech not in ("N/A", "") else ""
            _activity_row(act_frame, "🔧",
                          f"{rep.get('repair_mode','')}  │  "
                          f"{rep.get('part_number','')}  │  "
                          f"{rep.get('repair_location','')}  │  "
                          f"Qty: {rep.get('quantity',1)}{tech_info}",
                          rep_date)

        post_results = []
        for k, lbl in [("post_visual","Visual"),("post_ict","ICT"),("post_fct","FCT")]:
            if rma[k] not in ("N/A", ""):
                post_results.append(f"{lbl}: {rma[k]}")
        if post_results:
            post_date = (rma["updated_at"][:10] if rma["updated_at"] else "N/A")
            _activity_row(act_frame, "✅",
                          "Post-Inspection  │  " + "  │  ".join(post_results),
                          post_date)

        if rma["failure_code"] not in ("N/A", ""):
            fc_date = rma["updated_at"][:10] if rma["updated_at"] else "N/A"
            _activity_row(act_frame, "🏷",
                          f"Failure Code: {rma['failure_code']}",
                          fc_date, fg=C_WARN)

        audit_events = DB.json_load(rma.get("audit_log") or "")
        rendered_shipped = False

        for ev in audit_events:
            ev_date = (ev.get("timestamp") or "")[:10]
            ev_time = (ev.get("timestamp") or "")[11:16]
            ev_label = f"{ev_date} {ev_time}".strip()
            nv  = (ev.get("new_status") or "").strip()
            ov  = (ev.get("old_status") or "").strip()
            act = (ev.get("action") or "").strip().upper()

            if act == "REOPEN":
                who     = ev.get("user") or rma["updated_by"]
                details = ev.get("details", "") or ""
                reason  = ""
                marker  = "Reason: "
                idx     = details.find(marker)
                if idx != -1:
                    reason = details[idx + len(marker):]
                suffix = f"  │  Reason: {reason}" if reason else ""
                _activity_row(act_frame, "↩",
                              f"Re-opened from '{ov}' by {who}{suffix}",
                              ev_label, fg=C_WARN)
            elif nv == "Scrapped":
                who = ev.get("user") or rma["updated_by"]
                _activity_row(act_frame, "🗑",
                              f"Marked as Scrapped by {who}",
                              ev_label, fg=C_DANGER)
            elif nv == "Closed":
                who     = ev.get("user") or rma["updated_by"]
                po_info = (f"  │  PO: {rma['po_number']}"
                           if rma["po_number"] not in ("N/A", "") else "")
                closed_ts   = ev.get("timestamp") or ""
                was_reopened = any(
                    (e.get("action") or "").upper() == "REOPEN" and
                    (e.get("timestamp") or "") > closed_ts
                    for e in audit_events
                )
                if was_reopened:
                    _activity_row(act_frame, "✈",
                                  f"Shipped{po_info}  │  (later undone)",
                                  ev_label, fg=C_SUCCESS)
                else:
                    rendered_shipped = True

        if not rendered_shipped and rma["shipment_date"] and \
                rma["shipment_date"] != "N/A":
            po_info = (f"  │  PO: {rma['po_number']}"
                       if rma["po_number"] not in ("N/A", "") else "")
            _activity_row(act_frame, "✈",
                          f"Shipped{po_info}",
                          rma["shipment_date"], fg=C_SUCCESS)

        if rma["status"] == "Scrapped" and not any(
            (ev.get("new_status") or "") == "Scrapped"
            for ev in audit_events
        ):
            scrapped_date = (rma["updated_at"][:10]
                             if rma["updated_at"] else "N/A")
            _activity_row(act_frame, "🗑",
                          f"Marked as Scrapped by {rma['updated_by']}",
                          scrapped_date, fg=C_DANGER)

        footer = tk.Frame(body, bg=C_CARD)
        footer.pack(fill="x", pady=(6, 0))
        if rma["cost_estimate"] and rma["cost_estimate"] > 0:
            tk.Label(footer,
                     text=f"Est. Cost: USD {rma['cost_estimate']:.2f}",
                     bg=C_CARD, fg=C_WARN, font=FONT_SMALL).pack(
                         side="left", padx=4)
        try:
            recv = datetime.strptime(rma["received_date"], "%Y-%m-%d").date()
            if rma["shipment_date"] and rma["shipment_date"] != "N/A":
                ship    = datetime.strptime(rma["shipment_date"], "%Y-%m-%d").date()
                tat     = (ship - recv).days
                tat_col = (C_SUCCESS if tat <= TAT_WARN_DAYS
                           else (C_WARN if tat <= TAT_DANGER_DAYS else C_DANGER))
            else:
                tat     = (date.today() - recv).days
                tat_col = C_WARN if tat > TAT_WARN_DAYS else C_MUTED
            tk.Label(footer, text=f"TAT: {tat} days",
                     bg=C_CARD, fg=tat_col, font=FONT_SMALL).pack(
                         side="right", padx=4)
        except Exception:
            pass


# ══════════════════════════════════════════════════════════════════════════════
#  REPORT EXPORTER
# ══════════════════════════════════════════════════════════════════════════════

class ReportExporter:
    FRIENDLY = {
        "rma_number":"RMA #","ncr_number":"NCR #",
        "project":"Project","rma_category":"RMA Category",
        "model":"Model","revision":"Revision",
        "latest_revision":"Latest Revision","serial_number":"Serial Number",
        "rework_notice":"Rework Notice #","received_date":"Received Date",
        "received_by":"Received By","failure_desc":"Failure Description",
        "pre_visual":"Pre Visual","pre_ict":"Pre ICT","pre_fct":"Pre FCT",
        "post_visual":"Post Visual","post_ict":"Post ICT","post_fct":"Post FCT",
        "failure_code":"Failure Code","shipment_date":"Shipment Date",
        "po_number":"PO #","shipment_remarks":"Shipment Remarks",
        "status":"Status","cost_estimate":"Cost Est. (USD)",
        "created_at":"Created","updated_at":"Updated",
        "created_by":"Created By","updated_by":"Updated By",
    }

    def __init__(self, db: DB, cfg: ConfigManager):
        self.db  = db
        self.cfg = cfg

    def export_detail(self, rows: list, path: str, columns: list,
                      repeat_serials: set = None):
        if not HAS_XLSX:
            return False, "openpyxl not installed"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "RMA Detail"
        self._style_sheet(ws, rows, columns, repeat_serials or set())
        ws2 = wb.create_sheet("Repair Entries")
        self._repair_sheet(ws2, rows)
        wb.save(path)
        return True, path

    def export_summary(self, path: str, filters: dict):
        if not HAS_XLSX:
            return False, "openpyxl not installed"
        wb = openpyxl.Workbook()

        ws1 = wb.active
        ws1.title = "By Status"
        # SQL Server: wrap [status], repeat expression in GROUP BY
        statuses = self.db.fetchall(
            "SELECT [status], COUNT(*) as cnt, SUM(cost_estimate) as total_cost "
            "FROM rma GROUP BY [status] ORDER BY cnt DESC")
        ws1.append(["Status", "Count", "Total Cost (USD)"])
        for r in statuses:
            ws1.append([r["status"], r["cnt"], round(r["total_cost"] or 0, 2)])

        ws2 = wb.create_sheet("By Project")
        projs = self.db.fetchall(
            "SELECT project, COUNT(*) as cnt, SUM(cost_estimate) as total_cost "
            "FROM rma GROUP BY project ORDER BY cnt DESC")
        ws2.append(["Project", "Count", "Total Cost (USD)"])
        for r in projs:
            ws2.append([r["project"], r["cnt"], round(r["total_cost"] or 0, 2)])

        ws3 = wb.create_sheet("By Category")
        cats = self.db.fetchall(
            "SELECT rma_category, COUNT(*) as cnt FROM rma "
            "GROUP BY rma_category ORDER BY cnt DESC")
        ws3.append(["Category", "Count"])
        for r in cats:
            ws3.append([r["rma_category"], r["cnt"]])

        ws4 = wb.create_sheet("By Model")
        models = self.db.fetchall(
            "SELECT model, revision, COUNT(*) as cnt, AVG(cost_estimate) as avg_cost "
            "FROM rma GROUP BY model, revision ORDER BY cnt DESC")
        ws4.append(["Model", "Revision", "Count", "Avg Cost (USD)"])
        for r in models:
            ws4.append([r["model"], r["revision"], r["cnt"],
                        round(r["avg_cost"] or 0, 2)])

        ws5 = wb.create_sheet("By Failure Code")
        fcs = self.db.fetchall(
            "SELECT failure_code, COUNT(*) as cnt FROM rma "
            "WHERE failure_code != 'N/A' "
            "GROUP BY failure_code ORDER BY cnt DESC")
        ws5.append(["Failure Code", "Count"])
        for r in fcs:
            ws5.append([r["failure_code"], r["cnt"]])

        ws6 = wb.create_sheet("Monthly Trend")
        # SQL Server: LEFT() instead of substr(), repeat expression in GROUP BY
        trend = self.db.fetchall(
            "SELECT LEFT(received_date, 7) as month, COUNT(*) as cnt, "
            "SUM(cost_estimate) as total_cost "
            "FROM rma "
            "GROUP BY LEFT(received_date, 7) "
            "ORDER BY LEFT(received_date, 7)"
        )
        ws6.append(["Month", "Count", "Total Cost (USD)"])
        for r in trend:
            ws6.append([r["month"], r["cnt"], round(r["total_cost"] or 0, 2)])

        ws7 = wb.create_sheet("Repeat Offenders")
        repeat_rows = self.db.fetchall(
            "SELECT serial_number, COUNT(*) as rma_count, model, project, "
            "MIN(received_date) as first_received, "
            "MAX(received_date) as last_received, "
            "SUM(cost_estimate) as total_cost "
            "FROM rma "
            "GROUP BY serial_number, model, project "
            "HAVING COUNT(*) > 1 "
            "ORDER BY rma_count DESC, last_received DESC"
        )
        ws7.append(["Serial Number", "RMA Count", "Model", "Project",
                    "First Received", "Last Received", "Total Cost (USD)"])
        repeat_fill = PatternFill("solid", fgColor=REPEAT_XLSX_COLOR)
        repeat_font = Font(name="Segoe UI", size=9, color="1A1400", bold=True)
        for r in repeat_rows:
            ws7.append([
                r["serial_number"], r["rma_count"], r["model"], r["project"],
                r["first_received"], r["last_received"],
                round(r["total_cost"] or 0, 2),
            ])

        for ws in wb.worksheets:
            self._auto_width(ws)
            self._header_style(ws)

        for row_cells in ws7.iter_rows(min_row=2):
            for cell in row_cells:
                cell.fill = repeat_fill
                cell.font = repeat_font

        wb.save(path)
        return True, path

    def _style_sheet(self, ws, rows, columns, repeat_serials: set):
        header_fill = PatternFill("solid", fgColor="141820")
        header_font = Font(bold=True, color="00B4D8", name="Segoe UI", size=9)
        ws.append([self.FRIENDLY.get(c, c) for c in columns])
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        fill_even   = PatternFill("solid", fgColor="2C3440")
        fill_odd    = PatternFill("solid", fgColor="303844")
        fill_repeat = PatternFill("solid", fgColor=REPEAT_XLSX_COLOR)
        font_data   = Font(name="Segoe UI", size=9, color="E8EDF2")
        font_repeat = Font(name="Segoe UI", size=9, color="1A1400", bold=True)

        for ri, row in enumerate(rows):
            row_data = [row.get(c, "") for c in columns]
            ws.append(row_data)
            sn        = row.get("serial_number", "")
            is_repeat = sn in repeat_serials
            fill = fill_repeat if is_repeat else (
                fill_even if ri % 2 == 0 else fill_odd)
            fnt  = font_repeat if is_repeat else font_data
            for cell in ws[ri + 2]:
                cell.fill = fill
                cell.font = fnt
                cell.alignment = Alignment(vertical="center")

        self._auto_width(ws)
        ws.freeze_panes = "A2"

    def _repair_sheet(self, ws, rma_rows: list):
        ws.append(["RMA #", "Repair Mode", "Location", "Part Number",
                   "Description", "Qty", "Technician", "Date", "Notes"])
        for rma in rma_rows:
            for r in DB.json_load(rma.get("repair_entries") or ""):
                ws.append([
                    rma.get("rma_number", ""),
                    r.get("repair_mode", ""),
                    r.get("repair_location", ""),
                    r.get("part_number", ""),
                    r.get("part_description", ""),
                    r.get("quantity", 1),
                    r.get("technician", ""),
                    r.get("repair_date", ""),
                    r.get("repair_notes", ""),
                ])
        self._auto_width(ws)
        self._header_style(ws)

    def _auto_width(self, ws):
        for col in ws.columns:
            max_len = max((len(str(c.value or "")) for c in col), default=10)
            ws.column_dimensions[
                get_column_letter(col[0].column)].width = min(max_len + 4, 50)

    def _header_style(self, ws):
        hf = Font(bold=True, name="Segoe UI", size=9, color="00B4D8")
        hb = PatternFill("solid", fgColor="141820")
        for cell in ws[1]:
            cell.font = hf
            cell.fill = hb
            cell.alignment = Alignment(horizontal="center")


# ══════════════════════════════════════════════════════════════════════════════
#  REPORT / FILTER DIALOG
# ══════════════════════════════════════════════════════════════════════════════

class ReportDialog(tk.Toplevel):
    def __init__(self, parent, db: DB, cfg: ConfigManager):
        super().__init__(parent)
        self.db  = db
        self.cfg = cfg
        self.title("Export Report")
        self.configure(bg=C_PANEL)
        self.geometry("560x580")
        self.grab_set()
        self._build()

    def _build(self):
        tk.Label(self, text="📊  Export Report", bg=C_PANEL, fg=C_TEXT,
                 font=FONT_TITLE, pady=12, padx=16).pack(anchor="w")

        fr_type = tk.LabelFrame(self, text="Report Type", bg=C_PANEL, fg=C_MUTED,
                                font=FONT_BOLD, padx=8, pady=8)
        fr_type.pack(fill="x", padx=16, pady=4)
        self.v_type = tk.StringVar(value="detail")
        for val, lbl in [("detail",  "Detail Report  (one row per RMA)"),
                         ("summary", "Summary Report  (breakdowns by project/category/trend)")]:
            tk.Radiobutton(fr_type, text=lbl, variable=self.v_type, value=val,
                           bg=C_PANEL, fg=C_TEXT, selectcolor=C_CARD,
                           activebackground=C_PANEL, font=FONT_MAIN).pack(anchor="w")

        fr_filt = tk.LabelFrame(self, text="Filters", bg=C_PANEL, fg=C_MUTED,
                                font=FONT_BOLD, padx=8, pady=8)
        fr_filt.pack(fill="x", padx=16, pady=4)

        def row_pair(label, widget, r):
            tk.Label(fr_filt, text=label, bg=C_PANEL, fg=C_MUTED, font=FONT_MAIN,
                     width=16, anchor="w").grid(row=r, column=0, sticky="w", pady=2)
            widget.grid(row=r, column=1, sticky="w", pady=2, padx=4)

        self.v_proj   = tk.StringVar(value="All")
        self.v_cat    = tk.StringVar(value="All")
        self.v_status = tk.StringVar(value="All")
        self.date_from = DatePickerEntry(fr_filt)
        self.date_to   = DatePickerEntry(fr_filt)

        row_pair("Project",      styled_combo(fr_filt, values=["All"] + self.cfg.get_projects(),
                                              textvariable=self.v_proj, width=22), 0)
        row_pair("RMA Category", styled_combo(fr_filt, values=["All"] + RMA_TYPES,
                                              textvariable=self.v_cat, width=22), 1)
        row_pair("Status",       styled_combo(fr_filt, values=["All"] + STATUS_LIST,
                                              textvariable=self.v_status, width=22), 2)
        row_pair("Date From",    self.date_from, 3)
        row_pair("Date To",      self.date_to, 4)

        self.v_repeat_only = tk.BooleanVar(value=False)
        repeat_fr = tk.Frame(fr_filt, bg=C_PANEL)
        repeat_fr.grid(row=5, column=0, columnspan=2, sticky="w", pady=(6, 2))
        tk.Checkbutton(repeat_fr,
                       text="🔁  Repeat Offender only  (serial received > 1 time)",
                       variable=self.v_repeat_only,
                       bg=C_PANEL, fg=C_REPEAT_FG, selectcolor=C_CARD,
                       activebackground=C_PANEL, activeforeground=C_REPEAT_FG,
                       font=FONT_MAIN).pack(side="left")

        fr_cols = tk.LabelFrame(self, text="Columns (Detail Report)", bg=C_PANEL,
                                fg=C_MUTED, font=FONT_BOLD, padx=8, pady=8)
        fr_cols.pack(fill="x", padx=16, pady=4)
        tk.Label(fr_cols, text="Configured via config/report_columns.xlsx",
                 bg=C_PANEL, fg=C_MUTED, font=FONT_SMALL).pack(anchor="w")

        btn_fr = tk.Frame(self, bg=C_PANEL, pady=10)
        btn_fr.pack(fill="x", padx=16)
        styled_button(btn_fr, "📁  Export to Excel", command=self._export,
                      style="primary").pack(side="right", padx=4)
        styled_button(btn_fr, "Cancel", command=self.destroy,
                      style="muted").pack(side="right")

    def _build_where(self):
        clauses, params = [], []
        proj  = self.v_proj.get()
        cat   = self.v_cat.get()
        stat  = self.v_status.get()
        d_from = self.date_from.get()
        d_to   = self.date_to.get()
        repeat_only = self.v_repeat_only.get()
        if proj != "All":
            clauses.append("project=?");        params.append(proj)
        if cat != "All":
            clauses.append("rma_category=?");   params.append(cat)
        if stat != "All":
            clauses.append("[status]=?");        params.append(stat)
        if d_from:
            clauses.append("received_date>=?"); params.append(d_from)
        if d_to:
            clauses.append("received_date<=?"); params.append(d_to)
        if repeat_only:
            clauses.append(
                "serial_number IN ("
                "SELECT serial_number FROM rma "
                "GROUP BY serial_number HAVING COUNT(*) > 1)"
            )
        where = ("WHERE " + " AND ".join(clauses)) if clauses else ""
        return where, params

    def _export(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            title="Save Report As",
            initialfile=f"RMA_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        )
        if not path:
            return
        exporter = ReportExporter(self.db, self.cfg)
        where, params = self._build_where()
        if self.v_type.get() == "summary":
            ok, msg = exporter.export_summary(path, {})
        else:
            rows = self.db.fetchall(
                f"SELECT * FROM rma {where} ORDER BY received_date DESC", params)
            cols = self.cfg.report_columns or list(ReportExporter.FRIENDLY.keys())
            repeat_serials = self.db.repeat_offender_serials()
            ok, msg = exporter.export_detail(rows, path, cols, repeat_serials)
        if ok:
            messagebox.showinfo("Export Complete", f"Report saved:\n{path}", parent=self)
            self.destroy()
        else:
            messagebox.showerror("Export Failed", msg, parent=self)


# ══════════════════════════════════════════════════════════════════════════════
#  TAT DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════

class TATDashboard(tk.Frame):
    def __init__(self, parent, db: DB):
        super().__init__(parent, bg=C_BG)
        self.db = db
        self._build()

    def _build(self):
        tile_row = tk.Frame(self, bg=C_BG)
        tile_row.pack(fill="x", padx=16, pady=12)
        self._tiles = {}
        for key, label, color in [
            ("open",      "Open",           C_ACCENT),
            ("in_repair", "In-Repair",      C_WARN),
            ("overdue",   "Overdue (>20d)", C_DANGER),
            ("closed",    "Closed (30d)",   C_SUCCESS),
        ]:
            fr = tk.Frame(tile_row, bg=C_CARD, padx=20, pady=14)
            fr.pack(side="left", padx=6, fill="y")
            num = tk.Label(fr, text="—", bg=C_CARD, fg=color,
                           font=("Segoe UI", 24, "bold"))
            num.pack()
            tk.Label(fr, text=label, bg=C_CARD, fg=C_MUTED,
                     font=FONT_SMALL).pack()
            self._tiles[key] = num

        tk.Label(self, text="Aging Report – Open RMAs", bg=C_BG, fg=C_TEXT,
                 font=FONT_BOLD, padx=16).pack(anchor="w", pady=(8, 0))

        cols = ("RMA #", "Model", "Serial", "Status", "Received",
                "Age (days)", "Received By")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=16)
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=120, anchor="w")
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(fill="both", expand=True, padx=16, pady=8, side="left")
        vsb.pack(side="right", fill="y", pady=8)

        self.tree.tag_configure("overdue",  background="#3B1A1A", foreground="#F44336")
        self.tree.tag_configure("warning",  background="#332800", foreground="#FF9800")
        self.tree.tag_configure("normal",   background=C_CARD,   foreground=C_TEXT)
        self.refresh()

    def refresh(self):
        today = date.today()

        # Use [status] to avoid SQL Server reserved-word conflict
        open_cnt = self.db.fetchone(
            "SELECT COUNT(*) as c FROM rma WHERE [status]='Open'")["c"]
        repair_cnt = self.db.fetchone(
            "SELECT COUNT(*) as c FROM rma WHERE [status]='In-Repair'")["c"]
        threshold = (today - timedelta(days=20)).strftime("%Y-%m-%d")
        overdue_cnt = self.db.fetchone(
            "SELECT COUNT(*) as c FROM rma "
            "WHERE [status] NOT IN ('Closed','Cancelled','Scrapped') "
            "AND received_date <= ?", (threshold,))["c"]
        month_ago = (today - timedelta(days=30)).strftime("%Y-%m-%d")
        closed_cnt = self.db.fetchone(
            "SELECT COUNT(*) as c FROM rma "
            "WHERE [status]='Closed' AND updated_at >= ?", (month_ago,))["c"]

        self._tiles["open"].config(text=str(open_cnt))
        self._tiles["in_repair"].config(text=str(repair_cnt))
        self._tiles["overdue"].config(text=str(overdue_cnt))
        self._tiles["closed"].config(text=str(closed_cnt))

        for item in self.tree.get_children():
            self.tree.delete(item)

        rows = self.db.fetchall(
            "SELECT rma_number, model, serial_number, [status], "
            "received_date, received_by "
            "FROM rma "
            "WHERE [status] NOT IN ('Closed','Cancelled','Scrapped') "
            "ORDER BY received_date"
        )
        for row in rows:
            try:
                recv = datetime.strptime(row["received_date"], "%Y-%m-%d").date()
                age  = (today - recv).days
            except Exception:
                age = 0
            tag = ("overdue" if age > TAT_DANGER_DAYS
                   else ("warning" if age > TAT_WARN_DAYS else "normal"))
            self.tree.insert("", "end", values=(
                row["rma_number"], row["model"], row["serial_number"],
                row["status"], row["received_date"], age, row["received_by"],
            ), tags=(tag,))


# ══════════════════════════════════════════════════════════════════════════════
#  AUDIT LOG VIEWER
# ══════════════════════════════════════════════════════════════════════════════

class AuditLogViewer(tk.Frame):
    def __init__(self, parent, db: DB):
        super().__init__(parent, bg=C_BG)
        self.db = db
        self._build()

    def _build(self):
        hdr = tk.Frame(self, bg=C_BG, pady=8)
        hdr.pack(fill="x", padx=16)
        tk.Label(hdr, text="Audit Log", bg=C_BG, fg=C_TEXT,
                 font=FONT_BOLD).pack(side="left")
        styled_button(hdr, "⟳ Refresh", command=self.refresh,
                      style="muted").pack(side="right")

        cols = ("Timestamp", "RMA #", "User", "Action", "Details")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=28)
        widths = [155, 130, 100, 120, 460]
        for c, w in zip(cols, widths):
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w, anchor="w")
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(fill="both", expand=True, padx=16, pady=8, side="left")
        vsb.pack(side="right", fill="y", pady=8)
        self.refresh()

    def refresh(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        rows = self.db.fetchall(
            "SELECT TOP 1000 rma_number, audit_log, updated_at FROM rma "
            "WHERE audit_log IS NOT NULL AND audit_log != '' "
            "ORDER BY updated_at DESC"
        )
        entries = []
        for row in rows:
            for ev in DB.json_load(row.get("audit_log") or ""):
                entries.append((
                    ev.get("timestamp", ""),
                    row["rma_number"],
                    ev.get("user", ""),
                    ev.get("action", ""),
                    ev.get("details", ""),
                ))
        entries.sort(key=lambda x: x[0], reverse=True)
        for e in entries[:500]:
            self.tree.insert("", "end", values=e)


# ══════════════════════════════════════════════════════════════════════════════
#  BULK SHIPMENT UPDATE
# ══════════════════════════════════════════════════════════════════════════════

class BulkShipmentDialog(tk.Toplevel):
    def __init__(self, parent, db: DB, current_user: str, on_done=None):
        super().__init__(parent)
        self.db      = db
        self.user    = current_user
        self.on_done = on_done
        self._queue: list = []
        self._queued_ids: set = set()

        self.title("Bulk Shipment Update")
        self.configure(bg=C_BG)
        self.geometry("960x820")
        self.minsize(800, 680)
        self.grab_set()
        self._build_ui()
        self.after(150, lambda: self.scan_entry.focus_set())

    def _build_ui(self):
        hdr = tk.Frame(self, bg=C_HEADER, pady=8)
        hdr.pack(fill="x")
        tk.Label(hdr, text="🚚  Bulk Shipment Update", bg=C_HEADER,
                 fg=C_TEXT, font=FONT_TITLE, padx=16).pack(side="left")

        info_fr = tk.LabelFrame(
            self,
            text="Shipment Details (applied to every scanned record)",
            bg=C_PANEL, fg=C_MUTED, font=FONT_BOLD, padx=12, pady=10)
        info_fr.pack(fill="x", padx=16, pady=(12, 6))
        row = tk.Frame(info_fr, bg=C_PANEL)
        row.pack(fill="x")
        styled_label(row, "Shipment Date", fg=C_MUTED).grid(row=0, column=0, sticky="w")
        self.shipment_date = DatePickerEntry(
            row, initial=datetime.now().strftime("%Y-%m-%d"))
        self.shipment_date.grid(row=1, column=0, sticky="w", padx=(0,24), pady=(2,0))
        styled_label(row, "PO #", fg=C_MUTED).grid(row=0, column=1, sticky="w")
        self.v_po = tk.StringVar()
        styled_entry(row, textvariable=self.v_po, width=22
                     ).grid(row=1, column=1, sticky="w", padx=(0,24), pady=(2,0))
        styled_label(row, "Shipment Remarks (optional)", fg=C_MUTED
                     ).grid(row=0, column=2, sticky="w")
        self.v_remarks = tk.StringVar()
        styled_entry(row, textvariable=self.v_remarks, width=30
                     ).grid(row=1, column=2, sticky="w", pady=(2,0))

        scan_fr = tk.LabelFrame(self, text="Scan / Enter Serial Numbers",
                                bg=C_PANEL, fg=C_MUTED, font=FONT_BOLD,
                                padx=12, pady=10)
        scan_fr.pack(fill="x", padx=16, pady=6)
        scan_row = tk.Frame(scan_fr, bg=C_PANEL)
        scan_row.pack(fill="x")
        tk.Label(scan_row, text="🔫", bg=C_PANEL, fg=C_ACCENT,
                 font=FONT_TITLE).pack(side="left", padx=(0,6))
        self.v_scan = tk.StringVar()
        self.scan_entry = styled_entry(
            scan_row, textvariable=self.v_scan, width=36, font=FONT_MONO)
        self.scan_entry.pack(side="left", padx=(0,8), ipady=3)
        self.scan_entry.bind("<Return>", lambda e: self._on_scan())
        styled_button(scan_row, "＋ Add", command=self._on_scan,
                      style="primary").pack(side="left", padx=4)
        tk.Label(scan_fr,
                 text="Scanner input auto-submits on Enter.",
                 bg=C_PANEL, fg=C_MUTED, font=FONT_SMALL).pack(anchor="w", pady=(6,0))
        self.lbl_status = tk.Label(scan_fr, text="", bg=C_PANEL,
                                   fg=C_MUTED, font=FONT_SMALL)
        self.lbl_status.pack(anchor="w", pady=(4,0))

        auto_fr = tk.LabelFrame(
            self,
            text="⚡  Auto-Populate from Database  (no scanning needed)",
            bg=C_PANEL, fg=C_ACCENT, font=FONT_BOLD, padx=12, pady=10)
        auto_fr.pack(fill="x", padx=16, pady=(0,6))
        tk.Label(auto_fr,
                 text="Select filters, then click Auto-Populate to queue all "
                      "matching open RMAs at once.",
                 bg=C_PANEL, fg=C_MUTED, font=FONT_SMALL,
                 wraplength=780, justify="left").pack(anchor="w", pady=(0,8))

        filter_row = tk.Frame(auto_fr, bg=C_PANEL)
        filter_row.pack(fill="x")

        styled_label(filter_row, "Received Date", fg=C_MUTED
                     ).grid(row=0, column=0, sticky="w", padx=(0,4))
        self._auto_recv_date = DatePickerEntry(filter_row, allow_blank=True)
        self._auto_recv_date.grid(row=1, column=0, sticky="w",
                                  padx=(0,20), pady=(2,0))

        proj_list = ["(Any)"] + [r["project"] for r in self.db.fetchall(
            "SELECT DISTINCT project FROM rma ORDER BY project")]
        styled_label(filter_row, "Project", fg=C_MUTED
                     ).grid(row=0, column=1, sticky="w", padx=(0,4))
        self._auto_proj_var = tk.StringVar(value="(Any)")
        self._auto_proj_cb  = styled_combo(
            filter_row, values=proj_list,
            textvariable=self._auto_proj_var, width=20)
        self._auto_proj_cb.grid(row=1, column=1, sticky="w",
                                padx=(0,20), pady=(2,0))
        self._auto_proj_var.trace_add("write", self._refresh_auto_model_list)

        styled_label(filter_row, "Model", fg=C_MUTED
                     ).grid(row=0, column=2, sticky="w", padx=(0,4))
        self._auto_model_var = tk.StringVar(value="(Any)")
        self._auto_model_cb  = styled_combo(
            filter_row, values=["(Any)"],
            textvariable=self._auto_model_var, width=20)
        self._auto_model_cb.grid(row=1, column=2, sticky="w",
                                 padx=(0,20), pady=(2,0))

        recvby_list = ["(Any)"] + [r["received_by"] for r in self.db.fetchall(
            "SELECT DISTINCT received_by FROM rma ORDER BY received_by")]
        styled_label(filter_row, "Received By", fg=C_MUTED
                     ).grid(row=0, column=3, sticky="w", padx=(0,4))
        self._auto_recvby_var = tk.StringVar(value="(Any)")
        self._auto_recvby_cb  = styled_combo(
            filter_row, values=recvby_list,
            textvariable=self._auto_recvby_var, width=18)
        self._auto_recvby_cb.grid(row=1, column=3, sticky="w",
                                  padx=(0,20), pady=(2,0))

        auto_btn_row = tk.Frame(auto_fr, bg=C_PANEL)
        auto_btn_row.pack(fill="x", pady=(10,0))
        self._auto_preview_lbl = tk.Label(
            auto_btn_row, text="", bg=C_PANEL, fg=C_MUTED, font=FONT_SMALL)
        self._auto_preview_lbl.pack(side="left")
        styled_button(auto_btn_row, "🔍  Preview Match",
                      command=self._auto_preview, style="muted").pack(
                          side="right", padx=(4,0))
        styled_button(auto_btn_row, "⚡  Auto-Populate Queue",
                      command=self._auto_populate, style="primary").pack(
                          side="right", padx=4)

        foot = tk.Frame(self, bg=C_HEADER, pady=8)
        foot.pack(fill="x", side="bottom")
        self.lbl_count = tk.Label(foot, text="0 record(s) queued",
                                  bg=C_HEADER, fg=C_MUTED, font=FONT_SMALL)
        self.lbl_count.pack(side="left", padx=16)
        styled_button(foot, "✕ Remove Selected",
                      command=self._remove_selected, style="muted").pack(
                          side="left", padx=4)
        styled_button(foot, "✔  Commit Update",
                      command=self._commit, style="success").pack(
                          side="right", padx=(4,12))
        styled_button(foot, "✕  Close",
                      command=self.destroy, style="muted").pack(
                          side="right", padx=4)

        queue_fr = tk.Frame(self, bg=C_BG)
        queue_fr.pack(fill="both", expand=True, padx=16, pady=(6,0))
        tk.Label(queue_fr, text="Queued for Update",
                 bg=C_BG, fg=C_TEXT, font=FONT_BOLD).pack(anchor="w", pady=(0,4))

        cols = ("RMA #", "Serial #", "Current Status", "Received", "Result")
        self.tree = ttk.Treeview(queue_fr, columns=cols, show="headings",
                                  height=12)
        widths = [130, 130, 120, 100, 220]
        for c, w in zip(cols, widths):
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w, anchor="w")
        vsb = ttk.Scrollbar(queue_fr, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.tree.tag_configure("ok",    foreground=C_SUCCESS)
        self.tree.tag_configure("warn",  foreground=C_WARN)
        self.tree.tag_configure("error", foreground=C_DANGER)
        self.tree.bind("<Delete>",    lambda e: self._remove_selected())
        self.tree.bind("<BackSpace>", lambda e: self._remove_selected())

    # ── Auto-populate helpers ─────────────────────────────────────────────────
    def _refresh_auto_model_list(self, *_):
        proj = self._auto_proj_var.get()
        if proj == "(Any)":
            models = ["(Any)"] + [r["model"] for r in self.db.fetchall(
                "SELECT DISTINCT model FROM rma ORDER BY model")]
        else:
            models = ["(Any)"] + [r["model"] for r in self.db.fetchall(
                "SELECT DISTINCT model FROM rma WHERE project=? ORDER BY model",
                (proj,))]
        self._auto_model_cb["values"] = models
        self._auto_model_var.set("(Any)")

    def _build_auto_filter_query(self):
        clauses = ["[status] NOT IN ('Closed','Cancelled','Scrapped')"]
        params  = []
        recv_date = self._auto_recv_date.get()
        if recv_date:
            clauses.append("received_date = ?"); params.append(recv_date)
        proj = self._auto_proj_var.get()
        if proj != "(Any)":
            clauses.append("project = ?"); params.append(proj)
        model = self._auto_model_var.get()
        if model != "(Any)":
            clauses.append("model = ?"); params.append(model)
        recvby = self._auto_recvby_var.get()
        if recvby != "(Any)":
            clauses.append("received_by = ?"); params.append(recvby)
        where = "WHERE " + " AND ".join(clauses)
        return where, params

    def _auto_preview(self):
        where, params = self._build_auto_filter_query()
        rows = self.db.fetchall(
            f"SELECT id FROM rma {where} ORDER BY received_date, id", params)
        new_count = sum(1 for r in rows if r["id"] not in self._queued_ids)
        already   = len(rows) - new_count
        msg = f"✓  {new_count} open record(s) would be added"
        if already:
            msg += f"  ({already} already queued, skipped)"
        if new_count == 0 and already == 0:
            msg = "⚠  No open records match the current filter."
        self._auto_preview_lbl.config(
            text=msg,
            fg=C_SUCCESS if new_count > 0 else (C_WARN if already > 0 else C_DANGER))

    def _auto_populate(self):
        where, params = self._build_auto_filter_query()
        rows = self.db.fetchall(
            f"SELECT id, rma_number, serial_number, [status], received_date "
            f"FROM rma {where} ORDER BY received_date, id", params)
        if not rows:
            messagebox.showinfo("No Matches",
                                "No open RMA records match the selected filters.",
                                parent=self)
            return
        added = skipped_dup = 0
        for row in rows:
            if row["id"] in self._queued_ids:
                skipped_dup += 1
                continue
            self._queue.append(row)
            self._queued_ids.add(row["id"])
            self.tree.insert("", "end", iid=str(row["id"]), values=(
                row["rma_number"], row["serial_number"], row["status"],
                row["received_date"], "Ready (auto)",
            ), tags=("ok",))
            added += 1
        msg = f"✓  {added} record(s) added to queue."
        if skipped_dup:
            msg += f"  {skipped_dup} already queued, skipped."
        self._auto_preview_lbl.config(
            text=msg, fg=C_SUCCESS if added > 0 else C_WARN)
        self._update_count()
        self._flash_status(
            f"⚡ Auto-populate: {added} record(s) added.", C_SUCCESS)

    # ── Scan handling ─────────────────────────────────────────────────────────
    def _on_scan(self):
        raw = self.v_scan.get().strip()
        self.v_scan.set("")
        self.scan_entry.focus_set()
        if raw:
            self._process_serial(raw)

    def _process_serial(self, serial: str):
        rows = self.db.fetchall(
            "SELECT id, rma_number, serial_number, [status], received_date "
            "FROM rma WHERE serial_number=? "
            "ORDER BY received_date DESC, id DESC", (serial,))
        if not rows:
            self._flash_status(
                f"✗  No RMA record found for serial '{serial}'.", C_DANGER)
            return
        candidates = [r for r in rows
                      if r["status"] not in ("Closed","Cancelled","Scrapped")]
        row = candidates[0] if candidates else rows[0]
        if row["id"] in self._queued_ids:
            self._flash_status(f"⚠  {serial} is already queued.", C_WARN)
            return
        self._queue.append(row)
        self._queued_ids.add(row["id"])
        tag = ("warn" if row["status"] in ("Closed","Cancelled","Scrapped")
               else "ok")
        result_text = ("Ready" if tag == "ok"
                       else f"Already {row['status']}")
        self.tree.insert("", "end", iid=str(row["id"]), values=(
            row["rma_number"], row["serial_number"], row["status"],
            row["received_date"], result_text,
        ), tags=(tag,))
        if tag == "ok":
            self._flash_status(
                f"✓  {serial} → {row['rma_number']} added.", C_SUCCESS)
        self._update_count()

    def _flash_status(self, text: str, color: str):
        self.lbl_status.config(text=text, fg=color)

    def _update_count(self):
        self.lbl_count.config(text=f"{len(self._queue)} record(s) queued")

    def _remove_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        for iid in sel:
            self.tree.delete(iid)
            rid = int(iid)
            self._queued_ids.discard(rid)
            self._queue = [q for q in self._queue if q["id"] != rid]
        self._update_count()

    # ── Commit ────────────────────────────────────────────────────────────────
    def _commit(self):
        if not self._queue:
            messagebox.showinfo("Nothing to Update",
                                "Scan at least one serial number first.")
            return
        ship_date = self.shipment_date.get()
        po        = self.v_po.get().strip()
        remarks   = self.v_remarks.get().strip()
        if not ship_date:
            messagebox.showwarning("Missing Shipment Date",
                                   "Please set a shipment date.", parent=self)
            return
        if not po:
            if not messagebox.askyesno(
                "No PO Number",
                "No PO # was entered. Continue without a PO number?",
                parent=self):
                return
        po_val      = po      if po      else "N/A"
        remarks_val = remarks if remarks else "N/A"
        if not messagebox.askyesno(
            "Confirm Bulk Update",
            f"Update {len(self._queue)} record(s):\n\n"
            f"  Shipment Date : {ship_date}\n"
            f"  PO #          : {po_val}\n\n"
            "All matched RMAs will be marked Closed. Continue?",
            parent=self):
            return

        now     = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        updated = skipped = 0

        for item in self._queue:
            row = self.db.fetchone("SELECT * FROM rma WHERE id=?", (item["id"],))
            if not row:
                skipped += 1
                continue
            final_remarks = remarks_val if remarks else row["shipment_remarks"]

            self.db.execute(
                "UPDATE rma SET shipment_date=?, po_number=?, "
                "shipment_remarks=?, [status]=?, updated_at=?, updated_by=? "
                "WHERE id=?",
                (ship_date, po_val, final_remarks, "Closed",
                 now, self.user, item["id"]),
            )

            existing_log = DB.json_load(row.get("audit_log") or "")
            existing_log.append({
                "timestamp":  now,
                "user":       self.user,
                "action":     "BULK_SHIP_UPDATE",
                "old_status": row["status"],
                "new_status": "Closed",
                "details":    (f"Bulk shipment: {row['rma_number']} "
                               f"SN={row['serial_number']} "
                               f"shipped {ship_date}, PO {po_val}"),
            })
            self.db.execute(
                "UPDATE rma SET [audit_log]=? WHERE id=?",
                (DB.json_dump(existing_log), item["id"]),
            )
            log.info(f"Bulk ship: RMA {row['rma_number']} id={item['id']} "
                     f"→ Closed ship={ship_date} po={po_val} by {self.user}")
            updated += 1

        msg = f"{updated} record(s) updated and marked Closed."
        if skipped:
            msg += f"\n{skipped} record(s) skipped."
        messagebox.showinfo("Bulk Update Complete", msg, parent=self)
        if self.on_done:
            self.on_done()
        self.destroy()


# ══════════════════════════════════════════════════════════════════════════════
#  RMA LIST VIEW
# ══════════════════════════════════════════════════════════════════════════════

class RMAListView(tk.Frame):
    COLS = [
        ("RMA #",       "rma_number",   130),
        ("NCR #",       "ncr_number",    90),
        ("Category",    "rma_category",  90),
        ("Project",     "project",       90),
        ("Model",       "model",        100),
        ("Rev",         "revision",      50),
        ("Serial #",    "serial_number",120),
        ("Received",    "received_date", 90),
        ("By",          "received_by",   90),
        ("Status",      "status",        90),
        ("TAT (d)",     "_tat",          60),
        ("Cost (USD)",  "cost_estimate", 80),
    ]

    def __init__(self, parent, db: DB, cfg: ConfigManager, current_user: str):
        super().__init__(parent, bg=C_BG)
        self.db   = db
        self.cfg  = cfg
        self.user = current_user
        self._all_rows       = []
        self._repeat_serials: set = set()
        self._build()

    def _build(self):
        tb = tk.Frame(self, bg=C_HEADER, pady=6)
        tb.pack(fill="x")

        styled_button(tb, "➕  New RMA", command=self._new_rma,
                      style="primary").pack(side="left", padx=(12,4))
        styled_button(tb, "✏  Edit",    command=self._edit_rma,
                      style="muted").pack(side="left", padx=4)
        styled_button(tb, "📋  Board History", command=self._board_history,
                      style="muted").pack(side="left", padx=4)
        styled_button(tb, "🚚  Bulk Shipment", command=self._bulk_shipment,
                      style="warn").pack(side="left", padx=4)
        styled_button(tb, "📋  Clone",  command=self._clone_rma,
                      style="muted").pack(side="left", padx=4)
        styled_button(tb, "🗑  Delete", command=self._delete_rma,
                      style="danger").pack(side="left", padx=4)
        styled_button(tb, "📊  Export", command=self._export,
                      style="success").pack(side="left", padx=4)
        styled_button(tb, "⟳ Refresh", command=self.refresh,
                      style="muted").pack(side="left", padx=4)

        tk.Label(tb, text="Quick filter:", bg=C_HEADER, fg=C_MUTED,
                 font=FONT_SMALL).pack(side="left", padx=(16,4))
        for st in ["All"] + STATUS_LIST:
            color = STATUS_COLOR.get(st, C_MUTED)
            tk.Button(tb, text=st, bg=C_HEADER, fg=color,
                      relief="flat", font=FONT_SMALL, cursor="hand2", padx=6,
                      command=lambda s=st: self._quick_filter(s)
                      ).pack(side="left", padx=2)

        tk.Frame(tb, bg=C_BORDER, width=1).pack(side="left", fill="y", padx=4, pady=2)
        self._repeat_filter_active = False
        self._repeat_btn = tk.Button(
            tb, text="🔁 Repeat Offender", bg=C_HEADER, fg=C_REPEAT_FG,
            relief="flat", font=FONT_SMALL, cursor="hand2", padx=6,
            command=self._toggle_repeat_filter)
        self._repeat_btn.pack(side="left", padx=2)

        sf = tk.Frame(self, bg=C_PANEL, pady=6, padx=12)
        sf.pack(fill="x")

        tk.Label(sf, text="🔍", bg=C_PANEL, fg=C_ACCENT,
                 font=FONT_MAIN).pack(side="left")
        self.v_search = tk.StringVar()
        self.v_search.trace_add("write", lambda *_: self._apply_filter())
        styled_entry(sf, textvariable=self.v_search, width=24).pack(
            side="left", padx=4)
        tk.Label(sf, text="(any field)", bg=C_PANEL, fg=C_MUTED,
                 font=FONT_SMALL).pack(side="left")

        tk.Label(sf, text="│", bg=C_PANEL, fg=C_BORDER,
                 font=FONT_MAIN).pack(side="left", padx=4)
        tk.Label(sf, text="Serial #:", bg=C_PANEL, fg=C_MUTED,
                 font=FONT_SMALL).pack(side="left", padx=(2,2))
        self.v_filter_serial = tk.StringVar()
        self.v_filter_serial.trace_add("write", lambda *_: self._apply_filter())
        styled_entry(sf, textvariable=self.v_filter_serial, width=18).pack(
            side="left", padx=(0,8))

        tk.Label(sf, text="Project:", bg=C_PANEL, fg=C_MUTED,
                 font=FONT_SMALL).pack(side="left", padx=(4,2))
        self.v_filter_proj = tk.StringVar(value="All")
        proj_cb = styled_combo(sf, values=["All"] + self.cfg.get_projects(),
                               textvariable=self.v_filter_proj, width=16)
        proj_cb.pack(side="left", padx=2)
        self.v_filter_proj.trace_add("write", lambda *_: self._apply_filter())

        tk.Label(sf, text="Type:", bg=C_PANEL, fg=C_MUTED,
                 font=FONT_SMALL).pack(side="left", padx=(8,2))
        self.v_filter_cat = tk.StringVar(value="All")
        cat_cb = styled_combo(sf, values=["All"] + RMA_TYPES,
                              textvariable=self.v_filter_cat, width=14)
        cat_cb.pack(side="left", padx=2)
        self.v_filter_cat.trace_add("write", lambda *_: self._apply_filter())

        tk.Label(sf, text="│", bg=C_PANEL, fg=C_BORDER,
                 font=FONT_MAIN).pack(side="left", padx=4)
        tk.Label(sf, text="Recv From:", bg=C_PANEL, fg=C_MUTED,
                 font=FONT_SMALL).pack(side="left", padx=(2,2))
        self.filter_date_from = DatePickerEntry(sf, allow_blank=True)
        self.filter_date_from.pack(side="left", padx=(0,4))
        tk.Label(sf, text="To:", bg=C_PANEL, fg=C_MUTED,
                 font=FONT_SMALL).pack(side="left", padx=(0,2))
        self.filter_date_to = DatePickerEntry(sf, allow_blank=True)
        self.filter_date_to.pack(side="left", padx=(0,4))

        self.filter_date_from._var.trace_add("write", lambda *_: self._apply_filter())
        self.filter_date_to._var.trace_add("write",   lambda *_: self._apply_filter())

        def _clear_dates():
            if not self.filter_date_from._is_blank:
                self.filter_date_from._clear()
            if not self.filter_date_to._is_blank:
                self.filter_date_to._clear()
            self._apply_filter()

        tk.Button(sf, text="✕ Clear", bg=C_PANEL, fg=C_MUTED,
                  relief="flat", font=FONT_SMALL, cursor="hand2",
                  command=_clear_dates).pack(side="left", padx=(0,4))

        self.lbl_count = tk.Label(sf, text="", bg=C_PANEL,
                                  fg=C_MUTED, font=FONT_SMALL)
        self.lbl_count.pack(side="right")

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("RMA.Treeview",
                        background=C_CARD, foreground=C_TEXT,
                        fieldbackground=C_CARD, rowheight=24, font=FONT_MAIN)
        style.configure("RMA.Treeview.Heading",
                        background=C_HEADER, foreground=C_ACCENT,
                        relief="flat", font=FONT_BOLD)
        style.map("RMA.Treeview",
                  background=[("selected", C_SEL)],
                  foreground=[("selected", C_TEXT)])

        tv_frame = tk.Frame(self, bg=C_BG)
        tv_frame.pack(fill="both", expand=True, padx=8, pady=(0,8))

        col_ids = [c[0] for c in self.COLS]
        self.tree = ttk.Treeview(tv_frame, columns=col_ids, show="headings",
                                  style="RMA.Treeview")
        for label, key, width in self.COLS:
            self.tree.heading(label, text=label,
                              command=lambda k=label: self._sort(k))
            self.tree.column(label, width=width, minwidth=50, anchor="w")

        vsb = ttk.Scrollbar(tv_frame, orient="vertical",   command=self.tree.yview)
        hsb = ttk.Scrollbar(tv_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tv_frame.rowconfigure(0, weight=1)
        tv_frame.columnconfigure(0, weight=1)

        for st, color in STATUS_COLOR.items():
            self.tree.tag_configure(st, foreground=color)
        self.tree.tag_configure("overdue_row",
                                background="#3B1A1A", foreground=C_DANGER)
        self.tree.tag_configure("repeat_offender",
                                background=C_REPEAT_BG, foreground=C_REPEAT_FG)

        self.tree.bind("<Double-1>", lambda e: self._edit_rma())
        self.tree.bind("<Button-3>", self._show_context_menu)

        self._context_menu = tk.Menu(
            self, tearoff=0, bg=C_CARD, fg=C_TEXT,
            activebackground=C_ACCENT, activeforeground=C_HEADER,
            relief="flat", font=FONT_MAIN)
        self._context_menu.add_command(label="✏  Edit RMA",       command=self._edit_rma)
        self._context_menu.add_command(label="📋  Clone RMA",     command=self._clone_rma)
        self._context_menu.add_command(label="📋  Board History", command=self._board_history)
        self._context_menu.add_separator()
        self._context_menu.add_command(label="🗑  Delete RMA",    command=self._delete_rma)

        self._sort_col    = None
        self._sort_rev    = False
        self._status_filter = "All"
        self.refresh()

    # ── Data ──────────────────────────────────────────────────────────────────
    def refresh(self):
        self._all_rows       = self.db.fetchall(
            "SELECT * FROM rma ORDER BY received_date DESC, id DESC")
        self._repeat_serials = self.db.repeat_offender_serials()
        rcount = len(self._repeat_serials)
        badge  = f"🔁 Repeat Offender ({rcount})" if rcount else "🔁 Repeat Offender"
        self._repeat_btn.config(text=badge)
        self._apply_filter()

    def _apply_filter(self):
        term   = self.v_search.get().lower()
        serial = self.v_filter_serial.get().lower().strip()
        proj   = self.v_filter_proj.get()
        cat    = self.v_filter_cat.get()
        today  = date.today()
        d_from = self.filter_date_from.get()
        d_to   = self.filter_date_to.get()

        filtered = []
        for r in self._all_rows:
            if self._status_filter != "All" and r["status"] != self._status_filter:
                continue
            if proj != "All" and r["project"] != proj:
                continue
            if cat != "All" and r["rma_category"] != cat:
                continue
            if serial and serial not in str(r["serial_number"]).lower():
                continue
            if self._repeat_filter_active and \
                    r["serial_number"] not in self._repeat_serials:
                continue
            if term and not any(
                term in str(r.get(k, "")).lower() for k in r
            ):
                continue
            rd = str(r["received_date"])
            if d_from and rd < d_from:
                continue
            if d_to and rd > d_to:
                continue
            filtered.append(r)

        for item in self.tree.get_children():
            self.tree.delete(item)

        active_statuses = {"Open","In-Repair","Pending Shipment","On-Hold"}

        def _row_age(r) -> int:
            try:
                recv = datetime.strptime(r["received_date"], "%Y-%m-%d").date()
                if r["shipment_date"] and r["shipment_date"] != "N/A":
                    return 0
                return (today - recv).days
            except Exception:
                return 0

        def _is_overdue(r, age) -> bool:
            return r["status"] in active_statuses and age > TAT_DANGER_DAYS

        def _is_repeat(r) -> bool:
            return r["serial_number"] in self._repeat_serials

        filtered.sort(key=lambda r: (
            0 if _is_overdue(r, _row_age(r)) else (1 if _is_repeat(r) else 2),
            -_row_age(r) if _is_overdue(r, _row_age(r)) else 0,
        ))

        for r in filtered:
            try:
                recv = datetime.strptime(r["received_date"], "%Y-%m-%d").date()
                if r["shipment_date"] and r["shipment_date"] != "N/A":
                    ship = datetime.strptime(r["shipment_date"], "%Y-%m-%d").date()
                    tat  = str((ship - recv).days)
                else:
                    tat  = str((today - recv).days) + "*"
            except Exception:
                tat = "—"

            age = _row_age(r)
            if _is_overdue(r, age):
                row_tags = ("overdue_row",)
            elif _is_repeat(r):
                row_tags = ("repeat_offender",)
            else:
                row_tags = (r["status"],)

            self.tree.insert("", "end", iid=str(r["id"]), values=(
                r["rma_number"], r["ncr_number"], r["rma_category"],
                r["project"], r["model"], r["revision"], r["serial_number"],
                r["received_date"], r["received_by"], r["status"], tat,
                f"{r['cost_estimate']:.2f}",
            ), tags=row_tags)

        self.lbl_count.config(text=f"{len(filtered)} record(s)")

    def _quick_filter(self, status: str):
        self._status_filter = status
        self._repeat_filter_active = False
        self._repeat_btn.config(bg=C_HEADER)
        self._apply_filter()

    def _toggle_repeat_filter(self):
        self._repeat_filter_active = not self._repeat_filter_active
        if self._repeat_filter_active:
            self._status_filter = "All"
            self._repeat_btn.config(bg=C_REPEAT_BG, fg=C_REPEAT_FG)
        else:
            self._repeat_btn.config(bg=C_HEADER, fg=C_REPEAT_FG)
        self._apply_filter()

    def _sort(self, col: str):
        items   = [(self.tree.set(k, col), k) for k in self.tree.get_children("")]
        reverse = (col == self._sort_col and not self._sort_rev)
        items.sort(key=lambda x: x[0], reverse=reverse)
        for idx, (_, k) in enumerate(items):
            self.tree.move(k, "", idx)
        self._sort_col = col
        self._sort_rev = reverse

    def _selected_id(self) -> Optional[int]:
        sel = self.tree.selection()
        if not sel:
            return None
        return int(sel[0])

    # ── Actions ───────────────────────────────────────────────────────────────
    def _new_rma(self):
        RMAForm(self, self.db, self.cfg, self.user, on_save=self.refresh)

    def _edit_rma(self):
        rma_id = self._selected_id()
        if not rma_id:
            messagebox.showinfo("No Selection", "Select an RMA row to edit.")
            return
        RMAForm(self, self.db, self.cfg, self.user,
                rma_id=rma_id, on_save=self.refresh)

    def _board_history(self):
        rma_id = self._selected_id()
        if not rma_id:
            messagebox.showinfo("No Selection",
                                "Select an RMA to view board history.")
            return
        row = self.db.fetchone(
            "SELECT serial_number FROM rma WHERE id=?", (rma_id,))
        if row:
            BoardHistoryViewer(self, self.db, row["serial_number"])

    def _bulk_shipment(self):
        BulkShipmentDialog(self, self.db, self.user, on_done=self.refresh)

    def _delete_rma(self):
        rma_id = self._selected_id()
        if not rma_id:
            return
        row = self.db.fetchone(
            "SELECT rma_number FROM rma WHERE id=?", (rma_id,))
        if not row:
            return
        if not messagebox.askyesno(
            "Confirm Delete",
            f"Permanently delete {row['rma_number']}?\nThis cannot be undone.",
        ):
            return
        self.db.execute("DELETE FROM rma WHERE id=?", (rma_id,))
        log.warning(f"Deleted RMA id={rma_id} by {self.user}")
        self.refresh()

    def _export(self):
        ReportDialog(self, self.db, self.cfg)

    def _show_context_menu(self, event):
        iid = self.tree.identify_row(event.y)
        if iid:
            self.tree.selection_set(iid)
            try:
                self._context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self._context_menu.grab_release()

    def _clone_rma(self):
        rma_id = self._selected_id()
        if not rma_id:
            messagebox.showinfo("No Selection", "Select an RMA row to clone.")
            return
        row = self.db.fetchone("SELECT * FROM rma WHERE id=?", (rma_id,))
        if not row:
            return
        clone_data = {
            "project":       row["project"],
            "model":         row["model"],
            "revision":      row["revision"],
            "rma_category":  row["rma_category"],
            "failure_code":  row["failure_code"],
            "failure_desc":  row["failure_desc"],
            "rework_notice": row["rework_notice"],
        }
        RMAForm(self, self.db, self.cfg, self.user,
                clone_data=clone_data, on_save=self.refresh)


# ══════════════════════════════════════════════════════════════════════════════
#  DASHBOARD TAB
# ══════════════════════════════════════════════════════════════════════════════

class DashboardTab(tk.Frame):
    KPI_DEFS = [
        ("total",      "Total",         C_TEXT),
        ("open",       "Open",          C_ACCENT),
        ("in_repair",  "In-Repair",     C_WARN),
        ("pending",    "Pending Ship.", "#9C27B0"),
        ("closed",     "Closed",        C_SUCCESS),
        ("on_hold",    "On-Hold",       C_DANGER),
        ("avg_tat",    "Avg TAT (d)",   C_MUTED),
        ("total_cost", "Cost (USD)",    C_WARN),
    ]

    def __init__(self, parent, db: DB, cfg: ConfigManager):
        super().__init__(parent, bg=C_BG)
        self.db  = db
        self.cfg = cfg
        self._proj_kpi_labels: dict = {}
        self._pareto_data: list = []
        self._build()

    def _build(self):
        hdr_row = tk.Frame(self, bg=C_BG)
        hdr_row.pack(fill="x", padx=16, pady=(12,4))
        tk.Label(hdr_row, text="RMA Operations Dashboard",
                 bg=C_BG, fg=C_TEXT, font=FONT_TITLE).pack(side="left")
        styled_button(hdr_row, "⟳  Refresh", command=self.refresh,
                      style="muted").pack(side="right")

        # Date range bar
        dr_row = tk.Frame(self, bg=C_PANEL, pady=6, padx=12)
        dr_row.pack(fill="x", padx=0, pady=(0,4))
        tk.Label(dr_row, text="📅  Date Range:", bg=C_PANEL, fg=C_MUTED,
                 font=FONT_SMALL).pack(side="left", padx=(4,6))
        for label, days in [("Last 7d",7),("Last 30d",30),
                             ("Last 90d",90),("This Year",0),("All Time",-1)]:
            tk.Button(dr_row, text=label, bg=C_CARD, fg=C_ACCENT,
                      relief="flat", font=FONT_SMALL, cursor="hand2", padx=6,
                      command=lambda d=days: self._preset_range(d)
                      ).pack(side="left", padx=2)
        tk.Label(dr_row, text="│", bg=C_PANEL, fg=C_BORDER,
                 font=FONT_MAIN).pack(side="left", padx=4)
        tk.Label(dr_row, text="From:", bg=C_PANEL, fg=C_MUTED,
                 font=FONT_SMALL).pack(side="left", padx=(2,2))
        self._dash_date_from = DatePickerEntry(dr_row, allow_blank=True)
        self._dash_date_from.pack(side="left", padx=(0,4))
        tk.Label(dr_row, text="To:", bg=C_PANEL, fg=C_MUTED,
                 font=FONT_SMALL).pack(side="left", padx=(0,2))
        self._dash_date_to = DatePickerEntry(dr_row, allow_blank=True)
        self._dash_date_to.pack(side="left", padx=(0,6))
        styled_button(dr_row, "Apply", command=self.refresh,
                      style="primary").pack(side="left", padx=2)
        self._dash_range_lbl = tk.Label(dr_row, text="(All time)",
                                        bg=C_PANEL, fg=C_MUTED, font=FONT_SMALL)
        self._dash_range_lbl.pack(side="left", padx=8)
        self._dash_date_from._var.trace_add("write", lambda *_: self._update_range_label())
        self._dash_date_to._var.trace_add("write",   lambda *_: self._update_range_label())

        # Overall KPI banner
        tk.Label(self, text="OVERALL", bg=C_BG, fg=C_MUTED,
                 font=FONT_SMALL, padx=16).pack(anchor="w")
        overall_outer  = tk.Frame(self, bg=C_BG)
        overall_outer.pack(fill="x", padx=16, pady=(2,6))
        overall_canvas = tk.Canvas(overall_outer, bg=C_BG,
                                   highlightthickness=0, height=74)
        overall_hsb    = ttk.Scrollbar(overall_outer, orient="horizontal",
                                       command=overall_canvas.xview)
        overall_inner  = tk.Frame(overall_canvas, bg=C_BG)
        overall_inner.bind(
            "<Configure>",
            lambda e: overall_canvas.configure(
                scrollregion=overall_canvas.bbox("all")))
        overall_canvas.create_window((0,0), window=overall_inner, anchor="nw")
        overall_canvas.configure(xscrollcommand=overall_hsb.set)
        overall_canvas.pack(fill="x")
        overall_hsb.pack(fill="x")

        self._kpi_labels: dict = {}
        for key, label, color in self.KPI_DEFS:
            fr = tk.Frame(overall_inner, bg=C_CARD, padx=14, pady=10)
            fr.pack(side="left", padx=3, fill="y")
            num = tk.Label(fr, text="—", bg=C_CARD, fg=color,
                           font=("Segoe UI", 18, "bold"))
            num.pack()
            tk.Label(fr, text=label, bg=C_CARD, fg=C_MUTED,
                     font=FONT_SMALL, wraplength=80).pack()
            self._kpi_labels[key] = num

        tk.Frame(self, bg=C_BORDER, height=1).pack(fill="x", padx=16, pady=(2,4))

        # By Project
        tk.Label(self, text="BY PROJECT", bg=C_BG, fg=C_MUTED,
                 font=FONT_SMALL, padx=16).pack(anchor="w")
        proj_outer = tk.Frame(self, bg=C_BG)
        proj_outer.pack(fill="x", padx=16, pady=(2,4))
        self._proj_canvas = tk.Canvas(proj_outer, bg=C_BG,
                                      highlightthickness=0, height=220)
        proj_vsb = ttk.Scrollbar(proj_outer, orient="vertical",
                                 command=self._proj_canvas.yview)
        self._proj_inner = tk.Frame(self._proj_canvas, bg=C_BG)
        self._proj_inner.bind(
            "<Configure>",
            lambda e: self._proj_canvas.configure(
                scrollregion=self._proj_canvas.bbox("all")))
        self._proj_canvas.create_window((0,0), window=self._proj_inner, anchor="nw")
        self._proj_canvas.configure(yscrollcommand=proj_vsb.set)
        self._proj_canvas.pack(side="left", fill="x", expand=True)
        proj_vsb.pack(side="right", fill="y")
        self._proj_canvas.bind(
            "<MouseWheel>",
            lambda e: self._proj_canvas.yview_scroll(-1*(e.delta//120), "units"))

        tk.Frame(self, bg=C_BORDER, height=1).pack(fill="x", padx=16, pady=(2,4))

        # Pareto
        pareto_hdr = tk.Frame(self, bg=C_BG)
        pareto_hdr.pack(fill="x", padx=16, pady=(2,0))
        tk.Label(pareto_hdr, text="TOP FAILURE CODES  (Pareto)",
                 bg=C_BG, fg=C_MUTED, font=FONT_SMALL).pack(side="left")
        self._pareto_canvas = tk.Canvas(self, bg=C_BG, highlightthickness=0,
                                        height=160, relief="flat")
        self._pareto_canvas.pack(fill="x", padx=16, pady=(2,4))
        self._pareto_canvas.bind("<Configure>", lambda e: self._draw_pareto())

        self.refresh()

    # ── Date range helpers ────────────────────────────────────────────────────
    def _preset_range(self, days: int):
        today = date.today()
        if days == -1:
            if not self._dash_date_from._is_blank:
                self._dash_date_from._clear()
            if not self._dash_date_to._is_blank:
                self._dash_date_to._clear()
        elif days == 0:
            self._dash_date_from.set(date(today.year, 1, 1).strftime("%Y-%m-%d"))
            self._dash_date_to.set(today.strftime("%Y-%m-%d"))
        else:
            self._dash_date_from.set((today - timedelta(days=days)).strftime("%Y-%m-%d"))
            self._dash_date_to.set(today.strftime("%Y-%m-%d"))
        self.refresh()

    def _update_range_label(self):
        d_from = self._dash_date_from.get()
        d_to   = self._dash_date_to.get()
        if not d_from and not d_to:
            self._dash_range_lbl.config(text="(All time)")
        elif d_from and not d_to:
            self._dash_range_lbl.config(text=f"From {d_from}")
        elif not d_from and d_to:
            self._dash_range_lbl.config(text=f"Up to {d_to}")
        else:
            self._dash_range_lbl.config(text=f"{d_from}  →  {d_to}")

    def _dash_where(self) -> tuple:
        d_from = self._dash_date_from.get()
        d_to   = self._dash_date_to.get()
        clauses, params = [], []
        if d_from:
            clauses.append("received_date >= ?"); params.append(d_from)
        if d_to:
            clauses.append("received_date <= ?"); params.append(d_to)
        fragment = " AND ".join(clauses) if clauses else ""
        return fragment, params

    def _dash_sql(self, base_sql: str, extra_params: list = None) -> tuple:
        fragment, params = self._dash_where()
        all_params = list(params)
        if extra_params:
            all_params.extend(extra_params)
        if not fragment:
            return base_sql, all_params
        connector = " AND " if "WHERE" in base_sql.upper() else " WHERE "
        return base_sql + connector + fragment, all_params

    # ── Refresh ───────────────────────────────────────────────────────────────
    def refresh(self):
        today = date.today()
        self._update_range_label()

        def _avg_tat(rows) -> str:
            tats = []
            for r in rows:
                try:
                    d1 = datetime.strptime(r["received_date"], "%Y-%m-%d").date()
                    d2 = datetime.strptime(r["shipment_date"], "%Y-%m-%d").date()
                    tats.append((d2 - d1).days)
                except Exception:
                    pass
            return f"{sum(tats)/len(tats):.1f}" if tats else "—"

        # Overall KPIs — use [status] to avoid reserved-word conflict
        total_sql, total_params = self._dash_sql("SELECT COUNT(*) as c FROM rma")
        self._kpi_labels["total"].config(
            text=str(self.db.fetchone(total_sql, total_params)["c"]))

        dr_frag, dr_params = self._dash_where()
        dr_and = (" AND " + dr_frag) if dr_frag else ""

        for key, status in [
            ("open",     "Open"),
            ("in_repair","In-Repair"),
            ("pending",  "Pending Shipment"),
            ("closed",   "Closed"),
            ("on_hold",  "On-Hold"),
        ]:
            cnt = self.db.fetchone(
                f"SELECT COUNT(*) as c FROM rma WHERE [status]=?{dr_and}",
                [status] + dr_params)["c"]
            self._kpi_labels[key].config(text=str(cnt))

        closed_sql = (
            "SELECT received_date, shipment_date FROM rma "
            "WHERE [status]='Closed' AND shipment_date != 'N/A'"
            + ((" AND " + dr_frag) if dr_frag else "")
        )
        closed_all = self.db.fetchall(closed_sql, dr_params)
        self._kpi_labels["avg_tat"].config(text=_avg_tat(closed_all))

        cost_sql, cost_params = self._dash_sql(
            "SELECT SUM(cost_estimate) as s FROM rma")
        cost_row = self.db.fetchone(cost_sql, cost_params)
        self._kpi_labels["total_cost"].config(
            text=f"{(cost_row['s'] or 0):,.2f}")

        # By Project cards
        for w in self._proj_inner.winfo_children():
            w.destroy()
        self._proj_kpi_labels.clear()

        proj_base = """
            SELECT project,
                   COUNT(*) AS total,
                   SUM(CASE WHEN [status]='Open'             THEN 1 ELSE 0 END) AS "open",
                   SUM(CASE WHEN [status]='In-Repair'        THEN 1 ELSE 0 END) AS in_repair,
                   SUM(CASE WHEN [status]='Pending Shipment' THEN 1 ELSE 0 END) AS pending,
                   SUM(CASE WHEN [status]='Closed'           THEN 1 ELSE 0 END) AS closed,
                   SUM(CASE WHEN [status]='On-Hold'          THEN 1 ELSE 0 END) AS on_hold,
                   SUM(CASE WHEN [status]='Cancelled'        THEN 1 ELSE 0 END) AS cancelled,
                   SUM(cost_estimate) AS cost
            FROM rma"""
        proj_sql, proj_params = self._dash_sql(proj_base)
        proj_sql += " GROUP BY project ORDER BY total DESC"
        proj_rows = self.db.fetchall(proj_sql, proj_params)

        tat_sql = (
            "SELECT project, received_date, shipment_date FROM rma "
            "WHERE [status]='Closed' AND shipment_date != 'N/A'"
            + ((" AND " + dr_frag) if dr_frag else "")
        )
        tat_rows = self.db.fetchall(tat_sql, dr_params)
        tat_by_proj: dict = {}
        for r in tat_rows:
            tat_by_proj.setdefault(r["project"], []).append(r)

        for idx, pr in enumerate(proj_rows):
            row_f = idx // 2
            col_f = idx % 2
            proj_tat_rows = tat_by_proj.get(pr["project"], [])
            cost_val      = pr["cost"] or 0
            kpi_values = {
                "total":      str(pr["total"]),
                "open":       str(pr["open"]      or 0),
                "in_repair":  str(pr["in_repair"] or 0),
                "pending":    str(pr["pending"]   or 0),
                "closed":     str(pr["closed"]    or 0),
                "on_hold":    str(pr["on_hold"]   or 0),
                "avg_tat":    _avg_tat(proj_tat_rows),
                "total_cost": f"{cost_val:,.2f}",
            }
            self._make_project_card(
                self._proj_inner, pr["project"], kpi_values, row_f, col_f)

        if not proj_rows:
            tk.Label(self._proj_inner, text="No project data yet.",
                     bg=C_BG, fg=C_MUTED, font=FONT_MAIN,
                     padx=20, pady=20).grid(row=0, column=0)

        # Pareto — SQL Server: TOP N instead of LIMIT N
        pareto_sql = (
            "SELECT TOP 8 failure_code, COUNT(*) as cnt FROM rma "
            "WHERE failure_code NOT IN ('N/A', '', 'None')"
            + ((" AND " + dr_frag) if dr_frag else "")
            + " GROUP BY failure_code ORDER BY cnt DESC"
        )
        pareto_rows = self.db.fetchall(pareto_sql, dr_params)
        self._pareto_data = [(r["failure_code"], r["cnt"]) for r in pareto_rows]
        self._draw_pareto()

    # ── Pareto chart ──────────────────────────────────────────────────────────
    def _draw_pareto(self):
        cv = self._pareto_canvas
        cv.delete("all")
        data = getattr(self, "_pareto_data", [])
        if not data:
            cv.create_text(20, 80, anchor="w",
                           text="No failure code data yet.",
                           fill=C_MUTED, font=FONT_SMALL)
            return

        W = cv.winfo_width() or 800
        H = cv.winfo_height() or 160
        PAD_LEFT  = 20
        PAD_RIGHT = 10
        PAD_TOP   = 10
        PAD_BOT   = 10
        LABEL_W   = 180
        COUNT_W   = 40
        bar_area_w = W - PAD_LEFT - LABEL_W - COUNT_W - PAD_RIGHT
        n = len(data)
        if n == 0 or bar_area_w <= 0:
            return

        total_count = max(r[1] for r in data) or 1
        row_h       = (H - PAD_TOP - PAD_BOT) / n
        BAR_COLORS  = ["#00B4D8","#0096B4","#0077B6","#005A8E",
                       "#48CAE4","#90E0EF","#ADE8F4","#CAF0F8"]
        cumulative  = 0
        total_all   = sum(r[1] for r in data)
        pareto_pts  = []

        for i, (code, cnt) in enumerate(data):
            y1 = PAD_TOP + i * row_h + 2
            y2 = PAD_TOP + (i + 1) * row_h - 2
            bar_w = int(bar_area_w * (cnt / total_count))
            x0    = PAD_LEFT + LABEL_W
            cv.create_rectangle(x0, y1, x0 + bar_area_w, y2,
                                 fill=C_CARD, outline="")
            cv.create_rectangle(x0, y1, x0 + bar_w, y2,
                                 fill=BAR_COLORS[i % len(BAR_COLORS)], outline="")
            label = code if len(code) <= 24 else code[:21] + "…"
            cy    = (y1 + y2) / 2
            cv.create_text(PAD_LEFT, cy, anchor="w",
                           text=label, fill=C_TEXT, font=FONT_SMALL)
            cv.create_text(x0 + bar_area_w + 4, cy, anchor="w",
                           text=str(cnt), fill=C_MUTED, font=FONT_SMALL)
            cumulative += cnt
            pct = cumulative / total_all if total_all else 0
            pareto_pts.append((x0 + int(bar_area_w * pct), cy))

        if len(pareto_pts) >= 2:
            for j in range(len(pareto_pts) - 1):
                cv.create_line(pareto_pts[j][0], pareto_pts[j][1],
                               pareto_pts[j+1][0], pareto_pts[j+1][1],
                               fill=C_WARN, width=2)
        for px, py in pareto_pts:
            cv.create_oval(px-3, py-3, px+3, py+3, fill=C_WARN, outline=C_BG)

    # ── Project card ──────────────────────────────────────────────────────────
    def _make_project_card(self, parent, project_name: str,
                           kpi_values: dict, grid_row: int, grid_col: int):
        card = tk.Frame(parent, bg=C_CARD, padx=10, pady=8,
                        highlightthickness=1, highlightbackground=C_BORDER)
        card.grid(row=grid_row, column=grid_col, padx=6, pady=5, sticky="nsew")
        parent.columnconfigure(0, weight=1)
        parent.columnconfigure(1, weight=1)

        hdr = tk.Frame(card, bg=C_CARD)
        hdr.grid(row=0, column=0, columnspan=8, sticky="w", pady=(0,6))
        tk.Frame(hdr, bg=C_ACCENT, width=3).pack(side="left", fill="y", padx=(0,6))
        tk.Label(hdr, text=project_name, bg=C_CARD, fg=C_ACCENT,
                 font=FONT_BOLD).pack(side="left")

        tile_layout = [
            (1,0,"total"),(1,1,"open"),(1,2,"in_repair"),(1,3,"pending"),
            (2,0,"closed"),(2,1,"on_hold"),(2,2,"avg_tat"),(2,3,"total_cost"),
        ]
        for t_row, t_col, key in tile_layout:
            _, label, color = next(d for d in self.KPI_DEFS if d[0] == key)
            tile = tk.Frame(card, bg=C_PANEL, padx=8, pady=5)
            tile.grid(row=t_row, column=t_col, padx=2, pady=2, sticky="nsew")
            card.columnconfigure(t_col, weight=1)
            tk.Label(tile, text=kpi_values.get(key, "—"),
                     bg=C_PANEL, fg=color,
                     font=("Segoe UI", 12, "bold")).pack()
            tk.Label(tile, text=label, bg=C_PANEL, fg=C_MUTED,
                     font=FONT_SMALL, wraplength=70).pack()


# ══════════════════════════════════════════════════════════════════════════════
#  CONFIG TAB
# ══════════════════════════════════════════════════════════════════════════════

class ConfigTab(tk.Frame):
    def __init__(self, parent, cfg: ConfigManager):
        super().__init__(parent, bg=C_BG)
        self.cfg = cfg
        self._build()

    def _build(self):
        tk.Label(self, text="Configuration Status", bg=C_BG, fg=C_TEXT,
                 font=FONT_TITLE, pady=12).pack(anchor="w", padx=16)

        fr_paths = tk.LabelFrame(self, text="File Locations", bg=C_BG, fg=C_MUTED,
                                 font=FONT_BOLD, padx=8, pady=8)
        fr_paths.pack(fill="x", padx=16, pady=4)
        for label, path in [
            ("Config dir:", CONFIG_DIR),
            ("DB config:",  DB_CONFIG),
            ("Attachments:", LOCAL_ATTACHMENTS_DIR),
        ]:
            r = tk.Frame(fr_paths, bg=C_BG)
            r.pack(fill="x", pady=2)
            tk.Label(r, text=label, bg=C_BG, fg=C_MUTED, font=FONT_MAIN,
                     width=14, anchor="w").pack(side="left")
            tk.Label(r, text=str(path), bg=C_BG, fg=C_TEXT,
                     font=FONT_MONO).pack(side="left")
            exists = path.exists()
            tk.Label(r, text="✔ exists" if exists else "✗ missing",
                     bg=C_BG,
                     fg=C_SUCCESS if exists else C_DANGER,
                     font=FONT_SMALL).pack(side="right")

        if self.cfg.errors:
            fr_err = tk.LabelFrame(self, text="⚠ Configuration Warnings",
                                   bg=C_BG, fg=C_WARN, font=FONT_BOLD,
                                   padx=8, pady=8)
            fr_err.pack(fill="x", padx=16, pady=4)
            for e in self.cfg.errors:
                tk.Label(fr_err, text=f"• {e}", bg=C_BG, fg=C_WARN,
                         font=FONT_SMALL, anchor="w").pack(fill="x")

        fr_stat = tk.LabelFrame(self, text="Loaded Data", bg=C_BG, fg=C_MUTED,
                                font=FONT_BOLD, padx=8, pady=8)
        fr_stat.pack(fill="x", padx=16, pady=4)
        stats = [
            ("Projects",       len(self.cfg.projects)),
            ("Total Models",   sum(len(v) for v in self.cfg.projects.values())),
            ("Users",          len(self.cfg.users)),
            ("Failure Codes",  len(self.cfg.failure_codes)),
            ("Standard Costs", len(self.cfg.standard_costs)),
            ("Report Columns", len(self.cfg.report_columns)),
        ]
        for i, (label, val) in enumerate(stats):
            tk.Label(fr_stat, text=f"{label}:  {val}", bg=C_BG,
                     fg=C_TEXT, font=FONT_MAIN).grid(
                row=i//3, column=i%3, sticky="w", padx=20, pady=2)

        fr_act = tk.Frame(self, bg=C_BG, pady=8)
        fr_act.pack(fill="x", padx=16)
        styled_button(fr_act, "⟳  Reload Config",
                      command=self._reload, style="primary").pack(side="left", padx=4)
        styled_button(fr_act, "📂  Open Config Dir",
                      command=lambda: (
                          os.startfile(str(CONFIG_DIR)) if os.name == "nt"
                          else os.system(f"xdg-open '{CONFIG_DIR}'")),
                      style="muted").pack(side="left", padx=4)
        styled_button(fr_act, "📂  Open Attachments Dir",
                      command=lambda: (
                          os.startfile(str(LOCAL_ATTACHMENTS_DIR)) if os.name == "nt"
                          else os.system(f"xdg-open '{LOCAL_ATTACHMENTS_DIR}'")),
                      style="muted").pack(side="left", padx=4)

    def _reload(self):
        self.cfg.reload()
        messagebox.showinfo("Config Reloaded",
                            f"Config reloaded.\n"
                            f"{len(self.cfg.projects)} projects, "
                            f"{len(self.cfg.users)} users.")


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN APPLICATION
# ══════════════════════════════════════════════════════════════════════════════

class RMAApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.withdraw()

        if not HAS_XLSX:
            messagebox.showerror(
                "Missing Dependency",
                "openpyxl is required.\n\n"
                "Install with:\n  pip install openpyxl tkcalendar pillow pyodbc"
            )
            sys.exit(1)

        if not HAS_PYODBC:
            messagebox.showerror(
                "Missing Dependency",
                "pyodbc is required for SQL Server.\n\n"
                "Install with:\n  pip install pyodbc"
            )
            sys.exit(1)

        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        BOM_DIR.mkdir(parents=True, exist_ok=True)
        LOCAL_ATTACHMENTS_DIR.mkdir(parents=True, exist_ok=True)

        self.cfg = ConfigManager()

        try:
            self.db = DB()
        except (FileNotFoundError, ConnectionError) as exc:
            messagebox.showerror("Database Error", str(exc))
            sys.exit(1)

        # No login — use Windows / OS username
        self.current_user = getpass.getuser()

        self._build_main()
        self.deiconify()

    def _build_main(self):
        self.title(f"{APP_TITLE}  –  {self.current_user}")
        self.configure(bg=C_BG)
        self.geometry("1280x800")
        self.minsize(900, 600)

        bar = tk.Frame(self, bg=C_HEADER, pady=6)
        bar.pack(fill="x")
        tk.Label(bar, text="⚡  RMA MANAGEMENT SYSTEM", bg=C_HEADER,
                 fg=C_ACCENT, font=FONT_TITLE, padx=14).pack(side="left")
        tk.Label(bar,
                 text=f"User: {self.current_user}   |   "
                      f"{datetime.now().strftime('%A, %d %b %Y')}",
                 bg=C_HEADER, fg=C_MUTED, font=FONT_SMALL).pack(
                     side="right", padx=14)

        nb_style = ttk.Style()
        nb_style.configure("Dark.TNotebook", background=C_BG, borderwidth=0)
        nb_style.configure("Dark.TNotebook.Tab",
                           background=C_PANEL, foreground=C_MUTED,
                           padding=[14,6], font=FONT_BOLD)
        nb_style.map("Dark.TNotebook.Tab",
                     background=[("selected", C_CARD)],
                     foreground=[("selected", C_ACCENT)])

        self.nb = ttk.Notebook(self, style="Dark.TNotebook")
        self.nb.pack(fill="both", expand=True, padx=4, pady=4)

        self.dash     = DashboardTab(self.nb, self.db, self.cfg)
        self.rma_list = RMAListView(self.nb, self.db, self.cfg, self.current_user)
        self.tat      = TATDashboard(self.nb, self.db)
        self.audit    = AuditLogViewer(self.nb, self.db)
        self.conf_tab = ConfigTab(self.nb, self.cfg)

        self.nb.add(self.dash,     text="  📊  Dashboard  ")
        self.nb.add(self.rma_list, text="  📋  RMA Records  ")
        self.nb.add(self.tat,      text="  ⏱  TAT  ")
        self.nb.add(self.audit,    text="  🔍  Audit Log  ")
        self.nb.add(self.conf_tab, text="  ⚙  Config  ")

        self.nb.bind("<<NotebookTabChanged>>", self._on_tab_switch)

        self.status_var = tk.StringVar(value="Ready")
        tk.Label(self, textvariable=self.status_var, bg=C_HEADER,
                 fg=C_MUTED, font=FONT_SMALL, anchor="w", padx=8).pack(
                     fill="x", side="bottom")

        log.info(f"App started – user: {self.current_user}")
        if self.cfg.errors:
            self.status_var.set(
                f"⚠ {len(self.cfg.errors)} config warning(s) – see Config tab")

    def _on_tab_switch(self, event):
        tab = self.nb.tab(self.nb.select(), "text").strip()
        if "TAT"       in tab: self.tat.refresh()
        elif "Dashboard" in tab: self.dash.refresh()
        elif "Audit"     in tab: self.audit.refresh()


# ══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    app = RMAApp()
    app.mainloop()
