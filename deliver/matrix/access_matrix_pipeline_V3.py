#!/usr/bin/env python3
"""
Access Matrix Pipeline — six stages, one script.

Every pipeline run creates a single timestamped output folder under
OUTPUT_PARENT (e.g. 'ACCESS MATRIX VS SYSTEM  RUN 25-MAY-2026 1 PM'),
and every stage's outputs land in a subfolder of that run folder:
    <run_folder>\01_Files\
    <run_folder>\02_Reports\
    <run_folder>\03_CSV\
    <run_folder>\04_Comparison\
    <run_folder>\failures\

Stage 1 (copy):             walks the network share and copies the latest
                            PRODUCTION file of every matrix into FILES_FOLDER
                            (one file per matrix over the whole tree; UAT /
                            temp files are ignored; the winner is the newest
                            date written in the file NAME, then the highest
                            version, not the Windows modified date) and writes an index
                            xlsx + log into REPORTS_FOLDER.
Stage 2 (process):          reads each xlsx in FILES_FOLDER, picks the matrix
                            sheet by scanning the top-left 10x10 for the
                            "BLK_ROLEDETAIL" marker (in order: sheet 2 -> 1 ->
                            3), applies the red-fill / white-text cleanup
                            rules to columns 3, 4 and 8+, and writes a cleaned
                            CSV into CSV_FOLDER. Files without the marker are
                            skipped. Files with broken internal drawing/image
                            references (e.g. dangling xl/drawings/NULL) are
                            recorded as CorruptDrawings failures and skipped
                            — the user is expected to repair the xlsx in
                            Excel ("Save As" a fresh copy) and re-run.
Stage 3 (compare):          reads each csv in CSV_FOLDER, normalises it,
                            queries the live Oracle system, and writes a
                            per-file comparison xlsx into
                            COMPARISON_FOLDER/all_differences/. Copies of
                            processed inputs land in successful/ or failed/
                            depending on outcome.
Stage 4 (summary_all_diff): reads every *_comparison.xlsx in
                            all_differences/ and aggregates them into a
                            single _summary.xlsx (Overview + By Role + By
                            Section Detail + per-section drill-down sheets).
Stage 5 (general_sql):      validates every diff row in all_differences/
                            against Oracle LOV tables and writes Oracle SQL
                            (INSERT / UPDATE / DELETE) into
                            all_differences/generated_sql/ to bring the
                            system into alignment with the access matrix.
Stage 6 (send_report):      emails three key output files via Outlook:
                            _summary.xlsx, all_remediation_summary.xlsx,
                            and all_remediation.sql. Recipient + attachments
                            list are read from config.yaml's `email:` section.

Usage:
    python access_matrix_pipeline.py                          # run all six stages
    python access_matrix_pipeline.py copy                     # stage 1 only
    python access_matrix_pipeline.py process                  # stage 2 only
    python access_matrix_pipeline.py compare                  # stage 3 only
    python access_matrix_pipeline.py summary_all_diff         # stage 4 only
    python access_matrix_pipeline.py general_sql              # stage 5 only
    python access_matrix_pipeline.py send_report              # stage 6 only
    python access_matrix_pipeline.py compare summary_all_diff # subset
    python access_matrix_pipeline.py <path/to/config.yaml>    # use a different config
    python access_matrix_pipeline.py <config.yaml> general_sql  # config + stage selection

To re-run stage 6 (or any individual stage) against an EXISTING run folder
rather than creating a new timestamped one, set `paths.run_folder_name` in
config.yaml to the exact folder name you want to reuse.

All settings (paths, DB credentials, Oracle lib dir, skip rules, email
recipient, etc.) live in config.yaml next to this script — see that file
for documentation.
"""

import csv
import glob
import hashlib
import logging
import os
import re
import shutil
import sys
import traceback
import zipfile
from collections import defaultdict
from datetime import datetime
from functools import wraps
from pathlib import Path

import yaml

# Stage 1 + 2 dependencies (always required)
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException

# Stage 3 dependencies — imported lazily inside stage3_compare() so the
# script can still run stages 1 and 2 on a machine that doesn't have
# oracledb / pandas installed (e.g. when the copy + process steps run on
# one PC and the compare step runs on another).
pd = None        # populated by _ensure_stage3_imports()
oracledb = None  # populated by _ensure_stage3_imports()


def _ensure_stage3_imports():
    """Import pandas and oracledb on demand (stage 3 only)."""
    global pd, oracledb
    if pd is None:
        import pandas as _pd
        pd = _pd
    if oracledb is None:
        import oracledb as _oracledb
        oracledb = _oracledb


# ============================================================================
# CONFIGURATION LOADING
# ============================================================================
# Order of resolution for config.yaml:
#   1. first CLI arg ending in .yaml/.yml
#   2. file next to this script
#   3. file in the current working directory
# Anything else on the CLI is treated as a list of stage names to run.

_STAGE_NAMES = {"copy", "process", "compare",
                "summary_all_diff", "general_sql", "send_report"}


# ----------------------------------------------------------------------------
# CROSS-STAGE FAILURE TRACKING
# ----------------------------------------------------------------------------
# Every stage funnels per-item failures into _PIPELINE_FAILURES via the
# _record_failure() helper. At the end of the pipeline the orchestrator
# prints a FAILURES block to the terminal and writes a debug report into
# FAILURES_FOLDER. Cleared at the start of every main() so re-running the
# script doesn't pile up stale entries.
#
# Each entry is a dict with:
#   stage        — one of _STAGE_NAMES
#   name         — short identifier (filename, role id, workbook name, ...)
#   reason_code  — short machine-friendly code (NoMatrixMarker, IOError, ...)
#   detail       — human-readable reason, suitable for terminal display
#   input_path   — optional, the file / source that failed
#   traceback    — optional, full Python traceback if available

_PIPELINE_FAILURES = []

# Per-stage stats accumulated by the orchestrator and read by stage 6
# (send_report) so the email body can include a per-stage summary. Keyed by
# stage name. Empty when stage 6 is run standalone — the email body falls
# back to a "stats not available" message in that case.
_RUN_STATS_BY_STAGE = {}

# Path to the failures_<timestamp>.txt file written by _write_failures_report.
# Set by whichever code path writes it first (stage 6 if email is enabled and
# there are failures, otherwise _print_pipeline_summary at end-of-run). The
# tracking exists so we don't write the file twice with two different
# timestamps within the same run.
_PIPELINE_FAILURES_REPORT_PATH = None


def _record_failure(stage, name, reason_code, detail,
                    input_path=None, tb=None):
    """Append one failure entry to the cross-stage failures list. Safe to call
    from any stage; the orchestrator consumes the list at end-of-run."""
    _PIPELINE_FAILURES.append({
        "stage":       stage,
        "name":        str(name) if name is not None else "(unnamed)",
        "reason_code": str(reason_code) if reason_code else "",
        "detail":      str(detail) if detail else "",
        "input_path":  str(input_path) if input_path else None,
        "traceback":   tb,
    })


def _reset_failures():
    """Clear cross-stage failures. Called once at the top of main()."""
    _PIPELINE_FAILURES.clear()


def _split_cli_args():
    """Return (config_path_or_None, [stage_names])."""
    config_arg = None
    stages = []
    for raw in sys.argv[1:]:
        low = raw.lower().strip()
        if low.endswith((".yaml", ".yml")):
            config_arg = Path(raw)
        elif low in _STAGE_NAMES:
            stages.append(low)
        else:
            raise SystemExit(
                f"Unrecognised argument: {raw!r}. "
                f"Expected a config file (*.yaml) or a stage name "
                f"({', '.join(sorted(_STAGE_NAMES))})."
            )
    return config_arg, stages


def _resolve_config_path() -> Path:
    cli_arg, _ = _split_cli_args()
    if cli_arg is not None:
        return cli_arg
    try:
        candidate = Path(__file__).resolve().parent / "config.yaml"
    except NameError:
        candidate = Path.cwd() / "config.yaml"
    return candidate


def _load_config():
    config_path = _resolve_config_path()
    if not config_path.exists():
        raise FileNotFoundError(
            f"Configuration file not found: {config_path}\n"
            f"Place config.yaml next to the script, or pass its path as "
            f"a command-line argument."
        )
    with open(config_path, "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f) or {}
    return cfg, config_path


CONFIG, CONFIG_PATH = _load_config()


def _require_path(key):
    try:
        return Path(CONFIG["paths"][key])
    except KeyError:
        raise KeyError(
            f"Missing required key in {CONFIG_PATH}: paths.{key}"
        ) from None


# --- Run-folder name formatting ---------------------------------------------
# All output folders for one pipeline run live inside a single timestamped
# folder under OUTPUT_PARENT, e.g.
#   D:\...\ACCESS MATRIX VS SYSTEM  RUN 25-MAY-2026 1 PM\
#       01_Files\
#       02_Reports\
#       03_CSV\
#       04_Comparison\
#       failures\
# The folder name comes from config.yaml's `run_folder_format` template, or
# can be pinned literally via `paths.run_folder_name` (handy for re-running
# a single stage against a previous run's outputs).

_DEFAULT_RUN_FOLDER_FORMAT = (
    "ACCESS MATRIX VS SYSTEM  RUN "
    "{day:02d}-{month_abbr}-{year} {hour_12} {ampm}"
)


def _format_run_folder_name(template, now=None):
    """Render a run-folder template into a folder name.

    Supports these tokens (substituted via str.format()):
        {day}, {day:02d}      — day of month
        {month_abbr}          — three-letter month, UPPERCASE (JAN, MAY ...)
        {month}, {month:02d}  — month number
        {year}                — 4-digit year
        {hour_12}             — 12-hour clock (1..12, no leading zero)
        {hour_24}             — 24-hour clock (0..23)
        {minute}, {minute:02d}— minute
        {ampm}                — 'AM' or 'PM'
    With the default template + now=2026-05-25 13:00, the result is:
        'ACCESS MATRIX VS SYSTEM  RUN 25-MAY-2026 1 PM'
    """
    if now is None:
        now = datetime.now()
    try:
        return template.format(
            day=now.day,
            month=now.month,
            month_abbr=now.strftime("%b").upper(),
            year=now.year,
            hour_12=now.hour % 12 or 12,
            hour_24=now.hour,
            minute=now.minute,
            ampm="AM" if now.hour < 12 else "PM",
        )
    except (KeyError, IndexError) as exc:
        raise ValueError(
            f"Invalid run_folder_format token in config.yaml: {exc}. "
            f"Supported tokens: day, month, month_abbr, year, hour_12, "
            f"hour_24, minute, ampm."
        ) from None


# --- Path settings used across stages ---------------------------------------
# Source folder (network share or local dir) walked by stage 1. Required.
SOURCE_FOLDER = _require_path("source_folder")

# Parent directory under which the timestamped run folder is created.
# Every pipeline run gets its own subdirectory here.
OUTPUT_PARENT = _require_path("output_parent")

# Pin the run folder name (so stage 6 etc. point at an existing run's
# outputs) OR auto-generate it from the format template.
_pinned_run_name = CONFIG.get("paths", {}).get("run_folder_name")
if _pinned_run_name and str(_pinned_run_name).strip():
    RUN_FOLDER_NAME = str(_pinned_run_name).strip()
else:
    RUN_FOLDER_NAME = _format_run_folder_name(
        CONFIG.get("run_folder_format", _DEFAULT_RUN_FOLDER_FORMAT)
    )

RUN_FOLDER = OUTPUT_PARENT / RUN_FOLDER_NAME

# Sub-folder names within the run folder (overridable in config; defaults match
# the layout the rest of the script + the existing user-side documentation).
_subnames = CONFIG.get("subfolder_names") or {}
FILES_FOLDER = RUN_FOLDER / _subnames.get("files",      "01_Files")
REPORTS_FOLDER = RUN_FOLDER / _subnames.get("reports",    "02_Reports")
CSV_FOLDER = RUN_FOLDER / _subnames.get("csv",        "03_CSV")
COMPARISON_FOLDER = RUN_FOLDER / _subnames.get("comparison", "04_Comparison")
FAILURES_FOLDER = RUN_FOLDER / _subnames.get("failures",   "failures")

# Stage 1 historically referred to its input as ROOT_FOLDER. Keep that alias
# so the rest of stage 1's code is untouched.
ROOT_FOLDER = SOURCE_FOLDER

try:
    ORACLE_LIB_DIR = str(CONFIG["paths"]["oracle_client_lib_dir"])
except KeyError:
    ORACLE_LIB_DIR = None  # only required by stages 3 and 5

# --- Excel extensions normalised to leading-dot lowercase -------------------
_raw_exts = CONFIG.get("excel_extensions", [".xlsx", ".xlsm", ".xls", ".xlsb"])
EXCEL_EXTENSIONS = {
    (ext if ext.startswith(".") else f".{ext}").lower().strip()
    for ext in _raw_exts
}

# --- Stage 1: skip rules + duplicate handling -------------------------------
SKIP_FOLDERS = set(CONFIG.get("skip_folders", []))
SKIP_FOLDERS_LOWER = {name.lower() for name in SKIP_FOLDERS}


def _normalize_skip_file(pattern: str) -> str:
    p = str(pattern).strip().lower()
    stem, ext = os.path.splitext(p)
    if ext in EXCEL_EXTENSIONS:
        return stem
    return p


SKIP_FILES = set(CONFIG.get("skip_files", []))
SKIP_FILE_STEMS = {_normalize_skip_file(p)
                   for p in SKIP_FILES if str(p).strip()}

DUPLICATE_HANDLING = str(CONFIG.get(
    "duplicate_handling", "keep_newest")).lower()
if DUPLICATE_HANDLING not in {"keep_newest", "keep_all"}:
    raise ValueError(
        f"Invalid duplicate_handling in {CONFIG_PATH}: {DUPLICATE_HANDLING!r}. "
        f"Use 'keep_newest' or 'keep_all'."
    )

# --- Stage 1: which matrix files are copied ---------------------------------
# Only PRODUCTION matrices are copied. A file is rejected when its name (or,
# if the name says nothing, its folder) carries one of `exclude_keywords`
# (UAT / temp ...). Step 1: each folder on its own gives ONE file, its latest
# production one. Step 2: when the same matrix won in several folders, the
# newest of those winners is kept. "Latest" = newest DATE written in the file
# name, then highest VERSION in the name (V10, V11 ...). The Windows "Date
# modified" is only the very last tie-break.
_sel = CONFIG.get("matrix_selection") or {}
EXCLUDE_KEYWORDS = [str(k).strip().lower()
                    for k in (_sel.get("exclude_keywords")
                              or ["uat", "temp", "temporary", "tmp"])
                    if str(k).strip()]
INCLUDE_UNMARKED = bool(_sel.get("include_unmarked", True))
CHECK_FOLDER_NAMES = bool(_sel.get("check_folder_names", True))
SELECTION_SCOPE = str(_sel.get("scope", "matrix")).strip().lower()
if SELECTION_SCOPE not in {"matrix", "folder", "name"}:
    raise ValueError(
        f"Invalid matrix_selection.scope in {CONFIG_PATH}: {SELECTION_SCOPE!r}. "
        f"Use 'matrix', 'folder' or 'name'."
    )

TIMESTAMP_FORMAT = CONFIG.get("timestamp_format", "%Y%m%d_%H%M%S")
TIMESTAMP = datetime.now().strftime(TIMESTAMP_FORMAT)

LOG_FILE_TEMPLATE = CONFIG.get("log_file_template", "copy_log_{timestamp}.log")
REPORT_FILE_TEMPLATE = CONFIG.get(
    "report_file_template", "access_matrix_index_{timestamp}.xlsx"
)

LOG_FILE = REPORTS_FOLDER / LOG_FILE_TEMPLATE.format(timestamp=TIMESTAMP)
REPORT_FILE = REPORTS_FOLDER / REPORT_FILE_TEMPLATE.format(timestamp=TIMESTAMP)

# --- Stage 3: database ------------------------------------------------------
_db = CONFIG.get("database") or {}
UAT_CONFIG = {
    "user":     str(_db.get("user", "")),
    "password": str(_db.get("password", "")),
    "dsn":      str(_db.get("dsn", "")),
}
# Note: we don't raise here because stages 1 and 2 don't need DB credentials.
# Stage 3 validates them on entry.

# --- Stage 3: optional Oracle database link --------------------------------
# When `database.db_link` is set in config.yaml, every FCPROD.<table> reference
# in the SQL queries below is automatically rewritten to FCPROD.<table>@<db_link>.
# This lets you connect to one Oracle service (e.g. OBMACOMPDB as COMMONCORE)
# and reach the real FCPROD schema on a different database (e.g. V14_UAT_RO)
# through a pre-existing public database link, without rewriting any queries.
#
# Example: with db_link='V14_UAT_RO',
#     select * from FCPROD.SMTB_ROLE_MASTER rm
# is executed as
#     select * from FCPROD.SMTB_ROLE_MASTER@V14_UAT_RO rm
DB_LINK = str(_db.get("db_link", "")).strip()

# Match `fcprod.<identifier>` (case-insensitive), excluding identifiers that
# already carry an @dblink suffix. The capture preserves the original casing,
# so 'FCPROD.SMTB_ROLE_MASTER' stays uppercase and 'fcprod.smtb_role_glrest'
# stays lowercase in the rewritten SQL.
import re as _re_sql  # noqa: E402  — local alias avoids shadowing later uses

_FCPROD_TABLE_RE = _re_sql.compile(
    r'\b(fcprod\.[A-Za-z_][A-Za-z0-9_]*)\b(?!@)',
    _re_sql.IGNORECASE,
)


def _apply_db_link(sql):
    """Rewrite every FCPROD.<table> reference in `sql` to FCPROD.<table>@<DB_LINK>.

    No-op if DB_LINK is empty. Idempotent — a table that already has '@something'
    is skipped, so calling this twice on the same string is safe.
    """
    if not DB_LINK:
        return sql
    return _FCPROD_TABLE_RE.sub(lambda m: f"{m.group(1)}@{DB_LINK}", sql)


# ============================================================================
# WINDOWS LONG-PATH HELPERS (used by stage 1)
# ============================================================================
# Default MAX_PATH on Windows is 260 chars. Prefixing with \\?\ (or
# \\?\UNC\ for UNC paths) tells Windows to bypass that limit.
def _to_long(p):
    if os.name != "nt":
        return str(p)
    s = str(p)
    if s.startswith("\\\\?\\"):
        return s
    if s.startswith("\\\\"):
        return "\\\\?\\UNC\\" + s[2:]
    if len(s) >= 2 and s[1] == ":":
        return "\\\\?\\" + s
    return s


def _from_long(p):
    s = str(p)
    if s.startswith("\\\\?\\UNC\\"):
        return "\\\\" + s[8:]
    if s.startswith("\\\\?\\"):
        return s[4:]
    return s


def _fmt_mtime(mtime):
    """Format an epoch mtime for display, gracefully handling missing values."""
    if not mtime:
        return "?"
    try:
        return datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")
    except (OSError, ValueError):
        return "?"


def _banner(title, char="=", width=72):
    """Print a visible stage banner to both stdout and stage 1's logger."""
    bar = char * width
    print()
    print(bar)
    print(f"  {title}")
    print(bar)
    try:
        logging.info(bar)
        logging.info(f"  {title}")
        logging.info(bar)
    except Exception:
        pass


# ============================================================================
# STAGE 1: COPY  —  network share → FILES_FOLDER
# ============================================================================

def _stage1_setup_logging():
    REPORTS_FOLDER.mkdir(parents=True, exist_ok=True)
    FILES_FOLDER.mkdir(parents=True, exist_ok=True)
    # Reset handlers in case stage 3 (or a prior stage) set up its own logger
    root = logging.getLogger()
    for h in list(root.handlers):
        root.removeHandler(h)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(_to_long(LOG_FILE), encoding="utf-8"),
            logging.StreamHandler(),
        ],
    )


# ----------------------------------------------------------------------------
# Matrix file selection (production only, newest DATE in the name wins)
# ----------------------------------------------------------------------------
# Two steps: (1) every folder on its own gives its latest production file;
# (2) when the same matrix won in several folders, only the newest of those
# folder winners is kept. UAT / temp files never take part.
#
# Typical names:
#   V10_Access_Matrix_FlexCube_SBO_PROD_ 09-11-2026.xlsx     -> prod, 09-11-2026
#   V10_Access_Matrix_FlexCube_SBO_UAT_ 05-11-2026(...).xlsx -> UAT  (ignored)
#   V10_..._Operation-TMO_PROD_02-06-2026_Temp.xlsx          -> temp (ignored)
#   Pre Security Access Matrix FlexCube of HR  Prod V1 19-10-2023.xlsx -> v1
#   V4__FC_Central Blockingv3 dormantv2 UnitProd_V4_13082026.xlsx      -> v4

_VERSION_LEAD_RE = re.compile(r"^\W*v\s*(\d+(?:\.\d+)*)(?!\d)", re.IGNORECASE)
_VERSION_ANY_RE = re.compile(r"(?<![a-z0-9])v\s*(\d+(?:\.\d+)*)(?!\d)",
                             re.IGNORECASE)
_PROD_RE = re.compile(r"prod(?:uction)?(?![a-z])", re.IGNORECASE)

_DATE_YEAR_MIN, _DATE_YEAR_MAX = 2010, 2040
_MONTH_NUM = {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
              "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12}
_DATE_DMY_RE = re.compile(
    r"(?<!\d)(\d{1,2})\s*[-_./]\s*(\d{1,2})\s*[-_./]\s*(\d{4}|\d{2})(?!\d)")
_DATE_DMY_SPACE_RE = re.compile(r"(?<!\d)(\d{1,2}) (\d{1,2}) (\d{4})(?!\d)")
_DATE_YMD_RE = re.compile(
    r"(?<!\d)(20\d{2})\s*[-_./]\s*(\d{1,2})\s*[-_./]\s*(\d{1,2})(?!\d)")
_DATE_MONTHNAME_RE = re.compile(
    r"(?<!\d)(\d{1,2})[\s\-_./]*"
    r"(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*"
    r"[\s\-_./,]*(\d{4}|\d{2})(?!\d)", re.IGNORECASE)
_DIGIT_RUN_RE = re.compile(r"\d{6,}")
_ARABIC_MARKS_RE = re.compile(r"[ً-ْـ]")


def _parse_version(stem):
    """Version number written in the file NAME, as a comparable tuple.

    The leading 'V<n>' wins ('V2_..._Prod V1 31-08-2025' -> (2,)); otherwise
    the first 'V<n>' token anywhere in the name ('... Prod V5 10-07-2023' ->
    (5,)). Files with no version return (-1,) so any versioned file beats
    them."""
    m = _VERSION_LEAD_RE.search(stem) or _VERSION_ANY_RE.search(stem)
    if not m:
        return (-1,)
    return tuple(int(x) for x in m.group(1).split("."))


def _mk_date(y, m, d):
    try:
        y, m, d = int(y), int(m), int(d)
        if y < 100:
            y += 2000
        if not (_DATE_YEAR_MIN <= y <= _DATE_YEAR_MAX):
            return None
        return datetime(y, m, d)
    except ValueError:
        return None


def _date_from_digit_run(run):
    """ddmmyyyy / yyyymmdd / dmmyyyy / ddmyyyy / ddmmyy. A run longer than 8
    digits is treated as a typo ('106092026') and its last 8 digits are used."""
    n = len(run)
    tries = []
    if n == 8:
        tries = [(run[4:], run[2:4], run[:2]), (run[:4], run[4:6], run[6:])]
    elif n == 7:
        tries = [(run[3:], run[1:3], run[:1]), (run[3:], run[2:3], run[:2])]
    elif n == 6:
        tries = [(run[4:], run[2:4], run[:2])]
    elif n > 8:
        t, h = run[-8:], run[:8]
        tries = [(t[4:], t[2:4], t[:2]), (h[4:], h[2:4], h[:2])]
    for y, m, d in tries:
        dt = _mk_date(y, m, d)
        if dt:
            return dt
    return None


def _parse_name_date(stem):
    """The date written in the file name (the right-most one), or None.

    Understands 09-11-2026, 9-11-2026, 9.11.26, '- 5-8-2026', 09112026,
    2026-11-09, 20261109, 9-Nov-2026, 9 November 2026."""
    found = []   # (end position, date)
    for m in _DATE_DMY_RE.finditer(stem):
        dt = _mk_date(m.group(3), m.group(2), m.group(1))
        if dt:
            found.append((m.end(), dt))
    for m in _DATE_DMY_SPACE_RE.finditer(stem):
        dt = _mk_date(m.group(3), m.group(2), m.group(1))
        if dt:
            found.append((m.end(), dt))
    for m in _DATE_YMD_RE.finditer(stem):
        dt = _mk_date(m.group(1), m.group(2), m.group(3))
        if dt:
            found.append((m.end(), dt))
    for m in _DATE_MONTHNAME_RE.finditer(stem):
        dt = _mk_date(m.group(3), _MONTH_NUM[m.group(2).lower()], m.group(1))
        if dt:
            found.append((m.end(), dt))
    for m in _DIGIT_RUN_RE.finditer(stem):
        dt = _date_from_digit_run(m.group(0))
        if dt:
            found.append((m.end(), dt))
    if not found:
        return None
    found.sort(key=lambda t: t[0])
    return found[-1][1]


def _keyword_hit(text, kw):
    """True when `kw` (uat / temp ...) is written in `text` as its own word.

    Catches '_UAT_', 'UAT-1-10', 'uat09', '(Temp)' and also glued spellings
    such as 'SMEUAT_', 'UATjob', 'ProdTemp', 'TempV2' - but NOT the letters
    inside a normal word ('Evaluation', 'EVALUATION', 'attempt', 'Template')."""
    if re.search(r"(?<![a-z])" + re.escape(kw) + r"(?![a-z])", text.lower()):
        return True
    upper = kw.upper()
    for m in re.finditer(re.escape(upper), text):
        prev = text[m.start() - 1] if m.start() > 0 else ""
        nxt = text[m.end()] if m.end() < len(text) else ""
        if not (prev.isupper() and nxt.isupper()):
            return True
    cap = kw.capitalize()
    if cap != upper:
        for m in re.finditer(re.escape(cap), text):
            nxt = text[m.end()] if m.end() < len(text) else ""
            if not nxt.islower():
                return True
    return False


def _classify_env(stem, rel_folders=()):
    """Return (env, why): env is 'excluded' (UAT / temp ...), 'prod' or
    'unmarked'. The FILE NAME decides; only when the file name says nothing
    are the folder names (nearest folder first) consulted."""
    for kw in EXCLUDE_KEYWORDS:
        if _keyword_hit(stem, kw):
            return "excluded", f"file name contains '{kw}'"
    if _PROD_RE.search(stem):
        return "prod", ""
    if CHECK_FOLDER_NAMES:
        for folder in reversed(list(rel_folders)):
            for kw in EXCLUDE_KEYWORDS:
                if _keyword_hit(folder, kw):
                    return "excluded", f"folder '{folder}' contains '{kw}'"
            if _PROD_RE.search(folder):
                return "prod", ""
    return "unmarked", ""


def _drop_non_prod(paths, stage, say):
    """SAFETY NET for stages 2-5: remove every path whose FILE NAME is marked
    UAT / temp (matrix_selection.exclude_keywords) so such a file can never
    be converted, compared with the database, summarised or turned into SQL -
    even if it was put into the run folder by hand or left by an old run.
    `say` is the stage's own print / log function. Returns the kept paths."""
    kept = []
    for p in paths:
        name = os.path.basename(str(p))
        env, why = _classify_env(os.path.splitext(name)[0])
        if env == "excluded":
            say(f"SKIPPED - not a production matrix ({why}): {name}")
            _record_failure(
                stage=stage, name=name, reason_code="NotProduction",
                detail=f"Ignored, not a production matrix ({why}). "
                       f"It was NOT compared with the database.",
                input_path=str(p))
        else:
            kept.append(p)
    return kept


_KEY_STOPWORDS = {"pre", "security", "access", "matrix", "flexcube", "fc",
                  "f", "c", "of", "the", "and", "v"}


def _matrix_name_key(stem):
    """Matrix name with version / environment / date / filler words removed,
    so different versions of the SAME matrix get the same key:
        'V14_Access Matrix_FC_CardCenter_Prod_ 9-11-2025'  -> 'cardcenter'
        'V15_Access Matrix_FC_CardCenter_Prod_ 04-01-2026' -> 'cardcenter'
        'Pre Security Access Matrix FlexCube of Credit Admin - SME Prod V7 ...'
        'V11_Access_Matrix_FlexCube_ Credit_Admin - SME Prod 25-09-2024'
                                                        -> 'credit admin sme'
    Returns '' when nothing meaningful is left."""
    t = _ARABIC_MARKS_RE.sub("", stem).lower()
    t = _DATE_MONTHNAME_RE.sub(" ", t)      # '9-Nov-2025' leaves no 'nov'
    t = _PROD_RE.sub(" ", t)
    t = re.sub(r"(?<![a-z0-9])v\s*\d+(?:\.\d+)*(?!\d)", " ", t)
    t = re.sub(r"[^a-z0-9؀-ۿ]+", " ", t)
    words = []
    for w in t.split():
        if w.isdigit():
            continue                      # dates, '(002)' copy markers
        w = re.sub(r"(?<=[a-z])v\d+$", "", w)   # 'blockingv3' -> 'blocking'
        if w and w not in _KEY_STOPWORDS:
            words.append(w)
    return " ".join(words)


def _fmt_version(v):
    return "V" + ".".join(map(str, v)) if v != (-1,) else "no version"


def _fmt_name_date(d):
    return d.strftime("%d-%m-%Y") if d else "no date"


def _select_matrices(candidates):
    """Pick the matrices to copy out of EVERY Excel file found in the tree.

    candidates: dicts with name / stem / source_long / source_folder /
                rel_folders / mtime.
    Returns (winners, rejected) where rejected is a list of (candidate, reason).

    Step 0  Files that are UAT / temp (by file name, or by folder name when
            the file name says nothing) are dropped.
    Step 1  EACH FOLDER ON ITS OWN: the folder's latest production file is
            chosen; every other file of that folder is out.
    Step 2  ACROSS FOLDERS: only the folder winners of step 1 are compared.
            When the same matrix (same name once version / date / Prod are
            removed) won in more than one folder, the newest one is kept.

    "Latest" in both steps: a file saying 'Prod' beats a file with no
    environment word; then the newest DATE written in the file name; then the
    highest VERSION in the name; then Windows modified time.

    matrix_selection.scope: 'matrix' = step 1 + step 2 (default),
    'folder' = step 1 only, 'name' = step 2 only (applied to every file)."""
    rejected = []
    pool = []
    for c in candidates:
        env, why = _classify_env(c["stem"], c.get("rel_folders", ()))
        c["env"] = env
        c["version"] = _parse_version(c["stem"])
        c["name_date"] = _parse_name_date(c["stem"])
        c["name_key"] = _matrix_name_key(c["stem"])
        if env == "excluded":
            rejected.append((c, f"Not production ({why})"))
        else:
            pool.append(c)

    def _key(c):
        return (c["name_date"] or datetime.min, c["version"], c["mtime"])

    def _pick_one(group, context):
        """Return the winner of `group` (or None); losers go to `rejected`."""
        prod = [c for c in group if c["env"] == "prod"]
        unmarked = [c for c in group if c["env"] != "prod"]
        if prod:
            tier = prod
            best = max(prod, key=_key)
            for c in unmarked:
                rejected.append((c, f"{context}: no 'Prod' in its name while "
                                    f"'{best['name']}' is marked Prod"))
        elif INCLUDE_UNMARKED:
            tier = unmarked
        else:
            for c in unmarked:
                rejected.append((c, "No 'Prod' in the file / folder name "
                                    "(matrix_selection.include_unmarked is false)"))
            return None

        tier = sorted(tier, key=_key, reverse=True)
        winner = tier[0]
        for c in tier[1:]:
            where = ("" if c["source_folder"] == winner["source_folder"]
                     else f" in folder '{winner['source_folder']}'")
            if c["name_date"] != winner["name_date"]:
                why = (f"older date in file name "
                       f"({_fmt_name_date(c['name_date'])} < "
                       f"{_fmt_name_date(winner['name_date'])})")
            elif c["version"] != winner["version"]:
                why = (f"same date in name, lower version "
                       f"({_fmt_version(c['version'])} < "
                       f"{_fmt_version(winner['version'])})")
            else:
                why = "same date and version, older Windows modified time"
            rejected.append(
                (c, f"{context}: {why} - kept '{winner['name']}'{where}"))
        return winner

    # ---- Step 1: each folder on its own ----
    if SELECTION_SCOPE in ("matrix", "folder"):
        by_folder = defaultdict(list)
        for c in pool:
            by_folder[c["source_folder"].lower()].append(c)
        stage1 = []
        for grp in by_folder.values():
            w = _pick_one(grp, "Same folder")
            if w is not None:
                stage1.append(w)
    else:
        stage1 = pool

    # ---- Step 2: same matrix that won in several folders ----
    if SELECTION_SCOPE in ("matrix", "name"):
        by_name = defaultdict(list)
        winners = []
        for c in stage1:
            if c["name_key"]:
                by_name[c["name_key"]].append(c)
            elif SELECTION_SCOPE == "name" and not INCLUDE_UNMARKED \
                    and c["env"] != "prod":
                rejected.append((c, "No 'Prod' in the file / folder name "
                                    "(matrix_selection.include_unmarked is false)"))
            else:
                winners.append(c)          # no usable name -> stands alone
        for grp in by_name.values():
            w = _pick_one(grp, "Same matrix in another folder"
                          if SELECTION_SCOPE == "matrix" else "Same matrix")
            if w is not None:
                winners.append(w)
    else:
        winners = stage1

    winners.sort(key=lambda c: c["name"].lower())
    return winners, rejected


def _possible_duplicates(winners):
    """Pairs of SELECTED files whose names look like the same matrix but were
    not merged automatically (name spelled differently AND different folder).
    Listed in the report for a human to review - nothing is dropped."""
    import difflib
    out = []
    for i in range(len(winners)):
        a = winners[i]
        ka = a.get("name_key") or ""
        sa = set(ka.split())
        if not sa:
            continue
        for j in range(i + 1, len(winners)):
            b = winners[j]
            kb = b.get("name_key") or ""
            sb = set(kb.split())
            if not sb:
                continue
            small, big = (sa, sb) if len(sa) <= len(sb) else (sb, sa)
            if len(small) >= 2 and small <= big:
                why = "one name contains the other"
            elif difflib.SequenceMatcher(None, ka, kb).ratio() >= 0.88:
                why = "names are almost identical (spelling?)"
            else:
                continue
            out.append((a, b, why))
    return out


def _excel_files_in_folder(folder_long: str):
    """Return (excel_files, list_of_all_entry_names) for one folder.
    excel_files: dicts {name, stem, source_long, mtime} for every Excel file
    that is not a lock file ('~$...') and not in skip_files."""
    raw_entries = []
    try:
        raw_entries = os.listdir(folder_long)
    except (PermissionError, OSError) as e:
        logging.warning(f"Cannot list folder {_from_long(folder_long)}: {e}")
        return [], raw_entries

    files = []
    for name in raw_entries:
        item_long = os.path.join(folder_long, name)
        try:
            if not os.path.isfile(item_long):
                continue
            if name.startswith("~$"):
                continue
            stem, ext = os.path.splitext(name)
            if ext.lower().strip() not in EXCEL_EXTENSIONS:
                continue
            if stem.lower().strip() in SKIP_FILE_STEMS:
                logging.info(
                    f"Skipping file (in skip_files): {_from_long(item_long)}")
                continue
            try:
                mtime = os.path.getmtime(item_long)
            except OSError as e:
                logging.warning(
                    f"Cannot stat file {_from_long(item_long)}: {e}")
                continue
            files.append({"name": name, "stem": stem,
                          "source_long": item_long, "mtime": mtime})
        except OSError as e:
            logging.warning(
                f"Error inspecting {name} in {_from_long(folder_long)}: {e}")
            continue
    return files, raw_entries


def _stage1_archive_leftovers():
    """A previous run that used the SAME run folder (same hour, or a pinned
    run_folder_name) leaves its files behind; they would be mixed with this
    run's selection (old UAT / temp copies, '_1' duplicates). Move them out of
    the way - nothing is deleted - into '<run folder>/_previous_run_<time>/'."""
    try:
        leftovers = [n for n in os.listdir(str(FILES_FOLDER))
                     if os.path.splitext(n)[1].lower() in EXCEL_EXTENSIONS]
    except OSError:
        leftovers = []
    if not leftovers:
        return
    archive = RUN_FOLDER / f"_previous_run_{TIMESTAMP}"
    logging.info(
        f"Run folder already holds {len(leftovers)} Excel file(s) from an "
        f"earlier run - moving the earlier outputs to: {archive}")
    moves = [(FILES_FOLDER / n, archive / FILES_FOLDER.name / n)
             for n in leftovers]
    for folder in (CSV_FOLDER, COMPARISON_FOLDER):
        if folder.exists():
            moves.append((folder, archive / folder.name))
    for src, dst in moves:
        try:
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(_to_long(src), _to_long(dst))
        except Exception as e:
            logging.warning(f"  Could not move {src} : {e}")


def _resolve_duplicates(candidates):
    """Group candidates by case-insensitive filename. Keep the newest per group.
    Returns (winners, losers).
    """
    if DUPLICATE_HANDLING == "keep_all":
        return list(candidates), []

    groups = defaultdict(list)
    for c in candidates:
        groups[c["name"].lower()].append(c)

    winners, losers = [], []
    for _, group in groups.items():
        if len(group) == 1:
            winners.append(group[0])
            continue
        group.sort(key=lambda x: x["mtime"], reverse=True)
        winner = group[0]
        winners.append(winner)
        winner_mtime = _fmt_mtime(winner["mtime"])
        winner_folder = winner["source_folder"]
        for loser in group[1:]:
            loser["reason"] = (
                f"Older than winner in '{winner_folder}' "
                f"(winner modified {winner_mtime})"
            )
            losers.append(loser)
    return winners, losers


def _autosize(ws, widths):
    for letter, width in widths:
        ws.column_dimensions[letter].width = width


def _write_xlsx_report(copied_files, folders_without_excel, duplicates_skipped,
                       not_selected=None, possible_dups=None):
    """5-sheet index xlsx: Access Matrix Index / Folders Without Excel /
    Duplicates Skipped / Not Selected (UAT, temp, older versions) /
    Possible Duplicates (selected files that look like the same matrix)."""
    wb = Workbook()

    ws = wb.active
    ws.title = "Access Matrix Index"
    ws.append(["File Name", "Source Path", "Modified Time",
               "Date In Name", "Version In Name", "Environment"])
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in copied_files:
        ws.append(list(row))
    name_w = max((len(r[0]) for r in copied_files), default=0)
    path_w = max((len(r[1]) for r in copied_files), default=0)
    _autosize(ws, [("A", max(30, name_w + 2)),
                   ("B", max(50, path_w + 2)),
                   ("C", 22), ("D", 14), ("E", 16), ("F", 14)])
    ws.freeze_panes = "A2"

    ws2 = wb.create_sheet(title="Folders Without Excel")
    ws2.append(["Folder Path", "Entries Found"])
    for cell in ws2[1]:
        cell.font = Font(bold=True)
    for folder_path, files_str in folders_without_excel:
        ws2.append([folder_path, files_str])
    p_w = max((len(p) for p, _ in folders_without_excel), default=0)
    f_w = max((len(f) for _, f in folders_without_excel), default=0)
    _autosize(ws2, [("A", max(60, p_w + 2)),
                    ("B", max(40, min(120, f_w + 2)))])
    ws2.freeze_panes = "A2"

    ws3 = wb.create_sheet(title="Duplicates Skipped")
    ws3.append(["File Name", "Source Path", "Modified Time", "Reason"])
    for cell in ws3[1]:
        cell.font = Font(bold=True)
    for d in duplicates_skipped:
        ws3.append([
            d["name"],
            d["source_folder"],
            _fmt_mtime(d["mtime"]),
            d.get("reason", ""),
        ])
    dn_w = max((len(d["name"]) for d in duplicates_skipped), default=0)
    dp_w = max((len(d["source_folder"])
               for d in duplicates_skipped), default=0)
    dr_w = max((len(d.get("reason", ""))
               for d in duplicates_skipped), default=0)
    _autosize(ws3, [("A", max(30, dn_w + 2)),
                    ("B", max(60, dp_w + 2)),
                    ("C", 22),
                    ("D", max(40, min(120, dr_w + 2)))])
    ws3.freeze_panes = "A2"

    ws4 = wb.create_sheet(title="Not Selected")
    ws4.append(["File Name", "Source Path", "Modified Time", "Reason"])
    for cell in ws4[1]:
        cell.font = Font(bold=True)
    not_selected = not_selected or []
    for n in not_selected:
        ws4.append([n["name"], n["source_folder"],
                    _fmt_mtime(n["mtime"]), n["reason"]])
    nn_w = max((len(n["name"]) for n in not_selected), default=0)
    np_w = max((len(n["source_folder"]) for n in not_selected), default=0)
    nr_w = max((len(n["reason"]) for n in not_selected), default=0)
    _autosize(ws4, [("A", max(30, nn_w + 2)),
                    ("B", max(60, np_w + 2)),
                    ("C", 22),
                    ("D", max(40, min(120, nr_w + 2)))])
    ws4.freeze_panes = "A2"

    ws5 = wb.create_sheet(title="Possible Duplicates")
    ws5.append(["File A", "Folder A", "File B", "Folder B", "Why listed"])
    for cell in ws5[1]:
        cell.font = Font(bold=True)
    for a, b, why in (possible_dups or []):
        ws5.append([a["name"], a["source_folder"],
                    b["name"], b["source_folder"], why])
    _autosize(ws5, [("A", 70), ("B", 60), ("C", 70), ("D", 60), ("E", 40)])
    ws5.freeze_panes = "A2"

    wb.save(_to_long(REPORT_FILE))


def stage1_copy():
    """Walk ROOT_FOLDER; step 1: each folder gives its latest PROD file
    (newest date, then highest version, both read from the file name; UAT /
    temp ignored); step 2: when the same matrix won in several folders only
    the newest is kept. The winners are copied to FILES_FOLDER.
    Returns a dict of run statistics (used by the pipeline summary)."""
    _stage1_setup_logging()
    _banner("STAGE 1 — COPY  (network share -> Files/)")

    logging.info(f"Config file       : {CONFIG_PATH}")
    logging.info(f"Source            : {ROOT_FOLDER}")
    logging.info(f"Files folder      : {FILES_FOLDER}")
    logging.info(f"Reports folder    : {REPORTS_FOLDER}")
    logging.info(f"Skip folders      : {sorted(SKIP_FOLDERS)}")
    logging.info(f"Skip files        : {sorted(SKIP_FILES)}")
    logging.info(f"Duplicate handling: {DUPLICATE_HANDLING}")
    logging.info(f"Selection         : PROD only, newest date in file name "
                 f"(then highest version) wins (scope={SELECTION_SCOPE}, "
                 f"exclude={EXCLUDE_KEYWORDS}, "
                 f"include_unmarked={INCLUDE_UNMARKED}, "
                 f"check_folder_names={CHECK_FOLDER_NAMES})")
    logging.info(f"Log file          : {LOG_FILE}")
    logging.info(f"Report file       : {REPORT_FILE}")
    logging.info("=" * 72)

    walk_root = _to_long(ROOT_FOLDER)
    if not os.path.exists(walk_root):
        logging.error(
            f"Source folder does not exist or is unreachable: {ROOT_FOLDER}")
        return {"stage": "copy", "ok": False, "error": "source not reachable"}

    _stage1_archive_leftovers()

    candidates = []
    not_selected = []
    all_excel = []
    folders_without_excel = []
    folders_scanned = 0
    folders_with_excel = 0

    for current_folder, subfolders, _files in os.walk(walk_root):
        current_long = current_folder
        current_display = _from_long(current_long)
        folders_scanned += 1
        has_subfolders = len(subfolders) > 0

        pruned = [d for d in subfolders if d.lower() in SKIP_FOLDERS_LOWER]
        for d in pruned:
            logging.info(
                f"Skipping folder: {os.path.join(current_display, d)}")
        subfolders[:] = [d for d in subfolders
                         if d.lower() not in SKIP_FOLDERS_LOWER]

        found, raw_entries = _excel_files_in_folder(current_long)

        if not found:
            if not has_subfolders:
                entries_str = ", ".join(raw_entries) if raw_entries else ""
                folders_without_excel.append((current_display, entries_str))
                if raw_entries:
                    logging.info(
                        f"Flagged 'no Excel' — folder has {len(raw_entries)} "
                        f"entries but none are matching Excel files: "
                        f"{current_display}")
                    for name in raw_entries:
                        entry_long = os.path.join(current_long, name)
                        try:
                            kind = ("FILE" if os.path.isfile(entry_long)
                                    else "DIR" if os.path.isdir(entry_long)
                                    else "OTHER")
                        except OSError as e:
                            kind = f"ERROR({type(e).__name__})"
                        logging.info(f"    [{kind}] {name}")
                else:
                    logging.info(
                        f"Flagged 'no Excel' — os.listdir returned 0 entries "
                        f"(check permissions / network): {current_display}")
            continue

        folders_with_excel += 1
        rel = os.path.relpath(current_long, walk_root)
        rel_folders = [] if rel in (".", "") else [
            p for p in re.split(r"[\\/]+", rel) if p]
        for f in found:
            f["source_folder"] = current_display
            f["rel_folders"] = rel_folders
            all_excel.append(f)

    logging.info("=" * 72)
    logging.info(f"Excel files found in the whole tree: {len(all_excel)}")
    picked, rejected = _select_matrices(all_excel)

    for rc, reason in sorted(
            rejected, key=lambda t: (t[0]["source_folder"], t[0]["name"])):
        not_selected.append({
            "name": rc["name"],
            "source_folder": rc["source_folder"],
            "mtime": rc["mtime"],
            "reason": reason,
        })
        logging.info(f"  Not selected: {rc['name']} "
                     f"[{rc['source_folder']}] — {reason}")

    for pk in picked:
        candidates.append(pk)
        logging.info(
            f"  Selected: {pk['name']} ({pk['env']}, "
            f"date {_fmt_name_date(pk['name_date'])}, "
            f"{_fmt_version(pk['version'])}) [{pk['source_folder']}]")

    possible_dups = _possible_duplicates(picked)
    if possible_dups:
        logging.info("=" * 72)
        logging.info(
            f"{len(possible_dups)} pair(s) of selected files look like the "
            f"same matrix under different names - please review "
            f"(sheet 'Possible Duplicates'):")
        for a, b, why in possible_dups:
            logging.info(f"  ? {a['name']}  <->  {b['name']}  ({why})")

    winners, losers = _resolve_duplicates(candidates)
    if losers:
        logging.info("=" * 72)
        logging.info(
            f"Duplicate resolution: {len(losers)} older version(s) will be "
            f"skipped (kept the newest of each filename)")
        for L in losers:
            logging.info(
                f"  SKIP: {L['name']} from {L['source_folder']} "
                f"(modified {_fmt_mtime(L['mtime'])}) — {L['reason']}")

    copied_files = []
    copy_errors = 0
    logging.info("=" * 72)
    logging.info(f"Copying {len(winners)} file(s) to {FILES_FOLDER}")
    for w in winners:
        destination = FILES_FOLDER / w["name"]
        if destination.exists():
            stem = Path(w["name"]).stem
            suffix = Path(w["name"]).suffix
            counter = 1
            while True:
                destination = FILES_FOLDER / f"{stem}_{counter}{suffix}"
                if not destination.exists():
                    break
                counter += 1
            logging.info(
                f"  Name conflict with existing file in output; "
                f"renamed to: {destination.name}")
        try:
            shutil.copy2(w["source_long"], _to_long(destination))
            mtime_str = _fmt_mtime(w["mtime"])
            copied_files.append(
                (destination.name, w["source_folder"], mtime_str,
                 _fmt_name_date(w.get("name_date")),
                 _fmt_version(w.get("version", (-1,))),
                 w.get("env", "")))
            logging.info(
                f"  Copied {w['name']} (modified {mtime_str}) "
                f"from {w['source_folder']} -> {destination.name}")
        except Exception as e:
            copy_errors += 1
            logging.error(
                f"  Copy FAILED for {_from_long(w['source_long'])}: {e}")
            _record_failure(
                stage="copy",
                name=w["name"],
                reason_code=type(e).__name__,
                detail=f"shutil.copy2 failed: {e}",
                input_path=_from_long(w["source_long"]),
                tb=traceback.format_exc(),
            )

    logging.info("=" * 72)
    logging.info("Stage 1 Summary")
    logging.info(f"  Folders scanned         : {folders_scanned}")
    logging.info(f"  Folders with Excel files: {folders_with_excel}")
    logging.info(f"  Folders without Excel   : {len(folders_without_excel)}")
    logging.info(f"  Excel files found       : {len(all_excel)}")
    logging.info(f"  Matrices selected       : {len(candidates)}")
    logging.info(f"  Possible duplicates     : {len(possible_dups)} pair(s)")
    logging.info(f"  Duplicates skipped      : {len(losers)}")
    logging.info(f"  UAT/temp/older not used : {len(not_selected)}")
    logging.info(f"  Files copied            : {len(copied_files)}")
    logging.info(f"  Copy errors             : {copy_errors}")
    logging.info("=" * 72)

    try:
        _write_xlsx_report(copied_files, folders_without_excel, losers,
                           not_selected, possible_dups)
        logging.info(f"Report written: {REPORT_FILE}")
    except Exception as e:
        logging.error(f"Failed to write xlsx report: {e}")

    return {
        "stage": "copy",
        "ok": True,
        "folders_scanned": folders_scanned,
        "candidates": len(candidates),
        "copied": len(copied_files),
        "duplicates_skipped": len(losers),
        "not_selected": len(not_selected),
        "copy_errors": copy_errors,
    }


# ============================================================================
# STAGE 2: PROCESS  —  FILES_FOLDER (xlsx) → CSV_FOLDER (csv)
# ============================================================================
# Reads the matrix sheet of each xlsx (auto-detected via BLK_ROLEDETAIL marker
# in candidate positions 2 -> 1 -> 3), applies red-fill / white-text rules to
# columns 3, 4, and every column from 8 onwards (cols 1, 2, 5, 6, 7 are
# passed through verbatim), drops hidden rows/columns, and writes one CSV
# per xlsx as "<stem>_sheet2.csv" into CSV_FOLDER.
#
# Robustness note: some xlsx files reference drawings/images that aren't
# actually present in the zip (e.g. dangling 'xl/drawings/NULL'). openpyxl
# raises KeyError when loading such files. We do NOT retry in read_only
# mode (that path is very slow on large workbooks): the file is recorded
# as a 'CorruptDrawings' failure with a clear root cause, and the user
# is expected to repair the xlsx (open in Excel, Save As a fresh copy)
# and re-run. The failure ends up in the run folder's failures/ subdir.


def _is_red_fill(cell):
    """True if the cell has a red background fill (any FF****0000 variant)."""
    if cell.fill and cell.fill.patternType != "none":
        if cell.fill.start_color and cell.fill.start_color.rgb:
            color = str(cell.fill.start_color.rgb).upper()
            if "FF" in color and "0000" in color[-4:]:
                return True
    return False


def _is_white_text(cell):
    """True if the cell's font colour renders as white.
    Handles rgb, theme, and indexed colour types."""
    if not cell.font or not cell.font.color:
        return False
    color = cell.font.color
    ctype = getattr(color, "type", None)
    if ctype == "rgb":
        rgb = color.rgb
        if not isinstance(rgb, str):
            return False
        return rgb.upper() in ("FFFFFF", "FFFFFFFF", "00FFFFFF")
    if ctype == "theme":
        return color.theme == 0   # theme 0 = "Background 1" = white
    if ctype == "indexed":
        return color.indexed in (1, 9)  # both 00FFFFFF in legacy palette
    return False


# --- Matrix-sheet detection (stage 2) ---------------------------------------
# Stage 3 needs the sheet that contains the access-matrix data, and it
# identifies that sheet by the literal cell value "BLK_ROLEDETAIL". Picking
# the 2nd visible sheet blindly is fragile: some files have the cover/legend
# on sheet 2 and the matrix on sheet 1 (or sheet 3 for upgrade docs). So we
# scan the first 10x10 cells of a few candidate sheets to find the right one.
#
# Order of candidates: sheet 2 → sheet 1 → sheet 3. The 2nd visible sheet is
# the most common location, so it's tried first; sheet 1 is next (some files
# have only one usable tab); sheet 3 catches the upgrade-doc layout. If none
# of the three contains BLK_ROLEDETAIL in the top-left 10x10, the file is
# skipped — stage 3 wouldn't be able to parse it anyway.

_MATRIX_MARKER = "BLK_ROLEDETAIL"
_MATRIX_SCAN_ROWS = 10
_MATRIX_SCAN_COLS = 10
_MATRIX_CANDIDATE_POSITIONS = (2, 1, 3)  # 1-based, in priority order


def _scan_for_matrix_marker(sheet):
    """Return True if `_MATRIX_MARKER` appears in the first
    _MATRIX_SCAN_ROWS rows x _MATRIX_SCAN_COLS columns of `sheet`."""
    max_r = min(_MATRIX_SCAN_ROWS, sheet.max_row or 0)
    max_c = min(_MATRIX_SCAN_COLS, sheet.max_column or 0)
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            v = sheet.cell(row=r, column=c).value
            if v is None:
                continue
            if str(v).strip() == _MATRIX_MARKER:
                return True
    return False


def _select_matrix_sheet(wb, visible_sheets):
    """Pick the worksheet that contains the access-matrix data.

    Walks `_MATRIX_CANDIDATE_POSITIONS` against `visible_sheets`, scans each
    candidate's top-left 10x10 for `_MATRIX_MARKER`, and returns the first
    match.

    Returns:
        (worksheet, position_1based, [(pos, name, found_bool), ...])
        or (None, None, [...])  when no candidate matches.
    """
    checked = []
    for pos in _MATRIX_CANDIDATE_POSITIONS:
        idx = pos - 1
        if idx >= len(visible_sheets):
            continue
        name = visible_sheets[idx]
        sheet = wb[name]
        found = _scan_for_matrix_marker(sheet)
        checked.append((pos, name, found))
        if found:
            return sheet, pos, checked
    return None, None, checked


def _process_one_xlsx(input_excel, output_csv):
    """Convert the access-matrix sheet of input_excel to a cleaned CSV.

    The matrix sheet is picked by scanning the first 10x10 cells of
    candidate sheets (in order: 2nd visible -> 1st visible -> 3rd visible)
    for the literal value "BLK_ROLEDETAIL". The first candidate that
    contains the marker is converted. If none match, the file is skipped
    with error_type='NoMatrixMarker'.

    Returns a dict:
        {'success': bool, 'input': str, 'output': str | None,
         'error_type': str | None, 'error_message': str | None,
         'stats': {...}}
    """
    result = {
        "success": False,
        "input": input_excel,
        "output": None,
        "error_type": None,
        "error_message": None,
        "stats": {},
    }

    print(f"Converting SECOND SHEET of: {input_excel}")

    if not os.path.exists(input_excel):
        result["error_type"] = "FileNotFound"
        result["error_message"] = f"File does not exist: {input_excel}"
        print(f"  ERROR: {result['error_message']}")
        print("-" * 60)
        return result
    if not os.path.isfile(input_excel):
        result["error_type"] = "NotAFile"
        result["error_message"] = f"Path is not a file: {input_excel}"
        print(f"  ERROR: {result['error_message']}")
        print("-" * 60)
        return result
    if os.path.getsize(input_excel) == 0:
        result["error_type"] = "EmptyFile"
        result["error_message"] = "File is 0 bytes (empty)"
        print(f"  ERROR: {result['error_message']}")
        print("-" * 60)
        return result

    try:
        try:
            wb = openpyxl.load_workbook(input_excel, data_only=True)
        except InvalidFileException as e:
            result["error_type"] = "InvalidExcelFile"
            result["error_message"] = (
                f"Not a valid .xlsx/.xlsm file (openpyxl can't read .xls — "
                f"convert to .xlsx first). Details: {e}"
            )
            print(f"  ERROR: {result['error_message']}")
            print("-" * 60)
            return result
        except PermissionError as e:
            result["error_type"] = "PermissionDenied"
            result["error_message"] = (
                f"Permission denied opening file (is it open in Excel?). Details: {e}"
            )
            print(f"  ERROR: {result['error_message']}")
            print("-" * 60)
            return result
        except zipfile.BadZipFile as e:
            result["error_type"] = "CorruptFile"
            result["error_message"] = f"File appears corrupt (bad zip): {e}"
            print(f"  ERROR: {result['error_message']}")
            print("-" * 60)
            return result
        except KeyError as e:
            # Some xlsx files have a dangling drawing/image relationship in
            # their internal XML — e.g. <drawing r:id="rId1"/> pointing to a
            # part like 'xl/drawings/NULL' that isn't actually in the zip.
            # openpyxl raises KeyError when trying to load such drawings.
            # We do NOT retry in read_only mode here because that path is
            # very slow on large workbooks and prone to hanging the run.
            # The file is recorded as a failure with a clear root cause so
            # the user can repair it (open in Excel, save as a new .xlsx,
            # re-run) and try again.
            result["error_type"] = "CorruptDrawings"
            result["error_message"] = (
                f"Workbook has a broken drawing/image reference "
                f"({e}). The .xlsx references a part that isn't present "
                f"in the zip archive (typically 'xl/drawings/NULL'). "
                f"To fix: open the file in Excel, save it as a new .xlsx "
                f"to repair the internal relationships, then re-run the "
                f"pipeline."
            )
            print(f"  ERROR: {result['error_message']}")
            print("-" * 60)
            return result

        visible_sheets = [name for name in wb.sheetnames
                          if wb[name].sheet_state == "visible"]
        hidden_sheets = [name for name in wb.sheetnames
                         if wb[name].sheet_state != "visible"]
        if hidden_sheets:
            print(
                f"  Skipping {len(hidden_sheets)} hidden sheet(s): {hidden_sheets}")
        if not visible_sheets:
            result["error_type"] = "NoVisibleSheets"
            result["error_message"] = (
                f"No visible sheets found (workbook has {len(wb.sheetnames)} "
                f"sheet(s), all hidden)"
            )
            print(f"  ERROR: {result['error_message']}")
            print("-" * 60)
            return result

        # Pick the matrix sheet by scanning the first 10x10 cells for
        # "BLK_ROLEDETAIL" in candidate order (2nd -> 1st -> 3rd visible).
        # If none match, skip the file rather than convert noise.
        ws, selected_position, checked = _select_matrix_sheet(
            wb, visible_sheets)
        print(f"  BLK_ROLEDETAIL scan (top {_MATRIX_SCAN_ROWS}x"
              f"{_MATRIX_SCAN_COLS}) across candidates:")
        for pos, name, found in checked:
            mark = "✓" if found else "·"
            print(f"    {mark} position {pos}: {name!r}")
        if ws is None:
            result["error_type"] = "NoMatrixMarker"
            result["error_message"] = (
                f"'{_MATRIX_MARKER}' not found in the first "
                f"{_MATRIX_SCAN_ROWS}x{_MATRIX_SCAN_COLS} cells of any candidate "
                f"sheet (checked positions "
                f"{[p for p, _, _ in checked]}). Skipping."
            )
            print(f"  SKIP: {result['error_message']}")
            print("-" * 60)
            return result

        used_fallback = (selected_position != 2)
        print(f"  Sheet: {ws.title} "
              f"(position {selected_position} of {len(visible_sheets)} visible — "
              f"selected by BLK_ROLEDETAIL scan)")
        print(f"  Original: {ws.max_row} rows, {ws.max_column} columns")

        visible_rows = []
        visible_columns = []
        for row in range(1, ws.max_row + 1):
            row_dim = ws.row_dimensions.get(row)
            if not row_dim or not row_dim.hidden:
                visible_rows.append(row)
        for col in range(1, ws.max_column + 1):
            col_letter = get_column_letter(col)
            col_dim = ws.column_dimensions.get(col_letter)
            if not col_dim or not col_dim.hidden:
                visible_columns.append(col)
        print(
            f"  Visible: {len(visible_rows)} rows, {len(visible_columns)} columns")

        if not visible_rows or not visible_columns:
            result["error_type"] = "EmptySheet"
            result["error_message"] = (
                f"Sheet '{ws.title}' has no visible rows/columns "
                f"({len(visible_rows)} rows, {len(visible_columns)} cols)"
            )
            print(f"  ERROR: {result['error_message']}")
            print("-" * 60)
            return result

        csv_data = []
        white_text_cleared_count = 0
        red_fill_converted_count = 0
        red_fill_cleared_count = 0

        for row_idx in visible_rows:
            row_data = []
            for col_idx in visible_columns:
                cell = ws.cell(row=row_idx, column=col_idx)
                cell_value = cell.value

                # Formatting rules apply to columns 3, 4, and 8+. Columns
                # 1, 2, 5, 6, 7 are passed through unchanged.
                if col_idx in [3, 4] or col_idx > 7:
                    has_red = _is_red_fill(cell)
                    has_white = _is_white_text(cell)
                    if has_white:
                        cell_value = ""
                        white_text_cleared_count += 1
                    elif has_red:
                        if cell_value in [1, 0, "1", "0"]:
                            cell_value = "0"
                            red_fill_converted_count += 1
                        else:
                            cell_value = ""
                            red_fill_cleared_count += 1
                    elif cell_value is None:
                        cell_value = ""
                    else:
                        cell_value = str(cell_value)
                else:
                    cell_value = "" if cell_value is None else str(cell_value)

                cell_value = str(cell_value)
                if "\n" in cell_value or "\r" in cell_value:
                    cell_value = (cell_value
                                  .replace("\r\n", " ")
                                  .replace("\n", " ")
                                  .replace("\r", " "))
                row_data.append(cell_value)
            csv_data.append(row_data)

        out_dir = os.path.dirname(os.path.abspath(output_csv))
        os.makedirs(out_dir, exist_ok=True)
        try:
            with open(output_csv, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerows(csv_data)
        except PermissionError as e:
            result["error_type"] = "OutputPermissionDenied"
            result["error_message"] = (
                f"Could not write CSV (file open elsewhere or no write perms): {e}"
            )
            print(f"  ERROR: {result['error_message']}")
            print("-" * 60)
            return result
        except OSError as e:
            result["error_type"] = "OutputWriteFailed"
            result["error_message"] = f"Could not write CSV: {e}"
            print(f"  ERROR: {result['error_message']}")
            print("-" * 60)
            return result

        print(f"  Cleared {white_text_cleared_count} white text cells to ''")
        print(
            f"  Converted {red_fill_converted_count} red-filled 1/0 cells to '0'")
        print(
            f"  Cleared {red_fill_cleared_count} red-filled other-value cells to ''")
        print(f"  Created CSV: {output_csv}")
        print(f"  CSV size: {len(csv_data)} rows, "
              f"{len(csv_data[0]) if csv_data else 0} columns")
        print("-" * 60)

        result["success"] = True
        result["output"] = output_csv
        result["stats"] = {
            "sheet_name": ws.title,
            "used_fallback_sheet": used_fallback,
            "selected_position": selected_position,
            "rows": len(csv_data),
            "cols": len(csv_data[0]) if csv_data else 0,
            "white_text_cleared": white_text_cleared_count,
            "red_fill_converted": red_fill_converted_count,
            "red_fill_cleared": red_fill_cleared_count,
        }
        return result

    except Exception as e:
        result["error_type"] = type(e).__name__
        result["error_message"] = str(e) or repr(e)
        print(f"  ERROR ({result['error_type']}): {result['error_message']}")
        tb = traceback.extract_tb(e.__traceback__)
        if tb:
            last = tb[-1]
            print(f"    at {os.path.basename(last.filename)}:{last.lineno} "
                  f"in {last.name}")
        print("-" * 60)
        return result


def _print_stage2_summary(results, output_dir):
    successes = [r for r in results if r["success"]]
    failures = [r for r in results if not r["success"]]

    print()
    print("=" * 70)
    print("STAGE 2 SUMMARY")
    print("=" * 70)
    print(f"Total files processed : {len(results)}")
    print(f"  Successful          : {len(successes)}")
    print(f"  Failed              : {len(failures)}")
    print(f"Output directory      : {output_dir}")
    print()

    if successes:
        print("-" * 70)
        print(f"SUCCESSFUL ({len(successes)})")
        print("-" * 70)
        for r in successes:
            stats = r["stats"]
            name = os.path.basename(r["input"])
            pos = stats.get("selected_position")
            if pos == 2:
                note = ""
            elif pos in (1, 3):
                note = f" [matrix found on position {pos} — fallback from 2]"
            else:
                # Older runs (or unusual layouts) — keep the legacy hint
                note = " [used 1st sheet — only one visible]" if stats.get(
                    "used_fallback_sheet") else ""
            print(f"  ✓ {name}{note}")
            print(f"      sheet: {stats['sheet_name']!r:<30} "
                  f"-> {os.path.basename(r['output'])}")
            print(f"      rows: {stats['rows']}, cols: {stats['cols']}, "
                  f"white-cleared: {stats['white_text_cleared']}, "
                  f"red->0: {stats['red_fill_converted']}, "
                  f"red-cleared: {stats['red_fill_cleared']}")
        print()

    if failures:
        print("-" * 70)
        print(f"FAILED ({len(failures)})")
        print("-" * 70)
        by_type = {}
        for r in failures:
            by_type.setdefault(r["error_type"] or "Unknown", []).append(r)
        for err_type, items in by_type.items():
            print(f"  [{err_type}] — {len(items)} file(s)")
            for r in items:
                name = os.path.basename(r["input"])
                print(f"    ✗ {name}")
                print(f"        reason: {r['error_message']}")
            print()

    print("=" * 70)


def stage2_process():
    """Read every xlsx in FILES_FOLDER, write cleaned csv into CSV_FOLDER.
    Returns a stats dict for the pipeline summary."""
    _banner("STAGE 2 — PROCESS  (Files/*.xlsx -> CSV/*.csv)")

    if not FILES_FOLDER.exists():
        print(f"ERROR: Files folder does not exist: {FILES_FOLDER}")
        print("       Did stage 1 run successfully?")
        return {"stage": "process", "ok": False,
                "error": "files_folder missing"}

    CSV_FOLDER.mkdir(parents=True, exist_ok=True)
    print(f"Input directory : {FILES_FOLDER}")
    print(f"Output directory: {CSV_FOLDER}")
    print(f"Extensions      : {sorted(EXCEL_EXTENSIONS)}")
    print("=" * 70)

    excel_files = []
    for name in os.listdir(str(FILES_FOLDER)):
        if name.startswith("~$"):
            continue
        ext = os.path.splitext(name)[1].lower().strip()
        if ext not in EXCEL_EXTENSIONS:
            continue
        full = FILES_FOLDER / name
        if full.is_file():
            excel_files.append(full)
    excel_files.sort()
    excel_files = _drop_non_prod(excel_files, "process", print)

    if not excel_files:
        print(f"No Excel files found in: {FILES_FOLDER}")
        return {"stage": "process", "ok": True,
                "processed": 0, "successful": 0, "failed": 0}

    print(f"Found {len(excel_files)} Excel file(s):")
    for i, f in enumerate(excel_files, 1):
        print(f"  {i}. {f.name}")
    print("=" * 70)

    results = []
    for excel_file in excel_files:
        csv_path = CSV_FOLDER / f"{excel_file.stem}_sheet2.csv"
        results.append(_process_one_xlsx(str(excel_file), str(csv_path)))

    _print_stage2_summary(results, str(CSV_FOLDER))

    successful = sum(1 for r in results if r["success"])

    # Surface every per-file failure so the pipeline-wide FAILURES section
    # has a reason for each xlsx that couldn't be converted (no matrix
    # marker, unreadable workbook, write error, etc).
    for r in results:
        if r["success"]:
            continue
        _record_failure(
            stage="process",
            name=os.path.basename(r["input"]) if r.get(
                "input") else "(unknown)",
            reason_code=r.get("error_type") or "ProcessError",
            detail=r.get("error_message") or "(no error message captured)",
            input_path=r.get("input"),
            tb=None,
        )

    return {
        "stage": "process",
        "ok": True,
        "processed": len(results),
        "successful": successful,
        "failed": len(results) - successful,
    }


# ============================================================================
# STAGE 3: COMPARE  —  CSV_FOLDER (csv) → COMPARISON_FOLDER (diff xlsx)
# ============================================================================

# Per-file logger; re-initialised inside _process_one_csv() for each csv.
log = None


def _stage3_setup_logger(log_path=None):
    """Set up a logger that writes to console + (optional) per-file debug log."""
    logger = logging.getLogger("stage3")
    logger.setLevel(logging.DEBUG)
    logger.propagate = False  # don't double-print via root logger from stage 1

    if logger.handlers:
        logger.handlers.clear()

    fmt = logging.Formatter(
        "%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S")
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    if log_path:
        try:
            fh = logging.FileHandler(log_path, encoding="utf-8")
            fh.setLevel(logging.DEBUG)
            fh.setFormatter(fmt)
            logger.addHandler(fh)
        except (OSError, FileNotFoundError) as exc:
            logger.warning(
                f"Could not open log file '{log_path}' "
                f"({exc.__class__.__name__}: {exc}). "
                f"Continuing with console-only logging.")
    return logger


# --- Function-ID label tolerance --------------------------------------------
# Different matrices label the function-ID row differently inside a
# BLK_ROLEDETAIL block. Accept both spellings everywhere we look up that row.
FN_ID_LABELS = ("ROLE_FUNCTION", "FUNCTION_ID", "ROLE.FUNCTION", "FUNCTION.ID")
REPORT_WIDTH = 84


def _pick_fn_id_row(fn_to_row):
    for label in FN_ID_LABELS:
        if label in fn_to_row:
            return fn_to_row[label]
    return None


# --- Exclusion list (orphan ROLE_FUNCTIONs with no SMTB_MENU entry) ---------
_EXCLUDED_FUNCTION_IDS_CACHE = None
_EXCLUSION_QUERY = """

select distinct FUNCTION_ID from FCPROD.SMTBS_MENU where function_id not in (SELECT FUNCTION_ID FROM (SELECT A.FUNCTION_ID FUNCTION_ID, (SELECT DESCRIPTION FROM FCPROD.SMTB_FUNCTION_DESCRIPTION
             WHERE FUNCTION_ID = A.FUNCTION_ID and LANG_CODE like 'ENG') BALOON_HELP 
            FROM FCPROD.SMTBS_MENU A, FCPROD.SMTBS_MODULES C 
            WHERE A.TYPE_STRING IN ('SM','M', 'D','BOTG','ELM','LNC','LNM','PMM','GQ','MSP','R', 'U', 'ELR','B','P', 'T','O', 'ELO', 'PMO') 
            AND A.MODULE = C.MODULE_ID and a.RECORD_STAT = 'O' AND a.ONCE_AUTH = 'Y' ) where 1=1 )

"""


def _get_excluded_function_ids(uat_config):
    global _EXCLUDED_FUNCTION_IDS_CACHE
    if _EXCLUDED_FUNCTION_IDS_CACHE is not None:
        log.debug(f"  [EXCLUDE] Using cached exclusion list "
                  f"({len(_EXCLUDED_FUNCTION_IDS_CACHE)} fn_ids).")
        return _EXCLUDED_FUNCTION_IDS_CACHE

    log.info("  [EXCLUDE] Fetching ROLE_FUNCTION exclusion list from DB …")
    try:
        conn = oracledb.connect(
            user=uat_config["user"],
            password=uat_config["password"],
            dsn=uat_config["dsn"],
        )
        _q = _apply_db_link(_EXCLUSION_QUERY)
        if DB_LINK:
            log.debug(f"  [EXCLUDE] SQL (with db_link={DB_LINK}):\n{_q}")
        df = pd.read_sql(_q, con=conn)
        conn.close()
        excluded = {
            str(v).strip().upper()
            for v in df.iloc[:, 0].dropna().tolist()
            if str(v).strip()
        }
        log.info(
            f"  [EXCLUDE] Loaded {len(excluded)} function IDs to exclude.")
        if excluded:
            sample = sorted(excluded)[:10]
            tail = " …" if len(excluded) > 10 else ""
            log.debug(f"  [EXCLUDE] Sample: {sample}{tail}")
        _EXCLUDED_FUNCTION_IDS_CACHE = excluded
        return excluded
    except Exception:
        log.error("  [EXCLUDE] FAILED to fetch exclusion list — "
                  "proceeding WITHOUT filter.\n" + traceback.format_exc())
        _EXCLUDED_FUNCTION_IDS_CACHE = set()
        return _EXCLUDED_FUNCTION_IDS_CACHE


# --- Validation registry (Other Values / GL value lookups) ------------------
_VALIDATION_QUERIES = {
    "PRODUCT_CODE": """
        SELECT DISTINCT PRODUCT_CODE FROM fcprod.CSTM_PRODUCT WHERE RECORD_STAT = 'O' AND ONCE_AUTH = 'Y'
    """,
    "ROLESOURCE": """
        SELECT DISTINCT SOURCE_CODE
        FROM   FCPROD.PMZM_SOURCE
        WHERE  RECORD_STAT = 'O'
        AND    AUTH_STAT   = 'A'
    """,
    "GRP_RESTRICTION": """
        SELECT DISTINCT GROUP_CODE
        FROM   FCPROD.STTM_GROUP_CODE
        WHERE  GROUP_TYPE IN ('C','I')
        AND    RECORD_STAT = 'O'
        AND    AUTH_STAT   = 'A'
    """,
    "ROLEACCCLASS": "SELECT DISTINCT ACCOUNT_CLASS FROM FCPROD.STTMS_ACCOUNT_CLASS where record_stat = 'O'",
    "QUEUERIGHTS": """
        SELECT DISTINCT QUEUE
        FROM   FCPROD.SMTB_QUEUES
        WHERE  RECORD_STAT = 'O'
        AND    ONCE_AUTH   = 'Y'
    """,
    "NODE_GL": """
        SELECT DISTINCT GL_CODE FROM (
            SELECT GL_CODE, GL_DESC FROM FCPROD.GLZM_GLMASTER WHERE LEAF = 'N'
            UNION
            SELECT GL_CODE, GL_DESC FROM FCPROD.GLZM_GLMASTER
            WHERE  LEAF = 'Y' AND PARENT_GL = '0'
        )
    """,
    "LEAF_GL": """
        SELECT DISTINCT GL_CODE
        FROM (
            SELECT GL_CODE, GL_DESC,
                   DECODE(LEAF,'Y','LEAF','NODE') LEAF
            FROM   FCPROD.GLZM_GLMASTER
        )
    """,
}
_VALIDATION_CACHES = {}


def _get_validation_set(uat_config, key):
    if key in _VALIDATION_CACHES:
        log.debug(f"  [VALIDATE:{key}] Using cached set "
                  f"({len(_VALIDATION_CACHES[key])} values).")
        return _VALIDATION_CACHES[key]
    query = _VALIDATION_QUERIES.get(key)
    if not query:
        log.warning(f"  [VALIDATE:{key}] No query registered — empty set.")
        _VALIDATION_CACHES[key] = set()
        return _VALIDATION_CACHES[key]
    log.info(f"  [VALIDATE:{key}] Fetching valid values from DB …")
    try:
        conn = oracledb.connect(
            user=uat_config["user"],
            password=uat_config["password"],
            dsn=uat_config["dsn"],
        )
        _q = _apply_db_link(query)
        if DB_LINK:
            log.debug(
                f"  [VALIDATE:{key}] SQL (with db_link={DB_LINK}):\n{_q}")
        df = pd.read_sql(_q, con=conn)
        conn.close()
        values = {
            str(v).strip().upper()
            for v in df.iloc[:, 0].dropna().tolist()
            if str(v).strip()
        }
        log.info(f"  [VALIDATE:{key}] Loaded {len(values)} values.")
        if values:
            sample = sorted(values)[:10]
            tail = " …" if len(values) > 10 else ""
            log.debug(f"  [VALIDATE:{key}] Sample: {sample}{tail}")
        _VALIDATION_CACHES[key] = values
        return values
    except Exception:
        log.error(f"  [VALIDATE:{key}] FAILED to fetch list — column will be "
                  f"empty for this source.\n" + traceback.format_exc())
        _VALIDATION_CACHES[key] = set()
        return _VALIDATION_CACHES[key]


def _annotate_validation(df, value_col, validation_key, uat_config,
                         label="VALIDATE"):
    try:
        valid_set = _get_validation_set(uat_config, validation_key)

        def _check(v):
            if pd.isna(v):
                return ""
            s = str(v).strip().upper()
            if not s:
                return ""
            return ("does fetch on system" if s in valid_set
                    else "does not fetch on system")

        if df.empty:
            df["value validation"] = pd.Series(dtype=object)
        else:
            df["value validation"] = df[value_col].map(_check)
            vc = df["value validation"].value_counts(dropna=False)
            log.info(f"  [{label}:{validation_key}] {vc.to_dict()}")
    except Exception:
        log.error(f"  [{label}:{validation_key}] FAILED — writing empty.\n"
                  + traceback.format_exc())
        if "value validation" not in df.columns:
            df["value validation"] = ""
    return df


def _section(title):
    bar = "─" * REPORT_WIDTH
    log.info(f"\n{bar}\n  ▶ {title}\n{bar}")


def _df_snapshot(name, df):
    displayed_cols = list(df.columns[:10])
    col_note = (f"{displayed_cols} … (+{len(df.columns) - 10} more)"
                if len(df.columns) > 10 else f"{displayed_cols}")
    log.debug(f"[DF SNAPSHOT] {name} → shape={df.shape}, columns={col_note}")
    if df.empty:
        log.debug("  (empty DataFrame)")


# --- Stage 3: summary/table formatting --------------------------------------

def _hr(char="─", width=REPORT_WIDTH):
    return char * width


def _format_table(headers, rows, footer=None):
    widths = [len(str(h)) for h in headers]
    for row in rows + ([footer] if footer else []):
        for i, cell in enumerate(row):
            widths[i] = max(widths[i], len(str(cell)))
    widths = [w + 2 for w in widths]

    def _line(left, mid, right, fill):
        return left + mid.join(fill * w for w in widths) + right

    def _row(cells):
        out = "│"
        for i, c in enumerate(cells):
            txt = str(c)
            if i == 0:
                out += " " + txt.ljust(widths[i] - 1)
            else:
                out += txt.center(widths[i])
            out += "│"
        return out

    top = _line("┌", "┬", "┐", "─")
    sep = _line("├", "┼", "┤", "─")
    bot = _line("└", "┴", "┘", "─")
    out = [top, _row(headers), sep]
    for r in rows:
        out.append(_row(r))
    if footer:
        out.append(sep)
        out.append(_row(footer))
    out.append(bot)
    return "\n".join(out)


_DIFF_LABEL = {
    "In access matrix not in system": "In matrix, not in system",
    "In system not in access matrix": "In system, not in matrix",
    "Value in matrix_diff":           "Value mismatch",
}

_SECTIONS = [
    ("Screens",   "df_diff_screens"),
    ("Flags",     "df_diff_flags"),
    ("GL Nodes",  "df_diff_gl_nodes"),
    ("GL Leaves", "df_diff_gl_leaves"),
    ("Msg Rgts",  "df_diff_msg_rights"),
    ("OtherVals", "df_diff_values"),
]


def _compute_role_stats(diffs_by_section, roles):
    stats = {role: {} for role in roles}
    for section, df in diffs_by_section.items():
        for role in roles:
            stats[role].setdefault(section, {"total": 0})
        if df is None or len(df) == 0 or "Role_id" not in df.columns:
            continue
        for role in roles:
            sub = df[df["Role_id"] == role]
            if sub.empty:
                continue
            counts = {}
            if "difference_type" in sub.columns:
                vc = sub["difference_type"].value_counts(
                    dropna=False).to_dict()
                for k, v in vc.items():
                    label = _DIFF_LABEL.get(str(k).strip(), str(k).strip())
                    counts[label] = int(v)
            counts["total"] = int(len(sub))
            stats[role][section] = counts
    return stats


def _role_total(role_stats):
    return sum(s.get("total", 0) for s in role_stats.values())


def _format_counts_table(stats):
    headers = ["Role ID"] + [s for s, _ in _SECTIONS] + ["Total"]
    rows, col_totals, grand_total = [], [0] * len(_SECTIONS), 0
    for role in sorted(stats.keys()):
        rs = stats[role]
        row, role_total = [role], 0
        for i, (section, _) in enumerate(_SECTIONS):
            n = rs.get(section, {}).get("total", 0)
            row.append(str(n) if n else "·")
            col_totals[i] += n
            role_total += n
        row.append(str(role_total) if role_total else "·")
        grand_total += role_total
        rows.append(row)
    footer = (["TOTAL"]
              + [str(t) if t else "·" for t in col_totals]
              + [str(grand_total)])
    return _format_table(headers, rows, footer=footer)


def _format_role_narrative(stats):
    lines = []
    for role in sorted(stats.keys()):
        rs = stats[role]
        total = _role_total(rs)
        lines.append("")
        lines.append(f"  Role: {role}")
        if total == 0:
            lines.append(
                "    Status: ✓ Matched cleanly — no differences detected.")
            continue
        lines.append(f"    Status: ⚠ {total} difference"
                     f"{'s' if total != 1 else ''} detected")
        for section, _ in _SECTIONS:
            ss = rs.get(section, {})
            sec_total = ss.get("total", 0)
            if sec_total == 0:
                continue
            lines.append(f"      • {section}  ({sec_total})")
            for k, v in ss.items():
                if k == "total":
                    continue
                lines.append(f"          ◦ {k:<28} {v}")
    return "\n".join(lines)


def _build_file_summary_text(base_name, stats, output_excel):
    total_roles = len(stats)
    grand_total = sum(_role_total(s) for s in stats.values())
    clean_roles = sum(1 for s in stats.values() if _role_total(s) == 0)
    diff_roles = total_roles - clean_roles
    status_icon = "✓" if grand_total == 0 else "⚠"

    parts = [
        "═" * REPORT_WIDTH,
        f"   FILE SUMMARY  —  {base_name}",
        "═" * REPORT_WIDTH,
        "",
        f"   Status              : {status_icon} "
        + ("All roles matched" if grand_total == 0
           else "Differences require review"),
        f"   Roles processed     : {total_roles}",
        f"   Roles matched clean : {clean_roles}",
        f"   Roles with diffs    : {diff_roles}",
        f"   Total differences   : {grand_total}",
        f"   Output workbook     : {output_excel}",
        "",
        _hr(),
        "   PER-ROLE DIFFERENCE COUNTS",
        _hr(),
        "",
        _format_counts_table(stats),
        "",
        _hr(),
        "   PER-ROLE DETAIL",
        _hr(),
        _format_role_narrative(stats),
        "",
        "═" * REPORT_WIDTH,
    ]
    return "\n".join(parts)


def _save_file_summary(text, base_name, debug_dir):
    path = os.path.join(debug_dir, f"{base_name}_role_summary.txt")
    try:
        with open(path, "w", encoding="utf-8") as f:
            f.write(text + "\n")
        log.info(f"  [SUMMARY] Saved role summary → {path}")
        return path
    except Exception:
        log.warning("  [SUMMARY] Could not save role summary.\n"
                    + traceback.format_exc())
        return None


def _format_stage3_run_summary(results):
    total = len(results)
    succ = sum(1 for r in results if r["status"].startswith("Success"))
    warn = sum(1 for r in results if r.get("warning"))
    fail = total - succ
    total_roles = sum(len(r.get("stats") or {}) for r in results)
    total_diffs = sum(
        sum(_role_total(s) for s in (r.get("stats") or {}).values())
        for r in results
    )

    parts = ["",
             "═" * REPORT_WIDTH,
             "   STAGE 3 SUMMARY",
             "═" * REPORT_WIDTH,
             "",
             f"   Files processed   : {total}",
             f"   Successful        : {succ}  ✓"]
    if warn:
        parts.append(
            f"   With warnings     : {warn}  ⚠  "
            f"(processed cleanly; post-processing note — see log)")
    parts.append(f"   Failed            : {fail}  {'✗' if fail else ' '}")
    parts.append(f"   Total roles       : {total_roles}")
    parts.append(f"   Total differences : {total_diffs}")
    parts.append("")

    headers = ["File", "Status", "Roles", "Diffs"]
    rows = []
    for r in results:
        base = os.path.basename(r["input"])
        if len(base) > 50:
            base = base[:47] + "…"
        status_raw = r["status"]
        if status_raw.startswith("Success"):
            stats = r.get("stats") or {}
            n_roles = len(stats)
            n_diffs = sum(_role_total(s) for s in stats.values())
            label = "⚠ Success (warn)" if r.get("warning") else "✓ Success"
            rows.append([base, label, str(n_roles),
                         str(n_diffs) if n_diffs else "·"])
        else:
            reason = status_raw.replace("Failed: ", "")
            if len(reason) > 40:
                reason = reason[:37] + "…"
            rows.append([base, f"✗ {reason}", "-", "-"])
    parts.append(_format_table(headers, rows))
    if warn:
        parts.extend([
            "",
            "   Files marked ⚠ produced their comparison.xlsx in",
            "   all_differences/ normally. The warning refers to",
            "   post-processing housekeeping (typically the debug-",
            "   folder move hitting Windows' 260-char path limit).",
            "   Check the debug log of each ⚠ file for details.",
        ])
    parts.extend(["", "═" * REPORT_WIDTH])
    return "\n".join(parts)


# --- Stage 3: per-CSV processing --------------------------------------------

def _process_one_csv(input_csv_path, shared_dirs=None):
    """Process a single CSV file and generate a comparison xlsx.
    Returns (output_excel_path, debug_dir, stats_dict)."""
    global log
    csv_dir = os.path.dirname(os.path.abspath(input_csv_path))
    base_name = os.path.splitext(os.path.basename(input_csv_path))[0]

    debug_dir = os.path.join(csv_dir, f"{base_name}_debug")
    os.makedirs(debug_dir, exist_ok=True)
    log_file = os.path.join(debug_dir, f"{base_name}_debug.log")
    log = _stage3_setup_logger(log_file)

    _section(f"Processing file: {input_csv_path}")
    log.info(f"Debug folder : {debug_dir}")
    log.info(f"Log file     : {log_file}")

    # ----------------------------------------------------------
    # STEP 0: Read CSV
    # ----------------------------------------------------------
    _section("STEP 0 – Read CSV")
    try:
        df_matrix = pd.read_csv(input_csv_path)
        log.info(f"CSV loaded successfully → shape={df_matrix.shape}")
        log.debug(f"Columns: {list(df_matrix.columns)}")
    except Exception:
        log.error(
            f"[file: {base_name}] FAILED to read CSV file.\n" + traceback.format_exc())
        raise

    df_screens = pd.DataFrame(columns=[
        "Role_id", "Function_id", "New", "Copy", "Delete", "Close", "Unlock",
        "Reopen", "Print", "Auth", "Reverse", "Rollover", "Confirm", "Liquidate",
        "Hold", "Template", "View", "Generate"
    ])
    df_gl_nodes = pd.DataFrame(columns=["Role_id", "Node_GL"])
    df_gl_leaves = pd.DataFrame(columns=["Role_id", "Node_GL", "Leaf_GL"])
    df_other_values = pd.DataFrame(columns=["Role_id", "Type", "Value"])
    df_msg_rights = pd.DataFrame(columns=[
        "Role_id", "Cancel", "Change_Node", "Release", "Change_Media", "Branch_Move", "Hold",
        "Test_Input", "Change_Address", "Reinstate", "Change_Priority", "Auth_Cancel", "Auth_Change_Node",
        "Release_auth", "Change_Media_Auth", "Auth_Branch_Move", "Hold_Auth", "Auth_Test_Input",
        "Change_Address_Auth", "Auth_Reinstate", "Change_Priority_Auth", "Install", "Test_Check",
        "Link_Contract", "Change_Branch_In", "Change_Message", "Change_Force_Release_Fund",
        "Suppress", "Delete", "Print", "FT_Upload", "Move_To_Queue", "Change_Address_In",
        "Auth_Change_Message", "Auth_Rights", "Change_Force_Cover_Match"
    ])

    # ----------------------------------------------------------
    # STEP 1: Oracle thick-mode init
    # ----------------------------------------------------------
    _section("STEP 1 – Oracle client init")
    try:
        log.info(
            f"Attempting oracledb thick-mode init with lib_dir={ORACLE_LIB_DIR}")
        oracledb.init_oracle_client(lib_dir=ORACLE_LIB_DIR)
        log.info("Oracle client initialised successfully.")
    except Exception:
        already_init_msg = "Oracle Client library has already been initialized"
        exc_text = traceback.format_exc()
        if already_init_msg in exc_text:
            log.info("Oracle client already initialised (continuing).")
        else:
            log.error(
                f"[file: {base_name}] FAILED to init Oracle client.\n" + exc_text)
            raise

    uat_config = dict(UAT_CONFIG)

    # ----------------------------------------------------------
    # STEP 2: Identify key columns
    # ----------------------------------------------------------
    _section("STEP 2 – Identify key columns in CSV")
    try:
        log.debug("Searching for 'BLK_ROLEDETAIL' in df_matrix …")
        matched_cols = df_matrix.columns[(df_matrix == "BLK_ROLEDETAIL").any()]
        log.debug(f"Columns containing 'BLK_ROLEDETAIL': {list(matched_cols)}")
        if len(matched_cols) == 0:
            raise ValueError("No column found that contains the value 'BLK_ROLEDETAIL'. "
                             "Check that the CSV normalisation step has already run, "
                             "or that the sentinel value exists in the raw file.")
        type_column = matched_cols[0]
        type_column_index = df_matrix.columns.get_loc(type_column)
        role_id_column_index = type_column_index - 4
        role_id_column = df_matrix.columns[role_id_column_index]
        log.info(f"type_column='{type_column}' (index {type_column_index}), "
                 f"role_id_column='{role_id_column}' (index {role_id_column_index})")
    except Exception:
        log.error(
            f"[file: {base_name}] FAILED to identify key columns.\n" + traceback.format_exc())
        raise

    # ----------------------------------------------------------
    # STEP 3: Normalise df_matrix
    # ----------------------------------------------------------
    _section("STEP 3 – Normalise df_matrix")
    try:
        BLK_TO_SCREEN = {
            "BLK_ROLEDETAIL":  "MAINTENANCE",
            "BLK_ROLEDETAILB": "REPORTS",
            "BLK_ROLEDETAILC": "BATCH",
            "BLK_ROLEDETAILD": "ONLINE",
        }
        SCREEN_LABEL_MAP = {
            "MAINTENANCE":           "MAINTENANCE",
            "REPORTS":               "REPORTS",
            "REPORT":                "REPORTS",
            "BATCH":                 "BATCH",
            "ONLINE":                "ONLINE",
            "WEB_BRANCH":            "WEB_BRANCH",
            "WEB.BRANCH":            "WEB_BRANCH",
            "WEBBRANCH":             "WEB_BRANCH",
            "PROCESS_STAGE_RIGHTS":  "PROCESS_STAGE_RIGHTS",
            "PROCESS.STAGE.RIGHTS":  "PROCESS_STAGE_RIGHTS",
            "PROCESSSTAGERIGHTS":    "PROCESS_STAGE_RIGHTS",
            "GRANT_RIGHTS":          "GRANT_RIGHTS",
            "GRANT.RIGHTS":          "GRANT_RIGHTS",
        }
        SCREEN_LABEL_MAP_SCREENS = {
            k: v for k, v in SCREEN_LABEL_MAP.items()
            if v in {"MAINTENANCE", "REPORTS", "BATCH", "ONLINE",
                     "WEB_BRANCH", "PROCESS_STAGE_RIGHTS"}
        }

        eff_col = []
        current_screen = None
        for _, row in df_matrix.iterrows():
            raw_type = str(row[type_column]).strip(
            ).upper() if pd.notna(row[type_column]) else ""
            raw_type = raw_type.replace(" ", "").replace(".", "")
            if raw_type in BLK_TO_SCREEN:
                current_screen = BLK_TO_SCREEN[raw_type]
                eff_col.append(current_screen)
            elif raw_type in SCREEN_LABEL_MAP_SCREENS:
                current_screen = SCREEN_LABEL_MAP_SCREENS[raw_type]
                eff_col.append(current_screen)
            elif raw_type == "":
                eff_col.append(
                    current_screen if current_screen is not None else pd.NA)
            else:
                current_screen = None
                eff_col.append(pd.NA)

        df_matrix["effective_screen_type"] = eff_col
        log.debug(f"effective_screen_type unique values (pre-ffill): "
                  f"{df_matrix['effective_screen_type'].dropna().unique().tolist()}")
        log.debug("Forward-filling merged cells …")

        df_matrix[type_column] = df_matrix[type_column].ffill()
        df_matrix[role_id_column] = df_matrix[role_id_column].ffill()

        log.debug("Stripping / uppercasing string columns …")
        for col in df_matrix.select_dtypes(include="object"):
            if col == "effective_screen_type":
                continue
            df_matrix[col] = (df_matrix[col]
                              .str.strip()
                              .str.upper()
                              .str.replace(" ", ".")
                              .str.replace(r"\s+", "", regex=True))

        log.debug("Replacing literal 'NULL' / 'NAN' with pd.NA …")
        df_matrix = df_matrix.replace("NULL", pd.NA).replace("NAN", pd.NA)
        log.info("Normalisation complete.")
        _df_snapshot("df_matrix (normalised)", df_matrix)
    except Exception:
        log.error(
            f"[file: {base_name}] FAILED during normalisation.\n" + traceback.format_exc())
        raise

    def find_data_row_after(start_idx, expected_labels, max_lookahead=2):
        field_name_col = df_matrix.columns[type_column_index + 1]
        expected = {s.upper() for s in expected_labels}
        candidates = list(range(max_lookahead, 0, -1))
        for offset in candidates:
            idx = start_idx + offset
            if idx >= len(df_matrix):
                continue
            val = df_matrix.iloc[idx][field_name_col]
            if pd.notna(val) and str(val).strip().upper() in expected:
                if offset != max_lookahead:
                    log.warning(f"  [SHIFT] Label {expected} at offset +{offset} "
                                f"(canonical +{max_lookahead}). Marker row={start_idx}, "
                                f"data row={idx}")
                return idx
        log.warning(f"  [SHIFT] Label {expected} NOT found within +1..+{max_lookahead} "
                    f"after marker row {start_idx}. Block will be skipped.")
        return None

    # ----------------------------------------------------------
    # STEP 4: Extract roles
    # ----------------------------------------------------------
    _section("STEP 4 – Extract roles")
    field_name_col = df_matrix.columns[type_column_index + 1]
    log.debug(f"field_name_col resolved to: '{field_name_col}'")
    try:
        raw_role_vals = df_matrix[df_matrix["effective_screen_type"]
                                  == "MAINTENANCE"][role_id_column]
        _bad_role_tokens = {"", "ROLE_ID_NAME", "ROLE__ID_NAME",
                            "ROLE_DESCRIPTION", "GROUP_NAME"}

        def _is_real_role(v):
            if pd.isna(v):
                return False
            s = str(v).strip().upper()
            return bool(s) and s not in _bad_role_tokens

        roles = sorted({r for r in raw_role_vals if _is_real_role(r)})
        log.info(f"Roles found ({len(roles)}): {roles}")
        if not roles:
            log.warning("No roles found under type 'MAINTENANCE'. "
                        "Subsequent queries will have empty IN-clauses and may fail.")
    except Exception:
        log.error(
            f"[file: {base_name}] FAILED to extract roles.\n" + traceback.format_exc())
        raise

    OV_BLOCK_MARKERS = {
        "BLK_RESTRICTIVENETWORK": {"BLK_RESTRICTIVENETWORK", "NETWORK.RESTRICTION"},
        "BLK_ROLESOURCE":         {"BLK_ROLESOURCE", "SOURCE.RESTRICTION"},
    }

    def _markers_for(block_type):
        return OV_BLOCK_MARKERS.get(block_type, {block_type})

    present_optional_blocks = set()
    for blk in ["BLK_RESTRICTIVENETWORK", "BLK_ROLESOURCE"]:
        if df_matrix[type_column].isin(_markers_for(blk)).any():
            present_optional_blocks.add(blk)
    skipped = {"BLK_RESTRICTIVENETWORK",
               "BLK_ROLESOURCE"} - present_optional_blocks
    if skipped:
        log.info(
            f"Optional blocks NOT present in matrix — will skip: {sorted(skipped)}")
    else:
        log.info("All optional blocks (Network, Source) are present in matrix.")

    # ----------------------------------------------------------
    # STEP 5: Build df_screens_matrix
    # ----------------------------------------------------------
    _section("STEP 5 – Build df_screens_matrix")
    df_screens_matrix = df_screens.copy()

    def _value_cols(df_like):
        return [c for c in df_like.columns[type_column_index + 2:]
                if c != "effective_screen_type"]

    screen_types = ["MAINTENANCE", "ONLINE", "REPORTS",
                    "BATCH", "WEB_BRANCH", "PROCESS_STAGE_RIGHTS"]

    for screen_type in screen_types:
        log.info(f"  Processing screen type: {screen_type}")
        try:
            # ---- PROCESS_STAGE_RIGHTS ----
            if screen_type == "PROCESS_STAGE_RIGHTS":
                report_screens = df_matrix[
                    (df_matrix["effective_screen_type"] == "PROCESS_STAGE_RIGHTS") &
                    (df_matrix[field_name_col].isin(
                        list(FN_ID_LABELS) + ["EDITABLE"]))
                ]
                log.debug(
                    f"    PROCESS_STAGE_RIGHTS rows found: {len(report_screens)}")
                value_cols = _value_cols(screens)
                report_screens = report_screens[
                    [role_id_column, screens.columns[type_column_index + 1]] +
                    list(value_cols)
                ].copy()
                for role in roles:
                    try:
                        temp_df = df_screens.copy()
                        role_reports = report_screens[report_screens[role_id_column] == role].T[2:].dropna(
                            how="all")
                        if role_reports.empty:
                            continue
                        if role_reports.shape[1] != 2:
                            role_reports['Editable'] = 0
                        temp_df[["Function_id", "Generate"]] = role_reports
                        temp_df = temp_df.dropna(subset=["Function_id"])
                        temp_df["Role_id"] = role
                        for col in temp_df.columns[3:]:
                            temp_df[col] = temp_df["New"]
                        df_screens_matrix = pd.concat(
                            [df_screens_matrix, temp_df], ignore_index=True)
                    except Exception:
                        log.error(f"[file: {base_name}] FAILED in PROCESS_STAGE_RIGHTS | Role: '{role}'\n"
                                  + traceback.format_exc())
                        raise
                # CHANGED from original: the hardcoded UNC debug dump
                # ("//10.104.46.54/.../nasser/rights.csv") was machine-
                # specific. We now write to the per-file debug folder so
                # the script works on any PC without depending on a share.
                try:
                    _rights_path = os.path.join(
                        debug_dir, f"{base_name}_extracted_rights.csv")
                    df_screens_matrix.to_csv(_rights_path, index=False)
                    log.info(
                        f"  [EXTRACT] Saved rights snapshot → {_rights_path}")
                except Exception:
                    log.warning("  [EXTRACT] Could not save rights snapshot.\n"
                                + traceback.format_exc())
                continue

            # ---- REPORTS ----
            if screen_type == "REPORTS":
                field_name_col = screens.columns[type_column_index + 1]
                report_screens = df_matrix[
                    (df_matrix["effective_screen_type"] == "REPORTS") &
                    (df_matrix[field_name_col].isin(
                        list(FN_ID_LABELS) + ["PRINT", "GENERATE"]))
                ]
                log.debug(f"    REPORTS rows found: {len(report_screens)}")
                value_cols = _value_cols(screens)
                report_screens = report_screens[
                    [role_id_column, field_name_col] + list(value_cols)
                ].copy()
                for role in roles:
                    role_subset = report_screens[report_screens[role_id_column] == role]
                    if role_subset.empty:
                        continue
                    fn_to_row = {}
                    for _, r in role_subset.iterrows():
                        fn = r[field_name_col]
                        if pd.notna(fn):
                            key = str(fn).strip().upper()
                            if key not in fn_to_row:
                                fn_to_row[key] = r
                    fn_id_row = _pick_fn_id_row(fn_to_row)
                    if fn_id_row is None:
                        log.debug(
                            f"      Role '{role}': REPORTS has no ROLE_FUNCTION/FUNCTION_ID row → skipped.")
                        continue
                    n_vc = len(value_cols)
                    function_ids = fn_id_row[list(value_cols)].tolist()
                    prints = (fn_to_row["PRINT"][list(value_cols)].tolist()
                              if "PRINT" in fn_to_row else ["0"] * n_vc)
                    generates = (fn_to_row["GENERATE"][list(value_cols)].tolist()
                                 if "GENERATE" in fn_to_row else ["0"] * n_vc)
                    role_reports = pd.DataFrame({
                        "Function_id": function_ids,
                        "Print":       prints,
                        "Generate":    generates,
                    }).dropna(subset=["Function_id"])
                    if role_reports.empty:
                        continue
                    temp_df = df_screens.copy()
                    temp_df[["Function_id", "Print", "Generate"]
                            ] = role_reports.values
                    temp_df["Role_id"] = role
                    temp_df[temp_df.columns[2:8]] = "0"
                    temp_df[temp_df.columns[9:17]] = "0"
                    df_screens_matrix = pd.concat(
                        [df_screens_matrix, temp_df], ignore_index=True)
                continue

            # ---- WEB_BRANCH ----
            if screen_type == "WEB_BRANCH":
                field_name_col = screens.columns[type_column_index + 1]
                report_screens = df_matrix[
                    (df_matrix["effective_screen_type"] == "WEB_BRANCH") &
                    (df_matrix[field_name_col].isin(
                        list(FN_ID_LABELS) + ["NEW"]))
                ]
                log.debug(f"    WEB_BRANCH rows found: {len(report_screens)}")
                value_cols = _value_cols(screens)
                report_screens = report_screens[
                    [role_id_column, field_name_col] + list(value_cols)
                ].copy()
                for role in roles:
                    role_subset = report_screens[report_screens[role_id_column] == role]
                    if role_subset.empty:
                        continue
                    fn_to_row = {}
                    for _, r in role_subset.iterrows():
                        fn = r[field_name_col]
                        if pd.notna(fn):
                            key = str(fn).strip().upper()
                            if key not in fn_to_row:
                                fn_to_row[key] = r
                    fn_id_row = _pick_fn_id_row(fn_to_row)
                    if fn_id_row is None:
                        log.debug(
                            f"      Role '{role}': WEB_BRANCH has no ROLE_FUNCTION/FUNCTION_ID row → skipped.")
                        continue
                    n_vc = len(value_cols)
                    function_ids = fn_id_row[list(value_cols)].tolist()
                    news = (fn_to_row["NEW"][list(value_cols)].tolist()
                            if "NEW" in fn_to_row else ["0"] * n_vc)
                    role_reports = pd.DataFrame({
                        "Function_id": function_ids,
                        "New":         news,
                    }).dropna(subset=["Function_id"])
                    if role_reports.empty:
                        continue
                    temp_df = df_screens.copy()
                    temp_df[["Function_id", "New"]] = role_reports.values
                    temp_df["Role_id"] = role
                    temp_df[temp_df.columns[3:17]] = "0"
                    df_screens_matrix = pd.concat(
                        [df_screens_matrix, temp_df], ignore_index=True)
                continue

            # ---- BATCH ----
            if screen_type == "BATCH":
                batch_screens = df_matrix[
                    (df_matrix["effective_screen_type"] == "BATCH") &
                    (df_matrix[field_name_col].isin(FN_ID_LABELS))
                ]
                log.debug(f"    BATCH rows found: {len(batch_screens)}")
                value_cols = _value_cols(screens)
                batch_screens = batch_screens[
                    [role_id_column, screens.columns[type_column_index + 1]] +
                    list(value_cols)
                ].copy()
                for role in roles:
                    temp_df = df_screens.copy()
                    temp_batch_screens = batch_screens[batch_screens[role_id_column] == role].T.dropna(how="all")[
                        2:]
                    if temp_batch_screens.empty:
                        continue
                    temp_df["Function_id"] = temp_batch_screens
                    temp_df = temp_df.dropna(subset=["Function_id"])
                    temp_df["Role_id"] = role
                    temp_df[temp_df.columns[2:]] = "0"
                    df_screens_matrix = pd.concat(
                        [df_screens_matrix, temp_df], ignore_index=True)
                continue

            # ---- MAINTENANCE / ONLINE ----
            maint_field_to_outcol = [
                ("ROLE_FUNCTION", "Function_id"),
                ("NEW", "New"), ("COPY", "Copy"), ("DELETE", "Delete"),
                ("CLOSE", "Close"), ("UNLOCK", "Unlock"), ("REOPEN", "Reopen"),
                ("PRINT", "Print"), ("AUTH", "Auth"), ("REVERSE", "Reverse"),
                ("ROLLOVER", "Rollover"), ("CONFIRM", "Confirm"),
                ("LIQUIDATE", "Liquidate"), ("HOLD", "Hold"),
                ("TEMPLATE", "Template"), ("VIEW", "View"),
            ]
            expected_fns = (
                {fn for fn, _ in maint_field_to_outcol} | set(FN_ID_LABELS))
            screens = df_matrix[
                (df_matrix["effective_screen_type"] == screen_type) &
                (df_matrix[field_name_col].isin(expected_fns))
            ]
            log.debug(f"    {screen_type} rows found: {len(screens)}")
            if screens.empty:
                log.warning(
                    f"    No rows for type '{screen_type}' – skipping.")
                continue

            value_cols = _value_cols(screens)
            n_vc = len(value_cols)

            for role in roles:
                role_rows = screens[screens[role_id_column] == role]
                if role_rows.empty:
                    log.debug(f"      Role '{role}': no rows → skipped.")
                    continue
                fn_to_row = {}
                for _, r in role_rows.iterrows():
                    fn = r[field_name_col]
                    if pd.notna(fn):
                        key = str(fn).strip().upper()
                        if key not in fn_to_row:
                            fn_to_row[key] = r
                fn_id_row = _pick_fn_id_row(fn_to_row)
                if fn_id_row is None:
                    log.debug(
                        f"      Role '{role}': no ROLE_FUNCTION/FUNCTION_ID row in "
                        f"{screen_type} → skipped.")
                    continue
                role_data = {}
                missing_fns = []
                for fn, outcol in maint_field_to_outcol:
                    if outcol == "Function_id":
                        role_data[outcol] = fn_id_row[value_cols].tolist()
                    elif fn in fn_to_row:
                        role_data[outcol] = fn_to_row[fn][value_cols].tolist()
                    else:
                        role_data[outcol] = ["0"] * n_vc
                        missing_fns.append(fn)
                role_df = pd.DataFrame(role_data)
                role_df.insert(0, "Role_id", role)
                role_df["Generate"] = "0"
                role_df = role_df.dropna(subset=["Function_id"])
                if role_df.empty:
                    continue
                df_screens_matrix = pd.concat(
                    [df_screens_matrix, role_df], ignore_index=True)
                if missing_fns:
                    log.debug(
                        f"      Role '{role}': {len(role_df)} rows added "
                        f"(defaulted missing: {missing_fns}).")
                else:
                    log.debug(
                        f"      Role '{role}': {len(role_df)} rows added.")

        except Exception:
            log.error(f"[file: {base_name}] FAILED while processing screen type '{screen_type}'.\n"
                      f"  Matrix: '{base_name}'\n" + traceback.format_exc())
            raise

    value_flag_cols = [c for c in df_screens_matrix.columns
                       if c not in ("Role_id", "Function_id")]
    df_screens_matrix[value_flag_cols] = (
        df_screens_matrix[value_flag_cols].fillna("0").replace("", "0"))
    log.info(f"  [SCREENS] Filled empty permission cells with '0' "
             f"across {len(value_flag_cols)} column(s): {value_flag_cols}")

    _df_snapshot("df_screens_matrix (final)", df_screens_matrix)
    try:
        _path = os.path.join(debug_dir, f"{base_name}_extracted_screens.csv")
        df_screens_matrix.to_csv(_path, index=False)
        log.info(f"  [EXTRACT] Saved screens matrix → {_path}")
    except Exception:
        log.warning(
            "  [EXTRACT] Could not save screens CSV.\n" + traceback.format_exc())

    # ----------------------------------------------------------
    # HELPER: run_query
    # ----------------------------------------------------------
    def run_query(config, query_str, label="query"):
        # Rewrite FCPROD.<table> -> FCPROD.<table>@<db_link> if a db_link
        # is configured. Doing this at the top of run_query keeps the
        # original SQL strings readable in the function bodies while
        # ensuring the debug log and the actual DB call both see the
        # rewritten form.
        query_str = _apply_db_link(query_str)
        log.info(
            f"  [DB] Executing {label} against {config['dsn']} as {config['user']} …")
        if DB_LINK:
            log.debug(f"  [DB] SQL (with db_link={DB_LINK}):\n{query_str}")
        else:
            log.debug(f"  [DB] SQL:\n{query_str}")
        try:
            conn = oracledb.connect(
                user=config["user"], password=config["password"], dsn=config["dsn"])
            df = pd.read_sql(query_str, con=conn)
            conn.close()
            log.info(
                f"  [DB] {label} returned {len(df)} rows, {len(df.columns)} columns.")
            return df
        except Exception:
            log.error(
                f"[file: {base_name}] [DB] FAILED {label}.\n" + traceback.format_exc())
            raise

    # ----------------------------------------------------------
    # HELPER: compare_dataframes_by_keys
    # ----------------------------------------------------------
    def compare_dataframes_by_keys(df_matrix_in, df_system_in, key_cols, type_value):
        log.info(f"  [COMPARE] type='{type_value}', key_cols={key_cols}, "
                 f"matrix={df_matrix_in.shape}, system={df_system_in.shape}")
        try:
            df_matrixk = df_matrix_in.set_index(key_cols).sort_index()
            df_systemk = df_system_in.set_index(key_cols).sort_index()
            for side, dfc in [("matrix", df_matrixk), ("system", df_systemk)]:
                if dfc.index.duplicated().any():
                    log.warning(
                        f"  [COMPARE] df_{side} has duplicated index entries → collapsing to first per key.")
                    dfc = dfc.groupby(level=list(
                        range(dfc.index.nlevels))).first()
                    if side == "matrix":
                        df_matrixk = dfc
                    else:
                        df_systemk = dfc

            keys1 = set(df_matrixk.index)
            keys2 = set(df_systemk.index)
            log.debug(f"  [COMPARE] unique matrix keys={len(keys1)}, system keys={len(keys2)}, "
                      f"common={len(keys1 & keys2)}, "
                      f"only-in-matrix={len(keys1 - keys2)}, only-in-system={len(keys2 - keys1)}")

            missing_in_df_matrix_keys = list(keys2 - keys1)
            missing_in_df_system_keys = list(keys1 - keys2)
            missing_in_df_matrix = (
                df_systemk.loc[missing_in_df_matrix_keys].reset_index()
                if missing_in_df_matrix_keys else pd.DataFrame(columns=df_system_in.columns)
            )
            missing_in_df_matrix.insert(0, "type", type_value)
            missing_in_df_matrix["difference_type"] = "In system not in access matrix"

            missing_in_df_system = (
                df_matrixk.loc[missing_in_df_system_keys].reset_index()
                if missing_in_df_system_keys else pd.DataFrame(columns=df_matrix_in.columns)
            )
            missing_in_df_system.insert(0, "type", type_value)
            missing_in_df_system["difference_type"] = "In access matrix not in system"

            common_keys = sorted(list(keys1 & keys2))
            if common_keys:
                df_matrix_common = df_matrixk.reindex(common_keys).copy()
                df_system_common = df_systemk.reindex(common_keys).copy()
                common_value_cols = df_matrix_common.columns.intersection(
                    df_system_common.columns)
                if len(common_value_cols) == 0:
                    log.warning(
                        "  [COMPARE] No common value columns to compare.")
                    differences = pd.DataFrame()
                else:
                    df_matrix_common = df_matrix_common.reindex(
                        columns=common_value_cols)
                    df_system_common = df_system_common.reindex(
                        columns=common_value_cols)
                    diff_mask = df_matrix_common != df_system_common
                    differing_keys = diff_mask.any(axis=1)
                    if differing_keys.any():
                        differences = df_matrix_common.loc[differing_keys].copy(
                        )
                        for col in common_value_cols:
                            differences[col] = differences[col].where(
                                diff_mask[col], pd.NA)
                        differences = differences.dropna(
                            how="all").reset_index()
                        differences.insert(0, "type", type_value)
                        differences["difference_type"] = "Value in matrix_diff"
                        log.info(
                            f"  [COMPARE] {len(differences)} differing rows found.")
                    else:
                        log.info("  [COMPARE] No value differences found.")
                        differences = pd.DataFrame()
            else:
                log.info(
                    "  [COMPARE] No common keys → nothing to value-compare.")
                differences = pd.DataFrame()

            combined = pd.concat(
                [missing_in_df_matrix, missing_in_df_system, differences],
                ignore_index=True, sort=False)
            log.info(f"  [COMPARE] Combined diff shape: {combined.shape}")
            return combined
        except Exception:
            log.error(f"[file: {base_name}] [COMPARE] FAILED for type='{type_value}'.\n"
                      + traceback.format_exc())
            raise

    # ----------------------------------------------------------
    # STEP 6: Query + diff screens
    # ----------------------------------------------------------
    _section("STEP 6 – Query system screens & diff")
    try:
        query_screens = (
            "select role_id,ROLE_FUNCTION,"
            "to_char(CONTROL_1),to_char(CONTROL_2),to_char(CONTROL_3),to_char(CONTROL_4),"
            "to_char(CONTROL_5),to_char(CONTROL_6),to_char(CONTROL_7),to_char(CONTROL_8),"
            "to_char(CONTROL_9),to_char(CONTROL_10),to_char(CONTROL_11),to_char(CONTROL_12),"
            "to_char(CONTROL_13),to_char(CONTROL_14),to_char(CONTROL_15),to_char(CONTROL_16) "
            "from fcprod.smtb_role_detail where ( role_id in ("
            + ",".join(["'" + i + "'" for i in list(roles)]) +
            ") and role_function in ("
            "SELECT FUNCTION_ID FROM (SELECT A.FUNCTION_ID FUNCTION_ID, (SELECT DESCRIPTION FROM FCPROD.SMTB_FUNCTION_DESCRIPTION"
            " WHERE FUNCTION_ID = A.FUNCTION_ID and LANG_CODE like 'ENG') BALOON_HELP "
            "FROM FCPROD.SMTBS_MENU A, FCPROD.SMTBS_MODULES C "
            "WHERE A.TYPE_STRING IN ('SM','M', 'D','BOTG','ELM','LNC','LNM','PMM','GQ','MSP','R', 'U', 'ELR','B','P', 'T','O', 'ELO', 'PMO') "
            "AND A.MODULE = C.MODULE_ID and a.RECORD_STAT = 'O' AND a.ONCE_AUTH = 'Y' ) where 1=1 "
            "))"
        )
        df_screens_sys = run_query(uat_config, query_screens, label="screens")
        df_screens_sys.columns = df_screens.columns
        _df_snapshot("df_screens_sys", df_screens_sys)

        try:
            df_screens_sys.to_csv(os.path.join(
                debug_dir, "DIAG_df_screens_sys_raw.csv"), index=False)
        except Exception:
            log.warning("  [DIAG] Could not save raw screens snapshot.")

        excluded_fns = _get_excluded_function_ids(uat_config)
        if excluded_fns:
            def _norm(s):
                return s.astype(str).str.strip().str.upper()
            before_m = len(df_screens_matrix)
            before_s = len(df_screens_sys)
            df_screens_matrix = df_screens_matrix[
                ~_norm(df_screens_matrix["Function_id"]).isin(excluded_fns)
            ].reset_index(drop=True)
            df_screens_sys = df_screens_sys[
                ~_norm(df_screens_sys["Function_id"]).isin(excluded_fns)
            ].reset_index(drop=True)
            log.info(f"  [EXCLUDE] Dropped {before_m - len(df_screens_matrix)} "
                     f"matrix row(s) and {before_s - len(df_screens_sys)} "
                     f"system row(s) by exclusion list.")
        else:
            log.info(
                "  [EXCLUDE] Exclusion list empty — no filtering applied.")

        df_diff_screens = compare_dataframes_by_keys(
            df_screens_matrix.iloc[:, : len(df_screens_matrix.columns) - 2],
            df_screens_sys.iloc[:, :len(df_screens_sys.columns) - 2],
            ["Role_id", "Function_id"], "Function ID"
        )
    except Exception:
        log.error(
            f"[file: {base_name}] FAILED in STEP 6.\n" + traceback.format_exc())
        raise

    # ----------------------------------------------------------
    # STEP 7: Extract flags from matrix
    # ----------------------------------------------------------
    _section("STEP 7 – Extract flags from matrix")
    flag_mappings = [
        ("BLK_ROLEBRANCHES", "Branch_restriction_flag"),
        ("BLK_ROLEACCCLASS", "Account_class_flag"),
        ("BLK_ROLEPROD", "Product_posting_flag"),
        ("BLK_ROLEACCESSPROD", "Product_access_flag"),
        ("BLK_ROLEGLREST", "gl_restriction_flag"),
        ("BLK_GRP_RESTRICTION", "group_restriction_flag"),
    ]
    # if "BLK_RESTRICTIVENETWORK" in present_optional_blocks:
    #     flag_mappings.append(("BLK_RESTRICTIVENETWORK", "Network_flag"))
    # if "BLK_ROLESOURCE" in present_optional_blocks:
    #     flag_mappings.append(("BLK_ROLESOURCE", "Source_flag"))

    cols_flags = ["Role_id", "Branch_restriction_flag", "Account_class_flag",
                  "Product_posting_flag", "Product_access_flag",
                  "gl_restriction_flag", "group_restriction_flag"]
    # if "BLK_RESTRICTIVENETWORK" in present_optional_blocks:
    #     cols_flags.append("Network_flag")
    # if "BLK_ROLESOURCE" in present_optional_blocks:
    #     cols_flags.append("Source_flag")
    df_flags = pd.DataFrame(columns=cols_flags)
    df_flags_matrix = pd.DataFrame(columns=cols_flags)

    try:
        for role in roles:
            role_flags = {"Role_id": role}
            log.debug(f"  Flags for role '{role}':")
            for block_type, flag_name in flag_mappings:
                block_indices = df_matrix[
                    (df_matrix[role_id_column] == role) &
                    (df_matrix[type_column].isin(_markers_for(block_type)))
                ].index
                if len(block_indices) == 0:
                    log.debug(
                        f"    {flag_name}: block '{block_type}' not found → N/A")
                    role_flags[flag_name] = "not in matrix"
                    continue
                block_idx = block_indices[0]
                start_idx = block_idx
                flag_found = False
                for i in range(start_idx, min(start_idx + 3, len(df_matrix))):
                    row = df_matrix.iloc[i]
                    for col_offset in [2, 3, 4, 5]:
                        col_idx = type_column_index + col_offset
                        if col_idx >= len(df_matrix.columns):
                            continue
                        val = row[df_matrix.columns[col_idx]]
                        if pd.isna(val):
                            continue
                        val_str = str(val).strip().upper()
                        if val_str.startswith("DISALLOW") or val_str == "D":
                            resolved = "D"
                        elif val_str.startswith("ALLOW") or val_str == "A":
                            resolved = "A"
                        else:
                            continue
                        log.debug(f"    {flag_name}: '{val_str}' → {resolved}")
                        role_flags[flag_name] = resolved
                        flag_found = True
                        break
                    if flag_found:
                        break
                if not flag_found:
                    log.debug(
                        f"    {flag_name}: no ALLOW/DISALLOW value found → N/A")
                    role_flags[flag_name] = "not in matrix"
            df_flags_matrix = pd.concat(
                [df_flags_matrix, pd.DataFrame([role_flags])], ignore_index=True)

        _df_snapshot("df_flags_matrix", df_flags_matrix)
        try:
            _path = os.path.join(debug_dir, f"{base_name}_extracted_flags.csv")
            df_flags_matrix.to_csv(_path, index=False)
            log.info(f"  [EXTRACT] Saved flags matrix → {_path}")
        except Exception:
            log.warning(
                "  [EXTRACT] Could not save flags CSV.\n" + traceback.format_exc())
    except Exception:
        log.error(
            f"[file: {base_name}] FAILED in STEP 7.\n" + traceback.format_exc())
        raise

    # ----------------------------------------------------------
    # STEP 8: Query + diff flags
    # ----------------------------------------------------------
    _section("STEP 8 – Query system flags & diff")
    try:
        select_cols = [
            "rm.role_id", "rm.branches_allowed", "rm.accclass_allowed",
            "products_allowed", "products_access_allowed", "gl_allowed",
            "rc.GROUP_CODE_ALLOWED",
        ]
        if "BLK_RESTRICTIVENETWORK" in present_optional_blocks:
            select_cols.append("rc.NETWORK_ALLOWED")
        if "BLK_ROLESOURCE" in present_optional_blocks:
            select_cols.append("rc.SOURCE_ALLOWED")

        query_flags = (
            "select " + ",".join(select_cols) + " "
            "from FCPROD.SMTB_ROLE_MASTER rm "
            "join FCPROD.SMTB_ROLE_MASTER_custom rc on rm.role_id = rc.role_id "
            "where rm.role_id in ("
            + ",".join(["'" + i + "'" for i in list(roles)]) + ")"
        )
        df_flags_sys = run_query(uat_config, query_flags, label="flags")
        df_flags_sys.columns = df_flags.columns
        df_diff_flags = compare_dataframes_by_keys(
            df_flags_matrix, df_flags_sys, ["Role_id"], "flag")
    except Exception:
        log.error(
            f"[file: {base_name}] FAILED in STEP 8.\n" + traceback.format_exc())
        raise

    # ----------------------------------------------------------
    # STEP 9: Other values
    # ----------------------------------------------------------
    _section("STEP 9 – Extract other values from matrix")
    types_ov = ["BLK_ROLEACCCLASS", "BLK_QUEUERIGHTS", "BLK_ROLEPROD",
                "BLK_ROLEACCESSPROD", "BLK_GRP_RESTRICTION"]
    if "BLK_RESTRICTIVENETWORK" in present_optional_blocks:
        types_ov.append("BLK_RESTRICTIVENETWORK")
    if "BLK_ROLESOURCE" in present_optional_blocks:
        types_ov.append("BLK_ROLESOURCE")

    df_other_values_matrix = df_other_values.copy()
    try:
        OV_LABELS = {
            "BLK_ROLEACCCLASS":       {"ACCOUNT_CLASS", "ACCOUNT.CLASS"},
            "BLK_QUEUERIGHTS":        {"QUEUE"},
            "BLK_ROLEPROD":           {"PRODUCT_CODE", "PRODUCT.CODE"},
            "BLK_ROLEACCESSPROD":     {"PRODUCT_CODE", "PRODUCT.CODE"},
            "BLK_GRP_RESTRICTION":    {"GROUP_CODE", "GROUP.CODE"},
            "BLK_RESTRICTIVENETWORK": {"NETWORK", "NETWORK_CODE", "NETWORK_ID",
                                       "NETWORK.CODE", "NETWORK.ID"},
            "BLK_ROLESOURCE":         {"SOURCE", "SOURCE_CODE", "SOURCE_ID",
                                       "SOURCE.CODE", "SOURCE.ID"},
        }
        OV_CANONICAL = {
            "BLK_ROLEACCCLASS": 2, "BLK_QUEUERIGHTS": 1,
            "BLK_ROLEPROD": 2, "BLK_ROLEACCESSPROD": 2,
            "BLK_GRP_RESTRICTION": 2,
            "BLK_RESTRICTIVENETWORK": 2, "BLK_ROLESOURCE": 2,
        }
        for ov_type in types_ov:
            cols_ov = list(df_matrix.columns[type_column_index + 2:])
            cols_ov.insert(0, role_id_column)
            canonical = OV_CANONICAL.get(ov_type, 2)
            marker_idxs = list(
                df_matrix.index[df_matrix[type_column].isin(_markers_for(ov_type))])
            data_idxs = []
            for m in marker_idxs:
                d = find_data_row_after(
                    m, OV_LABELS[ov_type], max_lookahead=canonical)
                if d is not None:
                    data_idxs.append(d)
            if not data_idxs:
                log.warning(
                    f"  OV type '{ov_type}': no data rows resolved — skipping.")
                continue
            df_temp_value = df_matrix.loc[data_idxs][cols_ov]
            if ov_type == "BLK_QUEUERIGHTS":
                df_temp_value = df_temp_value.replace(" ", "_")
            log.debug(
                f"  OV type '{ov_type}' resolved {len(data_idxs)} data row(s)")
            TYPE_LABEL = {
                "BLK_ROLEACCCLASS": "ROLEACCCLASS", "BLK_QUEUERIGHTS": "QUEUERIGHTS",
                "BLK_ROLEPROD": "ROLEPROD", "BLK_ROLEACCESSPROD": "ROLEACCESSPROD",
                "BLK_GRP_RESTRICTION": "GRP_RESTRICTION",
                "BLK_RESTRICTIVENETWORK": "RESTRICTIVENETWORK",
                "BLK_ROLESOURCE": "ROLESOURCE",
            }
            for role in roles:
                role_marker_rows = df_temp_value[df_temp_value[role_id_column] == role]
                if role_marker_rows.empty:
                    continue
                collected_values = []
                for _, marker_row in role_marker_rows.iterrows():
                    for v in marker_row[cols_ov[1:]]:
                        if pd.notna(v):
                            collected_values.append(v)
                if not collected_values:
                    continue
                sub = pd.DataFrame({
                    "Role_id": role,
                    "Type":    TYPE_LABEL[ov_type],
                    "Value":   collected_values,
                })
                df_other_values_matrix = pd.concat(
                    [df_other_values_matrix, sub], ignore_index=True)
        _df_snapshot("df_other_values_matrix", df_other_values_matrix)
        try:
            _path = os.path.join(
                debug_dir, f"{base_name}_extracted_other_values.csv")
            df_other_values_matrix.to_csv(_path, index=False)
            log.info(f"  [EXTRACT] Saved other values matrix → {_path}")
        except Exception:
            log.warning(
                "  [EXTRACT] Could not save other_values CSV.\n" + traceback.format_exc())
    except Exception:
        log.error(
            f"[file: {base_name}] FAILED in STEP 9.\n" + traceback.format_exc())
        raise

    # ----------------------------------------------------------
    # STEP 10: Query + diff other values
    # ----------------------------------------------------------
    _section("STEP 10 – Query system other values & diff")
    try:
        roles_sql = ",".join(["'" + i + "'" for i in list(roles)])
        query_ov = (
            f"select role_id,'ROLEACCCLASS',account_class from FCPROD.SMTB_ROLE_ACCCLASS where role_id in ({roles_sql})"
            f" union all select role_id,'ROLEACCESSPROD',product_code from FCPROD.SMTB_ROLE_ACCESS_PRODUCTS where role_id in ({roles_sql})"
            f" union all select role_id,'ROLEPROD',product_code from FCPROD.SMTB_ROLE_PRODUCTS where role_id in ({roles_sql})"
            f" union all select user_role_id as role_id,'QUEUERIGHTS',queue from FCPROD.SMTB_QUEUE_RIGHTS where user_role_flag = 'R' and user_role_id in ({roles_sql})"
            f" union all select role_id,'GRP_RESTRICTION',group_code from FCPROD.SMTB_ROLE_GROUP where role_id in ({roles_sql})"
        )
        df_other_values_sys = run_query(
            uat_config, query_ov, label="other_values")
        df_other_values_sys.columns = df_other_values.columns

        def _canonicalize_value(s):
            if pd.isna(s):
                return s
            s = str(s).strip().upper()
            import re
            s = re.sub(r"\s+", ".", s)
            s = re.sub(r"\.+", ".", s)
            s = re.sub(r"\.", " ", s)
            return s

        before_matrix_sample = df_other_values_matrix["Value"].head(3).tolist()
        before_system_sample = df_other_values_sys["Value"].head(3).tolist()
        df_other_values_matrix["Value"] = df_other_values_matrix["Value"].map(
            _canonicalize_value)
        df_other_values_sys["Value"] = df_other_values_sys["Value"].map(
            _canonicalize_value)
        log.debug(
            f"  [CANON] matrix Value sample before: {before_matrix_sample}")
        log.debug(
            f"  [CANON] matrix Value sample after:  {df_other_values_matrix['Value'].head(3).tolist()}")
        log.debug(
            f"  [CANON] system Value sample before: {before_system_sample}")
        log.debug(
            f"  [CANON] system Value sample after:  {df_other_values_sys['Value'].head(3).tolist()}")

        def _strip_spaces_for_queuerights(df, side_label):
            if df is None or df.empty or "Type" not in df.columns:
                return df
            mask = (df["Type"].astype(
                str).str.strip().str.upper() == "QUEUERIGHTS")
            n = int(mask.sum())
            if n == 0:
                log.debug(f"  [QUEUE_FIX:{side_label}] No QUEUERIGHTS rows.")
                return df
            before = df.loc[mask, "Value"].astype(str).tolist()
            df.loc[mask, "Value"] = (df.loc[mask, "Value"].astype(str)
                                       .str.replace(r"\s+", "", regex=True))
            after = df.loc[mask, "Value"].astype(str).tolist()
            changed = sum(1 for b, a in zip(before, after) if b != a)
            log.info(f"  [QUEUE_FIX:{side_label}] Stripped spaces in "
                     f"{changed}/{n} QUEUERIGHTS Value(s).")
            if changed:
                sample = [(b, a) for b, a in zip(before, after) if b != a][:5]
                log.debug(
                    f"  [QUEUE_FIX:{side_label}] Sample changes: {sample}")
            return df

        df_other_values_matrix = _strip_spaces_for_queuerights(
            df_other_values_matrix, "matrix")
        df_other_values_sys = _strip_spaces_for_queuerights(
            df_other_values_sys, "system")

        df_diff_values = compare_dataframes_by_keys(
            df_other_values_matrix, df_other_values_sys,
            ["Role_id", "Type", "Value"], "")
        df_diff_values = df_diff_values[df_diff_values.columns[1:]]

        try:
            TYPE_TO_VALIDATION_KEY = {
                "ROLEPROD": "PRODUCT_CODE", "ROLEACCESSPROD": "PRODUCT_CODE",
                "GRP_RESTRICTION": "GRP_RESTRICTION", "ROLEACCCLASS": "ROLEACCCLASS",
                "QUEUERIGHTS": "QUEUERIGHTS",
            }
            _validate_caches = {}

            def _validate_row(row):
                t = row.get("Type")
                if pd.isna(t):
                    return ""
                t_norm = str(t).strip().upper()
                key = TYPE_TO_VALIDATION_KEY.get(t_norm)
                if not key:
                    return ""
                v = row.get("Value")
                if pd.isna(v):
                    return ""
                v_norm = str(v).strip().upper()
                if not v_norm:
                    return ""
                if t_norm == "QUEUERIGHTS":
                    v_norm = v_norm.replace(" ", "")
                    if "queue_nospace" not in _validate_caches:
                        src = _get_validation_set(uat_config, "QUEUERIGHTS")
                        _validate_caches["queue_nospace"] = {
                            s.replace(" ", "") for s in src}
                    valid_set = _validate_caches["queue_nospace"]
                else:
                    valid_set = _get_validation_set(uat_config, key)
                return ("does fetch on system" if v_norm in valid_set
                        else "does not fetch on system")

            if df_diff_values.empty:
                df_diff_values["value validation"] = pd.Series(dtype=object)
                log.info(
                    "  [VALIDATE] Other Values diff is empty — added empty 'value validation' column.")
            else:
                df_diff_values["value validation"] = df_diff_values.apply(
                    _validate_row, axis=1)
                vc = df_diff_values["value validation"].value_counts(
                    dropna=False)
                log.info(
                    f"  [VALIDATE] 'value validation' column populated: {vc.to_dict()}")
        except Exception:
            log.error("  [VALIDATE] FAILED while annotating value validation — "
                      "writing column as empty.\n" + traceback.format_exc())
            if "value validation" not in df_diff_values.columns:
                df_diff_values["value validation"] = ""
    except Exception:
        log.error(
            f"[file: {base_name}] FAILED in STEP 10.\n" + traceback.format_exc())
        raise

    # ----------------------------------------------------------
    # STEP 11: GL nodes & leaves
    # ----------------------------------------------------------
    _section("STEP 11 – GL nodes & leaves")
    df_gl_nodes_matrix = df_gl_nodes.copy()
    df_gl_leaves_matrix = df_gl_leaves.copy()
    gl_blocks = [
        ("BLK_ROLEGLREST",   df_gl_nodes,  "NODE_GL"),
        ("BLK_ROLEGLEXCEPT", df_gl_leaves, "LEAF_GL"),
    ]
    try:
        for block_name, target_df, type_label in gl_blocks:
            marker_idxs = list(
                df_matrix.index[df_matrix[type_column] == block_name])
            log.debug(
                f"  GL block '{block_name}': {len(marker_idxs)} marker row(s) found.")
            if not marker_idxs:
                continue
            cols_gl = list(df_matrix.columns[type_column_index + 2:])
            cols_gl.insert(0, role_id_column)
            node_idxs, leaf_idxs = [], []
            for m in marker_idxs:
                n = find_data_row_after(
                    m, {"NODE_GL", "NODE.GL"}, max_lookahead=4)
                if n is None:
                    continue
                node_idxs.append(n)
                if type_label == "LEAF_GL":
                    l = find_data_row_after(
                        n, {"EXCEPTION_LEAF_GL", "EXCEPTION.LEAF.GL"}, max_lookahead=6)
                    leaf_idxs.append(l if l is not None else None)
            if not node_idxs:
                log.warning(
                    f"  GL block '{block_name}': no NODE_GL row resolved — skipping.")
                continue
            df_block_nodes = df_matrix.loc[node_idxs][cols_gl]
            if type_label == "LEAF_GL":
                valid_leaf_idxs = [l for l in leaf_idxs if l is not None]
                df_block_leaves = df_matrix.loc[valid_leaf_idxs][cols_gl] if valid_leaf_idxs else pd.DataFrame(
                    columns=cols_gl)
            for role in roles:
                temp_df = df_block_nodes[df_block_nodes[role_id_column] == role].T.dropna(how="all")[
                    1:]
                if type_label == "LEAF_GL":
                    role_leaves_raw = (df_block_leaves[df_block_leaves[role_id_column] == role]
                                       if not df_block_leaves.empty else pd.DataFrame())
                    if role_leaves_raw.empty:
                        log.debug(f"      Role '{role}': NODE_GL present but no EXCEPTION_LEAF_GL "
                                  f"row → skipped for GL_LEAVES.")
                        continue
                    leaves_frame = role_leaves_raw.T.dropna(how="all")[1:]
                    if isinstance(leaves_frame, pd.DataFrame) and leaves_frame.shape[1] == 0:
                        log.debug(f"      Role '{role}': EXCEPTION_LEAF_GL row exists but has no "
                                  f"values → skipped for GL_LEAVES.")
                        continue
                    temp_df["leaves"] = leaves_frame
                if temp_df.empty:
                    continue
                temp_df.insert(0, "Role_id", role)
                temp_df.columns = target_df.columns
                if type_label == "NODE_GL":
                    df_gl_nodes_matrix = pd.concat(
                        [df_gl_nodes_matrix, temp_df], ignore_index=True)
                else:
                    df_gl_leaves_matrix = pd.concat(
                        [df_gl_leaves_matrix, temp_df], ignore_index=True)
        _df_snapshot("df_gl_nodes_matrix", df_gl_nodes_matrix)
        _df_snapshot("df_gl_leaves_matrix", df_gl_leaves_matrix)
        try:
            _path = os.path.join(
                debug_dir, f"{base_name}_extracted_gl_nodes.csv")
            df_gl_nodes_matrix.to_csv(_path, index=False)
            log.info(f"  [EXTRACT] Saved GL nodes matrix → {_path}")
            _path = os.path.join(
                debug_dir, f"{base_name}_extracted_gl_leaves.csv")
            df_gl_leaves_matrix.to_csv(_path, index=False)
            log.info(f"  [EXTRACT] Saved GL leaves matrix → {_path}")
        except Exception:
            log.warning(
                "  [EXTRACT] Could not save GL CSVs.\n" + traceback.format_exc())
    except Exception:
        log.error(
            f"[file: {base_name}] FAILED in STEP 11.\n" + traceback.format_exc())
        raise

    # ----------------------------------------------------------
    # STEP 12: Query + diff GL
    # ----------------------------------------------------------
    _section("STEP 12 – Query system GL & diff")
    try:
        roles_sql = ",".join(["'" + i + "'" for i in list(roles)])
        query_gl_nodes = f"SELECT role_id, node_gl FROM fcprod.smtb_role_glrest WHERE role_id IN ({roles_sql})"
        df_gl_nodes_sys = run_query(
            uat_config, query_gl_nodes, label="gl_nodes")
        df_gl_nodes_sys.columns = df_gl_nodes.columns
        query_gl_leaves = f"SELECT ROLE_ID, NODE_GL, EXCEPTION_LEAF_GL FROM fcprod.smtb_role_glexcept WHERE role_id IN ({roles_sql})"
        df_gl_leaves_sys = run_query(
            uat_config, query_gl_leaves, label="gl_leaves")
        df_gl_leaves_sys.columns = df_gl_leaves.columns
        df_diff_gl_nodes = compare_dataframes_by_keys(
            df_gl_nodes_matrix, df_gl_nodes_sys, ["Role_id", "Node_GL"], "GL_NODES")
        df_diff_gl_leaves = compare_dataframes_by_keys(
            df_gl_leaves_matrix, df_gl_leaves_sys, ["Role_id", "Node_GL", "Leaf_GL"], "GL_LEAVES")
        _annotate_validation(df_diff_gl_nodes, "Node_GL", "NODE_GL",
                             uat_config, label="VALIDATE_GL_NODES")
        _annotate_validation(df_diff_gl_leaves, "Leaf_GL", "LEAF_GL",
                             uat_config, label="VALIDATE_GL_LEAVES")
    except Exception:
        log.error(
            f"[file: {base_name}] FAILED in STEP 12.\n" + traceback.format_exc())
        raise

    # ----------------------------------------------------------
    # STEP 13: Message rights
    # ----------------------------------------------------------
    _section("STEP 13 – Message rights")
    df_msg_rights_matrix = df_msg_rights.copy()
    try:
        msg_rights = df_matrix[
            (df_matrix[type_column] == "GRANT_RIGHTS") &
            (df_matrix[field_name_col].isin([
                "CANCEL", "CHANGE_NODE", "RELEASE", "CHANGE_MEDIA", "BRANCH_MOVE", "HOLD", "TEST_INPUT",
                "CHANGE_ADDRESS", "REINSTATE", "CHANGE_PRIORITY", "AUTH_CANCEL", "AUTH_CHANGE_NODE",
                "AUTH_BRANCH_MOVE", "AUTH_TEST_INPUT", "AUTH_REINSTATE", "INSTALL", "TEST_CHECK",
                "LINK_CONTRACT", "CHANGE_BRANCH_IN", "CHANGE_MESSAGE", "CHANGE_FORCE_RELEASE_FUND",
                "SUPPRESS", "DELETE", "PRINT", "FT_UPLOAD", "MOVE_TO_QUEUE", "CHANGE_ADDRESS_IN",
                "AUTH_CHANGE_MESSAGE", "AUTH_RIGHTS", "CHANGE_FORCE_COVER_MATCH"
            ])) &
            ~(df_matrix[df_matrix.columns[type_column_index + 3]].isna())
        ]
        log.debug(f"  GRANT_RIGHTS rows found: {len(msg_rights)}")

        df_system_msg = msg_rights[[
            role_id_column,
            msg_rights.columns[type_column_index + 1],
            msg_rights.columns[type_column_index + 2],
            msg_rights.columns[type_column_index + 3],
        ]].copy()
        _df_snapshot("df_system_msg (message rights slice)", df_system_msg)

        for role in roles:
            temp_df = df_system_msg[df_system_msg[role_id_column] == role].T.dropna(how="all")[
                2:]
            temp_df = temp_df.fillna("N")
            if temp_df.empty or temp_df.shape[1] == 1:
                log.debug(
                    f"  Role '{role}': no msg_rights rows → inserting default N row.")
                temp_df = pd.DataFrame(
                    [{col: "N" for col in df_msg_rights.columns}])
                temp_df["Role_id"] = role
            else:
                temp_df.insert(0, "Role", role)
            temp_df.columns = df_msg_rights.columns
            df_msg_rights_matrix = pd.concat(
                [df_msg_rights_matrix, temp_df], ignore_index=True)

        _df_snapshot("df_msg_rights_matrix", df_msg_rights_matrix)
        try:
            _path = os.path.join(
                debug_dir, f"{base_name}_extracted_msg_rights.csv")
            df_msg_rights_matrix.to_csv(_path, index=False)
            log.info(f"  [EXTRACT] Saved message rights matrix → {_path}")
        except Exception:
            log.warning(
                "  [EXTRACT] Could not save msg_rights CSV.\n" + traceback.format_exc())
    except Exception:
        log.error(
            f"[file: {base_name}] FAILED in STEP 13.\n" + traceback.format_exc())
        raise

    # ----------------------------------------------------------
    # STEP 14: Query + diff message rights
    # ----------------------------------------------------------
    _section("STEP 14 – Query system msg rights & diff")
    try:
        roles_sql = ",".join(["'" + i + "'" for i in list(roles)])
        query_msg = (
            "select user_role_id,CANCEL,CHANGE_NODE,RELEASE,CHANGE_MEDIA,BRANCH_MOVE,HOLD,TEST_INPUT,"
            "CHANGE_ADDR,REINSTATE,CHANGE_PRIOR,CANCEL_AUTH,CHANGE_NODE_AUTH,RELEASE_AUTH,CHANGE_MEDIA_AUTH,"
            "BRANCH_MOVE_AUTH,HOLD_AUTH,TEST_INPUT_AUTH,CHANGE_ADDR_AUTH,REINSTATE_AUTH,CHANGE_PRIOR_AUTH,"
            "GENERATE,TEST_CHECK,LINK_CONTRACT,CHANGE_BRANCH_IN,CHANGE_MSG,CHANGE_FORCE_RELEASE_FUND,SUPPRESS,"
            "DELETE_MSG,PRINT,FT_UPLOAD,MOVE_QUEUE,CHANGE_ADDRESS_IN,CHANGE_MSG_AUTH,AUTH_RIGHTS,"
            f"CHANGE_FORCE_COVER_MATCH from fcprod.smtb_msgs_rights where user_role_id in ({roles_sql})"
        )
        df_msg_rights_sys = run_query(
            uat_config, query_msg, label="msg_rights")
        df_msg_rights_sys.columns = df_msg_rights.columns
        df_diff_msg_rights = compare_dataframes_by_keys(
            df_msg_rights_matrix, df_msg_rights_sys, ["Role_id"], "msg_rights")
    except Exception:
        log.error(
            f"[file: {base_name}] FAILED in STEP 14.\n" + traceback.format_exc())
        raise

    # ----------------------------------------------------------
    # STEP 15: Write Excel output + STEP 16: per-role summary
    # ----------------------------------------------------------
    _section("STEP 15 – Write Excel output")
    try:
        dfs = {
            "Screens":        df_diff_screens,
            "Flags":          df_diff_flags,
            "GL_nodes":       df_diff_gl_nodes,
            "GL_leaves":      df_diff_gl_leaves,
            "Message rights": df_diff_msg_rights,
            "Other Values":   df_diff_values,
        }

        if shared_dirs and "all_differences" in shared_dirs:
            differences_dir = shared_dirs["all_differences"]
        else:
            differences_dir = os.path.join(debug_dir, "differences")
        os.makedirs(differences_dir, exist_ok=True)
        output_excel = os.path.join(
            differences_dir, f"{base_name}_comparison.xlsx")
        log.info(f"Writing output to: {output_excel}")

        with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
            start_row = 0
            for name, df in dfs.items():
                log.debug(
                    f"  Writing section '{name}' at row {start_row}, shape={df.shape}")
                ws_name = "Sheet1"
                if ws_name not in writer.sheets:
                    writer.sheets[ws_name] = writer.book.add_worksheet(ws_name)
                ws = writer.sheets[ws_name]
                ws.write(start_row, 0, name)
                df.to_excel(writer, sheet_name=ws_name,
                            startrow=start_row + 1, index=False)
                start_row += len(df) + 3

        log.info(f"Output saved: {output_excel}")

        # STEP 16: per-role summary — best-effort; never lose the Excel
        # already written above.
        try:
            diffs_by_section = {
                "Screens":   df_diff_screens,
                "Flags":     df_diff_flags,
                "GL Nodes":  df_diff_gl_nodes,
                "GL Leaves": df_diff_gl_leaves,
                "Msg Rgts":  df_diff_msg_rights,
                "OtherVals": df_diff_values,
            }
            stats = _compute_role_stats(diffs_by_section, roles)
            summary_text = _build_file_summary_text(
                base_name, stats, output_excel)
            print(summary_text, flush=True)
            _save_file_summary(summary_text, base_name, debug_dir)
        except Exception:
            log.error("  [SUMMARY] FAILED to build per-role summary — "
                      "Excel output is unaffected.\n" + traceback.format_exc())
            stats = {}

        return output_excel, debug_dir, stats

    except Exception:
        log.error(
            f"[file: {base_name}] FAILED in STEP 15.\n" + traceback.format_exc())
        raise


# --- Stage 3: directory walker ----------------------------------------------

def _process_directory_csv(input_dir, output_dir):
    """Process every CSV in input_dir, segregating outputs under output_dir."""
    if not os.path.exists(input_dir):
        log.error(f"Input directory not found: {input_dir}")
        return []
    csv_files = sorted(glob.glob(os.path.join(input_dir, "*.csv")))
    # Only production matrices are ever compared with the database.
    csv_files = _drop_non_prod(csv_files, "compare", log.warning)
    if not csv_files:
        log.warning(f"No CSV files in directory: {input_dir}")
        return []
    log.info(f"Found {len(csv_files)} CSV file(s) in: {input_dir}")

    shared_dirs = {
        "successful":      os.path.join(output_dir, "successful"),
        "failed":          os.path.join(output_dir, "failed"),
        "all_differences": os.path.join(output_dir, "all_differences"),
    }
    for d in shared_dirs.values():
        os.makedirs(d, exist_ok=True)
    log.info("Output folders:")
    log.info(f"  successful      → {shared_dirs['successful']}")
    log.info(f"  failed          → {shared_dirs['failed']}")
    log.info(f"  all_differences → {shared_dirs['all_differences']}")

    def close_file_log_handlers():
        """Close FileHandlers on the stage-3 logger so Windows releases the
        debug.log file before we try to move its parent folder."""
        for handler in list(log.handlers):
            if isinstance(handler, logging.FileHandler):
                handler.flush()
                handler.close()
                log.removeHandler(handler)

    def move_debug_folder(staged_debug_dir, destination_parent):
        """Relocate the per-file debug folder into successful/ or failed/.
        Robust against Windows MAX_PATH (260 chars):
          1) os.rename — atomic directory-level move, sidesteps MAX_PATH on children
          2) per-file copy fallback — surface a shutil.Error to caller if any
             individual file fails
        """
        close_file_log_handlers()
        dest_debug = os.path.join(
            destination_parent, os.path.basename(staged_debug_dir))
        if os.path.exists(dest_debug):
            try:
                shutil.rmtree(dest_debug)
            except Exception:
                log.warning(f"  [MOVE] Could not clear pre-existing "
                            f"destination {dest_debug}; continuing.")
        try:
            os.rename(staged_debug_dir, dest_debug)
            return dest_debug
        except OSError:
            pass
        os.makedirs(dest_debug, exist_ok=True)
        errors = []
        for entry in os.listdir(staged_debug_dir):
            s = os.path.join(staged_debug_dir, entry)
            d = os.path.join(dest_debug, entry)
            try:
                if os.path.isdir(s):
                    shutil.copytree(s, d, dirs_exist_ok=True)
                else:
                    shutil.copy2(s, d)
            except Exception as exc:
                errors.append((s, d, str(exc)))
        if not errors:
            try:
                shutil.rmtree(staged_debug_dir)
            except Exception:
                log.warning(f"  [MOVE] Copied OK but couldn't remove source "
                            f"{staged_debug_dir}; leaving it in place.")
            return dest_debug
        raise shutil.Error(errors)

    results = []
    for i, csv_file in enumerate(csv_files, 1):
        log.info(f"\nFile {i}/{len(csv_files)}: {os.path.basename(csv_file)}")
        staged_debug_dir = os.path.join(
            input_dir, f"{os.path.splitext(os.path.basename(csv_file))[0]}_debug")

        # --- Stage A: real processing -------------------------------------
        try:
            output_file, staged_debug_dir, stats = _process_one_csv(
                csv_file, shared_dirs=shared_dirs)
        except Exception as exc:
            # Processing genuinely failed — comparison.xlsx was not written.
            failure_tb = traceback.format_exc()
            try:
                shutil.copy2(csv_file, os.path.join(
                    shared_dirs["failed"], os.path.basename(csv_file)))
            except Exception:
                log.warning("  [COPY] Could not copy CSV into failed/ "
                            "(non-fatal).\n" + traceback.format_exc())
            if os.path.exists(staged_debug_dir):
                try:
                    move_debug_folder(staged_debug_dir, shared_dirs["failed"])
                except Exception as move_exc:
                    log.warning(f"  [MOVE] Could not move debug folder into "
                                f"failed/ (non-fatal): {move_exc}")
            log.info(f"  ✘ {os.path.basename(csv_file)} → failed/")
            _record_failure(
                stage="compare",
                name=os.path.basename(csv_file),
                reason_code=type(exc).__name__,
                detail=str(exc),
                input_path=csv_file,
                tb=failure_tb,
            )
            results.append({
                "input": csv_file, "output": None,
                "status": f"Failed: {exc}",
                "stats": None, "warning": None,
            })
            continue

        # --- Stage B: post-processing housekeeping ------------------------
        # comparison.xlsx is already saved at this point. Copying the input
        # CSV and moving the debug folder are housekeeping; if they fail
        # (typically Windows MAX_PATH on the debug move) we log a warning
        # but DO NOT downgrade the file's status to "Failed".
        warning = None
        try:
            shutil.copy2(csv_file, os.path.join(
                shared_dirs["successful"], os.path.basename(csv_file)))
        except Exception as exc:
            msg = f"could not copy CSV into successful/: {exc}"
            log.warning(f"  [COPY] {msg}")
            warning = msg
        try:
            move_debug_folder(staged_debug_dir, shared_dirs["successful"])
        except Exception as exc:
            short_exc = str(exc)
            if len(short_exc) > 140:
                short_exc = short_exc[:137] + "…"
            msg = (f"debug folder left at staging "
                   f"(move into successful/ failed: {short_exc})")
            log.warning(f"  [MOVE] {msg}")
            warning = msg

        status = "Success" if warning is None else "Success (with warning)"
        icon = "✔" if warning is None else "⚠"
        log.info(f"  {icon} {os.path.basename(csv_file)} → successful/"
                 + ("" if warning is None else " (see warning)"))
        results.append({
            "input": csv_file, "output": output_file,
            "status": status, "stats": stats, "warning": warning,
        })

    try:
        print(_format_stage3_run_summary(results), flush=True)
    except Exception:
        log.error("Failed to render stage 3 run summary.\n"
                  + traceback.format_exc())

    return results


def stage3_compare():
    """Stage 3 entry point — compare every CSV in CSV_FOLDER with Oracle,
    write comparison xlsx files into COMPARISON_FOLDER. Returns a stats dict."""
    global log
    log = _stage3_setup_logger()
    _ensure_stage3_imports()

    _banner("STAGE 3 — COMPARE  (CSV/*.csv -> Comparison/all_differences/*.xlsx)")

    # Validate config before doing real work
    if not UAT_CONFIG.get("user") or not UAT_CONFIG.get("dsn"):
        print("ERROR: stage 3 needs database credentials in config.yaml under `database:` "
              "(user, password, dsn). Aborting.")
        return {"stage": "compare", "ok": False, "error": "missing db config"}
    if not ORACLE_LIB_DIR:
        print("ERROR: stage 3 needs `paths.oracle_client_lib_dir` set in config.yaml. "
              "Aborting.")
        return {"stage": "compare", "ok": False, "error": "missing oracle lib dir"}
    if not CSV_FOLDER.exists():
        print(f"ERROR: CSV folder does not exist: {CSV_FOLDER}")
        print("       Did stage 2 run successfully?")
        return {"stage": "compare", "ok": False, "error": "csv folder missing"}

    COMPARISON_FOLDER.mkdir(parents=True, exist_ok=True)
    log.info(f"Input folder    : {CSV_FOLDER}")
    log.info(f"Output folder   : {COMPARISON_FOLDER}")
    log.info(f"Oracle lib dir  : {ORACLE_LIB_DIR}")
    log.info(
        f"Database DSN    : {UAT_CONFIG['dsn']} (user {UAT_CONFIG['user']})")
    if DB_LINK:
        log.info(f"Database link   : {DB_LINK}  "
                 f"(queries rewritten as FCPROD.<table>@{DB_LINK})")
    else:
        log.info("Database link   : (none — queries run against the local schema)")

    results = _process_directory_csv(str(CSV_FOLDER), str(COMPARISON_FOLDER))

    successful = sum(1 for r in results if r["status"].startswith("Success"))
    failed = sum(1 for r in results if r["status"].startswith("Failed"))
    return {
        "stage": "compare",
        "ok": True,
        "processed": len(results),
        "successful": successful,
        "failed": failed,
    }


# ============================================================================
# STAGE 4: SUMMARY_ALL_DIFF  —  all_differences/*.xlsx -> _summary.xlsx
# ============================================================================
# Reads every *_comparison.xlsx produced by stage 3 and aggregates them into
# a single multi-sheet summary workbook:
#   Overview          One row per file: total diffs per section + grand total.
#   By Role           One row per (file, role): counts per section + total.
#   By Section Detail One row per (file, role, section), split by
#                     difference_type (matrix-only / system-only / value diff).
#   <Section>         Per-section drill-down sheets containing the full diff
#                     rows from every file, with a File column prepended.

_SUMMARY_SECTIONS = [
    "Screens", "Flags", "GL_nodes", "GL_leaves", "Message rights", "Other Values"
]
_SUMMARY_DIFF_TYPES = [
    "In access matrix not in system",
    "In system not in access matrix",
    "Value in matrix_diff",
]


def _parse_comparison_xlsx(path):
    """Parse one *_comparison.xlsx (the layout stage 3 writes) into
    {section_name: DataFrame}. Robust to empty sections."""
    raw = pd.read_excel(path, sheet_name="Sheet1", header=None)
    starts = []
    for i, v in enumerate(raw.iloc[:, 0]):
        if pd.notna(v) and str(v).strip() in _SUMMARY_SECTIONS:
            starts.append((i, str(v).strip()))
    starts.append((len(raw), None))  # sentinel

    sections = {}
    for k in range(len(starts) - 1):
        lbl_row, name = starts[k]
        end_row, _ = starts[k + 1]
        header_row = lbl_row + 1
        if header_row >= len(raw):
            sections[name] = pd.DataFrame()
            continue
        headers = raw.iloc[header_row].tolist()
        data = raw.iloc[header_row + 1:end_row].dropna(how="all")
        valid_cols = [c for c in headers if pd.notna(c)]
        if data.empty:
            sections[name] = pd.DataFrame(columns=valid_cols)
        else:
            df = data.copy()
            df.columns = headers
            sections[name] = df[valid_cols].reset_index(drop=True)
    for s in _SUMMARY_SECTIONS:
        sections.setdefault(s, pd.DataFrame())
    return sections


def _generate_summary_workbook(differences_dir, output_path=None):
    """Aggregate every *_comparison.xlsx in `differences_dir` into a single
    summary workbook. Returns the workbook path, or None if no inputs."""
    if output_path is None:
        output_path = os.path.join(differences_dir, "_summary.xlsx")

    xlsx_files = sorted(
        p for p in glob.glob(os.path.join(differences_dir, "*_comparison.xlsx"))
        if not os.path.basename(p).startswith("_summary")
    )
    xlsx_files = _drop_non_prod(xlsx_files, "summary_all_diff", log.warning)
    if not xlsx_files:
        log.warning(
            f"[SUMMARY] No *_comparison.xlsx files in {differences_dir}")
        return None

    log.info(f"[SUMMARY] Aggregating {len(xlsx_files)} comparison file(s) …")

    overview_rows = []
    by_role_rows = []
    by_section_detail = []
    per_section_full = {s: [] for s in _SUMMARY_SECTIONS}

    for xlsx_path in xlsx_files:
        file_name = os.path.basename(xlsx_path).replace("_comparison.xlsx", "")
        try:
            sections = _parse_comparison_xlsx(xlsx_path)
        except Exception as exc:
            log.error(f"[SUMMARY] Failed to parse {xlsx_path}\n"
                      + traceback.format_exc())
            _record_failure(
                stage="summary_all_diff",
                name=os.path.basename(xlsx_path),
                reason_code=type(exc).__name__,
                detail=f"_parse_comparison_xlsx failed: {exc}",
                input_path=xlsx_path,
                tb=traceback.format_exc(),
            )
            continue

        # --- Overview row (per file) ---
        ov = {"File": file_name}
        file_total = 0
        for s in _SUMMARY_SECTIONS:
            n = len(sections[s])
            ov[s] = n
            file_total += n
        ov["Total"] = file_total
        overview_rows.append(ov)

        # --- Collect every role mentioned anywhere in this file ---
        all_roles = set()
        for s in _SUMMARY_SECTIONS:
            df = sections[s]
            if not df.empty and "Role_id" in df.columns:
                all_roles.update(df["Role_id"].dropna().astype(str).unique())

        # --- By-role + by-section detail rows ---
        for role in sorted(all_roles):
            role_row = {"File": file_name, "Role_id": role}
            role_total = 0
            for s in _SUMMARY_SECTIONS:
                df = sections[s]
                if df.empty or "Role_id" not in df.columns:
                    role_row[s] = 0
                    continue
                sub = df[df["Role_id"].astype(str) == role]
                role_row[s] = len(sub)
                role_total += len(sub)
                if len(sub) > 0:
                    detail = {"File": file_name, "Role_id": role, "Section": s}
                    for dt in _SUMMARY_DIFF_TYPES:
                        if "difference_type" in sub.columns:
                            detail[dt] = int(
                                (sub["difference_type"] == dt).sum())
                        else:
                            detail[dt] = 0
                    detail["Total"] = len(sub)
                    by_section_detail.append(detail)
            role_row["Total"] = role_total
            by_role_rows.append(role_row)

        # --- Full per-section rows for drill-down sheets ---
        for s in _SUMMARY_SECTIONS:
            df = sections[s]
            if df.empty:
                continue
            df_copy = df.copy()
            df_copy.insert(0, "File", file_name)
            per_section_full[s].append(df_copy)

    df_overview = pd.DataFrame(
        overview_rows,
        columns=["File"] + _SUMMARY_SECTIONS + ["Total"],
    )
    if not df_overview.empty:
        totals = {"File": "TOTAL"}
        for c in _SUMMARY_SECTIONS + ["Total"]:
            totals[c] = int(df_overview[c].sum())
        df_overview = pd.concat(
            [df_overview, pd.DataFrame([totals])], ignore_index=True)

    df_by_role = pd.DataFrame(
        by_role_rows,
        columns=["File", "Role_id"] + _SUMMARY_SECTIONS + ["Total"],
    )
    df_section_detail = pd.DataFrame(
        by_section_detail,
        columns=["File", "Role_id", "Section"] +
        _SUMMARY_DIFF_TYPES + ["Total"],
    )

    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        wb = writer.book
        header_fmt = wb.add_format({
            "bold": True, "bg_color": "#305496", "font_color": "white",
            "align": "center", "valign": "vcenter", "border": 1,
        })
        total_fmt = wb.add_format({
            "bold": True, "bg_color": "#FFF2CC", "border": 1,
        })
        num_fmt = wb.add_format({"align": "center"})

        def _write(df, sheet, freeze=(1, 2)):
            df.to_excel(writer, sheet_name=sheet, index=False)
            ws = writer.sheets[sheet]
            for col, name in enumerate(df.columns):
                ws.write(0, col, name, header_fmt)
                sample_vals = df[name].head(50).tolist() if len(df) else []
                w = max(10, min(40, max(
                    [len(str(name))] + [len(str(v)) for v in sample_vals]) + 2))
                ws.set_column(col, col, w, num_fmt if col >= 1 else None)
            ws.freeze_panes(*freeze)
            ws.autofilter(0, 0, max(len(df), 1), max(len(df.columns) - 1, 0))

        _write(df_overview, "Overview", freeze=(1, 1))
        if len(df_overview) > 0:
            ws = writer.sheets["Overview"]
            last_row = len(df_overview)
            for col in range(len(df_overview.columns)):
                ws.write(last_row, col, df_overview.iloc[-1, col], total_fmt)

        _write(df_by_role, "By Role", freeze=(1, 2))
        _write(df_section_detail, "By Section Detail", freeze=(1, 3))

        for s in _SUMMARY_SECTIONS:
            if not per_section_full[s]:
                continue
            combined = pd.concat(per_section_full[s], ignore_index=True)
            sheet_name = s[:31]  # Excel sheet name limit
            _write(combined, sheet_name, freeze=(1, 1))

    log.info(f"[SUMMARY] Workbook written → {output_path}")
    return output_path


def stage4_summary_all_diff():
    """Stage 4 entry point — aggregate every *_comparison.xlsx in
    COMPARISON_FOLDER/all_differences/ into a single _summary.xlsx."""
    global log
    log = _stage3_setup_logger()
    _ensure_stage3_imports()

    _banner("STAGE 4 — SUMMARY_ALL_DIFF  (all_differences/*.xlsx -> _summary.xlsx)")

    diffs_dir = COMPARISON_FOLDER / "all_differences"
    if not diffs_dir.exists():
        print(f"ERROR: stage 4 expected {diffs_dir} to exist. "
              f"Did stage 3 run successfully?")
        return {"stage": "summary_all_diff", "ok": False,
                "error": "all_differences folder missing"}

    files = sorted(
        p for p in glob.glob(os.path.join(str(diffs_dir), "*_comparison.xlsx"))
        if not os.path.basename(p).startswith("_summary")
        and _classify_env(
            os.path.splitext(os.path.basename(p))[0])[0] != "excluded"
    )
    log.info(f"Input folder    : {diffs_dir}")
    log.info(f"Input files     : {len(files)} comparison workbook(s)")
    log.info(f"Output workbook : {diffs_dir / '_summary.xlsx'}")

    try:
        out = _generate_summary_workbook(str(diffs_dir))
    except Exception as exc:
        log.error("Stage 4 failed.\n" + traceback.format_exc())
        return {"stage": "summary_all_diff", "ok": False, "error": str(exc)}

    return {
        "stage": "summary_all_diff",
        "ok": True,
        "files_processed": len(files),
        "output": out,
    }


# ============================================================================
# STAGE 5: GENERAL_SQL  —  all_differences/*.xlsx -> generated_sql/*.sql
# ============================================================================
# For each *_comparison.xlsx produced by stage 3, validate every diff row
# against Oracle LOV tables, then generate Oracle SQL that brings the SYSTEM
# into alignment with the ACCESS MATRIX:
#
#   "In access matrix not in system"  →  INSERT (WHERE NOT EXISTS guarded)
#   "In system not in access matrix"  →  DELETE
#   "Value in matrix_diff"            →  UPDATE
#
# Outputs (under COMPARISON_FOLDER/all_differences/generated_sql/):
#   <basename>_remediation.sql   — one SQL file per workbook (valid rows only)
#   all_remediation.sql          — every workbook's SQL concatenated
#   validation_errors.xlsx       — every invalid row across every workbook
#   all_remediation_summary.xlsx — multi-sheet audit of all_remediation.sql
#
# Behaviour preserved verbatim from the standalone generate_sql.py:
#   - LOV-based validation against fcprod.* reference tables.
#   - role_id sanitisation (strip '.', truncate to 15 chars). The original
#     value is kept for ROLE_DESCRIPTION columns.
#   - group_code sanitisation (strip non-ASCII / invisible marks).
#   - GRP_RESTRICTION rows with value '0' are silently skipped.
#   - NETWORK_ALLOWED / SOURCE_ALLOWED only accept 'A' or 'D'; other values
#     are dropped from the generated SQL.
#   - Every INSERT into SMTB_role_detail is preceded by companion inserts
#     into SMTB_ROLE_MASTER and COZM_ROLE_DETAILS (WHERE NOT EXISTS guarded).
#   - Windows long-path support via the \\?\ prefix.
#
# Integration changes vs the standalone script:
#   - Reuses pipeline UAT_CONFIG / DB_LINK / ORACLE_LIB_DIR (no hard-coded
#     credentials inside this stage).
#   - LOV queries are routed through _apply_db_link() so FCPROD.<table>
#     becomes FCPROD.<table>@<db_link> when a link is configured.
#   - Reuses the stage-3 logger and the lazy _ensure_stage3_imports() so
#     pandas/oracledb stay optional for stages 1+2.

_SQL_DIFF_INSERT = "In access matrix not in system"
_SQL_DIFF_DELETE = "In system not in access matrix"
_SQL_DIFF_UPDATE = "Value in matrix_diff"

# Correct order for CONTROL_STRING concatenation: control_1 through control_16.
_SQL_CONTROL_STRING_ORDER = [
    "control_1", "control_2", "control_3", "control_4", "control_5",
    "control_6", "control_7", "control_8", "control_9", "control_10",
    "control_11", "control_12", "control_13", "control_14",
    "control_15", "control_16",
]

# Mapping from comparison-xlsx column names to DB column names for Screens.
_SQL_SCREENS_COL_MAP = {
    "New":       "control_1",
    "Copy":      "control_2",
    "Delete":    "control_3",
    "Close":     "control_4",
    "Unlock":    "control_5",
    "Reopen":    "control_6",
    "Print":     "control_7",
    "Auth":      "control_8",
    "Reverse":   "control_9",
    "Rollover":  "control_10",
    "Confirm":   "control_11",
    "Liquidate": "control_12",
    "Hold":      "control_13",
    "Template":  "control_14",
    "View":      "control_15",
    "Generate":  "control_16",
}

# Flags live across two tables — SMTB_ROLE_MASTER (the two columns below)
# and SMTB_ROLE_MASTER_CUSTOM (everything in FLAGS_CUSTOM_COLS).
_SQL_FLAGS_MASTER_COLS = {
    "Branch_restriction_flag": "branches_allowed",
    "Account_class_flag":      "accclass_allowed",
}
_SQL_FLAGS_CUSTOM_COLS = [
    ("Product_posting_flag",   "products_allowed"),
    ("Product_access_flag",    "products_access_allowed"),
    ("gl_restriction_flag",    "gl_allowed"),
    ("group_restriction_flag", "group_code_allowed"),
    ("Network_flag",           "NETWORK_ALLOWED"),
    ("Source_flag",            "SOURCE_ALLOWED"),
]

# NETWORK_ALLOWED / SOURCE_ALLOWED accept only 'A' or 'D' — values like
# 'not in matrix' get dropped from generated SQL.
_SQL_NETWORK_SOURCE_FLAG_COLS = {"Network_flag", "Source_flag"}
_SQL_NETWORK_SOURCE_VALID_VALUES = {"A", "D"}

# Routes OV-type literals to (table, role_col, value_col).
_SQL_OTHER_VALUES_TABLE_MAP = {
    "ROLEACCCLASS":       ("fcprod.smtb_role_accclass",          "role_id",      "account_class"),
    "ROLEACCESSPROD":     ("fcprod.smtb_role_access_products",   "role_id",      "product_code"),
    "ROLEPROD":           ("fcprod.smtb_role_products",          "role_id",      "product_code"),
    "QUEUERIGHTS":        ("fcprod.smtb_queue_rights",           "user_role_id", "queue"),
    "GRP_RESTRICTION":    ("fcprod.smtb_role_group",             "role_id",      "group_code"),
}
_SQL_OTHER_VALUES_EXTRA_WHERE = {
    "QUEUERIGHTS": "user_role_flag = 'R'",
}

# Comparison-xlsx column name -> DB column for SMTB_MSGS_RIGHTS.
_SQL_MSG_COL_MAP = {
    "Cancel":                    "cancel",
    "Change_Node":               "change_node",
    "Release":                   "release",
    "Change_Media":              "change_media",
    "Branch_Move":               "branch_move",
    "Hold":                      "hold",
    "Test_Input":                "test_input",
    "Change_Address":            "change_addr",
    "Reinstate":                 "reinstate",
    "Change_Priority":           "change_prior",
    "Auth_Cancel":               "cancel_auth",
    "Auth_Change_Node":          "change_node_auth",
    "Release_auth":              "release_auth",
    "Change_Media_Auth":         "change_media_auth",
    "Auth_Branch_Move":          "branch_move_auth",
    "Hold_Auth":                 "hold_auth",
    "Auth_Test_Input":           "test_input_auth",
    "Change_Address_Auth":       "change_addr_auth",
    "Auth_Reinstate":            "reinstate_auth",
    "Change_Priority_Auth":      "change_prior_auth",
    "Install":                   "generate",
    "Test_Check":                "test_check",
    "Link_Contract":             "link_contract",
    "Change_Branch_In":          "change_branch_in",
    "Change_Message":            "change_msg",
    "Change_Force_Release_Fund": "change_force_release_fund",
    "Suppress":                  "suppress",
    "Delete":                    "delete_msg",
    "Print":                     "print",
    "FT_Upload":                 "ft_upload",
    "Move_To_Queue":             "move_queue",
    "Change_Address_In":         "change_address_in",
    "Auth_Change_Message":       "change_msg_auth",
    "Auth_Rights":               "auth_rights",
    "Change_Force_Cover_Match":  "change_force_cover_match",
}
_SQL_MSG_RIGHTS_ROLE_FLAG = "R"

# role_id sanitisation
_SQL_ROLE_ID_MAX_LEN = 15

# Windows-path safety
_SQL_MAX_BASENAME_LEN = 120


# --- Windows long-path helpers ----------------------------------------------

def _sql_long_path(path):
    r"""Wrap a path with the Win32 \\?\ (or \\?\UNC\) prefix so >260-char
    paths work. Idempotent. No-op on non-Windows."""
    if not path:
        return path
    if os.name != "nt":
        return path
    p = os.path.abspath(path)
    if p.startswith("\\\\?\\"):
        return p
    if p.startswith("\\\\"):
        return "\\\\?\\UNC\\" + p[2:]
    return "\\\\?\\" + p


def _sql_shorten_basename(base, max_len=_SQL_MAX_BASENAME_LEN):
    """Truncate-and-hash a basename when it's too long for a stable filename."""
    if len(base) <= max_len:
        return base
    digest = hashlib.md5(base.encode("utf-8")).hexdigest()[:8]
    keep = max_len - len(digest) - 1
    return base[:keep] + "_" + digest


# --- LOV fetcher ------------------------------------------------------------
# All queries reference fcprod.<table> and pass through _apply_db_link()
# at runtime, so when DB_LINK is set every table becomes fcprod.<table>@<link>.

_SQL_LOVS_CACHE = None


def _sql_fetch_lovs():
    """Connect to Oracle once and fetch every LOV set the SQL generator needs.
    Cached in _SQL_LOVS_CACHE for the lifetime of the process. Each query is
    routed through _apply_db_link() so the configured db_link is honoured."""
    global _SQL_LOVS_CACHE
    if _SQL_LOVS_CACHE is not None:
        log.debug(f"[SQL-LOV] Reusing cached LOVs "
                  f"({sum(len(v) for v in _SQL_LOVS_CACHE.values())} total values).")
        return _SQL_LOVS_CACHE

    log.info("[SQL-LOV] Initialising Oracle thick-mode client …")
    try:
        oracledb.init_oracle_client(lib_dir=ORACLE_LIB_DIR)
    except Exception as exc:
        # Re-init in the same process raises a friendly error — proceed.
        if "DPI-1047" in str(exc) or "already" in str(exc).lower():
            log.debug(
                "[SQL-LOV] Oracle client already initialised — proceeding.")
        else:
            raise

    log.info(
        f"[SQL-LOV] Connecting to {UAT_CONFIG['dsn']} as {UAT_CONFIG['user']} …")
    conn = oracledb.connect(
        user=UAT_CONFIG["user"], password=UAT_CONFIG["password"], dsn=UAT_CONFIG["dsn"])
    log.info("[SQL-LOV] Connected. Fetching LOV tables …")

    queries = {
        "screens": (
            "SELECT FUNCTION_ID FROM (SELECT A.FUNCTION_ID FUNCTION_ID, "
            "(SELECT DESCRIPTION FROM FCPROD.SMTB_FUNCTION_DESCRIPTION "
            "WHERE FUNCTION_ID = A.FUNCTION_ID and LANG_CODE like 'ENG') BALOON_HELP "
            "FROM FCPROD.SMTBS_MENU A, FCPROD.SMTBS_MODULES C "
            "WHERE A.TYPE_STRING IN ('SM','M','D','BOTG','ELM','LNC','LNM','PMM','GQ','MSP','R','U','ELR','B','P','T','O','ELO','PMO') "
            "AND A.MODULE = C.MODULE_ID and a.RECORD_STAT = 'O' AND a.ONCE_AUTH = 'Y') where 1=1"
        ),
        "queues": (
            "SELECT DISTINCT QUEUE FROM FCPROD.SMTB_QUEUES "
            "WHERE RECORD_STAT = 'O' AND ONCE_AUTH = 'Y'"
        ),
        "group_code": (
            """    
         SELECT DISTINCT GROUP_CODE
        FROM   FCPROD.STTM_GROUP_CODE
        WHERE  GROUP_TYPE IN ('C','I')
        AND    RECORD_STAT = 'O'
        AND    AUTH_STAT   = 'A'   
        """
        ),
        "account_classes": "SELECT DISTINCT ACCOUNT_CLASS FROM FCPROD.STTMS_ACCOUNT_CLASS where record_stat = 'O'",
        "products": (
            """
            SELECT DISTINCT PRODUCT_CODE FROM fcprod.CSTM_PRODUCT WHERE RECORD_STAT = 'O' AND ONCE_AUTH = 'Y'
            """
        ),
        "gl_nodes": (
            "SELECT DISTINCT GL_CODE FROM ("
            "SELECT GL_CODE, GL_DESC FROM FCPROD.GLZM_GLMASTER WHERE LEAF = 'N' "
            "UNION "
            "SELECT GL_CODE, GL_DESC FROM FCPROD.GLZM_GLMASTER "
            "WHERE LEAF = 'Y' AND PARENT_GL = '0')"
        ),
        "gl_leaves": (
            "SELECT DISTINCT GL_CODE FROM ("
            "SELECT GL_CODE, GL_DESC, DECODE(LEAF,'Y','LEAF','NODE') LEAF "
            "FROM FCPROD.GLZM_GLMASTER)"
        ),
        # 'gl_all' is used only for the diagnostic "exists as LEAF, not NODE"
        # hint in validate_gl_*. Preserved exactly as in the original
        # generate_sql.py — including the GLTM table name, which differs
        # from the GLZM used by gl_nodes / gl_leaves above.
        "gl_all": "SELECT GL_CODE FROM fcprod.GLTM_GLMASTER",
        "network_code": (
            "SELECT NETWORK_CODE FROM fcprod.PMzM_NETWORK_CODE_MAINT "
            "WHERE RECORD_STAT = 'O' AND AUTH_STAT = 'A' "
            "AND PAYMENT_TYPE IN ('X','Y','R','B','S','Z','G')"
        ),
    }

    lovs = {}
    for name, sql in queries.items():
        try:
            q_rewritten = _apply_db_link(sql)
            if DB_LINK:
                log.debug(
                    f"[SQL-LOV:{name}] (with db_link={DB_LINK}):\n{q_rewritten}")
            df = pd.read_sql(q_rewritten, con=conn)
            lovs[name] = set(
                df.iloc[:, 0].dropna().astype(str).str.strip().str.upper()
            )
            log.info(
                f"[SQL-LOV] {name:<18} → {len(lovs[name])} value(s) loaded.")
        except Exception:
            log.error(f"[SQL-LOV] Failed to fetch '{name}'.\n"
                      + traceback.format_exc())
            raise

    conn.close()
    log.info("[SQL-LOV] LOV fetch complete. Connection closed.")
    _SQL_LOVS_CACHE = lovs
    return lovs


# --- Value / cell helpers ---------------------------------------------------

def _sql_q(val):
    """Wrap a value in Oracle single quotes, escaping embedded quotes."""
    if val is None:
        return "NULL"
    try:
        if pd.isna(val):
            return "NULL"
    except (TypeError, ValueError):
        pass
    return "'" + str(val).replace("'", "''") + "'"


def _sql_numeric_val(val):
    """Convert a control-column value to 0 or 1 (integer, not quoted).
    Accepts: 'Y'/'N', '1'/'0', True/False, etc."""
    if val is None:
        return 0
    try:
        if pd.isna(val):
            return 0
    except (TypeError, ValueError):
        pass
    s = str(val).strip().upper()
    if s in ("Y", "1", "TRUE", "YES", "X"):
        return 1
    return 0


def _sql_norm(val):
    """Uppercase + strip for LOV lookups. Returns '' for None/NA."""
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except (TypeError, ValueError):
        pass
    return str(val).strip().upper()


def _sql_is_na(val):
    """True for None or pandas NA."""
    try:
        return pd.isna(val)
    except (TypeError, ValueError):
        return val is None


def _sql_is_notna(val):
    return not _sql_is_na(val)


def _sql_safe_get(row, *keys):
    """First non-NA value across `keys` from `row`. None if all missing."""
    for key in keys:
        try:
            v = row.get(key) if hasattr(row, "get") else row[key]
            if _sql_is_na(v):
                continue
            return v
        except (KeyError, IndexError):
            continue
    return None


def _sql_row_to_dict(row):
    """Convert a pandas Series (from iterrows) to a plain dict, keeping the
    first occurrence on duplicate column names."""
    if isinstance(row, dict):
        return row
    seen = {}
    for col, val in zip(row.index, row.values):
        if col not in seen:
            seen[col] = val
    return seen


def _sql_diff_col(df):
    """Return the 'difference_type' column name (case-insensitive)."""
    for c in df.columns:
        if str(c).strip().lower() == "difference_type":
            return c
    return "difference_type"


def _sql_read_section(xl, sheet, section_name):
    """Pull one section out of a stage-3 comparison workbook (which writes
    every section stacked vertically on Sheet1)."""
    raw = xl.parse(sheet, header=None, dtype=str).fillna("")
    title_rows = raw.index[
        raw.iloc[:, 0].str.strip().str.lower() == section_name.lower()
    ].tolist()
    if not title_rows:
        return pd.DataFrame()
    title_row = title_rows[0]
    header_row = title_row + 1
    if header_row >= len(raw):
        return pd.DataFrame()
    headers = list(raw.iloc[header_row])
    data_rows = []
    for i in range(header_row + 1, len(raw)):
        row = list(raw.iloc[i])
        if all(str(c).strip() == "" for c in row):
            break
        data_rows.append(row)
    if not data_rows:
        return pd.DataFrame(columns=headers)
    df = pd.DataFrame(data_rows, columns=headers).replace("", pd.NA)
    return df


def _sql_make_err(matrix_file, section_nm, role, vtype, value, detail):
    """Build a standardised validation-error dict."""
    return {
        "matrix_file":  str(matrix_file),
        "section":      str(section_nm),
        "role_id":      str(role) if role is not None else "(blank)",
        "type":         str(vtype),
        "value":        str(value),
        "error_detail": str(detail),
    }


# --- Sanitisation -----------------------------------------------------------

def _sql_clean_role_id(role):
    """Strip '.' from a role_id and truncate to 15 chars.
    The original value is kept for ROLE_DESCRIPTION columns elsewhere."""
    if role is None:
        return role
    try:
        if pd.isna(role):
            return role
    except (TypeError, ValueError):
        pass
    s = str(role).replace(".", "").strip()
    if len(s) > _SQL_ROLE_ID_MAX_LEN:
        s = s[:_SQL_ROLE_ID_MAX_LEN]
    return s


def _sql_role_changed(original, cleaned):
    if original is None or cleaned is None:
        return False
    return str(original) != str(cleaned)


def _sql_clean_group_code(val):
    """Strip non-ASCII / invisible-combining marks from a group_code.
    Example: 'ٍSTAFF' (leading Arabic Kasratan U+064D) → 'STAFF'."""
    if val is None:
        return val
    try:
        if pd.isna(val):
            return val
    except (TypeError, ValueError):
        pass
    return str(val).encode("ascii", errors="ignore").decode("ascii").strip()


def _sql_is_valid_network_source_value(val):
    """True iff `val` is exactly 'A' or 'D' (case-insensitive)."""
    if _sql_is_na(val):
        return False
    return _sql_norm(val) in _SQL_NETWORK_SOURCE_VALID_VALUES


def _sql_companion_role_inserts(role):
    """Two companion inserts (WHERE NOT EXISTS guarded) that must accompany
    every INSERT into SMTB_role_detail: SMTB_ROLE_MASTER + COZM_ROLE_DETAILS.
    `role` is the ORIGINAL value; ROLE_ID columns get the sanitised version,
    ROLE_DESCRIPTION gets the original."""
    role_clean = _sql_clean_role_id(role)
    rid = _sql_q(role_clean)
    rdesc = _sql_q(role)
    return [
        (
            f"-- Companion: ensure role exists in FCPROD.SMTB_ROLE_MASTER\n"
            f"INSERT INTO FCPROD.SMTB_ROLE_MASTER\n"
            f"  (ROLE_ID, ROLE_DESCRIPTION, MAKER_ID, MAKER_DT_STAMP,\n"
            f"   CHECKER_ID, CHECKER_DT_STAMP, MOD_NO, RECORD_STAT, ONCE_AUTH, AUTH_STAT)\n"
            f"  SELECT {rid}, {rdesc}, '239582', SYSDATE, '22594', SYSDATE, 1, 'O', 'Y', 'A'\n"
            f"  FROM dual\n"
            f"  WHERE NOT EXISTS (\n"
            f"    SELECT 1 FROM FCPROD.SMTB_ROLE_MASTER WHERE ROLE_ID = {rid}\n"
            f"  );"
        ),
        (
            f"-- Companion: ensure role exists in FCPROD.COZM_ROLE_DETAILS\n"
            f"INSERT INTO FCPROD.COZM_ROLE_DETAILS (ROLE_ID, ROLE_DESCRIPTION, AUTH_STAT)\n"
            f"  SELECT {rid}, {rdesc}, 'A' FROM dual\n"
            f"  WHERE NOT EXISTS (\n"
            f"    SELECT 1 FROM FCPROD.COZM_ROLE_DETAILS WHERE ROLE_ID = {rid}\n"
            f"  );"
        ),
    ]


# --- Validators -------------------------------------------------------------
# Each validator returns (valid_rows_df, errors_list). Errors are dropped
# from SQL generation but recorded in validation_errors.xlsx.

def _sql_validate_screens(df, lovs, matrix_file, section_nm="Screens"):
    diff_col_name = _sql_diff_col(df)
    valid_rows, errors = [], []
    for _, raw_row in df.iterrows():
        row = _sql_row_to_dict(raw_row)
        diff = str(row.get(diff_col_name, "")).strip()
        role = _sql_safe_get(row, "Role_id", "role_id")
        func = _sql_safe_get(row, "Function_id", "function_id")
        if diff == _SQL_DIFF_DELETE:
            valid_rows.append(row)
            continue
        fn = _sql_norm(func)
        if not fn:
            errors.append(_sql_make_err(matrix_file, section_nm, role,
                                        "Screen", func, "Function_id is blank"))
            continue
        if fn not in lovs["screens"]:
            errors.append(_sql_make_err(matrix_file, section_nm, role, "Screen",
                                        func,
                                        f"Function_id '{func}' not found in fcprod.SMTBS_MENU"))
            continue
        valid_rows.append(row)
    return pd.DataFrame(valid_rows), errors


def _sql_validate_gl_nodes(df, lovs, matrix_file, section_nm="GL_nodes"):
    diff_col_name = _sql_diff_col(df)
    valid_rows, errors = [], []
    for _, raw_row in df.iterrows():
        row = _sql_row_to_dict(raw_row)
        diff = str(row.get(diff_col_name, "")).strip()
        role = _sql_safe_get(row, "Role_id", "role_id")
        node = _sql_safe_get(row, "Node_GL", "node_gl")
        if diff == _SQL_DIFF_DELETE:
            valid_rows.append(row)
            continue
        nd = _sql_norm(node)
        if not nd:
            errors.append(_sql_make_err(matrix_file, section_nm, role,
                                        "GL_NODE", node, "Node_GL is blank"))
            continue
        if nd not in lovs["gl_nodes"]:
            if nd in lovs["gl_all"]:
                detail = f"GL code '{node}' exists in GLTM_GLMASTER but is a LEAF, not a NODE"
            else:
                detail = f"GL code '{node}' not found in fcprod.GLTM_GLMASTER"
            errors.append(_sql_make_err(matrix_file, section_nm, role,
                                        "GL_NODE", node, detail))
            continue
        valid_rows.append(row)
    return pd.DataFrame(valid_rows), errors


def _sql_validate_gl_leaves(df, lovs, matrix_file, section_nm="GL_leaves"):
    diff_col_name = _sql_diff_col(df)
    valid_rows, errors = [], []
    for _, raw_row in df.iterrows():
        row = _sql_row_to_dict(raw_row)
        diff = str(row.get(diff_col_name, "")).strip()
        role = _sql_safe_get(row, "Role_id", "role_id")
        node = _sql_safe_get(row, "Node_GL", "node_gl")
        leaf = _sql_safe_get(row, "Leaf_GL", "exception_leaf_gl")
        if diff == _SQL_DIFF_DELETE:
            valid_rows.append(row)
            continue
        row_errors = []
        nd = _sql_norm(node)
        if not nd:
            row_errors.append("Node_GL is blank")
        elif nd not in lovs["gl_nodes"]:
            if nd in lovs["gl_all"]:
                row_errors.append(
                    f"Node_GL '{node}' exists but is a LEAF, not a NODE")
            else:
                row_errors.append(
                    f"Node_GL '{node}' not found in GLTM_GLMASTER")
        lf = _sql_norm(leaf)
        if not lf:
            row_errors.append("Leaf_GL is blank")
        elif lf not in lovs["gl_leaves"]:
            if lf in lovs["gl_all"]:
                row_errors.append(
                    f"Leaf_GL '{leaf}' exists but is a NODE, not a LEAF")
            else:
                row_errors.append(
                    f"Leaf_GL '{leaf}' not found in GLTM_GLMASTER")
        if row_errors:
            errors.append(_sql_make_err(matrix_file, section_nm, role, "GL_LEAF",
                                        f"node={node}, leaf={leaf}",
                                        "; ".join(row_errors)))
        else:
            valid_rows.append(row)
    return pd.DataFrame(valid_rows), errors


def _sql_validate_network_codes(df, lovs, matrix_file, section_nm="Network codes"):
    diff_col_name = _sql_diff_col(df)
    valid_rows, errors = [], []
    for _, raw_row in df.iterrows():
        row = _sql_row_to_dict(raw_row)
        diff = str(row.get(diff_col_name, "")).strip()
        role = _sql_safe_get(row, "Role_id", "role_id")
        net = _sql_safe_get(row, "Network_code", "network_code")
        if diff == _SQL_DIFF_DELETE:
            valid_rows.append(row)
            continue
        net_norm = _sql_norm(net)
        if not net_norm:
            errors.append(_sql_make_err(matrix_file, section_nm, role,
                                        "NETWORK_CODE", net, "Network_code is blank"))
            continue
        if net_norm not in lovs.get("network_code", set()):
            errors.append(_sql_make_err(matrix_file, section_nm, role, "NETWORK_CODE",
                                        net,
                                        f"Network_code '{net}' not found in "
                                        f"fcprod.PMzM_NETWORK_CODE_MAINT "
                                        f"(active, authorized, with allowed payment types)"))
            continue
        valid_rows.append(row)
    return pd.DataFrame(valid_rows), errors


def _sql_validate_flags(df, lovs, matrix_file, section_nm="Flags"):
    diff_col_name = _sql_diff_col(df)
    flag_cols = [
        "Branch_restriction_flag", "Account_class_flag",
        "Product_posting_flag",    "Product_access_flag",
        "gl_restriction_flag",     "group_restriction_flag",
    ]
    valid_rows, errors = [], []
    VALID_FLAG_VALUES = {"A", "D", "N/A", "NA", ""}
    for _, raw_row in df.iterrows():
        row = _sql_row_to_dict(raw_row)
        diff = str(row.get(diff_col_name, "")).strip()
        role = _sql_safe_get(row, "Role_id", "role_id")
        if diff == _SQL_DIFF_DELETE:
            valid_rows.append(row)
            continue
        bad = []
        for fc in flag_cols:
            val = row.get(fc)
            if _sql_is_na(val):
                continue
            if _sql_norm(val) not in VALID_FLAG_VALUES:
                bad.append(f"{fc}='{val}' (expected A / D / N/A)")
        # Network_flag / Source_flag silently dropped by gen_flags_sql, not a row-level error.
        if bad:
            errors.append(_sql_make_err(matrix_file, section_nm, role,
                                        "Flag", "",
                                        "Unexpected flag value(s): " + "; ".join(bad)))
        else:
            valid_rows.append(row)
    return pd.DataFrame(valid_rows), errors


def _sql_validate_other_values(df, lovs, matrix_file, section_nm="Other Values"):
    diff_col_name = _sql_diff_col(df)
    valid_rows, errors = [], []
    type_lov_map = {
        "ROLEACCCLASS":       ("account_classes", "fcprod.STTMS_ACCOUNT_CLASS"),
        "ROLEACCESSPROD":     ("products",        "fcprod.CSTM_PRODUCT"),
        "ROLEPROD":           ("products",        "fcprod.CSTM_PRODUCT"),
        "QUEUERIGHTS":        ("queues",          "fcprod.SMTB_QUEUES"),
        "GRP_RESTRICTION":    ("group_code",              "fcprod.SMTB_role_group"),
    }
    for _, raw_row in df.iterrows():
        row = _sql_row_to_dict(raw_row)
        diff = str(row.get(diff_col_name, "")).strip()
        role = _sql_safe_get(row, "Role_id", "role_id")
        vtype = _sql_norm(row.get("Type", ""))
        val = _sql_safe_get(row, "Value", "value")
        if diff == _SQL_DIFF_DELETE:
            if vtype == "GRP_RESTRICTION":
                cleaned = _sql_clean_group_code(val)
                if str(cleaned or "").strip() == "0":
                    log.debug(
                        f"    Skipping GRP_RESTRICTION DELETE with value '0' for role {role}")
                    continue
            valid_rows.append(row)
            continue
        if not vtype:
            errors.append(_sql_make_err(matrix_file, section_nm, role,
                                        "(blank type)", val, "Type column is blank"))
            continue
        mapping = type_lov_map.get(vtype)
        if mapping is None:
            errors.append(_sql_make_err(matrix_file, section_nm, role,
                                        vtype, val, f"Unknown OtherValues type '{vtype}'"))
            continue
        lov_key, lov_table = mapping
        if lov_key is None:
            cleaned = _sql_clean_group_code(val)
            if str(cleaned or "").strip() == "0":
                log.debug(
                    f"    Skipping GRP_RESTRICTION INSERT/UPDATE with value '0' for role {role}")
                continue
            if not cleaned:
                errors.append(_sql_make_err(matrix_file, section_nm, role,
                                            vtype, val,
                                            "group_code is empty after stripping non-ASCII characters"))
                continue
            valid_rows.append(row)
            continue
        vn = _sql_norm(val)
        if not vn:
            errors.append(_sql_make_err(matrix_file, section_nm, role,
                                        vtype, val, "Value is blank"))
            continue
        if vn not in lovs[lov_key]:
            errors.append(_sql_make_err(matrix_file, section_nm, role,
                                        vtype, val, f"'{val}' not found in {lov_table}"))
            continue
        valid_rows.append(row)
    return pd.DataFrame(valid_rows), errors


def _sql_validate_msg_rights(df, lovs, matrix_file, section_nm="Message rights"):
    """Message-rights values are Y/N — no LOV lookup required."""
    valid_rows = [_sql_row_to_dict(row) for _, row in df.iterrows()]
    return pd.DataFrame(valid_rows), []


# --- SQL generators ---------------------------------------------------------

def _sql_gen_screens_sql(df):
    """fcprod.SMTB_role_detail   PK: (role_id, role_function).
    Every INSERT also fires the two companion inserts (SMTB_ROLE_MASTER +
    COZM_ROLE_DETAILS)."""
    stmts = []
    if df.empty:
        return stmts
    diff_col_name = _sql_diff_col(df)
    table = "fcprod.SMTB_role_detail"
    for _, raw_row in df.iterrows():
        row = _sql_row_to_dict(raw_row)
        diff = str(row.get(diff_col_name, "")).strip()
        role = _sql_safe_get(row, "Role_id", "role_id")
        func = _sql_safe_get(row, "Function_id", "function_id")
        if _sql_is_na(role) or _sql_is_na(func):
            continue
        role_clean = _sql_clean_role_id(role)
        if _sql_role_changed(role, role_clean):
            log.debug(
                f"    [Screens] role_id sanitised: '{role}' → '{role_clean}'")

        if diff == _SQL_DIFF_DELETE:
            stmts.append(
                f"DELETE FROM {table}\n"
                f"  WHERE role_id = {_sql_q(role_clean)}\n"
                f"    AND role_function = {_sql_q(func)};"
            )
        elif diff == _SQL_DIFF_INSERT:
            control_nums = {}
            for excel_col, db_col in _SQL_SCREENS_COL_MAP.items():
                control_nums[db_col] = _sql_numeric_val(row.get(excel_col))
            control_string = "".join(
                str(control_nums.get(c, 0)) for c in _SQL_CONTROL_STRING_ORDER)
            insert_cols = ["role_id", "role_function", "control_string",
                           "auth_stat", "branch_code", "rad_function_id"]
            insert_vals = [_sql_q(role_clean), _sql_q(func), _sql_q(control_string),
                           _sql_q("A"), "NULL", "NULL"]
            for db_col in _SQL_CONTROL_STRING_ORDER:
                insert_cols.append(db_col)
                insert_vals.append(str(control_nums.get(db_col, 0)))
            stmts.extend(_sql_companion_role_inserts(role))
            stmts.append(
                f"INSERT INTO {table} ({', '.join(insert_cols)})\n"
                f"  SELECT {', '.join(insert_vals)} FROM dual\n"
                f"  WHERE NOT EXISTS (\n"
                f"    SELECT 1 FROM {table}\n"
                f"    WHERE role_id = {_sql_q(role_clean)} AND role_function = {_sql_q(func)}\n"
                f"  );"
            )
        elif diff == _SQL_DIFF_UPDATE:
            set_clauses = []
            changed_positions = {}
            for excel_col, db_col in _SQL_SCREENS_COL_MAP.items():
                val = row.get(excel_col)
                if _sql_is_notna(val):
                    num = _sql_numeric_val(val)
                    pos = _SQL_CONTROL_STRING_ORDER.index(db_col) + 1
                    set_clauses.append(f"    {db_col} = {num}")
                    changed_positions[db_col] = (pos, num)
            if set_clauses:
                cs_expr = "control_string"
                for db_col, (pos, num) in changed_positions.items():
                    cs_expr = (f"SUBSTR({cs_expr}, 1, {pos - 1}) || '{num}' || "
                               f"SUBSTR({cs_expr}, {pos + 1})")
                set_clauses.insert(0, f"    control_string = {cs_expr}")
                stmts.append(
                    f"UPDATE {table}\n"
                    f"  SET\n" + ",\n".join(set_clauses) + "\n"
                    f"  WHERE role_id = {_sql_q(role_clean)}\n"
                    f"    AND role_function = {_sql_q(func)};"
                )
    return stmts


def _sql_gen_flags_sql(df):
    """SMTB_ROLE_MASTER (2 cols) + SMTB_ROLE_MASTER_CUSTOM (6 cols).
    UPDATE first; if the role doesn't exist, INSERT full audit row."""
    stmts = []
    if df.empty:
        return stmts
    diff_col_name = _sql_diff_col(df)
    tbl_main = "fcprod.smtb_role_master"
    tbl_cust = "fcprod.smtb_role_master_custom"
    for _, raw_row in df.iterrows():
        row = _sql_row_to_dict(raw_row)
        diff = str(row.get(diff_col_name, "")).strip()
        role = _sql_safe_get(row, "Role_id", "role_id")
        if _sql_is_na(role):
            continue
        role_clean = _sql_clean_role_id(role)
        if _sql_role_changed(role, role_clean):
            log.debug(
                f"    [Flags] role_id sanitised: '{role}' → '{role_clean}'")

        if diff == _SQL_DIFF_DELETE:
            reset_main = ", ".join(
                f"{c} = NULL" for c in _SQL_FLAGS_MASTER_COLS.values())
            stmts.append(
                f"-- FLAGS DELETE (reset master cols to NULL) for role {_sql_q(role_clean)}\n"
                f"UPDATE {tbl_main}\n  SET {reset_main}\n  WHERE role_id = {_sql_q(role_clean)};"
            )
            reset_cust = ", ".join(
                f"{c} = NULL" for _, c in _SQL_FLAGS_CUSTOM_COLS)
            stmts.append(
                f"UPDATE {tbl_cust}\n  SET {reset_cust}\n  WHERE role_id = {_sql_q(role_clean)};"
            )
        elif diff in (_SQL_DIFF_INSERT, _SQL_DIFF_UPDATE):
            branches_val = row.get("Branch_restriction_flag")
            accclass_val = row.get("Account_class_flag")
            set_clauses_main = []
            if _sql_is_notna(branches_val):
                set_clauses_main.append(
                    f"    branches_allowed = {_sql_q(branches_val)}")
            if _sql_is_notna(accclass_val):
                set_clauses_main.append(
                    f"    accclass_allowed = {_sql_q(accclass_val)}")
            if set_clauses_main:
                stmts.append(
                    f"-- Update smtb_role_master flags for role {_sql_q(role_clean)}\n"
                    f"UPDATE {tbl_main}\n  SET\n"
                    + ",\n".join(set_clauses_main)
                    + f"\n  WHERE role_id = {_sql_q(role_clean)};"
                )
            br_insert = _sql_q(branches_val) if _sql_is_notna(
                branches_val) else "NULL"
            ac_insert = _sql_q(accclass_val) if _sql_is_notna(
                accclass_val) else "NULL"
            stmts.append(
                f"-- Insert smtb_role_master row for role {_sql_q(role_clean)} if missing\n"
                f"INSERT INTO {tbl_main}\n"
                f"  (ROLE_ID, ROLE_DESCRIPTION, BRANCHES_ALLOWED, ACCCLASS_ALLOWED,\n"
                f"   MAKER_ID, MAKER_DT_STAMP, CHECKER_ID, CHECKER_DT_STAMP,\n"
                f"   MOD_NO, RECORD_STAT, ONCE_AUTH, AUTH_STAT,\n"
                f"   BRANCH_VLT_ROLE, BRANCH_ROLE_CAT, BRANCH_ROLE_LEVEL,\n"
                f"   BRANCH_PWD_RESET_FREQ, BRANCH_ROLE, BRANCH_AUTH_ROLE,\n"
                f"   CENTRALISATION_ROLE)\n"
                f"  SELECT\n"
                f"    {_sql_q(role_clean)}, {_sql_q(role)}, {br_insert}, {ac_insert},\n"
                f"    '239582', SYSDATE, '222594', SYSDATE,\n"
                f"    1, 'O', 'Y', 'A',\n"
                f"    NULL, NULL, NULL,\n"
                f"    NULL, 'N', NULL,\n"
                f"    NULL\n"
                f"  FROM dual\n"
                f"  WHERE NOT EXISTS (\n"
                f"    SELECT 1 FROM {tbl_main} WHERE role_id = {_sql_q(role_clean)}\n"
                f"  );"
            )
            # Build the custom-table pair list, dropping NETWORK_ALLOWED /
            # SOURCE_ALLOWED rows whose value isn't 'A' or 'D'.
            cust_pairs = []
            for df_col, db_col in _SQL_FLAGS_CUSTOM_COLS:
                val = row.get(df_col)
                if not _sql_is_notna(val):
                    continue
                if df_col in _SQL_NETWORK_SOURCE_FLAG_COLS and not _sql_is_valid_network_source_value(val):
                    log.debug(f"    Skipping {df_col}={val!r} for role {role_clean} "
                              f"— only 'A' or 'D' allowed in {db_col}")
                    continue
                cust_pairs.append((db_col, _sql_q(val)))
            if cust_pairs:
                update_clause = ",\n        ".join(
                    f"{db_col} = {val}" for db_col, val in cust_pairs)
                insert_cols = "role_id, " + \
                    ", ".join(db_col for db_col, _ in cust_pairs)
                insert_vals = f"{_sql_q(role_clean)}, " + \
                    ", ".join(val for _, val in cust_pairs)
                stmts.append(
                    f"-- Update custom flags for role {_sql_q(role_clean)}; insert row if missing\n"
                    f"UPDATE {tbl_cust}\n"
                    f"  SET\n        {update_clause}\n"
                    f"  WHERE role_id = {_sql_q(role_clean)};\n"
                    f"\n"
                    f"INSERT INTO {tbl_cust} ({insert_cols})\n"
                    f"  SELECT {insert_vals} FROM dual\n"
                    f"  WHERE NOT EXISTS (\n"
                    f"    SELECT 1 FROM {tbl_cust} WHERE role_id = {_sql_q(role_clean)}\n"
                    f"  );"
                )
            else:
                log.debug(f"    All custom flag values were filtered for role {role_clean} "
                          f"— smtb_role_master_custom UPDATE/INSERT skipped.")
    return stmts


def _sql_gen_gl_nodes_sql(df):
    """fcprod.SMTB_role_glrest   PK: (role_id, node_gl)."""
    stmts = []
    if df.empty:
        return stmts
    diff_col_name = _sql_diff_col(df)
    table = "fcprod.SMTB_role_glrest"
    for _, raw_row in df.iterrows():
        row = _sql_row_to_dict(raw_row)
        diff = str(row.get(diff_col_name, "")).strip()
        role = _sql_safe_get(row, "Role_id", "role_id")
        node = _sql_safe_get(row, "Node_GL", "node_gl")
        if _sql_is_na(role) or _sql_is_na(node):
            continue
        role_clean = _sql_clean_role_id(role)
        if diff == _SQL_DIFF_DELETE:
            stmts.append(
                f"DELETE FROM {table}\n"
                f"  WHERE role_id = {_sql_q(role_clean)}\n"
                f"    AND node_gl = {_sql_q(node)};"
            )
        elif diff == _SQL_DIFF_INSERT:
            stmts.append(
                f"INSERT INTO {table} (role_id, node_gl)\n"
                f"  SELECT {_sql_q(role_clean)}, {_sql_q(node)} FROM dual\n"
                f"  WHERE NOT EXISTS (\n"
                f"    SELECT 1 FROM {table}\n"
                f"    WHERE role_id = {_sql_q(role_clean)} AND node_gl = {_sql_q(node)}\n"
                f"  );"
            )
    return stmts


def _sql_gen_gl_leaves_sql(df):
    """fcprod.SMTB_role_glexcept   PK: (role_id, node_gl, exception_leaf_gl)."""
    stmts = []
    if df.empty:
        return stmts
    diff_col_name = _sql_diff_col(df)
    table = "fcprod.SMTB_role_glexcept"
    for _, raw_row in df.iterrows():
        row = _sql_row_to_dict(raw_row)
        diff = str(row.get(diff_col_name, "")).strip()
        role = _sql_safe_get(row, "Role_id", "role_id")
        node = _sql_safe_get(row, "Node_GL", "node_gl")
        leaf = _sql_safe_get(row, "Leaf_GL", "exception_leaf_gl")
        if _sql_is_na(role) or _sql_is_na(node) or _sql_is_na(leaf):
            continue
        role_clean = _sql_clean_role_id(role)
        if diff == _SQL_DIFF_DELETE:
            stmts.append(
                f"DELETE FROM {table}\n"
                f"  WHERE role_id = {_sql_q(role_clean)}\n"
                f"    AND node_gl = {_sql_q(node)}\n"
                f"    AND exception_leaf_gl = {_sql_q(leaf)};"
            )
        elif diff == _SQL_DIFF_INSERT:
            stmts.append(
                f"INSERT INTO {table} (role_id, node_gl, exception_leaf_gl)\n"
                f"  SELECT {_sql_q(role_clean)}, {_sql_q(node)}, {_sql_q(leaf)} FROM dual\n"
                f"  WHERE NOT EXISTS (\n"
                f"    SELECT 1 FROM {table}\n"
                f"    WHERE role_id = {_sql_q(role_clean)}\n"
                f"      AND node_gl = {_sql_q(node)}\n"
                f"      AND exception_leaf_gl = {_sql_q(leaf)}\n"
                f"  );"
            )
    return stmts


def _sql_gen_network_codes_sql(df):
    """fcprod.SMZM_ROLE_NETWORK_DET_CUSTOM   PK: (role_id, network_code).
    Only INSERT and DELETE; no UPDATE."""
    stmts = []
    if df.empty:
        return stmts
    diff_col_name = _sql_diff_col(df)
    table = "fcprod.SMZM_ROLE_NETWORK_DET_CUSTOM"
    for _, raw_row in df.iterrows():
        row = _sql_row_to_dict(raw_row)
        diff = str(row.get(diff_col_name, "")).strip()
        role = _sql_safe_get(row, "Role_id", "role_id")
        net = _sql_safe_get(row, "Network_code", "network_code")
        if _sql_is_na(role) or _sql_is_na(net):
            continue
        role_clean = _sql_clean_role_id(role)
        if diff == _SQL_DIFF_DELETE:
            stmts.append(
                f"DELETE FROM {table}\n"
                f"  WHERE role_id = {_sql_q(role_clean)}\n"
                f"    AND network_code = {_sql_q(net)};"
            )
        elif diff == _SQL_DIFF_INSERT:
            stmts.append(
                f"INSERT INTO {table} (role_id, network_code)\n"
                f"  SELECT {_sql_q(role_clean)}, {_sql_q(net)} FROM dual\n"
                f"  WHERE NOT EXISTS (\n"
                f"    SELECT 1 FROM {table}\n"
                f"    WHERE role_id = {_sql_q(role_clean)} AND network_code = {_sql_q(net)}\n"
                f"  );"
            )
    return stmts


def _sql_gen_other_values_sql(df):
    """Routes each row to one of the tables in _SQL_OTHER_VALUES_TABLE_MAP.
    GRP_RESTRICTION rows with cleaned value '0' or '' are silently dropped."""
    stmts = []
    if df.empty:
        return stmts
    diff_col_name = _sql_diff_col(df)
    for _, raw_row in df.iterrows():
        row = _sql_row_to_dict(raw_row)
        diff = str(row.get(diff_col_name, "")).strip()
        role = _sql_safe_get(row, "Role_id", "role_id")
        vtype = _sql_norm(row.get("Type", ""))
        val = _sql_safe_get(row, "Value", "value")
        if _sql_is_na(role) or not vtype or _sql_is_na(val):
            continue
        role_clean = _sql_clean_role_id(role)
        if vtype == "GRP_RESTRICTION":
            original_val = val
            val = _sql_clean_group_code(val)
            if val != original_val:
                log.debug(f"    group_code sanitised: {original_val!r} → {val!r} "
                          f"for role {role_clean}")
            if not val or str(val).strip() == "0":
                log.debug(f"    Skipping GRP_RESTRICTION with cleaned value {val!r} "
                          f"for role {role_clean} (SQL generator)")
                continue
        mapping = _SQL_OTHER_VALUES_TABLE_MAP.get(vtype)
        if not mapping:
            continue
        table, role_col, value_col = mapping
        extra_where = _SQL_OTHER_VALUES_EXTRA_WHERE.get(vtype)
        if diff == _SQL_DIFF_DELETE:
            where = f"{role_col} = {_sql_q(role_clean)}\n    AND {value_col} = {_sql_q(val)}"
            if extra_where:
                where += f"\n    AND {extra_where}"
            stmts.append(f"DELETE FROM {table}\n  WHERE {where};")
        elif diff == _SQL_DIFF_INSERT:
            if vtype == "QUEUERIGHTS":
                stmts.append(
                    f"INSERT INTO {table} (user_role_id, queue, user_role_flag)\n"
                    f"  SELECT {_sql_q(role_clean)}, {_sql_q(val)}, 'R' FROM dual\n"
                    f"  WHERE NOT EXISTS (\n"
                    f"    SELECT 1 FROM {table}\n"
                    f"    WHERE user_role_id = {_sql_q(role_clean)} AND queue = {_sql_q(val)} AND user_role_flag = 'R'\n"
                    f"  );"
                )
            else:
                stmts.append(
                    f"INSERT INTO {table} ({role_col}, {value_col})\n"
                    f"  SELECT {_sql_q(role_clean)}, {_sql_q(val)} FROM dual\n"
                    f"  WHERE NOT EXISTS (\n"
                    f"    SELECT 1 FROM {table}\n"
                    f"    WHERE {role_col} = {_sql_q(role_clean)} AND {value_col} = {_sql_q(val)}\n"
                    f"  );"
                )
    return stmts


def _sql_gen_msg_rights_sql(df):
    """fcprod.SMTB_msgs_rights   PK: (user_role_id, user_role_flag='R')."""
    stmts = []
    if df.empty:
        return stmts
    diff_col_name = _sql_diff_col(df)
    table = "fcprod.SMTB_msgs_rights"
    flag_val = _sql_q(_SQL_MSG_RIGHTS_ROLE_FLAG)
    for _, raw_row in df.iterrows():
        row = _sql_row_to_dict(raw_row)
        diff = str(row.get(diff_col_name, "")).strip()
        role = _sql_safe_get(row, "Role_id", "role_id")
        if _sql_is_na(role):
            continue
        role_clean = _sql_clean_role_id(role)
        if diff == _SQL_DIFF_DELETE:
            stmts.append(
                f"DELETE FROM {table}\n"
                f"  WHERE user_role_id = {_sql_q(role_clean)}\n"
                f"    AND user_role_flag = {flag_val};"
            )
        elif diff == _SQL_DIFF_INSERT:
            cols = ["user_role_id", "user_role_flag"]
            vals = [_sql_q(role_clean), flag_val]
            for df_col, db_col in _SQL_MSG_COL_MAP.items():
                val = row.get(df_col)
                if _sql_is_notna(val):
                    cols.append(db_col)
                    vals.append(_sql_q(val))
            stmts.append(
                f"INSERT INTO {table} ({', '.join(cols)})\n"
                f"  SELECT {', '.join(vals)} FROM dual\n"
                f"  WHERE NOT EXISTS (\n"
                f"    SELECT 1 FROM {table}\n"
                f"    WHERE user_role_id = {_sql_q(role_clean)} AND user_role_flag = {flag_val}\n"
                f"  );"
            )
        elif diff == _SQL_DIFF_UPDATE:
            set_clauses = [
                f"    {db_col} = {_sql_q(row.get(df_col))}"
                for df_col, db_col in _SQL_MSG_COL_MAP.items()
                if _sql_is_notna(row.get(df_col))
            ]
            if set_clauses:
                stmts.append(
                    f"UPDATE {table}\n"
                    f"  SET\n" + ",\n".join(set_clauses) + "\n"
                    f"  WHERE user_role_id = {_sql_q(role_clean)}\n"
                    f"    AND user_role_flag = {flag_val};"
                )
    return stmts


# --- Section registry: drives which validator + generator runs per section --

_SQL_SECTION_REGISTRY = [
    ("Screens",        _sql_validate_screens,       _sql_gen_screens_sql),
    ("Flags",          _sql_validate_flags,         _sql_gen_flags_sql),
    ("GL_nodes",       _sql_validate_gl_nodes,      _sql_gen_gl_nodes_sql),
    ("GL_leaves",      _sql_validate_gl_leaves,     _sql_gen_gl_leaves_sql),
    ("Network codes",  _sql_validate_network_codes, _sql_gen_network_codes_sql),
    ("Message rights", _sql_validate_msg_rights,    _sql_gen_msg_rights_sql),
    ("Other Values",   _sql_validate_other_values,  _sql_gen_other_values_sql),
]


# --- Per-workbook processor -------------------------------------------------

def _sql_process_workbook(xlsx_path, output_dir, lovs):
    """Validate every section in one *_comparison.xlsx and write a
    per-file *_remediation.sql containing only the rows that passed LOV
    validation. Returns (sql_path, [validation_error_dicts])."""
    base_full = os.path.splitext(os.path.basename(xlsx_path))[0]
    base = _sql_shorten_basename(base_full)
    if base != base_full:
        log.info(f"  Basename shortened: '{base_full}' → '{base}'")
    log.info(f"  Processing: {base_full}")

    all_stmts = []
    all_errors = []
    try:
        xl = pd.ExcelFile(_sql_long_path(xlsx_path), engine="openpyxl")
    except Exception as exc:
        log.error(
            f"  Cannot open workbook: {xlsx_path}\n" + traceback.format_exc())
        _record_failure(
            stage="general_sql",
            name=os.path.basename(xlsx_path),
            reason_code=type(exc).__name__,
            detail=f"pd.ExcelFile open failed: {exc}",
            input_path=xlsx_path,
            tb=traceback.format_exc(),
        )
        return None, []
    sheet = xl.sheet_names[0]

    for section_name, validate_fn, sql_fn in _SQL_SECTION_REGISTRY:
        log.debug(f"    Section: {section_name}")
        try:
            df = _sql_read_section(xl, sheet, section_name)
            if df.empty:
                log.debug(f"      → empty / not found, skipped.")
                continue
            valid_df, errors = validate_fn(df, lovs, base_full, section_name)
            all_errors.extend(errors)
            if errors:
                log.warning(
                    f"      {section_name}: {len(errors)} validation error(s) "
                    f"— those rows are EXCLUDED from SQL."
                )
                for e in errors:
                    log.debug(
                        f"        ✘  matrix={e['matrix_file']} | "
                        f"section={e['section']} | role={e['role_id']} | "
                        f"type={e['type']} | value={e['value']} | "
                        f"reason: {e['error_detail']}"
                    )
            stmts = sql_fn(valid_df)
            if stmts:
                all_stmts.append(f"\n-- {'─'*56}")
                all_stmts.append(
                    f"-- Section: {section_name}  "
                    f"({len(stmts)} statement(s), {len(errors)} row(s) excluded)"
                )
                all_stmts.append(f"-- {'─'*56}\n")
                all_stmts.extend(stmts)
                log.info(f"      {section_name}: {len(stmts)} statement(s) generated, "
                         f"{len(errors)} row(s) excluded.")
            else:
                log.info(
                    f"      {section_name}: no valid differences to remediate.")
        except Exception as exc:
            log.error(f"  FAILED processing section '{section_name}' in {base_full}\n"
                      + traceback.format_exc())
            _record_failure(
                stage="general_sql",
                name=f"{base_full} :: {section_name}",
                reason_code=type(exc).__name__,
                detail=f"section processing failed: {exc}",
                input_path=xlsx_path,
                tb=traceback.format_exc(),
            )

    if not all_stmts:
        log.info(
            f"    No valid differences in {base_full} — no SQL file written.")
        return None, all_errors

    output_path = os.path.join(output_dir, f"{base}_remediation.sql")
    header = (
        f"-- {'═'*58}\n"
        f"-- Remediation SQL  –  generated by access_matrix_pipeline (stage 5)\n"
        f"-- Source   : {xlsx_path}\n"
        f"-- Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        f"--\n"
        f"-- PURPOSE : Bring the Oracle environment into alignment with\n"
        f"--           the access matrix.\n"
        f"-- WARNING : Review ALL statements before executing.\n"
        f"--           Run in a test / UAT environment first.\n"
        f"-- NOTE    : Rows that failed LOV validation are EXCLUDED.\n"
        f"--           See validation_errors.xlsx for the full list.\n"
        f"-- NOTE    : Each INSERT into SMTB_role_detail is preceded by\n"
        f"--           companion inserts into SMTB_ROLE_MASTER and\n"
        f"--           COZM_ROLE_DETAILS (both guarded by WHERE NOT EXISTS).\n"
        f"-- NOTE    : GRP_RESTRICTION rows with value '0' are skipped.\n"
        f"-- NOTE    : role_id values have '.' removed and are truncated to\n"
        f"--           15 chars before being written to ROLE_ID columns.\n"
        f"-- NOTE    : group_code values have non-ASCII characters stripped.\n"
        f"-- NOTE    : NETWORK_ALLOWED / SOURCE_ALLOWED accept only 'A'/'D';\n"
        f"--           other values are dropped from the generated SQL.\n"
        f"-- {'═'*58}\n"
    )
    try:
        with open(_sql_long_path(output_path), "w", encoding="utf-8") as f:
            f.write(header)
            f.write("\n".join(all_stmts))
            f.write("\n\nCOMMIT;\n")
    except OSError:
        log.error(
            f"  Cannot write SQL file: {output_path}\n"
            f"  Path length = {len(os.path.abspath(output_path))} chars.\n"
            + traceback.format_exc()
        )
        return None, all_errors
    log.info(f"    Written → {output_path}")
    return output_path, all_errors


def _sql_write_validation_report(all_errors, output_dir):
    """Write validation_errors.xlsx (red-banded, autofiltered, frozen header)."""
    if not all_errors:
        log.info(
            "[SQL] No validation errors — validation_errors.xlsx not written.")
        return None
    report_path = os.path.join(output_dir, "validation_errors.xlsx")
    cols = ["matrix_file", "section", "role_id",
            "type", "value", "error_detail"]
    df = pd.DataFrame(all_errors, columns=cols)
    try:
        with pd.ExcelWriter(_sql_long_path(report_path), engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name="Validation Errors", index=False)
            wb = writer.book
            ws = writer.sheets["Validation Errors"]
            hdr_fmt = wb.add_format({
                "bold": True, "bg_color": "#C00000", "font_color": "#FFFFFF",
                "border": 1, "align": "center", "valign": "vcenter",
            })
            row_fmt = wb.add_format(
                {"bg_color": "#FFE0E0", "border": 1, "valign": "vcenter"})
            err_fmt = wb.add_format({
                "bg_color": "#FFE0E0", "border": 1,
                "text_wrap": True, "valign": "vcenter",
            })
            col_widths = {"matrix_file": 32, "section": 18, "role_id": 22,
                          "type": 22, "value": 30, "error_detail": 58}
            for ci, col_name in enumerate(cols):
                ws.set_column(ci, ci, col_widths.get(col_name, 20))
                ws.write(0, ci, col_name, hdr_fmt)
            for ri, row_data in df.iterrows():
                ws.set_row(ri + 1, 30)
                for ci, val in enumerate(row_data):
                    fmt = err_fmt if ci == len(cols) - 1 else row_fmt
                    ws.write(ri + 1, ci, str(val), fmt)
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, len(df), len(cols) - 1)
    except OSError:
        log.error(f"Cannot write validation report: {report_path}\n"
                  + traceback.format_exc())
        return None
    log.info(f"[SQL] Validation report written → {report_path}  "
             f"({len(df)} error(s) across all workbooks)")
    return report_path


# ----------------------------------------------------------------------------
# REMEDIATION-SQL SUMMARY WORKBOOK
# ----------------------------------------------------------------------------
# Parses the freshly-written all_remediation.sql and produces a human-readable
# xlsx audit of what every statement will do. Useful for sign-off before any
# remediation SQL is executed against Oracle.
#
# Sheets produced:
#   Overview            One-line totals + key counts.
#   By Table            Per target-table breakdown (INSERT/UPDATE/DELETE).
#   By Role             Per role_id breakdown.
#   By Section          Counts grouped by the original section banner
#                       (Screens / Flags / GL_nodes / GL_leaves / Network codes
#                        / Message rights / Other Values).
#   By Source Workbook  Counts grouped by the originating *_comparison.xlsx.
#   All Statements      Every parsed statement with metadata + a 200-char
#                       excerpt of the SQL itself.

# Match the leading verb of a statement plus its target table (case-insensitive,
# multiline so '^' lines up with the start of a statement-line, not just the
# whole text).
_SQL_RE_VERB = re.compile(
    r"^\s*(INSERT\s+INTO|DELETE\s+FROM|UPDATE)\s+([\w.@]+)",
    re.IGNORECASE | re.MULTILINE,
)
# role_id / user_role_id capture — take the FIRST occurrence in the statement.
# That's normally the WHERE clause for UPDATE/DELETE and the SELECT values for
# INSERT. Subqueries in WHERE NOT EXISTS clauses also use role_id but they
# carry the SAME value, so picking the first match is correct either way.
_SQL_RE_ROLE_ID = re.compile(
    r"\b(?:role_id|user_role_id)\s*=\s*'([^']*)'",
    re.IGNORECASE,
)
# Comment-line patterns. The combined SQL preserves each per-workbook header
# (including '-- Source   : <path>') and each section banner.
_SQL_RE_SOURCE = re.compile(r"^--\s*Source\s*:\s*(.+?)\s*$", re.IGNORECASE)
_SQL_RE_SECTION = re.compile(r"^--\s*Section:\s*(.+?)\s*\(", re.IGNORECASE)
_SQL_RE_COMPANION = re.compile(r"^--\s*Companion:", re.IGNORECASE)
# Additional key columns to extract beyond role_id. Order matters only for
# display; the dict in _sql_parse_one_statement walks them all.
_SQL_SUMMARY_KEY_COLS = [
    "role_function", "node_gl", "exception_leaf_gl",
    "network_code", "account_class", "product_code",
    "queue", "group_code",
]


def _sql_parse_one_statement(stmt_text, source, section, is_companion):
    """Pull operation / target_table / role_id / key columns out of one
    SQL statement string. Returns None if no recognisable verb is present
    (purely-comment chunks, blank lines, etc.)."""
    m_verb = _SQL_RE_VERB.search(stmt_text)
    if not m_verb:
        return None
    op_raw = m_verb.group(1).upper()
    if op_raw.startswith("INSERT"):
        operation = "INSERT"
    elif op_raw.startswith("DELETE"):
        operation = "DELETE"
    else:
        operation = "UPDATE"
    target_table = m_verb.group(2)

    m_role = _SQL_RE_ROLE_ID.search(stmt_text)
    role_id = m_role.group(1) if m_role else ""

    key_values = []
    for col in _SQL_SUMMARY_KEY_COLS:
        m = re.search(rf"\b{col}\s*=\s*'([^']*)'", stmt_text, re.IGNORECASE)
        if m:
            key_values.append(f"{col}={m.group(1)}")

    # Compact one-line excerpt (collapse whitespace so the cell stays readable).
    excerpt = re.sub(r"\s+", " ", stmt_text).strip()
    if len(excerpt) > 240:
        excerpt = excerpt[:237] + "…"

    return {
        "operation":    operation,
        "target_table": target_table,
        "role_id":      role_id,
        "key_values":   "; ".join(key_values),
        "section":      section or "",
        "source_file":  os.path.basename(source) if source else "",
        "is_companion": bool(is_companion),
        "sql_excerpt":  excerpt,
    }


def _sql_parse_remediation_file(sql_path):
    """Walk all_remediation.sql line by line, buffering statements until each
    semicolon, and returning a list of parsed-statement dicts. Tracks the
    current section banner and source workbook from comment lines so each
    statement carries that context."""
    try:
        with open(_sql_long_path(str(sql_path)), encoding="utf-8") as fh:
            text = fh.read()
    except OSError as exc:
        log.error(f"[SQL-SUMMARY] Cannot read {sql_path}: {exc}")
        return []

    statements = []
    current_source = None
    current_section = None
    next_is_companion = False
    buf = []

    for raw in text.splitlines():
        stripped = raw.strip()

        if stripped.startswith("--"):
            # Comment line: update context but don't flush the buffer. The
            # companion flag only sticks until the NEXT real statement.
            m_src = _SQL_RE_SOURCE.match(stripped)
            if m_src:
                current_source = m_src.group(1).strip()
                continue
            m_sec = _SQL_RE_SECTION.match(stripped)
            if m_sec:
                current_section = m_sec.group(1).strip()
                continue
            if _SQL_RE_COMPANION.match(stripped):
                next_is_companion = True
            continue

        if not stripped:
            # Preserve blank lines inside buffered statements (rare, but safe)
            if buf:
                buf.append(raw)
            continue

        if stripped.upper().startswith("COMMIT"):
            # Final COMMIT — ignore.
            buf = []
            continue

        buf.append(raw)
        if stripped.endswith(";"):
            stmt_text = "\n".join(buf).strip()
            buf = []
            parsed = _sql_parse_one_statement(
                stmt_text,
                source=current_source,
                section=current_section,
                is_companion=next_is_companion,
            )
            next_is_companion = False
            if parsed:
                statements.append(parsed)
    return statements


def _sql_summarize_remediation(sql_path, output_xlsx_path):
    """Produce a multi-sheet xlsx audit of every statement in `sql_path`.
    Returns the output path on success, or None if the SQL file was missing
    or empty. Safe to call repeatedly; any error during parsing or writing
    is logged and returns None instead of bubbling up — stage 5's main
    deliverables (the SQL itself) must not be jeopardised by a problem in
    the summary."""
    if not os.path.exists(_sql_long_path(str(sql_path))):
        log.warning(f"[SQL-SUMMARY] Source file not found: {sql_path}")
        return None

    try:
        statements = _sql_parse_remediation_file(sql_path)
    except Exception:
        log.error(f"[SQL-SUMMARY] Failed to parse {sql_path}\n"
                  + traceback.format_exc())
        return None

    if not statements:
        log.warning(f"[SQL-SUMMARY] No statements parsed from {sql_path}")
        return None

    log.info(f"[SQL-SUMMARY] Parsed {len(statements)} statement(s) "
             f"from {os.path.basename(str(sql_path))}.")

    df = pd.DataFrame(statements)
    # Make sure every expected operation appears as a column even if zero,
    # so the pivots below don't blow up on partial data.
    for col in ["operation", "target_table", "role_id", "section",
                "source_file"]:
        if col not in df.columns:
            df[col] = ""

    # --- Overview ----------------------------------------------------------
    by_op_counts = df["operation"].value_counts().to_dict()
    n_companion = int(df["is_companion"].sum())
    n_main = len(df) - n_companion
    overview_rows = [
        ("Total SQL statements",                 len(df)),
        ("Main statements (excluding companions)", n_main),
        ("Companion role-master inserts",         n_companion),
        ("INSERT statements",                     int(
            by_op_counts.get("INSERT", 0))),
        ("UPDATE statements",                     int(
            by_op_counts.get("UPDATE", 0))),
        ("DELETE statements",                     int(
            by_op_counts.get("DELETE", 0))),
        ("Distinct target tables",
         int(df["target_table"].nunique())),
        ("Distinct role_id values",
         int(df["role_id"].replace("", pd.NA).dropna().nunique())),
        ("Source workbooks affected",
         int(df["source_file"].replace("", pd.NA).dropna().nunique())),
        ("Source SQL file",                       str(sql_path)),
    ]
    df_overview = pd.DataFrame(overview_rows, columns=["Metric", "Value"])

    def _pivot_by(col, header_name):
        """Group df by `col` and pivot operation counts into columns,
        always including INSERT/UPDATE/DELETE even when zero."""
        sub = df[df[col].astype(str) != ""].copy()
        if sub.empty:
            return pd.DataFrame(columns=[header_name, "INSERT", "UPDATE",
                                         "DELETE", "Total"])
        piv = (sub.groupby(col)["operation"].value_counts()
               .unstack(fill_value=0))
        for op in ("INSERT", "UPDATE", "DELETE"):
            if op not in piv.columns:
                piv[op] = 0
        piv = piv[["INSERT", "UPDATE", "DELETE"]]
        piv["Total"] = piv.sum(axis=1)
        piv = piv.reset_index().rename(columns={col: header_name})
        return piv.sort_values("Total", ascending=False).reset_index(drop=True)

    df_by_table = _pivot_by("target_table", "Target Table")
    df_by_role = _pivot_by("role_id",      "Role ID")
    df_by_section = _pivot_by("section",      "Section")
    df_by_source = _pivot_by("source_file",  "Source Workbook")

    # --- All statements sheet ---------------------------------------------
    df_all = df.copy()
    df_all.insert(0, "Statement #", range(1, len(df_all) + 1))
    df_all["is_companion"] = df_all["is_companion"].map(
        {True: "Yes", False: "No"})
    df_all = df_all.rename(columns={
        "operation":    "Operation",
        "target_table": "Target Table",
        "role_id":      "Role ID",
        "key_values":   "Key Columns",
        "section":      "Section",
        "source_file":  "Source Workbook",
        "is_companion": "Companion?",
        "sql_excerpt":  "SQL Excerpt",
    })
    df_all = df_all[["Statement #", "Operation", "Target Table", "Role ID",
                     "Key Columns", "Section", "Source Workbook",
                     "Companion?", "SQL Excerpt"]]

    # --- Write workbook ---------------------------------------------------
    try:
        with pd.ExcelWriter(_sql_long_path(str(output_xlsx_path)),
                            engine="xlsxwriter") as writer:
            wb = writer.book
            header_fmt = wb.add_format({
                "bold": True, "bg_color": "#305496", "font_color": "white",
                "align": "center", "valign": "vcenter", "border": 1,
            })
            total_fmt = wb.add_format({
                "bold": True, "bg_color": "#FFF2CC", "border": 1,
            })

            def _write_sheet(d, sheet, freeze=(1, 1), wide_last=False):
                d.to_excel(writer, sheet_name=sheet, index=False)
                ws = writer.sheets[sheet]
                for col, name in enumerate(d.columns):
                    ws.write(0, col, name, header_fmt)
                    sample = (d[name].head(80).astype(str).tolist()
                              if len(d) else [])
                    w = max(12, min(60, max(
                        [len(str(name))] + [len(s) for s in sample]) + 2))
                    if wide_last and col == len(d.columns) - 1:
                        w = max(w, 80)
                    ws.set_column(col, col, w)
                ws.freeze_panes(*freeze)
                if len(d):
                    ws.autofilter(0, 0, len(d), len(d.columns) - 1)

            _write_sheet(df_overview,   "Overview",            freeze=(1, 1))
            _write_sheet(df_by_table,   "By Table",            freeze=(1, 1))
            _write_sheet(df_by_role,    "By Role",             freeze=(1, 1))
            _write_sheet(df_by_section, "By Section",          freeze=(1, 1))
            _write_sheet(df_by_source,  "By Source Workbook",  freeze=(1, 1))
            _write_sheet(df_all,        "All Statements",      freeze=(1, 2),
                         wide_last=True)
    except Exception:
        log.error(f"[SQL-SUMMARY] Failed to write {output_xlsx_path}\n"
                  + traceback.format_exc())
        return None

    log.info(f"[SQL-SUMMARY] Workbook written → {output_xlsx_path}")
    return str(output_xlsx_path)


def stage5_general_sql():
    """Stage 5 entry point — validate every *_comparison.xlsx in
    COMPARISON_FOLDER/all_differences/ against Oracle LOV tables and
    write per-file *_remediation.sql + all_remediation.sql +
    validation_errors.xlsx into all_differences/generated_sql/."""
    global log
    log = _stage3_setup_logger()
    _ensure_stage3_imports()

    _banner("STAGE 5 — GENERAL_SQL  (all_differences/*.xlsx -> generated_sql/*.sql)")

    # --- Config validation ----------------------------------------------------
    if not UAT_CONFIG.get("user") or not UAT_CONFIG.get("dsn"):
        print("ERROR: stage 5 needs database credentials in config.yaml under "
              "`database:` (user, password, dsn). Aborting.")
        return {"stage": "general_sql", "ok": False, "error": "missing db config"}
    if not ORACLE_LIB_DIR:
        print("ERROR: stage 5 needs `paths.oracle_client_lib_dir` set in config.yaml. "
              "Aborting.")
        return {"stage": "general_sql", "ok": False, "error": "missing oracle lib dir"}

    diffs_dir = COMPARISON_FOLDER / "all_differences"
    if not diffs_dir.exists():
        print(f"ERROR: stage 5 expected {diffs_dir} to exist. "
              f"Did stage 3 run successfully?")
        return {"stage": "general_sql", "ok": False,
                "error": "all_differences folder missing"}

    xlsx_files = sorted(
        p for p in glob.glob(os.path.join(str(diffs_dir), "*_comparison.xlsx"))
        if not os.path.basename(p).startswith("~$")
    )
    xlsx_files = _drop_non_prod(xlsx_files, "general_sql", print)
    if not xlsx_files:
        print(f"WARNING: no *_comparison.xlsx files found under {diffs_dir}")
        return {"stage": "general_sql", "ok": True, "processed": 0,
                "sql_files": 0, "errors": 0}

    output_dir = diffs_dir / "generated_sql"
    output_dir.mkdir(parents=True, exist_ok=True)

    log.info(f"Input folder    : {diffs_dir}")
    log.info(f"Input files     : {len(xlsx_files)} workbook(s)")
    log.info(f"Output folder   : {output_dir}")
    log.info(
        f"Database DSN    : {UAT_CONFIG['dsn']} (user {UAT_CONFIG['user']})")
    if DB_LINK:
        log.info(f"Database link   : {DB_LINK}  "
                 f"(LOV queries rewritten as FCPROD.<table>@{DB_LINK})")
    else:
        log.info("Database link   : (none)")
    if os.name == "nt":
        log.info("Windows long-path mode: ENABLED (paths wrapped with \\\\?\\).")

    # --- Phase 1: fetch every LOV up front ----------------------------------
    try:
        lovs = _sql_fetch_lovs()
    except Exception as exc:
        log.error("Cannot fetch LOV data from Oracle. Aborting stage 5.\n"
                  + traceback.format_exc())
        return {"stage": "general_sql", "ok": False, "error": f"LOV fetch failed: {exc}"}

    # --- Phase 2: per-workbook validation + SQL generation ------------------
    all_output_files = []
    all_combined = []
    all_errors = []
    for xlsx_path in xlsx_files:
        out, errors = _sql_process_workbook(xlsx_path, str(output_dir), lovs)
        all_errors.extend(errors)
        if out:
            all_output_files.append(out)
            try:
                with open(_sql_long_path(out), encoding="utf-8") as f:
                    all_combined.append(f.read())
            except OSError:
                log.error(f"Cannot re-read SQL file for combined output: {out}\n"
                          + traceback.format_exc())

    # --- Phase 3: combined SQL + validation report --------------------------
    combined_path = None
    if all_combined:
        combined_path = output_dir / "all_remediation.sql"
        try:
            with open(_sql_long_path(str(combined_path)), "w", encoding="utf-8") as f:
                f.write(
                    f"-- {'═'*58}\n"
                    f"-- COMBINED Remediation SQL (stage 5)\n"
                    f"-- Generated : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                    f"-- Workbooks : {len(all_combined)}\n"
                    f"-- WARNING   : Review ALL statements before executing.\n"
                    f"-- {'═'*58}\n\n"
                )
                f.write("\n\n".join(all_combined))
            log.info(f"[SQL] Combined SQL written → {combined_path}")
        except OSError:
            log.error(f"Cannot write combined SQL file: {combined_path}\n"
                      + traceback.format_exc())

    err_report = _sql_write_validation_report(all_errors, str(output_dir))

    # --- Phase 4: audit workbook describing what all_remediation.sql does --
    # Parse the freshly-written combined SQL and emit a multi-sheet xlsx
    # for sign-off / review. Best-effort; failure here never poisons the
    # SQL deliverables that have already been written.
    summary_xlsx = None
    if combined_path and os.path.exists(_sql_long_path(str(combined_path))):
        summary_xlsx_path = output_dir / "all_remediation_summary.xlsx"
        summary_xlsx = _sql_summarize_remediation(
            combined_path, summary_xlsx_path)

    # --- Summary line ---------------------------------------------------------
    log.info(f"[SQL] SUMMARY: {len(xlsx_files)} workbook(s) processed, "
             f"{len(all_output_files)} SQL file(s) written, "
             f"{len(all_errors)} validation error(s).")

    return {
        "stage": "general_sql",
        "ok": True,
        "processed": len(xlsx_files),
        "sql_files": len(all_output_files),
        "errors": len(all_errors),
        "combined_sql": str(combined_path) if combined_path else None,
        "error_report": err_report,
        "summary_xlsx": summary_xlsx,
    }


# ============================================================================
# STAGE 6: SEND_REPORT  —  email key output files via Outlook
# ============================================================================
# Sends the three "headline" deliverables of a pipeline run via Outlook
# (using win32com.client, same pattern as the bank's existing reporting
# scripts). All settings live under `email:` in config.yaml:
#
#   email:
#     enabled: true
#     recipient: 'hany.fekry@nbe.com.eg'
#     subject: 'Access Matrix Pipeline Report — {run_folder}'
#     attachments:
#       - '04_Comparison/all_differences/_summary.xlsx'
#       - '04_Comparison/all_differences/generated_sql/all_remediation_summary.xlsx'
#       - '04_Comparison/all_differences/generated_sql/all_remediation.sql'
#
# Attachment paths are RELATIVE to the run folder (so they keep working as
# the run folder name changes each day). Missing attachments don't block
# sending — the email body lists which ones were missing so the recipient
# can chase them up.
#
# Default subject substitutions:
#     {run_folder}  — the run folder name (e.g. 'ACCESS MATRIX VS SYSTEM  RUN ...')
#     {date}        — ISO date (YYYY-MM-DD)


_DEFAULT_EMAIL_ATTACHMENTS = [
    "04_Comparison/all_differences/_summary.xlsx",
    "04_Comparison/all_differences/generated_sql/all_remediation_summary.xlsx",
    "04_Comparison/all_differences/generated_sql/all_remediation.sql",
    # Stage 1's per-run index workbook (filename has a timestamp, hence glob).
    "02_Reports/access_matrix_index_*.xlsx",
    # Per-run failures debug report, written by _write_failures_report when
    # any per-item failures occurred. Glob means the file is silently
    # skipped when there were no failures.
    "failures/failures_*.txt",
]


def _format_bytes(n):
    """Human-readable byte count: 1234 -> '1.2 KB'."""
    try:
        n = float(n)
    except (TypeError, ValueError):
        return "?"
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.1f} {unit}" if unit != "B" else f"{int(n)} {unit}"
        n /= 1024
    return f"{n:.1f} GB"


def _stage_detail_text(stage_name, s, recipient=None, n_attached=None,
                       cc_count=None, html=False):
    """Render the per-stage detail string that _print_pipeline_summary
    writes to the terminal. Used in two places:

      - in _print_pipeline_summary (with html=False) to print plain text
      - in _build_email_body (with html=True) to fill the PIPELINE
        COMPLETE table; the html variant uses <br> for line breaks so
        long detail strings split nicely in the email

    For stage 6 (send_report), the stage's own stats aren't yet in
    _RUN_STATS_BY_STAGE while the email is being assembled, so the caller
    can pass `recipient` + `n_attached` + `cc_count` to render an
    in-flight row.
    """
    br = "<br>" if html else "\n                     "
    if s is None:
        # Synthesized row for stage 6 in-flight
        if stage_name == "send_report":
            txt = (f"emailing {n_attached or 0} file(s) to "
                   f"{recipient or '(no recipient)'}")
            if cc_count:
                txt += f" (cc {cc_count})"
            return txt
        return "(no detail)"
    if not s.get("ok"):
        if s.get("skipped"):
            return f"skipped ({s.get('reason', '')})"
        return f"FAILED — {s.get('error', 'unknown error')}"
    if stage_name == "copy":
        return (f"{s.get('folders_scanned', 0)} folders scanned, "
                f"{s.get('copied', 0)} files copied, "
                f"{s.get('duplicates_skipped', 0)} duplicates, "
                f"{s.get('copy_errors', 0)} errors")
    if stage_name == "process":
        return (f"{s.get('processed', 0)} files processed "
                f"({s.get('successful', 0)} ok, "
                f"{s.get('failed', 0)} failed)")
    if stage_name == "compare":
        return (f"{s.get('processed', 0)} files compared "
                f"({s.get('successful', 0)} ok, "
                f"{s.get('failed', 0)} failed)")
    if stage_name == "summary_all_diff":
        text = f"{s.get('files_processed', 0)} comparison file(s) aggregated"
        out = s.get("output")
        if out:
            text += f"{br}→ {out}"
        return text
    if stage_name == "general_sql":
        text = (f"{s.get('processed', 0)} workbook(s) processed, "
                f"{s.get('sql_files', 0)} SQL file(s) written, "
                f"{s.get('errors', 0)} validation error(s)")
        audit = s.get("summary_xlsx")
        if audit:
            text += f"{br}audit workbook → {audit}"
        return text
    if stage_name == "send_report":
        if s.get("skipped"):
            return f"skipped ({s.get('reason', '')})"
        text = (f"emailed {s.get('attached', 0)} file(s) to "
                f"{s.get('recipient', '(?)')}")
        cc_list = s.get("cc") or []
        if cc_list:
            text += f" (cc {len(cc_list)})"
        if s.get("missing"):
            text += f" ⚠ {len(s['missing'])} attachment(s) missing"
        return text
    return str(s)


# --- Failure-reason helpers for the PIPELINE COMPLETE email table ----------

# Descriptive labels shown in the "Stage" column. send_report is intentionally
# absent — it never appears in the table.
_PIPELINE_STAGE_LABELS = {
    "copy":             "Copy from SMS-ITSOP share folder",
    "process":          "Process xlsx into CSV for cleanup",
    "compare":          "Compare CSV against system",
    "summary_all_diff": "Summary of all extracted differences",
    "general_sql":      "Generating SQL file for batch correction",
}

# Stage order for the table (no send_report).
_PIPELINE_TABLE_ORDER = (
    "copy", "process", "compare", "summary_all_diff", "general_sql",
)

# Oracle / python-oracledb connectivity errors. If every failure in a stage
# matches one of these, that stage's failure cell collapses to a single
# "Database is down" line instead of repeating the same ORA-xxxxx N times.
_DB_DOWN_RE = re.compile(
    r"ORA-12514|ORA-12541|ORA-12537|ORA-12170|ORA-12545|ORA-12528|"
    r"ORA-12521|ORA-01034|ORA-12560|DPY-4011|DPY-6005|DPY-3010|\bTNS:",
    re.IGNORECASE,
)


def _looks_like_db_down(text):
    """True if the failure text is an Oracle/listener connectivity error."""
    return bool(text and _DB_DOWN_RE.search(str(text)))


def _esc_html(s):
    """Minimal HTML escaping for failure text going into a table cell."""
    return (str(s).replace("&", "&amp;")
                  .replace("<", "&lt;")
                  .replace(">", "&gt;"))


def _stage_failure_reason_html(stage_name, max_lines=25):
    """Build the 'Failure reason (if any)' cell for one stage by reading the
    per-item failures already collected in _PIPELINE_FAILURES. Each failed
    item is shown as '<file>: <reason>' so the recipient can see exactly
    which xlsx/csv failed and why.

      * no failures          -> "" (empty cell)
      * all DB connectivity   -> "Database is down" (global outage; the
                                individual file names are omitted on purpose
                                — every file failed for the same reason)
      * otherwise            -> one '<file>: <reason>' line per failed item,
                                capped at `max_lines` with a "...and N more"
                                overflow line.
    """
    entries = []
    for f in _PIPELINE_FAILURES:
        if f.get("stage") != stage_name:
            continue
        name = (f.get("name") or "").strip()
        detail = (f.get("detail") or "").strip()
        code = (f.get("reason_code") or "").strip()
        entries.append((name, detail or code or "Unknown error"))

    if not entries:
        return ""

    # Global outage shortcut — don't list every file, just say it once.
    if all(_looks_like_db_down(reason) for _, reason in entries):
        return "Database is down"

    lines = []
    for name, reason in entries[:max_lines]:
        disp = reason
        if len(disp) > 200:
            disp = disp[:197] + "\u2026"
        if name:
            lines.append(f"<b>{_esc_html(name)}</b>: {_esc_html(disp)}")
        else:
            lines.append(_esc_html(disp))

    remaining = len(entries) - max_lines
    if remaining > 0:
        lines.append(f"\u2026and {remaining} more")
    return "<br>".join(lines)


def _build_email_body(found, missing, pending_send_info=None,
                      failures_report_path=None):
    """Compose the HTML body for the stage-6 report email.

    The PIPELINE COMPLETE table uses descriptive stage labels and carries a
    'Failure reason (if any)' column populated from _PIPELINE_FAILURES. The
    send_report stage and the To/CC recipients block are intentionally not
    rendered here.
    """
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    pending_send_info = pending_send_info or {}  # kept for signature compat

    # ----- PIPELINE COMPLETE table -------------------------------------
    rows = []
    for stage_name in _PIPELINE_TABLE_ORDER:
        s = _RUN_STATS_BY_STAGE.get(stage_name)
        if s is None:
            continue
        if s.get("ok"):
            icon = "\u2022" if s.get("skipped") else "\u2714"
        else:
            icon = "\u2718"
        label = _PIPELINE_STAGE_LABELS.get(stage_name, stage_name)
        detail = _stage_detail_text(stage_name, s, html=True)
        failure = _stage_failure_reason_html(stage_name)
        rows.append(
            f"<tr>"
            f"<td style='padding:6px 10px'>{label}</td>"
            f"<td style='padding:6px 10px; text-align:center; "
            f"font-size:14pt'>{icon}</td>"
            f"<td style='padding:6px 10px'>{detail}</td>"
            f"<td style='padding:6px 10px; color:#C00000'>{failure}</td>"
            f"</tr>"
        )

    if rows:
        pipeline_html = (
            "<h3 style='margin-bottom:4px'>PIPELINE COMPLETE</h3>"
            "<table border='1' cellpadding='0' "
            "style='border-collapse:collapse; font-size:10pt'>"
            "<tr style='background:#305496; color:white'>"
            "<th style='padding:6px 10px; text-align:left'>Stage</th>"
            "<th style='padding:6px 10px'>Status</th>"
            "<th style='padding:6px 10px; text-align:left'>Details</th>"
            "<th style='padding:6px 10px; text-align:left'>"
            "Failure reason (if any)</th></tr>"
            + "".join(rows) +
            "</table>"
        )
    else:
        pipeline_html = (
            "<h3>PIPELINE COMPLETE</h3>"
            "<p><i>Per-stage statistics not available — stage "
            "'send_report' was invoked on its own rather than as part of a "
            "full pipeline run. Attached files reflect the most recent "
            "outputs in the run folder.</i></p>"
        )

    # ----- Attachments table ------------------------------------------
    if found:
        att_rows = []
        for p in found:
            try:
                size_str = _format_bytes(p.stat().st_size)
            except OSError:
                size_str = "?"
            att_rows.append(
                f"<tr>"
                f"<td style='padding:4px 8px; font-family:Consolas,monospace'>"
                f"{p.name}</td>"
                f"<td style='padding:4px 8px; text-align:right'>{size_str}</td>"
                f"<td style='padding:4px 8px'><small>{p}</small></td>"
                f"</tr>"
            )
        attachments_html = (
            f"<h3>Attached Files ({len(found)})</h3>"
            "<table border='1' cellpadding='0' "
            "style='border-collapse:collapse; font-size:10pt'>"
            "<tr style='background:#305496; color:white'>"
            "<th style='padding:4px 8px; text-align:left'>File</th>"
            "<th style='padding:4px 8px'>Size</th>"
            "<th style='padding:4px 8px; text-align:left'>Full path</th></tr>"
            + "".join(att_rows) +
            "</table>"
        )
    else:
        attachments_html = (
            "<h3>Attached Files (0)</h3>"
            "<p><i>None of the configured attachments were found in the "
            "run folder. See 'Missing Attachments' below.</i></p>"
        )

    if missing:
        items = "".join(f"<li><code>{m}</code></li>" for m in missing)
        missing_html = (
            f"<h3 style='color:#C00000'>Missing Attachments "
            f"({len(missing)})</h3>"
            "<p>The following files were configured under "
            "<code>email.attachments</code> but were not present in the "
            "run folder:</p>"
            f"<ul>{items}</ul>"
            "<p>For literal file names this usually means the upstream "
            "stage that produces the file didn't run or produced no "
            "output. Glob patterns (with <code>*</code>) that match zero "
            "files don't appear here — they're skipped silently.</p>"
        )
    else:
        missing_html = ""

    # ----- Per-item failures snapshot ---------------------------------
    if _PIPELINE_FAILURES:
        by_stage_count = defaultdict(int)
        for f in _PIPELINE_FAILURES:
            by_stage_count[f.get("stage", "?")] += 1
        rows_f = "".join(
            f"<tr>"
            f"<td style='padding:4px 8px; font-family:Consolas,monospace'>"
            f"{stage}</td>"
            f"<td style='padding:4px 8px; text-align:center'>{n}</td>"
            f"</tr>"
            for stage, n in sorted(by_stage_count.items())
        )
        failures_note = ""
        if failures_report_path:
            failures_note = (
                f"<p>Full debug report (with tracebacks) attached: "
                f"<code>{Path(str(failures_report_path)).name}</code>.</p>"
            )
        failures_html = (
            f"<h3 style='color:#C00000'>Per-Item Failures "
            f"({len(_PIPELINE_FAILURES)} item(s))</h3>"
            "<p>These stages had at least one item (file, role, etc.) "
            "that couldn't be processed. The pipeline continued and "
            "produced what it could.</p>"
            "<table border='1' cellpadding='0' "
            "style='border-collapse:collapse; font-size:10pt'>"
            "<tr style='background:#C00000; color:white'>"
            "<th style='padding:4px 8px; text-align:left'>Stage</th>"
            "<th style='padding:4px 8px'>Failures</th></tr>"
            + rows_f + "</table>"
            + failures_note
        )
    else:
        failures_html = ""

    # NOTE: the To/CC recipients block has been removed on purpose.

    return f"""<html><body style="font-family:'Segoe UI', Arial, sans-serif; font-size:11pt">
<h2 style="color:#305496; margin-bottom:8px">Access Matrix Pipeline — Report Status</h2>

<table cellpadding="4" style="font-size:10pt">
  <tr><td><b>Run folder</b></td>
      <td><code>{RUN_FOLDER_NAME}</code></td></tr>
  <tr><td><b>Full path</b></td>
      <td><code>{RUN_FOLDER}</code></td></tr>
  <tr><td><b>Generated</b></td>
      <td>{now_str}</td></tr>
</table>

{pipeline_html}

{attachments_html}

{missing_html}

{failures_html}

<hr>
<p style="color:#888; font-size:9pt">
Generated automatically by access_matrix_pipeline.py (stage 6: send_report).
</p>
</body></html>"""


def stage6_send_report():
    """Stage 6 entry point — send a status email with the headline output
    files attached, via Outlook.

    Reads its settings from `email:` in config.yaml. When `email.enabled`
    is false, the stage is skipped (returns ok=True, skipped=True) so the
    pipeline can still be run end-to-end during testing without spamming
    real recipients.

    Attachment paths may be either literal file names or glob patterns
    (containing '*' or '?'). Glob patterns are expanded against the run
    folder and any matches are attached; patterns that match zero files
    are silently skipped (e.g. failures_*.txt when nothing failed). Literal
    paths that don't exist are reported in the email body's
    "Missing Attachments" section.

    If there are per-item failures from earlier stages, this stage writes
    the failures report file FIRST (so it can be attached) and stashes
    the path in _PIPELINE_FAILURES_REPORT_PATH; the orchestrator's
    end-of-run summary then skips re-writing it.
    """
    _banner("STAGE 6 — SEND_REPORT  (email summary + attachments)")

    email_cfg = CONFIG.get("email") or {}

    if not email_cfg.get("enabled", False):
        print("Email is disabled in config.yaml (email.enabled = false). "
              "Stage 6 will skip sending.")
        return {"stage": "send_report", "ok": True, "skipped": True,
                "reason": "email.enabled = false"}

    recipient = str(email_cfg.get("recipient", "")).strip()
    if not recipient:
        print("ERROR: email.recipient is not set in config.yaml. Cannot send.")
        return {"stage": "send_report", "ok": False,
                "error": "missing email.recipient"}

    # Normalise the CC list. Accepts either a YAML list:
    #   cc:
    #     - 'a@b.com'
    #     - 'c@d.com'
    # or a single string (comma- or semicolon-separated). Empty / missing
    # is fine and means no CC.
    raw_cc = email_cfg.get("cc")
    cc_list = []
    if raw_cc:
        if isinstance(raw_cc, (list, tuple)):
            cc_list = [str(x).strip() for x in raw_cc if str(x).strip()]
        else:
            cc_list = [
                x.strip() for x in re.split(r"[,;]", str(raw_cc))
                if x.strip()
            ]
    # De-dupe + drop anything that exactly matches the To address
    seen = set()
    cc_list = [
        x for x in cc_list
        if x and x.lower() != recipient.lower()
        and not (x.lower() in seen or seen.add(x.lower()))
    ]

    # Write the failures report now (if there are failures) so it can be
    # picked up by the glob in email.attachments. The orchestrator's
    # _print_pipeline_summary() will reuse this same file at end-of-run
    # rather than writing a second one with a slightly later timestamp.
    failures_report_path = None
    if _PIPELINE_FAILURES:
        failures_report_path = _write_failures_report(_PIPELINE_FAILURES)
        if failures_report_path:
            print(f"Failures report written for attachment: "
                  f"{failures_report_path}")

    rel_attachments = email_cfg.get(
        "attachments") or _DEFAULT_EMAIL_ATTACHMENTS

    # Resolve each attachment path. Relative paths are interpreted against
    # the run folder; absolute paths are used as-is. Patterns containing
    # '*' or '?' are expanded with glob.glob.
    found, missing, no_matches = [], [], []
    for rel in rel_attachments:
        rel_str = str(rel)
        is_glob = ("*" in rel_str) or ("?" in rel_str) or ("[" in rel_str)
        p = Path(rel_str)
        # Resolve relative-to-run-folder if not absolute
        if p.is_absolute():
            absolute = str(p)
        else:
            absolute = str(RUN_FOLDER / rel_str)

        if is_glob:
            matches = sorted(glob.glob(absolute))
            if matches:
                for m in matches:
                    mp = Path(m)
                    if mp.is_file() and mp not in found:
                        found.append(mp)
            else:
                # Glob with no matches: gloss over silently (e.g.
                # failures_*.txt when nothing failed). Note it for the log
                # but don't surface in "Missing Attachments" — that section
                # is for literal paths that should have been there.
                no_matches.append(rel_str)
        else:
            target = Path(absolute)
            if target.exists() and target.is_file():
                if target not in found:
                    found.append(target)
            else:
                missing.append(rel_str)

    # Build subject + body
    subject_template = email_cfg.get(
        "subject", "Access Matrix Pipeline Report — {run_folder}")
    try:
        subject = subject_template.format(
            run_folder=RUN_FOLDER_NAME,
            date=datetime.now().strftime("%Y-%m-%d"),
        )
    except (KeyError, IndexError) as exc:
        print(f"WARNING: bad token in email.subject template: {exc}. "
              f"Falling back to a default subject.")
        subject = f"Access Matrix Pipeline Report — {RUN_FOLDER_NAME}"

    # Console preview of what we're about to send
    print(f"Recipient        : {recipient}")
    if cc_list:
        print(f"CC               : {'; '.join(cc_list)}")
    print(f"Run folder       : {RUN_FOLDER}")
    print(f"Attachments found: {len(found)}")
    for p in found:
        try:
            sz = _format_bytes(p.stat().st_size)
        except OSError:
            sz = "?"
        print(f"  ✓ {p}  ({sz})")
    if missing:
        print(f"Attachments missing (expected): {len(missing)}")
        for r in missing:
            print(f"  ✘ {r}")
    if no_matches:
        print(f"Glob patterns with no matches (skipped): {len(no_matches)}")
        for r in no_matches:
            print(f"  · {r}")

    body = _build_email_body(
        found, missing,
        pending_send_info={
            "recipient":  recipient,
            "cc":         cc_list,
            "n_attached": len(found),
        },
        failures_report_path=failures_report_path,
    )

    # Outlook send. win32com is imported lazily so stages 1-5 still work
    # on machines that don't have pywin32 installed.
    try:
        import win32com.client
    except ImportError as exc:
        print(f"ERROR: cannot import win32com.client (need pywin32 on "
              f"Windows): {exc}")
        _record_failure(
            stage="send_report",
            name="outlook",
            reason_code="ImportError",
            detail=f"win32com.client not available: {exc}",
        )
        return {"stage": "send_report", "ok": False,
                "error": f"win32com missing: {exc}"}

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = recipient
        if cc_list:
            # Outlook expects a single string with addresses separated by
            # semicolons (or commas — semicolons are the safer choice on
            # the Windows / Outlook desktop client).
            mail.CC = "; ".join(cc_list)
        mail.Subject = subject
        mail.HTMLBody = body
        for p in found:
            try:
                mail.Attachments.Add(str(p.resolve()))
            except Exception as exc:
                print(f"  WARN: could not attach {p}: {exc}")
        mail.Send()
    except Exception as exc:
        print(f"ERROR: Outlook send failed: {type(exc).__name__}: {exc}")
        traceback.print_exc()
        _record_failure(
            stage="send_report",
            name="outlook",
            reason_code=type(exc).__name__,
            detail=str(exc),
            tb=traceback.format_exc(),
        )
        return {"stage": "send_report", "ok": False, "error": str(exc)}

    if cc_list:
        print(f"✓ Email sent to {recipient} (cc {len(cc_list)}) "
              f"with {len(found)} attachment(s).")
    else:
        print(f"✓ Email sent to {recipient} with {len(found)} attachment(s).")
    return {
        "stage":     "send_report",
        "ok":        True,
        "recipient": recipient,
        "cc":        cc_list,
        "attached":  len(found),
        "missing":   missing,
        "no_matches": no_matches,
        "subject":   subject,
    }


# ============================================================================
# PIPELINE ORCHESTRATOR
# ============================================================================

def _write_failures_report(failures):
    """Persist the cross-stage failures list into FAILURES_FOLDER as a
    timestamped text file. Returns the path written, or None on error.
    Caller should already have decided there's at least one failure."""
    try:
        FAILURES_FOLDER.mkdir(parents=True, exist_ok=True)
    except Exception as exc:
        print(f"[FAILURES] Could not create failures folder "
              f"{FAILURES_FOLDER}: {exc}")
        return None

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_path = FAILURES_FOLDER / f"failures_{stamp}.txt"

    # Build the report body. Group entries by stage so the user can scan
    # a failure cluster in one place. Each entry prints reason + detail +
    # input path; full tracebacks come last so they don't drown the
    # human-readable bit.
    by_stage = defaultdict(list)
    for f in failures:
        by_stage[f.get("stage", "?")].append(f)

    lines = []
    lines.append("=" * 80)
    lines.append(f"  ACCESS MATRIX PIPELINE — FAILURE REPORT")
    lines.append(
        f"  Generated : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(
        f"  Failures  : {len(failures)} across {len(by_stage)} stage(s)")
    lines.append("=" * 80)
    lines.append("")

    # Stage-by-stage human-readable section
    stage_order = ["copy", "process", "compare",
                   "summary_all_diff", "general_sql"]
    seen = set()
    ordered_stages = (
        [s for s in stage_order if s in by_stage]
        + [s for s in by_stage if s not in stage_order]
    )
    for stage in ordered_stages:
        stage_failures = by_stage[stage]
        if not stage_failures:
            continue
        lines.append("-" * 80)
        lines.append(f"  STAGE: {stage}   ({len(stage_failures)} failure(s))")
        lines.append("-" * 80)
        for i, f in enumerate(stage_failures, 1):
            lines.append("")
            lines.append(f"  [{i}] {f.get('name', '(unnamed)')}")
            lines.append(f"      reason : {f.get('reason_code', '')}")
            detail = f.get("detail") or ""
            # Wrap long detail lines so they don't overflow.
            for j, chunk in enumerate(detail.split("\n")):
                prefix = "      detail : " if j == 0 else "               "
                lines.append(prefix + chunk)
            if f.get("input_path"):
                lines.append(f"      source : {f['input_path']}")
        lines.append("")
        seen.update(id(x) for x in stage_failures)

    # Full tracebacks block at the bottom — useful for deep debugging
    # but kept separate so the top of the file stays scannable.
    tb_entries = [f for f in failures if f.get("traceback")]
    if tb_entries:
        lines.append("")
        lines.append("=" * 80)
        lines.append("  FULL TRACEBACKS")
        lines.append("=" * 80)
        for i, f in enumerate(tb_entries, 1):
            lines.append("")
            lines.append(
                f"  [{i}] stage={f.get('stage')}  name={f.get('name')}")
            lines.append("-" * 80)
            lines.append(f["traceback"].rstrip())
        lines.append("")

    try:
        with open(report_path, "w", encoding="utf-8") as fh:
            fh.write("\n".join(lines) + "\n")
        global _PIPELINE_FAILURES_REPORT_PATH
        _PIPELINE_FAILURES_REPORT_PATH = report_path
        return report_path
    except Exception as exc:
        print(f"[FAILURES] Could not write failures report "
              f"{report_path}: {exc}")
        return None


def _print_failures_block(failures, report_path):
    """Print the FAILURES section to the terminal. Compact one-line-per-
    failure format so a glance shows which item broke and why."""
    if not failures:
        return
    print()
    print("=" * 72)
    print(f"  FAILURES  ({len(failures)} item(s) need investigation)")
    print("=" * 72)
    # Group by stage for readability, same order as the report file.
    by_stage = defaultdict(list)
    for f in failures:
        by_stage[f.get("stage", "?")].append(f)
    stage_order = ["copy", "process", "compare",
                   "summary_all_diff", "general_sql"]
    for stage in stage_order + [s for s in by_stage if s not in stage_order]:
        stage_failures = by_stage.get(stage, [])
        if not stage_failures:
            continue
        print(f"  [{stage}]  {len(stage_failures)} failure(s):")
        for f in stage_failures:
            reason = f.get("reason_code") or "Error"
            name = f.get("name") or "(unnamed)"
            detail = (f.get("detail") or "").strip()
            # Truncate detail so the terminal stays readable; full text is in the file.
            if len(detail) > 140:
                detail = detail[:137] + "…"
            print(f"    ✘ {name}")
            print(f"        reason : {reason}")
            if detail:
                print(f"        detail : {detail}")
    if report_path:
        print()
        print(f"  Full report (with tracebacks): {report_path}")
    print("=" * 72)


def _print_pipeline_summary(stage_stats):
    """One-line-per-stage roll-up shown at the very end of the run, followed
    by a FAILURES block (and a written debug report) if any per-item
    failures were recorded across the stages."""
    print()
    print("=" * 72)
    print("  PIPELINE COMPLETE")
    print("=" * 72)
    for s in stage_stats:
        name = s.get("stage", "?")
        if not s.get("ok", False):
            print(f"  {name:<8s}: ✘ {s.get('error', 'failed')}")
            continue
        if name == "copy":
            print(f"  copy             : ✓ {s.get('folders_scanned', 0)} folders scanned, "
                  f"{s.get('copied', 0)} files copied, "
                  f"{s.get('duplicates_skipped', 0)} duplicates, "
                  f"{s.get('copy_errors', 0)} errors")
        elif name == "process":
            print(f"  process          : ✓ {s.get('processed', 0)} files processed "
                  f"({s.get('successful', 0)} ok, {s.get('failed', 0)} failed)")
        elif name == "compare":
            print(f"  compare          : ✓ {s.get('processed', 0)} files compared "
                  f"({s.get('successful', 0)} ok, {s.get('failed', 0)} failed)")
        elif name == "summary_all_diff":
            print(f"  summary_all_diff : ✓ {s.get('files_processed', 0)} comparison "
                  f"file(s) aggregated → {s.get('output', '(no output)')}")
        elif name == "general_sql":
            print(f"  general_sql      : ✓ {s.get('processed', 0)} workbook(s) processed, "
                  f"{s.get('sql_files', 0)} SQL file(s) written, "
                  f"{s.get('errors', 0)} validation error(s)")
            if s.get("summary_xlsx"):
                print(
                    f"                     audit workbook → {s['summary_xlsx']}")
        elif name == "send_report":
            if s.get("skipped"):
                print(f"  send_report      : • skipped "
                      f"({s.get('reason', 'email disabled in config')})")
            else:
                line = (f"  send_report      : ✓ emailed {s.get('attached', 0)} "
                        f"file(s) to {s.get('recipient', '(?)')}")
                cc_list = s.get("cc") or []
                if cc_list:
                    line += f" (cc {len(cc_list)})"
                if s.get("missing"):
                    line += f"  ⚠ {len(s['missing'])} attachment(s) missing"
                print(line)
        else:
            print(f"  {name:<16s} : ✓ {s}")
    print("=" * 72)

    # Surface any per-item failures with reasons + a debug report file.
    # If stage 6 (send_report) already wrote the report (so it could be
    # attached to the email), reuse that path instead of writing a second
    # copy with a slightly later timestamp.
    if _PIPELINE_FAILURES:
        if _PIPELINE_FAILURES_REPORT_PATH is not None:
            report_path = _PIPELINE_FAILURES_REPORT_PATH
        else:
            report_path = _write_failures_report(_PIPELINE_FAILURES)
        _print_failures_block(_PIPELINE_FAILURES, report_path)


def main():
    """Dispatch the requested stages in order:
    copy → process → compare → summary_all_diff → general_sql → send_report.
    If no stages are specified on the CLI, run all six."""
    _reset_failures()
    _RUN_STATS_BY_STAGE.clear()
    _, stages = _split_cli_args()
    if not stages:
        stages = ["copy", "process", "compare",
                  "summary_all_diff", "general_sql", "send_report"]

    # Preserve order, drop duplicates
    seen, ordered = set(), []
    for s in stages:
        if s not in seen:
            seen.add(s)
            ordered.append(s)

    print()
    print("=" * 72)
    print(f"  ACCESS MATRIX PIPELINE — running stages: {', '.join(ordered)}")
    print(f"  Config     : {CONFIG_PATH}")
    print(f"  Run folder : {RUN_FOLDER}")
    print("=" * 72)

    dispatcher = {
        "copy":             stage1_copy,
        "process":          stage2_process,
        "compare":          stage3_compare,
        "summary_all_diff": stage4_summary_all_diff,
        "general_sql":      stage5_general_sql,
        "send_report":      stage6_send_report,
    }

    stage_stats = []
    try:
        for stage in ordered:
            try:
                stats = dispatcher[stage]()
            except KeyboardInterrupt:
                print(f"\n[INTERRUPTED] during stage '{stage}'")
                raise
            except Exception as exc:
                print(f"\n[ERROR] Stage '{stage}' aborted: "
                      f"{type(exc).__name__}: {exc}")
                traceback.print_exc()
                err_stats = {"stage": stage, "ok": False, "error": str(exc)}
                stage_stats.append(err_stats)
                _RUN_STATS_BY_STAGE[stage] = err_stats
                # Don't try later stages if an earlier one failed — they'd
                # be operating on missing/incomplete inputs. Exception:
                # send_report still runs so the user gets an email about
                # the partial run (with whatever files were produced).
                if stage != "send_report" and "send_report" in ordered[
                        ordered.index(stage) + 1:]:
                    print("[INFO] Earlier stage failed, but send_report is "
                          "scheduled — running it so the user gets notified.")
                    try:
                        late_stats = stage6_send_report()
                        stage_stats.append(
                            late_stats or {"stage": "send_report", "ok": True})
                        _RUN_STATS_BY_STAGE["send_report"] = late_stats or {}
                    except Exception as exc2:
                        print(f"[ERROR] send_report also failed: {exc2}")
                break
            stage_stats.append(stats or {"stage": stage, "ok": True})
            _RUN_STATS_BY_STAGE[stage] = stats or {"stage": stage, "ok": True}
    finally:
        _print_pipeline_summary(stage_stats)


if __name__ == "__main__":
    main()
