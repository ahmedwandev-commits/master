"""
GW Amend Scripts Generator  (inc1 / inc2 / inc3)
================================================
Same idea as SA_USER.PY, but for the gateway "amend" configuration:
reads the rows of one or more (EXT_SYSTEM, SOURCE_OPERATION) pairs from the
SOURCE Oracle DB (UAT: fcuatdb-scan/PRODPDB) and writes the three ready-to-run
scripts for the TARGET DB (BS: fcshr-scan/PRODPDB) in ONE run:

    inc1.sql  ->  GWTM_AMEND_NODES    (one row per node)
    inc2.sql  ->  GWTM_AMEND_MASTER   (one row per ext-system / operation)
    inc3.sql  ->  GWTM_AMEND_FIELDS   (one row per node field)

HOW TO RUN
----------
  1. Edit the configuration below if needed (or use environment variables).
  2. Run the script:
         python GW_AMEND_SCRIPTS.py
             -> uses DEFAULT_PAIRS below (SUBMWFCPR / ModifyCustomer)
         python GW_AMEND_SCRIPTS.py --ext-system SUBMWFCPR --operation ModifyCustomer
         python GW_AMEND_SCRIPTS.py --ext-system SUBMWFCPR
             -> every SOURCE_OPERATION of that ext system
         python GW_AMEND_SCRIPTS.py --pairs "SUBMWFCPR:ModifyCustomer;SUBMWFCPR:CreateCustomer"
         python GW_AMEND_SCRIPTS.py --excel "D:\\...\\Operations.xlsx"
             -> Excel with columns EXT_SYSTEM and SOURCE_OPERATION
                (SOURCE_OPERATION may be empty = all operations)
         python GW_AMEND_SCRIPTS.py --skip-existing
             -> every INSERT gets "WHERE NOT EXISTS" so the scripts can be
                re-run on TARGET without ORA-00001 (nothing is ever deleted)
  3. Output (next to this script, in generated_amend_sql\\):
         inc1.sql, inc2.sql, inc3.sql   <- the three scripts
         run_all.sql                    <- runs them in the safe order
                                           (inc2 master -> inc1 nodes -> inc3 fields)

Key design rules
----------------
  - Only SELECT queries are run against SOURCE. Nothing is written anywhere
    except the output files.
  - Column lists are taken from the SOURCE table itself (SELECT *, in table
    column order), so if a column is added in UAT it is carried over
    automatically.
  - Date columns keep their ORIGINAL value from SOURCE (like the manual
    scripts: to_date('01-03-2016', 'dd-mm-yyyy')), with the time added only
    when it is not midnight.
  - Rows are sorted by their key columns (binary order, same as the manual
    scripts) so two runs give identical files.

Requirements
------------
  pip install oracledb openpyxl      (openpyxl only needed for --excel)
"""

from __future__ import annotations

import argparse
import getpass
import os
import re
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

try:
    import oracledb
except ImportError as exc:
    raise SystemExit(
        "Missing dependency: oracledb\n"
        "Install with:  pip install oracledb"
    ) from exc


# =============================================================================
# ★  CONFIGURATION  — edit these values if anything changes  ★
# =============================================================================

# ── SOURCE DB (read-only; UAT / PRODPDB on fcuatdb) ──────────────────────────
# Alternative environments:
#     PREPROD : fcpreprd-scan:1521/PRODPDB
#     BS/SHR  : fcshr-scan:1521/PRODPDB
#     UAT     : fcuatdb-scan:1521/PRODPDB   (default below)
# Password: set the SRC_PWD environment variable, or the script asks for it.
SRC_USER   = os.environ.get("SRC_USER",   "A_231232")
SRC_PWD    = os.environ.get("SRC_PWD",    "")
SRC_DSN    = os.environ.get("SRC_DSN",    "fcuatdb-scan:1521/PRODPDB")   # UAT
SRC_SCHEMA = os.environ.get("SRC_SCHEMA", "FCPROD")

# ── TARGET DB (only written into the script headers) ─────────────────────────
TGT_DSN    = "fcshr-scan:1521/PRODPDB"      # BS
# Schema prefix used in the generated INSERTs. "" = no prefix, exactly like
# the manual inc scripts (run them connected as the owner). e.g. "FCPROD".
TGT_SCHEMA = ""

# ── Oracle Instant Client (thick mode; leave as-is if not needed) ────────────
ORA_CLIENT_DIR = r"C:\oracle\instantclient_23_9"

# ── Default pairs used when nothing is given on the command line ─────────────
#    (EXT_SYSTEM, SOURCE_OPERATION)   — SOURCE_OPERATION None = all operations
DEFAULT_PAIRS: List[Tuple[str, Optional[str]]] = [
    ("SUBMWFCPR", "ModifyCustomer"),
]

# ── Output ───────────────────────────────────────────────────────────────────
HERE    = Path(__file__).parent
OUT_DIR = HERE / "generated_amend_sql"

# ── The three scripts ────────────────────────────────────────────────────────
# (file name, table, key columns used for ORDER BY / WHERE NOT EXISTS)
SCRIPTS: List[Tuple[str, str, List[str]]] = [
    ("inc1.sql", "GWTM_AMEND_NODES",  ["EXT_SYSTEM", "SOURCE_OPERATION", "NODE_NAME"]),
    ("inc2.sql", "GWTM_AMEND_MASTER", ["EXT_SYSTEM", "SOURCE_OPERATION"]),
    ("inc3.sql", "GWTM_AMEND_FIELDS", ["EXT_SYSTEM", "SOURCE_OPERATION", "NODE_NAME", "FIELD_NAME"]),
]
# Safe run order on TARGET: parent (master) first, then nodes, then fields.
RUN_ORDER = ["inc2.sql", "inc1.sql", "inc3.sql"]


# =============================================================================
# Input: which (EXT_SYSTEM, SOURCE_OPERATION) pairs
# =============================================================================

Pair = Tuple[str, Optional[str]]


def _clean(v: Any) -> Optional[str]:
    if v is None:
        return None
    s = str(v).strip()
    return s or None


def parse_pairs(raw: str) -> List[Pair]:
    """'SYS1:Op1;SYS1:Op2;SYS2' -> [(SYS1,Op1),(SYS1,Op2),(SYS2,None)]"""
    pairs: List[Pair] = []
    for item in re.split(r"[;,\s]+", raw.strip()):
        if not item:
            continue
        ext, _, op = item.partition(":")
        ext = _clean(ext)
        if ext:
            pairs.append((ext, _clean(op)))
    return pairs


def read_pairs_from_excel(path: Path, sheet: Any = 0) -> List[Pair]:
    try:
        import openpyxl
    except ImportError as exc:
        raise SystemExit("Missing dependency: openpyxl  (pip install openpyxl)") from exc
    if not path.exists():
        raise SystemExit(f"❌  Excel file not found: {path}")

    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    ws = wb.worksheets[sheet] if isinstance(sheet, int) else wb[sheet]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()

    ext_idx = op_idx = hdr_idx = None
    for i, row in enumerate(rows[:10]):
        hdr = [str(c).strip().upper().replace(" ", "_") if c is not None else "" for c in row]
        if "EXT_SYSTEM" in hdr:
            hdr_idx, ext_idx = i, hdr.index("EXT_SYSTEM")
            for name in ("SOURCE_OPERATION", "OPERATION", "SOURCE_OP"):
                if name in hdr:
                    op_idx = hdr.index(name)
                    break
            break
    if hdr_idx is None:
        raise SystemExit("❌  Excel: no 'EXT_SYSTEM' header found in the first 10 rows.")

    pairs: List[Pair] = []
    for row in rows[hdr_idx + 1:]:
        ext = _clean(row[ext_idx]) if ext_idx < len(row) else None
        op  = _clean(row[op_idx]) if op_idx is not None and op_idx < len(row) else None
        if ext:
            pairs.append((ext, op))
    return pairs


def dedupe_pairs(pairs: List[Pair]) -> List[Pair]:
    """Remove duplicates; (EXT, None) already covers every (EXT, op)."""
    all_ops = {e for e, o in pairs if o is None}
    seen, out = set(), []
    for e, o in pairs:
        if o is not None and e in all_ops:
            continue
        if (e, o) not in seen:
            seen.add((e, o))
            out.append((e, o))
    return out


# =============================================================================
# Oracle (SOURCE, read-only)
# =============================================================================

def connect_oracle() -> "oracledb.Connection":
    pwd = SRC_PWD or getpass.getpass(f"  Password for {SRC_USER}@{SRC_DSN}: ")
    try:
        oracledb.init_oracle_client(lib_dir=ORA_CLIENT_DIR)
    except Exception:
        pass   # already initialised, or thin mode
    return oracledb.connect(user=SRC_USER, password=pwd, dsn=SRC_DSN)


def fetch_rows(
    conn:  "oracledb.Connection",
    table: str,
    pairs: List[Pair],
) -> Tuple[List[str], List[Dict[str, Any]]]:
    """Return (columns in table order, rows) for every requested pair."""
    conds, binds = [], {}
    for i, (ext, op) in enumerate(pairs):
        binds[f"e{i}"] = ext
        if op is None:
            conds.append(f"(EXT_SYSTEM = :e{i})")
        else:
            binds[f"o{i}"] = op
            conds.append(f"(EXT_SYSTEM = :e{i} AND SOURCE_OPERATION = :o{i})")
    sql = f"SELECT * FROM {SRC_SCHEMA}.{table} WHERE " + " OR ".join(conds)

    with conn.cursor() as cur:
        cur.arraysize = 2000
        cur.execute(sql, binds)
        cols = [d[0].upper() for d in cur.description]
        rows = [dict(zip(cols, r)) for r in cur.fetchall()]
    return cols, rows


# =============================================================================
# SQL generation  (same layout as the manual inc scripts)
# =============================================================================

def _lit(v: Any) -> str:
    if v is None:
        return "null"
    if isinstance(v, datetime):
        if (v.hour, v.minute, v.second) == (0, 0, 0):
            return f"to_date('{v:%d-%m-%Y}', 'dd-mm-yyyy')"
        return f"to_date('{v:%d-%m-%Y %H:%M:%S}', 'dd-mm-yyyy hh24:mi:ss')"
    if isinstance(v, date):
        return f"to_date('{v:%d-%m-%Y}', 'dd-mm-yyyy')"
    if isinstance(v, bool):
        return "'Y'" if v else "'N'"
    if isinstance(v, (int, Decimal)):
        return str(v)
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else repr(v)
    if isinstance(v, bytes):
        return f"hextoraw('{v.hex().upper()}')"
    if hasattr(v, "read"):                 # CLOB / LOB
        v = v.read()
    return "'" + str(v).replace("'", "''") + "'"


def _sort_key(row: Dict[str, Any], keys: List[str]) -> tuple:
    # Binary (ASCII) order on the key columns, like Oracle's BINARY sort.
    return tuple("" if row.get(k) is None else str(row[k]) for k in keys if k in row)


def build_script(
    file_name:     str,
    table:         str,
    key_cols:      List[str],
    cols:          List[str],
    rows:          List[Dict[str, Any]],
    pairs:         List[Pair],
    skip_existing: bool,
) -> str:
    tgt      = f"{TGT_SCHEMA}.{table}" if TGT_SCHEMA else table
    tgt_low  = tgt.lower()
    col_list = ", ".join(cols)
    keys     = [k for k in key_cols if k in cols] or cols

    out: List[str] = [
        "-- ================================================================",
        f"-- {file_name}  ->  {table}",
        f"-- Generated     : {datetime.now():%Y-%m-%d %H:%M}",
        f"-- Source DB     : {SRC_DSN}  (schema {SRC_SCHEMA})",
        f"-- Target DB     : {TGT_DSN}  <-- run here",
        "-- For           : " + ", ".join(f"{e}/{o or '*'}" for e, o in pairs),
        f"-- Rows          : {len(rows)}",
        "-- ================================================================",
        "SET DEFINE OFF;",
        "",
    ]

    for row in sorted(rows, key=lambda r: _sort_key(r, keys)):
        vals = ", ".join(_lit(row[c]) for c in cols)
        if skip_existing:
            where = " and ".join(
                f"{k} = {_lit(row[k])}" if row[k] is not None else f"{k} is null"
                for k in keys
            )
            out.append(f"insert into {tgt_low} ({col_list})")
            out.append(f"select {vals} from dual")
            out.append(f"where not exists (select 1 from {tgt_low} where {where});")
        else:
            out.append(f"insert into {tgt_low} ({col_list})")
            out.append(f"values ({vals});")
        out.append("")

    if not rows:
        out.append(f"-- NOTE: no rows found in SOURCE {table} for the requested pairs.")
    out.append("commit;")
    return "\n".join(out) + "\n"


def build_run_all(out_dir: Path) -> str:
    lines = [
        "-- Runs the three amend scripts on TARGET in the safe order:",
        "--   inc2 (master) -> inc1 (nodes) -> inc3 (fields)",
        "SET DEFINE OFF;",
        f"SPOOL {out_dir.name}_run_all.log",
    ]
    for f in RUN_ORDER:
        lines += [f"PROMPT ===== {f} =====", f"@@{f}"]
    lines.append("SPOOL OFF;")
    return "\n".join(lines) + "\n"


# =============================================================================
# Entry point
# =============================================================================

def main() -> int:
    p = argparse.ArgumentParser(
        description="Generate inc1/inc2/inc3 (GWTM_AMEND_NODES/MASTER/FIELDS) "
                    "from SOURCE (UAT) for TARGET (BS) in one run.",
    )
    p.add_argument("--ext-system", metavar="EXT", help="EXT_SYSTEM, e.g. SUBMWFCPR")
    p.add_argument("--operation", metavar="OP",
                   help="SOURCE_OPERATION, e.g. ModifyCustomer (omit = all operations)")
    p.add_argument("--pairs", metavar="EXT:OP;...",
                   help='Several pairs: "SUBMWFCPR:ModifyCustomer;SUBMWFCPR:CreateCustomer"')
    p.add_argument("--excel", metavar="PATH",
                   help="Excel with columns EXT_SYSTEM and SOURCE_OPERATION")
    p.add_argument("--sheet", default="0", help="Excel sheet name or 0-based index")
    p.add_argument("--output-dir", default=str(OUT_DIR), metavar="DIR")
    p.add_argument("--skip-existing", action="store_true",
                   help="Generate INSERT ... WHERE NOT EXISTS (re-runnable on TARGET)")
    args = p.parse_args()

    # ── pairs ────────────────────────────────────────────────────────────────
    pairs: List[Pair] = []
    if args.pairs:
        pairs += parse_pairs(args.pairs)
    if args.ext_system:
        pairs.append((args.ext_system.strip(), _clean(args.operation)))
    if args.excel:
        sheet: Any = int(args.sheet) if args.sheet.isdigit() else args.sheet
        pairs += read_pairs_from_excel(Path(args.excel), sheet)
    if not pairs:
        pairs = list(DEFAULT_PAIRS)
    pairs = dedupe_pairs(pairs)

    out_dir = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    print("=" * 60)
    print("  GW Amend Scripts Generator  (inc1 / inc2 / inc3)")
    print("=" * 60)
    print(f"  Source DB  : {SRC_DSN}  (schema {SRC_SCHEMA})")
    print(f"  Target DB  : {TGT_DSN}")
    print(f"  Output dir : {out_dir.resolve()}")
    print(f"  Mode       : {'INSERT ... WHERE NOT EXISTS' if args.skip_existing else 'plain INSERT'}")
    for e, o in pairs:
        print(f"    • {e} / {o or '* (all operations)'}")
    print()

    print(f"Connecting to {SRC_DSN} …")
    try:
        conn = connect_oracle()
    except Exception as exc:
        raise SystemExit(f"❌  Could not connect to Oracle:\n   {exc}") from exc
    print("  ✔  Connected.\n")

    try:
        for file_name, table, key_cols in SCRIPTS:
            cols, rows = fetch_rows(conn, table, pairs)
            sql = build_script(file_name, table, key_cols, cols, rows, pairs, args.skip_existing)
            (out_dir / file_name).write_text(sql, encoding="utf-8")
            flag = "" if rows else "   ⚠  no rows!"
            print(f"  ✔  {file_name:<9} {table:<18} {len(rows):>5} row(s){flag}")
    finally:
        conn.close()

    (out_dir / "run_all.sql").write_text(build_run_all(out_dir), encoding="utf-8")
    print(f"  ✔  run_all.sql  (order: {' -> '.join(RUN_ORDER)})")
    print("\nDone.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
