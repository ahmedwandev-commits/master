"""
GW Amend Scripts Generator  (inc1 / inc2 / inc3)
================================================
Same idea as SA_USER.PY, but for the gateway "amend" configuration:
reads the rows of one or more (EXT_SYSTEM, SOURCE_OPERATION) pairs from the
SOURCE Oracle DB (UAT: fcuatdb-scan/PRODPDB) and writes the three ready-to-run
scripts for the TARGET DB (BS: fcshr-scan/PRODPDB) in ONE run, all in ONE file:

    inc1.sql  ->  GWTM_AMEND_NODES    (one row per node)
    inc2.sql  ->  GWTM_AMEND_MASTER   (one row per ext-system / operation)
    inc3.sql  ->  GWTM_AMEND_FIELDS   (one row per node field)

HOW TO RUN
----------
  "User" here = the EXT_SYSTEM value (e.g. SUBMWFCPR). For every user, all of
  its rows in the 3 tables are taken (all SOURCE_OPERATIONs, unless
  --operation is given).

  1. Edit the configuration below if needed (or use environment variables).
  2. Give the users the same way as SA_USER.PY:
         python GW_AMEND_SCRIPTS.py
             -> reads UserGW.xlsx (EXCEL_FILE; column USER_ID / EXT_SYSTEM, else column A);
                if that file doesn't exist it asks you to type the users
         python GW_AMEND_SCRIPTS.py --users SUBMWFCPR,SUBMWFCXX
         python GW_AMEND_SCRIPTS.py --users-file users.txt
         python GW_AMEND_SCRIPTS.py --excel "D:\\...\\Users.xlsx"
         python GW_AMEND_SCRIPTS.py --users SUBMWFCPR --operation ModifyCustomer
             -> only that operation
         python GW_AMEND_SCRIPTS.py --users SUBMWFCPR --skip-existing
             -> every INSERT gets "WHERE NOT EXISTS" so the scripts can be
                re-run on TARGET without ORA-00001 (nothing is ever deleted)
  3. Output: ONE file  generated_amend_sql\\GW_AMEND_SCRIPT.sql  (or --output)
     holding inc1 + inc2 + inc3 together, in the safe order
     inc2 (master) -> inc1 (nodes) -> inc3 (fields), with one COMMIT at the end.

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
# Same login as SA_USER.PY (environment variables override these defaults).
SRC_USER   = os.environ.get("SRC_USER",   "A_231232")
SRC_PWD    = os.environ.get("SRC_PWD",    "")   # set SRC_PWD or you will be asked
SRC_DSN    = os.environ.get("SRC_DSN",    "fcuatdb-scan:1521/PRODPDB")   # UAT
SRC_SCHEMA = os.environ.get("SRC_SCHEMA", "FCPROD")

# ── TARGET DB (only written into the script headers) ─────────────────────────
TGT_DSN    = "fcshr-scan:1521/PRODPDB"      # BS
# Schema prefix used in the generated INSERTs. "" = no prefix, exactly like
# the manual inc scripts (run them connected as the owner). e.g. "FCPROD".
TGT_SCHEMA = ""

# ── Oracle Instant Client (thick mode; leave as-is if not needed) ────────────
ORA_CLIENT_DIR = r"C:\oracle\instantclient_23_9"

# ── USERS (= EXT_SYSTEM values) — same input style as SA_USER.PY ─────────────
# Priority: --users  >  --users-file  >  --excel / EXCEL_FILE  >  typed at prompt
EXCEL_FILE = Path(
    r"D:\Felex\access matrix  vs system pipeline\BS Scripts\Scripts\SA_USERS\UserGW.xlsx"
)
EXCEL_SHEET = 0
EXCEL_USER_ID_COLUMN_ALIASES = [
    "USER_ID", "USERID", "USER ID", "EXT_SYSTEM", "EXT SYSTEM",
    "ID", "USERNAME", "USER NAME",
]

# ── Output ───────────────────────────────────────────────────────────────────
HERE    = Path(__file__).parent
OUT_FILE   = HERE / "generated_amend_sql" / "GW_AMEND_SCRIPT.sql"   # inc1+inc2+inc3 together
SPOOL_FILE = r"D:\GW_AMEND.TXT"                                   # log when run on TARGET

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
# Input: which users (EXT_SYSTEM values)   — same rules as SA_USER.PY
# =============================================================================

Pair = Tuple[str, Optional[str]]      # (EXT_SYSTEM, SOURCE_OPERATION or None=all)


def _clean(v: Any) -> Optional[str]:
    if v is None or isinstance(v, (bool, date, datetime)):
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    return s or None


def parse_user_ids(raw: str) -> List[str]:
    """Users separated by comma / space / semicolon / new line."""
    skip = {a.upper() for a in EXCEL_USER_ID_COLUMN_ALIASES}
    seen, out = set(), []
    for item in re.split(r"[\s,;]+", raw.strip()):
        uid = _clean(item)
        if uid and uid.upper() not in skip and uid not in seen:
            seen.add(uid)
            out.append(uid)
    return out


def read_user_ids_from_excel(path: Path, sheet: Any = EXCEL_SHEET) -> List[str]:
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

    # header row (first 10 rows) with one of the aliases, else column A
    col, start = 0, 0
    for i, row in enumerate(rows[:10]):
        hdr = [str(c).strip().upper() if c is not None else "" for c in row]
        hit = next((hdr.index(a) for a in EXCEL_USER_ID_COLUMN_ALIASES if a in hdr), None)
        if hit is not None:
            col, start = hit, i + 1
            print(f"  ✔  Excel: user column '{row[hit]}' (column {hit + 1}, header row {i + 1})")
            break
    else:
        print("  ⚠  Excel: no user header found — reading column A.")

    vals = [str(_clean(r[col])) for r in rows[start:] if col < len(r) and _clean(r[col])]
    return parse_user_ids("\n".join(vals))


def load_user_ids(args: argparse.Namespace) -> List[str]:
    if args.users:
        print("  ℹ  Users from: --users")
        return parse_user_ids(args.users)
    if args.users_file:
        p = Path(args.users_file)
        if not p.exists():
            raise SystemExit(f"❌  Users file not found: {p}")
        print(f"  ℹ  Users from: text file {p}")
        return parse_user_ids(p.read_text(encoding="utf-8", errors="ignore"))
    excel = Path(args.excel) if args.excel else EXCEL_FILE
    if args.excel or excel.exists():
        sheet: Any = int(args.sheet) if str(args.sheet).isdigit() else args.sheet
        print(f"  ℹ  Users from: Excel {excel}")
        return read_user_ids_from_excel(excel, sheet)
    print(f"  ℹ  Excel not found ({excel}).")
    return parse_user_ids(input("  Type the user(s) (EXT_SYSTEM), comma separated: "))


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
    cols: List[str] = []
    rows: List[Dict[str, Any]] = []
    for start in range(0, len(pairs), 500):          # bounded bind count per query
        chunk = pairs[start:start + 500]
        conds, binds = [], {}
        for i, (ext, op) in enumerate(chunk):
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
            rows += [dict(zip(cols, r)) for r in cur.fetchall()]
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


def build_section(
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
        "",
        "-- ================================================================",
        f"-- {file_name.replace('.sql', '')}  ->  {table}   ({len(rows)} rows)",
        "-- ================================================================",
        f"PROMPT {file_name.replace('.sql', '')} : {table} ({len(rows)} rows)",
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
        out.append(f"-- NOTE: no rows found in SOURCE {table} for the requested users.")
    return "\n".join(out) + "\n"


def build_combined(sections: Dict[str, str], pairs: List[Pair]) -> str:
    """inc1 + inc2 + inc3 in ONE file, in RUN_ORDER, with a single COMMIT."""
    head = [
        "-- ================================================================",
        "-- GW amend script  (inc1 + inc2 + inc3 in one file)",
        f"-- Generated     : {datetime.now():%Y-%m-%d %H:%M}",
        f"-- Source DB     : {SRC_DSN}  (schema {SRC_SCHEMA})",
        f"-- Target DB     : {TGT_DSN}  <-- run here",
        "-- Users         : " + ", ".join(sorted({e for e, _ in pairs})),
        "-- Order         : " + " -> ".join(f.replace(".sql", "") for f in RUN_ORDER)
        + "   (master first, then nodes, then fields)",
        "-- ================================================================",
        "SET DEFINE OFF;",
        f"SPOOL {SPOOL_FILE}",
    ]
    body = "".join(sections[f] for f in RUN_ORDER)
    return "\n".join(head) + "\n" + body + "\ncommit;\nSPOOL OFF;\n"


# =============================================================================
# Entry point
# =============================================================================

def main() -> int:
    p = argparse.ArgumentParser(
        description="Generate inc1/inc2/inc3 (GWTM_AMEND_NODES/MASTER/FIELDS) "
                    "from SOURCE (UAT) for TARGET (BS) in one run.",
    )
    p.add_argument("--users", metavar="U1,U2,...",
                   help="Users (EXT_SYSTEM values), e.g. SUBMWFCPR,SUBMWFCXX — overrides Excel")
    p.add_argument("--users-file", metavar="FILE",
                   help="Text file with users (one per line or comma separated)")
    p.add_argument("--excel", metavar="PATH", help=f"Excel with a USER_ID column (default: {EXCEL_FILE})")
    p.add_argument("--sheet", default=str(EXCEL_SHEET), help="Excel sheet name or 0-based index")
    p.add_argument("--operation", metavar="OP",
                   help="Optional: only this SOURCE_OPERATION (e.g. ModifyCustomer). Default = all.")
    p.add_argument("--output", default=str(OUT_FILE), metavar="FILE",
                   help=f"The single output SQL file (default: {OUT_FILE})")
    p.add_argument("--skip-existing", action="store_true",
                   help="Generate INSERT ... WHERE NOT EXISTS (re-runnable on TARGET)")
    args = p.parse_args()

    # ── users ────────────────────────────────────────────────────────────────
    user_ids = load_user_ids(args)
    if not user_ids:
        raise SystemExit("❌  No users given.")
    op = _clean(args.operation)
    pairs: List[Pair] = [(u, op) for u in user_ids]

    out_file = Path(args.output)
    out_file.parent.mkdir(parents=True, exist_ok=True)

    print("=" * 60)
    print("  GW Amend Scripts Generator  (inc1 / inc2 / inc3)")
    print("=" * 60)
    print(f"  Source DB  : {SRC_DSN}  (schema {SRC_SCHEMA})")
    print(f"  Target DB  : {TGT_DSN}")
    print(f"  Output     : {out_file.resolve()}")
    print(f"  Mode       : {'INSERT ... WHERE NOT EXISTS' if args.skip_existing else 'plain INSERT'}")
    print(f"  Users      : {len(user_ids)}   operation: {op or '* (all)'}")
    for u in user_ids[:10]:
        print(f"    • {u}")
    if len(user_ids) > 10:
        print(f"    … and {len(user_ids) - 10} more")
    print()

    print(f"Connecting to {SRC_DSN} …")
    try:
        conn = connect_oracle()
    except Exception as exc:
        raise SystemExit(f"❌  Could not connect to Oracle:\n   {exc}") from exc
    print("  ✔  Connected.\n")

    sections: Dict[str, str] = {}
    try:
        for file_name, table, key_cols in SCRIPTS:
            cols, rows = fetch_rows(conn, table, pairs)
            sections[file_name] = build_section(
                file_name, table, key_cols, cols, rows, pairs, args.skip_existing
            )
            flag = "" if rows else "   ⚠  no rows!"
            print(f"  ✔  {file_name[:-4]:<5} {table:<18} {len(rows):>5} row(s){flag}")
    finally:
        conn.close()

    out_file.write_text(build_combined(sections, pairs), encoding="utf-8")
    print(f"\n  ✔  One file written: {out_file.resolve()}")
    print("\nDone.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
