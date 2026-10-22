#!/usr/bin/env python3
"""
Loxley "Green Nut Daily Sales & Stock" Excel -> sales-out database.

Loxley Trading (our distributor) emails a workbook every working day, e.g.
"Green Nut Daily Sales & Stock as of 22 Oct 2026" with attachment
"Green Nut Daily Stock 2026 22.10.26.xlsx". The workbook has one sheet per
month (e.g. "Oct.26", "Sep.26", "June25", "Jan24", ...). Every month sheet
shares the same *logical* layout, but column positions move every month, so
this parser locates every block by its header text (row 4 = block title,
row 5 = date label) rather than by fixed column letters.

Per month sheet, one row per SKU (Part code), with these blocks:

  Part | Description | LTP | Avg Sale/mth | New Weight | Forecast
  Actual  <date> <date> ...        cumulative month-to-date (MTD) sales-out qty
  %                                MTD actual / forecast
  Plan PO | Add PO | Pending PO | Total PO
  PO Arrival in <month>  1..31 | Total   PO quantity received per day
  PO ขาดส่ง                         PO outstanding (not yet delivered)
  Available Stock <date> <date>    Loxley warehouse stock per day
  Cover Day | Demand | Promotion | ของเข้าวันที่ (incoming-goods note)

Key point: the "Actual" columns are CUMULATIVE month-to-date, recorded only
on report days. Daily sales-out = MTD(today) - MTD(previous report in the
same month). That difference is exposed by the view v_sales_out_daily.

Hidden reference sheets also loaded:
  DATA_ไซส์    SKU -> price point (Baht), selling channel / retailer notes
  DATA_Cocode  Loxley customer master -> channel group (7-11, Makro, Lotus, CJ...)

The script is idempotent: re-ingesting the same file (or an older file whose
month sheets overlap) upserts, never duplicates. Later reports win.

Usage
-----
  # daily run: ingest the newest email attachment (current + previous month)
  python loxley_sales_parser.py ingest "Green Nut Daily Stock 2026 22.10.26.xlsx"

  # one-off historical backfill from every month sheet in the workbook
  python loxley_sales_parser.py ingest file.xlsx --all-sheets

  # Postgres / Supabase instead of the default SQLite file
  python loxley_sales_parser.py ingest file.xlsx --db postgresql://user:pw@host:5432/db

  # quick look at what is in the database
  python loxley_sales_parser.py summary

The database URL can also come from the LOXLEY_DB_URL environment variable.
"""
from __future__ import annotations

import argparse
import calendar
import datetime as dt
import hashlib
import json
import logging
import os
import re
import sqlite3
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

import openpyxl

log = logging.getLogger("loxley")

DEFAULT_DB = "sqlite:///loxley_sales.db"
SOURCE = "loxley"

MONTHS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}
PART_RE = re.compile(r"^\d{9,15}$")
SIZE_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(?:ก\.|กรัม|g\b)")


# --------------------------------------------------------------------------
# Schema (portable between SQLite >= 3.25 and PostgreSQL)
# --------------------------------------------------------------------------
SCHEMA = [
    """
    CREATE TABLE IF NOT EXISTS ingest_log (
        file_sha256      TEXT PRIMARY KEY,
        file_name        TEXT NOT NULL,
        report_date      DATE,
        email_message_id TEXT,
        email_subject    TEXT,
        sheets           TEXT,
        rows_written     INTEGER,
        ingested_at      TIMESTAMP NOT NULL
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS dim_product (
        part_code     TEXT PRIMARY KEY,
        description   TEXT,
        size_g        DOUBLE PRECISION,
        price_point   TEXT,
        sales_channel TEXT,
        retailers     TEXT,
        first_seen    DATE,
        last_seen     DATE,
        updated_at    TIMESTAMP NOT NULL
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS dim_customer (
        customer_code TEXT NOT NULL,
        customer_name TEXT NOT NULL,
        channel_group TEXT,
        updated_at    TIMESTAMP NOT NULL,
        PRIMARY KEY (customer_code, customer_name)
    )
    """,
    # Raw cumulative MTD sales-out exactly as reported on each report day.
    """
    CREATE TABLE IF NOT EXISTS fact_sales_mtd (
        part_code     TEXT NOT NULL,
        snapshot_date DATE NOT NULL,
        month_start   DATE NOT NULL,
        mtd_qty       DOUBLE PRECISION NOT NULL,
        source_file   TEXT,
        source_sheet  TEXT,
        loaded_at     TIMESTAMP NOT NULL,
        PRIMARY KEY (part_code, snapshot_date)
    )
    """,
    # Loxley available stock per SKU per day.
    """
    CREATE TABLE IF NOT EXISTS fact_stock (
        part_code     TEXT NOT NULL,
        stock_date    DATE NOT NULL,
        available_qty DOUBLE PRECISION NOT NULL,
        source_file   TEXT,
        source_sheet  TEXT,
        loaded_at     TIMESTAMP NOT NULL,
        PRIMARY KEY (part_code, stock_date)
    )
    """,
    # PO quantities that arrived at Loxley per day.
    """
    CREATE TABLE IF NOT EXISTS fact_po_arrival (
        part_code     TEXT NOT NULL,
        arrival_date  DATE NOT NULL,
        qty           DOUBLE PRECISION NOT NULL,
        source_file   TEXT,
        source_sheet  TEXT,
        loaded_at     TIMESTAMP NOT NULL,
        PRIMARY KEY (part_code, arrival_date)
    )
    """,
    # Everything else on the SKU row, as of each report (history of plan,
    # forecast, PO, cover days, promotions ...).
    """
    CREATE TABLE IF NOT EXISTS fact_sku_status (
        part_code          TEXT NOT NULL,
        report_date        DATE NOT NULL,
        month_start        DATE NOT NULL,
        description        TEXT,
        ltp                TEXT,
        avg_sale_per_month DOUBLE PRECISION,
        avg_sale_label     TEXT,
        new_weight         TEXT,
        forecast_qty       DOUBLE PRECISION,
        mtd_actual_qty     DOUBLE PRECISION,
        pct_of_forecast    DOUBLE PRECISION,
        plan_po_qty        DOUBLE PRECISION,
        add_po_qty         DOUBLE PRECISION,
        pending_po_qty     DOUBLE PRECISION,
        total_po_qty       DOUBLE PRECISION,
        po_arrived_qty     DOUBLE PRECISION,
        po_outstanding_qty DOUBLE PRECISION,
        available_stock    DOUBLE PRECISION,
        cover_days         DOUBLE PRECISION,
        demand             DOUBLE PRECISION,
        promotion          TEXT,
        incoming_note      TEXT,
        extra_json         TEXT,
        source_file        TEXT,
        source_sheet       TEXT,
        loaded_at          TIMESTAMP NOT NULL,
        PRIMARY KEY (part_code, report_date)
    )
    """,
    "CREATE INDEX IF NOT EXISTS ix_sales_mtd_month ON fact_sales_mtd (month_start, part_code)",
    "CREATE INDEX IF NOT EXISTS ix_stock_date ON fact_stock (stock_date)",
    "CREATE INDEX IF NOT EXISTS ix_status_month ON fact_sku_status (month_start)",
    # Daily (per report interval) sales-out derived from the MTD snapshots.
    # days_in_interval > 1 means the qty covers several calendar days
    # (weekends / holidays without a report).
    "DROP VIEW IF EXISTS v_sales_out_daily",
    """
    CREATE VIEW v_sales_out_daily AS
    SELECT
        s.part_code,
        p.description,
        p.size_g,
        s.month_start,
        s.snapshot_date AS sales_date,
        s.prev_snapshot_date,
        s.mtd_qty,
        s.mtd_qty - COALESCE(s.prev_mtd_qty, 0) AS sales_qty
    FROM (
        SELECT
            part_code, month_start, snapshot_date, mtd_qty,
            LAG(snapshot_date) OVER w AS prev_snapshot_date,
            LAG(mtd_qty)       OVER w AS prev_mtd_qty
        FROM fact_sales_mtd
        WINDOW w AS (PARTITION BY part_code, month_start ORDER BY snapshot_date)
    ) s
    LEFT JOIN dim_product p ON p.part_code = s.part_code
    """,
    "DROP VIEW IF EXISTS v_sales_out_monthly",
    """
    CREATE VIEW v_sales_out_monthly AS
    SELECT
        m.part_code,
        p.description,
        p.size_g,
        m.month_start,
        m.snapshot_date AS last_report_date,
        m.mtd_qty       AS sales_qty
    FROM fact_sales_mtd m
    JOIN (
        SELECT part_code, month_start, MAX(snapshot_date) AS snapshot_date
        FROM fact_sales_mtd GROUP BY part_code, month_start
    ) last ON last.part_code = m.part_code
          AND last.snapshot_date = m.snapshot_date
    LEFT JOIN dim_product p ON p.part_code = m.part_code
    """,
]


# --------------------------------------------------------------------------
# Small DB adapter: sqlite3 by default, psycopg (v3) or psycopg2 for Postgres
# --------------------------------------------------------------------------
class DB:
    def __init__(self, url: str):
        self.url = url
        if url.startswith("sqlite:///"):
            path = url[len("sqlite:///"):]
            self.conn = sqlite3.connect(path)
            self.conn.execute("PRAGMA foreign_keys = ON")
            self.paramstyle = "?"
            self.kind = "sqlite"
        elif url.startswith(("postgres://", "postgresql://")):
            try:
                import psycopg  # type: ignore
                self.conn = psycopg.connect(url)
            except ImportError:
                import psycopg2  # type: ignore
                self.conn = psycopg2.connect(url)
            self.paramstyle = "%s"
            self.kind = "postgres"
        else:
            raise SystemExit(f"Unsupported database URL: {url!r}")

    def execute(self, sql: str, params: Iterable[Any] = ()) -> Any:
        cur = self.conn.cursor()
        cur.execute(sql.replace("?", self.paramstyle), tuple(params))
        return cur

    def executemany(self, sql: str, rows: list[tuple]) -> None:
        if rows:
            self.conn.cursor().executemany(sql.replace("?", self.paramstyle), rows)

    def init_schema(self) -> None:
        for stmt in SCHEMA:
            self.execute(stmt)
        self.conn.commit()

    def upsert(self, table: str, keys: list[str], rows: list[dict]) -> int:
        if not rows:
            return 0
        cols = list(rows[0].keys())
        updates = [c for c in cols if c not in keys]
        sql = (
            f"INSERT INTO {table} ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))}) "
            f"ON CONFLICT ({', '.join(keys)}) DO "
            + (f"UPDATE SET {', '.join(f'{c} = excluded.{c}' for c in updates)}" if updates else "NOTHING")
        )
        self.executemany(sql, [tuple(r[c] for c in cols) for r in rows])
        return len(rows)


# --------------------------------------------------------------------------
# Parsing helpers
# --------------------------------------------------------------------------
def clean_text(v: Any) -> str | None:
    if v is None:
        return None
    s = str(v).replace("\xa0", " ").strip()
    s = re.sub(r"\s+", " ", s)
    return s or None


def to_num(v: Any) -> float | None:
    """Numbers stay numbers; Excel errors ('#DIV/0!') and text become None."""
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.strip().replace(",", "")
        try:
            return float(s)
        except ValueError:
            return None
    return None


def sheet_month(name: str) -> tuple[int | None, int]:
    """'Oct.26' -> (2026, 10); 'June25' -> (2025, 6); 'Aug' -> (None, 8)."""
    m = re.match(r"\s*([A-Za-z]+)\.?\s*'?(\d{2,4})?\s*$", name)
    if not m or m.group(1)[:3].lower() not in MONTHS:
        raise ValueError(f"not a month sheet: {name!r}")
    month = MONTHS[m.group(1)[:3].lower()]
    year = m.group(2)
    if year is None:
        return None, month
    y = int(year)
    return (y + 2000 if y < 100 else y), month


def parse_header_date(label: Any, year: int, month: int) -> dt.date | None:
    """Parse the many date-label spellings Loxley uses in row 5:
    '1 Sep.', ' 7 Sep.', "1 Sep.'26", '31 May.25', '1Jun', '31.Dec', '01.Feb',
    '1 June 25', '22Mar.', 'Acc. 31 Mar24', '30Sep', 'update 1-29' ..."""
    if label is None:
        return None
    if isinstance(label, dt.datetime):
        return label.date()
    if isinstance(label, dt.date):
        return label
    s = str(label).strip()
    m = re.search(r"(\d{1,2})\s*[.\-/ ]?\s*([A-Za-z]{3,})\.?\s*'?\s*(\d{2,4})?", s)
    if m and m.group(2)[:3].lower() in MONTHS:
        day, mon = int(m.group(1)), MONTHS[m.group(2)[:3].lower()]
        if m.group(3):
            y = int(m.group(3))
            y = y + 2000 if y < 100 else y
        else:
            # month label without a year: neighbour of the sheet month
            y = year
            if month == 1 and mon == 12:
                y -= 1
            elif month == 12 and mon == 1:
                y += 1
    else:
        nums = re.findall(r"\d{1,2}", s)
        if not nums:
            return None
        day, mon, y = int(nums[-1]), month, year
    try:
        return dt.date(y, mon, day)
    except ValueError:
        return None


@dataclass
class Block:
    title: str
    cols: list[int] = field(default_factory=list)


@dataclass
class SheetResult:
    sheet: str
    year: int
    month: int
    report_date: dt.date | None = None
    products: list[dict] = field(default_factory=list)
    sales_mtd: list[dict] = field(default_factory=list)
    stock: list[dict] = field(default_factory=list)
    po_arrival: list[dict] = field(default_factory=list)
    status: list[dict] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


# Map of row-4 header text (normalised, prefix match) -> status field.
SINGLE_COLS = [
    ("ltp", "ltp"),
    ("avg sale", "avg_sale_per_month"),
    ("new weight", "new_weight"),
    ("forecast", "forecast_qty"),
    ("acc.actual", "acc_actual"),
    ("%", "pct_of_forecast"),
    ("plan po", "plan_po_qty"),
    ("po in", "plan_po_qty"),
    ("add po", "add_po_qty"),
    ("po add", "add_po_qty"),
    ("pending po", "pending_po_qty"),
    ("total po", "total_po_qty"),
    ("po ขาดส่ง", "po_outstanding_qty"),
    ("cover day", "cover_days"),
    ("demand", "demand"),
    ("promotion", "promotion"),
    ("ของเข้า", "incoming_note"),
]
TEXT_FIELDS = {"ltp", "new_weight", "promotion", "incoming_note"}


def find_header_row(ws) -> int | None:
    for r in range(1, 9):
        for c in range(1, 4):
            if clean_text(ws.cell(r, c).value) == "Part":
                return r
    return None


def parse_month_sheet(ws, year: int, month: int, source_file: str, loaded_at: str) -> SheetResult:
    res = SheetResult(sheet=ws.title, year=year, month=month)
    hr = find_header_row(ws)
    if hr is None:
        res.warnings.append("no 'Part' header row found")
        return res
    dr = hr + 1  # date-label row
    max_col = ws.max_column
    month_start = dt.date(year, month, 1)

    # Walk the header row and group columns into blocks: a titled column starts a
    # block, untitled columns continue the previous one (Available Stock in old
    # sheets is titled only on its first column).
    blocks: list[Block] = []
    part_col = desc_col = None
    for c in range(1, max_col + 1):
        title = clean_text(ws.cell(hr, c).value)
        sub = ws.cell(dr, c).value
        if title:
            t = title.lower()
            if t == "part":
                part_col = c
            elif t == "description":
                desc_col = c
            blocks.append(Block(title, [c]))
        elif blocks and sub not in (None, ""):
            blocks[-1].cols.append(c)
    if part_col is None or desc_col is None:
        res.warnings.append("Part/Description columns not found")
        return res

    actual_cols: dict[dt.date, int] = {}
    stock_cols: dict[dt.date, int] = {}
    po_day_cols: dict[dt.date, int] = {}
    po_total_col = None
    single: dict[str, int] = {}
    extra_cols: dict[str, int] = {}

    for b in blocks:
        t = b.title.lower()
        if t in ("part", "description"):
            continue
        if t.startswith("actual"):
            for c in b.cols:
                d = parse_header_date(ws.cell(dr, c).value, year, month)
                if d is None:
                    continue
                if (d.year, d.month) != (year, month):
                    # e.g. '31 May.25' at the start of the June sheet = prior
                    # month's closing total; that month's own sheet owns it.
                    continue
                actual_cols[d] = c  # duplicate labels: rightmost wins
        elif t.startswith("available stock"):
            for c in b.cols:
                d = parse_header_date(ws.cell(dr, c).value, year, month)
                if d is not None:
                    stock_cols[d] = c
        elif t.startswith("po arrival"):
            for c in b.cols:
                sub = ws.cell(dr, c).value
                if clean_text(sub) and str(sub).strip().lower() == "total":
                    po_total_col = c
                    continue
                day = to_num(sub)
                if day is not None and 1 <= day <= calendar.monthrange(year, month)[1]:
                    po_day_cols[dt.date(year, month, int(day))] = c
        else:
            matched = False
            for prefix, fld in SINGLE_COLS:
                if t.startswith(prefix):
                    if fld not in single:
                        single[fld] = b.cols[0]
                        if fld == "avg_sale_per_month":
                            single["_avg_label"] = b.cols[0]
                    matched = True
                    break
            if not matched:
                extra_cols[b.title] = b.cols[0]
    # Fallback: the PO-arrival 'Total' column can sit outside the titled block.
    if po_total_col is None:
        for c in range(1, max_col + 1):
            if clean_text(ws.cell(dr, c).value) == "Total" and not clean_text(ws.cell(hr, c).value):
                po_total_col = c
                break

    # Old Feb–Jun 2024 sheets have a single 'Acc.Actual' month total column.
    if not actual_cols and "acc_actual" in single:
        c = single["acc_actual"]
        d = parse_header_date(ws.cell(dr, c).value, year, month)
        if d and (d.year, d.month) == (year, month):
            actual_cols[d] = c

    # SKU rows
    sku_rows: list[tuple[int, str, str | None]] = []
    for r in range(dr + 1, ws.max_row + 1):
        code = clean_text(ws.cell(r, part_col).value)
        if code and PART_RE.match(code):
            sku_rows.append((r, code, clean_text(ws.cell(r, desc_col).value)))
        elif clean_text(ws.cell(r, desc_col).value) in ("Total", "Stock LTC"):
            break
    if not sku_rows:
        res.warnings.append("no SKU rows")
        return res

    def column_reported(c: int) -> bool:
        return any(to_num(ws.cell(r, c).value) is not None for r, _, _ in sku_rows)

    # Only keep snapshot dates that were actually reported (any numeric value);
    # within a reported column a blank cell means 0 sold so far this month.
    actual_cols = {d: c for d, c in actual_cols.items() if column_reported(c)}
    stock_cols = {d: c for d, c in stock_cols.items() if column_reported(c)}
    res.report_date = max(actual_cols) if actual_cols else (max(stock_cols) if stock_cols else None)
    if not actual_cols:
        res.warnings.append("no reported 'Actual' (sales MTD) columns")

    for r, code, desc in sku_rows:
        res.products.append({"part_code": code, "description": desc})
        for d, c in sorted(actual_cols.items()):
            res.sales_mtd.append({
                "part_code": code, "snapshot_date": d.isoformat(),
                "month_start": month_start.isoformat(),
                "mtd_qty": to_num(ws.cell(r, c).value) or 0.0,
                "source_file": source_file, "source_sheet": ws.title, "loaded_at": loaded_at,
            })
        for d, c in sorted(stock_cols.items()):
            v = to_num(ws.cell(r, c).value)
            if v is not None:
                res.stock.append({
                    "part_code": code, "stock_date": d.isoformat(), "available_qty": v,
                    "source_file": source_file, "source_sheet": ws.title, "loaded_at": loaded_at,
                })
        for d, c in sorted(po_day_cols.items()):
            v = to_num(ws.cell(r, c).value)
            if v:
                res.po_arrival.append({
                    "part_code": code, "arrival_date": d.isoformat(), "qty": v,
                    "source_file": source_file, "source_sheet": ws.title, "loaded_at": loaded_at,
                })

        if res.report_date is None:
            continue
        val = lambda f: ws.cell(r, single[f]).value if f in single else None  # noqa: E731
        num = lambda f: to_num(val(f))  # noqa: E731
        txt = lambda f: clean_text(val(f))  # noqa: E731
        latest_actual = actual_cols.get(max(actual_cols)) if actual_cols else None
        latest_stock = stock_cols.get(max(stock_cols)) if stock_cols else None
        extra = {k: clean_text(ws.cell(r, c).value) for k, c in extra_cols.items()
                 if clean_text(ws.cell(r, c).value) is not None}
        res.status.append({
            "part_code": code,
            "report_date": res.report_date.isoformat(),
            "month_start": month_start.isoformat(),
            "description": desc,
            "ltp": txt("ltp"),
            "avg_sale_per_month": num("avg_sale_per_month"),
            "avg_sale_label": clean_text(ws.cell(hr, single["_avg_label"]).value) if "_avg_label" in single else None,
            "new_weight": txt("new_weight"),
            "forecast_qty": num("forecast_qty"),
            "mtd_actual_qty": (to_num(ws.cell(r, latest_actual).value) or 0.0) if latest_actual else None,
            "pct_of_forecast": num("pct_of_forecast"),
            "plan_po_qty": num("plan_po_qty"),
            "add_po_qty": num("add_po_qty"),
            "pending_po_qty": num("pending_po_qty"),
            "total_po_qty": num("total_po_qty"),
            "po_arrived_qty": to_num(ws.cell(r, po_total_col).value) if po_total_col else None,
            "po_outstanding_qty": num("po_outstanding_qty"),
            "available_stock": to_num(ws.cell(r, latest_stock).value) if latest_stock else None,
            "cover_days": num("cover_days"),
            "demand": num("demand"),
            "promotion": txt("promotion"),
            "incoming_note": txt("incoming_note"),
            "extra_json": json.dumps(extra, ensure_ascii=False) if extra else None,
            "source_file": source_file, "source_sheet": ws.title, "loaded_at": loaded_at,
        })
    return res


def parse_reference_sheets(wb, loaded_at: str) -> tuple[list[dict], list[dict]]:
    """DATA_ไซส์ (SKU attributes) and DATA_Cocode (customer master)."""
    product_attrs: list[dict] = []
    customers: dict[tuple[str, str], dict] = {}
    for ws in wb.worksheets:
        title = ws.title.strip()
        if title.startswith("DATA_") and title != "DATA_Cocode":
            rows = ws.iter_rows(values_only=True)
            header = [clean_text(h) or "" for h in next(rows, ())]
            if "Part" not in header:
                continue
            ix = {h: i for i, h in enumerate(header)}
            for row in rows:
                code = clean_text(row[ix["Part"]]) if len(row) > ix["Part"] else None
                if not code or not PART_RE.match(code):
                    continue
                g = lambda i: clean_text(row[i]) if i is not None and i < len(row) else None  # noqa: E731
                product_attrs.append({
                    "part_code": code,
                    "description": g(ix.get("Description")),
                    "price_point": g(ix.get("ไซส์บาท")),
                    "sales_channel": g(ix.get("ช่องทางที่ขาย")),
                    "retailers": g(ix.get("ห้างที่ขาย")),
                })
        elif title == "DATA_Cocode":
            for row in ws.iter_rows(min_row=2, values_only=True):
                code, name = clean_text(row[0]), clean_text(row[1])
                if code and name:
                    customers[(code, name)] = {
                        "customer_code": code, "customer_name": name,
                        "channel_group": clean_text(row[2]) if len(row) > 2 else None,
                        "updated_at": loaded_at,
                    }
    return product_attrs, list(customers.values())


def month_sheets(wb) -> list[tuple[Any, int, int]]:
    """Month sheets in workbook order with (year, month) resolved. Sheets with no
    year in their name (2023's 'Aug'..'Dec') borrow it from the next dated sheet."""
    raw = []
    for ws in wb.worksheets:
        try:
            y, m = sheet_month(ws.title)
        except ValueError:
            continue
        raw.append([ws, y, m])
    for i, item in enumerate(raw):
        if item[1] is None:
            nxt = next((x for x in raw[i + 1:] if x[1] is not None), None)
            if nxt is None:
                continue
            item[1] = nxt[1] - 1 if item[2] > nxt[2] else nxt[1]
    return [(ws, y, m) for ws, y, m in raw if y is not None]


def report_date_from_name(name: str) -> dt.date | None:
    """'Green Nut Daily Stock 2026 22.10.26.xlsx' -> 2026-10-22."""
    m = re.search(r"(\d{1,2})\.(\d{1,2})\.(\d{2,4})(?!.*\d)", Path(name).stem)
    if not m:
        return None
    d, mo, y = (int(x) for x in m.groups())
    try:
        return dt.date(y + 2000 if y < 100 else y, mo, d)
    except ValueError:
        return None


def size_from_description(desc: str | None) -> float | None:
    if not desc:
        return None
    m = SIZE_RE.search(desc)
    return float(m.group(1)) if m else None


# --------------------------------------------------------------------------
# Ingest
# --------------------------------------------------------------------------
def ingest_file(db: DB, path: Path, all_sheets: bool, months_back: int,
                message_id: str | None, subject: str | None, force: bool) -> dict:
    data = path.read_bytes()
    sha = hashlib.sha256(data).hexdigest()
    if not force:
        seen = db.execute("SELECT ingested_at FROM ingest_log WHERE file_sha256 = ?", (sha,)).fetchone()
        if seen:
            log.info("%s already ingested at %s (use --force to reload)", path.name, seen[0])
            return {"file": path.name, "skipped": True}

    loaded_at = dt.datetime.now(dt.timezone.utc).replace(microsecond=0).isoformat()
    wb = openpyxl.load_workbook(path, data_only=True)  # cached values, not formulas
    sheets = month_sheets(wb)
    if not sheets:
        raise SystemExit(f"{path.name}: no month sheets found")

    name_date = report_date_from_name(path.name)
    sheets.sort(key=lambda s: (s[1], s[2]))
    if not all_sheets:
        # Current month (from file name, else newest sheet) and the N before it,
        # so a report on the 1st still finalises last month.
        anchor = (name_date.year, name_date.month) if name_date else (sheets[-1][1], sheets[-1][2])
        keep = []
        y, m = anchor
        for _ in range(months_back + 1):
            keep.append((y, m))
            y, m = (y - 1, 12) if m == 1 else (y, m - 1)
        sheets = [s for s in sheets if (s[1], s[2]) in keep]
        if not sheets:
            raise SystemExit(f"{path.name}: no sheet for {anchor}")

    totals = {"products": 0, "sales_mtd": 0, "stock": 0, "po_arrival": 0, "status": 0}
    products: dict[str, dict] = {}
    report_date: dt.date | None = None
    parsed_sheets = []
    for ws, y, m in sheets:  # chronological, so later sheets win on overlaps
        r = parse_month_sheet(ws, y, m, path.name, loaded_at)
        for w in r.warnings:
            log.warning("%s [%s]: %s", path.name, ws.title, w)
        if not r.sales_mtd and not r.stock:
            continue
        parsed_sheets.append(ws.title.strip())
        for p in r.products:
            d = r.report_date.isoformat() if r.report_date else None
            cur = products.setdefault(p["part_code"], {"part_code": p["part_code"], "description": p["description"],
                                                       "first_seen": d, "last_seen": d})
            cur["description"] = p["description"] or cur["description"]
            if d:
                cur["first_seen"] = min(filter(None, [cur["first_seen"], d]))
                cur["last_seen"] = max(filter(None, [cur["last_seen"], d]))
        totals["sales_mtd"] += db.upsert("fact_sales_mtd", ["part_code", "snapshot_date"], r.sales_mtd)
        totals["stock"] += db.upsert("fact_stock", ["part_code", "stock_date"], r.stock)
        totals["po_arrival"] += db.upsert("fact_po_arrival", ["part_code", "arrival_date"], r.po_arrival)
        totals["status"] += db.upsert("fact_sku_status", ["part_code", "report_date"], r.status)
        if r.report_date and (report_date is None or r.report_date > report_date):
            report_date = r.report_date
        log.info("%-8s %d-%02d report=%s skus=%d mtd_snapshots=%d stock=%d po=%d",
                 ws.title.strip(), y, m, r.report_date, len(r.products),
                 len({x['snapshot_date'] for x in r.sales_mtd}), len(r.stock), len(r.po_arrival))

    attrs, customers = parse_reference_sheets(wb, loaded_at)
    attr_by_code = {a["part_code"]: a for a in attrs}
    for code, a in attr_by_code.items():
        products.setdefault(code, {"part_code": code, "description": a["description"],
                                   "first_seen": None, "last_seen": None})
    prod_rows = []
    for code, p in products.items():
        a = attr_by_code.get(code, {})
        existing = db.execute("SELECT first_seen, last_seen FROM dim_product WHERE part_code = ?", (code,)).fetchone()
        # Postgres returns date objects, SQLite returns ISO strings
        old_first, old_last = (str(v)[:10] if v else None for v in (existing or (None, None)))
        first = min(filter(None, [p["first_seen"], old_first]), default=None)
        last = max(filter(None, [p["last_seen"], old_last]), default=None)
        prod_rows.append({
            "part_code": code,
            "description": p["description"] or a.get("description"),
            "size_g": size_from_description(p["description"] or a.get("description")),
            "price_point": a.get("price_point"),
            "sales_channel": a.get("sales_channel"),
            "retailers": a.get("retailers"),
            "first_seen": str(first) if first else None,
            "last_seen": str(last) if last else None,
            "updated_at": loaded_at,
        })
    totals["products"] = db.upsert("dim_product", ["part_code"], prod_rows)
    totals["customers"] = db.upsert("dim_customer", ["customer_code", "customer_name"], customers)

    if name_date and report_date and name_date != report_date:
        log.warning("file name says %s but newest sales column is %s", name_date, report_date)
    report_date = report_date or name_date
    db.upsert("ingest_log", ["file_sha256"], [{
        "file_sha256": sha, "file_name": path.name,
        "report_date": report_date.isoformat() if report_date else None,
        "email_message_id": message_id, "email_subject": subject,
        "sheets": ",".join(parsed_sheets), "rows_written": sum(totals.values()),
        "ingested_at": loaded_at,
    }])
    db.conn.commit()
    return {"file": path.name, "report_date": str(report_date), "sheets": parsed_sheets, **totals}


def summary(db: DB, month: str | None) -> None:
    if month is None:
        row = db.execute("SELECT MAX(month_start) FROM fact_sales_mtd").fetchone()
        month = row[0] if row else None
    if month is None:
        print("database is empty")
        return
    month = str(month)[:10]
    print(f"Sales-out by SKU, month starting {month}")
    rows = db.execute(
        "SELECT part_code, description, last_report_date, sales_qty FROM v_sales_out_monthly "
        "WHERE month_start = ? ORDER BY sales_qty DESC, part_code", (month,)).fetchall()
    total = 0.0
    for code, desc, last, qty in rows:
        total += qty or 0
        print(f"  {code}  {(desc or '')[:45]:45s}  {qty:10.2f}  (as of {last})")
    print(f"  {'TOTAL':>60s}  {total:10.2f}")
    print("\nMonthly totals:")
    for m, q in db.execute(
            "SELECT month_start, SUM(sales_qty) FROM v_sales_out_monthly GROUP BY month_start ORDER BY month_start").fetchall():
        print(f"  {str(m)[:7]}  {q:12.2f}")


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--db", default=os.environ.get("LOXLEY_DB_URL", DEFAULT_DB),
                    help=f"database URL (default: $LOXLEY_DB_URL or {DEFAULT_DB})")
    ap.add_argument("-v", "--verbose", action="store_true")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p_in = sub.add_parser("ingest", help="load one or more Loxley workbooks")
    p_in.add_argument("files", nargs="+", type=Path)
    p_in.add_argument("--all-sheets", action="store_true",
                      help="load every month sheet (historical backfill)")
    p_in.add_argument("--months-back", type=int, default=1,
                      help="previous month sheets to reload besides the current one (default 1)")
    p_in.add_argument("--message-id", help="source Gmail message id, stored in ingest_log")
    p_in.add_argument("--subject", help="source email subject, stored in ingest_log")
    p_in.add_argument("--force", action="store_true", help="reload even if this exact file was ingested")

    p_sum = sub.add_parser("summary", help="print monthly sales-out from the database")
    p_sum.add_argument("--month", help="YYYY-MM-01 (default: latest)")

    args = ap.parse_args(argv)
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO,
                        format="%(levelname)s %(message)s")
    db = DB(args.db)
    db.init_schema()

    if args.cmd == "ingest":
        results = []
        for f in args.files:
            if not f.exists():
                log.error("file not found: %s", f)
                return 2
            results.append(ingest_file(db, f, args.all_sheets, args.months_back,
                                       args.message_id, args.subject, args.force))
        print(json.dumps(results, ensure_ascii=False, indent=2))
    elif args.cmd == "summary":
        summary(db, args.month)
    return 0


if __name__ == "__main__":
    sys.exit(main())
