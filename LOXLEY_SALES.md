# Loxley Green Nut sales-out parser

`loxley_sales_parser.py` turns Loxley's daily email attachment
**"Green Nut Daily Sales & Stock as of DD Mon YYYY"**
(`Green Nut Daily Stock 2026 22.10.26.xlsx`) into a database that grows each
day, so we stop depending on the Excel file.

## What the Excel file contains

The workbook holds **one sheet per month** (`Oct.26`, `Sep.26`, … back to `Aug`
2023; older ones are hidden). The current month on 22 Oct 2026 is `Oct.26`, with
report columns for 1, 2, 5, 6, 7, 8, 9, 12, 14, 15, 19, 20, 21 and 22 Oct (working
days only). Each row is one SKU (Loxley part code, e.g. `006010130301` =
กรีนนัท 70 ก. รสปลาหมึก, 32 active SKUs). The columns are:

| Block | Meaning |
|---|---|
| Part, Description, LTP | SKU |
| Avg Sale/mth, New Weight, Forecast | monthly reference / plan |
| **Actual ‹date›** … | **cumulative month-to-date sales-out**, one column per report day |
| % | MTD actual ÷ forecast |
| Plan PO, Add PO, Pending PO, Total PO | purchase orders for the month |
| PO Arrival 1…31, Total | PO quantity received at Loxley per day |
| PO ขาดส่ง | PO still outstanding |
| **Available Stock ‹date›** … | Loxley warehouse stock per day |
| Cover Day, Demand, Promotion, ของเข้าวันที่ | stock cover, demand, active retailer promos, incoming-goods note |

Hidden reference sheets: `DATA_ไซส์` (SKU price point and retailer, e.g.
Makro / CJ / 7-Eleven) and `DATA_Cocode` (666 Loxley customers mapped to
channel groups: 7-11, Makro, Lotus, CJ, BigC, Tops, TT_North, …).

Things to know:
* **"Actual" is cumulative.** Daily sales = MTD today − MTD on the previous
  report day. The view `v_sales_out_daily` does this calculation.
* **Quantities are given per SKU only.** The file has no per-retailer
  breakdown; the promotion text is the only retailer-level detail. Units are as
  reported (cartons/ลัง, with fractions for loose packs).
* Column positions move every month, so the parser finds each block by its
  header text, not by fixed column letters.
* Loxley sometimes corrects earlier numbers, so a day can show negative sales.
  For example, 9 Jan 2025 was typed as 1,272.9 and the next day showed 2. These
  rows are stored as reported. A spike and its correction cancel out, so monthly
  totals stay correct.

## Database

SQLite by default (`loxley_sales.db`). PostgreSQL (for example Supabase) works
with `--db postgresql://…` or the `LOXLEY_DB_URL` environment variable.

| Table / view | Grain | Notes |
|---|---|---|
| `fact_sales_mtd` | SKU × report date | raw cumulative MTD qty, as reported |
| `v_sales_out_daily` | SKU × report date | **sales-out for the interval since the previous report** |
| `v_sales_out_monthly` | SKU × month | month total (last MTD value) |
| `fact_stock` | SKU × date | Loxley available stock |
| `fact_po_arrival` | SKU × date | PO qty received |
| `fact_sku_status` | SKU × report date | forecast, PO plan, outstanding, cover days, demand, promotion text, notes; one row per email, so the history is kept |
| `dim_product` | SKU | description, size (g), price point, retailers, first/last seen |
| `dim_customer` | Loxley customer | channel group (from `DATA_Cocode`) |
| `change_log` | revision | every update or delete of a value already loaded (old → new, which file) |
| `ingest_log` | file | SHA-256, report date, Gmail message id, subject, data-quality warnings |

### How loading works

Loads are incremental. Loxley only ever changes the current month sheet, by
inserting one "Actual" and one "Available Stock" column per day (see
`LOXLEY_DIFF_PLAN.md` for the analysis). A daily run:

* streams only the current month sheet. During the first 7 days of a month
  it also reads the previous month, because the 1 Oct file still finalised
  `Sep.26`. The customer master (`DATA_Cocode`, 8 MB) is read only on the
  first file of each month or with `--refresh-reference`;
* takes the report date from the newest sales column, not the file name
  (September files had no date in the name);
* diffs the month against the database: inserts new rows, updates changed
  ones, deletes rows Loxley removed (e.g. a PO moved to another SKU), and
  logs every update or delete in `change_log`;
* skips a file or Gmail message id that was already loaded, and any file
  older than the newest loaded report (`--force` overrides). Even a forced
  older file never deletes data dated after its own report date;
* refuses to delete more than 20% of a month's rows (a sign of a broken
  file) unless `--allow-mass-delete` is given;
* records warnings in `ingest_log.warnings`:
  * a repeated or out-of-order date label, fixed from the row-3 "time gone"
    formula where possible, otherwise the column is skipped;
  * SKU rows not adding up to the Total row;
  * a sharp MTD drop;
  * the file name or subject date not matching the data.

Measured on the real 19–22 Oct files: **0.6–0.8 s and about 95 row writes per
day**, versus 4.5 s and about 2,300 upserts before. Loading the files day by
day gives exactly the same tables as one full load of the latest file.

## Usage

```bash
pip install -r requirements.txt

# one-off: backfill all history (Sep 2023 → today) from any recent workbook
python loxley_sales_parser.py ingest "Green Nut Daily Stock 2026 22.10.26.xlsx" --all-sheets

# daily: incremental load of one email attachment
python loxley_sales_parser.py ingest "Green Nut Daily Stock 2026 23.10.26.xlsx" \
    --message-id 1a1615e08f1d26fa --subject "Green Nut Daily Sales & Stock as of 23 Oct 2026"

python loxley_sales_parser.py summary            # latest month by SKU + monthly totals
python -m pytest -q tests                         # tests (synthetic workbook)
```

Example analysis query:

```sql
SELECT p.size_g, substr(d.month_start,1,7) AS month, SUM(d.sales_qty) AS qty
FROM v_sales_out_daily d JOIN dim_product p USING (part_code)
GROUP BY 1, 2 ORDER BY 2, 1;
```

## Automation (next phase)

A scheduled Claude task would do the following on each run:
1. Search Gmail (Composio, account `witchawat.bb@gmail.com`) for
   `from:loxtrade.com subject:"Green Nut Daily" has:attachment newer_than:3d`.
   The subject varies ("Sales & Stock" / "Sales&Stock") and the attachment
   name sometimes has no date, so don't match on either.
2. Skip message ids already in `ingest_log.email_message_id`. The script
   also skips them itself.
3. Download each `.xlsx` attachment, **oldest email first**, and run
   `python loxley_sales_parser.py ingest <file> --message-id <id> --subject "<subject>"`.
4. Report anything in the output's `warnings` list.

Cloud sessions are ephemeral, so a scheduled run must write to a persistent
PostgreSQL database (set `LOXLEY_DB_URL`). A SQLite file inside the container
would be lost after each run.
