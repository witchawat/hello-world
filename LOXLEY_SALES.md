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
| `ingest_log` | file | SHA-256, report date, Gmail message id, subject |

Loads are idempotent. Every table is upserted on its natural key, the same
file is never loaded twice (use `--force` to reload it), and the newest report
wins when values overlap.

## Usage

```bash
pip install -r requirements.txt

# one-off: backfill all history (Sep 2023 → today) from any recent workbook
python loxley_sales_parser.py ingest "Green Nut Daily Stock 2026 22.10.26.xlsx" --all-sheets

# daily: reloads the current month + previous month (so the 1st-of-month email
# still finalises last month)
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
   `subject:"Green Nut Daily Sales & Stock" has:attachment newer_than:2d`.
2. Skip message ids already in `ingest_log.email_message_id`.
3. Download the `.xlsx` attachment and run
   `python loxley_sales_parser.py ingest <file> --message-id <id> --subject "<subject>"`.

Cloud sessions are ephemeral, so a scheduled run must write to a persistent
PostgreSQL database (set `LOXLEY_DB_URL`). A SQLite file inside the container
would be lost after each run.
