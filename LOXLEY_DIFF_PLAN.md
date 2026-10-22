# Plan: incremental ("diff") daily loader

Goal: tomorrow's email should load only what changed. Today the script
re-reads the whole 2 MB workbook and upserts about 2,300 rows for each email.

## 1. How Loxley updates the workbook (observed 30 Sep, 1 Oct, 15, 19, 20, 21, 22 Oct 2026)

Cell-by-cell comparison of consecutive files:

* **Only the current month sheet changes.** Between 15→19→20→21→22 Oct every
  other sheet is value-identical, including `Sep.26` and all hidden sheets.
* **Each day Loxley inserts two columns into the current sheet**: a new
  `Actual ‹today›` column after the last Actual column, and a new
  `Available Stock ‹today›` column after the last stock column. Everything to
  the right shifts by 1–2 letters (sheet width BQ → BS → BU → BW → BY).
  That is why a positional diff shows 800–900 "changed" cells per day.
  Compared by meaning (SKU + block + date), each day adds only:
  * 32 new MTD-sales values (one per SKU) and about 29 new stock values;
  * a few PO arrivals, sometimes **back-dated** (the 19 Oct file added
    arrivals dated 12 and 16 Oct);
  * new values in the "current state" columns (Demand, Cover Day, PO totals,
    % of forecast). These are recalculated daily, and Demand changes every day.
* **Past MTD and stock columns are frozen once written.** No historical value
  changed between 15 and 22 Oct.
* **But Loxley does make corrections, and the next file repairs them:**
  * The 30 Sep file had two columns both labelled `28 Sep.`. The second
    actually held 29 Sep data, and the label was fixed in the 1 Oct file.
  * The 19 Oct file moved a 350-unit PO (12 Oct) from `006010140201`
    (17 g squid) to `006010140203` (17 g squid Shelf Ready).
* **Month rollover**: the 1 Oct file is the last one that changed `Sep.26`.
  It added 29 and 30 Sep and fixed the label. `Oct.26` appeared the same day,
  and `Sep.26` was untouched after that.
* **File name is not reliable for the date.** September files were named
  `Green Nut Daily Stock 2026.xlsx` with no date.
* **The subject line varies**: "Sales & Stock" and "Sales&Stock", and the
  sender changed from Arthita Takeoka to Nattha Puiaroon.
* Raw sheet XML inside the xlsx changes for **every** sheet every day,
  because Excel rewrites the shared strings. Hashing the zip parts therefore
  cannot detect unchanged sheets.

## 2. What the current script gets right and wrong (verified)

Loaded 15 Oct (full backfill), then 19, 20, 21, 22 Oct one per day. Compared
the result with a fresh load of the 22 Oct file alone:

| table | result |
|---|---|
| fact_sales_mtd, fact_stock, dim_product, dim_customer | **identical** |
| fact_sku_status | sequential run keeps 5 daily snapshots (the history we want); single load keeps 1 |
| fact_po_arrival | **1 stale row**: `006010140201 2026-10-12 350`, the PO Loxley moved. Upserts never delete. |

Also confirmed: with the duplicate `28 Sep.` label, the parser stores the
29 Sep value under 28 Sep until the next file overwrites it.

Timing per daily file (current script): about 4.5 s, of which:
* `load_workbook` (all 48 sheets): 3.7 s
* read-only stream of `Oct.26` alone: 0.27 s
* `DATA_Cocode` alone (7.9 MB XML, 58k rows): 1.1 s

## 3. Design of the diff loader

1. **Open only what is needed.** Use `openpyxl` read-only mode and stream just:
   * the current month sheet;
   * the previous month sheet, only while it can still change: the report date
     is within the first 5 days of the month, **or** the DB's last snapshot
     for that month is earlier than the month's last working day;
   * `DATA_ไซส์` (33 rows, cheap).
   * Skip `DATA_Cocode` unless `--refresh-reference` is set, or it's the first
     run of a month.

   Read each sheet once into a row list (`iter_rows(values_only=True)`).
   Random `ws.cell()` access is very slow in read-only mode. So
   `parse_month_sheet` will take a grid instead of a worksheet; the same
   parser serves the backfill and the diff paths.
2. **Report date comes from the sheet content**: the newest reported
   `Actual` column. The email subject and file name are cross-checks only;
   log a warning on mismatch.
3. **Diff against the DB, scoped to the sheet's month.** For each fact table,
   load the existing rows for that month (one query), compare by key, then:
   * insert new keys (normally 32 MTD + about 29 stock + a few PO);
   * update changed values;
   * **delete keys that disappeared** (fixes the moved-PO bug). Only within
     the month the sheet owns. Carry-over columns, such as `31 Aug` stock in
     `Sep.26`, never delete rows owned by another month.
   * Everything runs in one transaction.
4. **`change_log` table** (new): for every update or delete of a past value,
   store table, key, old value, new value, report date, file, message id.
   This gives an audit trail of Loxley's corrections, for example
   `mtd 006010140101 2026-09-28 81.29 → 65.29` and the moved PO. Brand-new
   keys are not logged.
5. **Status rows**: insert one `fact_sku_status` row per SKU for the report
   date (32 rows). This is already the history we want.
6. **Data-quality guards**, as warnings in `ingest_log` that don't stop the run:
   * duplicate date labels in a block, resolved by position (each column
     must be later than the one to its left) and checked against the row-3
     "time gone" formula (`=22/31`);
   * MTD for a SKU drops by more than X versus the previous report (likely
     a typo, like 9 Jan 2025's 1,272.9);
   * sum of SKU rows ≠ the sheet's `Total` row for the newest columns;
   * the subject or file date doesn't match the content's report date.
7. **Dedup by Gmail message id as well as file SHA.** Late or duplicate
   emails are a no-op, and an older file never overwrites a newer report
   (compare report dates before writing).
8. **Backfill stays as is** (`--all-sheets`) and goes through the same diff
   path, so one code path serves both.

Expected daily cost: under 1 s, and about 100 row writes instead of about
2,300 upserts. Without `--refresh-reference`, the xlsx bytes for
`DATA_Cocode` are still decompressed but never parsed.

## 4. Scheduler flow (unchanged apart from the query)

```
Gmail query: from:loxtrade.com subject:"Green Nut Daily" has:attachment newer_than:3d
  -> skip message ids already in ingest_log
  -> download .xlsx (process oldest first)
  -> python loxley_sales_parser.py ingest <file> --message-id <id> --subject "<subject>"
```

## 5. Tests to add

Synthetic workbooks reproducing each observed case:
* two inserted columns (shifted layout);
* a back-dated PO arrival;
* a PO moved between SKUs (the stale row must be deleted and logged);
* a duplicate date label that is fixed the next day;
* month rollover (the previous sheet is finalised on the 1st, then ignored);
* the same file twice, and an older file after a newer one (no-op).

Plus the regression check used here: sequential daily loads must equal a
single full load of the last file, apart from status history.
