# Data pipeline analysis (2026-02-25)

## Current issue summary

Dashboard totals are incorrect because the Google Sheet now contains **many rows per same date** from rolling 30-day imports, and those rows often have different values than manual/daily entries.

Example anomalies from provided rows:

- `2026-02-21`: manual `14157`, rolling import `27715` (major inflation)
- `2026-02-22`: manual `15262`, rolling import `30409` (major inflation)
- `2026-02-24`: rolling import `25846` while expected is `12993`

So the problem is not only duplicate rows: rolling files can contain **different (inflated) values** for same date.

## Is rolling 30-day file useful?

Short answer: **not as a primary source** in the current pipeline.

- It currently increases complexity and creates conflicting rows.
- It does not appear to consistently correct missing daily values.
- In this dataset it introduces large overcounts on recent days.

Recommendation:

1. Use one primary source for dashboard totals (daily/manual validated rows).
2. Keep rolling file only as a secondary audit signal (separate sheet), not merged directly into production table.

## Why dashboard can show 25 846 for 24.2

Likely root cause:

- Latest rolling import wrote `2026-02-24 = 25846`.
- Backend payload used that value (either latest row per date or aggregation polluted by duplicates).

Expected for 24.2 is `12993`, so production table selection logic is currently trusting wrong source data.

## Minimal-risk stabilization plan

### Phase 1 (immediate)

- Stop writing rolling 30-day rows into production sheet used by endpoint.
- Keep only one row per date in production table.
- For 12.2–24.2, overwrite production values with your verified truth list.

### Phase 2 (safe pipeline)

- Use separate tabs:
  - `steps_raw_daily` (raw ingest)
  - `steps_raw_rolling30` (raw ingest)
  - `steps_fact` (one row/date for dashboard)
- Endpoint reads only `steps_fact`.

### Phase 3 (optional reconciliation)

- Reconciliation rule example:
  - prefer `manual` if exists
  - else prefer `daily` snapshot
  - use rolling only if date is missing in both
  - never auto-sum multiple rows for same date

## Quick sheet cleanup guidance

For each date, keep exactly one final value and archive extra rows.

Priority order for final value:

1. `manual`
2. trusted daily export
3. rolling export

Then enforce uniqueness by date for the production tab.
