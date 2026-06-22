# Fuel Station — MASTER Sync

A small Streamlit app that pulls the **operational** columns from a source
`MASTER` sheet and merges them with the **financial** columns you maintain in
your own workbook — matched on `(DATE, STATION)`.

## What gets synced

| Side | Columns | Source of truth |
|------|---------|-----------------|
| Operational | everything **except** the 7 below | the **source** MASTER (Google Sheet or uploaded .xlsx) |
| Financial | `BANKED`, `Banking Date`, `Bank`, `Amount Deposited`, `Balance Left`, `PMS Cost`, `AGO Cost` | your **target** workbook (preserved) |

Rows are matched on `DATE` + `STATION`. A row that exists in the source but not
the target is added with the financial columns left blank. Financial values you
already entered in the target are never overwritten.

## Setup

```bash
pip install -r requirements.txt
cp .env.example .env        # then edit .env with your links
streamlit run app.py
```

### `.env`

The Google links and layout config live in `.env` (never commit the real one —
it's in `.gitignore`):

```
SOURCE_SHEET_URL=https://docs.google.com/spreadsheets/d/<source-id>/edit   # operational, read-only
TARGET_SHEET_URL=https://docs.google.com/spreadsheets/d/<target-id>/edit   # the file we populate into
SHEET_NAME=MASTER
HEADER_ROW=2
GOOGLE_SERVICE_ACCOUNT_FILE=                 # required only for private read / write-back
```

## How to use

1. **Source** — defaults to `SOURCE_SHEET_URL` from your `.env`. (You can still
   switch to "Upload an .xlsx" in the sidebar.)
   - For a public link, share it as **Anyone with the link → Viewer**.
2. **Target** — the file you populate into. Either upload it, or leave the upload
   empty to use `TARGET_SHEET_URL` from `.env`. Its financial columns are read
   and preserved; its other tabs are kept intact.
3. Press **Sync now**.
4. Review the counts + preview, then either:
   - **Download updated workbook**, or
   - tick **"Write the result back into the target Google Sheet"** to push it
     straight into the target file (see write-back setup below).

## Writing back into the target Google file

Pushing the result into the target Google file (rather than downloading) needs a
service account, because you can't write to a Drive file through a share link
alone:

1. In Google Cloud, create a service account and download its JSON key. Enable
   the **Google Drive API**.
2. Put the key path in `.env` → `GOOGLE_SERVICE_ACCOUNT_FILE=/path/to/key.json`.
3. Share the **target** file with the service account's email
   (`...@...iam.gserviceaccount.com`) as **Editor**.
4. `pip install google-auth`.
5. In the app, tick **"Write the result back into the target Google Sheet"**.

The target keeps the same file id and links; only its contents are replaced (with
the synced `MASTER` plus every other tab unchanged). If you skip the service
account, just use the **Download** button and re-upload to Drive yourself.

By default the result is written to a new sheet called `MASTER_SYNCED`; every
other tab in your target workbook is left untouched. Tick "Overwrite the MASTER
sheet" if you'd rather replace it in place.

## Notes / things to know

- **Header row.** The column names live on row 2 of this workbook (row 1 is the
  grouped banner: PRICE / TRUCKS / DISCHARGE …). The default header-row setting
  reflects that; change it in the sidebar if your layout differs.
- **Computed columns become values.** Columns like `Value`, `PMS C`, dip
  variances and "days left" are carried across as the numbers the source already
  computed, not as live Excel formulas. If you need live formulas preserved in
  place, use the "Overwrite the MASTER sheet" option against a target that
  already has those formulas, and sync only the raw input columns (edit the
  financial-column list in **Advanced** to also exclude the computed ones from
  the source so they aren't overwritten).
- **Private sheets.** Set `GOOGLE_SERVICE_ACCOUNT_FILE` in `.env` to a service
  account JSON key and share the sheet with that account's email. Reading then
  uses an authenticated Drive download instead of the public export. (Install the
  optional `google-auth` dependency for this.)

## Files

- `app.py` — the Streamlit UI (reads config from `.env`).
- `sync_core.py` — fetch / merge / write logic (no Streamlit, easy to test).
- `.env.example` — template; copy to `.env` and fill in.
- `requirements.txt` — dependencies.
