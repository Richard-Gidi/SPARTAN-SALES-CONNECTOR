"""
Fuel Station MASTER Sync — Streamlit app
========================================

Reads the operational MASTER from a SOURCE Google Sheet, preserves the 7
financial columns from a TARGET (an uploaded .xlsx, or a second Google Sheet),
merges them on (DATE, STATION), and writes the result to a new sheet of a
downloadable workbook.

Google links + config live in a `.env` file (see .env.example):

    SOURCE_SHEET_URL=...            # operational source (required for Google mode)
    TARGET_SHEET_URL=...            # optional: financial source, if it's a sheet
    SHEET_NAME=MASTER
    HEADER_ROW=2
    GOOGLE_SERVICE_ACCOUNT_FILE=... # optional: for private links

Run with:  streamlit run app.py
"""

import io
import os

import streamlit as st
from dotenv import load_dotenv

from sync_core import (
    FINANCIAL_COLS,
    KEY_COLS,
    DEFAULT_SHEET,
    HEADER_ROW,
    fetch_google_xlsx_bytes,
    read_master_from_bytes,
    build_synced_master,
    write_synced_workbook,
    write_back_to_drive,
)

load_dotenv()  # pull .env into os.environ

ENV_SOURCE_URL = os.getenv("SOURCE_SHEET_URL", "").strip()
ENV_TARGET_URL = os.getenv("TARGET_SHEET_URL", "").strip()
ENV_SHEET_NAME = os.getenv("SHEET_NAME", DEFAULT_SHEET).strip() or DEFAULT_SHEET
ENV_HEADER_ROW = int(os.getenv("HEADER_ROW", str(HEADER_ROW)) or HEADER_ROW)
ENV_CREDS = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "").strip() or None

st.set_page_config(page_title="MASTER Sync", page_icon="⛽", layout="wide")
st.title("⛽ Fuel Station — MASTER Sync")
st.caption(
    "Operational data from the source MASTER + the financial columns from your "
    "own workbook, merged on (DATE, STATION). Links are read from `.env`."
)

# --------------------------------------------------------------------------- #
# Sidebar
# --------------------------------------------------------------------------- #
with st.sidebar:
    st.header("1 · Source (operational data)")
    src_mode = st.radio(
        "Where is the source MASTER?",
        ["Google Sheet (from .env)", "Upload an .xlsx"],
        index=0,
    )

    src_upload = None
    source_url = ENV_SOURCE_URL
    if src_mode == "Google Sheet (from .env)":
        if ENV_SOURCE_URL:
            st.success("SOURCE_SHEET_URL loaded from .env")
            st.code(ENV_SOURCE_URL, language=None)
        else:
            st.error("SOURCE_SHEET_URL is not set in your .env file.")
    else:
        src_upload = st.file_uploader("Source workbook (.xlsx)", type=["xlsx"], key="src")

    st.divider()
    st.header("2 · Target (financial columns)")
    tgt_upload = st.file_uploader(
        "Target workbook to update (.xlsx)", type=["xlsx"], key="tgt"
    )
    if ENV_TARGET_URL:
        st.caption(
            "If you leave this empty, financial columns are read from "
            "**TARGET_SHEET_URL** in your .env."
        )
        st.code(ENV_TARGET_URL, language=None)
    else:
        st.caption("Upload the workbook that holds your financial columns.")

    st.divider()
    st.header("3 · Options")
    sheet_name = st.text_input("Sheet/tab name", value=ENV_SHEET_NAME)
    header_row = st.number_input(
        "Header row (1-indexed)", min_value=1, max_value=10, value=ENV_HEADER_ROW
    )
    overwrite = st.checkbox(
        "Overwrite the MASTER sheet instead of writing a new one", value=False
    )
    out_sheet = st.text_input(
        "Output sheet name", value="MASTER_SYNCED", disabled=overwrite
    )

    can_write_back = bool(ENV_TARGET_URL and ENV_CREDS)
    write_back = st.checkbox(
        "Write the result back into the target Google Sheet",
        value=False,
        disabled=not can_write_back,
        help=(
            "Requires TARGET_SHEET_URL and GOOGLE_SERVICE_ACCOUNT_FILE in .env, "
            "and the service account must have Editor access to the file."
        ),
    )
    if not can_write_back:
        st.caption(
            "Write-back needs TARGET_SHEET_URL + GOOGLE_SERVICE_ACCOUNT_FILE in "
            ".env. Without them you'll get a downloadable workbook instead."
        )

    with st.expander("Advanced — columns & keys"):
        fin_text = st.text_area(
            "Financial columns to PRESERVE from the target (one per line)",
            value="\n".join(FINANCIAL_COLS),
            height=170,
        )
        key_text = st.text_input("Match keys (comma separated)", value=", ".join(KEY_COLS))
        if ENV_CREDS:
            st.caption(f"Private-sheet access via service account: `{ENV_CREDS}`")

financial_cols = [c.strip() for c in fin_text.splitlines() if c.strip()]
key_cols = [c.strip() for c in key_text.split(",") if c.strip()]

st.info(
    "**Operational** columns come from the source. **Financial** columns kept "
    f"from the target: {', '.join(financial_cols)}. Matched on "
    f"**{' + '.join(key_cols)}**."
)

run = st.button("🔄  Sync now", type="primary", use_container_width=True)


# --------------------------------------------------------------------------- #
# Run
# --------------------------------------------------------------------------- #
if run:
    hdr = int(header_row)

    # ---- read SOURCE (df + raw bytes; bytes used as a write-base fallback) ----
    try:
        with st.spinner("Reading the source MASTER…"):
            if src_mode == "Google Sheet (from .env)":
                if not source_url:
                    st.error("SOURCE_SHEET_URL is not set in .env."); st.stop()
                source_bytes = fetch_google_xlsx_bytes(source_url, ENV_CREDS)
            else:
                if src_upload is None:
                    st.error("Please upload the source workbook."); st.stop()
                source_bytes = src_upload.getvalue()
            source_df = read_master_from_bytes(source_bytes, sheet_name, hdr)
    except Exception as e:  # noqa: BLE001
        st.error(f"Could not read the source MASTER:\n\n{e}"); st.stop()

    # ---- read TARGET (financial columns) + choose the write-base workbook ----
    try:
        with st.spinner("Reading the target (financial) data…"):
            if tgt_upload is not None:
                base_bytes = tgt_upload.getvalue()
                target_df = read_master_from_bytes(base_bytes, sheet_name, hdr)
            elif ENV_TARGET_URL:
                target_bytes = fetch_google_xlsx_bytes(ENV_TARGET_URL, ENV_CREDS)
                target_df = read_master_from_bytes(target_bytes, sheet_name, hdr)
                base_bytes = target_bytes  # write into the target file itself
            else:
                st.error(
                    "No target provided. Upload a target workbook, or set "
                    "TARGET_SHEET_URL in your .env."
                ); st.stop()
    except Exception as e:  # noqa: BLE001
        st.error(f"Could not read the target MASTER:\n\n{e}"); st.stop()

    # ---- key sanity ----
    miss_s = [c for c in key_cols if c not in source_df.columns]
    miss_t = [c for c in key_cols if c not in target_df.columns]
    if miss_s or miss_t:
        st.error(
            f"Match keys missing — source: {miss_s or 'none'}, target: {miss_t or 'none'}. "
            "Check the sheet name / header row."
        ); st.stop()

    # ---- merge ----
    merged, report = build_synced_master(source_df, target_df, financial_cols, key_cols)

    # ---- write workbook ----
    try:
        with st.spinner("Writing the updated workbook…"):
            out_bytes = write_synced_workbook(
                base_bytes,
                merged,
                out_sheet_name=out_sheet or "MASTER_SYNCED",
                source_sheet_for_headers=sheet_name,
                overwrite_master=overwrite,
            )
    except Exception as e:  # noqa: BLE001
        st.error(f"Merge succeeded but writing the workbook failed:\n\n{e}"); st.stop()

    # ---- report ----
    st.success("Sync complete.")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Output rows", report["output_rows"])
    c2.metric("Matched rows", report["matched_rows"])
    c3.metric("New rows (no financial)", report["new_rows_no_financial"])
    c4.metric("Financial cells kept", report["financial_cells_preserved"])

    if report["new_rows_no_financial"]:
        st.warning(
            f"{report['new_rows_no_financial']} row(s) are in the source but not the "
            "target; added with financial columns left blank for you to fill in."
        )
    missing_fin = [c for c in financial_cols if c not in report["financial_cols_found_in_target"]]
    if missing_fin:
        st.info(f"Not found in the target (so not preserved): {', '.join(missing_fin)}.")

    where = (f"the **{sheet_name}** sheet" if overwrite
             else f"a new **{out_sheet or 'MASTER_SYNCED'}** sheet")
    st.markdown(f"Result written to {where}; all other tabs are unchanged.")

    # ---- optional: push back into the target Google file ----
    if write_back and can_write_back:
        try:
            with st.spinner("Writing back into the target Google Sheet…"):
                info = write_back_to_drive(ENV_TARGET_URL, out_bytes, ENV_CREDS)
            st.success(
                f"Target file updated on Drive "
                f"(id {info.get('id', '?')}, modified {info.get('modifiedTime', '?')})."
            )
        except Exception as e:  # noqa: BLE001
            st.error(
                "Write-back failed (the download below still works). "
                "Check that the service account has Editor access to the file.\n\n"
                f"{e}"
            )

    st.download_button(
        "⬇️  Download updated workbook",
        data=out_bytes,
        file_name="MASTER_synced.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    st.subheader("Preview")
    st.dataframe(merged.head(200), use_container_width=True, height=420)
    st.caption(f"Showing first 200 of {len(merged)} rows.")
else:
    st.markdown("Confirm the source/target in the sidebar, then press **Sync now**.")
