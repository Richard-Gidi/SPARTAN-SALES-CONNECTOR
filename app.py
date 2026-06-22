"""
Fuel Station MASTER Sync — Streamlit app
========================================

Reads the operational MASTER from a SOURCE Google Sheet, preserves the 7
financial columns from a TARGET file, merges on (DATE, STATION), and writes the
result into a new sheet — or straight back into the target Google file.

All links + layout config live in `.env` (see .env.example) and are never shown
in the interface.

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

load_dotenv()

ENV_SOURCE_URL = os.getenv("SOURCE_SHEET_URL", "").strip()
ENV_TARGET_URL = os.getenv("TARGET_SHEET_URL", "").strip()
SHEET_NAME = os.getenv("SHEET_NAME", DEFAULT_SHEET).strip() or DEFAULT_SHEET
HEADER = int(os.getenv("HEADER_ROW", str(HEADER_ROW)) or HEADER_ROW)
ENV_CREDS = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "").strip() or None

st.set_page_config(page_title="MASTER Sync", page_icon="⛽", layout="wide")


# --------------------------------------------------------------------------- #
# Look & feel
# --------------------------------------------------------------------------- #
def inject_css() -> None:
    st.markdown(
        """
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500;700&display=swap');

:root{
  --ink:#0F1B23; --panel:#16242E; --panel-2:#1B2D38; --line:#273D49;
  --amber:#F4A63B; --amber-soft:#F7C173; --teal:#34C7B5;
  --text:#EAF1F5; --muted:#8DA2AE;
}

.stApp{ background:
   radial-gradient(1200px 600px at 80% -10%, #16323b 0%, rgba(22,50,59,0) 60%),
   var(--ink); }
.block-container{ max-width:1120px; padding-top:1.4rem; padding-bottom:4rem; }

/* hide default chrome for a cleaner console */
#MainMenu, footer{ visibility:hidden; }
header[data-testid="stHeader"]{ background:transparent; height:0; }
[data-testid="stToolbar"], [data-testid="stDecoration"], .stAppDeployButton{ display:none !important; }

html, body, .stApp, p, label, span, div{ font-family:'Inter',sans-serif; color:var(--text); }
h1,h2,h3,h4{ font-family:'Space Grotesk',sans-serif; letter-spacing:-.01em; }

/* ---------- console header ---------- */
.console{
  display:flex; align-items:center; gap:18px;
  padding:22px 26px; margin-bottom:18px;
  background:linear-gradient(180deg,var(--panel) 0%, #132029 100%);
  border:1px solid var(--line); border-radius:18px;
  box-shadow:0 1px 0 rgba(255,255,255,.03) inset, 0 18px 40px -28px #000;
}
.console__mark{
  width:54px;height:54px;flex:0 0 54px;border-radius:14px;
  display:grid;place-items:center;
  background:radial-gradient(120% 120% at 30% 20%, #24333d, #101b22);
  border:1px solid var(--line);
}
.console__title{ font-size:1.7rem;font-weight:700;line-height:1.05;margin:0; }
.console__sub{ color:var(--muted);font-size:.92rem;margin-top:4px; }
.eyebrow{
  font-family:'JetBrains Mono',monospace;font-size:.66rem;letter-spacing:.28em;
  text-transform:uppercase;color:var(--amber);margin:0 0 2px;
}
.gauge-rule{ height:4px;border-radius:99px;margin-top:14px;overflow:hidden;
  background:#0d161d;border:1px solid var(--line); }
.gauge-rule > i{ display:block;height:100%;width:62%;border-radius:99px;
  background:linear-gradient(90deg,var(--amber),var(--amber-soft));
  animation:fill 1.1s cubic-bezier(.2,.7,.2,1) both; }
@keyframes fill{ from{width:0} }
@media (prefers-reduced-motion:reduce){ .gauge-rule>i{animation:none} }

/* ---------- flow strip ---------- */
.flow{ display:flex;align-items:stretch;gap:10px;margin:6px 0 18px;flex-wrap:wrap; }
.flow .node{ flex:1 1 200px;background:var(--panel);border:1px solid var(--line);
  border-radius:14px;padding:14px 16px; }
.flow .node .k{ font-family:'JetBrains Mono',monospace;font-size:.64rem;
  letter-spacing:.22em;text-transform:uppercase;color:var(--muted); }
.flow .node .v{ font-family:'Space Grotesk';font-weight:600;font-size:1.02rem;margin-top:3px; }
.flow .node.keep{ border-color:#3a3320;background:linear-gradient(180deg,#1d2920,#15232c); }
.flow .arrow{ align-self:center;color:var(--amber);font-size:1.2rem;flex:0 0 auto; }

/* ---------- status chips ---------- */
.chip{ display:inline-flex;align-items:center;gap:8px;font-size:.82rem;
  padding:6px 12px;border-radius:99px;border:1px solid var(--line);
  background:#10202a;color:var(--text);font-weight:500; }
.chip .dot{ width:8px;height:8px;border-radius:99px;background:var(--teal);
  box-shadow:0 0 0 3px rgba(52,199,181,.18); }
.chip.warn .dot{ background:var(--amber);box-shadow:0 0 0 3px rgba(244,166,59,.18); }

/* ---------- gauge cluster (results) ---------- */
.cluster{ display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin:8px 0 6px; }
@media (max-width:760px){ .cluster{ grid-template-columns:repeat(2,1fr); } }
.gauge{ background:linear-gradient(180deg,var(--panel) 0%, #11202a 100%);
  border:1px solid var(--line);border-radius:16px;padding:16px 18px 14px; }
.gauge .lab{ font-family:'JetBrains Mono',monospace;font-size:.62rem;
  letter-spacing:.2em;text-transform:uppercase;color:var(--muted); }
.gauge .num{ font-family:'JetBrains Mono',monospace;font-weight:700;
  font-size:2.0rem;line-height:1.1;margin-top:6px;color:var(--text);
  text-shadow:0 0 18px rgba(244,166,59,.10); }
.gauge .bar{ height:5px;border-radius:99px;background:#0c151b;margin-top:12px;overflow:hidden;
  border:1px solid var(--line); }
.gauge .bar > i{ display:block;height:100%;border-radius:99px; }
.gauge.amber .num{ color:var(--amber-soft); }
.gauge.amber .bar > i{ background:linear-gradient(90deg,var(--amber),var(--amber-soft)); }
.gauge.teal  .bar > i{ background:linear-gradient(90deg,#1f9c8d,var(--teal)); }

/* ---------- buttons ---------- */
.stButton > button, .stDownloadButton > button{
  font-family:'Space Grotesk',sans-serif;font-weight:600;letter-spacing:.01em;
  border-radius:12px;padding:.7rem 1rem;border:1px solid var(--line);
  background:var(--panel-2);color:var(--text);transition:transform .08s ease, box-shadow .2s ease, filter .2s; }
.stButton > button:hover, .stDownloadButton > button:hover{ transform:translateY(-1px); }
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"]{
  background:linear-gradient(180deg,var(--amber),#e0902a);
  color:#1a1206;border:0;box-shadow:0 14px 30px -14px rgba(244,166,59,.6); }
.stButton > button[kind="primary"]:hover, .stDownloadButton > button[kind="primary"]:hover{
  filter:brightness(1.04); }

/* ---------- sidebar ---------- */
section[data-testid="stSidebar"]{ background:#0c161d;border-right:1px solid var(--line); }
section[data-testid="stSidebar"] .block-container{ padding-top:1.2rem; }
.side-h{ font-family:'JetBrains Mono',monospace;font-size:.66rem;letter-spacing:.22em;
  text-transform:uppercase;color:var(--amber);margin:.2rem 0 .4rem; }

/* inputs */
.stTextInput input, .stNumberInput input, textarea{
  background:#0d171e !important;border:1px solid var(--line) !important;color:var(--text) !important;
  border-radius:10px !important; }
div[data-baseweb="input"]{ border-radius:10px; }

/* dataframe + alerts soften */
[data-testid="stDataFrame"]{ border:1px solid var(--line);border-radius:14px; }
hr{ border-color:var(--line); }
</style>
        """,
        unsafe_allow_html=True,
    )


PUMP_SVG = """
<svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#F4A63B"
 stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
 <rect x="3" y="3" width="9" height="18" rx="1.5"/>
 <line x1="3" y1="9" x2="12" y2="9"/>
 <path d="M12 7h3a2 2 0 0 1 2 2v7a1.5 1.5 0 0 0 3 0V10l-2.2-2.2"/>
 <circle cx="7.5" cy="6" r="0.6" fill="#F4A63B"/>
</svg>
"""


def header_band() -> None:
    st.markdown(
        f"""
<div class="console">
  <div class="console__mark">{PUMP_SVG}</div>
  <div style="flex:1">
    <div class="eyebrow">Station Reconciliation</div>
    <h1 class="console__title">MASTER&nbsp;Sync</h1>
    <div class="console__sub">Pull operational figures from the source &amp; keep your banking columns intact.</div>
    <div class="gauge-rule"><i></i></div>
  </div>
</div>
        """,
        unsafe_allow_html=True,
    )


def flow_panel(fin_cols) -> None:
    keep = ", ".join(fin_cols[:3]) + (f" +{len(fin_cols)-3} more" if len(fin_cols) > 3 else "")
    st.markdown(
        f"""
<div class="flow">
  <div class="node"><div class="k">Source · operational</div>
    <div class="v">Every column except banking</div></div>
  <div class="arrow">→</div>
  <div class="node keep"><div class="k">Preserved from target</div>
    <div class="v">{keep}</div></div>
  <div class="arrow">→</div>
  <div class="node"><div class="k">Target · matched on</div>
    <div class="v">DATE + STATION</div></div>
</div>
        """,
        unsafe_allow_html=True,
    )


def chip(text: str, warn: bool = False) -> str:
    cls = "chip warn" if warn else "chip"
    return f'<span class="{cls}"><span class="dot"></span>{text}</span>'


def render_gauges(report: dict) -> None:
    out = max(report["output_rows"], 1)
    matched_pct = round(100 * report["matched_rows"] / out)
    new_pct = round(100 * report["new_rows_no_financial"] / out)
    cells = report["financial_cells_preserved"]
    st.markdown(
        f"""
<div class="cluster">
  <div class="gauge teal"><div class="lab">Output rows</div>
    <div class="num">{report['output_rows']:,}</div>
    <div class="bar"><i style="width:100%"></i></div></div>
  <div class="gauge teal"><div class="lab">Matched</div>
    <div class="num">{report['matched_rows']:,}</div>
    <div class="bar"><i style="width:{matched_pct}%"></i></div></div>
  <div class="gauge amber"><div class="lab">New rows</div>
    <div class="num">{report['new_rows_no_financial']:,}</div>
    <div class="bar"><i style="width:{max(new_pct,2)}%"></i></div></div>
  <div class="gauge amber"><div class="lab">Financial cells kept</div>
    <div class="num">{cells:,}</div>
    <div class="bar"><i style="width:100%"></i></div></div>
</div>
        """,
        unsafe_allow_html=True,
    )


inject_css()
header_band()

# --------------------------------------------------------------------------- #
# Sidebar (links + layout config are read silently from .env)
# --------------------------------------------------------------------------- #
with st.sidebar:
    st.markdown('<div class="side-h">Source</div>', unsafe_allow_html=True)
    src_mode = st.radio(
        "Operational data",
        ["Linked source", "Upload an .xlsx"],
        index=0,
        label_visibility="collapsed",
    )
    src_upload = None
    if src_mode == "Linked source":
        st.markdown(
            chip("Source connected") if ENV_SOURCE_URL
            else chip("No source in .env", warn=True),
            unsafe_allow_html=True,
        )
    else:
        src_upload = st.file_uploader("Source workbook (.xlsx)", type=["xlsx"], key="src")

    st.markdown('<div class="side-h" style="margin-top:1.1rem">Target</div>', unsafe_allow_html=True)
    tgt_upload = st.file_uploader(
        "Target workbook (.xlsx)", type=["xlsx"], key="tgt",
        label_visibility="collapsed",
    )
    st.markdown(
        chip("Target file linked") if ENV_TARGET_URL
        else chip("Upload a target, or set it in .env", warn=True),
        unsafe_allow_html=True,
    )

    st.markdown('<div class="side-h" style="margin-top:1.1rem">Output</div>', unsafe_allow_html=True)
    overwrite = st.checkbox("Overwrite the MASTER sheet", value=False)
    out_sheet = st.text_input(
        "New sheet name", value="MASTER_SYNCED", disabled=overwrite
    )

    can_write_back = bool(ENV_TARGET_URL and ENV_CREDS)
    write_back = st.checkbox(
        "Write back into the target file", value=False, disabled=not can_write_back,
        help="Needs a service account with Editor access (see README).",
    )
    if not can_write_back:
        st.caption("Set a service account in .env to enable write-back.")

    with st.expander("Advanced"):
        fin_text = st.text_area(
            "Financial columns to preserve (one per line)",
            value="\n".join(FINANCIAL_COLS), height=160,
        )
        key_text = st.text_input("Match keys", value=", ".join(KEY_COLS))

financial_cols = [c.strip() for c in fin_text.splitlines() if c.strip()]
key_cols = [c.strip() for c in key_text.split(",") if c.strip()]

flow_panel(financial_cols)
run = st.button("Sync now", type="primary", use_container_width=True)


# --------------------------------------------------------------------------- #
# Run
# --------------------------------------------------------------------------- #
if run:
    # ---- source ----
    try:
        with st.spinner("Reading the source MASTER…"):
            if src_mode == "Linked source":
                if not ENV_SOURCE_URL:
                    st.error("No source link configured. Add SOURCE_SHEET_URL to your .env."); st.stop()
                source_bytes = fetch_google_xlsx_bytes(ENV_SOURCE_URL, ENV_CREDS)
            else:
                if src_upload is None:
                    st.error("Upload a source workbook, or switch to the linked source."); st.stop()
                source_bytes = src_upload.getvalue()
            source_df = read_master_from_bytes(source_bytes, SHEET_NAME, HEADER)
    except Exception as e:  # noqa: BLE001
        st.error(f"Couldn't read the source MASTER. {e}"); st.stop()

    # ---- target / write base ----
    try:
        with st.spinner("Reading the target…"):
            if tgt_upload is not None:
                base_bytes = tgt_upload.getvalue()
                target_df = read_master_from_bytes(base_bytes, SHEET_NAME, HEADER)
            elif ENV_TARGET_URL:
                base_bytes = fetch_google_xlsx_bytes(ENV_TARGET_URL, ENV_CREDS)
                target_df = read_master_from_bytes(base_bytes, SHEET_NAME, HEADER)
            else:
                st.error("No target. Upload a target workbook, or set TARGET_SHEET_URL in .env."); st.stop()
    except Exception as e:  # noqa: BLE001
        st.error(f"Couldn't read the target. {e}"); st.stop()

    miss_s = [c for c in key_cols if c not in source_df.columns]
    miss_t = [c for c in key_cols if c not in target_df.columns]
    if miss_s or miss_t:
        st.error(f"Match keys missing — source: {miss_s or 'none'}, target: {miss_t or 'none'}."); st.stop()

    merged, report = build_synced_master(source_df, target_df, financial_cols, key_cols)

    try:
        with st.spinner("Building the workbook…"):
            out_bytes = write_synced_workbook(
                base_bytes, merged,
                out_sheet_name=out_sheet or "MASTER_SYNCED",
                source_sheet_for_headers=SHEET_NAME,
                overwrite_master=overwrite,
            )
    except Exception as e:  # noqa: BLE001
        st.error(f"The merge worked, but building the workbook failed. {e}"); st.stop()

    # ---- results ----
    st.markdown('<div class="eyebrow" style="margin-top:.4rem">Result</div>', unsafe_allow_html=True)
    render_gauges(report)

    if report["new_rows_no_financial"]:
        st.warning(
            f"{report['new_rows_no_financial']} row(s) are in the source but not the target — "
            "added with the banking columns left blank for you to fill in."
        )
    missing_fin = [c for c in financial_cols if c not in report["financial_cols_found_in_target"]]
    if missing_fin:
        st.info(f"Not found in the target, so not preserved: {', '.join(missing_fin)}.")

    where = (f"the {SHEET_NAME} sheet" if overwrite else f"a new {out_sheet or 'MASTER_SYNCED'} sheet")
    st.markdown(
        chip(f"Written to {where} · other tabs untouched"),
        unsafe_allow_html=True,
    )
    st.write("")

    if write_back and can_write_back:
        try:
            with st.spinner("Writing back into the target Google file…"):
                info = write_back_to_drive(ENV_TARGET_URL, out_bytes, ENV_CREDS)
            st.success(f"Target file updated on Drive · modified {info.get('modifiedTime', '?')}.")
        except Exception as e:  # noqa: BLE001
            st.error(
                "Write-back failed (the download below still works). "
                f"Confirm the service account has Editor access. {e}"
            )

    st.download_button(
        "Download updated workbook", data=out_bytes,
        file_name="MASTER_synced.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary", use_container_width=True,
    )

    st.markdown('<div class="side-h" style="margin-top:1.4rem">Preview</div>', unsafe_allow_html=True)
    st.dataframe(merged.head(200), use_container_width=True, height=420)
    st.caption(f"First 200 of {len(merged):,} rows.")
else:
    st.markdown(
        '<div style="color:var(--muted);margin-top:.2rem">Ready when you are — press '
        '<b style="color:var(--amber-soft)">Sync&nbsp;now</b> to pull the latest figures.</div>',
        unsafe_allow_html=True,
    )