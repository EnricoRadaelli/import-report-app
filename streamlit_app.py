import datetime as _dt
from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path
import html

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from dateutil.relativedelta import relativedelta, MO

# --------------------------------------------------------
# CONFIG — tweak only if your business rules change
# --------------------------------------------------------
ALLOWED_IMPORTERS = {
    "Andrea Sergi",
    "Enrico Radaelli",
    "Alessia Pellegrino",
    "Andrea Pedroli",
}

BUSINESS_LINE_MAP = {
    "individuals": "Individual",
    "gp": "Individual",
    "gp gruppi": "Facility",
    "gipo": "Facility",
    "dp phone": "Facility",
    "clinic agenda": "Facility",
}

IMPORTER_SHORT = {
    "Alessia Pellegrino": "Alessia",
    "Andrea Sergi": "Andrea",
    "Enrico Radaelli": "Enrico",
    "Andrea Pedroli": "Pedro",
}

KEEP_COLS = [
    "Url",
    "Ticket Name",
    "Importer",
    "Import Type",
    "Business Line",
    "Close Date",
]

CLOSE_FMT = "%b %d, %Y, %I:%M:%S %p"  # e.g. "Jul 14, 2025, 04:20:01 PM"

# --------------------------------------------------------
# UTILITY FUNCTIONS (largely unchanged from your script)
# --------------------------------------------------------

STYLE_CELL = 'style="border:1px solid #cccccc;padding:4px;text-align:left;"'
STYLE_TABLE = 'style="border-collapse:collapse;border:1px solid #cccccc;"'


def previous_week_window(reference_date=None):
    """Return Monday–Sunday dates of the *previous* week w.r.t reference_date."""
    if reference_date is None:
        reference_date = _dt.date.today()
    this_week_monday = reference_date + relativedelta(weekday=MO(-1))
    prev_monday = this_week_monday - timedelta(weeks=1)
    prev_sunday = prev_monday + timedelta(days=6)
    return prev_monday, prev_sunday


def business_line_cat(raw_bl: str) -> str:
    return BUSINESS_LINE_MAP.get(str(raw_bl).strip().lower(), "Individual")


def load_and_filter(df: pd.DataFrame, reference_date=None):
    """Apply column subset + filters to the dataframe."""
    df = df[KEEP_COLS].copy()

    # Normalize & parse date
    df["Close Date"] = (
        df["Close Date"].astype(str).str.replace("\u202f", " ", regex=False)
        .pipe(pd.to_datetime, format=CLOSE_FMT, errors="coerce")
        .dt.date
    )

    start, end = previous_week_window(reference_date)
    mask_period = df["Close Date"].between(start, end, inclusive="both")
    mask_type = df["Import Type"].str.startswith(("Complete", "No Importation"), na=False)
    mask_imp = df["Importer"].isin(ALLOWED_IMPORTERS)

    df = df[mask_period & mask_type & mask_imp].reset_index(drop=True)
    df["BL_CAT"] = df["Business Line"].apply(business_line_cat)
    df["ImporterShort"] = df["Importer"].map(IMPORTER_SHORT)
    return df


def build_messages(df: pd.DataFrame):
    """Return (plain_markdown, html_email) strings ready to copy."""
    bl_counts = df["BL_CAT"].value_counts()
    facility = int(bl_counts.get("Facility", 0))
    individual = int(bl_counts.get("Individual", 0))
    tot_total = len(df)

    imp_counts = (
        df["ImporterShort"].value_counts().reindex(["Alessia", "Andrea", "Enrico", "Pedro"], fill_value=0)
    )

    fac_df = (
        df[df["BL_CAT"] == "Facility"][["Ticket Name", "Url"]]
        .dropna()
        .drop_duplicates(subset="Url")
        .sort_values("Ticket Name")
    )

    if fac_df.empty:
        links_md = "_Nessuna facility questa settimana_"
        links_html = "<em>Nessuna facility questa settimana</em>"
    else:
        links_md = "\n".join(f"- [{n}]({u})" for n, u in fac_df.values)
        links_html = "<ul>" + "".join(
            f'<li><a href="{html.escape(u)}">{html.escape(n)}</a></li>' for n, u in fac_df.values
        ) + "</ul>"

    plain = f"""\
Ciao Luisa,
di seguito gli import di questa settimana. Sono stati fatti {tot_total} import così divisi:

Business Line\tVolumi
Facility\t{facility}
Individual\t{individual}
Totale complessivo\t{tot_total}

*Il dato relativo alle Facility comprende anche le Cliniche GP, GIPO e DPP.*

Di seguito le lavorazioni suddivise per importer:

Importer\tVolumi
Alessia\t{imp_counts['Alessia']}
Andrea\t{imp_counts['Andrea']}
Enrico\t{imp_counts['Enrico']}
Pedro\t{imp_counts['Pedro']}
Totale complessivo\t{tot_total}

Di seguito i link delle cliniche (sia CRM che GIPO che Gruppi GP che Cliniche DPP) interessate:

{links_md}
"""

    # --- Gmail‑friendly HTML: inline CSS ---
    html_msg = f"""\
<p>Ciao Luisa,</p>

<p>di seguito gli import di questa settimana.<br/>
Sono stati fatti <strong>{tot_total}</strong> import così divisi:</p>

<table {STYLE_TABLE}>
  <tr>
    <th {STYLE_CELL}>Business Line</th>
    <th {STYLE_CELL}>Volumi</th>
  </tr>
  <tr>
    <td {STYLE_CELL}>Facility</td><td {STYLE_CELL}>{facility}</td>
  </tr>
  <tr>
    <td {STYLE_CELL}>Individual</td><td {STYLE_CELL}>{individual}</td>
  </tr>
  <tr>
    <td {STYLE_CELL}><strong>Totale complessivo</strong></td>
    <td {STYLE_CELL}><strong>{tot_total}</strong></td>
  </tr>
</table>

<p><em>Il dato relativo alle Facility comprende anche le Cliniche GP, GIPO e DPP.</em></p>

<p>Di seguito le lavorazioni suddivise per importer:</p>

<table {STYLE_TABLE}>
  <tr>
    <th {STYLE_CELL}>Importer</th>
    <th {STYLE_CELL}>Volumi</th>
  </tr>
  <tr><td {STYLE_CELL}>Alessia</td><td {STYLE_CELL}>{imp_counts['Alessia']}</td></tr>
  <tr><td {STYLE_CELL}>Andrea</td><td {STYLE_CELL}>{imp_counts['Andrea']}</td></tr>
  <tr><td {STYLE_CELL}>Enrico</td><td {STYLE_CELL}>{imp_counts['Enrico']}</td></tr>
  <tr><td {STYLE_CELL}>Pedro</td><td {STYLE_CELL}>{imp_counts['Pedro']}</td></tr>
  <tr>
    <td {STYLE_CELL}><strong>Totale complessivo</strong></td>
    <td {STYLE_CELL}><strong>{tot_total}</strong></td>
  </tr>
</table>

<p>Di seguito i link delle cliniche (sia CRM che GIPO che Gruppi GP che Cliniche DPP) interessate:</p>

{links_html}
"""
    return plain, html_msg


# --------------------------------------------------------
# STREAMLIT UI
# --------------------------------------------------------
st.set_page_config(page_title="Weekly Import Report", page_icon="📊", layout="centered")
st.title("📊 Weekly Import Report Generator")

with st.expander("ℹ️ Istruzioni", expanded=False):
    st.markdown(
        "1. Scarica il CSV dal CRM come fai di solito.\n"
        "2. Caricalo qui sotto.\n"
        "3. Facoltativo: cambia la *data di riferimento* se vuoi calcolare su un'altra settimana.\n"
        "4. Scarica l'Excel filtrato o copia il messaggio HTML pronto per Gmail."
    )

uploaded_file = st.file_uploader("Carica il CSV export dal CRM", type="csv")

col1, col2 = st.columns(2)
with col1:
    ref_date = st.date_input("Data di riferimento", value=_dt.date.today(), format="DD/MM/YYYY")
with col2:
    pass  # spacer per opzioni future

if uploaded_file:
    # Read & process
    try:
        raw_df = pd.read_csv(uploaded_file)
    except UnicodeDecodeError:
        raw_df = pd.read_csv(uploaded_file, encoding="latin1")

    df_filtered = load_and_filter(raw_df, reference_date=ref_date)

    tot_imports = len(df_filtered)
    bl_counts = df_filtered["BL_CAT"].value_counts()
    facility = int(bl_counts.get("Facility", 0))
    individual = int(bl_counts.get("Individual", 0))

    # KPI Cards
    kpi1, kpi2, kpi
