import datetime as _dt
from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path

import html

import pandas as pd
import streamlit as st
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

    html_msg = f"""\
<p>Ciao Luisa,</p>

<p>di seguito gli import di questa settimana.<br>
Sono stati fatti <strong>{tot_total}</strong> import così divisi:</p>

<table border="1" cellspacing="0" cellpadding="4">
  <tr><th>Business Line</th><th>Volumi</th></tr>
  <tr><td>Facility</td><td>{facility}</td></tr>
  <tr><td>Individual</td><td>{individual}</td></tr>
  <tr><td><strong>Totale complessivo</strong></td><td><strong>{tot_total}</strong></td></tr>
</table>

<p><em>Il dato relativo alle Facility comprende anche le Cliniche GP, GIPO e DPP.</em></p>

<p>Di seguito le lavorazioni suddivise per importer:</p>

<table border="1" cellspacing="0" cellpadding="4">
  <tr><th>Importer</th><th>Volumi</th></tr>
  <tr><td>Alessia</td><td>{imp_counts['Alessia']}</td></tr>
  <tr><td>Andrea</td><td>{imp_counts['Andrea']}</td></tr>
  <tr><td>Enrico</td><td>{imp_counts['Enrico']}</td></tr>
  <tr><td>Pedro</td><td>{imp_counts['Pedro']}</td></tr>
  <tr><td><strong>Totale complessivo</strong></td><td><strong>{tot_total}</strong></td></tr>
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
        "   (La app considera sempre la settimana precedente a quella della data scelta.)\n"
        "4. Scarica l'Excel filtrato o copia il messaggio HTML pronto per Gmail."
    )

uploaded_file = st.file_uploader("Carica il CSV export dal CRM", type="csv")

col1, col2 = st.columns(2)
with col1:
    ref_date = st.date_input("Data di riferimento", value=_dt.date.today(), format="DD/MM/YYYY")
with col2:
    pass  # empty spacer or we could put future options

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
    kpi1, kpi2, kpi3 = st.columns(3)
    kpi1.metric("Total Imports", f"{tot_imports}")
    kpi2.metric("Facility", f"{facility}")
    kpi3.metric("Individual", f"{individual}")

    st.subheader("📑 Risultati filtrati")
    st.dataframe(df_filtered, hide_index=True)

    # Messages
    plain_msg, html_msg = build_messages(df_filtered)

    with st.expander("Messaggio plain/Markdown"):
        st.text_area("", plain_msg, height=250)

    with st.expander("Messaggio HTML per Gmail"):
        st.text_area("", html_msg, height=300)

    # Download buttons
    def _to_excel_bytes(df):
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False)
        buf.seek(0)
        return buf

    excel_bytes = _to_excel_bytes(df_filtered)
    st.download_button(
        "⬇️ Scarica Excel filtrato",
        data=excel_bytes,
        file_name="import_filtered.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.download_button(
        "⬇️ Scarica HTML Gmail",
        data=html_msg,
        file_name="report_message.html",
        mime="text/html",
    )
else:
    st.info("Carica un file CSV per iniziare.")

# --------------------------------------------------------
# FOOTER / LINKS
# --------------------------------------------------------
st.markdown("---")
st.markdown(
    "Made with ❤️ by the Imports team  •  Source code: <https://github.com/your-org/import-report-app>  •  Free hosting via Streamlit Community Cloud"
)

# --------------------------------------------------------
# REQUIREMENTS (to put in requirements.txt):
#   streamlit
#   pandas
#   python-dateutil
#   xlsxwriter
# --------------------------------------------------------
