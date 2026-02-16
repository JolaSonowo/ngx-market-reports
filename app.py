import streamlit as st
import pandas as pd
import requests
from datetime import datetime, timedelta
import pytz
import io
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

# -----------------------------------
# TIME LOGIC (LAGOS MARKET RULE)
# -----------------------------------

def get_effective_trading_date():
    lagos = pytz.timezone("Africa/Lagos")
    now = datetime.now(lagos)

    # Weekend fallback
    if now.weekday() == 5:  # Saturday
        now -= timedelta(days=1)
    elif now.weekday() == 6:  # Sunday
        now -= timedelta(days=2)

    cutoff = now.replace(hour=14, minute=40, second=0, microsecond=0)

    # Before 2:40PM → use previous trading day
    if now < cutoff:
        if now.weekday() == 0:  # Monday
            now -= timedelta(days=3)
        else:
            now -= timedelta(days=1)

    return now.strftime("%d %b %Y").upper()


# -----------------------------------
# FETCH NGX DATA (NO SELENIUM)
# -----------------------------------

@st.cache_data(ttl=900)  # cache 15 minutes
def fetch_ngx_data():

    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/json"
    }

    # NGX equity snapshot endpoint
    url = "https://ngxgroup.com/api/marketdata/equities"

    response = requests.get(url, headers=headers, timeout=20)
    response.raise_for_status()

    data = response.json()

    df = pd.DataFrame(data)

    # Standardize columns
    df.columns = df.columns.str.strip()

    df["percentChange"] = pd.to_numeric(df["percentChange"], errors="coerce")
    df["last"] = pd.to_numeric(df["last"], errors="coerce")
    df["change"] = pd.to_numeric(df["change"], errors="coerce")

    df = df.dropna(subset=["percentChange"])

    # Top 5 Advancers
    adv = df.sort_values("percentChange", ascending=False).head(5)

    # Top 5 Decliners
    dec = df.sort_values("percentChange", ascending=True).head(5)

    return adv, dec


# -----------------------------------
# REPORT GENERATION
# -----------------------------------

def create_excel(adv, dec, date_str):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        adv[["symbol", "percentChange", "last", "change"]].to_excel(
            writer, sheet_name="Summary", startrow=2, index=False
        )

        dec[["symbol", "percentChange", "last", "change"]].to_excel(
            writer, sheet_name="Summary", startrow=12, index=False
        )

        ws = writer.sheets["Summary"]
        fmt = writer.book.add_format(
            {"bold": True, "align": "center", "font_size": 14}
        )
        ws.merge_range("A1:D1", f"DAILY EQUITY SUMMARY FOR {date_str}", fmt)

    return output.getvalue()


def create_word(adv, dec, date_str):
    doc = Document()
    heading = doc.add_heading(
        f"DAILY EQUITY SUMMARY FOR {date_str}", level=1
    )
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for df, label in [(adv, "Top 5 Advancers"), (dec, "Top 5 Decliners")]:
        doc.add_heading(label, level=2)
        table = doc.add_table(rows=1, cols=4)
        table.style = "Table Grid"

        headers = ["Symbol", "% Change", "Last Price", "Change"]
        for i, h in enumerate(headers):
            table.rows[0].cells[i].text = h

        for _, row in df.iterrows():
            cells = table.add_row().cells
            cells[0].text = str(row["symbol"])
            cells[1].text = f"{row['percentChange']:.2f}%"
            cells[2].text = f"{row['last']:.2f}"
            cells[3].text = f"{row['change']:.2f}"

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


# -----------------------------------
# STREAMLIT UI
# -----------------------------------

st.set_page_config(page_title="NGX Live Reporter", page_icon="🇳🇬")
st.title("📈 NGX Top 5 Advancers & Decliners")

date_header = get_effective_trading_date()
st.caption(f"Effective Trading Date: {date_header}")

if st.button("Load Market Data"):

    with st.spinner("Fetching live NGX data..."):
        adv, dec = fetch_ngx_data()

    st.success("Data Loaded Successfully")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Top 5 Advancers")
        st.dataframe(
            adv[["symbol", "percentChange", "last", "change"]],
            use_container_width=True
        )

    with col2:
        st.subheader("Top 5 Decliners")
        st.dataframe(
            dec[["symbol", "percentChange", "last", "change"]],
            use_container_width=True
        )

    st.download_button(
        "Download Excel Report",
        create_excel(adv, dec, date_header),
        file_name="NGX_Summary.xlsx"
    )

    st.download_button(
        "Download Word Report",
        create_word(adv, dec, date_header),
        file_name="NGX_Summary.docx"
    )
