import streamlit as st
import pandas as pd
import requests
from datetime import datetime, timedelta
import pytz
import io
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH


# -------------------------------------------------
# TRADING DATE LOGIC (LAGOS TIME, 2:40PM CUTOFF)
# -------------------------------------------------

def get_effective_trading_date():
    lagos = pytz.timezone("Africa/Lagos")
    now = datetime.now(lagos)

    # Weekend → fallback to Friday
    if now.weekday() == 5:  # Saturday
        now -= timedelta(days=1)
    elif now.weekday() == 6:  # Sunday
        now -= timedelta(days=2)

    cutoff = now.replace(hour=14, minute=40, second=0, microsecond=0)

    # Before 2:40PM → previous trading day
    if now < cutoff:
        if now.weekday() == 0:  # Monday before 2:40
            now -= timedelta(days=3)
        else:
            now -= timedelta(days=1)

    return now.strftime("%d %b %Y").upper()


# -------------------------------------------------
# FETCH NGX DATA (NO SELENIUM)
# -------------------------------------------------

@st.cache_data(ttl=900)
def fetch_ngx_data():

    url = "https://ngxgroup.com/wp-admin/admin-ajax.php"

    payload = {
        "action": "wpdatatables_get_table_data",
        "table_id": "2"  # If this fails, try "1"
    }

    headers = {
        "User-Agent": "Mozilla/5.0",
        "X-Requested-With": "XMLHttpRequest",
        "Referer": "https://ngxgroup.com/market-data/equities/",
    }

    response = requests.post(url, data=payload, headers=headers, timeout=30)

    if response.status_code != 200:
        st.error(f"NGX request failed: {response.status_code}")
        return None, None

    json_data = response.json()

    if "data" not in json_data:
        st.error("Unexpected NGX response structure.")
        return None, None

    rows = json_data["data"]

    df = pd.DataFrame(rows)

    # wpDataTables returns indexed columns — rename manually
    df.columns = [
        "symbol",
        "securityName",
        "market",
        "last",
        "change",
        "percentChange",
        "volume",
        "value"
    ]

    # Clean numeric columns
    df["percentChange"] = (
        df["percentChange"]
        .astype(str)
        .str.replace('%', '', regex=False)
    )
    df["percentChange"] = pd.to_numeric(df["percentChange"], errors="coerce")

    df["last"] = (
        df["last"]
        .astype(str)
        .str.replace(',', '', regex=False)
    )
    df["last"] = pd.to_numeric(df["last"], errors="coerce")

    df["change"] = (
        df["change"]
        .astype(str)
        .str.replace(',', '', regex=False)
    )
    df["change"] = pd.to_numeric(df["change"], errors="coerce")

    df = df.dropna(subset=["percentChange"])

    adv = df.sort_values("percentChange", ascending=False).head(5)
    dec = df.sort_values("percentChange", ascending=True).head(5)

    return adv, dec


# -------------------------------------------------
# REPORT GENERATION
# -------------------------------------------------

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

        fmt = writer.book.add_format({
            "bold": True,
            "align": "center",
            "font_size": 14
        })

        ws.merge_range(
            "A1:D1",
            f"DAILY EQUITY SUMMARY FOR {date_str}",
            fmt
        )

    return output.getvalue()


def create_word(adv, dec, date_str):

    doc = Document()

    heading = doc.add_heading(
        f"DAILY EQUITY SUMMARY FOR {date_str}",
        level=1
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


# -------------------------------------------------
# STREAMLIT UI
# -------------------------------------------------

st.set_page_config(page_title="NGX Market Reporter", page_icon="🇳🇬")
st.title("📈 NGX Top 5 Advancers & Decliners")

date_header = get_effective_trading_date()
st.caption(f"Effective Trading Date: {date_header}")

if st.button("Load Market Data"):

    with st.spinner("Fetching live NGX data..."):
        adv, dec = fetch_ngx_data()

    if adv is None:
        st.stop()

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
