import streamlit as st
import pandas as pd
import requests
import io
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime
import pytz

# --- 1. TIME & DATE LOGIC ---
def get_lagos_time():
    return datetime.now(pytz.timezone('Africa/Lagos'))

def get_ordinal_date():
    now = get_lagos_time()
    day = now.day
    suffix = "TH" if 11 <= day <= 13 else {1: "ST", 2: "ND", 3: "RD"}.get(day % 10, "TH")
    return now.strftime(f"{day}{suffix} %b %Y").upper()

def is_report_available():
    now = get_lagos_time()

    if now.weekday() > 4:
        return False, "The market is closed for the weekend."
    

    report_time = now.replace(hour=14, minute=40, second=0, microsecond=0)
    if now < report_time:
        return False, f"The report will be available at 2:40 PM. Current time: {now.strftime('%I:%M %p')}"
    
    return True, "Report is ready for download!"

# --- 2. DATA FETCHING ---
@st.cache_data(ttl=600)
def fetch_ngx_data():
    session = requests.Session()
    url = "https://ngxgroup.com/exchange/data/equities-price-list/"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Referer": "https://ngxgroup.com/"
    }
    try:
        response = session.get(url, headers=headers, timeout=15)
        tables = pd.read_html(io.StringIO(response.text))
        df = tables[0]
        

        df = df.rename(columns={
            'Symbol': 'Ticker', 
            'Current': 'Close Price', 
            'Change': 'Naira Change', 
            '% Change': '% Change'
        }, errors='ignore')

        for col in ['% Change', 'Close Price', 'Naira Change']:
            df[col] = pd.to_numeric(df[col], errors='coerce')

        adv = df.nlargest(5, '% Change')[['Ticker', '% Change', 'Close Price', 'Naira Change']]
        dec = df.nsmallest(5, '% Change')[['Ticker', '% Change', 'Close Price', 'Naira Change']]
        return adv, dec
    except:
        return None, None

# --- 3. FILE GENERATORS (EXCEL & WORD) ---
def create_excel(adv, dec):
    output = io.BytesIO()
    date_header = get_ordinal_date()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        adv.rename(columns={'Ticker': 'Gainers'}).to_excel(writer, sheet_name='Summary', startrow=2, index=False)
        dec.rename(columns={'Ticker': 'Decliners'}).to_excel(writer, sheet_name='Summary', startrow=10, index=False)
        worksheet = writer.sheets['Summary']
        title_fmt = writer.book.add_format({'bold': True, 'align': 'center', 'font_size': 14})
        worksheet.merge_range('A1:D1', f"DAILY EQUITY SUMMARY FOR {date_header}", title_fmt)
    return output.getvalue()

def create_word(adv, dec):
    doc = Document()
    title = doc.add_heading(f"DAILY EQUITY SUMMARY FOR {get_ordinal_date()}", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for df, label in [(adv, "Gainers"), (dec, "Decliners")]:
        doc.add_heading(label, level=2)
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        cols = [label, "% Change", "Close Price", "Naira Change"]
        for i, text in enumerate(cols):
            table.rows[0].cells[i].text = text
        for _, row in df.iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text, row_cells[1].text = str(row['Ticker']), f"{row['% Change']:.2f}%"
            row_cells[2].text, row_cells[3].text = f"{row['Close Price']:.2f}", f"{row['Naira Change']:.2f}"
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 4. STREAMLIT INTERFACE ---
st.set_page_config(page_title="NGX Reporter", page_icon="📈")
st.title("NGX Daily Market Reporter")

available, message = is_report_available()

if not available:
    st.warning(f"{message}")
    st.info("Please check back after 2:40 PM Lagos time.")
else:
    st.success(f"{message}")
    if st.button("Fetch and Prepare Documents"):
        adv, dec = fetch_ngx_data()
        if adv is not None:
            filename = f"DAILY EQUITY SUMMARY FOR {get_ordinal_date()}"
            st.download_button("Download Excel Report", create_excel(adv, dec), f"{filename}.xlsx")
            st.download_button("Download Word Report", create_word(adv, dec), f"{filename}.docx")
        else:
            st.error("Could not reach NGX. Please refresh.")
