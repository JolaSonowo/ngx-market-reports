import streamlit as st
import pandas as pd
import requests
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
from datetime import datetime
import pytz

# --- 1. DATE FORMATTING ---
def get_ordinal_date():
    lagos_tz = pytz.timezone('Africa/Lagos')
    now = datetime.now(lagos_tz)
    day = now.day
    if 11 <= day <= 13:
        suffix = "TH"
    else:
        suffix = {1: "ST", 2: "ND", 3: "RD"}.get(day % 10, "TH")
    return now.strftime(f"{day}{suffix} %b %Y").upper()

# --- 2. UPDATED DATA FETCHING (Using Requests to avoid HTTPError) ---
@st.cache_data(ttl=3600)
def fetch_ngx_data():
    url = "https://www.investing.com/equities/nigeria"
    # Stronger headers to "pretend" we are a real Chrome browser
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
        "Accept-Language": "en-US,en;q=0.9",
        "Referer": "https://www.google.com/"
    }
    
    # 1. Get the page content using requests
    response = requests.get(url, headers=headers, timeout=10)
    
    # Check if the request was successful
    if response.status_code != 200:
        st.error(f"Failed to fetch data. Error code: {response.status_code}")
        return None, None

    # 2. Use pandas to read the tables from the downloaded HTML text
    tables = pd.read_html(io.StringIO(response.text))
    df = tables[0]
    
    # Cleaning data for your exact headers
    df['Close Price'] = pd.to_numeric(df['Last'], errors='coerce')
    df['% Change'] = df['Chg. %'].str.replace('%', '').astype(float)
    df['Naira Change'] = pd.to_numeric(df['Chg'], errors='coerce')
    
    adv = df.nlargest(5, '% Change')[['Symbol', '% Change', 'Close Price', 'Naira Change']]
    dec = df.nsmallest(5, '% Change')[['Symbol', '% Change', 'Close Price', 'Naira Change']]
    return adv, dec

# --- 3. EXCEL GENERATOR ---
def create_excel(adv, dec):
    output = io.BytesIO()
    date_header = get_ordinal_date()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        sheet_name = 'Equity Summary'
        adv.rename(columns={'Symbol': 'Gainers'}).to_excel(writer, sheet_name=sheet_name, startrow=2, index=False)
        dec.rename(columns={'Symbol': 'Decliners'}).to_excel(writer, sheet_name=sheet_name, startrow=10, index=False)
        
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        title_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
        worksheet.merge_range('A1:D1', f"DAILY EQUITY SUMMARY FOR {date_header}", title_format)
    return output.getvalue()

# --- 4. WORD GENERATOR ---
def create_word(adv, dec):
    doc = Document()
    date_header = get_ordinal_date()
    title = doc.add_heading(f"DAILY EQUITY SUMMARY FOR {date_header}", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    for df, label in [(adv, "Gainers"), (dec, "Decliners")]:
        doc.add_heading(label, level=2)
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        hdr_cells = table.rows[0].cells
        cols = [label, "% Change", "Close Price", "Naira Change"]
        for i, text in enumerate(cols):
            hdr_cells[i].text = text
            hdr_cells[i].paragraphs[0].runs[0].font.bold = True
        for _, row in df.iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = str(row['Symbol'])
            row_cells[1].text = f"{row['% Change']:.2f}%"
            row_cells[2].text = f"{row['Close Price']:.2f}"
            row_cells[3].text = f"{row['Naira Change']:.2f}"
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 5. UI ---
st.set_page_config(page_title="NGX Reporter", page_icon="📈")
st.title("NGX Daily Market Reporter")
st.write(f"Today is: **{get_ordinal_date()}**")

if st.button("Generate Today's Market Summary"):
    with st.spinner("Fetching data from NGX..."):
        adv, dec = fetch_ngx_data()
        
    if adv is not None:
        filename_base = f"DAILY EQUITY SUMMARY FOR {get_ordinal_date()}"
        col1, col2 = st.columns(2)
        with col1:
            st.download_button("Download Excel", create_excel(adv, dec), f"{filename_base}.xlsx")
        with col2:
            st.download_button("Download Word", create_word(adv, dec), f"{filename_base}.docx")
