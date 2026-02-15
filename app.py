import streamlit as st
import pandas as pd
import requests
import io
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime, timedelta
import pytz

# --- 1. TIME & DATE LOGIC (NGX WEEKEND HANDLING) ---
def get_lagos_time():
    return datetime.now(pytz.timezone('Africa/Lagos'))

def format_ordinal(dt):
    day = dt.day
    suffix = "TH" if 11 <= day <= 13 else {1: "ST", 2: "ND", 3: "RD"}.get(day % 10, "TH")
    return dt.strftime(f"{day}{suffix} %b %Y").upper()

def get_report_info():
    now = get_lagos_time()
    day_of_week = now.weekday() # 0=Mon, 4=Fri, 5=Sat, 6=Sun
    cutoff = now.replace(hour=14, minute=40, second=0, microsecond=0)
    
    # If Saturday or Sunday, always show Friday's data
    if day_of_week >= 5:
        report_date = now - timedelta(days=(day_of_week - 4))
        msg = "Weekend Mode: Downloading Friday's Closing Data"
    # If Monday-Friday before 2:40 PM, show previous day
    elif now < cutoff:
        days_back = 3 if day_of_week == 0 else 1
        report_date = now - timedelta(days=days_back)
        msg = f"Market Open: Previewing {format_ordinal(report_date)} Data"
    else:
        report_date = now
        msg = "Market Closed: Today's Summary is Ready"
    return report_date, msg

# --- 2. THE ANTI-BLOCK DATA FETCH ---
@st.cache_data(ttl=600)
def fetch_ngx_data():
    session = requests.Session()
    # The direct equities table URL
    url = "https://ngxgroup.com/exchange/data/equities-price-list/"
    
    # Advanced headers to bypass 403 blocks
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Referer": "https://ngxgroup.com/",
        "Accept-Language": "en-US,en;q=0.9"
    }

    try:
        # Pre-visit homepage to establish a session cookie
        session.get("https://ngxgroup.com/", headers=headers, timeout=10)
        
        # Pull the table data
        response = session.get(url, headers=headers, timeout=20)
        if response.status_code != 200:
            return None, None

        # Parse tables (usually index 0 is the equities list)
        tables = pd.read_html(io.StringIO(response.text))
        df = tables[0]
        
        # Normalize NGX columns to your specific requirements
        # NGX uses 'Symbol', 'Current', 'Change', '% Change'
        df = df.rename(columns={
            'Symbol': 'Ticker', 
            'Current': 'Close Price', 
            'Change': 'Naira Change'
        }, errors='ignore')

        # CLEAN DATA: Remove commas and percentage signs before converting to numbers
        for col in ['% Change', 'Close Price', 'Naira Change']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col].astype(str).str.replace('%', '').str.replace(',', ''), errors='coerce')

        # Select Top 5 for report
        adv = df.nlargest(5, '% Change')[['Ticker', '% Change', 'Close Price', 'Naira Change']]
        dec = df.nsmallest(5, '% Change')[['Ticker', '% Change', 'Close Price', 'Naira Change']]
        return adv, dec
    except Exception as e:
        st.error(f"NGX Connection Refused: {e}")
        return None, None

# --- 3. DOCUMENT EXPORT (EXCEL & WORD) ---
def create_excel(adv, dec, date_str):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        adv.rename(columns={'Ticker': 'Gainers'}).to_excel(writer, sheet_name='Equity Summary', startrow=2, index=False)
        dec.rename(columns={'Ticker': 'Decliners'}).to_excel(writer, sheet_name='Equity Summary', startrow=10, index=False)
        
        ws = writer.sheets['Equity Summary']
        fmt = writer.book.add_format({'bold': True, 'align': 'center', 'font_size': 14})
        ws.merge_range('A1:D1', f"DAILY EQUITY SUMMARY FOR {date_str}", fmt)
    return output.getvalue()

def create_word(adv, dec, date_str):
    doc = Document()
    title = doc.add_heading(f"DAILY EQUITY SUMMARY FOR {date_str}", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    for df, label in [(adv, "Gainers"), (dec, "Decliners")]:
        doc.add_heading(label, level=2)
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        headers = [label, "% Change", "Close Price", "Naira Change"]
        for i, h in enumerate(headers):
            table.rows[0].cells[i].text = h
        for _, row in df.iterrows():
            cells = table.add_row().cells
            cells[0].text, cells[1].text = str(row['Ticker']), f"{row['% Change']:.2f}%"
            cells[2].text, cells[3].text = f"{row['Close Price']:.2f}", f"{row['Naira Change']:.2f}"
    
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 4. STREAMLIT INTERFACE ---
st.set_page_config(page_title="NGX Summary Tool", page_icon="🇳🇬")
st.title("NGX Market Report Generator")

r_date, r_msg = get_report_info()
date_str = format_ordinal(r_date)

st.info(f"**Active Report:** {date_str}")
st.caption(f"Status: {r_msg}")

if st.button("Fetch Today's Market Data"):
    with st.spinner("Bypassing NGX firewall..."):
        adv, dec = fetch_ngx_data()
    
    if adv is not None:
        st.success("Success! Data retrieved.")
        fname = f"DAILY EQUITY SUMMARY FOR {date_str}"
        
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("Download Excel", create_excel(adv, dec, date_str), f"{fname}.xlsx")
        with c2:
            st.download_button("Download Word", create_word(adv, dec, date_str), f"{fname}.docx")
        
        st.divider()
        st.subheader("Data Preview")
        st.table(adv)
    else:
        st.error("NGX is still blocking requests. Please click 'Fetch' again in 60 seconds.")
