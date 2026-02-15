import streamlit as st
import pandas as pd
import requests
import io
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime, timedelta
import pytz

# --- 1. TIME & DATE LOGIC ---
def get_lagos_time():
    return datetime.now(pytz.timezone('Africa/Lagos'))

def format_ordinal(dt):
    day = dt.day
    suffix = "TH" if 11 <= day <= 13 else {1: "ST", 2: "ND", 3: "RD"}.get(day % 10, "TH")
    return dt.strftime(f"{day}{suffix} %b %Y").upper()

def get_report_date_info():
    now = get_lagos_time()
    day_of_week = now.weekday() # 0=Mon, 5=Sat, 6=Sun
    
    # Define 2:40 PM today
    today_cutoff = now.replace(hour=14, minute=40, second=0, microsecond=0)
    
    # Logic for which date the report represents
    if day_of_week == 5: # Saturday
        report_date = now - timedelta(days=1) # Friday
        status = "Weekend Mode: Showing Friday's Closing Summary"
    elif day_of_week == 6: # Sunday
        report_date = now - timedelta(days=2) # Friday
        status = "Weekend Mode: Showing Friday's Closing Summary"
    elif now < today_cutoff:
        # Before 2:40 PM on a weekday, show previous trading day
        days_to_subtract = 3 if day_of_week == 0 else 1
        report_date = now - timedelta(days=days_to_subtract)
        status = f"Market Open: Showing Summary for {format_ordinal(report_date)}"
    else:
        # After 2:40 PM on a weekday
        report_date = now
        status = "Market Closed: Today's Summary is Ready"
        
    return report_date, status

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
        # We read the table and force the headers you need
        tables = pd.read_html(io.StringIO(response.text))
        df = tables[0]
        
        # Rename columns based on your NGX observation
        df = df.rename(columns={
            'Symbol': 'Ticker', 
            'Current': 'Close Price', 
            'Change': 'Naira Change', 
            '% Change': '% Change'
        }, errors='ignore')

        # Clean numeric data
        for col in ['% Change', 'Close Price', 'Naira Change']:
            df[col] = pd.to_numeric(df[col], errors='coerce')

        adv = df.nlargest(5, '% Change')[['Ticker', '% Change', 'Close Price', 'Naira Change']]
        dec = df.nsmallest(5, '% Change')[['Ticker', '% Change', 'Close Price', 'Naira Change']]
        return adv, dec
    except:
        return None, None

# --- 3. FILE GENERATORS ---
def create_excel(adv, dec, date_str):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        adv.rename(columns={'Ticker': 'Gainers'}).to_excel(writer, sheet_name='Summary', startrow=2, index=False)
        dec.rename(columns={'Ticker': 'Decliners'}).to_excel(writer, sheet_name='Summary', startrow=10, index=False)
        worksheet = writer.sheets['Summary']
        workbook = writer.book
        title_fmt = workbook.add_format({'bold': True, 'align': 'center', 'font_size': 14})
        worksheet.merge_range('A1:D1', f"DAILY EQUITY SUMMARY FOR {date_str}", title_fmt)
    return output.getvalue()

def create_word(adv, dec, date_str):
    doc = Document()
    title = doc.add_heading(f"DAILY EQUITY SUMMARY FOR {date_str}", level=1)
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

report_date, status_msg = get_report_date_info()
date_display = format_ordinal(report_date)

st.info(f"**Report Date:** {date_display}")
st.write(f"{status_msg}")

if st.button("Generate Reports"):
    with st.spinner("Fetching live data..."):
        adv, dec = fetch_ngx_data()
        
    if adv is not None:
        filename = f"DAILY EQUITY SUMMARY FOR {date_display}"
        
        col1, col2 = st.columns(2)
        with col1:
            st.download_button("Excel Format", create_excel(adv, dec, date_display), f"{filename}.xlsx")
        with col2:
            st.download_button("Word Format", create_word(adv, dec, date_display), f"{filename}.docx")
        
        st.divider()
        st.subheader("Preview")
        st.write("**Top 5 Gainers**")
        st.dataframe(adv, hide_index=True)
    else:
        st.error("Could not connect to NGX. Please check your internet or try again.")
