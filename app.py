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
        msg = "Weekend Mode: Friday's Closing Data"
    # If Monday-Friday before 2:40 PM, show previous day
    elif now < cutoff:
        days_back = 3 if day_of_week == 0 else 1
        report_date = now - timedelta(days=days_back)
        msg = f"Market Open: Showing {format_ordinal(report_date)} Data"
    else:
        report_date = now
        msg = "Market Closed: Today's Summary Ready"
    return report_date, msg

# --- 2. THE FULL MARKET AUTOMATION ENGINE ---
@st.cache_data(ttl=600)
def fetch_ngx_full_market():
    # Target URL: The official NGX Equities Price List
    url = "https://ngxgroup.com/exchange/data/equities-price-list/"
    
    # These headers are the "Secret Sauce" to bypass 403 blocks.
    # They mimic a high-end Chrome browser on Windows.
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
        "Referer": "https://ngxgroup.com/",
        "Connection": "keep-alive",
        "Upgrade-Insecure-Requests": "1"
    }

    try:
        # Step 1: Create a session to handle cookies like a real human
        session = requests.Session()
        session.get("https://ngxgroup.com/", headers=headers, timeout=10)
        
        # Step 2: Request the price list page
        response = session.get(url, headers=headers, timeout=20)
        
        if response.status_code != 200:
            st.error(f"NGX Server returned error code: {response.status_code}")
            return None, None

        # Step 3: Parse the HTML tables
        # NGX data is usually in the first table object [0]
        tables = pd.read_html(io.StringIO(response.text))
        df = tables[0]
        
        # Step 4: Map the data to YOUR specific header requirements
        # NGX raw headers: Symbol, Current, Change, % Change
        df = df.rename(columns={
            'Symbol': 'Ticker', 
            'Current': 'Close Price', 
            'Change': 'Naira Change'
        }, errors='ignore')

        # Step 5: DATA CLEANING (Commas and % signs prevent math)
        for col in ['% Change', 'Close Price', 'Naira Change']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col].astype(str).str.replace('%', '').str.replace(',', ''), errors='coerce')

        # Step 6: Identify the REAL Movers from the FULL market list
        # We sort by % Change to find the actual top 5 gainers and losers
        full_sorted = df.dropna(subset=['% Change']).sort_values(by='% Change', ascending=False)
        
        advancers = full_sorted.head(5)[['Ticker', '% Change', 'Close Price', 'Naira Change']]
        decliners = full_sorted.tail(5).iloc[::-1][['Ticker', '% Change', 'Close Price', 'Naira Change']]
        
        return advancers, decliners

    except Exception as e:
        st.error(f"Automation Refused: {e}")
        return None, None

# --- 3. DOCUMENT GENERATORS ---
def create_excel(adv, dec, date_str):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        adv.rename(columns={'Ticker': 'Gainers'}).to_excel(writer, sheet_name='Equity Summary', startrow=2, index=False)
        dec.rename(columns={'Ticker': 'Decliners'}).to_excel(writer, sheet_name='Equity Summary', startrow=10, index=False)
        
        ws = writer.sheets['Equity Summary']
        wb = writer.book
        title_fmt = wb.add_format({'bold': True, 'align': 'center', 'font_size': 14, 'font_color': '#1E40AF'})
        ws.merge_range('A1:D1', f"DAILY EQUITY SUMMARY FOR {date_str}", title_fmt)
        
        # Auto-adjust column width for readability
        for i, col in enumerate(adv.columns):
            ws.set_column(i, i, 15)
            
    return output.getvalue()

def create_word(adv, dec, date_str):
    doc = Document()
    title = doc.add_heading(f"DAILY EQUITY SUMMARY FOR {date_str}", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    for df, label in [(adv, "Top 5 Gainers"), (dec, "Top 5 Decliners")]:
        doc.add_heading(label, level=2)
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        hdr_cells = table.rows[0].cells
        hdrs = ["Ticker", "% Change", "Close Price", "Naira Change"]
        for i, h in enumerate(hdrs):
            hdr_cells[i].text = h
            
        for _, row in df.iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = str(row['Ticker'])
            row_cells[1].text = f"{row['% Change']:.2f}%"
            row_cells[2].text = f"{row['Close Price']:.2f}"
            row_cells[3].text = f"{row['Naira Change']:.2f}"
            
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 4. STREAMLIT INTERFACE ---
st.set_page_config(page_title="NGX Reporter PRO", page_icon="📈", layout="centered")

st.title("🇳🇬 NGX Daily Market Reporter")
st.markdown("---")

r_date, r_msg = get_report_info()
date_header = format_ordinal(r_date)

st.subheader(f"Report for: {date_header}")
st.info(f"💡 {r_msg}")

if st.button("🚀 Fetch & Process Full Market Data", use_container_width=True):
    with st.spinner("Extracting dynamic data from NGX portal..."):
        adv, dec = fetch_ngx_full_market()
        
    if adv is not None:
        st.success("Full Market Scan Complete!")
        
        # Layout for downloads
        col1, col2 = st.columns(2)
        filename = f"DAILY EQUITY SUMMARY FOR {date_header}"
        
        with col1:
            st.download_button("📊 Download Excel", create_excel(adv, dec, date_header), f"{filename}.xlsx", use_container_width=True)
        with col2:
            st.download_button("📝 Download Word", create_word(adv, dec, date_header), f"{filename}.docx", use_container_width=True)
            
        st.markdown("### Data Preview")
        tab1, tab2 = st.tabs(["🔥 Top Gainers", "❄️ Top Decliners"])
        with tab1:
            st.table(adv)
        with tab2:
            st.table(dec)
    else:
        st.error("The automated script was blocked or the data hasn't been published yet. Please try again.")

st.markdown("---")
st.caption("This tool scans the complete NGX price list and automatically selects the top 5 movers.")
