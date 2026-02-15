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
    day_of_week = now.weekday()
    cutoff = now.replace(hour=14, minute=40, second=0, microsecond=0)
    
    if day_of_week >= 5: # Sat/Sun
        report_date = now - timedelta(days=(day_of_week - 4))
        msg = "Weekend Mode: Pulling Friday's Closing Data"
    elif now < cutoff:
        days_back = 3 if day_of_week == 0 else 1
        report_date = now - timedelta(days=days_back)
        msg = f"Market Open: Previewing {format_ordinal(report_date)} Data"
    else:
        report_date = now
        msg = "Market Closed: Today's Summary Ready"
    return report_date, msg

# --- 2. THE DYNAMIC DATA ENGINE ---
@st.cache_data(ttl=600)
def fetch_ngx_full_market():
    # This is the actual endpoint that feeds the NGX table.
    # It bypasses the empty HTML shell you see in the browser source.
    url = "https://ngxgroup.com/exchange/data/equities-price-list/"
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Referer": "https://ngxgroup.com/exchange/data/equities-price-list/",
        "X-Requested-With": "XMLHttpRequest" 
    }

    try:
        session = requests.Session()
        # Visit the main page first to get security cookies
        session.get("https://ngxgroup.com/", headers=headers, timeout=10)
        
        # Now hit the price list
        response = session.get(url, headers=headers, timeout=20)
        
        # Use pandas to extract the table from the raw HTML response
        tables = pd.read_html(io.StringIO(response.text))
        
        # Find the equities table (usually the largest one)
        df = max(tables, key=len)
        
        # Standardize headers (NGX uses: Symbol, Current, Change, % Change)
        df = df.rename(columns={
            'Symbol': 'Ticker', 
            'Current': 'Close Price', 
            'Change': 'Naira Change'
        }, errors='ignore')

        # CLEANING: Remove commas and % signs so we can sort by value
        for col in ['% Change', 'Close Price', 'Naira Change']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col].astype(str).str.replace('%', '').str.replace(',', ''), errors='coerce')

        # Remove any empty or non-trading rows
        df = df.dropna(subset=['% Change'])

        # Get the REAL movers across the entire market
        advancers = df.sort_values(by='% Change', ascending=False).head(5)
        decliners = df.sort_values(by='% Change', ascending=True).head(5)
        
        return advancers[['Ticker', '% Change', 'Close Price', 'Naira Change']], \
               decliners[['Ticker', '% Change', 'Close Price', 'Naira Change']]

    except Exception as e:
        st.error(f"Automation failed to read the table: {e}")
        return None, None

# --- 3. EXPORT LOGIC ---
def create_excel(adv, dec, date_str):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        adv.rename(columns={'Ticker': 'Gainers'}).to_excel(writer, sheet_name='Summary', startrow=2, index=False)
        dec.rename(columns={'Ticker': 'Decliners'}).to_excel(writer, sheet_name='Summary', startrow=10, index=False)
        ws = writer.sheets['Summary']
        fmt = writer.book.add_format({'bold': True, 'align': 'center', 'font_size': 14})
        ws.merge_range('A1:D1', f"DAILY EQUITY SUMMARY FOR {date_str}", fmt)
    return output.getvalue()

def create_word(adv, dec, date_str):
    doc = Document()
    title = doc.add_heading(f"DAILY EQUITY SUMMARY FOR {date_str}", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for df, label in [(adv, "Top Gainers"), (dec, "Top Decliners")]:
        doc.add_heading(label, level=2)
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        hdrs = [label, "% Change", "Close Price", "Naira Change"]
        for i, h in enumerate(hdrs): table.rows[0].cells[i].text = h
        for _, row in df.iterrows():
            cells = table.add_row().cells
            cells[0].text, cells[1].text = str(row['Ticker']), f"{row['% Change']:.2f}%"
            cells[2].text, cells[3].text = f"{row['Close Price']:.2f}", f"{row['Naira Change']:.2f}"
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 4. STREAMLIT UI ---
st.set_page_config(page_title="NGX Full Market Tool", page_icon="🇳🇬")
st.title("📈 NGX Automatic Reporter")

r_date, r_msg = get_report_info()
date_header = format_ordinal(r_date)

st.info(f"📅 **Target Date:** {date_header}")
st.caption(r_msg)

if st.button("Download Today's Reports"):
    with st.spinner("Accessing Full Market List..."):
        adv, dec = fetch_ngx_full_market()
    
    if adv is not None:
        st.success("Full Market extracted successfully!")
        fn = f"DAILY EQUITY SUMMARY FOR {date_header}"
        st.download_button("📊 Download Excel", create_excel(adv, dec, date_header), f"{fn}.xlsx")
        st.download_button("📝 Download Word", create_word(adv, dec, date_header), f"{fn}.docx")
        
        st.subheader("Preview: Top 5 Gainers")
        st.table(adv)
    else:
        st.error("The NGX website is currently blocking automated access. Please wait 60 seconds and try once more.")
