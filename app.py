import streamlit as st
import pandas as pd
import cloudscraper
import io
import requests
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime, time, timedelta
import pytz

# --- 1. SMART DATE LOGIC (Lagos Timezone) ---
def get_report_info():
    lagos_tz = pytz.timezone('Africa/Lagos')
    now = datetime.now(lagos_tz)
    weekday = now.weekday()  # 0=Mon, 6=Sun
    cutoff_time = time(14, 40) # 2:40 PM
    
    if weekday == 5: # Saturday -> Friday
        target_date = now - timedelta(days=1)
    elif weekday == 6: # Sunday -> Friday
        target_date = now - timedelta(days=2)
    elif now.time() < cutoff_time:
        # Before 2:40 PM: Show previous trading day
        target_date = now - timedelta(days=3 if weekday == 0 else 1)
    else:
        target_date = now
        
    return target_date.strftime("%d %B %Y").upper()

# --- 2. SMART DATA FETCHING ---
def fetch_ngx_data():
    # Attempt 1: Using Cloudscraper to bypass NGX Firewall
    scraper = cloudscraper.create_scraper(browser={'browser': 'chrome', 'platform': 'windows', 'desktop': True})
    
    try:
        # Targeting the AJAX data feed NGX uses
        url = "https://ngxgroup.com/wp-admin/admin-ajax.php?action=get_wdtable&table_id=2"
        response = scraper.get(url, timeout=15)
        
        if response.status_code == 200:
            data = response.json()
            df = pd.DataFrame(data['data'])
            df = df.rename(columns={
                'symbol': 'Ticker',
                'current': 'Close Price',
                'pchange': '% Change',
                'change': 'Naira Change'
            })
        else:
            raise Exception("Main site blocked")

    except Exception:
        # Attempt 2: Backup Mirror (Always works if main site is down)
        try:
            backup_url = "https://afx.kwayisi.org/ngx/"
            df_list = pd.read_html(backup_url)
            df = df_list[0]
            df = df.rename(columns={'Ticker': 'Ticker', 'Price': 'Close Price', 'Gain': '% Change', 'Change': 'Naira Change'})
        except Exception as e:
            st.error(f"Critical Error: {e}")
            return None, None

    # Clean numeric data
    for col in ['% Change', 'Close Price', 'Naira Change']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace('%', '').str.replace(',', '').str.replace('+', '').strip(), errors='coerce')

    df = df.dropna(subset=['Ticker', '% Change'])
    
    # Get Top 5 Advancers and Decliners
    adv = df.sort_values(by='% Change', ascending=False).head(5)
    dec = df.sort_values(by='% Change', ascending=True).head(5)
    
    return adv, dec

# --- 3. EXCEL/WORD GENERATORS ---
def create_excel(adv, dec, date_str):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        adv[['Ticker', '% Change', 'Close Price', 'Naira Change']].to_excel(writer, sheet_name='Summary', startrow=2, index=False)
        dec[['Ticker', '% Change', 'Close Price', 'Naira Change']].to_excel(writer, sheet_name='Summary', startrow=10, index=False)
        ws = writer.sheets['Summary']
        bold = writer.book.add_format({'bold': True, 'align': 'center'})
        ws.merge_range('A1:D1', f"DAILY EQUITY SUMMARY FOR {date_str}", bold)
        ws.write('A2', 'TOP 5 ADVANCERS', bold)
        ws.write('A10', 'TOP 5 DECLINERS', bold)
    return output.getvalue()

def create_word(adv, dec, date_str):
    doc = Document()
    doc.add_heading(f"DAILY EQUITY SUMMARY FOR {date_str}", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    for df, label in [(adv, "Top 5 Advancers"), (dec, "Top 5 Decliners")]:
        doc.add_heading(label, 2)
        table = doc.add_table(rows=1, cols=4); table.style = 'Table Grid'
        hdrs = ["Ticker", "% Change", "Close Price", "Naira Change"]
        for i, h in enumerate(hdrs): table.rows[0].cells[i].text = h
        for _, row in df.iterrows():
            c = table.add_row().cells
            c[0].text, c[1].text = str(row['Ticker']), f"{row['% Change']:.2f}%"
            c[2].text, c[3].text = f"{row['Close Price']:.2f}", f"{row['Naira Change']:.2f}"
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 4. STREAMLIT UI ---
st.set_page_config(page_title="NGX Reporter", page_icon="🇳🇬")
st.title("📈 NGX Market Reporter")

report_date = get_report_info()
st.subheader(f"📅 Report Date: {report_date}")

if st.button("🚀 Fetch Top 5 Advancers & Decliners"):
    with st.spinner("Fetching market data..."):
        adv, dec = fetch_ngx_data()
        
        if adv is not None and not adv.empty:
            st.success("Data Captured!")
            c1, c2 = st.columns(2)
            c1.download_button("📊 Excel", create_excel(adv, dec, report_date), f"NGX_{report_date}.xlsx")
            c2.download_button("📝 Word", create_word(adv, dec, report_date), f"NGX_{report_date}.docx")
            
            st.write("### 🟢 Top 5 Advancers")
            st.table(adv[['Ticker', '% Change', 'Close Price']])
            st.write("### 🔴 Top 5 Decliners")
            st.table(dec[['Ticker', '% Change', 'Close Price']])
