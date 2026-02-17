import streamlit as st
import pandas as pd
import requests
import io
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime, time, timedelta
import pytz

# --- 1. DATE LOGIC (2:40 PM LAGOS CUTOFF) ---
def get_market_date():
    lagos_tz = pytz.timezone('Africa/Lagos')
    now = datetime.now(lagos_tz)
    weekday = now.weekday()  # 0=Mon, 4=Fri, 5=Sat, 6=Sun
    cutoff = time(14, 40)

    # Weekend logic (Sat/Sun) -> Always show Friday
    if weekday == 5:
        report_date = now - timedelta(days=1)
    elif weekday == 6:
        report_date = now - timedelta(days=2)
    # Weekday logic
    elif now.time() < cutoff:
        # If Monday morning, go back 3 days to Friday
        if weekday == 0:
            report_date = now - timedelta(days=3)
        else:
            report_date = now - timedelta(days=1)
    else:
        # After 2:40 PM, show today
        report_date = now
        
    return report_date.strftime("%d %B %Y").upper()

# --- 2. DATA FETCHING (FIREWALL-SAFE) ---
def fetch_data():
    # Using the AFX mirror which provides the exact NGX data without the Sucuri block
    url = "https://afx.kwayisi.org/ngx/"
    headers = {"User-Agent": "Mozilla/5.0"}
    
    try:
        response = requests.get(url, headers=headers, timeout=15)
        # Read the first table on the page
        df = pd.read_html(io.StringIO(response.text))[0]
        
        # Clean the data to match your requirements
        # Mirror uses: Ticker, Name, Price, Change, Gain
        df = df.rename(columns={
            'Ticker': 'Ticker',
            'Price': 'Close Price',
            'Gain': '% Change',
            'Change': 'Naira Change'
        })

        # Convert strings to actual numbers
        for col in ['% Change', 'Close Price', 'Naira Change']:
            df[col] = df[col].astype(str).str.replace('%', '').str.replace(',', '').str.replace('+', '')
            df[col] = pd.to_numeric(df[col], errors='coerce')

        df = df.dropna(subset=['% Change'])
        
        # Get Top 5 Advancers and Decliners
        adv = df.sort_values(by='% Change', ascending=False).head(5)
        dec = df.sort_values(by='% Change', ascending=True).head(5)
        
        return adv, dec
    except Exception as e:
        st.error(f"Error: {e}")
        return None, None

# --- 3. EXPORT TO EXCEL ---
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

# --- 4. EXPORT TO WORD ---
def create_word(adv, dec, date_str):
    doc = Document()
    title = doc.add_heading(f"DAILY EQUITY SUMMARY FOR {date_str}", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    for df, label in [(adv, "Top 5 Advancers"), (dec, "Top 5 Decliners")]:
        doc.add_heading(label, level=2)
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        hdr = table.rows[0].cells
        hdr[0].text, hdr[1].text, hdr[2].text, hdr[3].text = "Ticker", "% Change", "Close Price", "Naira Change"
        for _, row in df.iterrows():
            c = table.add_row().cells
            c[0].text = str(row['Ticker'])
            c[1].text = f"{row['% Change']:.2f}%"
            c[2].text = f"{row['Close Price']:.2f}"
            c[3].text = f"{row['Naira Change']:.2f}"
            
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 5. STREAMLIT INTERFACE ---
st.set_page_config(page_title="NGX Reporter", page_icon="🇳🇬")
st.title("🇳🇬 NGX Market Report")

market_date = get_market_date()
st.subheader(f"📅 Report Date: {market_date}")

if st.button("🚀 Generate Reports"):
    with st.spinner("Fetching market data..."):
        adv, dec = fetch_data()
        
        if adv is not None:
            st.success("Data loaded!")
            
            c1, c2 = st.columns(2)
            c1.download_button("📊 Download Excel", create_excel(adv, dec, market_date), f"NGX_{market_date}.xlsx")
            c2.download_button("📝 Download Word", create_word(adv, dec, market_date), f"NGX_{market_date}.docx")
            
            st.write("### 🟢 Top 5 Advancers")
            st.table(adv[['Ticker', '% Change', 'Close Price']])
            
            st.write("### 🔴 Top 5 Decliners")
            st.table(dec[['Ticker', '% Change', 'Close Price']])
