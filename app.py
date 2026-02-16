import streamlit as st
import pandas as pd
import requests
import io
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime, time, timedelta
import pytz

# --- 1. SMART DATE LOGIC (Lagos Timezone) ---
def get_market_date():
    lagos_tz = pytz.timezone('Africa/Lagos')
    now = datetime.now(lagos_tz)
    weekday = now.weekday()  # 0=Mon, 4=Fri, 5=Sat, 6=Sun
    cutoff_time = time(14, 40) # 2:40 PM

    if weekday == 5: # Saturday -> Show Friday
        report_date = now - timedelta(days=1)
    elif weekday == 6: # Sunday -> Show Friday
        report_date = now - timedelta(days=2)
    elif now.time() < cutoff_time:
        # Before 2:40 PM: Show previous trading day
        # If Monday morning, go back 3 days to Friday
        report_date = now - timedelta(days=3 if weekday == 0 else 1)
    else:
        # After 2:40 PM: Show today
        report_date = now
        
    return report_date.strftime("%d %B %Y").upper()

# --- 2. SMART DATA FETCHING (No Chrome/Selenium Needed) ---
def fetch_ngx_data():
    # Source: Kwayisi NGX Mirror (Fast, No Firewall, Reliable)
    url = "https://afx.kwayisi.org/ngx/"
    headers = {"User-Agent": "Mozilla/5.0"}
    
    try:
        response = requests.get(url, headers=headers, timeout=10)
        # Read the first table on the page
        df = pd.read_html(io.StringIO(response.text))[0]
        
        # Mapping Mirror columns to NGX standard names
        # Mirror: Ticker | Name | Price | Change | Gain
        df = df.rename(columns={
            'Ticker': 'Ticker',
            'Price': 'Close Price',
            'Gain': '% Change',
            'Change': 'Naira Change'
        })

        # Clean numeric columns
        for col in ['% Change', 'Close Price', 'Naira Change']:
            df[col] = df[col].astype(str).str.replace('%', '').str.replace(',', '').str.replace('+', '')
            df[col] = pd.to_numeric(df[col], errors='coerce')

        df = df.dropna(subset=['% Change'])
        
        # Get Top 5 Advancers and Decliners
        adv = df.sort_values(by='% Change', ascending=False).head(5)
        dec = df.sort_values(by='% Change', ascending=True).head(5)
        
        return adv, dec
    except Exception as e:
        st.error(f"Smart Fetch Error: {e}")
        return None, None

# --- 3. DOCUMENT GENERATION ---
def create_excel(adv, dec, date_str):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        adv[['Ticker', '% Change', 'Close Price', 'Naira Change']].to_excel(writer, sheet_name='Summary', startrow=2, index=False)
        dec[['Ticker', '% Change', 'Close Price', 'Naira Change']].to_excel(writer, sheet_name='Summary', startrow=10, index=False)
        ws = writer.sheets['Summary']
        bold = writer.book.add_format({'bold': True, 'align': 'center', 'font_size': 12})
        ws.merge_range('A1:D1', f"DAILY EQUITY SUMMARY FOR {date_str}", bold)
        ws.write('A2', 'TOP 5 ADVANCERS', bold)
        ws.write('A10', 'TOP 5 DECLINERS', bold)
    return output.getvalue()

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
            c[0].text, c[1].text = str(row['Ticker']), f"{row['% Change']:.2f}%"
            c[2].text, c[3].text = f"{row['Close Price']:.2f}", f"{row['Naira Change']:.2f}"
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 4. STREAMLIT INTERFACE ---
st.set_page_config(page_title="NGX Smart Reporter", page_icon="🇳🇬")
st.title("🇳🇬 NGX Market Reporter")

market_date = get_market_date()
st.subheader(f"📅 Market Date: {market_date}")
st.caption("Data logic: Before 2:40pm WAT shows previous day. Weekends show Friday.")

if st.button("🚀 Generate Top 5 Advancers & Decliners"):
    with st.spinner("Analyzing market data..."):
        adv, dec = fetch_ngx_data()
        
        if adv is not None:
            st.success("Data loaded successfully!")
            
            c1, c2 = st.columns(2)
            c1.download_button("📊 Download Excel", create_excel(adv, dec, market_date), f"NGX_Summary_{market_date}.xlsx")
            c2.download_button("📝 Download Word", create_word(adv, dec, market_date), f"NGX_Summary_{market_date}.docx")
            
            st.markdown("---")
            col_a, col_b = st.columns(2)
            with col_a:
                st.write("### 🟢 Top 5 Advancers")
                st.dataframe(adv[['Ticker', '% Change', 'Close Price']], hide_index=True)
            with col_b:
                st.write("### 🔴 Top 5 Decliners")
                st.dataframe(dec[['Ticker', '% Change', 'Close Price']], hide_index=True)
        else:
            st.error("Market data is currently unreachable. Please try again in 1 minute.")
