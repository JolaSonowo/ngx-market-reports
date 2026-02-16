import streamlit as st
import pandas as pd
import requests
import io
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime, time, timedelta
import pytz

# --- 1. DATE LOGIC (Lagos Time, 2:40 PM Cutoff) ---
def get_market_date():
    lagos_tz = pytz.timezone('Africa/Lagos')
    now = datetime.now(lagos_tz)
    weekday = now.weekday()  # 0=Mon, 4=Fri, 5=Sat, 6=Sun
    current_time = now.time()
    cutoff_time = time(14, 40) # 2:40 PM

    if weekday == 5: # Saturday -> Show Friday
        report_date = now - timedelta(days=1)
    elif weekday == 6: # Sunday -> Show Friday
        report_date = now - timedelta(days=2)
    elif current_time < cutoff_time:
        if weekday == 0: # Monday morning -> Show Friday
            report_date = now - timedelta(days=3)
        else: # Weekday morning -> Show Yesterday
            report_date = now - timedelta(days=1)
    else: # Weekday after 2:40 PM -> Show Today
        report_date = now
        
    return report_date.strftime("%d %B %Y").upper()

# --- 2. DIRECT DATA FETCHING (Firewall-Safe & Fast) ---
def fetch_ngx_data():
    # Direct API used by NGX for their document library statistics
    url = "https://doclib.ngxgroup.com/REST/api/statistics/equities/?market=&sector=&orderby=&pageSize=300&pageNo=0"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
        "Accept": "application/json"
    }
    
    try:
        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()
        data = response.json()
        
        # Create DataFrame from JSON
        df = pd.DataFrame(data)
        
        # Map the internal API keys to your required names
        # 'symbol' -> Ticker
        # 'close' -> Close Price
        # 'changePercentage' -> % Change
        # 'change' -> Naira Change
        df = df[['symbol', 'changePercentage', 'close', 'change']].copy()
        df.columns = ['Ticker', '% Change', 'Close Price', 'Naira Change']

        # Ensure columns are numeric for sorting
        for col in ['% Change', 'Close Price', 'Naira Change']:
            df[col] = pd.to_numeric(df[col], errors='coerce')

        df = df.dropna(subset=['% Change'])
        
        # Get Top 5 Advancers and Decliners
        adv = df.sort_values(by='% Change', ascending=False).head(5)
        dec = df.sort_values(by='% Change', ascending=True).head(5)
        
        return adv, dec
    except Exception as e:
        st.error(f"⚠️ NGX Feed unavailable: {e}")
        return None, None

# --- 3. EXPORT GENERATION ---
def create_excel(adv, dec, date_str):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        adv.to_excel(writer, sheet_name='Summary', startrow=2, index=False)
        dec.to_excel(writer, sheet_name='Summary', startrow=10, index=False)
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
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        hdrs = ["Ticker", "% Change", "Close Price", "Naira Change"]
        for i, h in enumerate(hdrs):
            table.rows[0].cells[i].text = h
        for _, row in df.iterrows():
            c = table.add_row().cells
            c[0].text = str(row['Ticker'])
            c[1].text = f"{row['% Change']:.2f}%"
            c[2].text = f"{row['Close Price']:.2f}"
            c[3].text = f"{row['Naira Change']:.2f}"
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 4. STREAMLIT INTERFACE ---
st.set_page_config(page_title="NGX Reporter", page_icon="🇳🇬")
st.title("🇳🇬 NGX Market Reporter")

market_date = get_market_date()
st.info(f"Generating data for: **{market_date}**")

if st.button("🚀 Fetch Top 5 Advancers & Decliners"):
    with st.spinner("Connecting to NGX Data Feed..."):
        adv, dec = fetch_ngx_data()
        
        if adv is not None and not adv.empty:
            st.success("Market Data Captured!")
            
            c1, c2 = st.columns(2)
            c1.download_button("📊 Excel Report", create_excel(adv, dec, market_date), f"NGX_Summary_{market_date}.xlsx")
            c2.download_button("📝 Word Report", create_word(adv, dec, market_date), f"NGX_Summary_{market_date}.docx")
            
            st.subheader("🟢 Top 5 Advancers")
            st.table(adv[['Ticker', '% Change', 'Close Price', 'Naira Change']])
            
            st.subheader("🔴 Top 5 Decliners")
            st.table(dec[['Ticker', '% Change', 'Close Price', 'Naira Change']])
        else:
            st.warning("Data currently unavailable. This usually happens if the NGX servers are undergoing daily maintenance. Please try again in 10 minutes.")
