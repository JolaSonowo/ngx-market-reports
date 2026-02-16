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
    weekday = now.weekday()  # 0=Mon, 6=Sun
    cutoff_time = time(14, 40) # 2:40 PM

    # Weekend logic
    if weekday == 5: # Sat
        report_date = now - timedelta(days=1)
    elif weekday == 6: # Sun
        report_date = now - timedelta(days=2)
    # Weekday logic
    elif now.time() < cutoff_time:
        # If Monday morning, go back to Friday
        report_date = now - timedelta(days=3 if weekday == 0 else 1)
    else:
        report_date = now
        
    return report_date.strftime("%d %B %Y").upper()

# --- 2. THE "SMART" DATA FETCH (Using official API) ---
def fetch_ngx_data():
    # Direct Official API endpoint - usually bypasses Sucuri Firewall
    url = "https://doclib.ngxgroup.com/REST/api/statistics/equities/?market=&sector=&orderby=&pageSize=300&pageNo=0"
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
        "Accept": "application/json, text/plain, */*",
        "Referer": "https://ngxgroup.com/"
    }
    
    try:
        # Use a session to keep connection alive and bypass simple bot checks
        session = requests.Session()
        response = session.get(url, headers=headers, timeout=15)
        response.raise_for_status()
        data = response.json()
        
        # API returns a list of dictionaries
        df = pd.DataFrame(data)
        
        # Map the API keys to your required headers
        # symbol -> Ticker
        # close -> Close Price
        # changePercentage -> % Change
        # change -> Naira Change
        df = df[['symbol', 'changePercentage', 'close', 'change']].copy()
        df.columns = ['Ticker', '% Change', 'Close Price', 'Naira Change']

        # Convert to numeric
        for col in ['% Change', 'Close Price', 'Naira Change']:
            df[col] = pd.to_numeric(df[col], errors='coerce')

        df = df.dropna(subset=['% Change'])
        
        # Get Top 5 Advancers and Decliners
        adv = df.sort_values(by='% Change', ascending=False).head(5)
        dec = df.sort_values(by='% Change', ascending=True).head(5)
        
        return adv, dec

    except Exception as e:
        st.error(f"⚠️ Market Data Error: The NGX Data Feed is currently blocking the connection. Try again in 2 minutes. (Details: {e})")
        return None, None

# --- 3. EXCEL/WORD GENERATION ---
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
    title = doc.add_heading(f"DAILY EQUITY SUMMARY FOR {date_str}", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for df, label in [(adv, "Top 5 Advancers"), (dec, "Top 5 Decliners")]:
        doc.add_heading(label, 2)
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
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

market_date = get_market_date()
st.info(f"Report for: **{market_date}**")

if st.button("🚀 Fetch Top 5 Advancers & Decliners"):
    with st.spinner("Connecting to Official NGX Data Feed..."):
        adv, dec = fetch_ngx_data()
        
        if adv is not None:
            st.success("Data Captured Successfully!")
            c1, c2 = st.columns(2)
            c1.download_button("📊 Excel", create_excel(adv, dec, market_date), f"NGX_{market_date}.xlsx")
            c2.download_button("📝 Word", create_word(adv, dec, market_date), f"NGX_{market_date}.docx")
            
            st.subheader("🟢 Top 5 Advancers")
            st.table(adv)
            st.subheader("🔴 Top 5 Decliners")
            st.table(dec)
