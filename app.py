import streamlit as st
import pandas as pd
import requests
import io
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime, time, timedelta
import pytz

# --- 1. DATE & TIME LOGIC (Lagos Timezone) ---
def get_market_date_and_status():
    lagos_tz = pytz.timezone('Africa/Lagos')
    now = datetime.now(lagos_tz)
    weekday = now.weekday()  # 0=Mon, 4=Fri, 5=Sat, 6=Sun
    cutoff_time = time(14, 40) # 2:40 PM
    
    # Logic for selecting the correct report date
    if weekday >= 5: # Weekend (Sat/Sun) -> Always Friday
        days_back = 1 if weekday == 5 else 2
        report_date = now - timedelta(days=days_back)
    elif now.time() < cutoff_time: # Weekday before 2:40 PM -> Previous trading day
        days_back = 3 if weekday == 0 else 1 # If Mon, go to Fri
        report_date = now - timedelta(days=days_back)
    else: # Weekday after 2:40 PM -> Today
        report_date = now
        
    date_str = report_date.strftime("%d %B %Y").upper()
    return date_str

# --- 2. DATA FETCHING (Firewall-Safe Method) ---
@st.cache_data(ttl=600) # Cache data for 10 minutes
def fetch_ngx_data():
    # We use the direct API endpoint that NGX uses to populate their own library
    # This is often less protected by the Sucuri WAF than the main homepage
    url = "https://doclib.ngxgroup.com/REST/api/statistics/equities/?market=&sector=&orderby=&pageSize=300&pageNo=0"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36"
    }
    
    try:
        response = requests.get(url, headers=headers, timeout=20)
        # If the direct API is blocked, we use a reliable 3rd party fallback
        if response.status_code != 200:
            fallback_url = "https://afx.kwayisi.org/ngx/"
            df_list = pd.read_html(fallback_url)
            df = df_list[0]
            # Standardize columns for the fallback
            df = df.rename(columns={'Ticker': 'Ticker', 'Price': 'Close Price', 'Change': 'Naira Change', 'Gain': '% Change'})
        else:
            data = response.json()
            df = pd.DataFrame(data)
            # Standardize columns for the primary API
            df = df.rename(columns={
                'symbol': 'Ticker',
                'close': 'Close Price',
                'changePercentage': '% Change',
                'change': 'Naira Change'
            })

        # Cleaning numeric columns
        for col in ['% Change', 'Close Price', 'Naira Change']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col].astype(str).str.replace('%', '').str.replace(',', '').str.strip(), errors='coerce')

        df = df.dropna(subset=['% Change'])
        
        # Get Top 5 Advancers (Gainers) and Decliners (Losers)
        adv = df.sort_values(by='% Change', ascending=False).head(5)
        dec = df.sort_values(by='% Change', ascending=True).head(5)
        
        return adv, dec
    except Exception as e:
        st.error(f"⚠️ Service Busy. Please try again in a few minutes. (Detail: {e})")
        return None, None

# --- 3. DOCUMENT GENERATION ---
def create_excel(adv, dec, date_str):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        bold_fmt = writer.book.add_format({'bold': True, 'align': 'center', 'font_size': 12})
        adv[['Ticker', '% Change', 'Close Price', 'Naira Change']].to_excel(writer, sheet_name='Summary', startrow=2, index=False)
        dec[['Ticker', '% Change', 'Close Price', 'Naira Change']].to_excel(writer, sheet_name='Summary', startrow=10, index=False)
        ws = writer.sheets['Summary']
        ws.merge_range('A1:D1', f"DAILY EQUITY SUMMARY FOR {date_str}", bold_fmt)
        ws.write('A2', 'TOP 5 ADVANCERS', bold_fmt)
        ws.write('A10', 'TOP 5 DECLINERS', bold_fmt)
    return output.getvalue()

def create_word(adv, dec, date_str):
    doc = Document()
    doc.add_heading(f"DAILY EQUITY SUMMARY FOR {date_str}", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    for df, label in [(adv, "Top 5 Advancers"), (dec, "Top 5 Decliners")]:
        doc.add_heading(label, 2)
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        for i, h in enumerate([label, "% Change", "Close Price", "Naira Change"]):
            table.rows[0].cells[i].text = h
        for _, row in df.iterrows():
            c = table.add_row().cells
            c[0].text, c[1].text = str(row['Ticker']), f"{row['% Change']:.2f}%"
            c[2].text, c[3].text = f"{row['Close Price']:.2f}", f"{row['Naira Change']:.2f}"
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 4. STREAMLIT UI ---
st.set_page_config(page_title="NGX Live Reporter", page_icon="🇳🇬")
st.title("📈 NGX Market Report")

report_date = get_market_date_and_status()
st.subheader(f"📅 Report for: {report_date}")
st.caption("Data is collected daily at 2:40 PM WAT. Weekends reflect the last Friday.")

if st.button("🚀 Generate Latest Reports"):
    with st.spinner("Fetching market data..."):
        adv, dec = fetch_ngx_data()
        
        if adv is not None and not adv.empty:
            st.success("Data loaded successfully!")
            
            c1, c2 = st.columns(2)
            c1.download_button("📊 Download Excel", create_excel(adv, dec, report_date), f"NGX_Report_{report_date}.xlsx")
            c2.download_button("📝 Download Word", create_word(adv, dec, report_date), f"NGX_Report_{report_date}.docx")
            
            st.markdown("---")
            st.write("### 🟢 Top 5 Advancers")
            st.dataframe(adv[['Ticker', '% Change', 'Close Price', 'Naira Change']], hide_index=True)
            
            st.write("### 🔴 Top 5 Decliners")
            st.dataframe(dec[['Ticker', '% Change', 'Close Price', 'Naira Change']], hide_index=True)
        else:
            st.warning("Could not retrieve data at this moment. The NGX servers might be down.")
