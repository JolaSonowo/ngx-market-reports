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
    weekday = now.weekday() 
    current_time = now.time()
    cutoff_time = time(14, 40) 

    if weekday == 5: # Sat
        report_date = now - timedelta(days=1)
    elif weekday == 6: # Sun
        report_date = now - timedelta(days=2)
    elif current_time < cutoff_time:
        report_date = now - timedelta(days=3 if weekday == 0 else 1)
    else:
        report_date = now
    return report_date.strftime("%d %B %Y").upper()

# --- 2. ROBUST DATA FETCHING ---
def fetch_ngx_data():
    url = "https://doclib.ngxgroup.com/REST/api/statistics/equities/?market=&sector=&orderby=&pageSize=300&pageNo=0"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36"
    }
    
    try:
        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()
        data = response.json()
        
        # API can return a list or a dict containing a list
        if isinstance(data, dict) and 'data' in data:
            df = pd.DataFrame(data['data'])
        elif isinstance(data, list):
            df = pd.DataFrame(data)
        else:
            df = pd.DataFrame(data)

        # --- FUZZY COLUMN MAPPING ---
        # We look for keywords in column names to avoid "KeyErrors"
        col_map = {}
        for col in df.columns:
            c_low = col.lower()
            if 'symbol' in c_low or 'ticker' in c_low:
                col_map[col] = 'Ticker'
            elif 'percentage' in c_low or 'pchange' in c_low or 'pct' in c_low:
                col_map[col] = '% Change'
            elif ('close' in c_low or 'price' in c_low or 'current' in c_low) and 'change' not in c_low:
                col_map[col] = 'Close Price'
            elif 'change' in c_low and 'percentage' not in c_low and 'pct' not in c_low:
                col_map[col] = 'Naira Change'

        # Safety check: Did we find the required columns?
        required = ['Ticker', '% Change', 'Close Price', 'Naira Change']
        if not all(k in col_map.values() for k in required):
            # If API fails, try Web Scraping as Fallback
            st.info("API structure changed, trying fallback method...")
            df_fallback = pd.read_html("https://afx.kwayisi.org/ngx/")[0]
            df_fallback.columns = ['Ticker', 'Name', 'Close Price', 'Naira Change', '% Change']
            df = df_fallback
        else:
            df = df.rename(columns=col_map)
            df = df[required]

        # Clean numeric data
        for col in ['% Change', 'Close Price', 'Naira Change']:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace('%', '').str.replace(',', '').str.strip(), errors='coerce')

        df = df.dropna(subset=['% Change'])
        adv = df.sort_values(by='% Change', ascending=False).head(5)
        dec = df.sort_values(by='% Change', ascending=True).head(5)
        
        return adv, dec
    except Exception as e:
        st.error(f"⚠️ Error: {e}")
        return None, None

# --- 3. EXPORTS ---
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
        table = doc.add_table(rows=1, cols=4); table.style = 'Table Grid'
        headers = ["Ticker", "% Change", "Close Price", "Naira Change"]
        for i, h in enumerate(headers): table.rows[0].cells[i].text = h
        for _, row in df.iterrows():
            c = table.add_row().cells
            c[0].text, c[1].text = str(row['Ticker']), f"{row['% Change']:.2f}%"
            c[2].text, c[3].text = f"{row['Close Price']:.2f}", f"{row['Naira Change']:.2f}"
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 4. UI ---
st.set_page_config(page_title="NGX Reporter", page_icon="🇳🇬")
st.title("🇳🇬 NGX Market Report")

market_date = get_market_date()
st.info(f"Report Date: **{market_date}**")

if st.button("🚀 Fetch Top 5 Advancers & Decliners"):
    with st.spinner("Fetching market data..."):
        adv, dec = fetch_ngx_data()
        if adv is not None and not adv.empty:
            st.success("Data Captured!")
            c1, c2 = st.columns(2)
            c1.download_button("📊 Excel", create_excel(adv, dec, market_date), f"NGX_{market_date}.xlsx")
            c2.download_button("📝 Word", create_word(adv, dec, market_date), f"NGX_{market_date}.docx")
            
            st.subheader("🟢 Top 5 Advancers")
            st.table(adv[['Ticker', '% Change', 'Close Price', 'Naira Change']])
            st.subheader("🔴 Top 5 Decliners")
            st.table(dec[['Ticker', '% Change', 'Close Price', 'Naira Change']])
        else:
            st.warning("Could not find data. The market feed might be down for maintenance.")
