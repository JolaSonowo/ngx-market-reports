import streamlit as st
import pandas as pd
import requests
import io
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime, time, timedelta
import pytz

# --- 1. DATE & TIME LOGIC ---
def get_market_date():
    lagos_tz = pytz.timezone('Africa/Lagos')
    now = datetime.now(lagos_tz)
    weekday = now.weekday() 
    cutoff_time = time(14, 40) # 2:40 PM

    if weekday == 5: # Saturday
        report_date = now - timedelta(days=1)
    elif weekday == 6: # Sunday
        report_date = now - timedelta(days=2)
    elif now.time() < cutoff_time:
        # If Monday morning, go back to Friday (3 days)
        report_date = now - timedelta(days=3 if weekday == 0 else 1)
    else:
        report_date = now
        
    return report_date.strftime("%d %B %Y").upper()

# --- 2. ROBUST DATA FETCHING ---
def fetch_ngx_data():
    # Attempt 1: Official NGX Data API
    url = "https://doclib.ngxgroup.com/REST/api/statistics/equities/?market=&sector=&orderby=&pageSize=300&pageNo=0"
    headers = {"User-Agent": "Mozilla/5.0"}
    
    try:
        response = requests.get(url, headers=headers, timeout=15)
        if response.status_code == 200:
            df = pd.DataFrame(response.json())
            # Map API keys to our standard names
            column_map = {
                'symbol': 'Ticker',
                'close': 'Close Price',
                'changePercentage': '% Change',
                'change': 'Naira Change'
            }
            df = df.rename(columns=column_map)
        else:
            # Attempt 2: Stable Fallback Source (Kwayisi)
            df = pd.read_html("https://afx.kwayisi.org/ngx/")[0]
            # Standardize fallback names
            df = df.rename(columns={'Price': 'Close Price', 'Gain': '% Change', 'Change': 'Naira Change'})

        # --- SMART COLUMN CLEANING ---
        # Find the % Change column even if renamed (look for 'Gain' or '%')
        potential_pct_cols = [c for c in df.columns if '%' in c or 'Gain' in str(c) or 'pchange' in str(c).lower()]
        if potential_pct_cols:
            df = df.rename(columns={potential_pct_cols[0]: '% Change'})

        # Clean numeric data
        for col in ['% Change', 'Close Price', 'Naira Change']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col].astype(str).str.replace('%', '').str.replace(',', '').str.strip(), errors='coerce')

        df = df.dropna(subset=['% Change'])
        
        # Get Top 5
        adv = df.sort_values(by='% Change', ascending=False).head(5)
        dec = df.sort_values(by='% Change', ascending=True).head(5)
        
        return adv, dec

    except Exception as e:
        st.error(f"⚠️ Connection Error: {e}")
        return None, None

# --- 3. EXPORT LOGIC ---
def create_excel(adv, dec, date_str):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Format the data for Excel
        adv_final = adv[['Ticker', '% Change', 'Close Price', 'Naira Change']]
        dec_final = dec[['Ticker', '% Change', 'Close Price', 'Naira Change']]
        
        adv_final.to_excel(writer, sheet_name='Summary', startrow=2, index=False)
        dec_final.to_excel(writer, sheet_name='Summary', startrow=10, index=False)
        
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
        for i, h in enumerate([label, "% Change", "Close Price", "Naira Change"]):
            table.rows[0].cells[i].text = h
        for _, row in df.iterrows():
            c = table.add_row().cells
            c[0].text, c[1].text = str(row['Ticker']), f"{row['% Change']:.2f}%"
            c[2].text, c[3].text = f"{row['Close Price']:.2f}", f"{row['Naira Change']:.2f}"
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 4. UI ---
st.set_page_config(page_title="NGX Reporter", page_icon="🇳🇬")
st.title("🇳🇬 NGX Market Reporter")

market_date = get_market_date()
st.info(f"Report Date: **{market_date}**")

if st.button("🚀 Fetch Latest Advancers & Decliners"):
    with st.spinner("Fetching live data..."):
        adv, dec = fetch_ngx_data()
        
        if adv is not None and not adv.empty:
            st.success("Data loaded successfully!")
            
            col1, col2 = st.columns(2)
            col1.download_button("📊 Download Excel", create_excel(adv, dec, market_date), f"NGX_Report_{market_date}.xlsx")
            col2.download_button("📝 Download Word", create_word(adv, dec, market_date), f"NGX_Report_{market_date}.docx")
            
            st.subheader("🟢 Top 5 Advancers")
            st.dataframe(adv[['Ticker', '% Change', 'Close Price', 'Naira Change']], hide_index=True)
            
            st.subheader("🔴 Top 5 Decliners")
            st.dataframe(dec[['Ticker', '% Change', 'Close Price', 'Naira Change']], hide_index=True)
        else:
            st.error("Could not find market data. Please check your internet connection and try again.")
