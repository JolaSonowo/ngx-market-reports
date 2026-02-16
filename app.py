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

    # Weekend or before 2:40 PM logic
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

# --- 2. FUZZY COLUMN FINDER ---
def find_column(df, keywords):
    """Finds a column name in df that matches any of the keywords."""
    for col in df.columns:
        if any(key.lower() in str(col).lower() for key in keywords):
            return col
    return None

# --- 3. ROBUST DATA FETCHING ---
def fetch_ngx_data():
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    
    # Sources: Official API, then secondary mirror
    sources = [
        "https://doclib.ngxgroup.com/REST/api/statistics/equities/?market=&sector=&orderby=&pageSize=300&pageNo=0",
        "https://afx.kwayisi.org/ngx/"
    ]
    
    for url in sources:
        try:
            if "json" in url or "doclib" in url:
                resp = requests.get(url, headers=headers, timeout=15)
                df = pd.DataFrame(resp.json())
            else:
                df = pd.read_html(url)[0]

            if df.empty: continue

            # Robust Column Mapping
            ticker_col = find_column(df, ['symbol', 'ticker', 'company', 'name'])
            price_col = find_column(df, ['close', 'price', 'current'])
            pct_col = find_column(df, ['%', 'gain', 'pchange', 'changepercentage'])
            naira_col = find_column(df, ['naira', 'change', 'absolute'])

            # Rename columns to standard format
            mapping = {}
            if ticker_col: mapping[ticker_col] = 'Ticker'
            if price_col: mapping[price_col] = 'Close Price'
            if pct_col: mapping[pct_col] = '% Change'
            if naira_col: mapping[naira_col] = 'Naira Change'
            
            df = df.rename(columns=mapping)

            # Keep only the columns we need
            df = df[['Ticker', '% Change', 'Close Price', 'Naira Change']]

            # Convert to numeric
            for col in ['% Change', 'Close Price', 'Naira Change']:
                df[col] = pd.to_numeric(df[col].astype(str).str.replace('%', '').str.replace(',', '').str.replace('+', '').str.strip(), errors='coerce')

            df = df.dropna(subset=['% Change'])
            adv = df.sort_values(by='% Change', ascending=False).head(5)
            dec = df.sort_values(by='% Change', ascending=True).head(5)
            return adv, dec

        except Exception as e:
            continue # Try next source if this one fails

    st.error("Could not retrieve data from any source. NGX servers may be down.")
    return None, None

# --- 4. EXPORT LOGIC ---
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
        hdrs = ["Ticker", "% Change", "Close Price", "Naira Change"]
        for i, h in enumerate(hdrs): table.rows[0].cells[i].text = h
        for _, row in df.iterrows():
            c = table.add_row().cells
            c[0].text, c[1].text = str(row['Ticker']), f"{row['% Change']:.2f}%"
            c[2].text, c[3].text = f"{row['Close Price']:.2f}", f"{row['Naira Change']:.2f}"
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 5. UI ---
st.set_page_config(page_title="NGX Reporter", page_icon="🇳🇬")
st.title("🇳🇬 NGX Top 5 Advancers & Decliners")

market_date = get_market_date()
st.info(f"Generating report for: **{market_date}**")

if st.button("🚀 Fetch Latest Data"):
    with st.spinner("Scanning market tables..."):
        adv, dec = fetch_ngx_data()
        
        if adv is not None and not adv.empty:
            st.success("Data loaded successfully!")
            c1, c2 = st.columns(2)
            c1.download_button("📊 Excel", create_excel(adv, dec, market_date), f"NGX_{market_date}.xlsx")
            c2.download_button("📝 Word", create_word(adv, dec, market_date), f"NGX_{market_date}.docx")
            
            st.subheader("🟢 Top 5 Advancers")
            st.dataframe(adv, hide_index=True)
            st.subheader("🔴 Top 5 Decliners")
            st.dataframe(dec, hide_index=True)
