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

    if weekday == 5: # Saturday
        report_date = now - timedelta(days=1)
    elif weekday == 6: # Sunday
        report_date = now - timedelta(days=2)
    elif current_time < cutoff_time:
        if weekday == 0: # Monday before 2:40 PM
            report_date = now - timedelta(days=3)
        else: # Weekday before 2:40 PM
            report_date = now - timedelta(days=1)
    else: # After 2:40 PM
        report_date = now
        
    return report_date.strftime("%d %B %Y").upper()

# --- 2. DATA EXTRACTION (Browser-less) ---
def fetch_ngx_data():
    # NGX uses this AJAX endpoint to load their tables without a browser
    url = "https://ngxgroup.com/wp-admin/admin-ajax.php?action=get_wdtable&table_id=2"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    
    try:
        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()
        json_data = response.json()
        
        # Load into DataFrame
        df = pd.DataFrame(json_data['data'])
        
        # Select and rename relevant columns
        # Based on NGX's internal JSON structure: symbol, current, pchange, change
        df = df[['symbol', 'pchange', 'current', 'change']].copy()
        df.columns = ['Ticker', '% Change', 'Close Price', 'Naira Change']

        # Convert strings to numbers
        for col in ['% Change', 'Close Price', 'Naira Change']:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace('%', '').str.replace(',', '').str.strip(), errors='coerce')

        # Drop any rows with missing data
        df = df.dropna(subset=['% Change'])

        # Get Top 5 Advancers and Decliners
        adv = df.sort_values(by='% Change', ascending=False).head(5)
        dec = df.sort_values(by='% Change', ascending=True).head(5)
        
        return adv, dec
    except Exception as e:
        st.error(f"Failed to fetch data: {e}")
        return None, None

# --- 3. EXPORT GENERATORS ---
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
        hdrs = table.rows[0].cells
        hdrs[0].text, hdrs[1].text, hdrs[2].text, hdrs[3].text = "Ticker", "% Change", "Close Price", "Naira Change"
        
        for _, row in df.iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = str(row['Ticker'])
            row_cells[1].text = f"{row['% Change']:.2f}%"
            row_cells[2].text = f"{row['Close Price']:.2f}"
            row_cells[3].text = f"{row['Naira Change']:.2f}"
            
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 4. STREAMLIT UI ---
st.set_page_config(page_title="NGX Reporter", page_icon="🇳🇬")
st.title("📈 NGX Market Data Extractor")

market_date = get_market_date()
st.info(f"Generating data for: **{market_date}**")

if st.button("🚀 Fetch NGX Figures"):
    with st.spinner("Connecting to NGX..."):
        adv, dec = fetch_ngx_data()
        
        if adv is not None:
            st.success("Data loaded successfully!")
            
            c1, c2 = st.columns(2)
            c1.download_button("📊 Download Excel", create_excel(adv, dec, market_date), f"NGX_Summary_{market_date}.xlsx")
            c2.download_button("📝 Download Word", create_word(adv, dec, market_date), f"NGX_Summary_{market_date}.docx")
            
            st.subheader("Top 5 Advancers")
            st.dataframe(adv, use_container_width=True)
            
            st.subheader("Top 5 Decliners")
            st.dataframe(dec, use_container_width=True)
