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
        if weekday == 0: # Monday morning -> Friday
            report_date = now - timedelta(days=3)
        else: # Weekday morning -> Yesterday
            report_date = now - timedelta(days=1)
    else: # Weekday afternoon -> Today
        report_date = now
        
    return report_date.strftime("%d %B %Y").upper()

# --- 2. LIGHTWEIGHT DATA FETCHING (No Chrome Needed) ---
def fetch_ngx_data():
    # We use a reliable mirror that is much faster than the official NGX site
    url = "https://afx.kwayisi.org/ngx/"
    headers = {"User-Agent": "Mozilla/5.0"}
    
    try:
        # Fetch tables from the page
        response = requests.get(url, headers=headers, timeout=10)
        df_list = pd.read_html(io.StringIO(response.text))
        
        # The main price table is usually the first one
        df = df_list[0]

        # Rename columns to match your requirements
        # Kwayisi format: Ticker, Name, Price, Change, Gain (%)
        # We hunt for the columns by keyword to be safe
        ticker_col = [c for c in df.columns if 'Ticker' in str(c)][0]
        price_col = [c for c in df.columns if 'Price' in str(c)][0]
        gain_col = [c for c in df.columns if 'Gain' in str(c) or '%' in str(c)][0]
        change_col = [c for c in df.columns if 'Change' in str(c)][0]

        df = df[[ticker_col, gain_col, price_col, change_col]]
        df.columns = ['Ticker', '% Change', 'Close Price', 'Naira Change']

        # Convert to numeric
        for col in ['% Change', 'Close Price', 'Naira Change']:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace('%', '').str.replace(',', '').str.replace('+', '').strip(), errors='coerce')

        df = df.dropna(subset=['% Change'])
        
        # Get Top 5 Advancers and Decliners
        adv = df.sort_values(by='% Change', ascending=False).head(5)
        dec = df.sort_values(by='% Change', ascending=True).head(5)
        
        return adv, dec
    except Exception as e:
        st.error(f"Fetch failed: {e}")
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
        hdrs = ["Ticker", "% Change", "Close Price", "Naira Change"]
        for i, h in enumerate(hdrs): table.rows[0].cells[i].text = h
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
st.info(f"Report for: **{market_date}**")

if st.button("🚀 Get Data Now"):
    with st.spinner("Fetching live market summary..."):
        adv, dec = fetch_ngx_data()
        if adv is not None:
            st.success("Data ready!")
            c1, c2 = st.columns(2)
            c1.download_button("📊 Excel", create_excel(adv, dec, market_date), f"NGX_{market_date}.xlsx")
            c2.download_button("📝 Word", create_word(adv, dec, market_date), f"NGX_{market_date}.docx")
            
            st.subheader("🟢 Top 5 Advancers")
            st.dataframe(adv, hide_index=True)
            st.subheader("🔴 Top 5 Decliners")
            st.dataframe(dec, hide_index=True)
