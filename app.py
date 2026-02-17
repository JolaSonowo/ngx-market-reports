import streamlit as st
import pandas as pd
import cloudscraper
import io
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime, time, timedelta
import pytz

# --- 1. SMART DATE LOGIC (Lagos Timezone) ---
def get_report_info():
    lagos_tz = pytz.timezone('Africa/Lagos')
    now = datetime.now(lagos_tz)
    weekday = now.weekday()  # 0=Mon, 6=Sun
    cutoff_time = time(14, 40) # 2:40 PM
    
    # Logic:
    # 1. Weekend -> Show Friday
    # 2. Weekday before 2:40 PM -> Show Previous Trading Day
    # 3. Weekday after 2:40 PM -> Show Today
    
    if weekday == 5: # Saturday
        target_date = now - timedelta(days=1)
    elif weekday == 6: # Sunday
        target_date = now - timedelta(days=2)
    elif now.time() < cutoff_time:
        if weekday == 0: # Monday morning -> Friday
            target_date = now - timedelta(days=3)
        else:
            target_date = now - timedelta(days=1)
    else:
        target_date = now
        
    return target_date.strftime("%d %B %Y").upper()

# --- 2. DATA EXTRACTION (Firewall Bypass) ---
def fetch_ngx_data():
    # Use cloudscraper to bypass Sucuri Firewall
    scraper = cloudscraper.create_scraper(
        browser={
            'browser': 'chrome',
            'platform': 'windows',
            'desktop': True
        }
    )
    
    # We target the AJAX endpoint NGX uses to load the price table
    # This is more reliable than scraping the visible HTML
    url = "https://ngxgroup.com/wp-admin/admin-ajax.php?action=get_wdtable&table_id=2"
    
    try:
        response = scraper.get(url, timeout=15)
        if response.status_code != 200:
            # Fallback: Try the direct Price List page
            response = scraper.get("https://afx.kwayisi.org/ngx/", timeout=15)
            df = pd.read_html(io.StringIO(response.text))[0]
            df = df.rename(columns={'Ticker': 'Ticker', 'Price': 'Close Price', 'Gain': '% Change', 'Change': 'Naira Change'})
        else:
            data = response.json()
            df = pd.DataFrame(data['data'])
            # Mapping API keys to readable names
            df = df.rename(columns={
                'symbol': 'Ticker',
                'current': 'Close Price',
                'pchange': '% Change',
                'change': 'Naira Change'
            })

        # Clean numeric columns
        for col in ['% Change', 'Close Price', 'Naira Change']:
            if col in df.columns:
                df[col] = df[col].astype(str).str.replace('%', '').str.replace(',', '').str.replace('+', '').strip()
                df[col] = pd.to_numeric(df[col], errors='coerce')

        df = df.dropna(subset=['Ticker', '% Change'])
        
        # Sort to get Top 5 Advancers and Decliners
        adv = df.sort_values(by='% Change', ascending=False).head(5)
        dec = df.sort_values(by='% Change', ascending=True).head(5)
        
        return adv, dec

    except Exception as e:
        st.error(f"Fetch Error: {e}")
        return None, None

# --- 3. FILE GENERATION ---
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
        headers = ["Ticker", "% Change", "Close Price", "Naira Change"]
        for i, h in enumerate(headers):
            table.rows[0].cells[i].text = h
        for _, row in df.iterrows():
            cells = table.add_row().cells
            cells[0].text = str(row['Ticker'])
            cells[1].text = f"{row['% Change']:.2f}%"
            cells[2].text = f"{row['Close Price']:.2f}"
            cells[3].text = f"{row['Naira Change']:.2f}"
            
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 4. STREAMLIT UI ---
st.set_page_config(page_title="NGX Smart Reporter", page_icon="🇳🇬")
st.title("🇳🇬 NGX Market Reporter")

report_date = get_report_info()
st.subheader(f"📅 Report Date: {report_date}")
st.caption("Auto-Logic: Weekends show Friday. Weekdays before 2:40 PM show previous trading day.")

if st.button("🚀 Fetch Data & Generate Reports"):
    with st.spinner("Connecting to NGX Secure Data Feed..."):
        adv, dec = fetch_ngx_data()
        
        if adv is not None and not adv.empty:
            st.success("Data Captured Successfully!")
            
            col1, col2 = st.columns(2)
            with col1:
                st.download_button("📊 Excel Report", create_excel(adv, dec, report_date), f"NGX_Summary_{report_date}.xlsx")
            with col2:
                st.download_button("📝 Word Report", create_word(adv, dec, report_date), f"NGX_Summary_{report_date}.docx")
            
            st.markdown("---")
            st.write("### 🟢 Top 5 Advancers")
            st.table(adv[['Ticker', '% Change', 'Close Price']])
            
            st.write("### 🔴 Top 5 Decliners")
            st.table(dec[['Ticker', '% Change', 'Close Price']])
        else:
            st.warning("Could not retrieve data. The NGX servers may be under high load or blocking requests. Please try again in 30 seconds.")
