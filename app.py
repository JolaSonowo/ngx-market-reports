import streamlit as st
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.core.os_manager import ChromeType
import io
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime, time, timedelta
import pytz

# --- DATE LOGIC ---
def get_market_date():
    lagos_tz = pytz.timezone('Africa/Lagos')
    now = datetime.now(lagos_tz)
    weekday = now.weekday()  # 0=Mon, 4=Fri, 5=Sat, 6=Sun
    current_time = now.time()
    cutoff_time = time(14, 40) # 2:40 PM

    # If Weekend (Saturday or Sunday)
    if weekday == 5: # Saturday
        report_date = now - timedelta(days=1)
    elif weekday == 6: # Sunday
        report_date = now - timedelta(days=2)
    # If Weekday but before 2:40 PM
    elif current_time < cutoff_time:
        if weekday == 0: # Monday morning -> Go back to Friday
            report_date = now - timedelta(days=3)
        else:
            report_date = now - timedelta(days=1)
    # If Weekday after 2:40 PM
    else:
        report_date = now
        
    return report_date.strftime("%d %B %Y").upper()

# --- BROWSER CONFIGURATION ---
def get_driver():
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
    
    return webdriver.Chrome(
        service=Service(ChromeDriverManager(chrome_type=ChromeType.CHROMIUM).install()), 
        options=options
    )

# --- DATA EXTRACTION ---
def fetch_ngx_figures():
    driver = get_driver()
    url = "https://ngxgroup.com/#tab2"
    
    try:
        driver.get(url)
        wait = WebDriverWait(driver, 30)
        # Wait for the table to be visible
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.wpDataTable")))
        
        html_content = driver.page_source
        driver.quit()
        
        df_list = pd.read_html(io.StringIO(html_content))
        # Usually the first table on that tab is the price list
        df = df_list[0]
        
        # Rename based on NGX table structure
        df = df.rename(columns={'Symbol': 'Ticker', 'Current': 'Close Price', 'Change': 'Naira Change'})
        
        # Clean numeric columns
        for col in ['% Change', 'Close Price', 'Naira Change']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col].astype(str).str.replace('%', '').str.replace(',', '').str.replace(' ', ''), errors='coerce')
        
        df = df.dropna(subset=['% Change'])
        
        # Get Top 5 Advancers and Decliners
        adv = df.sort_values(by='% Change', ascending=False).head(5)
        dec = df.sort_values(by='% Change', ascending=True).head(5)
        
        return adv, dec
    except Exception as e:
        if 'driver' in locals(): driver.quit()
        st.error(f"Figure Extraction Failed: {e}")
        return None, None

# --- FILE GENERATION ---
def create_excel(adv, dec, date_str):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        adv.rename(columns={'Ticker': 'Top Advancers'}).to_excel(writer, sheet_name='Summary', startrow=2, index=False)
        dec.rename(columns={'Ticker': 'Top Decliners'}).to_excel(writer, sheet_name='Summary', startrow=10, index=False)
        ws = writer.sheets['Summary']
        fmt = writer.book.add_format({'bold': True, 'align': 'center', 'font_size': 14})
        ws.merge_range('A1:D1', f"DAILY EQUITY SUMMARY FOR {date_str}", fmt)
    return output.getvalue()

def create_word(adv, dec, date_str):
    doc = Document()
    p = doc.add_heading(f"DAILY EQUITY SUMMARY FOR {date_str}", level=1)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    for df, label in [(adv, "Top Advancers"), (dec, "Top Decliners")]:
        doc.add_heading(label, level=2)
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        headers = [label, "% Change", "Close Price", "Naira Change"]
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

# --- STREAMLIT UI ---
st.set_page_config(page_title="NGX Live Reporter", page_icon="🇳🇬")
st.title("📈 NGX Top 5 Advancers & Decliners")

market_date = get_market_date()
st.info(f"Report Date: **{market_date}**")
st.caption("Note: Data automatically reflects the previous trading day if viewed before 2:40 PM WAT.")

if st.button("🚀 Fetch NGX Data"):
    with st.spinner("Scraping live data from NGX..."):
        adv, dec = fetch_ngx_figures()
        
        if adv is not None and not adv.empty:
            st.success(f"Data for {market_date} captured successfully!")
            
            col1, col2 = st.columns(2)
            with col1:
                st.download_button("📊 Download Excel", create_excel(adv, dec, market_date), f"NGX_Summary_{market_date}.xlsx")
            with col2:
                st.download_button("📝 Download Word", create_word(adv, dec, market_date), f"NGX_Summary_{market_date}.docx")
            
            st.subheader("Top 5 Advancers")
            st.table(adv[['Ticker', '% Change', 'Close Price', 'Naira Change']])
            
            st.subheader("Top 5 Decliners")
            st.table(dec[['Ticker', '% Change', 'Close Price', 'Naira Change']])
        else:
            st.warning("No data found. The market might be closed or the website structure has changed.")
