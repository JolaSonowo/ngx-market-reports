import streamlit as st
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium_stealth import stealth
import io
import time
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime, time as dt_time, timedelta
import pytz

# --- DATE LOGIC ---
def get_market_date():
    lagos_tz = pytz.timezone('Africa/Lagos')
    now = datetime.now(lagos_tz)
    weekday = now.weekday() 
    current_time = now.time()
    cutoff_time = dt_time(14, 40) 

    if weekday == 5: report_date = now - timedelta(days=1)
    elif weekday == 6: report_date = now - timedelta(days=2)
    elif current_time < cutoff_time:
        report_date = now - timedelta(days=3 if weekday == 0 else 1)
    else: report_date = now
    return report_date.strftime("%d %B %Y").upper()

# --- STEALTH DRIVER SETUP ---
def get_driver():
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    
    # Standard paths for Streamlit Cloud
    options.binary_location = "/usr/bin/chromium"
    service = Service("/usr/bin/chromedriver")
    
    driver = webdriver.Chrome(service=service, options=options)
    
    # Apply Stealth settings to bypass Sucuri Firewall
    stealth(driver,
        languages=["en-US", "en"],
        vendor="Google Inc.",
        platform="Win32",
        webgl_vendor="Intel Inc.",
        renderer="Intel Iris OpenGL Engine",
        fix_hairline=True,
    )
    return driver

def fetch_ngx_data():
    driver = None
    try:
        driver = get_driver()
        # Navigate directly to the Price List page
        driver.get("https://ngxgroup.com/exchange/data/equities-price-list/")
        
        wait = WebDriverWait(driver, 45) # Increased timeout
        
        # Wait for the table to appear (searching for the wpDataTable class)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.wpDataTable")))
        time.sleep(2) # Extra time for JS to populate rows
        
        # Extract data
        html_content = driver.page_source
        df_list = pd.read_html(io.StringIO(html_content))
        
        # Filter for the correct table
        df = None
        for table in df_list:
            if 'Symbol' in table.columns:
                df = table
                break
        
        if df is None or df.empty:
            return None, None

        # Clean columns to match NGX website headers
        df = df.rename(columns={'Symbol': 'Ticker', 'Current': 'Close Price', 'Change': 'Naira Change'})
        
        # Process numeric columns
        for col in ['% Change', 'Close Price', 'Naira Change']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col].astype(str).str.replace('%', '').str.replace(',', '').str.strip(), errors='coerce')

        df = df.dropna(subset=['% Change'])
        
        # Top 5 Advancers and Decliners
        adv = df.sort_values(by='% Change', ascending=False).head(5)
        dec = df.sort_values(by='% Change', ascending=True).head(5)
        
        return adv, dec

    except Exception as e:
        st.error(f"Scraping Error: {e}")
        return None, None
    finally:
        if driver:
            driver.quit()

# --- EXPORT GENERATORS ---
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
        headers = [label, "% Change", "Close Price", "Naira Change"]
        for i, h in enumerate(headers):
            table.rows[0].cells[i].text = h
        for _, row in df.iterrows():
            c = table.add_row().cells
            c[0].text, c[1].text = str(row['Ticker']), f"{row['% Change']:.2f}%"
            c[2].text, c[3].text = f"{row['Close Price']:.2f}", f"{row['Naira Change']:.2f}"
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- STREAMLIT UI ---
st.set_page_config(page_title="NGX Reporter")
st.title("🇳🇬 NGX Market Report")

market_date = get_market_date()
st.info(f"Report for: **{market_date}**")

if st.button("🚀 Fetch Data"):
    with st.spinner("Bypassing firewall... Please wait."):
        adv, dec = fetch_ngx_data()
        if adv is not None:
            st.success(f"Successfully captured data for {market_date}!")
            c1, c2 = st.columns(2)
            c1.download_button("📊 Download Excel", create_excel(adv, dec, market_date), f"NGX_{market_date}.xlsx")
            c2.download_button("📝 Download Word", create_word(adv, dec, market_date), f"NGX_{market_date}.docx")
            
            st.subheader("Top 5 Advancers")
            st.table(adv[['Ticker', '% Change', 'Close Price', 'Naira Change']])
            st.subheader("Top 5 Decliners")
            st.table(dec[['Ticker', '% Change', 'Close Price', 'Naira Change']])
