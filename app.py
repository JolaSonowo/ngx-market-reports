import streamlit as st
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
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
    options.add_argument("--headless=new") # Optimized headless mode
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    
    # --- FIREWALL BYPASS STEALTH SETTINGS ---
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument("--disable-blink-features=AutomationControlled")
    
    options.binary_location = "/usr/bin/chromium"
    service = Service("/usr/bin/chromedriver")
    
    return webdriver.Chrome(service=service, options=options)

def fetch_ngx_data():
    driver = None
    try:
        driver = get_driver()
        # Remove the 'webdriver' flag so the site thinks we are human
        driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        })
        
        # 1. Visit the home page
        driver.get("https://ngxgroup.com/")
        
        # 2. Wait and click the Price List tab
        wait = WebDriverWait(driver, 30)
        tab = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='#tab2']")))
        
        # Scroll to it and click
        driver.execute_script("arguments[0].scrollIntoView();", tab)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", tab)
        
        # 3. Wait for the table to actually populate with data
        # We look for a table cell that contains data, not just the header
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#tab2 table.wpDataTable tbody tr td")))
        
        # 4. Extract data
        html_content = driver.page_source
        df_list = pd.read_html(io.StringIO(html_content))
        
        # Find the correct table in the list
        df = None
        for table in df_list:
            if 'Symbol' in table.columns:
                df = table
                break
        
        if df is None:
            st.error("Table found but columns don't match. Website structure might have changed.")
            return None, None

        # Clean and filter
        df = df.rename(columns={'Symbol': 'Ticker', 'Current': 'Close Price', 'Change': 'Naira Change'})
        for col in ['% Change', 'Close Price', 'Naira Change']:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace('%', '').str.replace(',', '').str.strip(), errors='coerce')

        df = df.dropna(subset=['% Change'])
        adv = df.sort_values(by='% Change', ascending=False).head(5)
        dec = df.sort_values(by='% Change', ascending=True).head(5)
        return adv, dec

    except Exception as e:
        st.error(f"Error fetching data: {e}")
        return None, None
    finally:
        if driver:
            driver.quit()

# --- FILE GENERATORS (Same as before) ---
def create_excel(adv, dec, date_str):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        adv.to_excel(writer, sheet_name='Summary', startrow=2, index=False)
        dec.to_excel(writer, sheet_name='Summary', startrow=10, index=False)
        ws = writer.sheets['Summary']
        bold = writer.book.add_format({'bold': True, 'align': 'center'})
        ws.merge_range('A1:D1', f"DAILY EQUITY SUMMARY FOR {date_str}", bold)
        ws.write('A2', 'TOP 5 ADVANCERS', bold); ws.write('A10', 'TOP 5 DECLINERS', bold)
    return output.getvalue()

def create_word(adv, dec, date_str):
    doc = Document()
    doc.add_heading(f"DAILY EQUITY SUMMARY FOR {date_str}", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    for df, label in [(adv, "Top 5 Advancers"), (dec, "Top 5 Decliners")]:
        doc.add_heading(label, 2)
        table = doc.add_table(rows=1, cols=4); table.style = 'Table Grid'
        for i, h in enumerate([label, "% Change", "Close Price", "Naira Change"]):
            table.rows[0].cells[i].text = h
        for _, row in df.iterrows():
            c = table.add_row().cells
            c[0].text, c[1].text = str(row['Ticker']), f"{row['% Change']:.2f}%"
            c[2].text, c[3].text = f"{row['Close Price']:.2f}", f"{row['Naira Change']:.2f}"
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- UI ---
st.set_page_config(page_title="NGX Reporter", page_icon="🇳🇬")
st.title("🇳🇬 NGX Market Reporter")
market_date = get_market_date()
st.info(f"Market Date: **{market_date}**")

if st.button("🚀 Run Report"):
    with st.spinner("Accessing NGX website... (This takes about 20-30 seconds)"):
        adv, dec = fetch_ngx_data()
        if adv is not None:
            st.success("Data Captured!")
            c1, c2 = st.columns(2)
            c1.download_button("📊 Excel", create_excel(adv, dec, market_date), f"NGX_{market_date}.xlsx")
            c2.download_button("📝 Word", create_word(adv, dec, market_date), f"NGX_{market_date}.docx")
            st.subheader("Top 5 Advancers")
            st.table(adv[['Ticker', '% Change', 'Close Price', 'Naira Change']])
            st.subheader("Top 5 Decliners")
            st.table(dec[['Ticker', '% Change', 'Close Price', 'Naira Change']])
