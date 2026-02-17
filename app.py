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

# --- 1. DATE LOGIC (Lagos Time, 2:40 PM Cutoff) ---
def get_market_date():
    lagos_tz = pytz.timezone('Africa/Lagos')
    now = datetime.now(lagos_tz)
    weekday = now.weekday()  # 0=Mon, 6=Sun
    current_time = now.time()
    cutoff_time = dt_time(14, 40) # 2:40 PM

    if weekday == 5: # Saturday -> Show Friday
        report_date = now - timedelta(days=1)
    elif weekday == 6: # Sunday -> Show Friday
        report_date = now - timedelta(days=2)
    elif current_time < cutoff_time:
        # Before 2:40 PM: Show previous trading day
        # If Monday, go back to Friday (3 days)
        report_date = now - timedelta(days=3 if weekday == 0 else 1)
    else:
        # After 2:40 PM: Show today
        report_date = now
        
    return report_date.strftime("%d %B %Y").upper()

# --- 2. THE STEALTH BROWSER ---
def get_driver():
    options = Options()
    options.add_argument("--headless=new") # New headless mode is harder to detect
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    
    # Hiding Selenium from the Sucuri Firewall
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    
    # Critical for Streamlit Cloud
    options.binary_location = "/usr/bin/chromium"
    service = Service("/usr/bin/chromedriver")
    
    driver = webdriver.Chrome(service=service, options=options)
    
    # Disable the 'webdriver' flag so the site doesn't know we are a bot
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    })
    
    return driver

# --- 3. DATA SCRAPER ---
def fetch_ngx_data():
    driver = None
    try:
        driver = get_driver()
        # Direct URL to the price list which is more stable
        driver.get("https://ngxgroup.com/exchange/data/equities-price-list/")
        
        # Wait up to 30 seconds for the table to populate
        wait = WebDriverWait(driver, 30)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.wpDataTable tbody tr td")))
        
        # Give it a moment to finish rendering numeric values
        time.sleep(2)
        
        html_content = driver.page_source
        df_list = pd.read_html(io.StringIO(html_content))
        
        # Look for the table with "Symbol" or "% Change"
        df = None
        for table in df_list:
            if any('% Change' in str(col) for col in table.columns):
                df = table
                break
        
        if df is None:
            st.error("Table structure not found. Site may be under maintenance.")
            return None, None

        # Clean column names (NGX sometimes adds extra spaces)
        df.columns = [str(c).strip() for c in df.columns]
        
        # Standardize names
        df = df.rename(columns={
            'Symbol': 'Ticker',
            'Current': 'Close Price',
            'Change': 'Naira Change'
        })

        # Process numbers (Handle %, commas, and signs)
        for col in ['% Change', 'Close Price', 'Naira Change']:
            if col in df.columns:
                df[col] = df[col].astype(str).str.replace('%', '').str.replace(',', '').str.replace('+', '').strip()
                df[col] = pd.to_numeric(df[col], errors='coerce')

        # Get Top 5 Advancers (Gainers) and Decliners (Losers)
        df = df.dropna(subset=['% Change'])
        adv = df.sort_values(by='% Change', ascending=False).head(5)
        dec = df.sort_values(by='% Change', ascending=True).head(5)
        
        return adv, dec

    except Exception as e:
        st.error(f"⚠️ Market Access Error: {e}")
        return None, None
    finally:
        if driver:
            driver.quit()

# --- 4. EXPORT FUNCTIONS ---
def create_excel(adv, dec, date_str):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        adv[['Ticker', '% Change', 'Close Price', 'Naira Change']].to_excel(writer, sheet_name='Summary', startrow=2, index=False)
        dec[['Ticker', '% Change', 'Close Price', 'Naira Change']].to_excel(writer, sheet_name='Summary', startrow=10, index=False)
        ws = writer.sheets['Summary']
        bold = writer.book.add_format({'bold': True, 'align': 'center'})
        ws.merge_range('A1:D1', f"DAILY EQUITY SUMMARY FOR {date_str}", bold)
        ws.write('A2', 'TOP 5 ADVANCERS', bold)
        ws.write('A10', 'TOP 5 DECLINERS', bold)
    return output.getvalue()

def create_word(adv, dec, date_str):
    doc = Document()
    title = doc.add_heading(f"DAILY EQUITY SUMMARY FOR {date_str}", 1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for df, label in [(adv, "Top 5 Advancers"), (dec, "Top 5 Decliners")]:
        doc.add_heading(label, 2)
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        hdr = table.rows[0].cells
        hdr[0].text, hdr[1].text, hdr[2].text, hdr[3].text = "Ticker", "% Change", "Close Price", "Naira Change"
        for _, row in df.iterrows():
            c = table.add_row().cells
            c[0].text, c[1].text = str(row['Ticker']), f"{row['% Change']:.2f}%"
            c[2].text, c[3].text = f"{row['Close Price']:.2f}", f"{row['Naira Change']:.2f}"
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 5. UI ---
st.set_page_config(page_title="NGX Live Reporter", page_icon="🇳🇬")
st.title("🇳🇬 NGX Market Reporter")

market_date = get_market_date()
st.subheader(f"📅 Market Date: {market_date}")
st.caption("Auto-Logic: Weekends show Friday. Weekdays before 2:40 PM show previous day.")

if st.button("🚀 Fetch Top 5 Advancers & Decliners"):
    with st.spinner("Opening Secure Browser to NGX..."):
        adv, dec = fetch_ngx_data()
        
        if adv is not None and not adv.empty:
            st.success("Data Captured!")
            c1, c2 = st.columns(2)
            c1.download_button("📊 Download Excel", create_excel(adv, dec, market_date), f"NGX_Summary_{market_date}.xlsx")
            c2.download_button("📝 Download Word", create_word(adv, dec, market_date), f"NGX_Summary_{market_date}.docx")
            
            st.markdown("---")
            col_a, col_b = st.columns(2)
            with col_a:
                st.write("### 🟢 Top 5 Advancers")
                st.table(adv[['Ticker', '% Change', 'Close Price']])
            with col_b:
                st.write("### 🔴 Top 5 Decliners")
                st.table(dec[['Ticker', '% Change', 'Close Price']])
        else:
            st.warning("The NGX website is currently blocking the connection. Please wait 1 minute and click again to retry.")
