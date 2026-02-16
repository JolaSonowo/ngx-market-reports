import streamlit as st
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import io
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime, time, timedelta
import pytz

# --- DATE LOGIC ---
def get_market_date():
    lagos_tz = pytz.timezone('Africa/Lagos')
    now = datetime.now(lagos_tz)
    weekday = now.weekday() 
    current_time = now.time()
    cutoff_time = time(14, 40) 

    if weekday == 5: report_date = now - timedelta(days=1)
    elif weekday == 6: report_date = now - timedelta(days=2)
    elif current_time < cutoff_time:
        report_date = now - timedelta(days=3 if weekday == 0 else 1)
    else: report_date = now
    return report_date.strftime("%d %B %Y").upper()

# --- THE "STREAMLIT CLOUD" DRIVER ---
def get_driver():
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    
    # Bypassing Firewall: Use a very common real-world User-Agent
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36")
    
    # CRITICAL: These paths are where Streamlit Cloud installs Chromium via packages.txt
    options.binary_location = "/usr/bin/chromium"
    service = Service("/usr/bin/chromedriver")
    
    return webdriver.Chrome(service=service, options=options)

def fetch_ngx_data():
    driver = None
    try:
        driver = get_driver()
        # Visit the main site first to establish a "human" session cookie
        driver.get("https://ngxgroup.com/")
        
        wait = WebDriverWait(driver, 25)
        # Find and click the 'Price List' tab (Tab 2)
        tab = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='#tab2']")))
        driver.execute_script("arguments[0].click();", tab)
        
        # Wait for the specific data table to appear
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.wpDataTable")))
        
        df_list = pd.read_html(io.StringIO(driver.page_source))
        df = df_list[0]
        
        # Rename and Clean
        df = df.rename(columns={'Symbol': 'Ticker', 'Current': 'Close Price', 'Change': 'Naira Change'})
        for col in ['% Change', 'Close Price', 'Naira Change']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col].astype(str).str.replace('%', '').str.replace(',', '').strip(), errors='coerce')

        df = df.dropna(subset=['% Change'])
        adv = df.sort_values(by='% Change', ascending=False).head(5)
        dec = df.sort_values(by='% Change', ascending=True).head(5)
        return adv, dec
    except Exception as e:
        st.error(f"Error fetching data: {e}")
        return None, None
    finally:
        if driver: driver.quit()

# --- FILE GENERATORS ---
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
st.set_page_config(page_title="NGX Reporter")
st.title("🇳🇬 NGX Market Reporter")
market_date = get_market_date()
st.info(f"Report Date: **{market_date}**")

if st.button("🚀 Generate Reports"):
    with st.spinner("Accessing NGX website..."):
        adv, dec = fetch_ngx_data()
        if adv is not None:
            st.success("Data Captured Successfully!")
            col1, col2 = st.columns(2)
            col1.download_button("📊 Excel", create_excel(adv, dec, market_date), f"NGX_{market_date}.xlsx")
            col2.download_button("📝 Word", create_word(adv, dec, market_date), f"NGX_{market_date}.docx")
            st.subheader("Top 5 Advancers")
            st.table(adv[['Ticker', '% Change', 'Close Price', 'Naira Change']])
            st.subheader("Top 5 Decliners")
            st.table(dec[['Ticker', '% Change', 'Close Price', 'Naira Change']])
