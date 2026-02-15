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
from datetime import datetime
import pytz

# --- BROWSER CONFIGURATION ---
def get_driver():
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
    
    # This setup works for both local and Streamlit Cloud
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
        # Wait specifically for the data rows to appear in the table body
        wait = WebDriverWait(driver, 25)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.wpDataTable tbody tr td")))
        
        # Get the finalized HTML after JS has rendered the figures
        html_content = driver.page_source
        driver.quit()
        
        df_list = pd.read_html(io.StringIO(html_content))
        df = df_list[0]
        
        # Rename columns to your standard requirements
        df = df.rename(columns={'Symbol': 'Ticker', 'Current': 'Close Price', 'Change': 'Naira Change'})
        
        # Convert strings to numeric (stripping % and commas)
        for col in ['% Change', 'Close Price', 'Naira Change']:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace('%', '').str.replace(',', ''), errors='coerce')
        
        # Filter out empty rows and get real Top 5
        df = df.dropna(subset=['% Change'])
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
        adv.rename(columns={'Ticker': 'Gainers'}).to_excel(writer, sheet_name='Summary', startrow=2, index=False)
        dec.rename(columns={'Ticker': 'Decliners'}).to_excel(writer, sheet_name='Summary', startrow=10, index=False)
        ws = writer.sheets['Summary']
        fmt = writer.book.add_format({'bold': True, 'align': 'center', 'font_size': 14})
        ws.merge_range('A1:D1', f"DAILY EQUITY SUMMARY FOR {date_str}", fmt)
    return output.getvalue()

def create_word(adv, dec, date_str):
    doc = Document()
    p = doc.add_heading(f"DAILY EQUITY SUMMARY FOR {date_str}", level=1)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    for df, label in [(adv, "Top Gainers"), (dec, "Top Decliners")]:
        doc.add_heading(label, level=2)
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        for i, h in enumerate([label, "% Change", "Close Price", "Naira Change"]):
            table.rows[0].cells[i].text = h
        for _, row in df.iterrows():
            cells = table.add_row().cells
            cells[0].text, cells[1].text = str(row['Ticker']), f"{row['% Change']:.2f}%"
            cells[2].text, cells[3].text = f"{row['Close Price']:.2f}", f"{row['Naira Change']:.2f}"
    
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- STREAMLIT UI ---
st.set_page_config(page_title="NGX Live Reporter", page_icon="🇳🇬")
st.title("📈 NGX Live Figure Extractor")

lagos_now = datetime.now(pytz.timezone('Africa/Lagos'))
date_header = lagos_now.strftime("%d%b %Y").upper() # Simplified for code clarity

if st.button("🚀 Load Website & Capture Figures"):
    with st.spinner("Opening browser... Please wait for JS figures to load."):
        adv, dec = fetch_ngx_figures()
    
    if adv is not None:
        st.success("Figures captured successfully!")
        
        col1, col2 = st.columns(2)
        with col1:
            st.download_button("📊 Excel Report", create_excel(adv, dec, date_header), "NGX_Summary.xlsx")
        with col2:
            st.download_button("📝 Word Report", create_word(adv, dec, date_header), "NGX_Summary.docx")
            
        st.subheader("Live Market Preview")
        st.dataframe(adv[['Ticker', '% Change', 'Close Price', 'Naira Change']])
