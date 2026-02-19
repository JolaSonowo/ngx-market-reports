import streamlit as st
from scraper import scrape_equities_table, get_top_movers, generate_excel, generate_word
import os

st.set_page_config(page_title="NGX Market Reporter", layout="wide")
st.title("🇳🇬 NGX Market Reporter")
st.write("Click the button below to fetch the latest market data and generate reports.")

if st.button("Generate Latest Market Report"):
    with st.spinner("Fetching NGX data..."):
        df = scrape_equities_table()
        gainers, decliners = get_top_movers(df)
        excel_path = generate_excel(gainers, decliners)
        word_path = generate_word(gainers, decliners)

    st.success("Reports Generated!")

    with open(excel_path, "rb") as f:
        st.download_button("Download Excel", f, file_name="NGX_Market_Report.xlsx")

    with open(word_path, "rb") as f:
        st.download_button("Download Word", f, file_name="NGX_Market_Report.docx")

st.write("Reports are generated in the `reports/` folder on the server.")
