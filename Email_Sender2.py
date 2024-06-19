import streamlit as st
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import time
import os
import shutil

# Streamlit app
st.title('Report Generation App')

# User inputs for date range
date_from = st.date_input('Report Date From')
date_to = st.date_input('Report Date To')

if st.button('Generate Report'):
    # Convert dates to required formats
    formatted_date_from = date_from.strftime('%m/%d/%Y').replace('.', '/')
    formatted_date_to = date_to.strftime('%m/%d/%Y').replace('.', '/')
    formatted_date_from_report_name = formatted_date_from.replace('/', '.')
    formatted_date_to_report_name = formatted_date_to.replace('/', '.')

    # Read configuration from Excel
    config_path = r'\\nseura0090\DXCITO\NBSES\DMT\02 TOOLS\Nestle_EDF DE\data\Config.xlsx'
    config = pd.read_excel(config_path, sheet_name=None)

    payroll_folder = config['Config'].iloc[3, 3]
    email_to = config['Config'].iloc[1, 0]
    email_account = config['Config'].iloc[1, 1]
    signature = config['Config'].iloc[1, 2]
    email_folder = config['Config'].iloc[3, 0]
    folder = config['Config'].iloc[3, 1]
    rm = config['Config'].iloc[3, 2]
    emails_recon = config['Emails_Recon'].iloc[:, 0].tolist()
    req_codes = config['Request_types'].values.tolist()
    out_of_scope = config['Out of scope'].values.tolist()

    # Create subfolder
    current_date_txt = datetime.now().strftime('%m.%d.%Y')
    current_subfolder = os.path.join(folder, current_date_txt)
    os.makedirs(current_subfolder, exist_ok=True)

    # Selenium web scraping
    options = webdriver.EdgeOptions()
    options.add_argument('headless')
    browser = webdriver.Edge(options=options)
    browser.get(rm)

    # Login and navigate through pages
    # ... (Assuming login and navigation is required and appropriate code is written here)

    # Set dropdown values and input dates
    Select(browser.find_element(By.ID, "phlMain_ddlRegion")).select_by_visible_text("Germany")
    Select(browser.find_element(By.ID, "phlMain_ddlCountry")).select_by_visible_text("Germany")
    Select(browser.find_element(By.ID, "phlMain_ddlReferenceDate")).select_by_visible_text("Closing")
    Select(browser.find_element(By.ID, "phlMain_ddlStatus")).select_by_visible_text("Closed")

    date_from_input = browser.find_element(By.ID, "x:362216426.0:mkr:3")
    date_from_input.clear()
    date_from_input.send_keys(formatted_date_from)
    date_to_input = browser.find_element(By.ID, "x:362216426.0:mkr:3")
    date_to_input.clear()
    date_to_input.send_keys(formatted_date_to)

    Select(browser.find_element(By.ID, "phlMain_ddlReportType")).select_by_visible_text("General data")
    browser.find_element(By.ID, "phlMain_btnExport").click()

    # Wait for download and move files
    downloads_folder = os.path.join(os.path.expanduser('~'), 'Downloads')
    time.sleep(10)  # Adjust this as needed based on download time
    downloaded_files = os.listdir(downloads_folder)
    for file in downloaded_files:
        if file.endswith('.xls') or file.endswith('.xlsx'):
            shutil.move(os.path.join(downloads_folder, file), os.path.join(current_subfolder, file))

    browser.quit()

    st.success('Report Generated and Saved Successfully!')
