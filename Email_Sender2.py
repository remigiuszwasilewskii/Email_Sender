import streamlit as st
import pandas as pd
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
import requests
from bs4 import BeautifulSoup

# Streamlit UI
st.title("Automated Report Generator")

# User Inputs
report_date_from = st.date_input("Report Date From")
report_date_to = st.date_input("Report Date To")

# Configurations
robot_path = r'\\nseura0090\DXCITO\NBSES\DMT\02 TOOLS\Nestle_EDF DE'
smtp_user = 'justyna.brzezniakiewicz@nestle.com'
smtp_password = 'Zadanie_dozrobienia2024'

# Read Config File
config_path = os.path.join(robot_path, 'data', 'Config.xlsx')
config_df = pd.read_excel(config_path, sheet_name=None)

# Extracting data from 'Config' sheet
config_sheet = config_df['Config']
payroll_folder = config_sheet.iloc[3, 3]
email_to = config_sheet.iloc[1, 0]
email_account = config_sheet.iloc[1, 1]
signature = config_sheet.iloc[1, 2]
email_folder = config_sheet.iloc[3, 0]
folder = config_sheet.iloc[3, 1]
rm = config_sheet.iloc[3, 2]

# Extracting data from 'Emails_Recon' sheet
emails_recon_sheet = config_df['Emails_Recon']
emails_recon = emails_recon_sheet.iloc[:, 0].dropna().tolist()

# Extracting data from 'Request_types' sheet
request_types_sheet = config_df['Request_types']
req_codes = request_types_sheet.iloc[:, :7].dropna().values.tolist()

# Extracting data from 'Out of scope' sheet
out_of_scope_sheet = config_df['Out of scope']
out_of_scope = out_of_scope_sheet.iloc[:, :2].dropna().values.tolist()

# Format Dates
formatted_date_from = report_date_from.strftime("%m.%d.%Y")
formatted_date_to = report_date_to.strftime("%m.%d.%Y")
formatted_date_from_report_name = formatted_date_from.replace(".", "/")
formatted_date_to_report_name = formatted_date_to.replace(".", "/")

# Create Subfolder
current_date_txt = datetime.now().strftime("%m.%d.%Y")
current_subfolder = os.path.join(folder, current_date_txt)
os.makedirs(current_subfolder, exist_ok=True)

# Start session for requests
session = requests.Session()

# Get the page
response = session.get(rm)
soup = BeautifulSoup(response.content, 'html.parser')

# Prepare the payload for form submission
payload = {
    "phlMain_ddlRegion": "Germany",
    "phlMain_ddlCountry": "Germany",
    "phlMain_ddlReferenceDate": "Closing",
    "phlMain_ddlStatus": "Closed",
    "x:362216426.0:mkr:3": formatted_date_from,
    "x:362216426.0:mkr:3": formatted_date_to
}

# Submit the form
response = session.post(form_action_url, data=payload)

# Check if the export was successful and get the report
# This part assumes the server returns a downloadable file in the response
report_path = os.path.join(robot_path, 'reports', f'EDF Report_{formatted_date_from_report_name} - {formatted_date_to_report_name}.xlsx')
with open(report_path, 'wb') as file:
    file.write(response.content)

# Prepare the report
prepare_report_path = os.path.join(robot_path, 'data', 'PrepareReport.xlsm')
prepare_report_df = pd.read_excel(prepare_report_path, sheet_name='Path', engine='openpyxl')
prepare_report_df.at[0, 'Path'] = report_path
prepare_report_df.to_excel(prepare_report_path, index=False, engine='openpyxl')

st.success("Report generated and saved successfully!")
