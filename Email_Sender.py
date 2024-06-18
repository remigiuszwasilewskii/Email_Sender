import win32com.client as win32
import json
from datetime import datetime
import openpyxl
import streamlit as st
from tempfile import NamedTemporaryFile


def load_email_data(file):
    """
    Load email data from a JSON file.

    Args:
        file: The uploaded JSON file.

    Returns:
        dict: A dictionary containing email data.
    """
    return json.load(file)


def get_account(outlook, sender_email):
    """
    Get the Outlook account corresponding to the sender's email address.

    Args:
        outlook: The Outlook application instance.
        sender_email (str): The sender's email address.

    Returns:
        Account: The Outlook account corresponding to the sender's email address, or None if not found.
    """
    for account in outlook.Session.Accounts:
        if account.SmtpAddress.lower() == sender_email.lower():
            return account
    return None


def log_to_excel(subject, status, error=None):
    """
    Log email status to an Excel file.

    Args:
        subject (str): The subject of the email.
        status (str): The status of the email.
        error (str, optional): Any error message encountered during email sending.

    Returns:
        None
    """
    log_file = 'Log.xlsx'
    try:
        workbook = openpyxl.load_workbook(log_file)
        sheet = workbook.active
        next_row = sheet.max_row + 1
        current_time = datetime.now().strftime("%H:%M:%S | %d.%m.%Y")
        sheet[f'A{next_row}'] = current_time
        sheet[f'B{next_row}'] = subject
        sheet[f'C{next_row}'] = status
        if error:
            sheet[f'D{next_row}'] = str(error)
        workbook.save(log_file)
    except Exception as e:
        st.error(f"Failed to log to Excel: {e}")


def send_email(outlook, sender_account, recipient, subject, body, attachments):
    """
    Send an email.

    Args:
        outlook: The Outlook application instance.
        sender_account: The Outlook account from which to send the email.
        recipient (str): The recipient's email address.
        subject (str): The subject of the email.
        body (str): The body content of the email.
        attachments (list): A list of file paths to attachments (optional).

    Returns:
        bool: True if the email was sent successfully, False otherwise.
    """
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Subject = subject
    mail.Body = body
    for attachment in attachments:
        mail.Attachments.Add(attachment)
    try:
        mail._oleobj_.Invoke(*(64209, 0, 8, 0, sender_account))
        mail.Send()
        return True
    except Exception as e:
        log_to_excel(subject, "Error", e)
        return False


def main():
    """
    Main function to send emails.

    Returns:
        None
    """
    st.title('Email Sender with Outlook')

    uploaded_file = st.file_uploader("Choose a JSON file", type="json")

    if uploaded_file is not None:
        try:
            email_data_list = load_email_data(uploaded_file)
            outlook = win32.Dispatch('outlook.application')

            for email_data in email_data_list:
                sender = email_data['sender']
                recipients = email_data['recipients']
                subject = email_data['subject']
                body = email_data['body']
                attachments = email_data.get('attachments', [])
                sender_account = get_account(outlook, sender)

                if sender_account is None:
                    log_to_excel(subject, f"Account with email {sender} not found.")
                else:
                    for recipient in recipients:
                        status = "Successful" if send_email(outlook, sender_account, recipient, subject, body,
                                                            attachments) else "Error"
                        log_to_excel(subject, status)

            st.success("Emails processed. Check the log file for details.")
        except Exception as e:
            st.error(f"Error processing the file: {e}")


if __name__ == "__main__":
    main()
