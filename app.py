import io
import os.path
import streamlit as st
import pdfplumber   # PDF to Text

from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


st.set_page_config(page_title="PDF to Google Sheets", page_icon="📄")

st.title("📄 PDF to Google Sheets")
st.write("Upload your PDF files. The tool will append every page as text to your Google Sheets.")

# 1. File uploader box
uploaded_pdfs = st.file_uploader("Upload your PDF files here", type=['pdf'], accept_multiple_files=True)

# 2. Google Sheets ID input
spreadsheet_id = st.text_input("Enter your Google Sheets ID")

# 3. Process and Append button
if st.button("Process and Append to Sheets"):
    with st.spinner('Processing documents... please wait.'):

        # --- Setting up the Sheets API ---
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.file']

        creds = None
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json')
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            with open('token.json', 'w') as token:
                token.write(creds.to_json())

        service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()

        # --- Process PDFs and append to Google Sheet ---
        for uploaded_pdf in uploaded_pdfs:
            try:
                with pdfplumber.open(uploaded_pdf) as pdf:
                    for page in pdf.pages:
                        page_text = page.extract_text()
                        values = [[line] for line in page_text.split('\n')]
                        body = {'values': values}
                        
                        # Append PDF page to Google Sheets
                        result = service.spreadsheets().values().append(spreadsheetId=spreadsheet_id, range="A1", 
                                     valueInputOption="USER_ENTERED", body=body).execute()
            except Exception as e:
                st.error(f"Error processing {uploaded_pdf.name}: {e}")

    st.success("Process completed successfully! 🎉 PDF Pages are appended to Google Sheets.")
