import streamlit as st
import datetime
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import google.generativeai as genai
import io

# --- 1. CORE CONFIGURATION ---
MASTER_SHEET_ID = "15wPQ9QWydGKF1OIW1QkaeXB3msRjhwiJix4ZVyf6DxA"
DEPARTMENTS = ["English & Languages", "Social Sciences & Humanities", "Sciences", "Management", "Commerce"]
ACADEMIC_YEARS = ["2024-25", "2025-26", "2026-27", "2027-28", "2028-29", "2029-30"]

DEPARTMENT_FOLDERS = {
    "English & Languages": "14Nhs3qve5vDBbIT6GmzaRue51hvTzAOG",
    "Social Sciences & Humanities": "1m0xEcv-WKQr8CWfHlZ5AuCWIFXAm1H5g",
    "Sciences": "1u_KRBhdZhcWQ55CyVI0v042bIpC5FQfs",
    "Management": "1VG3xY_SmhqmQ9BvSh6KvDXOptO3kHhsj",
    "Commerce": "1HMBoNkhksNpaitlBaGfq3JeoHsb_jmo-"
}

# DIRECT BRIDGE: Secure backend Gemini connectivity
API_KEY = "AQ.Ab8RN6LKO3uh7uul5PXN870CGqBYigqEiVGCvBJl12elIT93uA"
genai.configure(api_key=API_KEY)

# --- 2. BACKEND API ENGINE ---
def get_google_credentials():
    """Generates the authenticated credentials object from Streamlit Secrets."""
    return service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive.file",
            "https://www.googleapis.com/auth/drive"
        ]
    )

def upload_file_to_drive(file_bytes, file_name, mime_type, target_folder_id, creds):
    """Uploads the file binary directly into the selected department folder."""
    try:
        drive_service = build('drive', 'v3', credentials=creds)
        
        file_metadata = {
            'name': file_name,
            'parents': [target_folder_id]
        }
        
        media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime_type, resumable=True)
        uploaded_file = drive_service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id, webViewLink',
            supportsAllDrives=True
        ).execute()
        
        try:
            drive_service.permissions().create(
                fileId=uploaded_file.get('id'),
                body={'type': 'anyone', 'role': 'reader'},
                supportsAllDrives=True
            ).execute()
        except Exception:
            pass
            
        return uploaded_file.get('webViewLink', "Link Error")
    except Exception as e:
        return "Drive Upload Pending (Permissions Check)"

def ai_extract_document_details(file_bytes, file_name, mime_type):
    """Uses Gemini 1.5 Flash vision capability to process raw image/PDF data structures safely."""
    try:
        model = genai.GenerativeModel("gemini-1.5-flash")
        
        prompt = """
        You are an institutional data extraction AI. Read all text in this academic certificate/document and extract information into a clean JSON object.
        
        Strictly output ONLY a valid JSON object with these exact keys:
        {
            "date_of_event": "In YYYY-MM-DD format. Look for terms like 'held on', 'date of issue', or completion date. Do NOT use today's date.",
            "submission_type": "Must strictly be chosen from one of these: FDP, Paper Presentation, Research Publication, Book Chapter, Book Publication, Workshop",
            "title": "The exact official name of the event, workshop, or paper title text"
        }
        Do not include markdown blocks, text explanations, or backticks around the JSON output.
        """
        
        document_payload = {
            "mime_type": mime_type,
            "data": file_bytes
        }
        
        response = model.generate_content([document_payload, prompt])
        
        response_text = response.text.strip()
        if response_text.startswith("```"):
            lines = response_text.splitlines()
            if lines[0].startswith("```"):
                lines = lines[1:]
            if lines[-1].startswith("```"):
                lines = lines[:-1]
            response_text = "\n".join(lines).strip()
            
        data = json.loads(response_text)
        return data
        
    except Exception as e:
        return {
            "date_of_event": "Check Original Document", 
            "submission_type": "FDP/Workshop",
            "title": f"Review Needed (AI Failed Parse): {file_name}"
        }

# --- 3. FRONTEND INTERFACE ---
st.set_page_config(page_title="RC Auto-Extraction Portal", page_icon="🤖", layout="centered")

st.title("🏢 St. Mary's Smart Research Desk")
st.markdown("Drop your documents below. The AI engine will read them, store them in your department's designated drive folder, and cluster them beautifully.")

# Main Form Formats
form_name = st.text_input("Faculty Member Name", placeholder="Type your full name...")
form_dept = st.selectbox("Select Department", DEPARTMENTS)
form_year = st.selectbox("Select Academic Year", ACADEMIC_YEARS)

uploaded_files = st.file_uploader(
    "Drop all your evidence certificates/papers here at once", 
    type=['pdf', 'png', 'jpg', 'jpeg'], 
    accept_multiple_files=True
)

submit_btn = st.button("🚀 Submit Documents to Master Sheet")

if submit_btn:
    if not form_name.strip():
        st.error("Please enter your name before submitting.")
    elif not uploaded_files:
        st.error("Please drop at least one file into the dropbox.")
    else:
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        try:
            creds = get_google_credentials()
            sheets_service = build('sheets', 'v4', credentials=creds)
            
            target_folder_id = DEPARTMENT_FOLDERS.get(form_dept, "14Nhs3qve5vDBbIT6GmzaRue51hvTzAOG")
            
            utc_now = datetime.datetime.utcnow()
            ist_offset = datetime.timedelta(hours=5, minutes=30)
            timestamp_now = (utc_now + ist_offset).strftime("%d-%m-%Y %H:%M:%S")
            
            for index, file in enumerate(uploaded_files):
                status_text.markdown(f"🤖 **AI is analyzing & sorting file {index+1}/{len(uploaded_files)}:** `{file.name}`...")
                
                file_bytes = file.read()
                
                # 1. Execute AI Parsing
                extracted = ai_extract_document_details(file_bytes, file.name, file.type)
                
                # 2. Sync to Dynamic Department Folder
                drive_link = upload_file_to_drive(file_bytes, file.name, file.type, target_folder_id, creds)
                
                # 3. Assemble structural row array matching Sheet headers
                new_entry = [
                    extracted.get("date_of_event", "Check Original Document"),
                    form_name.strip(),
                    extracted.get("submission_type", "FDP/Workshop"),
                    extracted.get("title", "Untitled Entry"),
                    drive_link,
                    form_dept,
                    timestamp_now
                ]
                
                # Force standard sheet target range matching manually selected dropdown year value
                sheet_range = f"'{form_year}'!A1:G1000"
                try:
                    result = sheets_service.spreadsheets().values().get(spreadsheetId=MASTER_SHEET_ID, range=sheet_range).execute()
                    rows = result.get('values', [])
                except Exception:
                    rows = []

                if not rows:
                    headers = ["Date", "Faculty Name", "Category", "Title", "Document Link", "Department", "Timestamp"]
                    rows = [headers, new_entry]
                else:
                    headers = rows[0]
                    data_rows = rows[1:]
                    data_rows.append(new_entry)
                    
                    data_rows.sort(key=lambda x: DEPARTMENTS.index(x[5]) if