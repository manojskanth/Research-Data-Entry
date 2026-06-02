import streamlit as st
import datetime
import json
import re
import io

try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseUpload
    import google.generativeai as genai
except ImportError as e:
    st.error(f"Environment Assembly Pending: Missing library package ({str(e)})")

# --- 1. CORE CONFIGURATION ---
MASTER_SHEET_ID = "15wPQ9QWydGKF1OIW1QkaeXB3msRjhwiJix4ZVyf6DxA"
DEPARTMENTS = [
    "English & Languages", 
    "Social Sciences & Humanities", 
    "Sciences", 
    "Management", 
    "Commerce"
]
ACADEMIC_YEARS = [
    "2024-25", 
    "2025-26", 
    "2026-27", 
    "2027-28", 
    "2028-29", 
    "2029-30"
]

DEPARTMENT_FOLDERS = {
    "English & Languages": "14Nhs3qve5vDBbIT6GmzaRue51hvTzAOG",
    "Social Sciences & Humanities": "1m0xEcv-WKQr8CWfHlZ5AuCWIFXAm1H5g",
    "Sciences": "1u_KRBhdZhcWQ55CyVI0v042bIpC5FQfs",
    "Management": "1VG3xY_SmhqmQ9BvSh6KvDXOptO3kHhsj",
    "Commerce": "1HMBoNkhksNpaitlBaGfq3JeoHsb_jmo-"
}

# --- 2. BACKEND API ENGINE ---
def get_google_credentials():
    """Generates clean authenticated credentials directly out of environment secrets."""
    g_sec = st.secrets["gcp_service_account"]
    processed_private_key = g_sec["private_key"].replace("\\n", "\n").replace("\n\n", "\n").strip()
    
    info_matrix = {
        "type": g_sec["type"],
        "project_id": g_sec["project_id"],
        "private_key_id": g_sec["private_key_id"],
        "private_key": processed_private_key,
        "client_email": g_sec["client_email"],
        "client_id": g_sec["client_id"],
        "auth_uri": g_sec["auth_uri"],
        "token_uri": g_sec["token_uri"],
        "auth_provider_x509_cert_url": g_sec["auth_provider_x509_cert_url"],
        "client_x509_cert_url": g_sec["client_x509_cert_url"],
        "universe_domain": g_sec["universe_domain"]
    }
    
    return service_account.Credentials.from_service_account_info(
        info_matrix,
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive.file",
            "https://www.googleapis.com/auth/drive"
        ]
    )

def upload_file_to_drive(file_bytes, file_name, mime_type, target_id, creds):
    try:
        drive_service = build('drive', 'v3', credentials=creds)
        file_metadata = {'name': file_name, 'parents': [target_id]}
        media = MediaIoBaseUpload(
            io.BytesIO(file_bytes), mimetype=mime_type, resumable=True
        )
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
    except Exception:
        return "Drive Pending"

def ai_extract_document_details(file_bytes, file_name, mime_type):
    try:
        api_key_source = st.secrets["GEMINI_API_KEY"]
        genai.configure(api_key=api_key_source)
        
        model = genai.GenerativeModel("models/gemini-1.5-flash")
        prompt = """
        Analyze this document image or PDF completely.
        Extract the values into a raw JSON map containing exactly these three keys:
        1. "date_of_event": The specific date when the event occurred or publication was issued in YYYY-MM-DD format. Look for phrases like 'held on', 'dated', or signatures. Do NOT provide today's date.
        2. "submission_type": Must be exactly one of these: FDP, Paper Presentation, Research Publication, Book Chapter, Book Publication, Workshop.
        3. "title": The precise full name or theme text of the workshop/event/paper. Do not use the file name.
        
        Provide raw text JSON output only.
        """
        response = model.generate_content(
            contents=[{"mime_type": mime_type, "data": file_bytes}, prompt]
        )
        match = re.search(r"\{.*\}", response.text, re.DOTALL)
        if match:
            return json.loads(match.group(0))
        raise ValueError("No JSON found")
    except Exception as e:
        st.sidebar.error(f"AI Matrix Log: {str(e)}")
        return {
            "date_of_event": "Check Original Document",
            "submission_type": "FDP/Workshop",
            "title": f"Review Needed: {file_name}"
        }

def get_department_sort_index(row_data):
    if len(row_data) > 5 and row_data[5] in DEPARTMENTS:
        return DEPARTMENTS.index(row_data[5])
    return 99

def fetch_existing_sheet_rows(sheets_service, sheet_range):
    try:
        res = sheets_service.spreadsheets().values().get(
            spreadsheetId=MASTER_SHEET_ID, range=sheet_range
        ).execute()
        return res.get('values', [])
    except Exception:
        return []

def update_master_sheet_rows(sheets_service, sheet_range, target_year, rows):
    sheets_service.spreadsheets().values().clear(
        spreadsheetId=MASTER_SHEET_ID, range=sheet_range
    ).execute()
    sheets_service.spreadsheets().values().update(
        spreadsheetId=MASTER_SHEET_ID, range=f"'{target_year}'!A1",
        valueInputOption="USER_ENTERED", body={'values': rows}
    ).execute()

# --- 3. FRONTEND INTERFACE ---
st.set_page_config(page_title="Portal", page_icon="🤖", layout="centered")

st.title("🏢 St. Mary's Smart Research Desk")
st.markdown("Drop your documents below to sort them beautifully.")

form_name = st.text_input("Faculty Member Name", placeholder="Your name...")
form_dept = st.selectbox("Select Department", DEPARTMENTS)
form_year = st.selectbox("Select Academic Year", ACADEMIC_YEARS)
uploaded_files = st.file_uploader("Upload files", accept_multiple_files=True)

if st.button("🚀 Submit Documents to Master Sheet"):
    if not form_name.strip():
        st.error("Name required")
    elif not uploaded_files:
        st.error("Files required")
    else:
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        try:
            creds = get_google_credentials()
            sheets_service = build('sheets', 'v4', credentials=creds)
            f_id = DEPARTMENT_FOLDERS.get(form_dept, "14Nhs3qve5vDBbIT6GmzaRue51hvTzAOG")
            
            utc_now = datetime.datetime.utcnow()
            ist_offset = datetime.timedelta(hours=5, minutes=30)
            t_now = (utc_now + ist_offset).strftime("%d-%m-%Y %H:%M:%S")
            
            for index, file in enumerate(uploaded_files):
                status_text.markdown(f"🤖 **Analyzing:** `{file.name}`...")
                file_bytes = file.read()
                
                extracted = ai_extract_document_details(file_bytes, file.name, file.type)
                drive_link = upload_file_to_drive(file_bytes, file.name, file.type, f_id, creds)
                
                new_entry = [
                    extracted.get("date_of_event", "Check Original Document"),
                    form_name.strip(),
                    extracted.get("submission_type", "FDP/Workshop"),
                    extracted.get("title", "Untitled Entry"),
                    drive_link,
                    form_dept,
                    t_now
                ]
                
                sheet_range = f"'{form_year}'!A1:G1000"
                rows = fetch_existing_sheet_rows(sheets_service, sheet_range)

                if not rows:
                    headers = ["Date", "Faculty Name", "Category", "Title", "Document Link", "Department", "Timestamp"]
                    rows = [headers, new_entry]
                else:
                    headers = rows[0]
                    data_rows = rows[1:]
                    data_rows.append(new_entry)
                    data_rows.sort(key=get_department_sort_index)
                    rows = [headers] + data_rows

                update_master_sheet_rows(sheets_service, sheet_range, form_year, rows)
                progress_bar.progress(int((index + 1) / len(uploaded_files) * 100))

            status_text.empty()
            progress_bar.empty()
            st.success(f"🎉 Complete for '{form_year}'!")
            st.balloons()
            
        except Exception as e:
            st.error(f"System Error: {str(e)}")