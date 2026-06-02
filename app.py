import streamlit as st
import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import io

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
    "2024-25", "2025-26", "2026-27", "2027-28", "2028-29", "2029-30"
]
RESEARCH_TYPES = [
    "FDP", "Workshop", "Conference Presentation", "Paper publication", "Book Chapter", "Full Book"
]
DEPARTMENT_FOLDERS = {
    "English & Languages": "14Nhs3qve5vDBbIT6GmzaRue51hvTzAOG",
    "Social Sciences & Humanities": "1m0xEcv-WKQr8CWfHlZ5AuCWIFXAm1H5g",
    "Sciences": "1u_KRBhdZhcWQ55CyVI0v042bIpC5FQfs",
    "Management": "1VG3xY_SmhqmQ9BvSh6KvDXOptO3kHhsj",
    "Commerce": "1HMBoNkhksNpaitlBaGfq3JeoHsb_jmo-"
}

USER_PASSWORD_KEYS = {
    "saikiran@stmaryscollege.in": "saikiran_pass",
    "sangeetha@stmaryscollege.in": "sangeetha_pass",
    "aditijuyal@stmaryscollege.in": "aditijuyal_pass",
    "maithry@stmaryscollege.in": "maithry_pass",
    "soumya@stmaryscollege.in": "soumya_pass",
    "rajita@stmaryscollege.in": "rajita_pass",
    "manojkanth@stmaryscollege.in": "manojkanth_pass"
}

# --- 2. BACKEND GOOGLE INTEGRATION ---
def get_google_credentials():
    g_sec = st.secrets["gcp_service_account"]
    raw_key = g_sec["private_key"]
    
    # 1. Strip raw markers, backslashes, tabs, or quotes injected by text parsers
    raw_key = raw_key.replace("\\n", "").replace("\n", "").replace("\r", "")
    raw_key = raw_key.replace("-----BEGIN PRIVATE KEY-----", "").replace("-----END PRIVATE KEY-----", "")
    clean_base64 = raw_key.strip().replace(" ", "")
    
    # 2. Re-slice the single flat string into perfect 64-character chunks (Strict standard PEM layout)
    chunks = [clean_base64[i:i+64] for i in range(0, len(clean_base64), 64)]
    
    # 3. Assemble with structural newline formatting blocks that cryptography module requires
    formatted_body = "\n".join(chunks)
    processed_private_key = f"-----BEGIN PRIVATE KEY-----\n{formatted_body}\n-----END PRIVATE KEY-----\n"

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
    return service_account.Credentials.from_service_account_info(info_matrix, scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"])

def upload_file_to_drive(file_bytes, file_name, mime_type, target_id, creds):
    try:
        drive_service = build('drive', 'v3', credentials=creds)
        file_metadata = {'name': file_name, 'parents': [target_id]}
        media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime_type, resumable=True)
        uploaded_file = drive_service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink', supportsAllDrives=True).execute()
        try:
            drive_service.permissions().create(fileId=uploaded_file.get('id'), body={'type': 'anyone', 'role': 'reader'}, supportsAllDrives=True).execute()
        except Exception:
            pass
        return uploaded_file.get('webViewLink', "Link Error")
    except Exception:
        return "Drive Pending"

def get_department_sort_index(row_data):
    if len(row_data) > 7 and row_data[7] in DEPARTMENTS:
        return DEPARTMENTS.index(row_data[7])
    return 99

# --- 3. SESSION STATE INITIALIZATION ---
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "logged_email" not in st.session_state:
    st.session_state.logged_email = ""

st.set_page_config(page_title="Faculty Portal", page_icon="📝", layout="wide")

# --- INTERFACE ROUTING: LOGIN WALL ---
if not st.session_state.authenticated:
    st.markdown("<h2 style='text-align: center;'>🔐 St. Mary's Secure Research Gateway</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center;'>Authorized Faculty Login Panel</p>", unsafe_allow_html=True)
    
    col_l1, col_l2, col_l3 = st.columns([1, 1.5, 1])
    with col_l2:
        input_email = st.text_input("College Email Address", placeholder="username@stmaryscollege.in").strip().lower()
        input_password = st.text_input("Password", type="password", placeholder="••••••••")
        
        if st.button("Sign In", type="primary", use_container_width=True):
            if input_email in USER_PASSWORD_KEYS:
                secret_key_name = USER_PASSWORD_KEYS[input_email]
                try:
                    correct_password = st.secrets[secret_key_name]
                    if input_password == correct_password:
                        st.session_state.authenticated = True
                        st.session_state.logged_email = input_email
                        st.success("Access Granted.")
                        st.rerun()
                    else:
                        st.error("Invalid password entry. Please try again.")
                except Exception:
                    st.error("Authentication value missing from system secrets configuration.")
            else:
                st.error("Access Denied: This email address is not authorized.")
    st.stop()

# --- INTERFACE ROUTING: APPLICATION WORKSPACE ---
with st.sidebar:
    st.markdown(f"**Logged in as:**\n`{st.session_state.logged_email}`")
    if st.button("🔒 Secure Sign Out", use_container_width=True):
        st.session_state.authenticated = False
        st.session_state.logged_email = ""
        st.rerun()
        
    st.markdown("---")
    st.sidebar.subheader("⚙️ Account Settings")
    with st.sidebar.expander("🔑 Change Password"):
        new_pass = st.text_input("Enter New Password", type="password", placeholder="New password...")
        confirm_pass = st.text_input("Confirm New Password", type="password", placeholder="Confirm password...")
        
        if st.button("Generate Update String", use_container_width=True):
            if not new_pass.strip():
                st.error("Password cannot be blank.")
            elif new_pass != confirm_pass:
                st.error("Passwords do not match.")
            else:
                st.success("Copy this string and share it with the administrator:")
                st.code(f'{USER_PASSWORD_KEYS[st.session_state.logged_email]} = "{new_pass.strip()}"')

st.title("🏢 St. Mary's Manual Research Logging Desk")
st.markdown("Fill out your entry matrix fields and submit them cleanly to the Master Sheet ledger.")

st.markdown("---")
st.subheader("👤 Faculty Member Profile")

col_a, col_b, col_c, col_d = st.columns(4)
with col_a:
    st.text_input("Faculty Email Address", value=st.session_state.logged_email, disabled=True)
with col_b:
    form_name = st.text_input("Faculty Member Name", placeholder="Your full name...")
with col_c:
    form_dept = st.selectbox("Select Department", DEPARTMENTS)
with col_d:
    form_year = st.selectbox("Select Academic Year", ACADEMIC_YEARS)

st.markdown("---")
st.subheader("📊 Research Matrix Log (Up to 10 Rows)")

hdr_1, hdr_2, hdr_3, hdr_4, hdr_5 = st.columns([0.6, 2.0, 3.5, 1.8, 2.5])
with hdr_1: st.markdown("**Sl No.**")
with hdr_2: st.markdown("**Research Type**")
with hdr_3: st.markdown("**Precise Title**")
with hdr_4: st.markdown("**Date of Event**")
with hdr_5: st.markdown("**Upload Document Certificate**")

row_data_collection = []

for i in range(1, 11):
    row_col1, row_col2, row_col3, row_col4, row_col5 = st.columns([0.6, 2.0, 3.5, 1.8, 2.5])
    with row_col1:
        st.markdown(f"<p style='padding-top:25px; text-align:center;'>{i}</p>", unsafe_allow_html=True)
    with row_col2:
        r_type = st.selectbox(f"Type-{i}", ["-- Select Entry --"] + RESEARCH_TYPES, label_visibility="collapsed")
    with row_col3:
        r_title = st.text_input(f"Title-{i}", placeholder="Enter title/theme text...", label_visibility="collapsed")
    with row_col4:
        r_date = st.date_input(f"Date-{i}", value=None, key=f"date_widget_{i}", label_visibility="collapsed")
    with row_col5:
        r_file = st.file_uploader(f"File-{i}", key=f"file_widget_{i}", label_visibility="collapsed")
        
    if r_type != "-- Select Entry --" or r_title.strip():
        row_data_collection.append({
            "sl_no": i,
            "type": r_type if r_type != "-- Select Entry --" else "Unspecified",
            "title": r_title.strip() if r_title.strip() else "Untitled Entry",
            "date": r_date.strftime("%Y-%m-%d") if r_date else "Check Original Document",
            "file": r_file
        })

st.markdown("---")

if st.button("🚀 Process Batch & Commit Records to Sheet", type="primary"):
    if not form_name.strip():
        st.error("Faculty Member Name field required.")
    elif not row_data_collection:
        st.error("Please fill out at least one item row in the matrix grid list.")
    else:
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        try:
            creds = get_google_credentials()
            sheets_service = build('sheets', 'v4', credentials=creds)
            f_id = DEPARTMENT_FOLDERS.get(form_dept, "14Nhs3qve5vDBbIT6GmzaRue51hvTzAOG")
            
            utc_now = datetime.datetime.utcnow()
            t_now = (utc_now + datetime.timedelta(hours=5, minutes=30)).strftime("%d-%m-%Y %H:%M:%S")
            
            sheet_range = f"'{form_year}'!A1:I1000"
            
            try:
                res = sheets_service.spreadsheets().values().get(spreadsheetId=MASTER_SHEET_ID, range=sheet_range).execute()
                existing_rows = res.get('values', [])
            except Exception:
                existing_rows = []

            if not existing_rows:
                headers = ["Sl No", "Research Type", "Title", "Date", "Document Link", "Email Address", "Faculty Name", "Department", "Timestamp"]
                data_rows = []
            else:
                headers = existing_rows[0]
                data_rows = existing_rows[1:]

            for idx, entry in enumerate(row_data_collection):
                status_text.markdown(f"💾 **Processing Entry Row {entry['sl_no']}:** `{entry['title']}`...")
                
                if entry["file"] is not None:
                    f_bytes = entry["file"].read()
                    drive_link = upload_file_to_drive(f_bytes, entry["file"].name, entry["file"].type, f_id, creds)
                else:
                    drive_link = "No File Uploaded"
                
                new_entry_record = [
                    str(entry["sl_no"]),
                    entry["type"],
                    entry["title"],
                    entry["date"],
                    drive_link,
                    st.session_state.logged_email,
                    form_name.strip(),
                    form_dept,
                    t_now
                ]
                data_rows.append(new_entry_record)
                progress_bar.progress(int((idx + 1) / len(row_data_collection) * 100))

            data_rows.sort(key=get_department_sort_index)
            final_write_payload = [headers] + data_rows

            sheets_service.spreadsheets().values().clear(spreadsheetId=MASTER_SHEET_ID, range=sheet_range).execute()
            sheets_service.spreadsheets().values().update(
                spreadsheetId=MASTER_SHEET_ID, range=f"'{form_year}'!A1", valueInputOption="USER_ENTERED", body={'values': final_write_payload}).execute()

            status_text.empty()
            progress_bar.empty()
            st.success(f"🎉 Successfully logged all structural matrix rows to the '{form_year}' sheet ledger!")
            st.balloons()
            
        except Exception as e:
            st.error(f"System Operational Mismatch: {str(e)}")
