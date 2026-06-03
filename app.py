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
JOURNAL_TYPES = [
    "UGC Care listed", "Scopus", "Pubmed", "Peer Reviewed", "Other"
]
SCOPES = [
    "National", "International"
]

FACULTY_DIRECTORY = {
    "saikiran@stmaryscollege.in": {"name": "Dr. Saikiran", "secret_key": "saikiran_pass"},
    "sangeetha@stmaryscollege.in": {"name": "Dr. Sangeetha", "secret_key": "sangeetha_pass"},
    "aditijuyal@stmaryscollege.in": {"name": "Prof. Aditi Juyal", "secret_key": "aditijuyal_pass"},
    "maithry@stmaryscollege.in": {"name": "Dr. Maithry", "secret_key": "maithry_pass"},
    "soumya@stmaryscollege.in": {"name": "Prof. Soumya", "secret_key": "soumya_pass"},
    "rajita@stmaryscollege.in": {"name": "Dr. Rajita", "secret_key": "rajita_pass"},
    "manojkanth@stmaryscollege.in": {"name": "Dr. Manoj Kanth", "secret_key": "manojkanth_pass"}
}

DEPARTMENT_FOLDERS = {
    "English & Languages": "14Nhs3qve5vDBbIT6GmzaRue51hvTzAOG",
    "Social Sciences & Humanities": "1m0xEcv-WKQr8CWfHlZ5AuCWIFXAm1H5g",
    "Sciences": "1u_KRBhdZhcWQ55CyVI0v042bIpC5FQfs",
    "Management": "1VG3xY_SmhqmQ9BvSh6KvDXOptO3kHhsj",
    "Commerce": "1HMBoNkhksNpaitlBaGfq3JeoHsb_jmo-"
}

# --- 2. BACKEND GOOGLE INTEGRATION ---
def get_google_credentials():
    g_sec = st.secrets["gcp_service_account"]
    raw_key = g_sec["private_key"]
    
    clean_base64 = raw_key.replace("\\n", "").replace("\n", "").replace("\r", "")
    clean_base64 = clean_base64.replace("-----BEGIN PRIVATE KEY-----", "")
    clean_base64 = clean_base64.replace("-----END PRIVATE KEY-----", "")
    clean_base64 = clean_base64.strip().replace(" ", "")
    
    chunks = [clean_base64[i:i+64] for i in range(0, len(clean_base64), 64)]
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
    if len(row_data) > 6 and row_data[6] in DEPARTMENTS:
        return DEPARTMENTS.index(row_data[6])
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
            if input_email in FACULTY_DIRECTORY:
                secret_key_name = FACULTY_DIRECTORY[input_email]["secret_key"]
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
top_col1, top_col2 = st.columns([6.5, 3.5])

with top_col1:
    st.title("🏢 St. Mary's Manual Research Logging Desk")
    st.markdown(f"**Logged in as:** `{st.session_state.logged_email}`")

with top_col2:
    st.markdown("<div style='padding-top: 10px;'></div>", unsafe_allow_html=True)
    acc_expander = st.expander("⚙️ Account Settings & Security")
    with acc_expander:
        with st.form("password_change_form", clear_on_submit=False):
            st.markdown("##### 🔑 Change Account Password")
            new_pass = st.text_input("Enter New Password", type="password", placeholder="New password...")
            confirm_pass = st.text_input("Confirm New Password", type="password", placeholder="Confirm password...")
            submit_pass = st.form_submit_button("Generate Update Code", use_container_width=True)
            
            if submit_pass:
                if not new_pass.strip():
                    st.error("Password cannot be blank.")
                elif new_pass != confirm_pass:
                    st.error("Passwords do not match.")
                else:
                    st.success("Share this string with the administrator:")
                    st.code(f'{FACULTY_DIRECTORY[st.session_state.logged_email]["secret_key"]} = "{new_pass.strip()}"')
        
        if st.button("🔒 Secure Sign Out", type="secondary", use_container_width=True):
            st.session_state.authenticated = False
            st.session_state.logged_email = ""
            st.rerun()

st.markdown("Fill out your entry matrix fields and submit them cleanly to the Master Sheet ledger.")
st.markdown("---")

st.subheader("👤 Faculty Member Profile")
current_faculty_name = FACULTY_DIRECTORY[st.session_state.logged_email]["name"]

col_a, col_b, col_c, col_d = st.columns(4)
with col_a:
    st.text_input("Faculty Email Address", value=st.session_state.logged_email, disabled=True)
with col_b:
    form_name = st.text_input("Faculty Member Name", value=current_faculty_name, disabled=True)
with col_c:
    form_dept = st.selectbox("Select Department", ["-- Select Department --"] + DEPARTMENTS)
with col_d:
    form_year = st.selectbox("Select Academic Year", ["-- Select Academic Year --"] + ACADEMIC_YEARS)

st.markdown("---")
st.subheader("📊 Research Matrix Log (Up to 10 Rows)")

row_data_collection = []

for i in range(1, 11):
    st.markdown(f"#### 🔘 Entry Row Record #{i}")
    
    col1, col2, col3 = st.columns([2.0, 5.0, 5.0])
    with col1:
        r_type = st.selectbox(f"Research Type", ["-- Select Entry --"] + RESEARCH_TYPES, key=f"type_{i}")
    with col2:
        r_title = st.text_input(f"Precise Title / Theme text", placeholder="Enter theme text...", key=f"title_{i}")
    with col3:
        r_url = st.text_input(f"URL of Publication (Optional)", placeholder="Paste web link if available...", key=f"url_{i}")
        
    j_type = "N/A"
    p_name = "N/A"
    p_scope = "N/A"
    c_scope = "N/A"
    org_body = "N/A"
    
    if r_type in ["FDP", "Workshop"]:
        sub_col1, sub_col2, sub_col3 = st.columns([3.0, 3.0, 6.0])
        with sub_col1:
            r_date_from = st.date_input(f"Date From", value=None, key=f"date_from_{i}")
        with sub_col2:
            r_date_to = st.date_input(f"Date To", value=None, key=f"date_to_{i}")
        with sub_col3:
            org_body = st.text_input(f"Organised By", placeholder="Enter hosting college/organization name...", key=f"org_{i}")
            
    elif r_type == "Conference Presentation":
        sub_col1, sub_col2, sub_col3, sub_col4 = st.columns([2.5, 2.5, 3.5, 3.5])
        with sub_col1:
            r_date_from = st.date_input(f"Date From", value=None, key=f"date_from_{i}")
        with sub_col2:
            r_date_to = st.date_input(f"Date To", value=None, key=f"date_to_{i}")
        with sub_col3:
            c_scope = st.selectbox(f"Conference Classification", SCOPES, key=f"cscope_{i}")
        with sub_col4:
            org_body = st.text_input(f"Conducted By", placeholder="Enter university/body name...", key=f"cond_{i}")
            
    else:
        sub_col1 = st.columns(1)[0]
        with sub_col1:
            r_date_from = st.date_input(f"Date of Event", value=None, key=f"date_single_{i}")
        r_date_to = None
        
        if r_type == "Paper publication":
            sub_col_pub = st.columns(1)[0]
            with sub_col_pub:
                j_type = st.selectbox(f"Journal Listing Index", JOURNAL_TYPES, key=f"jtype_{i}")
                
        elif r_type in ["Book Chapter", "Full Book"]:
            sub_col_bk1, sub_col_bk2 = st.columns([6.0, 4.0])
            with sub_col_bk1:
                p_name = st.text_input(f"Name of the Publisher", placeholder="Enter publishing house...", key=f"pname_{i}")
            with sub_col_bk2:
                p_scope = st.selectbox(f"Publisher Classification", SCOPES, key=f"pscope_{i}")
            
    r_file = st.file_uploader(f"Upload Document Certificate Support Asset", key=f"file_{i}")
    st.markdown("<hr style='margin:10px 0px; border-top: 1px dashed #ddd;' />", unsafe_allow_html=True)
        
    if r_type != "-- Select Entry --" or r_title.strip():
        row_data_collection.append({
            "sl_no": i,
            "type": r_type if r_type != "-- Select Entry --" else "Unspecified",
            "title": r_title.strip() if r_title.strip() else "Untitled Entry",
            "pub_url": r_url.strip() if r_url.strip() else "No Link Provided",
            "date_from": r_date_from,
            "date_to": r_date_to,
            "file": r_file,
            "j_type": j_type,
            "p_name": p_name,
            "p_scope": p_scope,
            "c_scope": c_scope,
            "org_body": org_body.strip() if (isinstance(org_body, str) and org_body.strip()) else org_body
        })

st.markdown("---")

if st.button("🚀 Process Batch & Commit Records to Sheet", type="primary", use_container_width=True):
    if form_dept == "-- Select Department --":
        st.error("Form Validation Error: Please select your explicit Department before submitting.")
    elif form_year == "-- Select Academic Year --":
        st.error("Form Validation Error: Please select the explicit Academic Year before submitting.")
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
            
            sheet_range = f"'{form_year}'!A1:P1000"
            
            try:
                res = sheets_service.spreadsheets().values().get(spreadsheetId=MASTER_SHEET_ID, range=sheet_range).execute()
                existing_rows = res.get('values', [])
            except Exception:
                existing_rows = []

            if not existing_rows:
                headers = [
                    "Date From", "Date To", "Faculty Name", "Category", "Title", "Document Link", 
                    "Department", "Timestamp", "Year", "Month", "Publication URL", 
                    "Journal Type", "Publisher Name", "Publisher Scope", "Conference Scope", "Organizing/Conducting Body"
                ]
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
                
                str_date_from = entry["date_from"].strftime("%Y-%m-%d") if entry["date_from"] else "Check Attachment"
                str_date_to = entry["date_to"].strftime("%Y-%m-%d") if entry["date_to"] else str_date_from
                
                str_year = entry["date_from"].strftime("%Y") if entry["date_from"] else ""
                str_month = entry["date_from"].strftime("%B") if entry["date_from"] else ""
                
                new_entry_record = [
                    str_date_from,                # Column A: Date From
                    str_date_to,                  # Column B: Date To
                    form_name.strip(),            # Column C: Faculty Name
                    entry["type"],                # Column D: Category
                    entry["title"],               # Column E: Title
                    drive_link,                   # Column F: Document Link
                    form_dept,                    # Column G: Department
                    t_now,                        # Column H: Timestamp
                    str_year,                     # Column I: Year
                    str_month,                    # Column J: Month
                    entry["pub_url"],             # Column K: Publication URL
                    entry["j_type"],              # Column L: Journal Type
                    entry["p_name"],              # Column M: Publisher Name
                    entry["p_scope"],             # Column N: Publisher Scope
                    entry["c_scope"],             # Column O: Conference Scope
                    entry["org_body"]             # Column P: Organizing/Conducting Body
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
            st.success(f"🎉 Successfully logged all organized entries sorted by Department layout to the master ledger!")
            st.balloons()
            
        except Exception as e:
            st.error(f"System Operational Mismatch: {str(e)}")
