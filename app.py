import streamlit as st
import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import io
import json
import base64
from docx import Document
from docx.shared import Pt

# --- 1. CORE SYSTEM CONFIGURATION ---
MASTER_SHEET_ID = st.secrets["MASTER_SHEET_ID"]

DEPARTMENTS = ["English & Languages", "Social Sciences & Humanities", "Sciences", "Management", "Commerce", "IQAC", "Research & Innovation", "Physical Education"]
COMMITTEES_CELLS_CLUBS = ["Alumni", "Anti-Ragging", "Disciplinary", "Equal Opportunity", "Grievance Redressal", "Internal Complaints", "Scholarship", "Library", "Placement", "Public Relations", "Women Empowerment", "Student Activity Clubs", "IIC"]
ACADEMIC_YEARS = ["2024-25", "2025-26", "2026-27", "2027-28", "2028-29", "2029-30"]
MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

DEPT_SORT_ORDER = {dept: index for index, dept in enumerate(DEPARTMENTS)}

# Dedicated folder vault repository mapping targets
COMMITTEE_FOLDER_ID = "1pzrbGsViKtzsQYBPt-p9ZqQ5WXbzdHcW"

DEPARTMENT_FOLDERS = {
    "English & Languages": "14Nhs3qve5vDBbIT6GmzaRue51hvTzAOG",
    "Social Sciences & Humanities": "1m0xEcv-WKQr8CWfHlZ5AuCWIFXAm1H5g",
    "Sciences": "1u_KRBhdZhcWQ55CyVI0v042bIpC5FQfs",
    "Management": "1VG3xY_SmhqmQ9BvSh6KvDXOptO3kHhsj",
    "Commerce": "1HMBoNkhksNpaitlBaGfq3JeoHsb_jmo-"
}

FACULTY_DIRECTORY = {
    "saikiran@stmaryscollege.in": {"name": "Dr. Saikiran", "secret_key": "saikiran_pass"},
    "sangeetha@stmaryscollege.in": {"name": "Dr. Sangeetha", "secret_key": "sangeetha_pass"},
    "aditijuyal@stmaryscollege.in": {"name": "Prof. Aditi Juyal", "secret_key": "aditijuyal_pass"},
    "maithry@stmaryscollege.in": {"name": "Dr. Maithry Shinde", "secret_key": "maithry_pass"},
    "soumya@stmaryscollege.in": {"name": "Dr. Soumya K", "secret_key": "soumya_pass"},
    "rajita@stmaryscollege.in": {"name": "Dr. Rajita Anand Singh", "secret_key": "rajita_pass"},
    "manojkanth@stmaryscollege.in": {"name": "Dr. Manoj Kanth", "secret_key": "manojkanth_pass"},
    "swathi@stmaryscollege.in": {"name": "Dr. Swathi", "secret_key": "swathi_pass"},
    "padmaleela@stmaryscollege.in": {"name": "Ms. Padmaleela", "secret_key": "padmaleela_pass"},
    "nsrinath@stmarycollege.in": {"name": "Dr. Srinath Naganathan", "secret_key": "nsrinath_pass"},
    "sowjanya@stmaryscollege.in": {"name": "Ms. D. Sowjanya", "secret_key": "sowjanya_pass"},
    "sandhyarani@stmaryscollege.in": {"name": "Ms. A. Sandhya Rani", "secret_key": "sandhyarani_pass"},
    "ragasudha@stmaryscollege.in": {"name": "Ms. Raga Sudha Jonnada", "secret_key": "ragasudha_pass"},
    "rajyalakshmi@stmaryscollege.in": {"name": "Ms. Rajalakshmi", "secret_key": "rajyalakshmi_pass"},
    "mahanta@stmaryscollege.in": {"name": "Mr. Mahanta Chauhan", "secret_key": "mahanta_pass"},
    "sharol@stmaryscollege.in": {"name": "Dr. Sharol Sebastian", "secret_key": "sharol_pass"},
    "deepthipriya@stmaryscollege.in": {"name": "Dr. Deepthi Priya", "secret_key": "deepthipriya_pass"},
    "satabdi@stmaryscollege.in": {"name": "Dr. Satabdi Roy", "secret_key": "satabdi_pass"},
    "shima@stmaryscollege.in": {"name": "Ms. Shima A.N", "secret_key": "shima_pass"},
    "anuvictor@stmaryscollege.in": {"name": "Ms. Anu Victor", "secret_key": "anuvictor_pass"},
    "sadbhavana@stmaryscollege.in": {"name": "Ms. Sadbhavana Sharat", "secret_key": "sadbhavana_pass"},
    "sriveda@stmaryscollege.in": {"name": "Ms. Sriveda Baswapoor", "secret_key": "sriveda_pass"},
    "rameshk@stmaryscollege.in": {"name": "Dr. Ramesh Kumar", "secret_key": "rameshk_pass"},
    "shivakumar@stmaryscollege.in": {"name": "Mr. Shiva Kumar Reddy", "secret_key": "shivakumar_pass"},
    "anamika@stmaryscollege.in": {"name": "Dr. Anamika Sukul", "secret_key": "anamika_pass"},
    "arunjose@stmaryscollege.in": {"name": "Mr. Arun B Jose", "secret_key": "arunjose_pass"},
    "elisheba@stmaryscollege.in": {"name": "Ms. P. Elisheba", "secret_key": "elisheba_pass"},
    "debanjalee@stmaryscollege.in": {"name": "Dr. Debanjalee Bose", "secret_key": "debanjalee_pass"},
    "kirtibdnr@stmaryscollege.in": {"name": "Dr. Kirti", "secret_key": "kirti_pass"},
    "shikhasharma@stmaryscollege.in": {"name": "Dr. Shikha Sharma", "secret_key": "shikha_pass"},
    "himani@stmaryscollege.in": {"name": "Dr. Himani", "secret_key": "himani_pass"},
    "roy@stmaryscollege.in": {"name": "Mr. MSS Roy", "secret_key": "roy_pass"},
    "phebi@stmaryscollege.in": {"name": "Ms. Phebi", "secret_key": "phebi_pass"},
    "vigneshwari@stmaryscollege.in": {"name": "Dr. Vigneswari", "secret_key": "vigneshwari_pass"},
    "nagarjuna@stmaryscollege.in": {"name": "Dr. Nagarjuna", "secret_key": "nagarjuna_pass"},
    "pavitrambika@stmaryscollege.in": {"name": "Dr. Pavitrambika", "secret_key": "pavitrambika_pass"},
    "anuradhaemani@stmaryscollege.in": {"name": "Dr. Anuradha", "secret_key": "anuradha_pass"},
    "kanthi@stmaryscollege.in": {"name": "Dr. Kanthi Sree", "secret_key": "kanthi_pass"},
    "timee@stmaryscollege.in": {"name": "Dr. Timee Ronra Shimray", "secret_key": "timee_pass"},
    "ismail@stmaryscollege.in": {"name": "Mr. Ismail C", "secret_key": "ismail_pass"},
    "aksharasingh@stmaryscollege.in": {"name": "Dr. Akshara Singh", "secret_key": "akshara_pass"},
    "vasantharao@stmaryscollege.in": {"name": "Mr. Vasantha Rao B", "secret_key": "vasantharao_pass"},
    "gisageorge@stmaryscollege.in": {"name": "Ms. Gisa George", "secret_key": "gisageorge_pass"},
    "research@stmaryscollege.in": {"name": "Research Admin", "secret_key": "research_pass"},
    "deepa@stmaryscollege.in": {"name": "Dr. Deepa", "secret_key": "deepa_pass"},
    "chrislenina@stmaryscollege.in: {"name": "Dr. Chris", "secret_key": "chris_pass"},
    "harini@stmaryscollege.in": {"name": "Ms. Harini", "secret_key": "harini_pass"},
    "jayalakshmi@stmaryscollege.in": {"name": "Ms. Jayalakshmi D", "secret_key": "jayalakshmi_pass"},
    "rupini@stmaryscollege.in": {"name": "Ms. B. Rupini", "secret_key": "rupini_pass"},
    "manali@stmaryscollege.in": {"name": "Ms. Manali Manoj Manwadkar", "secret_key": "manali_pass"},
    "kusuma@stmaryscollege.in": {"name": "Dr. Kusuma C", "secret_key": "kusuma_pass"},
    "bikshapathi@stmaryscollege.in": {"name": "Mr. Bikshapathi M", "secret_key": "bikshapathi_pass"},
    "vijaybhaskar@stmaryscollege.in": {"name": "Mr. Vijay Bhaskar Reddy", "secret_key": "vijaybhaskar_pass"},
    "poojasharma@stmaryscollege.in": {"name": "Ms. Pooja Sharma", "secret_key": "poojasharma_pass"},
    "kavithathakur@stmaryscollege.in": {"name": "Dr. Kavitha Thakur", "secret_key": "kavithathakur_pass"},
    "priyamishra@stmaryscollege.in": {"name": "Dr. Priya Mishra", "secret_key": "priyamishra_pass"},
    "deepaagraval@stmaryscollege.in": {"name": "Ms. Deepa Agraval", "secret_key": "deepaagraval_pass"}
}

# --- 2. GOOGLE SERVICE INTEGRATION HANDSHAKE ---
def get_google_credentials():
    try:
        encoded_json = st.secrets["GCP_COMPLETE_B64"]
        decoded_json_string = base64.b64decode(encoded_json.encode('utf-8')).decode('utf-8')
        info = json.loads(decoded_json_string)
        return service_account.Credentials.from_service_account_info(
            info, 
            scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        )
    except Exception as e:
        st.error(f"Ecosystem Verification Block Error: {str(e)}")
        st.stop()

def upload_file_to_drive(file_bytes, file_name, mime_type, parent_ids, creds):
    try:
        drive_service = build('drive', 'v3', credentials=creds)
        links = []
        for p_id in parent_ids:
            file_metadata = {'name': file_name, 'parents': [p_id]}
            media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime_type, resumable=True)
            uploaded = drive_service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink', supportsAllDrives=True).execute()
            try:
                drive_service.permissions().create(fileId=uploaded.get('id'), body={'type': 'anyone', 'role': 'reader'}, supportsAllDrives=True).execute()
            except:
                pass
            links.append(uploaded.get('webViewLink', ""))
        return links[0] if links else "Drive Error"
    except Exception as e:
        st.warning(f"⚠️ Drive Sync Notification: Written cleanly to spreadsheet layout.")
        return "Pending Folder Permissions Link"

def append_and_sort_sheet_by_department(sheet_name, new_row, dept_column_index, creds):
    try:
        sheets_service = build('sheets', 'v4', credentials=creds)
        result = sheets_service.spreadsheets().values().get(spreadsheetId=MASTER_SHEET_ID, range=f"'{sheet_name}'!A1:N2000").execute()
        rows = result.get('values', [])
        
        if not rows:
            sheets_service.spreadsheets().values().update(
                spreadsheetId=MASTER_SHEET_ID, range=f"'{sheet_name}'!A1",
                valueInputOption="USER_ENTERED", body={"values": [new_row]}
            ).execute()
            return

        header, data_rows = rows[0], rows[1:]
        data_rows.append(new_row)
        
        if sheet_name in ["Research_Database", "Faculty_Achievements", "Student_Activities"]:
            data_rows.sort(key=lambda r: DEPT_SORT_ORDER.get(r[dept_column_index], len(DEPARTMENTS)) if len(r) > dept_column_index else len(DEPARTMENTS))
        else:
            data_rows.sort(key=lambda r: r[dept_column_index] if len(r) > dept_column_index else "")
            
        sorted_matrix = [header] + data_rows
        
        sheets_service.spreadsheets().values().clear(spreadsheetId=MASTER_SHEET_ID, range=f"'{sheet_name}'!A1:N2000").execute()
        sheets_service.spreadsheets().values().update(
            spreadsheetId=MASTER_SHEET_ID, range=f"'{sheet_name}'!A1",
            valueInputOption="USER_ENTERED", body={"values": sorted_matrix}
        ).execute()
    except Exception as e:
        st.error(f"Sorting Error: {str(e)}")

# --- 3. THE WORD DOCUMENT NARRATIVE COMPILER ENGINE ---
def build_monthly_word_document(name_focus, active_month, active_year, creds):
    doc = Document()
    doc.styles['Normal'].font.name = 'Times New Roman'
    doc.styles['Normal'].font.size = Pt(12)
    
    title_p = doc.add_paragraph()
    title_p.add_run(f"Monthly Achievements Summary Report: {active_month}, {active_year}\n").bold = True
    title_p.runs[0].font.size = Pt(14)
    
    scope_p = doc.add_paragraph()
    
    sheets_service = build('sheets', 'v4', credentials=creds)
    month_map = {"jan": "january", "feb": "february", "mar": "march", "apr": "april", "may": "may", "jun": "june", "jul": "july", "aug": "august", "sep": "september", "oct": "october", "nov": "november", "dec": "december"}
    target_month_clean = str(active_month).strip().lower()

    # --- 🏢 UNIFIED COMMITTEE MASTER DOCUMENT DISPATCH (FROM 4TH SHEET) ---
    if name_focus == "Committees / Cells / Clubs":
        scope_p.add_run("COMMITTEES / CELLS / CLUBS MASTER DOSSIER\n").bold = True
        scope_p.runs[0].font.size = Pt(13)
        
        doc.add_paragraph().add_run("I. Consolidated Committee Activity Logs & Event Narratives").bold = True
        doc.add_paragraph().add_run("Chronological record of organized events, initiatives, and execution statements across all campus cells.").font.italic = True
        
        try:
            res = sheets_service.spreadsheets().values().get(spreadsheetId=MASTER_SHEET_ID, range="'Committees_Cells_Clubs'!A1:F1000").execute()
            rows = res.get('values', [])
        except:
            rows = []
            
        has_data = False
        if len(rows) > 1:
            for row in rows[1:]:
                if len(row) >= 5:
                    row_comm, row_faculty, row_month, row_year, row_narrative = row[0], row[1], row[2], row[3], row[4]
                    normalized_row_month = str(row_month).strip().lower()
                    
                    month_match = (target_month_clean == normalized_row_month) or \
                                  (target_month_clean[:3] in normalized_row_month) or \
                                  (normalized_row_month in month_map and month_map[normalized_row_month] == target_month_clean)
                    
                    if month_match and str(row_year).strip() == str(active_year).strip():
                        p = doc.add_paragraph(style='List Bullet')
                        p.add_run(f"[{row_comm}] ").bold = True
                        p.add_run(f"{row_narrative} (In-charge: {row_faculty})")
                        has_data = True
                        
        if not has_data:
            doc.add_paragraph().add_run("\t- Nil -")
            
        doc_stream = io.BytesIO()
        doc.save(doc_stream)
        return doc_stream.getvalue()

    # --- 🔬 STANDARD DEPARTMENT ENGINE SCOPE (FROM ORIGINAL 3 SHEETS) ---
    scope_p.add_run(f"DEPARTMENT OF {name_focus.upper()}\n").bold = True
    scope_p.runs[0].font.size = Pt(13)
    
    # --- MODIFIED DATA EXTRACTION LAYER ---
    sections = [
        {"title": "I. Research Publications & Paper Presentations", "sheet": "Research_Database", "filter": ["Paper Publication", "Book Chapter", "Full Book", "Paper Presentation"], "desc": "Include journal articles, book chapters, full books, or papers presented at conferences."},
        {"title": "II. Faculty Development Programs (FDPs) & Workshops", "sheet": "Research_Database", "filter": ["FDP", "Workshop"], "desc": "Include training programs attended or successfully completed."},
        {"title": "III. Professional Certifications & Training", "sheet": "Faculty_Achievements", "filter": ["Certification/Course"], "desc": "Include NPTEL courses, Innovation Ambassador training, or other professional certifications."},
        {"title": "IV. Resource Person Roles & Invited Lectures", "sheet": "Faculty_Achievements", "filter": ["Presentation/Resource Person"], "desc": "Include acting as a Judge, Guest Speaker, Keynote Facilitator, or Resource Person for academic colloquiums."},
        {"title": "V. Research Milestones (For Doctoral Scholars)", "sheet": "Faculty_Achievements", "filter": ["Doctoral Milestone"], "desc": "Include milestones such as Synopsis Seminars, Pre-Ph.D. exams, or Thesis submission."},
        {"title": "VI. Awards, Honors, & Recognitions", "sheet": "Faculty_Achievements", "filter": ["Award/Honor"], "desc": "Include any special awards, titles, or professional recognitions."},
        {"title": "VII. Departmental & Student Contribution", "sheet": "Student_Activities", "filter": ["Institutional Contribution"], "desc": "Include organized events, Institutional Social Responsibility (ISR) activities, or specialized student activities."}
    ]
    
    def pad_row(target_row, required_length=15):
        return target_row + [""] * (required_length - len(target_row))

    for sec in sections:
        doc.add_paragraph().add_run(sec["title"]).bold = True
        doc.add_paragraph().add_run(sec["desc"]).font.italic = True
        
        try:
            res = sheets_service.spreadsheets().values().get(spreadsheetId=MASTER_SHEET_ID, range=f"'{sec['sheet']}'!A1:N1000").execute()
            rows = res.get('values', [])
        except: 
            rows = []
            
        has_data = False
        if len(rows) > 1:
            for row in rows[1:]:
                if len(row) >= 2:
                    padded = pad_row(row, required_length=15)
                    
                    # Exact tracking offsets assigned based on the structural column layouts of each tab
                    if sec["sheet"] == "Research_Database":
                        row_dept, row_cat, row_month = padded[1], padded[2], padded[13]
                    elif sec["sheet"] == "Faculty_Achievements":
                        row_dept, row_cat, row_month = padded[1], padded[4], padded[2]
                    elif sec["sheet"] == "Student_Activities":
                        row_dept, row_cat, row_month = padded[1], padded[4], padded[2]
                    else:
                        row_dept, row_cat, row_month = padded[0], padded[4], padded[2]
                    
                    normalized_row_month = str(row_month).strip().lower()
                    
                    month_match = (target_month_clean == normalized_row_month) or \
                                  (target_month_clean[:3] in normalized_row_month) or \
                                  (normalized_row_month in month_map and month_map[normalized_row_month] == target_month_clean)
                    
                    if str(row_dept).strip().lower() == str(name_focus).strip().lower() and month_match and \
                       any(str(row_cat).strip().lower() == str(f).strip().lower() for f in sec["filter"]):
                        
                        p = doc.add_paragraph(style='List Bullet')
                        if sec["sheet"] == "Research_Database":
                            f_name, f_cat, j_type, title_text, pub_url, pub_name, pub_scope, conf_scope, org_body, isbn_issn, duration_dates = \
                                padded[0], padded[2], padded[3], padded[4], padded[7], padded[8], padded[9], padded[10], padded[11], padded[12], padded[6]
                            
                            if f_cat in ["Paper Publication", "Book Chapter", "Full Book"]:
                                narr = f'{f_name} published a {f_cat} titled "{title_text}" in {pub_name}. Journal Type: {j_type}, ISSN/ISBN: [{isbn_issn}], Scope: {pub_scope}. URL: {pub_url}'
                            elif f_cat == "Paper Presentation":
                                narr = f'{f_name} presented a research paper titled "{title_text}" at the conference organized by {org_body or pub_name} ({duration_dates or "NA"}). Scope: {conf_scope}.'
                            else:
                                narr = f'{f_name} completed a {duration_dates} {conf_scope if (conf_scope and conf_scope != "NA") else "Institutional"} {f_cat} on "{title_text}," organized by {org_body}.'
                        else:
                            narr = padded[5]
                        
                        p.add_run(narr)
                        has_data = True
                        
        if not has_data:
            doc.add_paragraph().add_run("\t- Nil -")
        doc.add_paragraph()
        
    doc_stream = io.BytesIO()
    doc.save(doc_stream)
    return doc_stream.getvalue()

def styled_block(format_text, example_text):
    html_string = f"""
<div style="background-color: #FFFFFF; padding: 16px; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); border: 1px solid #EAECEF; margin-bottom: 20px; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;">
    <div style="display: flex; align-items: flex-start; margin-bottom: 14px;">
        <div style="background-color: #E8EAF6; color: #1A237E; font-weight: 700; font-size: 11px; text-transform: uppercase; letter-spacing: 0.8px; padding: 4px 8px; border-radius: 4px; margin-right: 12px; min-width: 70px; text-align: center; border-left: 3px solid #1A237E;">Format</div>
        <div style="color: #2C3E50; font-size: 14px; line-height: 1.5; font-weight: 500;">{format_text}</div>
    </div>
    <div style="height: 1px; background-color: #F1F3F5; margin: 12px 0;"></div>
    <div style="display: flex; align-items: flex-start;">
        <div style="background-color: #E8F5E9; color: #1B5E20; font-weight: 700; font-size: 11px; text-transform: uppercase; letter-spacing: 0.8px; padding: 4px 8px; border-radius: 4px; margin-right: 12px; min-width: 70px; text-align: center; border-left: 3px solid #2E7D32;">Example</div>
        <div style="color: #455A64; font-size: 14px; line-height: 1.5; font-style: italic; font-weight: 500;">{example_text}</div>
    </div>
</div>
""".strip()
    st.markdown(html_string, unsafe_allow_html=True)

# --- 4. STREAMLIT FRAMEWORK DESK ---
if "authenticated" not in st.session_state: st.session_state.authenticated = False
if "logged_email" not in st.session_state: st.session_state.logged_email = ""
if "admin_enabled" not in st.session_state: st.session_state.admin_enabled = True

st.set_page_config(page_title="St. Mary's Integrated Portal", page_icon="🏫", layout="wide")

if not st.session_state.authenticated:
    _, img_col, _ = st.columns([2, 1, 2])
    with img_col:
        st.image("logo.png", use_container_width=True)
        
    st.markdown("<h2 style='text-align: center;'>St. Mary's Central Achievements Portal</h2>", unsafe_allow_html=True)
    _, col_l2, _ = st.columns([1, 1.5, 1])
    with col_l2:
        input_email = st.text_input("College Email Address").strip().lower()
        input_password = st.text_input("Password", type="password")
        if st.button("Sign In", type="primary", use_container_width=True):
            if input_email in FACULTY_DIRECTORY:
                if input_password == st.secrets.get(FACULTY_DIRECTORY[input_email]["secret_key"], "welcome@2026"):
                    st.session_state.authenticated, st.session_state.logged_email = True, input_email
                    st.rerun()
                else: st.error("Invalid credentials entry.")
            else: st.error("Email address not authorized inside profile system.")
    st.stop()

# --- HEADER WORKSPACE WITH LOGOUT TOOL ---
current_faculty_name = FACULTY_DIRECTORY[st.session_state.logged_email]["name"]

logo_col, header_col, logout_col = st.columns([1, 7, 1.5])
with logo_col:
    st.image("logo.png", width=65)
with header_col:
    st.markdown(f"### Welcome back, **{current_faculty_name}**")
with logout_col:
    if st.button("🚪 Log Out", type="secondary", use_container_width=True):
        st.session_state.authenticated = False
        st.session_state.logged_email = ""
        st.rerun()

tab_submit, tab_document, tab_admin = st.tabs(["📝 Submit Achievement Log", "📊 Monthly Achievement Generator", "🔒 Admin Control"])

with tab_admin:
    if st.session_state.logged_email == "research@stmaryscollege.in":
        st.toggle(
            "Enable Data Entry for Users", 
            key="admin_toggle_widget", 
            value=st.session_state.get("admin_enabled", True)
        )
        st.session_state.admin_enabled = st.session_state.admin_toggle_widget
    else: 
        st.warning("Unauthorized access.")

with tab_submit:
    is_locked = not st.session_state.get("admin_enabled", True)
    is_admin = st.session_state.get("logged_email") == "research@stmaryscollege.in"

    if is_locked and not is_admin:
        st.error("🔒 Data entry is currently disabled by the Administrator.")
    else:
        st.subheader("Add Monthly Achievement Entry")
        
        scope_type = st.radio("Select Reporting Scope*", ["Department", "Committee / Cell / Club"], horizontal=True)
        
        col1, col2, col3 = st.columns(3)
        with col1:
            if scope_type == "Department":
                form_focus = st.selectbox("Department Focus*", DEPARTMENTS)
            else:
                form_focus = st.selectbox("Committees / Cells / Clubs Focus*", COMMITTEES_CELLS_CLUBS)
        with col2: 
            form_month = st.selectbox("Reporting Month*", MONTHS)
        with col3: 
            form_year = st.selectbox("Academic Year*", ACADEMIC_YEARS)
            
        st.markdown("---")
        
        if scope_type == "Department":
            classification = st.selectbox("Select Classification", [
                "--- Select Category ---", 
                "🔬 Research Database", 
                "🏆 Faculty Profiles & Milestones", 
                "👥 Departmental & Student Contributions"
            ])

            if classification != "--- Select Category ---":
                creds = get_google_credentials()
                
                if classification == "🔬 Research Database":
                    r_type = st.selectbox("Research Type", ["Paper Publication", "Book Chapter", "Full Book", "Paper Presentation", "FDP", "Workshop"])
                    collab_check = st.checkbox("Collaboration involved?", key="collab_box")
                    
                    with st.form("research_db_form", clear_on_submit=True):
                        title = st.text_input("Title*")
                        org = st.text_input("Organised By/Journal Name*")
                        
                        if r_type in ["Paper Publication", "Book Chapter", "Full Book"]:
                            index_type = st.selectbox("Indexing/Journal Type*", ["UGC Care Listed", "Scopus", "PubMed", "Peer Reviewed", "DOAJ", "ABDC", "SCIE", "Embase"])
                            issn = st.text_input("ISSN/ISBN Number*")
                            url = st.text_input("URL*")
                            date_span, scope = "NA", "NA"
                        elif r_type in ["Paper Presentation", "FDP", "Workshop"]:
                            date_span = st.text_input("Date Span*")
                            scope = st.selectbox("Scope*", ["International", "National", "State", "Institutional"])
                            index_type, issn, url = "NA", "NA", "NA"
                        
                        collab_names = st.text_input("Enter Collaborator Names*") if st.session_state.collab_box else ""
                        upload = st.file_uploader("Upload Verification Document (Mandatory)*")
                        
                        if st.form_submit_button("Commit Entry"):
                            if not st.session_state.get("admin_enabled", True) and st.session_state.logged_email != "research@stmaryscollege.in":
                                st.error("Submission rejected: Data entry is currently disabled.")
                            elif not upload: st.error("Verification mandatory!")
                            elif st.session_state.collab_box and not collab_names.strip(): st.error("Collaboration names mandatory!")
                            elif not title or not org: st.error("Title and Organisation are mandatory!")
                            else:
                                drive_folder_id = DEPARTMENT_FOLDERS.get(form_focus, "1HMBoNkhksNpaitlBaGfq3JeoHsb_jmo-")
                                drive_link = upload_file_to_drive(upload.read(), upload.name, upload.type, [drive_folder_id], creds)
                                new_row = [current_faculty_name, form_focus, r_type, index_type, title, drive_link, date_span, url, org, scope, scope, org, issn, form_month]
                                append_and_sort_sheet_by_department("Research_Database", new_row, 1, creds)
                                st.success("🎉 Research entry submitted successfully!")

                elif classification == "🏆 Faculty Profiles & Milestones":
                    target_sheet = "Faculty_Achievements"
                    subtype = st.selectbox("Select Profile Subtype", ["Certification/Course", "Presentation/Resource Person", "Doctoral Milestone", "Award/Honor"])
                    if subtype == "Certification/Course": 
                        styled_block("[Name], [Certification Title/Course Name], [Issuing Body], [Result/Grade/Medal if applicable].", "Mr. MSS Roy successfully completed an 8-week NPTEL certification course in 'Advanced Corporate Governance' with an Elite Silver Medal, organized by IIT Madras.")
                    elif subtype == "Presentation/Resource Person": 
                        styled_block("[Name], [Role: Guest Speaker/Judge/Facilitator], '[Topic/Title],' [Organizing Event Name/Department/Institution], [Date].", "Dr. Rajita Anand Singh acted as a Resource Person and delivered an invited lecture on 'Emerging Trends in Literary Criticism' for the National Colloquium organized by the Department of English, St. Mary's College on June 15, 2026.")
                    elif subtype == "Doctoral Milestone": 
                        styled_block("[Name], [Milestone Achieved], '[Research Topic],' [University/Institution], [Date].", "Ms. Shima A.N successfully completed her Ph.D. Viva-Voce examination for her doctoral thesis titled 'A Comprehensive Evaluation of Cloud Workloads' at Osmania University.")
                    elif subtype == "Award/Honor": 
                        styled_block("[Name], [Title of Award/Recognition], [Awarding Body/Organization], [Date].", "Dr. Deepthi Priya was conferred with the 'Best Faculty Researcher Award 2026' by the Institute of Scholar Recognitions on May 12, 2026.")
                    
                    with st.form("faculty_form", clear_on_submit=True):
                        narrative_input = st.text_area("Achievement Narrative*")
                        upload = st.file_uploader("Upload Verification Document (Mandatory)*")
                        if st.form_submit_button("Submit Profile"):
                            if not st.session_state.get("admin_enabled", True) and st.session_state.logged_email != "research@stmaryscollege.in":
                                st.error("Submission rejected: Data entry is currently disabled.")
                            elif not upload or not narrative_input.strip(): st.error("Verification and narrative statement mandatory!")
                            else:
                                drive_folder_id = DEPARTMENT_FOLDERS.get(form_focus, "1HMBoNkhksNpaitlBaGfq3JeoHsb_jmo-")
                                drive_link = upload_file_to_drive(upload.read(), upload.name, upload.type, [drive_folder_id], creds)
                                new_row = [datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), form_focus, form_month, form_year, subtype, narrative_input.strip(), current_faculty_name, drive_link]
                                append_and_sort_sheet_by_department(target_sheet, new_row, 1, creds)
                                st.success("🎉 Profile submitted!")

                elif classification == "👥 Departmental & Student Contributions":
                    target_sheet = "Student_Activities"
                    styled_block("[Coordinator/Dept], [Type of Event/Activity], [Beneficiaries/Location], [Date].", "The Department of Sciences hosted an Inter-Collegiate Science Exhibition titled 'Eco-Innovate 2026' for undergraduate students of regional colleges on April 22, 2026.")
                    with st.form("student_form", clear_on_submit=True):
                        description = st.text_area("Description*")
                        upload = st.file_uploader("Upload Verification Document (Mandatory)*")
                        if st.form_submit_button("Submit Contribution"):
                            if not st.session_state.get("admin_enabled", True) and st.session_state.logged_email != "research@stmaryscollege.in":
                                st.error("Submission rejected: Data entry is currently disabled.")
                            elif not upload or not description.strip(): st.error("Verification and description mandatory!")
                            else:
                                drive_folder_id = DEPARTMENT_FOLDERS.get(form_focus, "1HMBoNkhksNpaitlBaGfq3JeoHsb_jmo-")
                                drive_link = upload_file_to_drive(upload.read(), upload.name, upload.type, [drive_folder_id], creds)
                                new_row = [datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), form_focus, form_month, form_year, "Institutional Contribution", description.strip(), current_faculty_name, drive_link]
                                append_and_sort_sheet_by_department(target_sheet, new_row, 1, creds)
                                st.success("🎉 Contribution submitted!")

        else:
            creds = get_google_credentials()
            target_sheet = "Committees_Cells_Clubs"
            styled_block("[Committee Name], organized [Event Type/Activity Details] on [Date].", "The Placement Cell coordinated a campus recruitment drive with Deloitte for final year commerce students on May 18, 2026.")
            
            with st.form("committees_ledger_form", clear_on_submit=True):
                narrative_input = st.text_area("Narrative Log Description*")
                event_date = st.date_input("Date of Event Activity*", value=datetime.date.today())
                upload = st.file_uploader("Upload Verification Document (Mandatory)*")
                
                if st.form_submit_button("Commit Committee Record"):
                    if not st.session_state.get("admin_enabled", True) and st.session_state.logged_email != "research@stmaryscollege.in":
                        st.error("Submission rejected: Data entry is currently disabled.")
                    elif not upload or not narrative_input.strip(): 
                        st.error("Log Description and verification attachment are strictly mandatory fields!")
                    else:
                        # Redirecting all committee uploads into your specified folder vault
                        drive_link = upload_file_to_drive(upload.read(), upload.name, upload.type, [COMMITTEE_FOLDER_ID], creds)
                        new_row = [
                            form_focus, 
                            current_faculty_name, 
                            form_month, 
                            form_year, 
                            narrative_input.strip(), 
                            str(event_date)
                        ]
                        append_and_sort_sheet_by_department(target_sheet, new_row, 0, creds)
                        st.success(f"🎉 Structured Activity Log written to '{target_sheet}' sheet successfully!")

with tab_document:
    st.subheader("Central Document Engine Dashboard Workspace")
    
    col_d1, col_d2, col_d3 = st.columns(3)
    with col_d1:
        view_focus = st.selectbox(
            "Target Department / Scope Scope", 
            DEPARTMENTS + ["Committees / Cells / Clubs"], 
            key="vd1"
        )
            
    with col_d2: view_month = st.selectbox("Target Month Scope", MONTHS, key="vm1")
    with col_d3: view_year = st.selectbox("Target Year Scope", ACADEMIC_YEARS, key="vy1")
        
    if st.button("🏗️ Construct Automated Monthly Document Package", use_container_width=True, type="primary"):
        creds = get_google_credentials()
        with st.spinner("Assembling structured records from sheets..."):
            docx_bytes = build_monthly_word_document(view_focus, view_month, view_year, creds)
            file_name_string = f"Monthly_Achievements_Report_{view_focus.replace(' ', '_')}_{view_month}_{view_year}.docx"
            
            # Dynamic destination choice: send to specific department folder or the committee folder vault
            target_folder = DEPARTMENT_FOLDERS.get(view_focus, COMMITTEE_FOLDER_ID) if view_focus != "Committees / Cells / Clubs" else COMMITTEE_FOLDER_ID
            upload_file_to_drive(docx_bytes, file_name_string, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", [target_folder], creds)
            
            st.success(f"🎯 Document synchronized into your Drive repository folder automatically!")
            st.download_button(label="📥 Download Report File Asset Directly", data=docx_bytes, file_name=file_name_string, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

st.markdown("---")
st.markdown("<div style='text-align: center; color: olive;'>Developed by Research Committee @ St. Mary's College</div>", unsafe_allow_html=True)
