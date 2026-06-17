import streamlit as st
import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import io
import os
import json

# Ensure python-docx is appended to your requirements.txt file
from docx import Document
from docx.shared import Pt

# --- 1. CORE SYSTEM CONFIGURATION ---
MASTER_SHEET_ID = "15wPQ9QWydGKF1OIW1QkaeXB3msRjhwiJix4ZVyf6DxA"
HR_DRIVE_FOLDER_ID = "1XFwkDsTLqYvuby6BWAJNZDt_10pBK3Hh"
IQAC_DRIVE_FOLDER_ID = "1uueAqk4PK4vEy6kc6kiPh2qz9_vjWW7l"

DEPARTMENTS = ["English & Languages", "Social Sciences & Humanities", "Sciences", "Management", "Commerce"]
ACADEMIC_YEARS = ["2024-25", "2025-26", "2026-27", "2027-28", "2028-29", "2029-30"]
MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

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
    "maithry@stmaryscollege.in": {"name": "Dr. Maithry", "secret_key": "maithry_pass"},
    "soumya@stmaryscollege.in": {"name": "Dr. Soumya", "secret_key": "soumya_pass"},
    "rajita@stmaryscollege.in": {"name": "Dr. Rajita", "secret_key": "rajita_pass"},
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
    "kanthi@stmaryscollege.in": {"name": "Dr. Kanthi Sree", "secret_key": "kanthi_pass"}
}

# --- 2. GOOGLE SERVICE INTEGRATION HANDSHAKE ---
def get_google_credentials():
    json_path = "credentials.json"
    if os.path.exists(json_path):
        with open(json_path, "r") as f:
            info_matrix = json.load(f)
            
        if "private_key" in info_matrix:
            key_content = info_matrix["private_key"]
            
            # 1. Clean out literal string-escaped '\n' sequences if present
            key_content = key_content.replace("\\n", "\n")
            
            # 2. Re-enforce standard cryptographic boundaries cleanly
            if "-----BEGIN PRIVATE KEY-----" not in key_content:
                key_content = f"-----BEGIN PRIVATE KEY-----\n{key_content}"
            if "-----END PRIVATE KEY-----" not in key_content:
                key_content = f"{key_content}\n-----END PRIVATE KEY-----\n"
                
            # 3. Strip any accidental duplicate header rows or trailing whitespace noise
            key_content = key_content.replace("-----BEGIN PRIVATE KEY-----\n-----BEGIN PRIVATE KEY-----", "-----BEGIN PRIVATE KEY-----")
            key_content = key_content.replace("-----END PRIVATE KEY-----\n-----END PRIVATE KEY-----", "-----END PRIVATE KEY-----")
            
            info_matrix["private_key"] = key_content.strip() + "\n"
            
        return service_account.Credentials.from_service_account_info(
            info_matrix, scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        )
    else:
        st.error("Critical configuration credential registry source missing (credentials.json).")
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
        return f"Upload Failed: {str(e)}"

# --- 3. THE WORD DOCUMENT NARRATIVE COMPILER ENGINE ---
def build_monthly_word_document(dept_name, active_month, active_year, creds):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    
    title_p = doc.add_paragraph()
    title_p.add_run(f"Monthly Staff Achievements Report: {active_month}, {active_year}\n").bold = True
    title_p.runs[0].font.size = Pt(14)
    
    dept_p = doc.add_paragraph()
    dept_p.add_run(f"DEPARTMENT OF {dept_name.upper()}\n").bold = True
    dept_p.runs[0].font.size = Pt(13)
    
    sheets_service = build('sheets', 'v4', credentials=creds)
    
    sections = [
        {"title": "I. Research Publications", "sheet": "Research_Database", "filter": ["Paper publication", "Book Chapter", "Full Book"], "desc": "Include journal articles, book chapters, or conference proceedings. Please specify indexing (e.g., Web of Science, Scopus, Peer-Reviewed)."},
        {"title": "II. Faculty Development Programs (FDPs) & Workshops", "sheet": "Research_Database", "filter": ["FDP", "Workshop"], "desc": "Include training programs attended or successfully completed."},
        {"title": "III. Professional Certifications & Training", "sheet": "Faculty_Achievements", "filter": ["Certification/Course"], "desc": "Include NPTEL courses, Innovation Ambassador training, or other professional certifications."},
        {"title": "IV. Paper Presentations & Resource Person Roles", "sheet": "Faculty_Achievements", "filter": ["Presentation/Resource Person"], "desc": "Include papers presented at conferences, acting as a Judge, Guest Speaker, or Facilitator for colloquiums."},
        {"title": "V. Research Milestones (For Doctoral Scholars)", "sheet": "Faculty_Achievements", "filter": ["Doctoral Milestone"], "desc": "Include milestones such as Synopsis Seminars, Pre-Ph.D. exams, or Thesis submission."},
        {"title": "VI. Awards, Honors, & Recognitions", "sheet": "Faculty_Achievements", "filter": ["Award/Honor"], "desc": "Include any special awards, titles, or professional recognitions."},
        {"title": "VII. Departmental & Institutional Contribution", "sheet": "Departmental_Student_Activities", "filter": ["Institutional Contribution"], "desc": "Include organized events, Institutional Social Responsibility (ISR) activities, or specialized student activities."}
    ]
    
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
                if len(row) >= 6:
                    if sec["sheet"] == "Research_Database":
                        row_dept, row_cat, row_month = row[1], row[2], row[13]
                    else:
                        row_dept, row_month, row_cat = row[1], row[2], row[4]
                    
                    if row_dept == dept_name and row_month == active_month and row_cat in sec["filter"]:
                        p = doc.add_paragraph(style='List Bullet')
                        
                        if sec["sheet"] == "Research_Database":
                            f_name = row[0]
                            f_cat = row[2]
                            j_type = row[3]
                            title_text = row[4]
                            pub_url = row[7]
                            pub_name = row[8]
                            pub_scope = row[9]
                            conf_scope = row[10]
                            org_body = row[11]
                            isbn_issn = row[12]
                            duration_dates = row[6]
                            
                            if f_cat in ["Paper publication", "Book Chapter", "Full Book"]:
                                narr = f'{f_name} published a {f_cat} titled "{title_text}" in {pub_name}. Journal Type: {j_type}, ISSN/ISBN: [{isbn_issn}], Scope: {pub_scope}. URL: {pub_url}'
                            else:
                                active_scope = conf_scope if (conf_scope and conf_scope != "NA") else "Institutional"
                                narr = f'{f_name} completed a {duration_dates} {active_scope} {f_cat} on "{title_text}," organized by {org_body}.'
                        else:
                            narr = row[5]
                            
                        p.add_run(narr)
                        has_data = True
                        
        if not has_data:
            doc.add_paragraph().add_run("\t- Nil -")
        doc.add_paragraph()
        
    doc_stream = io.BytesIO()
    doc.save(doc_stream)
    return doc_stream.getvalue()

# --- 4. STREAMLIT FRAMEWORK DESK ---
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "logged_email" not in st.session_state:
    st.session_state.logged_email = ""

st.set_page_config(page_title="St. Mary's Integrated Portal", page_icon="🏫", layout="wide")

if not st.session_state.authenticated:
    st.markdown("<h2 style='text-align: center;'>🔐 St. Mary's Central Achievements Gateway</h2>", unsafe_allow_html=True)
    col_l1, col_l2, col_l3 = st.columns([1, 1.5, 1])
    with col_l2:
        input_email = st.text_input("College Email Address").strip().lower()
        input_password = st.text_input("Password", type="password")
        if st.button("Sign In", type="primary", use_container_width=True):
            if input_email in FACULTY_DIRECTORY:
                secret_key_name = FACULTY_DIRECTORY[input_email]["secret_key"]
                correct_password = st.secrets.get(secret_key_name, "welcome@2026")
                if input_password == correct_password:
                    st.session_state.authenticated = True
                    st.session_state.logged_email = input_email
                    st.rerun()
                else:
                    st.error("Invalid credentials entry.")
            else:
                st.error("Email address not authorized inside profile system.")
    st.stop()

current_faculty_name = FACULTY_DIRECTORY[st.session_state.logged_email]["name"]
tab_submit, tab_document = st.tabs(["📝 Submit Achievement Log", "📊 Live Document Lounge & Analytics"])

with tab_submit:
    st.subheader("Add Monthly Achievement Entry")
    col_a, col_b, col_c = st.columns(3)
    with col_a: form_dept = st.selectbox("Select Department Focus", DEPARTMENTS)
    with col_b: form_month = st.selectbox("Reporting Month", MONTHS)
    with col_c: form_year = st.selectbox("Reporting Academic Year", ACADEMIC_YEARS)
        
    st.markdown("---")
    classification = st.selectbox("Select Entry Classification Category", [
        "-- Select Sub-Ledger Direction --",
        "🔬 Research Database (Publications, FDPs, Workshops)",
        "🏆 Faculty Profiles & Milestones (Certifications, Presentations, Ph.D. Milestones, Awards)",
        "👥 Departmental & Student Contributions"
    ])
    
    if classification != "-- Select Sub-Ledger Direction --":
        target_sheet = "Faculty_Achievements" if "Faculty Profiles" in classification else "Departmental_Student_Activities"
        specific_category = "Institutional Contribution"
        
        if "Faculty Profiles" in classification:
            specific_category = st.selectbox("Sub-Category Type", [
                "Certification/Course", 
                "Presentation/Resource Person", 
                "Doctoral Milestone", 
                "Award/Honor"
            ])

        # -------------------------------------------------------------
        # 🎯 FIX: Helper containers remain ABOVE the static form block to prevent layout disruption
        # -------------------------------------------------------------
        if "Research Database" not in classification:
            st.markdown("### 📝 Required Formatting Helper")
            
            if specific_category == "Certification/Course":
                st.warning("**Format:** `[Name], [Certification Title/Course Name], [Issuing Body], [Result/Grade/Medal if applicable].`")
                st.info("**Example:** `Mr. Roy attended a 3-day International Workshop focused on advanced research techniques, specifically \"Mastering Research Reviews and Meta-Analysis\"`")
            
            elif specific_category == "Presentation/Resource Person":
                st.warning("**Format:** `[Name], [Role: e.g., Presenter/Judge/Facilitator], \"[Topic/Title],\" [Event Name/Department], [Date].`")
                st.info("**Example:** `Dr. C. Kusuma Reddy conducted a Department Colloquium on GST Types and Return`")
            
            elif specific_category == "Doctoral Milestone":
                st.warning("**Format:** `[Name], [Milestone Achieved], \"[Research Topic],\" [University/Institution], [Date].`")
                st.info("**Example:** `Ms. Shanti has successfully completed her PHD thesis onn`")
            
            elif specific_category == "Award/Honor":
                st.warning("**Format:** `[Name], [Title of Award/Recognition], [Awarding Body/Organization], [Date].`")
                st.info("**Example:** `Dr. Vigneshwari K was officially recognized as an Innovation Ambassador at the \"Foundation Level\" by the Ministry of Education`")
            
            elif specific_category == "Institutional Contribution":
                st.warning("**Format:** `[Coordinator/Dept], [Type of Event/Activity], [Beneficiaries/Location], [Date].`")
                st.info("**Example:** `The Department of Commerce hosted the \"IPR Diaries\" event, where first-year students delivered presentations on Intellectual Property Rights`")
        # -------------------------------------------------------------

        # Stable form container setup
        with st.form("achievement_universal_form", clear_on_submit=True):
            uploaded_file = st.file_uploader("Upload Supporting Verification Document")
            
            if "Research Database" in classification:
                f_cat = st.selectbox("Category/ Research Type", ["Paper publication", "Book Chapter", "Full Book", "FDP", "Workshop"])
                j_type = st.selectbox("Journal Type", ["UGC Care listed", "Scopus", "Pubmed", "Peer Reviewed", "Other", "NA"])
                title_text = st.text_input("Title (Ttile)")
                duration_dates = st.text_input("Date Span Text (e.g., 5-day Online: June 10-14, 2026)")
                pub_url = st.text_input("Publication URL")
                pub_name = st.text_input("Publisher Name / Journal context name")
                pub_scope = st.selectbox("Publisher Scope", ["International", "National", "NA"])
                conf_scope = st.selectbox("Conference Scope", ["International", "National", "NA"])
                org_body = st.text_input("Organizing/Conducting Body")
                isbn_issn = st.text_input("ISSN/ISBN Number")
                
                submit_log = st.form_submit_button("Commit Entry to Central Cloud Repository", type="primary")
                if submit_log:
                    creds = get_google_credentials()
                    drive_link = upload_file_to_drive(uploaded_file.read(), uploaded_file.name, uploaded_file.type, [DEPARTMENT_FOLDERS[form_dept]], creds) if uploaded_file else "No File Linked"
                    
                    new_row = [
                        current_faculty_name,  # Col A: Faculty Name
                        form_dept,             # Col B: Department
                        f_cat,                 # Col C: Category/ Research Type
                        j_type,                # Col D: Journal Type
                        title_text,            # Col E: Ttile
                        drive_link,            # Col F: Document Link
                        duration_dates,        # Col G: Date
                        pub_url,               # Col H: Publication URL
                        pub_name,              # Col I: Publisher Name
                        pub_scope,             # Col J: Publisher Scope
                        conf_scope,            # Col K: Conference Scope
                        org_body,              # Col L: Organizing/Conducting Body
                        isbn_issn,             # Col M: ISSN/ISBN Number
                        form_month             # Col N: Reporting Month Helper
                    ]
                    
                    build('sheets', 'v4', credentials=creds).spreadsheets().values().append(
                        spreadsheetId=MASTER_SHEET_ID, range="'Research_Database'!A1", valueInputOption="USER_ENTERED", insertDataOption="INSERT_ROWS", body={"values": [new_row]}
                    ).execute()
                    st.success("🎉 Structured Research Entry compiled into database ledger completely!")
                    
            else:
                narrative_input = st.text_area("Enter Achievement Narrative Text Statement String", placeholder="Write your paragraph matching the sample pattern display block above...")
                
                submit_log = st.form_submit_button("Commit Entry to Central Cloud Repository", type="primary")
                if submit_log:
                    if not narrative_input.strip():
                        st.error("Input Error: The achievement narrative text statement block cannot be left empty.")
                    else:
                        creds = get_google_credentials()
                        drive_link = upload_file_to_drive(uploaded_file.read(), uploaded_file.name, uploaded_file.type, [DEPARTMENT_FOLDERS[form_dept]], creds) if uploaded_file else "No File Linked"
                        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        new_row = [timestamp, form_dept, form_month, form_year, specific_category, narrative_input.strip(), current_faculty_name, drive_link]
                        build('sheets', 'v4', credentials=creds).spreadsheets().values().append(
                            spreadsheetId=MASTER_SHEET_ID, range=f"'{target_sheet}'!A1", valueInputOption="USER_ENTERED", insertDataOption="INSERT_ROWS", body={"values": [new_row]}
                        ).execute()
                        st.success(f"🎉 Achievement string appended to the `{target_sheet}` database ledger!")

with tab_document:
    st.subheader("Central HR Document Engine Dashboard Workspace")
    col_d1, col_d2, col_d3 = st.columns(3)
    with col_d1: view_dept = st.selectbox("Target Department File Scope", DEPARTMENTS, key="vd1")
    with col_d2: view_month = st.selectbox("Target Month Scope", MONTHS, key="vm1")
    with col_d3: view_year = st.selectbox("Target Year Scope", ACADEMIC_YEARS, key="vy1")
        
    if st.button("🏗️ Construct & Synchronize Automated Monthly Document Package", use_container_width=True, type="primary"):
        creds = get_google_credentials()
        with st.spinner("Assembling structured records from sheets sub-matrices..."):
            docx_bytes = build_monthly_word_document(view_dept, view_month, view_year, creds)
            file_name_string = f"Monthly_Staff_Achievements_Report_{view_dept.replace(' ', '_')}_{view_month}_{view_year}.docx"
            
            destination_sync_folders = [DEPARTMENT_FOLDERS[view_dept], HR_DRIVE_FOLDER_ID, IQAC_DRIVE_FOLDER_ID]
            upload_file_to_drive(docx_bytes, file_name_string, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", destination_sync_folders, creds)
            
            st.success(f"🎯 Document synchronized into Drive ecosystem folders automatically!")
            st.download_button(label="📥 Download Report File Asset Directly", data=docx_bytes, file_name=file_name_string, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
