import streamlit as st
import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import io
import os
import json
import base64

# Ensure python-docx is appended to your requirements.txt file
from docx import Document
from docx.shared import Pt

# --- 1. CORE SYSTEM CONFIGURATION FROM SECRETS ---
MASTER_SHEET_ID = st.secrets["MASTER_SHEET_ID"]

DEPARTMENTS = ["English & Languages", "Social Sciences & Humanities", "Sciences", "Management", "Commerce"]
ACADEMIC_YEARS = ["2024-25", "2025-26", "2026-27", "2027-28", "2028-29", "2029-30"]
MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

# Exact order map used to sort rows mathematically before pushing to Google Sheets
DEPT_SORT_ORDER = {dept: index for index, dept in enumerate(DEPARTMENTS)}

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
    # --- NEWLY RECONFIGURED FACULTY ADDITIONS ---
    "timee@stmaryscollege.in": {"name": "Dr. Timee Ronra Shimray", "secret_key": "timee_pass"},
    "ismail@stmaryscollege.in": {"name": "Mr. Ismail C", "secret_key": "ismail_pass"},
    "aksharasingh@stmaryscollege.in": {"name": "Dr. Akshara Singh", "secret_key": "akshara_pass"},
    "vasantharao@stmaryscollege.in": {"name": "Mr. Vasantha Rao B", "secret_key": "vasantharao_pass"}
}

# --- 2. GOOGLE SERVICE INTEGRATION HANDSHAKE ---
def get_google_credentials():
    try:
        b64_string = st.secrets["BASE64_GCP_CREDENTIALS"]
        decoded_bytes = base64.b64decode(b64_string)
        info_matrix = json.loads(decoded_bytes)
        return service_account.Credentials.from_service_account_info(
            info_matrix, 
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
        return f"Upload Failed: {str(e)}"

def append_and_sort_sheet_by_department(sheet_name, new_row, dept_column_index, creds):
    try:
        sheets_service = build('sheets', 'v4', credentials=creds)
        
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=MASTER_SHEET_ID, range=f"'{sheet_name}'!A1:N2000"
        ).execute()
        rows = result.get('values', [])
        
        if not rows:
            sheets_service.spreadsheets().values().update(
                spreadsheetId=MASTER_SHEET_ID, range=f"'{sheet_name}'!A1",
                valueInputOption="USER_ENTERED", body={"values": [new_row]}
            ).execute()
            return

        header = rows[0]
        data_rows = rows[1:]
        data_rows.append(new_row)
        
        def sort_key_resolver(row):
            if len(row) <= dept_column_index:
                return len(DEPARTMENTS)
            dept_name = row[dept_column_index]
            return DEPT_SORT_ORDER.get(dept_name, len(DEPARTMENTS))

        data_rows.sort(key=sort_key_resolver)
        sorted_matrix = [header] + data_rows
        
        sheets_service.spreadsheets().values().clear(
            spreadsheetId=MASTER_SHEET_ID, range=f"'{sheet_name}'!A1:N2000"
        ).execute()
        
        sheets_service.spreadsheets().values().update(
            spreadsheetId=MASTER_SHEET_ID, range=f"'{sheet_name}'!A1",
            valueInputOption="USER_ENTERED", body={"values": sorted_matrix}
        ).execute()
        
    except Exception as e:
        st.error(f"Rearrangement Sorting Algorithm Pipeline Failure Block: {str(e)}")

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
        {"title": "I. Research Publications & Paper Presentations", "sheet": "Research_Database", "filter": ["Paper publication", "Book Chapter", "Full Book", "Paper Presentation"], "desc": "Include journal articles, book chapters, full books, or papers presented at conferences."},
        {"title": "II. Faculty Development Programs (FDPs) & Workshops", "sheet": "Research_Database", "filter": ["FDP", "Workshop"], "desc": "Include training programs attended or successfully completed."},
        {"title": "III. Professional Certifications & Training", "sheet": "Faculty_Achievements", "filter": ["Certification/Course"], "desc": "Include NPTEL courses, Innovation Ambassador training, or other professional certifications."},
        {"title": "IV. Resource Person Roles & Invited Lectures", "sheet": "Faculty_Achievements", "filter": ["Presentation/Resource Person"], "desc": "Include acting as a Judge, Guest Speaker, Keynote Facilitator, or Resource Person for academic colloquiums."},
        {"title": "V. Research Milestones (For Doctoral Scholars)", "sheet": "Faculty_Achievements", "filter": ["Doctoral Milestone"], "desc": "Include milestones such as Synopsis Seminars, Pre-Ph.D. exams, or Thesis submission."},
        {"title": "VI. Awards, Honors, & Recognitions", "sheet": "Faculty_Achievements", "filter": ["Award/Honor"], "desc": "Include any special awards, titles, or professional recognitions."},
        {"title": "VII. Departmental & Student Contribution", "sheet": "Student_Activities", "filter": ["Institutional Contribution"], "desc": "Include organized events, Institutional Social Responsibility (ISR) activities, or specialized student activities."}
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
                            elif f_cat == "Paper Presentation":
                                narr = f'{f_name} presented a research paper titled "{title_text}" at the conference organized by {org_body or pub_name} ({duration_dates or "NA"}). Scope: {conf_scope}.'
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
        "🔬 Research Database (Publications, Paper Presentations, FDPs, Workshops)",
        "🏆 Faculty Profiles & Milestones (Certifications, Resource Person Roles, Ph.D. Milestones, Awards)",
        "👥 Departmental & Student Contributions"
    ])
    
    if classification != "-- Select Sub-Ledger Direction --":
        if "Research Database" in classification:
            target_sheet = "Research_Database"
            specific_category = "Research"
        elif "Faculty Profiles" in classification:
            target_sheet = "Faculty_Achievements"
            specific_category = st.selectbox("Sub-Category Type", [
                "Certification/Course", 
                "Presentation/Resource Person",
                "Doctoral Milestone", 
                "Award/Honor"
            ])
        else:
            target_sheet = "Student_Activities"
            specific_category = "Institutional Contribution"

        if "Research Database" not in classification:
            st.markdown("### 📝 Required Formatting Helper")
            
            if specific_category == "Certification/Course":
                st.warning("**Format:** `[Name], [Certification Title/Course Name], [Issuing Body], [Result/Grade/Medal if applicable].`")
                st.info("**Example:** `Mr. Roy attended a 3-day International Workshop focused on advanced research techniques, specifically \"Mastering Research Reviews and Meta-Analysis\"`")
            
            elif specific_category == "Presentation/Resource Person":
                st.warning("**Format (Resource Person):** `[Name], [Role: e.g., Guest Speaker/Judge/Facilitator], \"[Topic/Title],\" [Organizing Event Name/Department/Institution], [Date].`")
                st.info("**Example:** `Dr. C. Kusuma Reddy acted as a Resource Person and conducted a Department Colloquium on GST Types and Returns.`")
            
            elif specific_category == "Doctoral Milestone":
                st.warning("**Format:** `[Name], [Milestone Achieved], \"[Research Topic],\" [University/Institution], [Date].`")
                st.info("**Example:** `Ms. Shanti has successfully completed her PHD thesis on...`")
            
            elif specific_category == "Award/Honor":
                st.warning("**Format:** `[Name], [Title of Award/Recognition], [Awarding Body/Organization], [Date].`")
                st.info("**Example:** `Dr. Vigneshwari K was officially recognized as an Innovation Ambassador at the \"Foundation Level\" by the Ministry of Education`")
            
            elif specific_category == "Institutional Contribution":
                st.warning("**Format:** `[Coordinator/Dept], [Type of Event/Activity], [Beneficiaries/Location], [Date].`")
                st.info("**Example:** `The Department of Commerce hosted the \"IPR Diaries\" event, where first-year students delivered presentations on Intellectual Property Rights`")

        with st.form("achievement_universal_form", clear_on_submit=True):
            uploaded_file = st.file_uploader("Upload Supporting Verification Document")
            
            if "Research Database" in classification:
                f_cat = st.selectbox("Category/ Research Type", ["Paper publication", "Book Chapter", "Full Book", "Paper Presentation", "FDP", "Workshop"])
                j_type = st.selectbox("Journal / Event Type", ["UGC Care listed", "Scopus", "Pubmed", "Peer Reviewed", "Conference", "Other", "NA"])
                title_text = st.text_input("Title of Paper / Book / Topic")
                duration_dates = st.text_input("Date Span Text (e.g., June 10-14, 2026 / Single Day: June 17, 2026)")
                pub_url = st.text_input("Publication / Event URL (If applicable)")
                pub_name = st.text_input("Publisher Name / Journal Name / Conference Name")
                pub_scope = st.selectbox("Publisher Scope (For books/papers)", ["International", "National", "NA"])
                conf_scope = st.selectbox("Conference / Event Scope", ["International", "National", "State", "Institutional", "NA"])
                org_body = st.text_input("Organizing/Conducting Body")
                isbn_issn = st.text_input("ISSN/ISBN Number (If applicable)")
                
                submit_log = st.form_submit_button("Commit Entry to Central Cloud Repository", type="primary")
                if submit_log:
                    creds = get_google_credentials()
                    drive_link = upload_file_to_drive(uploaded_file.read(), uploaded_file.name, uploaded_file.type, [DEPARTMENT_FOLDERS[form_dept]], creds) if uploaded_file else "No File Linked"
                    
                    new_row = [
                        current_faculty_name,
                        form_dept,
                        f_cat,
                        j_type,
                        title_text,
                        drive_link,
                        duration_dates,
                        pub_url,
                        pub_name,
                        pub_scope,
                        conf_scope,
                        org_body,
                        isbn_issn,
                        form_month
                    ]
                    
                    append_and_sort_sheet_by_department("Research_Database", new_row, 1, creds)
                    st.success("🎉 Structured Research Entry compiled into database ledger and perfectly sorted by department hierarchy!")
                    
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
                        
                        append_and_sort_sheet_by_department(target_sheet, new_row, 1, creds)
                        st.success(f"🎉 Achievement string appended to the `{target_sheet}` ledger and perfectly sorted by department hierarchy!")

with tab_document:
    st.subheader("Central Document Engine Dashboard Workspace")
    col_d1, col_d2, col_d3 = st.columns(3)
    with col_d1: view_dept = st.selectbox("Target Department File Scope", DEPARTMENTS, key="vd1")
    with col_d2: view_month = st.selectbox("Target Month Scope", MONTHS, key="vm1")
    with col_d3: view_year = st.selectbox("Target Year Scope", ACADEMIC_YEARS, key="vy1")
        
    if st.button("🏗️ Construct Automated Monthly Document Package", use_container_width=True, type="primary"):
        creds = get_google_credentials()
        with st.spinner("Assembling structured records from sheets..."):
            docx_bytes = build_monthly_word_document(view_dept, view_month, view_year, creds)
            file_name_string = f"Monthly_Staff_Achievements_Report_{view_dept.replace(' ', '_')}_{view_month}_{view_year}.docx"
            
            upload_file_to_drive(docx_bytes, file_name_string, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", [DEPARTMENT_FOLDERS[view_dept]], creds)
            
            st.success(f"🎯 Document synchronized into your Department Drive folder automatically!")
            st.download_button(label="📥 Download Report File Asset Directly", data=docx_bytes, file_name=file_name_string, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
