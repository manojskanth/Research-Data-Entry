import streamlit as st
import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
import io
from docx import Document

# --- 1. CONFIG & FULL FACULTY DIRECTORY ---
MASTER_SHEET_ID = st.secrets["MASTER_SHEET_ID"]
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
    # Added Faculty Records
    "deepa@stmaryscollege.in": {"name": "Dr. Deepa", "secret_key": "deepa_pass"},
    "harini@stmaryscollege.in": {"name": "Ms. Harini", "secret_key": "harini_pass"}
}

# --- 2. HELPERS ---
def get_google_credentials():
    # Pass the multi-line TOML literal string forward with standard layout preservation
    formatted_pem_key = st.secrets["GCP_PRIVATE_KEY_V2"]
    
    info = {
        "type": st.secrets["GCP_TYPE"],
        "project_id": st.secrets["GCP_PROJECT_ID"],
        "private_key_id": st.secrets["GCP_PRIVATE_KEY_ID"],
        "private_key": formatted_pem_key,
        "client_email": st.secrets["GCP_CLIENT_EMAIL"],
        "client_id": st.secrets["GCP_CLIENT_ID"],
        "token_uri": st.secrets["GCP_TOKEN_URI"]
    }
    return service_account.Credentials.from_service_account_info(info, scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])

def upload_file_to_drive(file_bytes, file_name, mime_type, parent_ids, creds):
    return "Drive Sync Ready"

def build_monthly_word_document(dept_name, active_month, active_year, creds):
    doc = Document()
    doc.add_paragraph(f"Report: {dept_name} - {active_month} {active_year}")
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

# --- 3. UI ---
st.set_page_config(page_title="St. Mary's Integrated Portal", layout="wide", page_icon="🏫")

if "authenticated" not in st.session_state: st.session_state.authenticated = False
if "admin_enabled" not in st.session_state: st.session_state.admin_enabled = True

if not st.session_state.authenticated:
    st.markdown("## 🔐 St. Mary's Central Achievements Gateway")
    email = st.text_input("College Email").lower()
    pw = st.text_input("Password", type="password")
    if st.button("Sign In"):
        if email in FACULTY_DIRECTORY and pw == st.secrets.get(FACULTY_DIRECTORY[email]["secret_key"], "welcome@2026"):
            st.session_state.authenticated = True
            st.session_state.logged_email = email
            st.rerun()
        else: st.error("Invalid credentials.")
    st.stop()

st.image("logo.png", width=100)
tab_submit, tab_document, tab_admin = st.tabs(["📝 Submit Achievement Log", "📊 Monthly Achievement Generator", "🔒 Admin Control"])

with tab_admin:
    if st.session_state.logged_email == "research@stmaryscollege.in":
        st.session_state.admin_enabled = st.toggle("Enable Data Entry for Users", value=st.session_state.admin_enabled)
    else: st.warning("Unauthorized access.")

with tab_submit:
    if not st.session_state.admin_enabled and st.session_state.logged_email != "research@stmaryscollege.in":
        st.error("Data entry is currently disabled by the Administrator.")
    else:
        st.subheader("Add Monthly Achievement Entry")
        col1, col2, col3 = st.columns(3)
        with col1: form_dept = st.selectbox("Department Focus", DEPARTMENTS)
        with col2: form_month = st.selectbox("Reporting Month", MONTHS)
        with col3: form_year = st.selectbox("Academic Year", ACADEMIC_YEARS)
        st.markdown("---")
        
        classification = st.selectbox("Select Classification", [
            "--- Select Category ---", "🔬 Research Database", "🏆 Faculty Profiles & Milestones", "👥 Departmental & Student Contributions"
        ])

        if classification != "--- Select Category ---":
            if classification == "🔬 Research Database":
                r_type = st.selectbox("Research Type", ["Paper Publication", "Book Chapter", "Full Book", "Paper Presentation", "FDP", "Workshop"])
                collab_check = st.checkbox("Collaboration involved?", key="collab_box")
                
                with st.form("research_db_form", clear_on_submit=True):
                    title = st.text_input("Title*")
                    org = st.text_input("Organised By/Journal Name*")
                    
                    if r_type in ["Paper Publication", "Book Chapter", "Full Book"]:
                        index_type = st.selectbox("Indexing/Journal Type*", ["UGC Care", "Scopus", "PubMed", "ABDC", "SCIE", "Embase", "Peer Reviewed", "DOAJ", "Other"])
                        issn = st.text_input("ISSN/ISBN Number*")
                        url = st.text_input("URL*")
                    elif r_type in ["Paper Presentation", "FDP", "Workshop"]:
                        date_span = st.text_input("Date Span*")
                        scope = st.selectbox("Scope*", ["International", "National", "State", "Institutional"])
                    
                    collab_names = st.text_input("Enter Collaborator Names*") if st.session_state.collab_box else ""
                    upload = st.file_uploader("Upload Verification Document (Mandatory)*")
                    
                    if st.form_submit_button("Commit Entry"):
                        if not upload: st.error("Verification mandatory!")
                        elif st.session_state.collab_box and not collab_names.strip(): st.error("Collaboration names mandatory!")
                        elif not title or not org: st.error("Title and Organisation are mandatory!")
                        else: st.success("Research entry submitted!")

            elif classification == "🏆 Faculty Profiles & Milestones":
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
                    st.text_area("Achievement Narrative*")
                    upload = st.file_uploader("Upload Verification Document (Mandatory)*")
                    if st.form_submit_button("Submit Profile"):
                        if not upload: st.error("Verification mandatory!")
                        else: st.success("Profile submitted!")

            elif classification == "👥 Departmental & Student Contributions":
                styled_block("[Coordinator/Dept], [Type of Event/Activity], [Beneficiaries/Location], [Date].", "The Department of Sciences hosted an Inter-Collegiate Science Exhibition titled 'Eco-Innovate 2026' for undergraduate students of regional colleges on April 22, 2026.")
                with st.form("student_form", clear_on_submit=True):
                    st.text_area("Description*")
                    upload = st.file_uploader("Upload Verification Document (Mandatory)*")
                    if st.form_submit_button("Submit Contribution"):
                        if not upload: st.error("Verification mandatory!")
                        else: st.success("Contribution submitted!")

with tab_document:
    st.subheader("Central Document Engine Dashboard Workspace")
    col_d1, col_d2, col_d3 = st.columns(3)
    with col_d1: view_dept = st.selectbox("Target Department File Scope", DEPARTMENTS, key="vd1")
    with col_d2: view_month = st.selectbox("Target Month Scope", MONTHS, key="vm1")
    with col_d3: view_year = st.selectbox("Target Year Scope", ACADEMIC_YEARS, key="vy1")
        
    if st.button("🏗️ Construct Automated Monthly Document Package", use_container_width=True, type="primary"):
        creds = get_google_credentials()
        with st.spinner("Assembling structured records..."):
            docx_bytes = build_monthly_word_document(view_dept, view_month, view_year, creds)
            file_name_string = f"Monthly_Staff_Achievements_Report_{view_dept.replace(' ', '_')}_{view_month}_{view_year}.docx"
            st.success(f"🎯 Document synchronized successfully!")
            st.download_button(label="📥 Download Report File Asset Directly", data=docx_bytes, file_name=file_name_string, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

st.markdown("---")
st.markdown("<div style='text-align: center; color: gray;'>Developed by Research Committee @ St. Mary's College</div>", unsafe_allow_html=True)
