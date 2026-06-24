import streamlit as st
import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
import io
from docx import Document

# --- 1. CONFIG & FACULTY ---
MASTER_SHEET_ID = st.secrets["MASTER_SHEET_ID"]
DEPARTMENTS = ["English & Languages", "Social Sciences & Humanities", "Sciences", "Management", "Commerce"]
ACADEMIC_YEARS = ["2024-25", "2025-26", "2026-27", "2027-28", "2028-29", "2029-30"]
MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

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
    "research@stmaryscollege.in": {"name": "Research Admin", "secret_key": "research_pass"}
}

# --- 2. AUTH ---
def get_google_credentials():
    clean_key = st.secrets["GCP_PRIVATE_KEY"].replace("\\n", "\n")
    info = {
        "type": st.secrets["GCP_TYPE"], "project_id": st.secrets["GCP_PROJECT_ID"],
        "private_key_id": st.secrets["GCP_PRIVATE_KEY_ID"], "private_key": clean_key,
        "client_email": st.secrets["GCP_CLIENT_EMAIL"], "client_id": st.secrets["GCP_CLIENT_ID"],
        "token_uri": st.secrets["GCP_TOKEN_URI"]
    }
    return service_account.Credentials.from_service_account_info(info, scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])

# --- 3. UI ---
st.set_page_config(page_title="St. Mary's Integrated Portal", layout="wide", page_icon="🏫")

if "authenticated" not in st.session_state: st.session_state.authenticated = False
if "admin_enabled" not in st.session_state: st.session_state.admin_enabled = True

if not st.session_state.authenticated:
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
        st.session_state.admin_enabled = st.toggle("Enable Data Entry", value=st.session_state.admin_enabled)
    else: st.warning("Unauthorized access.")

with tab_submit:
    st.subheader("Add Monthly Achievement Entry")
    classification = st.selectbox("Select Classification", [
        "--- Select Category ---", "🔬 Research Database", "🏆 Faculty Profiles & Milestones", "👥 Departmental & Student Contributions"
    ])
    
    if classification == "🔬 Research Database":
        r_type = st.selectbox("Research Type", ["Paper Publication", "Book Chapter", "Full Book", "Paper Presentation", "FDP", "Workshop"])
        with st.form("research_db_form", clear_on_submit=True):
            title = st.text_input("Title*")
            org = st.text_input("Organised By/Journal Name*")
            if r_type in ["Paper Publication", "Book Chapter", "Full Book"]:
                issn = st.text_input("ISSN/ISBN Number*")
                url = st.text_input("URL*")
            elif r_type in ["Paper Presentation", "FDP", "Workshop"]:
                date_span = st.text_input("Date Span*")
                scope = st.selectbox("Scope*", ["International", "National", "State", "Institutional"])
            
            collab_check = st.checkbox("Collaboration involved?", key="collab_box")
            collab_names = st.text_input("Enter Collaborator Names*") if st.session_state.collab_box else ""
            upload = st.file_uploader("Upload Verification*")
            
            if st.form_submit_button("Submit"):
                if not upload: st.error("Verification document is mandatory!")
                elif st.session_state.collab_box and not collab_names.strip(): st.error("Collaboration names are mandatory!")
                elif not title or not org: st.error("Title and Organisation are mandatory!")
                elif r_type in ["Paper Publication", "Book Chapter", "Full Book"] and (not issn or not url): st.error("ISSN and URL are mandatory!")
                else: st.success(f"{r_type} submitted successfully!")

    elif classification == "🏆 Faculty Profiles & Milestones":
        with st.form("faculty_form", clear_on_submit=True):
            st.text_area("Achievement Narrative*")
            upload = st.file_uploader("Upload Verification*")
            if st.form_submit_button("Submit"):
                if not upload: st.error("Verification mandatory!")
                else: st.success("Profile submitted!")

    elif classification == "👥 Departmental & Student Contributions":
        with st.form("student_form", clear_on_submit=True):
            st.text_area("Description*")
            upload = st.file_uploader("Upload Verification*")
            if st.form_submit_button("Submit"):
                if not upload: st.error("Verification mandatory!")
                else: st.success("Contribution submitted!")

with tab_document:
    st.subheader("Monthly Achievement Generator")
    if st.button("🏗️ Construct Automated Monthly Document Package"):
        st.info("Generating report...")

st.markdown("---")
st.markdown("<div style='text-align: center; color: gray;'>Developed by Research Committee @ St. Mary's College</div>", unsafe_allow_html=True)
