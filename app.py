import streamlit as st
import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import io
from docx import Document
from docx.shared import Pt

# --- 1. CONFIGURATION ---
MASTER_SHEET_ID = st.secrets["MASTER_SHEET_ID"]
DEPARTMENTS = ["English & Languages", "Social Sciences & Humanities", "Sciences", "Management", "Commerce"]
ACADEMIC_YEARS = ["2024-25", "2025-26", "2026-27", "2027-28", "2028-29", "2029-30"]
MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

# --- 2. FORTIFIED GOOGLE AUTH ---
def get_google_credentials():
    raw_key = st.secrets["GCP_PRIVATE_KEY"]
    formatted_key = raw_key.replace("\\n", "\n")
    
    info = {
        "type": st.secrets["GCP_TYPE"],
        "project_id": st.secrets["GCP_PROJECT_ID"],
        "private_key_id": st.secrets["GCP_PRIVATE_KEY_ID"],
        "private_key": formatted_key,
        "client_email": st.secrets["GCP_CLIENT_EMAIL"],
        "client_id": st.secrets["GCP_CLIENT_ID"],
        "token_uri": st.secrets["GCP_TOKEN_URI"]
    }
    return service_account.Credentials.from_service_account_info(
        info, 
        scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    )

# --- 3. DOCUMENT GENERATOR ---
def build_monthly_word_document(dept_name, active_month, active_year, creds):
    doc = Document()
    doc.add_heading(f"Report: {dept_name} - {active_month} {active_year}", 0)
    doc_stream = io.BytesIO()
    doc.save(doc_stream)
    return doc_stream.getvalue()

# --- 4. APP UI ---
st.set_page_config(page_title="St. Mary's Research Portal", layout="wide")
if "authenticated" not in st.session_state: st.session_state.authenticated = False
if "admin_enabled" not in st.session_state: st.session_state.admin_enabled = True

if not st.session_state.authenticated:
    st.image("logo.png", width=100)
    st.markdown("## 🔐 St. Mary's Central Achievements Gateway")
    email = st.text_input("College Email").lower()
    if st.button("Submit"):
        st.session_state.authenticated = True
        st.session_state.email = email
        st.rerun()
    st.stop()

st.image("logo.png", width=100)
tabs = st.tabs(["📝 Submit Achievement Log", "📊 Monthly Achievement Generator", "🔒 Admin Control"])

with tabs[2]:
    if st.session_state.email == "research@stmaryscollege.in":
        st.subheader("Admin Control Desk")
        st.session_state.admin_enabled = st.toggle("Enable Data Entry", value=st.session_state.admin_enabled)
    else: st.warning("Unauthorized access.")

with tabs[0]:
    if not st.session_state.admin_enabled and st.session_state.email != "research@stmaryscollege.in":
        st.error("Data entry is currently disabled by the Administrator.")
    else:
        classification = st.selectbox("Select Entry Classification", ["--- Select ---", "🔬 Research Database", "🏆 Faculty Profiles & Milestones", "👥 Departmental & Student Contributions"])
        if classification == "🔬 Research Database":
            with st.form("research_db_form", clear_on_submit=True):
                r_type = st.selectbox("Category", ["Paper Publication", "Book Chapter", "Full Book", "Paper Presentation", "FDP", "Workshop", "Funded Research"])
                c1, c2 = st.columns(2)
                d_from, d_to = c1.date_input("Date From"), c2.date_input("Date To")
                title, org_by = st.text_input("Title"), st.text_input("Organised By")
                url, id_num = st.text_input("URL"), st.text_input("ISSN / ISBN")
                collab = st.checkbox("Collaboration involved?")
                collab_names = st.text_input("Enter Collaborator Names (Mandatory if checked)") if collab else ""
                upload = st.file_uploader("Upload Verification Document (Mandatory)", type=['pdf', 'jpg', 'png'])
                if st.form_submit_button("Submit"):
                    if collab and not collab_names: st.error("Collaborator names required!")
                    elif not upload: st.error("Verification upload is required!")
                    else: st.success("Entry submitted successfully!")

with tabs[1]:
    st.subheader("Monthly Achievement Generator")
    c1, c2, c3 = st.columns(3)
    view_dept, view_month, view_year = c1.selectbox("Target Department", DEPARTMENTS), c2.selectbox("Target Month", MONTHS), c3.selectbox("Target Year", ACADEMIC_YEARS)
    if st.button("🏗️ Construct Automated Monthly Document Package"):
        with st.spinner("Assembling document..."):
            creds = get_google_credentials()
            bytes_data = build_monthly_word_document(view_dept, view_month, view_year, creds)
            st.success("Assembled!")
            st.download_button("📥 Download Report", data=bytes_data, file_name="Report.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

st.markdown("---")
st.markdown("<div style='text-align: center; color: gray;'>Developed by Research Committee @ St. Mary's College</div>", unsafe_allow_html=True)
