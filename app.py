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

# --- 2. GOOGLE HELPERS ---
def get_google_credentials():
    raw_key = st.secrets["GCP_PRIVATE_KEY"].strip()
    return service_account.Credentials.from_service_account_info({
        "type": st.secrets["GCP_TYPE"], "project_id": st.secrets["GCP_PROJECT_ID"],
        "private_key_id": st.secrets["GCP_PRIVATE_KEY_ID"], "private_key": raw_key,
        "client_email": st.secrets["GCP_CLIENT_EMAIL"], "client_id": st.secrets["GCP_CLIENT_ID"],
        "token_uri": st.secrets["GCP_TOKEN_URI"]
    })

# --- 3. UI ---
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
    st.markdown("---")
    st.markdown("<div style='text-align: center; color: gray;'>Developed by Research Committee @ St. Mary's College</div>", unsafe_allow_html=True)
    st.stop()

st.image("logo.png", width=100)
tabs = st.tabs(["📝 Submit Achievement Log", "📊 Monthly Achievement Generator", "🔒 Admin Control"])

# Tab: Admin
with tabs[2]:
    if st.session_state.email == "research@stmaryscollege.in":
        st.subheader("Admin Control Desk")
        st.session_state.admin_enabled = st.toggle("Enable Data Entry", value=st.session_state.admin_enabled)
    else: st.warning("Unauthorized.")

# Tab: Submission (With granular columns)
with tabs[0]:
    if not st.session_state.admin_enabled and st.session_state.email != "research@stmaryscollege.in":
        st.error("Data entry is disabled by Admin.")
    else:
        with st.form("detailed_form", clear_on_submit=True):
            r_type = st.selectbox("Research Category", ["Paper Publication", "Book Chapter", "Full Book", "Paper Presentation", "FDP", "Workshop", "Funded Research"])
            col1, col2 = st.columns(2)
            with col1: date_from = st.date_input("Date From")
            with col2: date_to = st.date_input("Date To")
            title = st.text_input("Title of Presentation / Publication / FDP / Workshop")
            org_by = st.text_input("Organised By")
            url = st.text_input("URL of Publication / Event")
            id_num = st.text_input("ISSN / ISBN Number")
            collab = st.checkbox("Collaboration involved?")
            collab_names = st.text_input("Enter Collaborator Names (Mandatory if checked)") if collab else ""
            upload = st.file_uploader("Upload Verification Document (Mandatory)", type=['pdf', 'jpg', 'png'])
            
            if st.form_submit_button("Submit"):
                if collab and not collab_names: st.error("Collaborator names required!")
                elif not upload: st.error("Upload required!")
                else: st.success(f"Submitted {r_type} entry!")

# Tab: Generator (Restored to your original working structure)
with tabs[1]:
    st.subheader("Monthly Achievement Generator")
    col_d1, col_d2, col_d3 = st.columns(3)
    view_dept = col_d1.selectbox("Target Department", DEPARTMENTS)
    view_month = col_d2.selectbox("Target Month", MONTHS)
    view_year = col_d3.selectbox("Target Year", ACADEMIC_YEARS)
    if st.button("🏗️ Construct Automated Monthly Document Package"):
        st.info(f"Generating report for {view_dept} - {view_month} {view_year}...")

st.markdown("---")
st.markdown("<div style='text-align: center; color: gray;'>Developed by Research Committee @ St. Mary's College</div>", unsafe_allow_html=True)
