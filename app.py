import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
import io
from docx import Document

# --- 1. CONFIG & FACULTY ---
# (Keep all existing FACULTY_DIRECTORY, DEPARTMENTS, etc. here)

# --- 2. AUTH & HELPERS ---
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
st.set_page_config(page_title="St. Mary's Integrated Portal", layout="wide")

# ... (Auth logic here) ...

st.image("logo.png", width=100)
tab_submit, tab_document, tab_admin = st.tabs(["📝 Submit Achievement Log", "📊 Monthly Achievement Generator", "🔒 Admin Control"])

with tab_submit:
    classification = st.selectbox("Select Classification", ["🔬 Research Database", "🏆 Faculty Profiles & Milestones", "👥 Departmental & Student Contributions"])
    
    with st.form("main_form", clear_on_submit=True):
        upload = st.file_uploader("Upload Verification Document (Mandatory)*")
        
        if classification == "🔬 Research Database":
            r_type = st.selectbox("Type", ["Paper Publication", "Book Chapter", "Full Book", "Paper Presentation", "FDP", "Workshop"])
            
            # FIELDS ONLY FOR PUBLICATIONS
            if r_type in ["Paper Publication", "Book Chapter", "Full Book"]:
                issn = st.text_input("ISSN/ISBN Number*")
                pub_name = st.text_input("Publisher/Journal Name*")
            
            # FIELDS FOR ALL RESEARCH
            title = st.text_input("Title*")
            
            # COLLABORATION LOGIC (ONLY SHOWN IF CHECKED)
            collab = st.checkbox("Collaboration involved?")
            collab_names = st.text_input("Enter Collaborator Names*") if collab else ""
            
            if st.form_submit_button("Submit Research"):
                if not upload: st.error("Verification upload is mandatory!")
                elif collab and not collab_names: st.error("Collaborator names are mandatory!")
                else: st.success(f"Submitted {r_type} successfully!")

        elif classification == "🏆 Faculty Profiles & Milestones":
            st.text_area("Achievement Narrative*")
            if st.form_submit_button("Submit Profile"): st.success("Profile submitted!")

        elif classification == "👥 Departmental & Student Contributions":
            st.text_area("Description*")
            if st.form_submit_button("Submit Contribution"): st.success("Contribution submitted!")

with tab_document:
    st.subheader("Monthly Achievement Generator")
    if st.button("🏗️ Construct Automated Monthly Document Package"):
        st.info("Generating report...")

st.markdown("---")
st.markdown("<div style='text-align: center; color: gray;'>Developed by Research Committee @ St. Mary's College</div>", unsafe_allow_html=True)
