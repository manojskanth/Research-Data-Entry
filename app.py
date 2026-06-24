import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import io

# --- GOOGLE AUTH ---
def get_google_credentials():
    # Helper to clean the multi-line key for the Google library
    raw_key = st.secrets["GCP_PRIVATE_KEY"].strip()
    info = {
        "type": st.secrets["GCP_TYPE"],
        "project_id": st.secrets["GCP_PROJECT_ID"],
        "private_key_id": st.secrets["GCP_PRIVATE_KEY_ID"],
        "private_key": raw_key,
        "client_email": st.secrets["GCP_CLIENT_EMAIL"],
        "client_id": st.secrets["GCP_CLIENT_ID"],
        "token_uri": st.secrets["GCP_TOKEN_URI"]
    }
    return service_account.Credentials.from_service_account_info(info)

# --- APP SETUP ---
st.set_page_config(page_title="St. Mary's Research Portal", layout="wide")
if "authenticated" not in st.session_state: st.session_state.authenticated = False
if "admin_enabled" not in st.session_state: st.session_state.admin_enabled = True

# --- LOGIN ---
if not st.session_state.authenticated:
    st.image("logo.png", width=100)
    st.markdown("## 🔐 St. Mary's Central Achievements Gateway")
    email = st.text_input("College Email").lower()
    pw = st.text_input("Password", type="password")
    if st.button("Submit"):
        if email: 
            st.session_state.authenticated = True
            st.session_state.email = email
            st.rerun()
    st.stop()

# --- MAIN APP ---
st.image("logo.png", width=100)
tabs = st.tabs(["📝 Submit Achievement Log", "📊 Monthly Achievement Generator", "🔒 Admin Control"])

# Admin Tab
with tabs[2]:
    if st.session_state.email == "research@stmaryscollege.in":
        st.subheader("Admin Control Desk")
        st.session_state.admin_enabled = st.toggle("Enable Data Entry for Users", value=st.session_state.admin_enabled)
    else: st.warning("Unauthorized access.")

# Submission Tab
with tabs[0]:
    if not st.session_state.admin_enabled and st.session_state.email != "research@stmaryscollege.in":
        st.error("Data entry is currently disabled by the Administrator.")
    else:
        with st.form("research_form", clear_on_submit=True):
            r_type = st.selectbox("Research Type", ["Paper Publication", "Book Chapter", "Paper Presentation"])
            collab = st.checkbox("Collaboration involved?")
            collab_names = st.text_input("Enter Collaborator Names (Mandatory if checked)") if collab else ""
            upload = st.file_uploader("Upload Verification Document (Mandatory)", type=['pdf', 'jpg', 'png'])
            
            # Dynamic fields
            if r_type == "Paper Publication": journal = st.text_input("Journal Name")
            
            if st.form_submit_button("Submit"):
                if collab and not collab_names: st.error("Collaborator names are mandatory when collaboration is checked!")
                elif not upload: st.error("Verification upload is mandatory!")
                else: st.success("Entry submitted successfully!")

# Footer
st.markdown("---")
st.markdown("<div style='text-align: center; color: gray;'>Developed by Research Committee @ St. Mary's College</div>", unsafe_allow_html=True)
