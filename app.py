import os
import io
import json
import time
import streamlit as st
from datetime import datetime, timedelta, timezone
import firebase_admin
from firebase_admin import credentials, firestore
import google.generativeai as genai
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# --------------------------------------------------
# PAGE CONFIGURATION
# --------------------------------------------------
st.set_page_config(
    page_title="DepEd ILAW Generator with COT Indicators",
    page_icon="📘",
    layout="wide"
)

# --------------------------------------------------
# FIREBASE INITIALIZATION
# --------------------------------------------------
@st.cache_resource
def init_firebase():
    if not firebase_admin._apps:
        try:
            # 1. Check Streamlit Cloud Secrets (Primary)
            if "firebase" in st.secrets:
                cred_dict = dict(st.secrets["firebase"])
                if "private_key" in cred_dict:
                    # Clean up escaping if present
                    cred_dict["private_key"] = cred_dict["private_key"].replace("\\n", "\n")
                cred = credentials.Certificate(cred_dict)
                return firebase_admin.initialize_app(cred)
            
            # 2. Check Environment Variables (Secondary)
            elif "FIREBASE_CREDENTIALS" in os.environ:
                cred_json = json.loads(os.environ["FIREBASE_CREDENTIALS"])
                if isinstance(cred_json, dict) and "private_key" in cred_json:
                    cred_json["private_key"] = cred_json["private_key"].replace("\\n", "\n")
                cred = credentials.Certificate(cred_json)
                return firebase_admin.initialize_app(cred)
            
            # 3. Local fallback (Only if local file exists)
            elif os.path.exists("serviceAccountKey.json"):
                cred = credentials.Certificate("serviceAccountKey.json")
                return firebase_admin.initialize_app(cred)
            
            else:
                st.warning("Firebase credentials not found in secrets or environment. Skipping Firebase init.")
                return None
        except Exception as e:
            st.error(f"Firebase initialization error: {e}")
            return None
    return firebase_admin.get_app()

# Initialize Firebase
db = init_firebase()

# --------------------------------------------------
# APP MAIN INTERFACE
# --------------------------------------------------
st.title("DepEd ILAW Generator with COT Indicators")
st.write("Enter your email address and personal Gemini API key to start generating DepEd Order No. 003 lesson plans.")

with st.form("auth_form"):
    email = st.text_input("Enter Email Address:", value="cherbbinondzo@gmail.com")
    api_key = st.text_input("Enter Your Gemini API Key:", type="password")
    submit_btn = st.form_submit_button("Continue to Dashboard")

if submit_btn:
    if not api_key:
        st.error("Please enter a valid Gemini API Key.")
    else:
        st.session_state["api_key"] = api_key
        st.session_state["user_email"] = email
        st.success("Successfully authenticated! Loading dashboard...")
        st.rerun()