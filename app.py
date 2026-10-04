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
            if "firebase" in st.secrets:
                cred_dict = dict(st.secrets["firebase"])
                if "private_key" in cred_dict:
                    cred_dict["private_key"] = cred_dict["private_key"].replace("\\n", "\n")
                cred = credentials.Certificate(cred_dict)
                return firebase_admin.initialize_app(cred)
            elif "FIREBASE_CREDENTIALS" in os.environ:
                cred_json = json.loads(os.environ["FIREBASE_CREDENTIALS"])
                if isinstance(cred_json, dict) and "private_key" in cred_json:
                    cred_json["private_key"] = cred_json["private_key"].replace("\\n", "\n")
                cred = credentials.Certificate(cred_json)
                return firebase_admin.initialize_app(cred)
            elif os.path.exists("serviceAccountKey.json"):
                cred = credentials.Certificate("serviceAccountKey.json")
                return firebase_admin.initialize_app(cred)
            else:
                return None
        except Exception as e:
            return None
    return firebase_admin.get_app()

db = init_firebase()

# Initialize session state for login
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

# --------------------------------------------------
# LOGIN / DASHBOARD ROUTING
# --------------------------------------------------
if not st.session_state["authenticated"]:
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
            genai.configure(api_key=api_key)
            st.session_state["api_key"] = api_key
            st.session_state["user_email"] = email
            st.session_state["authenticated"] = True
            st.rerun()

else:
    # --------------------------------------------------
    # MAIN DASHBOARD WORKSPACE
    # --------------------------------------------------
    st.sidebar.title("Navigation")
    st.sidebar.write(f"Logged in as: **{st.session_state.get('user_email')}**")
    if st.sidebar.button("Log Out"):
        st.session_state["authenticated"] = False
        st.rerun()

    st.title("📘 DepEd ILAW Lesson Plan Generator")
    st.success("Authentication successful! Welcome to your workspace.")

    st.subheader("Generate Lesson Plan")
    grade_level = st.selectbox("Select Grade Level:", ["Grade 7", "Grade 8", "Grade 9", "Grade 10"])
    subject = st.text_input("Subject / Learning Area:", value="Mathematics")
    topic = st.text_input("Lesson Topic / Competency:", value="Quadrilaterals and Parallelograms")

    if st.button("Generate Lesson Plan (DepEd Order No. 003)"):
        st.info("Generating lesson plan with Classroom Observation Tool (COT) indicators...")