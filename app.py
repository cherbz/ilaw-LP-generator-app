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

# --------------------------------------------------
# SESSION STATE INITIALIZATION
# --------------------------------------------------
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

# --------------------------------------------------
# AUTHENTICATION & LICENSE CHECK
# --------------------------------------------------
if not st.session_state["authenticated"]:
    st.title("DepEd ILAW Generator with COT Indicators")
    st.write("Enter your email address, personal Gemini API key, and License Key to start generating DepEd Order No. 003 lesson plans.")

    with st.form("auth_form"):
        email = st.text_input("Enter Email Address:", value="cherbbinondzo@gmail.com")
        api_key = st.text_input("Enter Your Gemini API Key:", type="password")
        license_key = st.text_input("Enter License Key:", type="password")
        submit_btn = st.form_submit_button("Continue to Dashboard")

    if submit_btn:
        if not api_key:
            st.error("Please enter a valid Gemini API Key.")
        elif not license_key:
            st.error("Please enter a valid License Key.")
        else:
            try:
                genai.configure(api_key=api_key)
                st.session_state["api_key"] = api_key
                st.session_state["license_key"] = license_key
                st.session_state["user_email"] = email
                st.session_state["authenticated"] = True
                st.success("Successfully authenticated!")
                st.rerun()
            except Exception as e:
                st.error(f"Authentication failed: {e}")

else:
    # --------------------------------------------------
    # MAIN DASHBOARD & COT INDICATORS WORKSPACE
    # --------------------------------------------------
    st.sidebar.title("Navigation")
    st.sidebar.write(f"Logged in as: **{st.session_state.get('user_email')}**")
    st.sidebar.write(f"License Status: **Active**")
    
    if st.sidebar.button("Log Out"):
        st.session_state["authenticated"] = False
        st.rerun()

    st.title("📘 DepEd ILAW Lesson Plan Generator")
    st.caption("Aligned with DepEd Order No. 003 & Classroom Observation Tool (COT) Indicators")

    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("Lesson Details")
        grade_level = st.selectbox("Grade Level:", ["Grade 7", "Grade 8", "Grade 9", "Grade 10"])
        subject = st.text_input("Learning Area / Subject:", value="Mathematics")
        quarter = st.selectbox("Quarter:", ["Quarter 1", "Quarter 2", "Quarter 3", "Quarter 4"])
        topic = st.text_input("Lesson Topic / Most Essential Learning Competency (MELC):", value="Quadrilaterals and Parallelograms")

    with col2:
        st.subheader("COT Indicators Alignment")
        cot_1 = st.checkbox("COT 1: Apply knowledge of content within and across curriculum teaching areas.", value=True)
        cot_2 = st.checkbox("COT 2: Use a range of teaching strategies that enhance learner achievement in literacy and numeracy skills.", value=True)
        cot_3 = st.checkbox("COT 3: Apply a range of teaching strategies to develop critical and creative thinking, as well as other higher-order thinking skills.", value=True)
        cot_4 = st.checkbox("COT 4: Display proficient use of Mother Tongue, Filipino and English to facilitate teaching and learning.", value=True)
        cot_5 = st.checkbox("COT 5: Establish safe and secure learning environments to enhance learning through the consistent implementation of policies.", value=True)

    st.divider()

    if st.button("🚀 Generate DepEd Order No. 003 Lesson Plan", type="primary"):
        with st.spinner("Generating lesson plan with embedded COT indicators via Gemini AI..."):
            try:
                model = genai.GenerativeModel('gemini-1.5-flash')
                
                selected_cots = []
                if cot_1: selected_cots.append("COT Indicator 1 (Content Integration)")
                if cot_2: selected_cots.append("COT Indicator 2 (Literacy & Numeracy)")
                if cot_3: selected_cots.append("COT Indicator 3 (Critical & Creative Thinking / HOTS)")
                if cot_4: selected_cots.append("COT Indicator 4 (Language Proficiency)")
                if cot_5: selected_cots.append("COT Indicator 5 (Safe & Secure Environment)")

                prompt = f"""
                Generate a complete DepEd Order No. 003 Daily Lesson Log (DLL) / Daily Lesson Plan (DLP) for:
                - Grade Level: {grade_level}
                - Subject: {subject}
                - Quarter: {quarter}
                - Topic: {topic}
                
                Ensure the following Classroom Observation Tool (COT) indicators are explicitly highlighted and integrated into the procedures:
                {', '.join(selected_cots)}

                Structure the lesson plan with the standard ILAW / DepEd components:
                I. Objectives
                II. Content
                III. Learning Resources
                IV. Procedures (Explicitly tag COT indicators in bold where applied)
                V. Remarks & Reflection
                """

                response = model.generate_content(prompt)
                
                st.subheader("Generated Lesson Plan")
                st.markdown(response.text)
                st.success("Lesson plan generated successfully!")

            except Exception as e:
                st.error(f"Error generating lesson plan: {e}")