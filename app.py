import io
import re
import json
import streamlit as st
import google.generativeai as genai
import firebase_admin
from firebase_admin import credentials, firestore
from docx import Document

# ---------------------------------------------------------
# 1. PAGE CONFIGURATION & STYLING
# ---------------------------------------------------------
st.set_page_config(
    page_title="Binonz Semi-Detailed ILAW Lesson Plan Generator",
    page_icon="💡",
    layout="wide"
)

# ---------------------------------------------------------
# 2. FIREBASE & GEMINI INITIALIZATION
# ---------------------------------------------------------
@st.cache_resource
def init_firebase():
    """Initialize Firebase Admin SDK using Streamlit Secrets."""
    if not firebase_admin._apps:
        try:
            # Check if secret string exists
            if "FIREBASE_CREDENTIALS" in st.secrets:
                cred_data = json.loads(st.secrets["FIREBASE_CREDENTIALS"])
                cred = credentials.Certificate(cred_data)
                firebase_admin.initialize_app(cred)
            else:
                st.error("Firebase credentials not found in Streamlit Secrets.")
                return None
        except Exception as e:
            st.error(f"Failed to initialize Firebase: {e}")
            return None
    return firestore.client()

db = init_firebase()

# ---------------------------------------------------------
# 3. HELPER & VALIDATION FUNCTIONS
# ---------------------------------------------------------
def validate_and_claim_license(license_key, email_input):
    """
    Validates the license key against Firestore.
    Includes explicit checks to prevent empty path crashes.
    """
    key = license_key.strip() if license_key else ""
    email = email_input.strip() if email_input else ""

    # Check empty input FIRST to avoid invalid Firestore path elements (must be even)
    if not key:
        return False, "Please enter a valid License Key."
    if not email:
        return False, "Please enter your Registered Email Address."

    if not db:
        return False, "Database connection not available."

    try:
        # Fetch document from Firestore
        key_ref = db.collection("license_keys").document(key)
        doc = key_ref.get()

        if not doc.exists:
            return False, "Invalid or unrecognized License Key."

        data = doc.to_dict()
        registered_email = data.get("email", "").strip().lower()

        if registered_email and registered_email != email.lower():
            return False, "Email does not match the registered key user."

        return True, "License validated successfully!"

    except Exception as e:
        return False, f"Database Error: {str(e)}"

def generate_lesson_plan(api_key, meta_data, cot_targets, teacher_info):
    """Generates lesson plan via Gemini API."""
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel("gemini-1.5-flash")

        prompt = f"""
        Act as an expert DepEd Educator. Create a Semi-Detailed ILAW Lesson Plan based on:
        
        Teacher Profile:
        - Name: {teacher_info['name']}
        - Position: {teacher_info['position']}
        - Career Stage: {teacher_info['stage']}
        
        Meta Details:
        - Grade Level: {meta_data['grade']}
        - Section: {meta_data['section']}
        - Subject: {meta_data['subject']}
        - Quarter: {meta_data['quarter']}
        - Lesson: {meta_data['lesson_name']}
        - Time Allotment: {meta_data['time']}
        - Date: {meta_data['date']}
        
        Selected COT/PPST Indicators:
        {', '.join(cot_targets)}
        
        Structure the lesson plan clearly using DepEd ILAW standards.
        """
        
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        return f"Error generating content from Gemini: {str(e)}"

# ---------------------------------------------------------
# 4. STREAMLIT FRONTEND UI
# ---------------------------------------------------------
st.title("💡 Binonz Semi-Detailed ILAW Lesson Plan Generator")

# Sidebar - Setup & Profile
with st.sidebar:
    st.header("🔑 Authentication & Setup")
    license_key = st.text_input("License Key", type="password", help="Enter your product license key")
    email_input = st.text_input("Registered Email Address", help="Enter your registered email")
    gemini_api_key = st.text_input("Gemini API Key", type="password", help="Enter your Google Gemini API Key")

    st.markdown("---")
    st.header("👤 Teacher Profile & Position")
    teacher_name = st.text_input("Teacher Name", value="NORBERTO P. BINONDO JR.")
    position = st.selectbox("Position / Rank", ["Teacher I", "Teacher II", "Teacher III", "Master Teacher I", "Master Teacher II"], index=2)
    career_stage = st.selectbox("Career Stage", ["Beginning to Proficient", "Proficient", "Highly Proficient", "Distinguished"])

# Main Form Area
st.subheader("1. Basic Information & Meta Details")

col1, col2, col3 = st.columns(3)
with col1:
    grade_level = st.selectbox("Grade Level", ["Grade 7", "Grade 8", "Grade 9", "Grade 10", "Grade 11", "Grade 12"])
    subject = st.text_input("Learning Area / Subject", value="Mathematics")
    lesson_name = st.text_input("Name of Lesson", value="Graphing Linear Functions")

with col2:
    section = st.text_input("Section", value="Kindness")
    quarter = st.selectbox("Quarter", ["Quarter 1", "Quarter 2", "Quarter 3", "Quarter 4"])
    teaching_date = st.date_input("Teaching Date")

with col3:
    time_allotment = st.text_input("Time Allotment", value="60 Minutes")
    school_name = st.text_input("School Name", value="DepEd High School")

st.markdown("---")
st.subheader("2. Classroom Observation Tool (COT) Indicators & PPST Targets")
st.caption("Select the PPST/COT Indicators to target in this lesson plan:")

cot_col1, cot_col2 = st.columns(2)
selected_cots = []

with cot_col1:
    if st.checkbox("[1.1.2] Apply knowledge of content within and across curriculum teaching areas", value=True):
        selected_cots.append("[1.1.2] Apply knowledge of content within and across curriculum teaching areas")
    if st.checkbox("[1.4.2] Use a range of teaching strategies that enhance learner achievement in literacy and numeracy skills", value=True):
        selected_cots.append("[1.4.2] Use a range of teaching strategies that enhance learner achievement in literacy and numeracy skills")
    if st.checkbox("[1.5.2] Apply a range of teaching strategies to develop critical and creative thinking", value=True):
        selected_cots.append("[1.5.2] Apply a range of teaching strategies to develop critical and creative thinking")

with cot_col2:
    if st.checkbox("[2.3.2] Manage classroom structure to engage learners in hands-on/collaborative activities", value=True):
        selected_cots.append("[2.3.2] Manage classroom structure to engage learners in hands-on/collaborative activities")
    if st.checkbox("[3.1.2] Use differentiated, developmentally appropriate learning experiences", value=True):
        selected_cots.append("[3.1.2] Use differentiated, developmentally appropriate learning experiences")
    if st.checkbox("[5.1.2] Design, select, organize and use diagnostic, formative and summative assessment strategies", value=True):
        selected_cots.append("[5.1.2] Design, select, organize and use diagnostic, formative and summative assessment strategies")

st.markdown("---")

# Submit Button
if st.button("🚀 Generate Semi-Detailed Lesson Plan", use_container_width=True, type="primary"):
    if not gemini_api_key:
        st.error("Please provide a valid Gemini API Key in the sidebar.")
    else:
        # Validate License First
        is_valid, msg = validate_and_claim_license(license_key, email_input)
        if not is_valid:
            st.error(msg)
        else:
            st.success(msg)
            with st.spinner("Generating lesson plan... Please wait..."):
                meta_data = {
                    "grade": grade_level,
                    "subject": subject,
                    "lesson_name": lesson_name,
                    "section": section,
                    "quarter": quarter,
                    "date": str(teaching_date),
                    "time": time_allotment,
                    "school": school_name
                }
                teacher_info = {
                    "name": teacher_name,
                    "position": position,
                    "stage": career_stage
                }
                
                result = generate_lesson_plan(gemini_api_key, meta_data, selected_cots, teacher_info)
                
                st.markdown("### Generated Lesson Plan")
                st.write(result)