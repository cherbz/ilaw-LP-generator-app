import os
import re
import json
import streamlit as st
import google.generativeai as genai
import firebase_admin
from firebase_admin import credentials, firestore

# ==========================================
# 1. PAGE CONFIGURATION
# ==========================================
st.set_page_config(
    page_title="Binonz ILAW Lesson Plan Generator",
    page_icon="📚",
    layout="wide"
)

# ==========================================
# 2. FIREBASE INITIALIZATION
# ==========================================
@st.cache_resource
def init_firebase():
    """Initializes Firebase Admin SDK safely across local, Streamlit, and Vercel environments."""
    if not firebase_admin._apps:
        key_dict = None

        # 1. Try reading from Streamlit Secrets (Dict or String)
        if "text_key" in st.secrets:
            secrets_val = st.secrets["text_key"]
            if isinstance(secrets_val, str):
                try:
                    key_dict = json.loads(secrets_val)
                except Exception:
                    pass
            elif hasattr(secrets_val, "to_dict"):
                key_dict = secrets_val.to_dict()
            elif isinstance(secrets_val, dict):
                key_dict = dict(secrets_val)

        # 2. Try reading from Vercel / OS Environment Variables
        if not key_dict:
            env_var = os.getenv("text_key") or os.getenv("FIREBASE_CREDENTIALS") or os.getenv("FIREBASE_SERVICE_ACCOUNT")
            if env_var:
                try:
                    key_dict = json.loads(env_var)
                except Exception as e:
                    st.error(f"Failed to parse Firebase JSON from environment variable: {e}")
                    st.stop()

        # 3. Fallback to local serviceAccountKey.json file
        if not key_dict and os.path.exists("serviceAccountKey.json"):
            try:
                with open("serviceAccountKey.json", "r") as f:
                    key_dict = json.load(f)
            except Exception as e:
                st.error(f"Failed to load serviceAccountKey.json: {e}")
                st.stop()

        # Initialize Firebase if credentials were found
        if key_dict:
            if "private_key" in key_dict and isinstance(key_dict["private_key"], str):
                key_dict["private_key"] = key_dict["private_key"].replace("\\n", "\n")
                
            cred = credentials.Certificate(key_dict)
            firebase_admin.initialize_app(cred)
        else:
            st.error("Firebase credentials not found! Ensure 'text_key' or 'FIREBASE_CREDENTIALS' environment variable is set in Vercel.")
            st.stop()

    return firestore.client()

db = init_firebase()

# ==========================================
# 3. HELPER FUNCTIONS
# ==========================================
def validate_and_claim_license(license_key: str, email: str) -> tuple[bool, str]:
    """Validates the license key in Firestore and binds it to the user's email."""
    if not license_key or not email:
        return False, "Please provide both a valid License Key and Email address."
    
    key_ref = db.collection("license_keys").document(license_key.strip())
    doc = key_ref.get()

    if not doc.exists:
        return False, "Invalid License Key. Please check and try again."

    data = doc.to_dict()
    if data.get("is_used"):
        if data.get("used_by") == email.strip().lower():
            return True, "License verified."
        else:
            return False, "This key is already registered to another email address."
    
    # Claim key for new user
    key_ref.update({
        "is_used": True,
        "used_by": email.strip().lower(),
        "used_at": firestore.SERVER_TIMESTAMP
    })
    return True, "License key successfully activated!"

def clean_math_syntax(text: str) -> str:
    """Formats raw model output for clean rendering."""
    return text.replace("\\[", "$$").replace("\\]", "$$").replace("\\(", "$").replace("\\)", "$")

def generate_lesson_plan_content(api_key: str, prompt: str) -> str:
    """Generates content using user key with dynamic model auto-selection & fallbacks."""
    genai.configure(api_key=api_key.strip())
    
    preferred_models = []
    try:
        available = [
            m.name.replace("models/", "") 
            for m in genai.list_models() 
            if "generateContent" in m.supported_generation_methods
        ]
        preferred_models = [m for m in available if "flash" in m or "pro" in m]
    except Exception:
        pass

    fallback_models = [
        "gemini-2.5-flash",
        "gemini-2.0-flash",
        "gemini-1.5-flash",
        "gemini-1.5-pro"
    ]

    candidate_models = list(dict.fromkeys(preferred_models + fallback_models))

    last_error = None
    for model_name in candidate_models:
        try:
            model = genai.GenerativeModel(model_name)
            response = model.generate_content(prompt)
            if response and response.text:
                return clean_math_syntax(response.text)
        except Exception as e:
            last_error = e
            continue

    raise Exception(f"Unable to generate content with provided key. Last error: {str(last_error)}")

# ==========================================
# 4. SIDEBAR SETUP
# ==========================================
st.sidebar.title("🔑 Authentication & Setup")

license_key_input = st.sidebar.text_input("License Key", type="password", help="Enter your product license key")
email_input = st.sidebar.text_input("Registered Email Address", help="Enter your registered email")
api_key_input = st.sidebar.text_input("Gemini API Key", type="password", help="Enter your personal Google Gemini API key")

st.sidebar.markdown("---")
st.sidebar.title("👤 Teacher Profile & Position")

teacher_name = st.sidebar.text_input("Teacher Name", value="NORBERTO P. BINONDO JR.")
position = st.sidebar.selectbox("Position / Rank", ["Teacher I", "Teacher II", "Teacher III", "Master Teacher I", "Master Teacher II"])

career_stage = "Beginning to Proficient" if "Teacher I" in position or "Teacher II" in position else "Proficient to Highly Proficient"
st.sidebar.info(f"**Career Stage:** {career_stage}")
st.sidebar.caption("COT Scale: 2 to 6")

# ==========================================
# 5. MAIN APP UI & INPUTS
# ==========================================
st.title("💡 Binonz ILAW Lesson Plan Generator")

# --- SECTION 1: BASIC INFORMATION ---
st.subheader("1. Basic Information & Meta Details")
col1, col2, col3 = st.columns(3)
with col1:
    grade_level = st.selectbox("Grade Level", ["Grade 7", "Grade 8", "Grade 9", "Grade 10", "Grade 11", "Grade 12"])
    learning_area = st.text_input("Learning Area / Subject", value="Mathematics")
with col2:
    quarter = st.selectbox("Quarter", ["Quarter 1", "Quarter 2", "Quarter 3", "Quarter 4"])
    teaching_date = st.date_input("Teaching Date")
with col3:
    time_allotment = st.text_input("Time Allotment", value="60 Minutes")
    school_name = st.text_input("School Name", value="DepEd High School")

st.markdown("---")

# --- SECTION 2: COT INDICATORS ---
st.subheader("2. Classroom Observation Tool (COT) Indicators & PPST Targets")
st.caption("Select the PPST/COT Indicators to target in this lesson plan:")

cot_col1, cot_col2 = st.columns(2)
with cot_col1:
    cot1 = st.checkbox("[1.1.2] Apply knowledge of content within and across curriculum teaching areas", value=True)
    cot2 = st.checkbox("[1.4.2] Use a range of teaching strategies that enhance learner achievement in literacy and numeracy skills", value=True)
    cot3 = st.checkbox("[1.5.2] Apply a range of teaching strategies to develop critical and creative thinking", value=True)
    cot4 = st.checkbox("[4.5.2] Select, develop, organize and use appropriate teaching and learning resources, including ICT, to address learning goals", value=True)

with cot_col2:
    cot5 = st.checkbox("[2.3.2] Manage classroom structure to engage learners in hands-on/collaborative activities", value=True)
    cot6 = st.checkbox("[3.1.2] Use differentiated, developmentally appropriate learning experiences", value=True)
    cot7 = st.checkbox("[5.1.2] Design, select, organize and use diagnostic, formative and summative assessment strategies consistent with curriculum requirements", value=True)
    cot8 = st.checkbox("[5.2.2] Monitor and evaluate learner progress and achievement using learner attainment data", value=False)

st.markdown("---")

# --- SECTION 3: LEARNING OBJECTIVES & CONTEXT ---
st.subheader("3. Learning Objectives & Context")

learning_competency = st.text_area(
    "Learning Competency",
    value="Graphs a linear function and values its real-life applications (domain, range, intercepts, and slope)."
)

learner_context = st.text_area(
    "Learner Context",
    value="The class is a mixed-ability group of learners with varied mathematical inclinations. Visual and kinesthetic learners benefit from coordinate plotting exercises."
)

# ==========================================
# 6. GENERATION ACTION
# ==========================================
if st.button("🚀 Generate Lesson Plan", type="primary", use_container_width=True):
    if not license_key_input or not email_input or not api_key_input:
        st.error("Please fill in your License Key, Registered Email, and Gemini API Key in the sidebar before proceeding.")
    else:
        with st.spinner("Verifying License..."):
            is_valid, msg = validate_and_claim_license(license_key_input, email_input)

        if not is_valid:
            st.error(f"License Error: {msg}")
        else:
            st.success(msg)
            with st.spinner("Generating Detailed Lesson Plan using Gemini API..."):
                
                # Gather active COT targets
                selected_cots = []
                if cot1: selected_cots.append("1.1.2 Knowledge across curriculum")
                if cot2: selected_cots.append("1.4.2 Literacy and Numeracy strategies")
                if cot3: selected_cots.append("1.5.2 Critical and Creative thinking")
                if cot4: selected_cots.append("4.5.2 ICT Integration & Learning Resources")
                if cot5: selected_cots.append("2.3.2 Hands-on & Collaborative learning")
                if cot6: selected_cots.append("3.1.2 Differentiated instruction")
                if cot7: selected_cots.append("5.1.2 Assessment Strategies")
                if cot8: selected_cots.append("5.2.2 Monitoring learner progress")

                prompt = f"""
                You are an expert DepEd Curriculum Specialist. Generate a detailed ILAW Lesson Plan based on the following details:

                **BASIC INFORMATION:**
                - Teacher Name: {teacher_name}
                - Position/Rank: {position}
                - Career Stage: {career_stage}
                - School: {school_name}
                - Grade & Subject: {grade_level} - {learning_area}
                - Quarter & Duration: {quarter} ({time_allotment})

                **TARGET COT/PPST INDICATORS:**
                {", ".join(selected_cots)}

                **CURRICULUM CONTEXT:**
                - Learning Competency: {learning_competency}
                - Learner Context: {learner_context}

                Please output a complete, professionally structured DepEd ILAW Lesson Plan highlighting explicit integration of the targeted COT indicators across learning objectives, ILAW procedure steps, ICT integration, and assessment strategies.
                """

                try:
                    generated_plan = generate_lesson_plan_content(api_key_input, prompt)
                    st.markdown("---")
                    st.subheader("📋 Generated Lesson Plan Output")
                    st.markdown(generated_plan)
                except Exception as err:
                    st.error(f"Generation Error: {str(err)}")