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
    """Initializes Firebase Admin SDK from Secrets, Env Vars, or local file."""
    if not firebase_admin._apps:
        key_dict = None

        # 1. Check Streamlit Secrets (TOML / Streamlit Cloud)
        if "text_key" in st.secrets:
            key_dict = dict(st.secrets["text_key"])

        # 2. Check Vercel Environment Variables (JSON string)
        elif os.getenv("text_key"):
            key_dict = json.loads(os.getenv("text_key"))
        elif os.getenv("FIREBASE_CREDENTIALS"):
            key_dict = json.loads(os.getenv("FIREBASE_CREDENTIALS"))

        # Initialize from dictionary if found
        if key_dict:
            cred = credentials.Certificate(key_dict)
            firebase_admin.initialize_app(cred)

        # 3. Fallback to local serviceAccountKey.json file
        elif os.path.exists("serviceAccountKey.json"):
            cred = credentials.Certificate("serviceAccountKey.json")
            firebase_admin.initialize_app(cred)
        else:
            st.error("Firebase credentials not found! Please check serviceAccountKey.json or Vercel Environment Variables.")
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
    
    # 1. Try to fetch active models supported by the key
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

    # 2. Hardcoded fallbacks in case list_models isn't permitted by API key scope
    fallback_models = [
        "gemini-2.5-flash",
        "gemini-2.0-flash",
        "gemini-1.5-flash",
        "gemini-1.5-pro"
    ]

    # Combine lists while maintaining order and removing duplicates
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

col1, col2 = st.columns(2)
with col1:
    st.checkbox("[4.5.2] Select, develop, organize and use appropriate teaching and learning resources, including ICT, to address learning goals", value=True)
with col2:
    st.checkbox("[5.1.2] Design, select, organize and use diagnostic, formative and summative assessment strategies consistent with curriculum requirements", value=True)

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
                prompt = f"""
                You are an expert DepEd Curriculum Specialist. Generate a detailed ILAW Lesson Plan based on the following details:

                Teacher Name: {teacher_name}
                Position/Rank: {position}
                Career Stage: {career_stage}
                Learning Competency: {learning_competency}
                Learner Context: {learner_context}

                Please output a fully structured, professional lesson plan with complete learning objectives, procedures, ICT integration, and assessment strategies.
                """

                try:
                    generated_plan = generate_lesson_plan_content(api_key_input, prompt)
                    st.markdown("---")
                    st.subheader("📋 Generated Lesson Plan Output")
                    st.markdown(generated_plan)
                except Exception as err:
                    st.error(f"Generation Error: {str(err)}")