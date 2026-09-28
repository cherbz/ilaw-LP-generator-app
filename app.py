import os
import json
import streamlit as st
from datetime import datetime, timedelta, timezone
import firebase_admin
from firebase_admin import credentials, firestore
import google.generativeai as genai

# -------------------------------------------------------------------
# FIREBASE INITIALIZATION (STREAMLIT CLOUD, VERCEL, & LOCAL FALLBACK)
# -------------------------------------------------------------------
@st.cache_resource
def init_firebase():
    if not firebase_admin._apps:
        # 1. Try reading from Streamlit Secrets (Streamlit Community Cloud)
        if "firebase" in st.secrets:
            cred_dict = dict(st.secrets["firebase"])
            if "private_key" in cred_dict:
                cred_dict["private_key"] = cred_dict["private_key"].replace("\\n", "\n")
            cred = credentials.Certificate(cred_dict)
        # 2. Try reading from Vercel Environment Variables
        elif "FIREBASE_CREDENTIALS" in os.environ:
            cred_json = json.loads(os.environ["FIREBASE_CREDENTIALS"])
            if isinstance(cred_json, dict) and "private_key" in cred_json:
                cred_json["private_key"] = cred_json["private_key"].replace("\\n", "\n")
            cred = credentials.Certificate(cred_json)
        # 3. Fallback to local serviceAccountKey.json for offline testing
        else:
            cred = credentials.Certificate("serviceAccountKey.json")
            
        firebase_admin.initialize_app(cred)
    return firestore.client()

db = init_firebase()

# -------------------------------------------------------------------
# GEMINI AI CONFIGURATION
# -------------------------------------------------------------------
def configure_gemini():
    api_key = None
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
    elif "GEMINI_API_KEY" in os.environ:
        api_key = os.environ["GEMINI_API_KEY"]
        
    if api_key:
        genai.configure(api_key=api_key)
        return True
    return False

# -------------------------------------------------------------------
# DATABASE HELPER FUNCTIONS
# -------------------------------------------------------------------
def get_or_create_user(email):
    user_ref = db.collection("users").document(email)
    doc = user_ref.get()
    
    if doc.exists:
        return doc.to_dict()
    else:
        new_user = {
            "email": email,
            "trial_used": False,
            "license_expires_at": None,
            "created_at": datetime.now(timezone.utc)
        }
        user_ref.set(new_user)
        return new_user

def mark_trial_as_used(email):
    db.collection("users").document(email).update({
        "trial_used": True
    })

def redeem_license_key(email, key_string):
    key_ref = db.collection("license_keys").document(key_string)
    key_doc = key_ref.get()

    if not key_doc.exists:
        return False, "Invalid License Key. Please check for typos."

    key_data = key_doc.to_dict()

    if key_data.get("is_used"):
        return False, "This license key has already been redeemed by another email account."

    now = datetime.now(timezone.utc)
    user_ref = db.collection("users").document(email)
    user_doc = user_ref.get().to_dict()

    current_expiry = user_doc.get("license_expires_at")
    
    # Extend license by 365 days if user already has an active plan, else start from today
    if current_expiry and current_expiry > now:
        new_expiry = current_expiry + timedelta(days=365)
    else:
        new_expiry = now + timedelta(days=365)

    # 1. Update user expiration
    user_ref.update({
        "license_expires_at": new_expiry
    })

    # 2. Mark license key as claimed permanently
    key_ref.update({
        "is_used": True,
        "used_by": email,
        "redeemed_at": now
    })

    return True, "Success! 1-Year License successfully activated."

# -------------------------------------------------------------------
# STREAMLIT UI LAYOUT & PAGE CONFIG
# -------------------------------------------------------------------
st.set_page_config(page_title="Binonz ILAW Lesson Plan Generator", page_icon="📘", layout="centered")

if "user_email" not in st.session_state:
    st.session_state["user_email"] = None

# -------------------------------------------------------------------
# STEP 1: LOGIN / IDENTIFICATION SCREEN (DIRECT ACCESS)
# -------------------------------------------------------------------
if not st.session_state["user_email"]:
    st.title("📘 Binonz ILAW Lesson Plan Generator")
    st.markdown("Welcome! Enter your email address below to access the generator. New users get **1 FREE Trial** lesson generation.")

    email_input = st.text_input("Enter Email Address:").strip().lower()

    if st.button("Continue"):
        if "@" in email_input and "." in email_input:
            st.session_state["user_email"] = email_input
            st.rerun()
        else:
            st.error("Please enter a valid email address.")
    st.stop()

# Fetch user subscription details
user_email = st.session_state["user_email"]
user_data = get_or_create_user(user_email)

now = datetime.now(timezone.utc)
has_active_license = False

if user_data.get("license_expires_at"):
    expires_at = user_data["license_expires_at"]
    if expires_at > now:
        has_active_license = True

# -------------------------------------------------------------------
# STEP 2: SIDEBAR - LICENSE ACTIVATION & ACCOUNT STATUS
# -------------------------------------------------------------------
with st.sidebar:
    st.subheader("Account Overview")
    st.write(f"Logged in as: **{user_email}**")
    
    if st.button("Switch Account"):
        st.session_state["user_email"] = None
        st.rerun()

    st.divider()

    if has_active_license:
        days_left = (user_data["license_expires_at"] - now).days
        st.success(f"🟢 **Active Subscription**\n\n{days_left} days remaining.")
    else:
        st.warning("🔴 **No Active Subscription**")

    st.subheader("Redeem License Key")
    key_input = st.text_input("Enter 1-Year Key:").strip()
    
    if st.button("Activate Key"):
        if key_input:
            success, msg = redeem_license_key(user_email, key_input)
            if success:
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)
        else:
            st.warning("Please enter a key.")

# -------------------------------------------------------------------
# STEP 3: ACCESS CONTROL LOGIC
# -------------------------------------------------------------------
can_generate = False

if has_active_license:
    can_generate = True
elif not user_data.get("trial_used", False):
    st.info("🎁 **Free Trial Available:** You have 1 free lesson plan generation remaining.")
    can_generate = True
else:
    st.error("🔒 **Trial Expired**")
    st.write("You have already used your 1 free trial generation. Please enter a valid 1-Year License Key in the sidebar to unlock unlimited access.")
    st.stop()

# -------------------------------------------------------------------
# STEP 4: MAIN LESSON PLAN GENERATOR INTERFACE (GEMINI AI INTEGRATED)
# -------------------------------------------------------------------
st.title("DepEd ILAW Lesson Plan Generator")

with st.form("generator_form"):
    subject = st.text_input("Subject Area (e.g., Mathematics 9, TLE 10 CSS)")
    grade_level = st.selectbox("Grade Level", ["Grade 7", "Grade 8", "Grade 9", "Grade 10", "Grade 11", "Grade 12"])
    topic = st.text_input("Lesson Topic / Competency")
    quarter = st.selectbox("Quarter", ["Quarter 1", "Quarter 2", "Quarter 3", "Quarter 4"])
    
    submit_button = st.form_submit_button("Generate Lesson Plan")

if submit_button:
    if not subject or not topic:
        st.warning("Please fill in all required fields.")
    else:
        if not configure_gemini():
            st.error("Missing GEMINI_API_KEY. Please set GEMINI_API_KEY in your Streamlit Cloud Secrets or Environment Variables.")
        else:
            with st.spinner("Generating detailed DepEd ILAW Lesson Plan using Gemini AI..."):
                try:
                    prompt = f"""
                    You are an expert DepEd Philippines curriculum developer and master teacher.
                    Generate a complete, highly detailed DepEd ILAW Lesson Plan for the following details:
                    
                    - Subject Area: {subject}
                    - Grade Level: {grade_level}
                    - Topic / Competency: {topic}
                    - Quarter: {quarter}

                    Strictly organize the lesson plan into the following DepEd ILAW sections using clear Markdown headers:

                    ## DepEd ILAW Lesson Plan: {topic}
                    **Subject Area:** {subject} | **Grade Level:** {grade_level} | **Quarter:** {quarter}

                    ---

                    ### I. Objectives
                    - **Knowledge:** 
                    - **Skills:** 
                    - **Attitudes/Values:** 

                    ### II. Content & Learning Resources
                    - **Topic:** {topic}
                    - **Reference Materials:** DepEd Learning Modules, Curriculum Guide
                    - **Tools/Equipment Needed:** 

                    ### III. Procedures (ILAW Framework)
                    - **I - Introduction (Hook & Motivation):** Activity to activate prior knowledge and state learning objectives.
                    - **L - Learning Activity (Direct Instruction & Exploration):** Step-by-step guided activity or demonstration.
                    - **A - Application (Hands-on Practice & Transfer):** Individual/Group activity to apply concepts.
                    - **W - Wrap-Up & Assessment (Evaluation & Synthesis):** Assessment questions/quiz and teacher synthesis.

                    ### IV. Remarks & Reflection
                    - Space for teacher comments, mastery rate, and remediation needs.
                    """

                    model = genai.GenerativeModel('gemini-1.5-flash')
                    response = model.generate_content(prompt)

                    st.markdown(response.text)

                    # Consume trial only after successful AI generation
                    if not has_active_license and not user_data.get("trial_used", False):
                        mark_trial_as_used(user_email)
                        st.warning("⚠️ You have used your 1-time free trial generation. Activate a 1-Year License Key in the sidebar to generate more.")

                except Exception as e:
                    st.error(f"Failed to generate lesson plan: {e}")