import os
import io
import json
import time
import streamlit as st
from datetime import datetime, timedelta, timezone
import firebase_admin
from firebase_admin import credentials, firestore
from google import genai
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# -------------------------------------------------------------------
# PAGE CONFIGURATION & CUSTOM CSS STYLING
# -------------------------------------------------------------------
st.set_page_config(
    page_title="DepEd ILAW Lesson Plan Generator",
    page_icon="📘",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS to replicate the UI layout
st.markdown("""
<style>
    /* Main Background & Font */
    .stApp {
        background-color: #F0F4F9;
        font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
    }

    /* Top Navigation Header Bar */
    .top-header {
        background: linear-gradient(90deg, #0D47A1 0%, #1565C0 60%, #1976D2 100%);
        padding: 12px 24px;
        border-radius: 0px 0px 12px 12px;
        color: white;
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 20px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    }
    
    .brand-title {
        font-size: 22px;
        font-weight: 800;
        color: #FFFFFF;
        display: flex;
        align-items: center;
        gap: 10px;
    }
    
    .brand-tagline {
        font-size: 12px;
        color: #BBDEFB;
        font-style: italic;
    }

    /* Sidebar Custom Styling */
    [data-testid="stSidebar"] {
        background-color: #0A2540 !important;
        color: #FFFFFF;
    }
    
    [data-testid="stSidebar"] * {
        color: #FFFFFF !important;
    }

    /* Custom Input Container Card */
    .input-card {
        background-color: #FFFFFF;
        border-radius: 16px;
        padding: 24px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.05);
        margin-bottom: 20px;
    }

    /* Input Field Label Styling */
    .field-label {
        font-weight: 700;
        font-size: 14px;
        margin-bottom: 6px;
        display: flex;
        align-items: center;
        gap: 8px;
    }

    /* Feature Badge Cards */
    .badge-card {
        background: #FFFFFF;
        border-radius: 12px;
        padding: 12px 16px;
        display: flex;
        align-items: center;
        gap: 12px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        border: 1px solid #E2E8F0;
        font-size: 13px;
        font-weight: 600;
    }

    /* Footer Feature Cards */
    .footer-bar {
        background: #FFFFFF;
        border-radius: 12px;
        padding: 16px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.05);
        margin-top: 20px;
    }
    .footer-item {
        text-align: center;
        border-right: 1px solid #E2E8F0;
    }
    .footer-item:last-child {
        border-right: none;
    }
    .footer-title {
        font-weight: 700;
        color: #0D47A1;
        font-size: 13px;
    }
    .footer-sub {
        font-size: 11px;
        color: #64748B;
    }

    /* Generate Button Styling */
    div.stButton > button {
        background: linear-gradient(90deg, #2563EB 0%, #3B82F6 100%) !important;
        color: white !important;
        font-weight: 700 !important;
        font-size: 16px !important;
        border-radius: 12px !important;
        padding: 12px 28px !important;
        border: none !important;
        box-shadow: 0 4px 14px rgba(37, 99, 235, 0.35) !important;
        width: 100% !important;
    }

    /* Hide standard Streamlit header/footer padding */
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 2rem !important;
    }
</style>
""", unsafe_allow_html=True)

# -------------------------------------------------------------------
# FIREBASE INITIALIZATION
# -------------------------------------------------------------------
@st.cache_resource
def init_firebase():
    if not firebase_admin._apps:
        if "firebase" in st.secrets:
            cred_dict = dict(st.secrets["firebase"])
            if "private_key" in cred_dict:
                cred_dict["private_key"] = cred_dict["private_key"].replace("\\n", "\n")
            cred = credentials.Certificate(cred_dict)
        elif "FIREBASE_CREDENTIALS" in os.environ:
            cred_json = json.loads(os.environ["FIREBASE_CREDENTIALS"])
            if isinstance(cred_json, dict) and "private_key" in cred_json:
                cred_json["private_key"] = cred_json["private_key"].replace("\\n", "\n")
            cred = credentials.Certificate(cred_json)
        else:
            cred = credentials.Certificate("serviceAccountKey.json")
            
        firebase_admin.initialize_app(cred)
    return firestore.client()

db = init_firebase()

# -------------------------------------------------------------------
# DYNAMIC GEMINI MODEL RESOLUTION WITH FALLBACKS
# -------------------------------------------------------------------
def get_candidate_models(user_api_key):
    if not user_api_key:
        raise Exception("API key is missing.")

    client = genai.Client(api_key=user_api_key)
    candidate_models = []

    try:
        all_models = list(client.models.list())
        clean_models = [m.name.replace("models/", "") for m in all_models]

        active_models = [
            m for m in clean_models 
            if "2.5-flash" not in m and "embedding" not in m and "tts" not in m
        ]

        flash_models = [m for m in active_models if "flash" in m.lower()]
        other_models = [m for m in active_models if m not in flash_models]
        
        candidate_models = flash_models + other_models
    except Exception:
        pass

    defaults = ["gemini-3.8-flash", "gemini-3.5-flash"]
    for d in defaults:
        if d not in candidate_models:
            candidate_models.append(d)

    return client, candidate_models

def generate_with_fallback(client, candidate_models, prompt):
    last_exception = None

    for model_name in candidate_models:
        for attempt in range(2):
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=prompt
                )
                if response and response.text:
                    return response.text, model_name
            except Exception as e:
                last_exception = e
                err_msg = str(e)
                if "503" in err_msg or "UNAVAILABLE" in err_msg or "high demand" in err_msg:
                    time.sleep(1.5)
                    continue
                else:
                    break

    raise Exception(f"All model attempts failed. Last error: {last_exception}")

# -------------------------------------------------------------------
# COT INDICATOR PRESETS
# -------------------------------------------------------------------
COT_INDICATOR_OPTIONS = {
    "Teacher I - III (Proficient)": [
        "1.1.2: Apply knowledge of content within and across curriculum teaching areas.",
        "1.4.2: Use a range of teaching strategies that enhance learner achievement in literacy and numeracy skills.",
        "1.5.2: Apply a range of teaching strategies to develop critical and creative thinking, as well as other higher-order thinking skills.",
        "2.3.2: Manage classroom structure to engage learners in meaningful exploration, discovery and hands-on activities.",
        "2.6.2: Manage learner behavior constructively by applying positive and non-violent discipline.",
        "3.1.2: Use differentiated, developmentally appropriate learning experiences to address learners' gender, needs, strengths, interests and experiences.",
        "4.1.2: Design, select, organize, and use diagnostic, formative, and summative assessment strategies.",
        "4.5.2: Select, develop, organize, and use appropriate teaching and learning resources, including ICT."
    ],
    "Teacher IV - VI (Highly Proficient)": [
        "1.1.3: Model effective applications of content knowledge within and across curriculum teaching areas.",
        "1.4.3: Model a range of teaching strategies that enhance learner achievement in literacy and numeracy skills.",
        "1.5.3: Model a range of teaching strategies to develop critical and creative thinking.",
        "2.3.3: Model effective management of classroom structure.",
        "3.1.3: Work with colleagues to design, adapt and implement differentiated learning experiences."
    ],
    "Master Teacher I - II (Highly Proficient)": [
        "1.1.3: Model effective applications of content knowledge within and across curriculum teaching areas.",
        "1.4.3: Evaluate with colleagues teaching strategies that enhance literacy and numeracy skills.",
        "1.5.3: Model effective teaching strategies to develop HOTS.",
        "3.1.3: Lead colleagues in evaluating differentiated strategies."
    ]
}

def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

def create_deped_annex_a_docx(plan_data, metadata):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    header_p = doc.add_paragraph()
    header_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = header_p.add_run(f"ILAW LESSON PLAN ON {metadata['subject'].upper()}\n")
    run_title.bold = True
    run_title.font.size = Pt(14)
    run_title.font.color.rgb = RGBColor(16, 44, 87)

    run_sub = header_p.add_run("DepEd Order No. 003, s. 2026 (Annex A Template) | COT Indicators Embedded\n")
    run_sub.font.size = Pt(9)
    run_sub.font.italic = True
    run_sub.font.color.rgb = RGBColor(100, 100, 100)

    meta_table = doc.add_table(rows=6, cols=2)
    meta_table.style = 'Table Grid'
    
    meta_rows = [
        ("Name of Lesson", metadata.get("topic", "")),
        ("Learning Area/s", metadata.get("subject", "")),
        ("Designed by Teacher/s", metadata.get("teacher_name", "NORBERTO P. BINONDO JR.")),
        ("Grade Level & Section", f"{metadata.get('grade_level', '')} - {metadata.get('section', '')}"),
        ("No. of Sessions", metadata.get("sessions", "1")),
        ("References", metadata.get("references", "DepEd K-12 Curriculum Guide & Learning Materials")),
    ]

    for i, (label, val) in enumerate(meta_rows):
        row = meta_table.rows[i]
        c1, c2 = row.cells[0], row.cells[1]
        c1.width = Inches(2.2)
        c2.width = Inches(4.8)
        set_cell_background(c1, "F0F4F8")
        
        p1 = c1.paragraphs[0]
        r1 = p1.add_run(label)
        r1.bold = True
        r1.font.size = Pt(10)
        
        p2 = c2.paragraphs[0]
        r2 = p2.add_run(val)
        r2.font.size = Pt(10)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    def add_section_header(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(title)
        run.bold = True
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(16, 44, 87)

    sections = [
        ("1. INTENTIONS", [
            ("Learning Competency", plan_data.get("learning_competency", "")),
            ("Learning Objectives", plan_data.get("learning_objectives", "")),
            ("Learner Context", plan_data.get("learner_context", ""))
        ]),
        ("2. LEARNING EXPERIENCE", [
            ("Pre-Lesson (Getting Ready)", plan_data.get("pre_lesson", "")),
            ("Instructional Flow & Direct Modeling", plan_data.get("instructional_flow", "")),
            ("Collaborative Group Activity", plan_data.get("collaborative_activity", "")),
            ("Synthesis & Resources", plan_data.get("synthesis", "")),
            ("Opportunities for Integration", plan_data.get("integration", ""))
        ]),
        ("3. ASSESSMENT", [
            ("Formative Assessment (Individual Evaluation)", plan_data.get("formative_assessment", ""))
        ]),
        ("4. WAYS FORWARD", [
            ("Extended Learning Opportunities", plan_data.get("extended_learning", "")),
            ("Teacher Reflections", plan_data.get("teacher_reflections", ""))
        ])
    ]

    for sec_title, fields in sections:
        add_section_header(sec_title)
        table = doc.add_table(rows=len(fields), cols=2)
        table.style = 'Table Grid'
        
        for idx, (label, value) in enumerate(fields):
            row = table.rows[idx]
            c1, c2 = row.cells[0], row.cells[1]
            c1.width = Inches(2.2)
            c2.width = Inches(4.8)
            set_cell_background(c1, "F9FAFB")
            
            p1 = c1.paragraphs[0]
            r1 = p1.add_run(label)
            r1.bold = True
            r1.font.size = Pt(10)
            
            p2 = c2.paragraphs[0]
            lines = value.split('\n')
            for line_idx, line in enumerate(lines):
                if line_idx > 0:
                    p2 = c2.add_paragraph()
                if "📌 [COT INDICATOR:" in line:
                    r2 = p2.add_run(line)
                    r2.bold = True
                    r2.font.size = Pt(9.5)
                    r2.font.color.rgb = RGBColor(180, 40, 40)
                else:
                    r2 = p2.add_run(line)
                    r2.font.size = Pt(10)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

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
    db.collection("users").document(email).update({"trial_used": True})

def redeem_license_key(email, key_string):
    key_ref = db.collection("license_keys").document(key_string)
    key_doc = key_ref.get()
    if not key_doc.exists:
        return False, "Invalid License Key."
    key_data = key_doc.to_dict()
    if key_data.get("is_used"):
        return False, "This key has already been redeemed."

    now = datetime.now(timezone.utc)
    user_ref = db.collection("users").document(email)
    user_doc = user_ref.get().to_dict()

    current_expiry = user_doc.get("license_expires_at")
    if current_expiry and current_expiry > now:
        new_expiry = current_expiry + timedelta(days=365)
    else:
        new_expiry = now + timedelta(days=365)

    user_ref.update({"license_expires_at": new_expiry})
    key_ref.update({"is_used": True, "used_by": email, "redeemed_at": now})

    return True, "1-Year License successfully activated!"

# -------------------------------------------------------------------
# USER AUTHENTICATION & LOGIN SESSION
# -------------------------------------------------------------------
if "user_email" not in st.session_state:
    st.session_state["user_email"] = None
if "user_gemini_key" not in st.session_state:
    st.session_state["user_gemini_key"] = None

if not st.session_state["user_email"] or not st.session_state["user_gemini_key"]:
    st.title("📘 DepEd ILAW Lesson Plan Generator")
    st.markdown("Enter your email address and personal Gemini API key to start generating DepEd Order No. 003 lesson plans.")

    with st.form("login_form"):
        email_input = st.text_input("Enter Email Address:").strip().lower()
        api_key_input = st.text_input("Enter Your Gemini API Key:", type="password").strip()
        login_submitted = st.form_submit_button("Continue to Dashboard")

        if login_submitted:
            if not ("@" in email_input and "." in email_input):
                st.error("Please enter a valid email address.")
            elif not api_key_input:
                st.error("Please enter your Gemini API Key.")
            else:
                st.session_state["user_email"] = email_input
                st.session_state["user_gemini_key"] = api_key_input
                st.rerun()
    st.stop()

user_email = st.session_state["user_email"]
user_gemini_key = st.session_state["user_gemini_key"]
user_data = get_or_create_user(user_email)

now = datetime.now(timezone.utc)
has_active_license = False
if user_data.get("license_expires_at") and user_data["license_expires_at"] > now:
    has_active_license = True

# -------------------------------------------------------------------
# SIDEBAR NAVIGATION
# -------------------------------------------------------------------
with st.sidebar:
    st.markdown("### 🏠 Navigation")
    st.markdown("- 📊 **Dashboard**")
    st.markdown("- 📁 **My Lesson Plans**")
    st.markdown("- 📑 **Templates (Annex A)**")
    st.markdown("- 🎯 **COT Indicators**")
    st.markdown("- 📚 **Resources**")
    st.markdown("- ⚙️ **Settings**")
    st.markdown("- ❓ **Help & Support**")
    
    st.divider()

    st.markdown("### 👤 Account Overview")
    st.caption(f"Logged in as:\n**{user_email}**")
    if st.button("🔑 Switch Account / API Key"):
        st.session_state["user_email"] = None
        st.session_state["user_gemini_key"] = None
        st.rerun()

    st.divider()

    if has_active_license:
        days_left = (user_data["license_expires_at"] - now).days
        st.success(f"👑 **Active Subscription** ({days_left} days left)")
    else:
        st.warning("🔴 **No Active License**")

    st.markdown("#### Redeem License Key")
    key_input = st.text_input("Enter Key:", key="license_key_sidebar").strip()
    if st.button("Activate Key"):
        if key_input:
            success, msg = redeem_license_key(user_email, key_input)
            if success:
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)

    st.divider()
    st.markdown("### 📌 COT Indicators")
    teacher_rank = st.selectbox("Select Rank:", list(COT_INDICATOR_OPTIONS.keys()))
    available_indicators = COT_INDICATOR_OPTIONS[teacher_rank]
    selected_cots = []
    for cot in available_indicators:
        if st.checkbox(cot, value=True, key=f"cot_{cot[:5]}"):
            selected_cots.append(cot)

# -------------------------------------------------------------------
# MAIN DASHBOARD UI
# -------------------------------------------------------------------

# Top Blue Navigation Banner
st.markdown("""
<div class="top-header">
    <div>
        <div class="brand-title">📘 DepEd ILAW <span style="font-size: 14px; font-weight: 400; opacity: 0.9;">Lesson Plan Generator</span></div>
        <div class="brand-tagline">Plan • Teach • Inspire — Quality Lessons for a Brighter Future</div>
    </div>
    <div style="font-size: 13px; background: rgba(255,255,255,0.15); padding: 6px 14px; border-radius: 20px;">
        👨‍🏫 <b>Teacher Mode</b> | DepEd Educator
    </div>
</div>
""", unsafe_allow_html=True)

# Hero Branding Banner
st.markdown("""
<div style="background: linear-gradient(135deg, #EBF3FE 0%, #C6E0FF 100%); border-radius: 16px; padding: 24px 32px; margin-bottom: 24px; border: 1px solid #B8D8FF; display: flex; justify-content: space-between; align-items: center;">
    <div>
        <h2 style="color: #0D47A1; font-size: 28px; font-weight: 900; margin: 0;">💡 DepEd ILAW Lesson Plan <span style="color: #F57C00;">Generator</span></h2>
        <p style="color: #1565C0; font-size: 14px; font-weight: 600; margin-top: 6px;">Aligned with DepEd Order No. 003, s. 2026 Annex A Template & COT Indicators</p>
    </div>
    <div style="background: #FFFFFF; padding: 10px 20px; border-radius: 30px; box-shadow: 0 4px 12px rgba(13,71,161,0.1); color: #0D47A1; font-weight: 800; font-size: 14px;">
        ✨ Quality Lesson Plans Made Easy!
    </div>
</div>
""", unsafe_allow_html=True)

# Form Fields Card Layout
with st.form("deped_ilaw_form"):
    st.markdown("<h4 style='color: #0D47A1; margin-bottom: 16px;'>📝 Lesson Configuration</h4>", unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        teacher_name = st.text_input("👤 Teacher Name", value="NORBERTO P. BINONDO JR.")
        subject = st.selectbox("📚 Learning Area / Subject", ["Mathematics", "Science", "English", "Filipino", "Apan", "TLE / CSS", "MAPEH"], index=0)
        grade_level = st.selectbox("🎓 Grade Level", ["Grade 7", "Grade 8", "Grade 9", "Grade 10", "Grade 11", "Grade 12"], index=2)
        section = st.text_input("🏫 Section", value="Kindness")

    with col2:
        topic = st.text_input("🎯 Lesson Topic / Competency", value="Graphing Linear Functions")
        sessions = st.selectbox("⏱️ No. of Sessions", ["1", "2", "3", "4", "5"], index=0)
        quarter = st.selectbox("📅 Quarter", ["Quarter 1", "Quarter 2", "Quarter 3", "Quarter 4"], index=0)
        references = st.text_input("📖 References", value="DepEd Curriculum Guide & Presentation Slides")

    st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
    submit_button = st.form_submit_button("✨ Generate DepEd ILAW Lesson Plan ➔")

# Feature Badges Bar
col_b1, col_b2, col_b3, col_b4 = st.columns(4)
with col_b1:
    st.markdown('<div class="badge-card">📄 <span>Aligned with DepEd Standards</span></div>', unsafe_allow_html=True)
with col_b2:
    st.markdown('<div class="badge-card">🎯 <span>COT Indicators Included</span></div>', unsafe_allow_html=True)
with col_b3:
    st.markdown('<div class="badge-card">⚡ <span>Easy to Use & Fast</span></div>', unsafe_allow_html=True)
with col_b4:
    st.markdown('<div class="badge-card">❤️ <span>Designed for Teachers</span></div>', unsafe_allow_html=True)

# Footer Highlights
st.markdown("""
<div class="footer-bar">
    <div style="display: flex; justify-content: space-around;">
        <div class="footer-item">
            <div class="footer-title">⏱️ Save Time</div>
            <div class="footer-sub">Generate in seconds</div>
        </div>
        <div class="footer-item">
            <div class="footer-title">📄 Professional Format</div>
            <div class="footer-sub">Ready-to-use template</div>
        </div>
        <div class="footer-item">
            <div class="footer-title">🎯 Curriculum Aligned</div>
            <div class="footer-sub">DepEd Order No. 003, s. 2026</div>
        </div>
        <div class="footer-item">
            <div class="footer-title">📥 Download / Export</div>
            <div class="footer-sub">Save or print your plan</div>
        </div>
        <div class="footer-item">
            <div class="footer-title">💙 For Better Learning</div>
            <div class="footer-sub">Support every Filipino learner</div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# -------------------------------------------------------------------
# GENERATION LOGIC
# -------------------------------------------------------------------
if submit_button:
    if not subject or not topic:
        st.warning("Please fill in all required fields.")
    elif not selected_cots:
        st.warning("Please select at least one COT Indicator in the sidebar.")
    else:
        with st.spinner("Generating DepEd Order No. 003 Annex A ILAW Lesson Plan..."):
            try:
                client, candidate_models = get_candidate_models(user_gemini_key)
                cot_list_str = "\n".join([f"- {c}" for c in selected_cots])

                prompt = f"""
                You are a DepEd Master Teacher and Curriculum Expert.
                Generate a DepEd ILAW Lesson Plan following DepEd Order No. 003, s. 2026 (Annex A Template).

                METADATA:
                - Subject: {subject}
                - Grade & Section: {grade_level} - {section}
                - Topic: {topic}
                - Sessions: {sessions}
                
                TARGET COT INDICATORS TO EMBED IN THE CONTENT:
                {cot_list_str}

                Respond ONLY in valid JSON format matching this exact schema:
                {{
                    "learning_competency": "...",
                    "learning_objectives": "1. ...\\n2. ...\\n3. ...\\n📌 [COT INDICATOR: ...]",
                    "learner_context": "...\\n📌 [COT INDICATOR: ...]",
                    "pre_lesson": "...\\n📌 [COT INDICATOR: ...]",
                    "instructional_flow": "...\\n📌 [COT INDICATOR: ...]",
                    "collaborative_activity": "...\\n📌 [COT INDICATOR: ...]",
                    "synthesis": "...\\n📌 [COT INDICATOR: ...]",
                    "integration": "...\\n📌 [COT INDICATOR: ...]",
                    "formative_assessment": "...\\n📌 [COT INDICATOR: ...]",
                    "extended_learning": "...",
                    "teacher_reflections": "..."
                }}

                CRITICAL REQUIREMENTS:
                1. Explicitly attach "📌 [COT INDICATOR: code: description]" at the end of paragraphs where that strategy is applied.
                2. Provide complete, detailed DepEd-aligned classroom activities.
                3. Output strictly raw JSON.
                """

                response_text, used_model = generate_with_fallback(client, candidate_models, prompt)
                clean_text = response_text.strip().replace("```json", "").replace("```", "")
                plan_data = json.loads(clean_text)

                metadata = {
                    "teacher_name": teacher_name,
                    "subject": subject,
                    "grade_level": grade_level,
                    "section": section,
                    "topic": topic,
                    "sessions": sessions,
                    "quarter": quarter,
                    "references": references
                }

                docx_file = create_deped_annex_a_docx(plan_data, metadata)

                st.success(f"Lesson Plan generated successfully using model `{used_model}`!")

                st.download_button(
                    label="📄 Download DepEd Annex A Word Document (.docx)",
                    data=docx_file,
                    file_name=f"ILAW_Lesson_Plan_{subject}_{topic}.docx".replace(" ", "_"),
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

                st.divider()
                st.subheader(f"ILAW LESSON PLAN ON {subject.upper()}")
                st.caption("DepEd Order No. 003, s. 2026 (Annex A Template)")

                st.json(plan_data)

                if not has_active_license and not user_data.get("trial_used", False):
                    mark_trial_as_used(user_email)
                    st.warning("⚠️ Free trial generation used.")

            except Exception as e:
                st.error(f"Error generating lesson plan: {e}")