import os
import io
import json
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
    ],
    "Master Teacher III - V (Distinguished)": [
        "1.1.4: Lead colleagues in exploring innovative content applications.",
        "1.4.4: Lead in the design and evaluation of teaching strategies for literacy and numeracy.",
        "1.5.4: Lead colleagues in developing high-order thinking strategies."
    ]
}

# Helper function to set table cell background color in docx
def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

# -------------------------------------------------------------------
# DEPED ANNEX A WORD DOCX BUILDER (TABLE-BASED)
# -------------------------------------------------------------------
def create_deped_annex_a_docx(plan_data, metadata):
    doc = Document()

    # Page setup - Margins
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # Document Header Title
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

    # --- METADATA TABLE ---
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

    # --- MAIN SECTIONS TABLE BUILDER ---
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

    # Save to byte stream
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

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
    db.collection("users").document(email).update({"trial_used": True})

def redeem_license_key(email, key_string):
    key_ref = db.collection("license_keys").document(key_string)
    key_doc = key_ref.get()

    if not key_doc.exists:
        return False, "Invalid License Key. Please check for typos."

    key_data = key_doc.to_dict()
    if key_data.get("is_used"):
        return False, "This license key has already been redeemed."

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

    return True, "Success! 1-Year License successfully activated."

# -------------------------------------------------------------------
# STREAMLIT UI CONFIG
# -------------------------------------------------------------------
st.set_page_config(page_title="Binonz ILAW Lesson Plan Generator", page_icon="📘", layout="wide")

if "user_email" not in st.session_state:
    st.session_state["user_email"] = None

# Step 1: Login
if not st.session_state["user_email"]:
    st.title("📘 Binonz ILAW Lesson Plan Generator")
    st.markdown("Enter your email address to log in or start your **1 Free Trial** generation.")

    email_input = st.text_input("Enter Email Address:").strip().lower()

    if st.button("Continue"):
        if "@" in email_input and "." in email_input:
            st.session_state["user_email"] = email_input
            st.rerun()
        else:
            st.error("Please enter a valid email address.")
    st.stop()

user_email = st.session_state["user_email"]
user_data = get_or_create_user(user_email)

now = datetime.now(timezone.utc)
has_active_license = False

if user_data.get("license_expires_at"):
    expires_at = user_data["license_expires_at"]
    if expires_at > now:
        has_active_license = True

# Step 2: Sidebar Config & COT Selection
with st.sidebar:
    st.subheader("Account Overview")
    st.write(f"Logged in as: **{user_email}**")
    
    if st.button("Switch Account"):
        st.session_state["user_email"] = None
        st.rerun()

    st.divider()

    if has_active_license:
        days_left = (user_data["license_expires_at"] - now).days
        st.success(f"🟢 **Active Subscription** ({days_left} days left)")
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

    st.divider()

    # --- COT INDICATORS SELECTION PANEL ---
    st.subheader("📌 Target COT Indicators")
    teacher_rank = st.selectbox("Select Your Teacher Rank:", list(COT_INDICATOR_OPTIONS.keys()))
    
    available_indicators = COT_INDICATOR_OPTIONS[teacher_rank]
    selected_cots = []
    
    st.write("Check target indicators to embed:")
    for cot in available_indicators:
        if st.checkbox(cot, value=True, key=f"cot_{cot[:5]}"):
            selected_cots.append(cot)

# Step 3: Access Control
if not has_active_license and user_data.get("trial_used", False):
    st.error("🔒 **Trial Expired**")
    st.write("You have used your 1 free trial generation. Redeem a 1-Year License Key in the sidebar to generate more.")
    st.stop()
elif not has_active_license:
    st.info("🎁 **Free Trial Available:** You have 1 free lesson plan generation remaining.")

# Step 4: Main Form Interface
st.title("DepEd ILAW Lesson Plan Generator")
st.caption("Aligned with DepEd Order No. 003, s. 2026 Annex A Template & COT Indicators")

with st.form("deped_ilaw_form"):
    col1, col2 = st.columns(2)
    with col1:
        teacher_name = st.text_input("Teacher Name", value="NORBERTO P. BINONDO JR.")
        subject = st.text_input("Learning Area / Subject", value="Mathematics")
        grade_level = st.selectbox("Grade Level", ["Grade 7", "Grade 8", "Grade 9", "Grade 10", "Grade 11", "Grade 12"], index=2)
        section = st.text_input("Section", value="Kindness")
    with col2:
        topic = st.text_input("Lesson Topic / Competency", value="Graphing Linear Functions")
        sessions = st.text_input("No. of Sessions", value="1")
        quarter = st.selectbox("Quarter", ["Quarter 1", "Quarter 2", "Quarter 3", "Quarter 4"])
        references = st.text_input("References", value="DepEd Curriculum Guide & Presentation Slides")

    submit_button = st.form_submit_button("Generate DepEd ILAW Lesson Plan")

if submit_button:
    if not subject or not topic:
        st.warning("Please fill in all required fields.")
    elif not selected_cots:
        st.warning("Please select at least one COT Indicator in the sidebar.")
    else:
        if not configure_gemini():
            st.error("Missing GEMINI_API_KEY in secrets/environment variables.")
        else:
            with st.spinner("Generating DepEd Order No. 003 Annex A ILAW Lesson Plan..."):
                try:
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
                    2. Provide complete, detailed DepEd-aligned classroom activities, not short summaries.
                    3. Output strictly raw JSON (no Markdown block fences, no prose outside JSON).
                    """

                    # Updated Model Fallback Chain to prevent 404 endpoint errors
                    model_names = ['gemini-2.5-flash', 'gemini-2.0-flash', 'gemini-1.5-flash-latest']
                    response = None
                    last_error = None

                    for model_name in model_names:
                        try:
                            model = genai.GenerativeModel(model_name)
                            response = model.generate_content(prompt)
                            if response and response.text:
                                break
                        except Exception as e:
                            last_error = e
                            continue

                    if not response or not response.text:
                        raise Exception(f"Failed to generate content with available models: {last_error}")
                    
                    # Parse JSON Output
                    clean_text = response.text.strip().replace("```json", "").replace("```", "")
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

                    # Render Document in Memory
                    docx_file = create_deped_annex_a_docx(plan_data, metadata)

                    st.success("Lesson Plan successfully generated!")

                    # Download Button
                    st.download_button(
                        label="📄 Download as DepEd Annex A Word Document (.docx)",
                        data=docx_file,
                        file_name=f"ILAW_Lesson_Plan_{subject}_{topic}.docx".replace(" ", "_"),
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

                    # Display Preview UI
                    st.divider()
                    st.subheader(f"ILAW LESSON PLAN ON {subject.upper()}")
                    st.caption("DepEd Order No. 003, s. 2026 (Annex A Template)")

                    st.json(plan_data)

                    # Consume trial if un-subscribed
                    if not has_active_license and not user_data.get("trial_used", False):
                        mark_trial_as_used(user_email)
                        st.warning("⚠️ Free trial generation used. Please activate a 1-Year License key for continued access.")

                except Exception as e:
                    st.error(f"Error generating lesson plan: {e}")