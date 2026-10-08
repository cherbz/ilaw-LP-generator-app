import io
import re
import json
import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
import google.generativeai as genai
import firebase_admin
from firebase_admin import credentials, firestore
import streamlit as st

# ==============================================================================
# 1. FIREBASE INITIALIZATION & LICENSE VALIDATION
# ==============================================================================

@st.cache_resource
def init_firebase():
    """Initializes Firebase Admin SDK using Streamlit Secrets."""
    if not firebase_admin._apps:
        try:
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

def validate_and_claim_license(license_key, email_input):
    """
    Validates license key in Firestore.
    Fixes the 'even number of path elements' error by checking for empty inputs FIRST.
    """
    key = license_key.strip() if license_key else ""
    email = email_input.strip() if email_input else ""

    # Prevent empty path queries in Firestore
    if not key:
        return False, "Please enter a valid License Key."
    if not email:
        return False, "Please enter your Registered Email Address."

    if not db:
        return False, "Database connection is unavailable."

    try:
        key_ref = db.collection("license_keys").document(key)
        doc = key_ref.get()

        if not doc.exists:
            return False, "Invalid or unrecognized License Key."

        data = doc.to_dict()
        registered_email = data.get("email", "").strip().lower()

        if registered_email and registered_email != email.lower():
            return False, "Email address does not match the registered license owner."

        return True, "License validated successfully!"

    except Exception as e:
        return False, f"Database Error: {str(e)}"


# ==============================================================================
# 2. COT DATA STRUCTURES & CAREER STAGE MAPPING
# ==============================================================================

CAREER_STAGES = {
    "Teacher I": {"stage": "Beginning to Proficient", "scale": "2 to 6"},
    "Teacher II": {"stage": "Beginning to Proficient", "scale": "2 to 6"},
    "Teacher III": {"stage": "Beginning to Proficient", "scale": "2 to 6"},
    "Teacher IV": {"stage": "Proficient", "scale": "3 to 7"},
    "Teacher V": {"stage": "Proficient", "scale": "3 to 7"},
    "Teacher VI": {"stage": "Proficient", "scale": "3 to 7"},
    "Teacher VII": {"stage": "Proficient", "scale": "3 to 7"},
    "Master Teacher I": {"stage": "Highly Proficient", "scale": "4 to 8"},
    "Master Teacher II": {"stage": "Highly Proficient", "scale": "4 to 8"},
    "Master Teacher III": {"stage": "Distinguished", "scale": "5 to 9"},
    "Master Teacher IV": {"stage": "Distinguished", "scale": "5 to 9"},
    "Master Teacher V": {"stage": "Distinguished", "scale": "5 to 9"},
}

COT_INDICATORS_BY_SY = {
    "2025-2026": [
        ("1.1.2", "Apply knowledge of content within and across curriculum teaching areas"),
        ("1.4.2", "Use a range of teaching strategies that enhance learner achievement in literacy and numeracy skills"),
        ("1.5.2", "Apply a range of teaching strategies to develop critical and creative thinking, as well as other higher-order thinking skills"),
        ("2.3.2", "Manage classroom structure to engage learners, individually or in groups, in meaningful exploration, discovery and hands-on activities within a range of physical learning environments"),
        ("2.6.2", "Manage learner behavior constructively by applying positive and non-violent discipline to ensure learning-focused environments"),
        ("3.1.2", "Use differentiated, developmentally appropriate learning experiences to address learners' gender, needs, strengths, interests and experiences"),
        ("4.1.2", "Plan, manage and implement developmentally sequenced teaching and learning process to meet curriculum requirements and varied teaching contexts"),
        ("4.5.2", "Select, develop, organize and use appropriate teaching and learning resources, including ICT, to address learning goals"),
        ("5.1.2", "Design, select, organize and use diagnostic, formative and summative assessment strategies consistent with curriculum requirements"),
    ],
    "2026-2027": [
        ("1.1.2", "Apply knowledge of content within and across curriculum teaching areas"),
        ("1.4.2", "Use a range of teaching strategies that enhance learner achievement in literacy and numeracy skills"),
        ("1.5.2", "Apply a range of teaching strategies to develop critical and creative thinking, as well as other higher-order thinking skills"),
        ("1.6.2", "Display proficient use of Mother Tongue, Filipino and English to facilitate teaching and learning"),
        ("2.1.2", "Establish safe and secure learning environments to enhance learning through the consistent implementation of policies, guidelines and procedures"),
        ("2.2.2", "Maintain learning environments that promote fairness, respect and care to encourage learning"),
        ("3.2.2", "Establish a learner-centered culture by using teaching strategies that respond to learners' linguistic, cultural, socio-economic and religious backgrounds"),
        ("3.5.2", "Adapt and use culturally appropriate teaching strategies to address the needs of learners from indigenous groups"),
        ("5.3.2", "Use strategies for providing timely, accurate and constructive feedback to improve learner performance"),
    ],
    "2027-2028": [
        ("1.1.2", "Apply knowledge of content within and across curriculum teaching areas"),
        ("1.4.2", "Use a range of teaching strategies that enhance learner achievement in literacy and numeracy skills"),
        ("1.3.2", "Ensure the positive use of ICT to facilitate the teaching and learning process"),
        ("1.7.2", "Use effective verbal and non-verbal classroom communication strategies to support learner understanding, participation, engagement and achievement"),
        ("2.4.2", "Maintain supportive learning environments that nurture and inspire learners to participate, cooperate and collaborate in continued learning"),
        ("2.5.2", "Apply a range of successful strategies that maintain learning environments that motivate learners to work productively by assuming responsibility for their own learning"),
        ("3.3.2", "Design, adapt and implement teaching strategies that are responsive to learners with disabilities, giftedness and talents"),
        ("3.4.2", "Plan and deliver teaching strategies that are responsive to the special educational needs of learners in difficult circumstances"),
    ],
}

# ==============================================================================
# 3. STREAMLIT APP LAYOUT
# ==============================================================================

st.set_page_config(page_title="Binonz ILAW Lesson Plan Generator", page_icon="📝", layout="wide")

st.title("💡 Binonz ILAW Lesson Plan Generator")
st.caption("DepEd Order No. 003, s. 2026 (Annex A Template) | Automated COT Indicator Embedding")

# SIDEBAR: CREDENTIALS & TEACHER PROFILE
with st.sidebar:
    st.header("🔑 Authentication")
    license_key = st.text_input("License Key", type="password")
    user_email = st.text_input("Registered Email Address")
    api_key = st.text_input("Gemini API Key", type="password")

    st.divider()
    st.header("👤 Teacher Profile & Position")
    teacher_name = st.text_input("Teacher Name", "NORBERTO P. BINONDO JR.")
    position_rank = st.selectbox("Position / Rank", list(CAREER_STAGES.keys()), index=2)

    stage_info = CAREER_STAGES[position_rank]
    st.info(f"**Career Stage:** {stage_info['stage']}\n\n**COT Scale:** {stage_info['scale']}")

# MAIN FORM: LESSON DETAILS
st.subheader("1. Lesson Details")
col1, col2, col3 = st.columns(3)

with col1:
    grade_level = st.text_input("Grade Level & Section", "Grade 9 - Newton")
    subject = st.text_input("Learning Area", "Mathematics")

with col2:
    school_year = st.text_input("School Year", "2025-2026")
    sessions = st.text_input("No. of Sessions", "1")

with col3:
    topic = st.text_input("Name of Lesson / Topic", "Graphing Linear Functions")
    references = st.text_input("References", "DepEd Curriculum Guide & Presentation Slides")

st.divider()
st.subheader("2. Select COT Indicators to Integrate")

available_indicators = COT_INDICATORS_BY_SY.get(
    school_year.strip(), COT_INDICATORS_BY_SY["2025-2026"]
)
selected_indicators = []

for code, desc in available_indicators:
    if st.checkbox(f"**[{code}]** {desc}", value=True):
        selected_indicators.append(f"COT INDICATOR {code}: {desc}")

st.divider()
st.subheader("3. Learning Objectives & Context")
learning_competency = st.text_area(
    "Learning Competency",
    "Graphs a linear function and values its real-life applications (domain, range, intercepts, and slope).",
)
learner_context = st.text_area(
    "Learner Context",
    "The class is a mixed-ability group of learners with varied mathematical inclinations. Visual and kinesthetic learners benefit from coordinate plotting exercises.",
)

# ==============================================================================
# 4. HELPER FUNCTIONS TO CLEAN MATH & BUILD DOCX
# ==============================================================================

def clean_math_syntax(text: str) -> str:
    """Removes LaTeX dollar signs ($) and cleans math formatting."""
    cleaned = re.sub(r"\$+", "", text)
    cleaned = (
        cleaned.replace("\\", "")
        .replace("angle", "∠")
        .replace("circ", "°")
        .replace("&", "&")
    )
    return cleaned

def set_cell_background(cell, hex_color):
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)

def build_deped_ilaw_docx(header_data, content_dict):
    doc = docx.Document()

    for s in doc.sections:
        s.top_margin = Inches(0.8)
        s.bottom_margin = Inches(0.8)
        s.left_margin = Inches(0.8)
        s.right_margin = Inches(0.8)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run(f"ILAW LESSON PLAN ON {clean_math_syntax(header_data['subject']).upper()}\n")
    run_title.bold = True
    run_title.font.size = Pt(16)
    run_title.font.color.rgb = RGBColor(15, 32, 67)

    run_sub = p_title.add_run("DepEd Order No. 003, s. 2026 (Annex A Template) | COT Indicators Embedded")
    run_sub.font.size = Pt(10)
    run_sub.font.italic = True
    run_sub.font.color.rgb = RGBColor(100, 100, 100)

    doc.add_paragraph()

    meta_table = doc.add_table(rows=6, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_data = [
        ("Name of Lesson", clean_math_syntax(header_data["topic"])),
        ("Learning Area/s", clean_math_syntax(header_data["subject"])),
        ("Designed by Teacher/s", clean_math_syntax(header_data["teacher"])),
        ("Grade Level & Section", clean_math_syntax(header_data["grade"])),
        ("No. of Sessions", clean_math_syntax(header_data["sessions"])),
        ("References", clean_math_syntax(header_data["references"])),
    ]

    for idx, (label, val) in enumerate(meta_data):
        row = meta_table.rows[idx]
        cell_lbl, cell_val = row.cells[0], row.cells[1]
        cell_lbl.width = Inches(2.2)
        cell_