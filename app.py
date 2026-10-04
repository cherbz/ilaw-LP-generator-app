import os
import json
import io
import re
import streamlit as st
import google.generativeai as genai
import firebase_admin
from firebase_admin import credentials, firestore
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

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

        if not key_dict:
            env_var = os.getenv("text_key") or os.getenv("FIREBASE_CREDENTIALS") or os.getenv("FIREBASE_SERVICE_ACCOUNT")
            if env_var:
                try:
                    key_dict = json.loads(env_var)
                except Exception as e:
                    st.error(f"Failed to parse Firebase JSON from environment variable: {e}")
                    st.stop()

        if not key_dict and os.path.exists("serviceAccountKey.json"):
            try:
                with open("serviceAccountKey.json", "r") as f:
                    key_dict = json.load(f)
            except Exception as e:
                st.error(f"Failed to load serviceAccountKey.json: {e}")
                st.stop()

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
# 3. HELPER FUNCTIONS & CLEANING
# ==========================================
def clean_latex_math(text: str) -> str:
    """Removes math dollar signs ($or$$) and LaTeX slashes to keep plain text formatting."""     if not isinstance(text, str):         return text     # Strip dollar signs     cleaned = text.replace("$$", "").replace("$", "")
    # Clean common LaTeX math symbols
    cleaned = cleaned.replace("\\f", "f").replace("\\[", "").replace("\\]", "").replace("\\(", "").replace("\\)", "")
    return cleaned

def validate_and_claim_license(license_key: str, email: str) -> tuple[bool, str]:
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
    
    key_ref.update({
        "is_used": True,
        "used_by": email.strip().lower(),
        "used_at": firestore.SERVER_TIMESTAMP
    })
    return True, "License key successfully activated!"

def generate_lesson_plan_content(api_key: str, prompt: str) -> str:
    genai.configure(api_key=api_key.strip())
    
    candidate_models = []
    try:
        all_models = list(genai.list_models())
        supported_models = [
            m.name.replace("models/", "") 
            for m in all_models 
            if "generateContent" in m.supported_generation_methods
        ]
        
        flash_models = [m for m in supported_models if "flash" in m]
        pro_models = [m for m in supported_models if "pro" in m]
        other_models = [m for m in supported_models if m not in flash_models and m not in pro_models]
        
        candidate_models = flash_models + pro_models + other_models
    except Exception:
        candidate_models = ["gemini-2.0-flash", "gemini-1.5-flash", "gemini-pro"]

    if not candidate_models:
        candidate_models = ["gemini-2.0-flash", "gemini-1.5-flash", "gemini-pro"]

    last_error = None
    for model_name in candidate_models:
        try:
            model = genai.GenerativeModel(model_name)
            response = model.generate_content(prompt)
            if response and response.text:
                return response.text
        except Exception as e:
            last_error = e
            continue

    raise Exception(f"Unable to generate content with provided key. Attempted models: {candidate_models}. Last error: {str(last_error)}")

def create_docx_ilaw_template(data_dict: dict) -> io.BytesIO:
    """Generates a structured DepEd Order No. 003, s. 2026 Semi-Detailed ILAW Word Document."""
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run(f"SEMI-DETAILED ILAW LESSON PLAN ON {clean_latex_math(data_dict.get('learning_area', 'MATHEMATICS')).upper()}")
    run_title.bold = True
    run_title.font.size = Pt(14)
    run_title.font.name = "Arial"

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = p_sub.add_run("DepEd Order No. 003, s. 2026 (Annex A Template) | COT Indicators Embedded")
    run_sub.font.size = Pt(9)
    run_sub.font.italic = True
    run_sub.font.color.rgb = RGBColor(100, 100, 100)

    meta_table = doc.add_table(rows=6, cols=2)
    meta_table.autofit = False
    
    meta_data = [
        ("Name of Lesson", clean_latex_math(data_dict.get("lesson_name", ""))),
        ("Learning Area/s", clean_latex_math(data_dict.get("learning_area", ""))),
        ("Designed by Teacher/s", clean_latex_math(data_dict.get("teacher_name", ""))),
        ("Grade Level & Section", f"{clean_latex_math(data_dict.get('grade_level', ''))} - {clean_latex_math(data_dict.get('section', 'Kindness'))}"),
        ("No. of Sessions", "1"),
        ("References", "DepEd Curriculum Guide & Presentation Slides")
    ]

    for i, (label, val) in enumerate(meta_data):
        row = meta_table.rows[i]
        cell_lbl, cell_val = row.cells[0], row.cells[1]
        cell_lbl.width = Inches(2.2)
        cell_val.width = Inches(4.8)
        
        p_lbl = cell_lbl.paragraphs[0]
        p_lbl.add_run(label).bold = True
        p_lbl.paragraph_format.space_after = Pt(2)

        p_val = cell_val.paragraphs[0]
        p_val.add_run(val)
        p_val.paragraph_format.space_after = Pt(2)

    doc.add_paragraph()

    ilaw_sections = [
        ("1. INTENTIONS", [
            ("Learning Competency", clean_latex_math(data_dict.get("competency", ""))),
            ("Learning Objectives", clean_latex_math(data_dict.get("objectives", ""))),
            ("Learner Context", clean_latex_math(data_dict.get("learner_context", "")))
        ]),
        ("2. LEARNING EXPERIENCE", [
            ("Pre-Lesson (Getting Ready)", clean_latex_math(data_dict.get("pre_lesson", ""))),
            ("Instructional Flow & Direct Modeling", clean_latex_math(data_dict.get("direct_modeling", ""))),
            ("Collaborative Group Activity", clean_latex_math(data_dict.get("group_activity", ""))),
            ("Synthesis & Resources", clean_latex_math(data_dict.get("synthesis", ""))),
            ("Opportunities for Integration", clean_latex_math(data_dict.get("integration", "")))
        ]),
        ("3. ASSESSMENT", [
            ("Formative Assessment (Individual Evaluation)", clean_latex_math(data_dict.get("assessment", "")))
        ]),
        ("4. WAYS FORWARD", [
            ("Extended Learning Opportunities", clean_latex_math(data_dict.get("extended_learning", ""))),
            ("Teacher Reflections", clean_latex_math(data_dict.get("reflections", "")))
        ])
    ]

    for sec_title, items in ilaw_sections:
        p_sec = doc.add_paragraph()
        r_sec = p_sec.add_run(sec_title)
        r_sec.bold = True
        r_sec.font.size = Pt(11)
        r_sec.font.color.rgb = RGBColor(0, 51, 102)

        table = doc.add_table(rows=len(items), cols=2)
        for idx, (lbl, text_val) in enumerate(items):
            row = table.rows[idx]
            c_lbl, c_val = row.cells[0], row.cells[1]
            c_lbl.width = Inches(2.2)
            c_val.width = Inches(4.8)

            p_l = c_lbl.paragraphs[0]
            p_l.add_run(lbl).bold = True
            
            p_v = c_val.paragraphs[0]
            p_v.add_run(text_val)

        doc.add_paragraph()

    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

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
st.title("💡 Binonz Semi-Detailed ILAW Lesson Plan Generator")

# --- SECTION 1: BASIC INFORMATION ---
st.subheader("1. Basic Information & Meta Details")
col1, col2, col3 = st.columns(3)
with col1:
    grade_level = st.selectbox("Grade Level", ["Grade 7", "Grade 8", "Grade 9", "Grade 10", "Grade 11", "Grade 12"])
    learning_area = st.text_input("Learning Area / Subject", value="Mathematics")
    lesson_name = st.text_input("Name of Lesson", value="Graphing Linear Functions")
with col2:
    section_name = st.text_input("Section", value="Kindness")
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
if st.button("🚀 Generate Semi-Detailed Lesson Plan", type="primary", use_container_width=True):
    if not license_key_input or not email_input or not api_key_input:
        st.error("Please fill in your License Key, Registered Email, and Gemini API Key in the sidebar before proceeding.")
    else:
        with st.spinner("Verifying License..."):
            is_valid, msg = validate_and_claim_license(license_key_input, email_input)

        if not is_valid:
            st.error(f"License Error: {msg}")
        else:
            st.success(msg)
            with st.spinner("Generating Semi-Detailed DepEd ILAW Lesson Plan (DO No. 003, s. 2026)..."):
                
                selected_cots = []
                if cot1: selected_cots.append("📌 [COT INDICATOR: 1.1.2: Apply knowledge of content within and across curriculum teaching areas.]")
                if cot2: selected_cots.append("📌 [COT INDICATOR: 1.4.2: Use a range of teaching strategies that enhance learner achievement in literacy and numeracy skills.]")
                if cot3: selected_cots.append("📌 [COT INDICATOR: 1.5.2: Apply a range of teaching strategies to develop critical and creative thinking.]")
                if cot4: selected_cots.append("📌 [COT INDICATOR: 4.5.2: Select, develop, organize, and use appropriate teaching and learning resources, including ICT.]")
                if cot5: selected_cots.append("📌 [COT INDICATOR: 2.3.2: Manage classroom structure to engage learners in meaningful exploration, discovery and hands-on activities.]")
                if cot6: selected_cots.append("📌 [COT INDICATOR: 3.1.2: Use differentiated, developmentally appropriate learning experiences.]")
                if cot7: selected_cots.append("📌 [COT INDICATOR: 5.1.2: Design, select, organize and use diagnostic, formative and summative assessment strategies.]")
                if cot8: selected_cots.append("📌 [COT INDICATOR: 5.2.2: Monitor and evaluate learner progress and achievement using learner attainment data.]")

                prompt = f"""
                You are a DepEd Curriculum Specialist. Generate a comprehensive SEMI-DETAILED DepEd Order No. 003, s. 2026 (Annex A) ILAW Lesson Plan.

                CRITICAL FORMAT RULES:
                1. DO NOT USE ANY DOLLAR SIGNS ($) OR LATEX FORMATTING anywhere in the response. Write mathematical expressions using plain text (e.g., f(x) = 2x + 1, y = 3x - 4).
                2. Write a SEMI-DETAILED lesson plan format (provide complete step-by-step procedures, teacher actions, student responses, and clear activity instructions in every section).
                3. Return ONLY a valid JSON object matching this schema without markdown codeblocks:

                {{
                  "lesson_name": "{lesson_name}",
                  "learning_area": "{learning_area}",
                  "teacher_name": "{teacher_name}",
                  "grade_level": "{grade_level}",
                  "section": "{section_name}",
                  "competency": "{learning_competency}",
                  "objectives": "1. Cognitive objective... 2. Psychomotor objective... 3. Affective objective... 📌 [COT INDICATOR: 1.4.2: ...]",
                  "learner_context": "Semi-detailed learner context description 📌 [COT INDICATOR: 3.1.2: ...]",
                  "pre_lesson": "Semi-detailed preliminary activity: 1. Drill/Warm-up, 2. Review, 3. Motivation with teacher and student tasks 📌 [COT INDICATOR: ...]",
                  "direct_modeling": "Semi-detailed direct modeling step-by-step: Example 1, Example 2 with plain text math formulas and teacher explanations 📌 [COT INDICATOR: ...]",
                  "group_activity": "Semi-detailed collaborative group activity instructions for Group 1, Group 2, Group 3, and Group 4 📌 [COT INDICATOR: ...]",
                  "synthesis": "Semi-detailed discussion, debrief questions, and key takeaways 📌 [COT INDICATOR: ...]",
                  "integration": "Semi-detailed cross-curricular integration (e.g., Science, Economics, Health) 📌 [COT INDICATOR: ...]",
                  "assessment": "Semi-detailed evaluation items and scoring guide 📌 [COT INDICATOR: ...]",
                  "extended_learning": "Detailed assignment/enrichment activity",
                  "reflections": "Teacher reflections and next steps"
                }}

                Target COT Indicators to embed inline:
                {", ".join(selected_cots)}
                """

                try:
                    raw_res = generate_lesson_plan_content(api_key_input, prompt)
                    
                    clean_json = raw_res.strip()
                    if clean_json.startswith("```json"):
                        clean_json = clean_json[7:]
                    if clean_json.startswith("```"):
                        clean_json = clean_json[3:]
                    if clean_json.endswith("```"):
                        clean_json = clean_json[:-3]
                    
                    ilaw_data = json.loads(clean_json.strip())

                    # Clean dollar signs across all dictionary entries
                    for k in ilaw_data:
                        ilaw_data[k] = clean_latex_math(ilaw_data[k])

                    st.markdown("---")
                    st.subheader("📋 Generated Semi-Detailed DepEd ILAW Lesson Plan (DO No. 003, s. 2026)")

                    st.markdown(f"### SEMI-DETAILED ILAW LESSON PLAN ON {ilaw_data.get('learning_area', 'MATHEMATICS').upper()}")
                    st.caption("DepEd Order No. 003, s. 2026 (Annex A Template) | COT Indicators Embedded")

                    st.markdown("#### 1. INTENTIONS")
                    st.write(f"**Learning Competency:** {ilaw_data.get('competency')}")
                    st.write(f"**Learning Objectives:** {ilaw_data.get('objectives')}")
                    st.write(f"**Learner Context:** {ilaw_data.get('learner_context')}")

                    st.markdown("#### 2. LEARNING EXPERIENCE")
                    st.write(f"**Pre-Lesson (Getting Ready):** {ilaw_data.get('pre_lesson')}")
                    st.write(f"**Instructional Flow & Direct Modeling:** {ilaw_data.get('direct_modeling')}")
                    st.write(f"**Collaborative Group Activity:** {ilaw_data.get('group_activity')}")
                    st.write(f"**Synthesis & Resources:** {ilaw_data.get('synthesis')}")
                    st.write(f"**Opportunities for Integration:** {ilaw_data.get('integration')}")

                    st.markdown("#### 3. ASSESSMENT")
                    st.write(f"**Formative Assessment:** {ilaw_data.get('assessment')}")

                    st.markdown("#### 4. WAYS FORWARD")
                    st.write(f"**Extended Learning Opportunities:** {ilaw_data.get('extended_learning')}")
                    st.write(f"**Teacher Reflections:** {ilaw_data.get('reflections')}")

                    docx_file = create_docx_ilaw_template(ilaw_data)

                    st.markdown("---")
                    st.download_button(
                        label="📄 Download Semi-Detailed Word File (.docx)",
                        data=docx_file,
                        file_name=f"Semi_Detailed_ILAW_Lesson_Plan_{lesson_name.replace(' ', '_')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary"
                    )

                except Exception as err:
                    st.error(f"Generation or Parsing Error: {str(err)}")