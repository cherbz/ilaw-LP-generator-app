import io
import re
import json
import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
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
    """Validates license key in Firestore safely."""
    key = license_key.strip() if license_key else ""
    email = email_input.strip() if email_input else ""

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
    teacher_name = st.text_input("Teacher Name", "JUAN DELA CRUZ")
    position_rank = st.selectbox("Position / Rank", list(CAREER_STAGES.keys()), index=2)

    stage_info = CAREER_STAGES[position_rank]
    st.info(f"**Career Stage:** {stage_info['stage']}\n\n**COT Scale:** {stage_info['scale']}")

# MAIN FORM: LESSON DETAILS
st.subheader("1. Lesson Details")
col1, col2, col3 = st.columns(3)

with col1:
    grade_level = st.text_input("Grade Level & Section", "Grade 9 - Kindness")
    subject = st.text_input("Learning Area", "Araling Panlipunan")

with col2:
    school_year = st.text_input("School Year", "2025-2026")
    sessions = st.text_input("No. of Sessions", "3")

with col3:
    topic = st.text_input("Name of Lesson / Topic", "Demand and Suplay")
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
st.subheader("3. Learning Objectives")
learning_competency = st.text_area(
    "Learning Competency",
    "Nasusuri ang interaksyon ng demand at suplay at implikasyon nito sa kalagayan ng presyo at ng pamilihan.",
)

# ==============================================================================
# 4. DOCX CLONING HELPER FUNCTIONS (MATCHING CAMBRIA & SCREENSHOT STYLING)
# ==============================================================================

def clean_math_syntax(text: str) -> str:
    """Removes LaTeX dollar signs ($) and cleans math/internal reasoning commentary."""
    cleaned = re.sub(r"\*+.*?\*+", "", text, flags=re.DOTALL)
    cleaned = re.sub(r"\(Mental draft.*?\)", "", cleaned, flags=re.DOTALL | re.IGNORECASE)
    cleaned = re.sub(r"\$+", "", cleaned)
    cleaned = (
        cleaned.replace("\\", "")
        .replace("angle", "∠")
        .replace("circ", "°")
        .replace("&", "&")
    )
    return cleaned.strip()

def set_cell_background(cell, hex_color):
    """Sets cell background fill."""
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    """Sets internal padding for table cells."""
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>\n'
        f'  <w:top w:w="{top}" w:type="dxa"/>\n'
        f'  <w:bottom w:w="{bottom}" w:type="dxa"/>\n'
        f'  <w:left w:w="{left}" w:type="dxa"/>\n'
        f'  <w:right w:w="{right}" w:type="dxa"/>\n'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)

def remove_table_borders(table):
    """Removes outer and inner borders from Word tables."""
    tblPr = table._element.xpath('w:tblPr')
    if tblPr:
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>\n'
            f'  <w:top w:val="none"/>\n'
            f'  <w:left w:val="none"/>\n'
            f'  <w:bottom w:val="none"/>\n'
            f'  <w:right w:val="none"/>\n'
            f'  <w:insideH w:val="none"/>\n'
            f'  <w:insideV w:val="none"/>\n'
            f'</w:tblBorders>'
        )
        tblPr[0].append(borders)

def build_deped_ilaw_docx(header_data, content_dict):
    """Clones the exact layout, font, sizes, and highlights from target document."""
    doc = docx.Document()

    # Set document margins
    for s in doc.sections:
        s.top_margin = Inches(0.8)
        s.bottom_margin = Inches(0.8)
        s.left_margin = Inches(0.8)
        s.right_margin = Inches(0.8)

    # Document Title Banner
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(2)

    run_title = p_title.add_run(f"ILAW LESSON PLAN ON {clean_math_syntax(header_data['subject']).upper()}\n")
    run_title.font.name = "Cambria"
    run_title.bold = True
    run_title.font.size = Pt(16)
    run_title.font.color.rgb = RGBColor(15, 32, 67)

    run_sub = p_title.add_run("DepEd Order No. 003, s. 2026 (Annex A Template) | COT Indicators Embedded")
    run_sub.font.name = "Cambria"
    run_sub.font.size = Pt(10)
    run_sub.font.italic = True
    run_sub.font.color.rgb = RGBColor(100, 100, 100)

    doc.add_paragraph()

    # Metadata Table
    meta_table = doc.add_table(rows=7, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(meta_table)

    meta_data = [
        ("Name of Lesson", clean_math_syntax(header_data["topic"])),
        ("Learning Area/s", clean_math_syntax(header_data["subject"])),
        ("Designed by Teacher/s", clean_math_syntax(header_data["teacher"])),
        ("Grade Level & Section", clean_math_syntax(header_data["grade"])),
        ("No. of Sessions", clean_math_syntax(header_data["sessions"])),
        ("References", clean_math_syntax(header_data["references"])),
        ("Declaration of AI Use", "AI was utilized to structure content into DepEd Order No. 003, s. 2026 Annex A template & align COT indicators."),
    ]

    for idx, (label, val) in enumerate(meta_data):
        row = meta_table.rows[idx]
        cell_lbl, cell_val = row.cells[0], row.cells[1]
        cell_lbl.width = Inches(2.2)
        cell_val.width = Inches(4.5)

        set_cell_background(cell_lbl, "F2F4F8")
        set_cell_margins(cell_lbl, top=80, bottom=80, left=120, right=120)
        set_cell_margins(cell_val, top=80, bottom=80, left=120, right=120)

        p_lbl = cell_lbl.paragraphs[0]
        p_lbl.paragraph_format.space_after = Pt(2)
        r_lbl = p_lbl.add_run(label)
        r_lbl.font.name = "Cambria"
        r_lbl.bold = True
        r_lbl.font.size = Pt(10)
        r_lbl.font.color.rgb = RGBColor(0, 0, 0)

        p_val = cell_val.paragraphs[0]
        p_val.paragraph_format.space_after = Pt(2)
        r_val = p_val.add_run(val)
        r_val.font.name = "Cambria"
        r_val.font.size = Pt(10)
        r_val.font.color.rgb = RGBColor(0, 0, 0)

    doc.add_paragraph()

    # Section Table Generator
    def add_section_table(section_title, rows_data):
        h_p = doc.add_paragraph()
        h_p.paragraph_format.space_before = Pt(10)
        h_p.paragraph_format.space_after = Pt(4)
        h_run = h_p.add_run(section_title)
        h_run.font.name = "Cambria"
        h_run.bold = True
        h_run.font.size = Pt(12)
        h_run.font.color.rgb = RGBColor(15, 32, 67)

        tbl = doc.add_table(rows=len(rows_data), cols=2)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        remove_table_borders(tbl)

        cot_regex = re.compile(r"((?:📌|📌|\*|\+)?\s*\[COT INDICATOR:[^\]]+\])", re.IGNORECASE)

        for r_idx, (lbl, text_content) in enumerate(rows_data):
            row = tbl.rows[r_idx]
            c_lbl, c_val = row.cells[0], row.cells[1]
            c_lbl.width = Inches(2.2)
            c_val.width = Inches(4.5)

            set_cell_background(c_lbl, "EBF3FC")
            set_cell_margins(c_lbl, top=80, bottom=80, left=120, right=120)
            set_cell_margins(c_val, top=80, bottom=80, left=120, right=120)

            p1 = c_lbl.paragraphs[0]
            p1.paragraph_format.space_after = Pt(2)
            r1 = p1.add_run(lbl)
            r1.font.name = "Cambria"
            r1.bold = True
            r1.font.size = Pt(10)
            r1.font.color.rgb = RGBColor(0, 0, 0)

            p2 = c_val.paragraphs[0]
            p2.paragraph_format.space_after = Pt(4)
            clean_text = clean_math_syntax(text_content.strip())

            lines = clean_text.split("\n")
            for l_idx, line in enumerate(lines):
                if not line.strip():
                    continue
                if l_idx > 0:
                    p2 = c_val.add_paragraph()
                    p2.paragraph_format.space_after = Pt(4)

                segments = cot_regex.split(line)
                for seg in segments:
                    if not seg:
                        continue
                    if "[COT INDICATOR" in seg.upper():
                        r_cot = p2.add_run(f" {seg.strip()} ")
                        r_cot.font.name = "Cambria"
                        r_cot.bold = True
                        r_cot.font.size = Pt(10)
                        r_cot.font.color.rgb = RGBColor(0, 0, 0)
                        r_cot.font.highlight_color = WD_COLOR_INDEX.YELLOW
                    else:
                        r_norm = p2.add_run(seg)
                        r_norm.font.name = "Cambria"
                        r_norm.font.size = Pt(10)
                        r_norm.font.color.rgb = RGBColor(0, 0, 0)
                        r_norm.font.highlight_color = WD_COLOR_INDEX.AUTO

        doc.add_paragraph()

    # Section 1: INTENTIONS
    add_section_table(
        "1. INTENTIONS",
        [
            ("Learning Competency", content_dict.get("Learning Competency", "")),
            ("Learning Objectives", content_dict.get("Learning Objectives", "")),
            ("Learner Context", content_dict.get("Learner Context", "")),
        ],
    )

    # Section 2: LEARNING EXPERIENCE
    add_section_table(
        "2. LEARNING EXPERIENCE",
        [
            ("Pre-Lesson (Getting Ready)", content_dict.get("Pre-Lesson", "")),
            ("Instructional Flow & Direct Modeling", content_dict.get("Instructional Flow", "")),
            ("Collaborative Group Activity", content_dict.get("Collaborative Group Activity", "")),
            ("Synthesis & Resources", content_dict.get("Synthesis & Resources", "")),
            ("Opportunities for Integration", content_dict.get("Opportunities for Integration", "")),
        ],
    )

    # Section 3: ASSESSMENT
    add_section_table(
        "3. ASSESSMENT",
        [("Formative Assessment (Individual Evaluation)", content_dict.get("Formative Assessment", ""))],
    )

    # Section 4: WAYS FORWARD
    add_section_table(
        "4. WAYS FORWARD",
        [
            ("Extended Learning Opportunities", content_dict.get("Extended Learning Opportunities", "")),
            ("Teacher Reflections", content_dict.get("Teacher Reflections", "")),
        ],
    )

    doc_buffer = io.BytesIO()
    doc.save(doc_buffer)
    doc_buffer.seek(0)
    return doc_buffer


# ==============================================================================
# 5. GENERATION ENGINE & DYNAMIC GEMINI MODEL FINDER
# ==============================================================================

st.divider()

if st.button("🚀 Generate Lesson Plan", type="primary", use_container_width=True):
    if not api_key:
        st.error("Please enter a valid Gemini API Key in the sidebar.")
    elif not license_key or not user_email:
        st.warning("Please enter your registered Email and License Key.")
    else:
        is_valid, msg = validate_and_claim_license(license_key, user_email)
        if not is_valid:
            st.error(msg)
        else:
            st.success(msg)
            with st.spinner(f"Generating DepEd Order No. 003, s. 2026 (Annex A) ILAW Lesson Plan for {sessions} session(s)..."):
                try:
                    genai.configure(api_key=api_key.strip())

                    available_models = []
                    try:
                        for m in genai.list_models():
                            if 'generateContent' in m.supported_generation_methods:
                                model_name = m.name.replace("models/", "")
                                available_models.append(model_name)
                    except Exception:
                        pass

                    fallback_models = [
                        "gemini-1.5-flash",
                        "gemini-1.5-flash-latest",
                        "gemini-1.5-pro",
                        "gemini-1.5-pro-latest",
                        "gemini-2.0-flash-exp",
                        "gemini-2.5-flash"
                    ]

                    model_candidates = list(dict.fromkeys(available_models + fallback_models))
                    cot_prompt_text = "\n".join([f"- {ind}" for ind in selected_indicators])

                    prompt = f"""
                    You are an expert DepEd Instructional Designer. Create an official DepEd ILAW Lesson Plan adhering strictly to DepEd Order No. 003, s. 2026 (Annex A Template).

                    STRICT NO-INTERNAL-THOUGHTS RULE:
                    - DO NOT include internal thoughts, mental drafts, commentary, reasoning notes, or rules checking (e.g. DO NOT include '*Check formatting rules again:*' or '*Developing the Content*').
                    - Start your response IMMEDIATELY with 'Learning Competency:::' and proceed directly with the template data.

                    CRITICAL RULE FOR OBJECTIVES:
                    - DO NOT place any COT indicator tags in the 'Learning Objectives' section. Keep Learning Objectives purely focused on SMART learning goals.

                    CRITICAL MULTI-SESSION INSTRUCTION:
                    - Total Sessions Specified: {sessions}
                    - When Total Sessions is greater than 1 (e.g., 3), write out distinct session paragraphs for 'Learning Objectives', 'Pre-Lesson', 'Instructional Flow', 'Collaborative Group Activity', 'Synthesis & Resources', and 'Formative Assessment'.
                    - Structure each session paragraph clearly on a new line as:
                      Session 1: [Detailed activity/objective...]
                      Session 2: [Detailed activity/objective...]
                      Session 3: [Detailed activity/objective...]

                    CRITICAL AUTOMATIC LEARNER CONTEXT INSTRUCTION:
                    - Automatically generate a detailed, realistic 'Learner Context' tailored specifically for Grade Level/Section: '{grade_level}', Subject: '{subject}', and Topic: '{topic}'. Include a COT indicator tag at the end.

                    CRITICAL FORMATTING RULES FOR OTHER SECTIONS:
                    1. DO NOT use LaTeX syntax or dollar signs ($) for mathematical formulas or variables. Write plain text.
                    2. IMPORTANT FOR COT INDICATORS: Embed selected COT indicators ONLY inside 'Learner Context', 'Pre-Lesson', 'Instructional Flow', 'Collaborative Group Activity', 'Synthesis & Resources', 'Opportunities for Integration', and 'Formative Assessment'. Place each indicator tag strictly at the VERY END of the sentence/paragraph that demonstrates it. Write as 📌 [COT INDICATOR: code: description]

                    HEADER DETAILS:
                    - Teacher: {teacher_name} ({position_rank} - Stage: {stage_info['stage']})
                    - Grade & Section: {grade_level} | Learning Area: {subject}
                    - School Year: {school_year} | No. of Sessions: {sessions}
                    - Lesson Name/Topic: {topic}
                    - References: {references}

                    INPUTS:
                    - Learning Competency: {learning_competency}

                    TARGET COT INDICATORS TO EMBED:
                    {cot_prompt_text if cot_prompt_text else "Apply standard pedagogical strategies."}

                    STRUCTURE OUTPUT USING ':::' AS DELIMITER:

                    Learning Competency::: {learning_competency}
                    Learning Objectives::: Write Session 1:, Session 2:, etc., SMART objectives without any COT indicator tags.
                    Learner Context::: Provide the automatically generated learner context. 📌 [COT INDICATOR: code: description]
                    Pre-Lesson::: Write detailed warmup/review for Session 1:, Session 2:, etc., on separate lines with COT tags at the end of each.
                    Instructional Flow::: Write detailed instruction/modeling for Session 1:, Session 2:, etc., on separate lines with COT tags at the end of each.
                    Collaborative Group Activity::: Write detailed group activities for Session 1:, Session 2:, etc., on separate lines with COT tags at the end of each.
                    Synthesis & Resources::: Write debriefing questions for Session 1:, Session 2:, etc., on separate lines with COT tags at the end of each.
                    Opportunities for Integration::: Write cross-curricular linkages. 📌 [COT INDICATOR: code: description]
                    Formative Assessment::: Write evaluation tasks for Session 1:, Session 2:, etc., on separate lines with COT tags at the end of each.
                    Extended Learning Opportunities::: Write homework or remedial tasks here.
                    Teacher Reflections::: Write reflective notes on student performance.
                    """

                    response = None
                    successful_model = None
                    last_error = None

                    for model_name in model_candidates:
                        try:
                            model = genai.GenerativeModel(model_name)
                            response = model.generate_content(prompt)
                            successful_model = model_name
                            break
                        except Exception as err:
                            last_error = err
                            continue

                    if response is None:
                        raise Exception(f"Unable to generate content with provided key. Last error: {str(last_error)}")

                    # Apply regex cleaning to remove any residual internal commentary
                    raw_text = clean_math_syntax(response.text)

                    keys_list = [
                        "Learning Competency",
                        "Learning Objectives",
                        "Learner Context",
                        "Pre-Lesson",
                        "Instructional Flow",
                        "Collaborative Group Activity",
                        "Synthesis & Resources",
                        "Opportunities for Integration",
                        "Formative Assessment",
                        "Extended Learning Opportunities",
                        "Teacher Reflections",
                    ]

                    parsed_content = {}
                    for i, k in enumerate(keys_list):
                        if i < len(keys_list) - 1:
                            next_k = keys_list[i + 1]
                            pattern = rf"{re.escape(k)}:::(.*?)(?={re.escape(next_k)}:::|$)"
                        else:
                            pattern = rf"{re.escape(k)}:::(.*)"

                        match = re.search(pattern, raw_text, re.DOTALL)
                        if match:
                            parsed_content[k] = clean_math_syntax(match.group(1).strip())
                        else:
                            parsed_content[k] = "N/A"

                    header_info = {
                        "topic": topic,
                        "subject": subject,
                        "teacher": teacher_name,
                        "grade": grade_level,
                        "sessions": sessions,
                        "references": references,
                    }

                    st.success(f"Official DepEd ILAW Lesson Plan generated using model: `{successful_model}`!")

                    # Streamlit UI Preview
                    st.markdown(f"### ILAW LESSON PLAN ON {subject.upper()}\n*DepEd Order No. 003, s. 2026 (Annex A)*")
                    for section_title, keys in [
                        ("1. INTENTIONS", ["Learning Competency", "Learning Objectives", "Learner Context"]),
                        ("2. LEARNING EXPERIENCE", ["Pre-Lesson", "Instructional Flow", "Collaborative Group Activity", "Synthesis & Resources", "Opportunities for Integration"]),
                        ("3. ASSESSMENT", ["Formative Assessment"]),
                        ("4. WAYS FORWARD", ["Extended Learning Opportunities", "Teacher Reflections"]),
                    ]:
                        st.markdown(f"#### {section_title}")
                        for k in keys:
                            if k in parsed_content:
                                st.write(f"**{k}:** {parsed_content[k]}")

                    # Build DOCX file
                    doc_file = build_deped_ilaw_docx(header_info, parsed_content)

                    st.download_button(
                        label="📄 Download Official DepEd ILAW Lesson Plan (.docx)",
                        data=doc_file,
                        file_name=f"DepEd_ILAW_Lesson_Plan_{topic.replace(' ', '_')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                    )

                except Exception as e:
                    st.error(f"Generation Error: {str(e)}")