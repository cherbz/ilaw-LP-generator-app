import json
import re
import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor
from google import genai
from google.genai import types
from pptx import Presentation
import streamlit as st

# Page Configuration
st.set_page_config(
    page_title="Binonz ILAW Lesson Plan Generator", 
    page_icon="📝", 
    layout="wide"
)

st.title("📝 Binonz ILAW Lesson Plan Generator")
st.caption(
    "Automated DepEd Order No. 003, s. 2026 Lesson Plan Generator with Embedded COT Indicators"
)

# -------------------------------------------------------------
# SIDEBAR CONFIGURATION
# -------------------------------------------------------------
st.sidebar.header("🔑 Configuration")

# Get API Key from user
gemini_key = st.sidebar.text_input("Enter Gemini API Key:", type="password")

st.sidebar.markdown("---")
st.sidebar.header("📌 Target COT Indicators")

# Choose Teacher Rank to filter corresponding indicators
teacher_rank = st.sidebar.selectbox(
    "Select Your Teacher Rank:",
    [
        "Teacher I - III (Proficient)", 
        "Teacher IV - VI (Highly Proficient Transition)", 
        "Master Teacher I - II (Highly Proficient)", 
        "Master Teacher III - V (Distinguished)"
    ]
)

# Define indicators based on rank selection
cot_database = {
    "Teacher I - III (Proficient)": [
        "1.1.2: Apply knowledge of content within and across curriculum teaching areas.",
        "1.4.2: Use a range of teaching strategies that enhance learner achievement in literacy and numeracy skills.",
        "1.5.2: Apply a range of teaching strategies to develop critical and creative thinking, as well as other higher-order thinking skills.",
        "2.3.2: Manage classroom structure to engage learners, individually or in groups, in meaningful exploration, discovery and hands-on activities.",
        "2.6.2: Manage learner behavior constructively by applying positive and non-violent discipline.",
        "3.1.2: Use differentiated, developmentally appropriate learning experiences to address learners' needs.",
        "4.1.2: Design, select, organize, and use diagnostic, formative, and summative assessment strategies.",
        "4.5.2: Select, develop, organize, and use appropriate teaching and learning resources, including ICT."
    ],
    "Teacher IV - VI (Highly Proficient Transition)": [
        "1.1.2: Apply high-level content knowledge to challenge learners cognitively.",
        "1.5.2: Lead colleagues in implementing teaching strategies to build critical and creative thinking.",
        "2.3.2: Refine classroom structures to maximize learning activities smoothly.",
        "3.1.2: Adapt and customize differentiated learning resources for diverse classroom conditions.",
        "5.1.2: Formulate dynamic diagnostic and feedback loops to support self-regulation."
    ],
    "Master Teacher I - II (Highly Proficient)": [
        "1.1.3: Model effective applications of content knowledge within and across curriculum teaching areas.",
        "1.4.3: Model and support colleagues in using teaching strategies that enhance literacy and numeracy.",
        "1.5.3: Model a range of teaching strategies to develop critical, creative, and higher-order thinking skills.",
        "2.3.3: Model effective classroom management designs that encourage collaborative responsibility.",
        "3.1.3: Model differentiated, developmentally appropriate experiences to address learners' needs and strengths.",
        "5.1.3: Model the design and constructive use of diagnostic, formative, and summative assessment strategies."
    ],
    "Master Teacher III - V (Distinguished)": [
        "1.1.4: Lead collaboration with colleagues to deepen content knowledge and pedagogical integration.",
        "1.5.4: Initiate systemic school-wide policies that nurture critical and creative thinking styles.",
        "3.1.4: Lead curriculum design innovations targeting marginalized, indigenous, and gifted learners.",
        "5.1.4: Lead institutional strategies for utilizing assessment data to inform school-wide instructional adjustments."
    ]
}

# Dynamic Checkboxes based on selected rank
available_indicators = cot_database[teacher_rank]
selected_indicators = []

st.sidebar.write(f"Select indicators to inject into the lesson:")
for ind in available_indicators:
    label_short = ind.split(":")[0] + "..."
    if st.sidebar.checkbox(ind, value=True, key=ind):
        selected_indicators.append(ind)

# -------------------------------------------------------------
# MAIN APP FORM
# -------------------------------------------------------------
col1, col2 = st.columns(2)

with col1:
    teacher_name = st.text_input("Teacher Name:", value="NORBERTO P. BINONDO JR.")
    learning_area = st.text_input("Learning Area / Subject:", value="Mathematics")
    grade_section = st.text_input("Grade Level & Section:", value="Grade 9 - Kindness")

with col2:
    lesson_name = st.text_input("Name of Lesson:", value="Graphing Linear Functions")
    num_sessions = st.text_input("No. of Sessions:", value="1")
    uploaded_ppt = st.file_uploader("Upload Lesson Presentation (.pptx)", type=["pptx"])

# Helper function to extract text from pptx
def extract_pptx_text(file):
    try:
        prs = Presentation(file)
        text_runs = []
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text.strip():
                    text_runs.append(shape.text.strip())
        return "\n".join(text_runs)
    except Exception as e:
        return f"Could not parse PPTX: {str(e)}"

# -------------------------------------------------------------
# WORD GENERATION HELPER FUNCTIONS
# -------------------------------------------------------------
def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=22, bottom=22, left=70, right=70):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table, color="C8D3E0", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="none"/>'
        f'<w:right w:val="none"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

def build_docx(data, teacher, rank, subject, grade, sessions, name_of_lesson):
    doc = docx.Document()
    
    # Page setup for 8.5 x 13 Long Bond Paper (Folio)
    for section in doc.sections:
        section.page_width = Inches(8.5)
        section.page_height = Inches(13.0)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    PRIMARY = RGBColor(27, 54, 93)     # #1B365D
    DARK_TEXT = RGBColor(34, 34, 34)   # #222222
    HIGHLIGHT_RED = RGBColor(180, 0, 0)

    # Header Banner Table
    banner_table = doc.add_table(rows=1, cols=1)
    banner_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = banner_table.cell(0, 0)
    set_cell_background(cell, "1B365D")
    set_cell_margins(cell, top=45, bottom=45, left=90, right=90)

    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(1)
    run = p.add_run("ILAW LESSON PLAN ON " + subject.upper())
    run.font.name = "Arial"
    run.font.size = Pt(12)
    run.font.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)

    p2 = cell.add_paragraph()
    p2.paragraph_format.space_after = Pt(0)
    r2 = p2.add_run(f"DepEd Order No. 003, s. 2026 (Annex A Format) | Target Stage: {rank}")
    r2.font.name = "Arial"
    r2.font.size = Pt(8.5)
    r2.font.color.rgb = RGBColor(224, 230, 237)

    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    def add_section_header(title_text):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        c = tbl.cell(0, 0)
        set_cell_background(c, "2C4D75")
        set_cell_margins(c, top=28, bottom=28, left=70, right=70)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(title_text)
        r.font.name = "Arial"
        r.font.size = Pt(9.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)
        doc.add_paragraph().paragraph_format.space_after = Pt(2)

    def append_highlight_cot(paragraph, text):
        run = paragraph.add_run(f"\n\n📌 [COT INDICATOR: {text}]")
        run.font.name = "Arial"
        run.font.size = Pt(8)
        run.font.bold = True
        run.font.color.rgb = HIGHLIGHT_RED
        rPr = run._r.get_or_add_rPr()
        highlight = OxmlElement('w:highlight')
        highlight.set(qn('w:val'), 'yellow')
        rPr.append(highlight)

    # Metadata Table
    metadata_fields = [
        ("Name of Lesson", name_of_lesson),
        ("Learning Area/s", subject),
        ("Designed by Teacher/s", f"{teacher} ({rank})"),
        ("Designed for Grade/Section", grade),
        ("No. of Sessions", sessions),
        ("Declaration of AI Use", "AI was utilized to format the plan structure matching DepEd standards.")
    ]

    meta_table = doc.add_table(rows=len(metadata_fields), cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(meta_table)

    for idx, (label, val) in enumerate(metadata_fields):
        c0 = meta_table.cell(idx, 0)
        c1 = meta_table.cell(idx, 1)
        set_cell_background(c0, "F0F4F8")
        set_cell_margins(c0)
        set_cell_margins(c1)
        c0.width = Inches(2.3)
        c1.width = Inches(5.2)
        
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(label)
        r0.font.name = "Arial"
        r0.font.size = Pt(8.5)
        r0.font.bold = True
        r0.font.color.rgb = PRIMARY
        
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(str(val))
        r1.font.name = "Arial"
        r1.font.size = Pt(8.5)
        r1.font.color.rgb = DARK_TEXT

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # Fill Sections from JSON Data
    for sec_num, (sec_title, fields) in enumerate(data.items(), start=1):
        add_section_header(f"{sec_num}. {sec_title.upper()}")
        
        sec_table = doc.add_table(rows=len(fields), cols=2)
        sec_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(sec_table)
        
        for idx, (field_label, field_data) in enumerate(fields.items()):
            c0 = sec_table.cell(idx, 0)
            c1 = sec_table.cell(idx, 1)
            set_cell_background(c0, "F0F4F8")
            set_cell_margins(c0)
            set_cell_margins(c1)
            c0.width = Inches(2.3)
            c1.width = Inches(5.2)
            
            p0 = c0.paragraphs[0]
            p0.paragraph_format.space_after = Pt(0)
            r0 = p0.add_run(field_label)
            r0.font.name = "Arial"
            r0.font.size = Pt(8.5)
            r0.font.bold = True
            r0.font.color.rgb = PRIMARY
            
            p1 = c1.paragraphs[0]
            p1.paragraph_format.space_after = Pt(0)
            
            # Extract content text and optional COT indicator
            content_text = field_data
            cot_found = None
            if isinstance(field_data, dict):
                content_text = field_data.get("content", "")
                cot_found = field_data.get("cot", None)
            
            r1 = p1.add_run(str(content_text))
            r1.font.name = "Arial"
            r1.font.size = Pt(8.5)
            r1.font.color.rgb = DARK_TEXT
            
            if cot_found:
                append_highlight_cot(p1, cot_found)
                
        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    return doc

# -------------------------------------------------------------
# RUNNING GENERATION
# -------------------------------------------------------------
if st.button("🚀 Generate ILAW Lesson Plan Document"):
    if not gemini_key:
        st.error("Please enter a valid Gemini API Key in the sidebar.")
    else:
        with st.spinner("Analyzing PPTX & generating standardized Lesson Plan..."):
            # Extract ppt context if available
            context_text = ""
            if uploaded_ppt is not None:
                context_text = extract_pptx_text(uploaded_ppt)
                st.info("Uploaded slides processed as lesson reference context.")
            
            # Construct a clear prompt demanding valid JSON structure matching Annex A sections
            prompt = f"""
            Generate a standardized DepEd Order No. 003, s. 2026 Lesson Plan (ILAW Framework, Annex A format) in JSON format.
            Subject: {learning_area}
            Lesson Name: {lesson_name}
            Grade/Section: {grade_section}
            Teacher Rank / Target Audience: {teacher_rank}
            Selected COT Indicators: {selected_indicators}
            
            Instructional slide raw context (if any):
            \"\"\"{context_text}\"\"\"

            Respond ONLY with a valid JSON document (do not include markdown syntax wrappers).
            Structure the JSON object strictly as follows:
            {{
                "Intentions": {{
                    "Learning Competency": {{"content": "..."}},
                    "Learning Objectives": {{"content": "..."}},
                    "Learner Context & Diagnostic": {{"content": "...", "cot": "{selected_indicators[0] if len(selected_indicators)>0 else ''}"}}
                }},
                "Learning Experience": {{
                    "Pre-Lesson Activities": {{"content": "...", "cot": "{selected_indicators[1] if len(selected_indicators)>1 else ''}"}},
                    "Instructional Flow & Direct Modeling": {{"content": "...", "cot": "{selected_indicators[2] if len(selected_indicators)>2 else ''}"}},
                    "Collaborative Activities": {{"content": "...", "cot": "{selected_indicators[3] if len(selected_indicators)>3 else ''}"}},
                    "Synthesis & Review": {{"content": "..."}}
                }},
                "Assessment": {{
                    "Formative Assessment Task": {{"content": "...", "cot": "{selected_indicators[4] if len(selected_indicators)>4 else ''}"}}
                }},
                "Ways Forward": {{
                    "Extended Learning (Remediation/Homework)": {{"content": "..."}},
                    "Teacher Self-Reflection Note": {{"content": "..."}}
                }}
            }}
            """

            try:
                # Initialize Google Gemini Client with the newer gemini-3.5-flash model
                client = genai.Client(api_key=gemini_key)
                
                response = client.models.generate_content(
                    model='gemini-3.5-flash',
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json"
                    )
                )
                
                # Parse output
                clean_json_str = re.sub(r"^```json|```$", "", response.text, flags=re.MULTILINE).strip()
                parsed_plan = json.loads(clean_json_str)
                
                # Build docx file
                out_doc = build_docx(
                    parsed_plan, 
                    teacher_name, 
                    teacher_rank, 
                    learning_area, 
                    grade_section, 
                    num_sessions, 
                    lesson_name
                )
                
                # Setup Download Interface
                import io
                bio = io.BytesIO()
                out_doc.save(bio)
                bio.seek(0)
                
                st.success("🎉 Lesson Plan generated successfully!")
                st.download_button(
                    label="💾 Download Lesson Plan Document (.docx)",
                    data=bio,
                    file_name=f"ILAW_Lesson_Plan_{lesson_name.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
            except Exception as e:
                st.error(f"Error processing document: {str(e)}")