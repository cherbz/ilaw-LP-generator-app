import os
import json
import re
import streamlit as st
from google import genai
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="Binonz ILAW Lesson Plan Generator",
    page_icon="📝",
    layout="wide"
)

# --- HELPER FUNCTIONS FOR XML TABLE & CELL FORMATTING ---
def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_formatted_text_with_highlights(paragraph, text):
    """
    Parses text for '📌 [COT INDICATOR: ...]' patterns and renders them 
    with Yellow Highlighting and Bold styling matching the target DOCX.
    """
    pattern = r'(📌\s*\[COT INDICATOR:[^\]]+\])'
    parts = re.split(pattern, text)
    
    for part in parts:
        if not part:
            continue
        if part.startswith('📌') or 'COT INDICATOR:' in part:
            run = paragraph.add_run(f" {part.strip()} ")
            run.font.name = 'Calibri'
            run.font.size = Pt(9.5)
            run.font.bold = True
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        else:
            run = paragraph.add_run(part)
            run.font.name = 'Calibri'
            run.font.size = Pt(10)

def add_table_header_section(table, title):
    """Adds a full-width header row for main sections like 1. INTENTIONS."""
    row = table.add_row()
    a = row.cells[0]
    b = row.cells[1]
    a.merge(b)
    
    set_cell_background(a, "E0E0E0")
    set_cell_margins(a, top=120, bottom=120, left=150, right=150)
    
    p = a.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(title)
    run.bold = True
    run.font.name = 'Calibri'
    run.font.size = Pt(10.5)

def add_styled_row(table, title, content):
    row = table.add_row()
    cell_title = row.cells[0]
    cell_content = row.cells[1]
    
    cell_title.width = Inches(2.2)
    cell_content.width = Inches(4.3)
    
    set_cell_background(cell_title, "F4F6F8")
    set_cell_margins(cell_title, top=100, bottom=100, left=150, right=150)
    set_cell_margins(cell_content, top=100, bottom=100, left=150, right=150)
    
    p_title = cell_title.paragraphs[0]
    run_title = p_title.add_run(title)
    run_title.bold = True
    run_title.font.name = 'Calibri'
    run_title.font.size = Pt(10)
    
    p_content = cell_content.paragraphs[0]
    add_formatted_text_with_highlights(p_content, content)

# --- PASSWORD PROTECTION FUNCTION ---
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if st.session_state["password_correct"]:
        return True

    st.title("🔒 Access Required")
    st.write("Please enter the access code to use the Binonz ILAW Lesson Plan Generator.")
    
    user_input = st.text_input("Enter Access Code:", type="password")

    if st.button("Unlock App"):
        if user_input == "depedilaw123$$":
            st.session_state["password_correct"] = True
            st.success("Access granted!")
            st.rerun()
        else:
            st.error("❌ Incorrect access code. Please try again.")

    return False

# --- MAIN APPLICATION CODE ---
if check_password():
    # --- COT INDICATORS DICTIONARY BY TEACHER RANK ---
    COT_INDICATORS_BY_RANK = {
        "Teacher I - III (Proficient)": [
            "1.4.2: Use a range of teaching strategies that enhance learner achievement in literacy and numeracy skills.",
            "1.5.2: Apply a range of teaching strategies to develop critical and creative thinking, as well as other higher-order thinking skills.",
            "2.3.2: Manage classroom structure to engage learners in meaningful exploration, discovery and hands-on activities.",
            "3.1.2: Use differentiated, developmentally appropriate learning experiences to address learners' gender, needs, strengths, interests and experiences."
        ],
        "Teacher IV - VI (Highly Proficient)": [
            "1.4.3: Model and support colleagues in using a range of teaching strategies that enhance learner achievement in literacy and numeracy skills.",
            "1.5.3: Develop and apply effective teaching strategies to promote critical and creative thinking, as well as other higher-order thinking skills.",
            "2.3.3: Work with colleagues to share differentiated, developmentally appropriate opportunities to address learners' needs.",
            "3.1.3: Design, adapt and implement teaching strategies that are responsive to learners with disabilities, giftedness and talents."
        ],
        "Master Teacher I - II (Highly Proficient)": [
            "1.4.3: Lead colleagues in evaluating and refining teaching strategies that enhance learner achievement in literacy and numeracy skills.",
            "1.5.3: Model exemplary teaching strategies to develop higher-order thinking skills among learners.",
            "2.3.3: Support colleagues in managing structured classroom environments for discovery learning.",
            "3.1.3: Advise and guide colleagues on differentiated instruction techniques tailored to diverse learner groups."
        ],
        "Master Teacher III - V (Distinguished)": [
            "1.4.4: Lead institutional initiatives that improve literacy and numeracy achievement across grade levels.",
            "1.5.4: Lead in the design and evaluation of instructional models that foster critical, creative, and transformative thinking.",
            "2.3.4: Establish school-wide standards for conducive and structured physical and virtual learning environments.",
            "3.1.4: Provide strategic leadership in creating inclusive and responsive learning policies and frameworks."
        ]
    }

    # --- SIDEBAR CONFIGURATION ---
    with st.sidebar:
        st.header("🔍 Configuration")
        api_key = st.secrets.get("GEMINI_API_KEY", os.getenv("GEMINI_API_KEY"))
        if not api_key:
            api_key = st.text_input("Enter Gemini API Key:", type="password")
        
        st.divider()
        st.header("📌 Target COT Indicators")
        
        teacher_rank = st.selectbox(
            "Select Your Teacher Rank:",
            list(COT_INDICATORS_BY_RANK.keys())
        )
        
        st.markdown("**Active COT Indicators for Generation:**")
        
        selected_cot_indicators = []
        available_indicators = COT_INDICATORS_BY_RANK[teacher_rank]
        
        for idx, indicator in enumerate(available_indicators):
            if st.checkbox(indicator, value=True, key=f"cot_{teacher_rank}_{idx}"):
                selected_cot_indicators.append(indicator)

    # --- MAIN UI LAYOUT ---
    st.title("📝 Binonz ILAW Lesson Plan Generator")
    st.caption("Automated DepEd Order No. 003, s. 2026 Lesson Plan Generator with Highlighted Embedded COT Indicators")

    if not api_key:
        st.warning("⚠️ Please provide a Gemini API Key in the sidebar to generate lesson plans.")
        st.stop()

    client = genai.Client(api_key=api_key)

    # General Information Inputs
    st.subheader("📋 General Information")
    col1, col2 = st.columns(2)
    with col1:
        teacher_name = st.text_input("Teacher Name:", "NORBERTO P. BINONDO JR.")
        subject = st.text_input("Learning Area / Subject:", "Mathematics")
        grade_section = st.text_input("Grade Level & Section:", "Grade 9 - Kindness")

    with col2:
        lesson_name = st.text_input("Name of Lesson:", "Graphing Linear Functions")
        sessions = st.number_input("No. of Sessions:", min_value=1, max_value=10, value=1)

    st.divider()

    # Competency Input Tabs
    st.subheader("🎯 DepEd Curriculum & Lesson Inputs")
    tab1, tab2 = st.tabs(["📝 Manual Input / Copy-Paste Competencies", "📚 Suggested Template Prompts"])

    with tab1:
        learning_competency = st.text_area(
            "Enter / Copy-Paste DepEd Learning Competency:",
            value="Graphs a linear function and values its real-life applications (domain, range, intercepts, and slope).",
            height=100
        )
        specific_objectives = st.text_area(
            "Enter Specific Objectives or Focus Points (Optional):",
            value="1. Graph linear equations by constructing a table of values and plotting points on the Cartesian plane.\n2. Interpret the slope and y-intercept within real-life contexts.\n3. Show cooperative engagement during differentiated group activities.",
            height=120
        )

    with tab2:
        st.info("💡 You can select or reference pre-formulated DepEd competencies for faster lesson plan generation.")
        st.markdown("""
        **Example Mathematics Competency:**
        > *Graphs a linear function and values its real-life applications (domain, range, intercepts, and slope).*

        **Example TLE / Computer Systems Servicing Competency:**
        > *Install and configure computer systems and networks in accordance with industry standards.*
        """)

    st.divider()

    # --- GENERATION LOGIC ---
    if st.button("🪄 Generate ILAW Lesson Plan Document", type="primary", use_container_width=True):
        with st.spinner("Generating DepEd Order No. 003, s. 2026 Annex A Lesson Plan with Highlighting..."):
            try:
                prompt = f"""
                You are a DepEd Master Teacher in the Philippines.
                Generate a complete DepEd Order No. 003, s. 2026 (Annex A Template) Lesson Plan.

                TEACHER DETAILS:
                - Name: {teacher_name}
                - Subject: {subject}
                - Grade & Section: {grade_section}
                - Lesson Name: {lesson_name}
                - No. of Sessions: {sessions}
                - Teacher Rank: {teacher_rank}

                DEPED LEARNING COMPETENCY:
                {learning_competency}

                SPECIFIC OBJECTIVES / FOCUS POINTS:
                {specific_objectives}

                TARGET COT INDICATORS TO EMBED IN TEXT (Include '📌 [COT INDICATOR: ...]' tags explicitly at the end of relevant sections):
                {chr(10).join(['- ' + c for c in selected_cot_indicators])}

                PROVIDE OUTPUT EXACTLY IN THIS JSON FORMAT (no additional markdown outside JSON):
                {{
                  "references": "DepEd K to 12 Curriculum Guide; Mathematics 9 Teacher's Guide; Learner's Materials for Grade 9 Mathematics",
                  "competency": "{learning_competency}",
                  "objectives": "1. Graph linear equations by constructing a table of values and plotting points on the Cartesian plane.\\n2. Interpret the slope and y-intercept within real-life contexts, such as airline baggage fees and infant weight monitoring.\\n3. Show cooperative engagement during differentiated group activities to complete mathematical challenges. 📌 [COT INDICATOR: 1.4.2: Use a range of teaching strategies that enhance learner achievement in literacy and numeracy skills.]",
                  "learner_context": "The {grade_section} class is a mixed-ability group of learners with varied mathematical inclinations. Visual and kinesthetic learners benefit from coordinate plotting exercises, while logical-mathematical thinkers appreciate contextual applications. Differentiated experiences help bridge the achievement gap. 📌 [COT INDICATOR: 3.1.2: Use differentiated, developmentally appropriate learning experiences to address learners' gender, needs, strengths, interests and experiences.]",
                  "pre_lesson": "Conduct a brief 'Coordinate Hunt' diagnostic game to test plotting concepts. Establish non-verbal signals and positive feedback systems to ensure a safe, organized environment. Review key numeracy concepts before introducing linear frameworks. 📌 [COT INDICATOR: 1.4.2: Use a range of teaching strategies that enhance learner achievement in literacy and numeracy skills.]",
                  "instructional_flow": "Explain how to graph the linear equation y = 3x - 4 and y = -0.5x - 2. Use color-coded ICT-assisted plotting animations to illustrate step-by-step substitution of values. Highlight cross-curricular science and health integration. 📌 [COT INDICATOR: 1.5.2: Apply a range of teaching strategies to develop critical and creative thinking, as well as other higher-order thinking skills.]",
                  "group_activity": "Assign group tasks with varied complexities: Group 1 and 3 use basic integers; Group 2 (Airline luggage fee) and Group 4 (Infant weight tracking) apply equations to real-life issues. Walk around to moderate discussions and manage cooperative structures. 📌 [COT INDICATOR: 2.3.2: Manage classroom structure to engage learners in meaningful exploration, discovery and hands-on activities.]",
                  "synthesis": "Facilitate a structured debrief centered on three essential questions regarding the steps in graphing linear functions and real-life mathematical models. 📌 [COT INDICATOR: 1.5.2: Apply a range of teaching strategies to develop critical and creative thinking, as well as other higher-order thinking skills.]",
                  "integration": "Incorporate cross-curricular linkages with Health/Pediatrics (monitoring growth) and Business/Economics (calculating baggage fee rates) to demonstrate practical uses of math.",
                  "assessment": "Conduct a formative 'Graph Checkpoint' evaluation on paper. Students are given an equation to find coordinates, construct a table of values, and trace the line.",
                  "extended_learning": "Assign the 'Infant Growth Challenge': Solve y = 0.9x + 3 for x = 0, 3, 6, and 12. Plot coordinates on graph paper and identify the y-intercept.",
                  "teacher_reflections": "Reflect on student responses to decimal coefficients and note the effectiveness of peer coaching during group work."
                }}
                """

                response = client.models.generate_content(
                    model="gemini-2.5-flash",
                    contents=prompt
                )
                
                clean_json = response.text.replace("```json", "").replace("```", "").strip()
                data = json.loads(clean_json)

                # --- GENERATE WORD (.DOCX) MATCHING TEMPLATE EXACTLY ---
                doc = Document()

                # Document Header Title
                p_head = doc.add_paragraph()
                p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_head1 = p_head.add_run(f"ILAW LESSON PLAN ON {subject.upper()}\n")
                r_head1.bold = True
                r_head1.font.name = 'Calibri'
                r_head1.font.size = Pt(13)
                
                r_head2 = p_head.add_run("DepEd Order No. 003, s. 2026 (Annex A Template) | COT Indicators Embedded\n")
                r_head2.font.italic = True
                r_head2.font.name = 'Calibri'
                r_head2.font.size = Pt(9.5)

                # Primary Table Structure
                table = doc.add_table(rows=0, cols=2)
                table.style = 'Table Grid'

                add_styled_row(table, "Name of Lesson", lesson_name)
                add_styled_row(table, "Learning Area/s", subject)
                add_styled_row(table, "Designed by Teacher/s", teacher_name)
                add_styled_row(table, "Grade Level & Section", grade_section)
                add_styled_row(table, "No. of Sessions", str(sessions))
                add_styled_row(table, "References", data['references'])
                add_styled_row(table, "Declaration of AI Use", "AI was utilized to structure content into DepEd Order No. 003, s. 2026 Annex A template & align COT indicators.")

                # Section 1: INTENTIONS
                add_table_header_section(table, "1. INTENTIONS")
                add_styled_row(table, "Learning Competency", data['competency'])
                add_styled_row(table, "Learning Objectives", data['objectives'])
                add_styled_row(table, "Learner Context", data['learner_context'])

                # Section 2: LEARNING EXPERIENCE
                add_table_header_section(table, "2. LEARNING EXPERIENCE")
                add_styled_row(table, "Pre-Lesson\n(Getting Ready)", data['pre_lesson'])
                add_styled_row(table, "Instructional Flow &\nDirect Modeling", data['instructional_flow'])
                add_styled_row(table, "Collaborative Group\nActivity", data['group_activity'])
                add_styled_row(table, "Synthesis & Resources", data['synthesis'])
                add_styled_row(table, "Opportunities for\nIntegration", data['integration'])

                # Section 3: ASSESSMENT
                add_table_header_section(table, "3. ASSESSMENT")
                add_styled_row(table, "Formative Assessment\n(Individual Evaluation)", data['assessment'])

                # Section 4: WAYS FORWARD
                add_table_header_section(table, "4. WAYS FORWARD")
                add_styled_row(table, "Extended Learning\nOpportunities", data['extended_learning'])
                add_styled_row(table, "Teacher Reflections", data['teacher_reflections'])

                # Save Document
                doc_path = "ILAW_Lesson_Plan_Annex_A.docx"
                doc.save(doc_path)

                # UI Preview
                st.subheader("📄 Generated Lesson Plan Preview")
                st.info("The generated Word document contains yellow text highlights for embedded COT indicators matching Annex A requirements.")
                st.markdown(f"**Lesson:** {lesson_name} | **Teacher:** {teacher_name}")
                st.markdown(f"**Competency:** {data['competency']}")

                with open(doc_path, "rb") as file:
                    st.download_button(
                        label="📥 Download Highlighted Annex A Word Document (.docx)",
                        data=file,
                        file_name=f"ILAW_Lesson_Plan_{lesson_name.replace(' ', '_')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )

            except Exception as e:
                st.error(f"Error generating lesson plan document: {str(e)}")