import os
import streamlit as st
from google import genai
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="Binonz ILAW Lesson Plan Generator",
    page_icon="📝",
    layout="wide"
)

# --- HELPER FUNCTIONS FOR DOCX TABLE FORMATTING ---
def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

def add_styled_row(table, title, content):
    row = table.add_row()
    cell_title = row.cells[0]
    cell_content = row.cells[1]
    
    cell_title.width = Inches(2.0)
    cell_content.width = Inches(4.5)
    
    set_cell_background(cell_title, "F0F2F6")
    
    p_title = cell_title.paragraphs[0]
    run_title = p_title.add_run(title)
    run_title.bold = True
    run_title.font.name = 'Calibri'
    run_title.font.size = Pt(10)
    
    p_content = cell_content.paragraphs[0]
    run_content = p_content.add_run(content)
    run_content.font.name = 'Calibri'
    run_content.font.size = Pt(10)

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
            [
                "Teacher I - III (Proficient)",
                "Teacher IV - VI (Highly Proficient)",
                "Master Teacher I - II (Highly Proficient)",
                "Master Teacher III - V (Distinguished)"
            ]
        )
        
        st.markdown("**Active COT Indicators for Generation:**")
        cot_1 = st.checkbox("1.4.2: Use a range of teaching strategies that enhance learner achievement in literacy and numeracy skills.", value=True)
        cot_2 = st.checkbox("1.5.2: Apply a range of teaching strategies to develop critical and creative thinking, as well as other higher-order thinking skills.", value=True)
        cot_3 = st.checkbox("2.3.2: Manage classroom structure to engage learners in meaningful exploration, discovery and hands-on activities.", value=True)
        cot_4 = st.checkbox("3.1.2: Use differentiated, developmentally appropriate learning experiences to address learners' gender, needs, strengths, interests and experiences.", value=True)

    # --- MAIN UI LAYOUT ---
    st.title("📝 Binonz ILAW Lesson Plan Generator")
    st.caption("Automated DepEd Order No. 003, s. 2026 Lesson Plan Generator with Embedded COT Indicators")

    if not api_key:
        st.warning("⚠️ Please provide a Gemini API Key in the sidebar to generate lesson plans.")
        st.stop()

    client = genai.Client(api_key=api_key)

    # Input Fields Layout
    col1, col2 = st.columns(2)
    with col1:
        teacher_name = st.text_input("Teacher Name:", "NORBERTO P. BINONDO JR.")
        subject = st.text_input("Learning Area / Subject:", "Mathematics")
        grade_section = st.text_input("Grade Level & Section:", "Grade 9 - Kindness")

    with col2:
        lesson_name = st.text_input("Name of Lesson:", "Graphing Linear Functions")
        sessions = st.number_input("No. of Sessions:", min_value=1, max_value=10, value=1)
        uploaded_pptx = st.file_uploader("Upload Lesson Presentation (.pptx)", type=["pptx"])

    st.divider()

    # --- GENERATION LOGIC ---
    if st.button("🪄 Generate ILAW Lesson Plan Document", type="primary", use_container_width=True):
        with st.spinner("Generating DepEd Order No. 003, s. 2026 Annex A Lesson Plan..."):
            try:
                cot_list = []
                if cot_1: cot_list.append("1.4.2: Use a range of teaching strategies that enhance learner achievement in literacy and numeracy skills.")
                if cot_2: cot_list.append("1.5.2: Apply a range of teaching strategies to develop critical and creative thinking, as well as other higher-order thinking skills.")
                if cot_3: cot_list.append("2.3.2: Manage classroom structure to engage learners in meaningful exploration, discovery and hands-on activities.")
                if cot_4: cot_list.append("3.1.2: Use differentiated, developmentally appropriate learning experiences to address learners' gender, needs, strengths, interests and experiences.")

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

                TARGET COT INDICATORS TO EMBED IN TEXT (Include '📌 [COT INDICATOR: ...]' tags explicitly at the end of relevant sections):
                {chr(10).join(['- ' + c for c in cot_list])}

                PROVIDE OUTPUT EXACTLY IN THIS JSON FORMAT (no additional markdown outside JSON):
                {{
                  "references": "DepEd Curriculum Guide & Presentation Slides",
                  "competency": "Graphs a linear function and values its real-life applications (domain, range, intercepts, and slope).",
                  "objectives": "1. Graph linear equations by constructing a table of values and plotting points on the Cartesian plane.\\n2. Interpret the slope and y-intercept within real-life contexts.\\n3. Show cooperative engagement during group activities. 📌 [COT INDICATOR: 1.4.2: Use a range of teaching strategies that enhance learner achievement in literacy and numeracy skills.]",
                  "learner_context": "Differentiated Experiences... 📌 [COT INDICATOR: 3.1.2: Use differentiated, developmentally appropriate learning experiences to address learners' gender, needs, strengths, interests and experiences.]",
                  "pre_lesson": "Conduct diagnostic review... 📌 [COT INDICATOR: 1.4.2: Use a range of teaching strategies that enhance learner achievement in literacy and numeracy skills.]",
                  "instructional_flow": "Explain step-by-step substitution... 📌 [COT INDICATOR: 1.1.2: Apply knowledge of content within and across curriculum teaching areas.]",
                  "group_activity": "Assign group tasks with varied complexities... 📌 [COT INDICATOR: 2.3.2: Manage classroom structure to engage learners in meaningful exploration, discovery and hands-on activities.]",
                  "synthesis": "Facilitate structured debrief... 📌 [COT INDICATOR: 1.5.2: Apply a range of teaching strategies to develop critical and creative thinking, as well as other higher-order thinking skills.]",
                  "integration": "Incorporate cross-curricular linkages... 📌 [COT INDICATOR: 1.1.2: Apply knowledge of content within and across curriculum teaching areas.]",
                  "assessment": "Formative evaluation on paper... 📌 [COT INDICATOR: 4.1.2: Design, select, organize, and use diagnostic, formative, and summative assessment strategies.]",
                  "extended_learning": "Assign infant growth challenge...",
                  "teacher_reflections": "Reflect on student mastery..."
                }}
                """

                response = client.models.generate_content(
                    model="gemini-2.5-flash",
                    contents=prompt
                )
                
                # Parse JSON Response
                import json
                clean_json = response.text.replace("```json", "").replace("```", "").strip()
                data = json.loads(clean_json)

                # --- DISPLAY PREVIEW ---
                st.subheader("📄 Generated Lesson Plan Preview")
                st.markdown(f"**Lesson:** {lesson_name} | **Teacher:** {teacher_name}")
                st.markdown(f"**Competency:** {data['competency']}")
                st.markdown("---")
                st.markdown(f"**Objectives:**\n{data['objectives']}")
                st.markdown(f"**Instructional Flow:**\n{data['instructional_flow']}")
                st.markdown(f"**Collaborative Group Activity:**\n{data['group_activity']}")

                # --- GENERATE WORD (.DOCX) TEMPLATE MATCHING ANNEX A ---
                doc = Document()
                
                # Header Title
                p_head = doc.add_paragraph()
                p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_head1 = p_head.add_run(f"ILAW LESSON PLAN ON {subject.upper()}\n")
                r_head1.bold = True
                r_head1.font.size = Pt(14)
                r_head2 = p_head.add_run("DepEd Order No. 003, s. 2026 (Annex A Template) | COT Indicators Embedded\n")
                r_head2.font.size = Pt(9)
                r_head2.font.italic = True

                # Primary Table
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
                add_styled_row(table, "1. INTENTIONS", "")
                add_styled_row(table, "Learning Competency", data['competency'])
                add_styled_row(table, "Learning Objectives", data['objectives'])
                add_styled_row(table, "Learner Context", data['learner_context'])

                # Section 2: LEARNING EXPERIENCE
                add_styled_row(table, "2. LEARNING EXPERIENCE", "")
                add_styled_row(table, "Pre-Lesson (Getting Ready)", data['pre_lesson'])
                add_styled_row(table, "Instructional Flow & Direct Modeling", data['instructional_flow'])
                add_styled_row(table, "Collaborative Group Activity", data['group_activity'])
                add_styled_row(table, "Synthesis & Resources", data['synthesis'])
                add_styled_row(table, "Opportunities for Integration", data['integration'])

                # Section 3: ASSESSMENT
                add_styled_row(table, "3. ASSESSMENT", "")
                add_styled_row(table, "Formative Assessment (Individual Evaluation)", data['assessment'])

                # Section 4: WAYS FORWARD
                add_styled_row(table, "4. WAYS FORWARD", "")
                add_styled_row(table, "Extended Learning Opportunities", data['extended_learning'])
                add_styled_row(table, "Teacher Reflections", data['teacher_reflections'])

                # Save Document
                doc_path = "ILAW_Lesson_Plan_Annex_A.docx"
                doc.save(doc_path)

                with open(doc_path, "rb") as file:
                    st.download_button(
                        label="📥 Download DepEd Order No. 003 Annex A Word Document (.docx)",
                        data=file,
                        file_name=f"ILAW_Lesson_Plan_{lesson_name.replace(' ', '_')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )

            except Exception as e:
                st.error(f"Error generating lesson plan: {str(e)}")