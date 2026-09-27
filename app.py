import os
import streamlit as st
from google import genai
from docx import Document

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="ILAW Lesson Plan Generator",
    page_icon="📚",
    layout="wide"
)

# --- PASSWORD PROTECTION FUNCTION ---
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if st.session_state["password_correct"]:
        return True

    st.title("🔒 Access Required")
    st.write("Please enter the access code to use the ILAW Lesson Plan Generator.")
    
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
    st.title("📚 ILAW Lesson Plan Generator")
    st.write("Generate DepEd-aligned Daily Lesson Logs and Lesson Plans instantly.")

    # Configure Gemini API Key
    api_key = st.secrets.get("GEMINI_API_KEY", os.getenv("GEMINI_API_KEY"))
    
    if not api_key:
        api_key = st.text_input("Enter your Google Gemini API Key:", type="password")
        
    if not api_key:
        st.info("💡 Please provide a Gemini API Key to proceed.")
        st.stop()

    # Initialize Gemini Client
    client = genai.Client(api_key=api_key)

    # Form inputs
    col1, col2 = st.columns(2)
    with col1:
        grade_level = st.selectbox("Grade Level", ["Grade 7", "Grade 8", "Grade 9", "Grade 10", "Grade 11", "Grade 12"])
        subject = st.text_input("Subject", "Mathematics")
        topic = st.text_input("Topic", "Quadrilaterals")

    with col2:
        quarter = st.selectbox("Quarter", ["Quarter 1", "Quarter 2", "Quarter 3", "Quarter 4"])
        duration = st.text_input("Duration", "45 minutes")
        learning_competency = st.text_area("Learning Competency", "Solves problems involving quadrilaterals.")

    # Generation Button
    if st.button("Generate Lesson Plan", type="primary"):
        with st.spinner("Generating DepEd ILAW Lesson Plan..."):
            try:
                prompt = f"""
                You are an expert DepEd Public School Master Teacher in the Philippines.
                Create a detailed DepEd Daily Lesson Log (DLL) / Lesson Plan in the ILAW format for:
                - Grade Level: {grade_level}
                - Subject: {subject}
                - Topic: {topic}
                - Quarter: {quarter}
                - Duration: {duration}
                - Learning Competency: {learning_competency}

                Format the lesson plan clearly with sections:
                I. OBJECTIVES (Content Standards, Performance Standards, Learning Competencies)
                II. CONTENT
                III. LEARNING RESOURCES
                IV. PROCEDURES (Introductory Activity, Activity, Analysis, Abstraction, Application, Assessment, Assignment)
                V. REMARKS & REFLECTION
                """

                # Call model using modern SDK
                response = client.models.generate_content(
                    model="gemini-2.5-flash",
                    contents=prompt
                )
                generated_text = response.text

                st.markdown("### Generated Lesson Plan")
                st.markdown(generated_text)

                # Generate Word Document (.docx)
                doc = Document()
                doc.add_heading(f"ILAW Lesson Plan: {topic}", 0)
                doc.add_paragraph(f"Grade Level: {grade_level} | Subject: {subject} | Quarter: {quarter}")
                doc.add_paragraph("=" * 50)
                
                for paragraph in generated_text.split("\n\n"):
                    doc.add_paragraph(paragraph)
                
                doc_path = "Generated_ILAW_Lesson_Plan.docx"
                doc.save(doc_path)

                with open(doc_path, "rb") as file:
                    st.download_button(
                        label="📥 Download Word Document (.docx)",
                        data=file,
                        file_name=f"ILAW_Lesson_Plan_{topic.replace(' ', '_')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

            except Exception as e:
                st.error(f"Error generating lesson plan: {str(e)}")