import io
import docx
import google.generativeai as genai
import streamlit as st

# ==============================================================================
# 1. COT DATA STRUCTURES & CAREER STAGE MAPPING
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
        (
            "1.1.2",
            "Apply knowledge of content within and across curriculum teaching areas",
        ),
        (
            "1.4.2",
            "Use a range of teaching strategies that enhance learner achievement in literacy and numeracy skills",
        ),
        (
            "1.5.2",
            "Apply a range of teaching strategies to develop critical and creative thinking, as well as other higher-order thinking skills",
        ),
        (
            "2.3.2",
            "Manage classroom structure to engage learners, individually or in groups, in meaningful exploration, discovery and hands-on activities within a range of physical learning environments",
        ),
        (
            "2.6.2",
            "Manage learner behavior constructively by applying positive and non-violent discipline to ensure learning-focused environments",
        ),
        (
            "3.1.2",
            "Use differentiated, developmentally appropriate learning experiences to address learners' gender, needs, strengths, interests and experiences",
        ),
        (
            "4.1.2",
            "Plan, manage and implement developmentally sequenced teaching and learning process to meet curriculum requirements and varied teaching contexts",
        ),
        (
            "4.5.2",
            "Select, develop, organize and use appropriate teaching and learning resources, including ICT, to address learning goals",
        ),
        (
            "5.1.2",
            "Design, select, organize and use diagnostic, formative and summative assessment strategies consistent with curriculum requirements",
        ),
    ],
    "2026-2027": [
        (
            "1.1.2",
            "Apply knowledge of content within and across curriculum teaching areas",
        ),
        (
            "1.4.2",
            "Use a range of teaching strategies that enhance learner achievement in literacy and numeracy skills",
        ),
        (
            "1.5.2",
            "Apply a range of teaching strategies to develop critical and creative thinking, as well as other higher-order thinking skills",
        ),
        (
            "1.6.2",
            "Display proficient use of Mother Tongue, Filipino and English to facilitate teaching and learning",
        ),
        (
            "2.1.2",
            "Establish safe and secure learning environments to enhance learning through the consistent implementation of policies, guidelines and procedures",
        ),
        (
            "2.2.2",
            "Maintain learning environments that promote fairness, respect and care to encourage learning",
        ),
        (
            "3.2.2",
            "Establish a learner-centered culture by using teaching strategies that respond to learners' linguistic, cultural, socio-economic and religious backgrounds",
        ),
        (
            "3.5.2",
            "Adapt and use culturally appropriate teaching strategies to address the needs of learners from indigenous groups",
        ),
        (
            "5.3.2",
            "Use strategies for providing timely, accurate and constructive feedback to improve learner performance",
        ),
    ],
    "2027-2028": [
        (
            "1.1.2",
            "Apply knowledge of content within and across curriculum teaching areas",
        ),
        (
            "1.4.2",
            "Use a range of teaching strategies that enhance learner achievement in literacy and numeracy skills",
        ),
        (
            "1.3.2",
            "Ensure the positive use of ICT to facilitate the teaching and learning process",
        ),
        (
            "1.7.2",
            "Use effective verbal and non-verbal classroom communication strategies to support learner understanding, participation, engagement and achievement",
        ),
        (
            "2.4.2",
            "Maintain supportive learning environments that nurture and inspire learners to participate, cooperate and collaborate in continued learning",
        ),
        (
            "2.5.2",
            "Apply a range of successful strategies that maintain learning environments that motivate learners to work productively by assuming responsibility for their own learning",
        ),
        (
            "3.3.2",
            "Design, adapt and implement teaching strategies that are responsive to learners with disabilities, giftedness and talents",
        ),
        (
            "3.4.2",
            "Plan and deliver teaching strategies that are responsive to the special educational needs of learners in difficult circumstances",
        ),
    ],
}

# ==============================================================================
# 2. STREAMLIT APP LAYOUT & CONFIGURATION
# ==============================================================================

st.set_page_config(
    page_title="Binonz ILAW Lesson Plan Generator", page_icon="📝", layout="wide"
)

st.title("💡 Binonz ILAW Lesson Plan Generator")
st.caption(
    "Generate DepEd-aligned Daily Lesson Logs (DLL) and Lesson Plans with automated COT indicator integration."
)

# ------------------------------------------------------------------------------
# SIDEBAR: CREDENTIALS & TEACHER PROFILE
# ------------------------------------------------------------------------------
with st.sidebar:
    st.header("🔑 Authentication")
    license_key = st.text_input("License Key", type="password")
    user_email = st.text_input("Registered Email Address")
    api_key = st.text_input("Gemini API Key", type="password")

    st.divider()
    st.header("👤 Teacher Profile & Position")
    teacher_name = st.text_input("Teacher Name", "Norberto P. Binondo Jr.")
    position_rank = st.selectbox("Position / Rank", list(CAREER_STAGES.keys()))

    stage_info = CAREER_STAGES[position_rank]
    st.info(
        f"**Career Stage:** {stage_info['stage']}\n\n**COT Scale:** {stage_info['scale']}"
    )

# ------------------------------------------------------------------------------
# MAIN FORM: LESSON DETAILS & COT SELECTION
# ------------------------------------------------------------------------------
st.subheader("1. Lesson Details")
col1, col2, col3 = st.columns(3)

with col1:
    grade_level = st.selectbox(
        "Grade Level",
        [
            "Grade 7",
            "Grade 8",
            "Grade 9",
            "Grade 10",
            "Grade 11",
            "Grade 12",
        ],
    )
    subject = st.text_input("Subject / Learning Area", "Mathematics")

with col2:
    school_year = st.selectbox(
        "School Year", list(COT_INDICATORS_BY_SY.keys())
    )
    quarter = st.selectbox(
        "Quarter", ["Quarter 1", "Quarter 2", "Quarter 3", "Quarter 4"]
    )

with col3:
    topic = st.text_input(
        "Lesson Topic / Title", "Quadrilaterals and Parallelograms"
    )
    duration = st.text_input("Duration / Time Allotment", "60 Minutes")

st.divider()
st.subheader("2. Select COT Indicators to Integrate")

available_indicators = COT_INDICATORS_BY_SY[school_year]
selected_indicators = []

for code, desc in available_indicators:
    if st.checkbox(f"**[{code}]** {desc}", value=False):
        selected_indicators.append(f"Indicator {code}: {desc}")

st.divider()
st.subheader("3. Learning Objectives & Content Standard")
content_standard = st.text_area(
    "Content Standard",
    "The learner demonstrates understanding of key concepts of quadrilaterals.",
)
learning_competency = st.text_area(
    "Learning Competency / Code",
    "Solves problems involving parallelograms, trapezoids and kites. (M9GE-IIIe-1)",
)

# ------------------------------------------------------------------------------
# 3. GENERATION ENGINE (UPDATED GEMINI MODEL FALLBACK)
# ------------------------------------------------------------------------------
st.divider()

if st.button("🚀 Generate Lesson Plan", type="primary", use_container_width=True):
    if not api_key:
        st.error(
            "Please provide a valid Gemini API Key in the sidebar to generate the lesson plan."
        )
    elif not license_key or not user_email:
        st.warning(
            "Please enter your registered Email and License Key to proceed."
        )
    else:
        with st.spinner(
            "Generating DepEd DLL and integrating COT indicators..."
        ):
            try:
                genai.configure(api_key=api_key.strip())

                # Updated candidates with latest supported flash models
                model_candidates = [
                    "gemini-3.8-flash",
                    "gemini-2.5-flash",
                    "gemini-1.5-flash",
                    "gemini-1.5-pro",
                ]

                cot_prompt_text = "\n".join(
                    [f"- {ind}" for ind in selected_indicators]
                )

                prompt = f"""
                You are an expert DepEd Instructional Designer. Create a detailed Daily Lesson Log (DLL) / Lesson Plan for:
                
                Teacher: {teacher_name} ({position_rank} - Stage: {stage_info['stage']})
                Grade & Subject: {grade_level} {subject}
                School Year: {school_year} | {quarter}
                Topic: {topic} ({duration})
                
                Content Standard: {content_standard}
                Learning Competency: {learning_competency}
                
                Integrate the following target COT Indicators naturally into the teaching-learning procedures:
                {cot_prompt_text if cot_prompt_text else "Apply standard pedagogical strategies."}
                
                Structure the output with clear headers:
                I. OBJECTIVES
                II. CONTENT & LEARNING RESOURCES
                III. PROCEDURES (Explicitly label where COT indicators are integrated)
                   A. Reviewing previous lesson
                   B. Establishing a purpose
                   C. Presenting examples
                   D. Discussing new concepts & practicing skills
                   E. Developing mastery
                   F. Finding practical applications
                   G. Making generalizations
                   H. Evaluating learning
                   I. Additional activities
                IV. REMARKS & REFLECTION
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
                    raise Exception(
                        f"Unable to generate content with provided key. Last error: {str(last_error)}"
                    )

                generated_text = response.text

                st.success(
                    f"Lesson Plan successfully generated using model: `{successful_model}`!"
                )
                st.markdown(generated_text)

                # Generate Word Document (.docx)
                doc = docx.Document()
                doc.add_heading(f"LESSON PLAN: {topic}", 0)
                doc.add_paragraph(f"Teacher: {teacher_name} ({position_rank})")
                doc.add_paragraph(f"Grade & Subject: {grade_level} - {subject}")
                doc.add_paragraph(f"School Year: {school_year} | {quarter}")
                doc.add_paragraph("\n" + generated_text)

                doc_buffer = io.BytesIO()
                doc.save(doc_buffer)
                doc_buffer.seek(0)

                st.download_button(
                    label="📄 Download as Word Document (.docx)",
                    data=doc_buffer,
                    file_name=f"Lesson_Plan_{topic.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )

            except Exception as e:
                st.error(f"Generation Error: {str(e)}")