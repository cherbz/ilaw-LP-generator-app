import streamlit as st

# ==============================================================================
# DATA STRUCTURES
# ==============================================================================

CAREER_STAGES = {
    "Beginning to Proficient": {
        "positions": ["Teacher I", "Teacher II", "Teacher III"],
        "scale": "2 to 6",
        "default_no": 2,
    },
    "Proficient": {
        "positions": ["Teacher IV", "Teacher V", "Teacher VI", "Teacher VII"],
        "scale": "3 to 7",
        "default_no": 3,
    },
    "Highly Proficient": {
        "positions": ["Master Teacher I", "Master Teacher II"],
        "scale": "4 to 8",
        "default_no": 4,
    },
    "Distinguished": {
        "positions": [
            "Master Teacher III",
            "Master Teacher IV",
            "Master Teacher V",
        ],
        "scale": "5 to 9",
        "default_no": 5,
    },
}

COT_INDICATORS_BY_SY = {
    "2025-2026": [
        "1.1.2 - Apply knowledge of content within and across curriculum teaching areas",
        "1.4.2 - Use a range of teaching strategies that enhance learner achievement in literacy and numeracy skills",
        "1.5.2 - Apply a range of teaching strategies to develop critical and creative thinking...",
        "2.3.2 - Manage classroom structure to engage learners...",
        "2.6.2 - Manage learner behavior constructively...",
        "3.1.2 - Use differentiated, developmentally appropriate learning experiences...",
        "4.1.2 - Plan, manage and implement developmentally sequenced teaching...",
        "4.5.2 - Select, develop, organize and use appropriate teaching and learning resources...",
        "5.1.2 - Design, select, organize and use diagnostic, formative and summative assessment...",
    ],
    "2026-2027": [
        "1.1.2 - Apply knowledge of content within and across curriculum teaching areas",
        "1.4.2 - Use a range of teaching strategies that enhance learner achievement...",
        "1.5.2 - Apply a range of teaching strategies to develop critical and creative thinking...",
        "1.6.2 - Display proficient use of Mother Tongue, Filipino and English...",
        "2.1.2 - Establish safe and secure learning environments...",
        "2.2.2 - Maintain learning environments that promote fairness, respect and care...",
        "3.2.2 - Establish a learner-centered culture...",
        "3.5.2 - Adapt and use culturally appropriate teaching strategies...",
        "5.3.2 - Use strategies for providing timely, accurate and constructive feedback...",
    ],
    "2027-2028": [
        "1.1.2 - Apply knowledge of content within and across curriculum teaching areas",
        "1.4.2 - Use a range of teaching strategies that enhance learner achievement...",
        "1.3.2 - Ensure the positive use of ICT to facilitate teaching and learning",
        "1.7.2 - Use effective verbal and non-verbal classroom communication strategies...",
        "2.4.2 - Maintain supportive learning environments...",
        "2.5.2 - Apply a range of successful strategies that maintain learning environments...",
        "3.3.2 - Design, adapt and implement teaching strategies for learners with disabilities...",
        "3.4.2 - Plan and deliver teaching strategies responsive to special educational needs...",
    ],
}

# ==============================================================================
# STREAMLIT APP UI
# ==============================================================================

st.title("Classroom Observation Tool (COT) Helper")

tab1, tab2 = st.tabs(["Indicators by SY", "Career Stage Evaluator"])

with tab1:
    st.header("COT Indicators")
    sy = st.selectbox("Select School Year", list(COT_INDICATORS_BY_SY.keys()))
    indicators = COT_INDICATORS_BY_SY[sy]
    st.write(f"### Indicators for {sy} ({len(indicators)} total)")
    for idx, ind in enumerate(indicators, 1):
        st.markdown(f"**{idx}.** {ind}")

with tab2:
    st.header("Rating Evaluator")
    stage_name = st.selectbox("Select Career Stage", list(CAREER_STAGES.keys()))
    stage = CAREER_STAGES[stage_name]

    st.info(
        f"**Target Positions:** {', '.join(stage['positions'])}\n\n**Rating Scale:** {stage['scale']}"
    )

    is_observed = st.checkbox("Was the indicator observed?", value=True)

    if not is_observed:
        st.warning(
            f"Not Observed (NO) selected. Standard default score: **{stage['default_no']}**"
        )
    else:
        min_s, max_s = map(int, stage["scale"].split(" to "))
        score = st.slider("Select Score", min_value=min_s, max_value=max_s)
        st.success(f"Selected Score: **{score}**")