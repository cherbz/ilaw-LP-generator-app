from flask import Flask, jsonify, request

app = Flask(__name__)

# ==============================================================================
# DATA STRUCTURES (COT Indicators & Career Stage Rules)
# ==============================================================================

CAREER_STAGES = {
    "beginning_to_proficient": {
        "title": "Beginning to Proficient",
        "positions": ["Teacher I", "Teacher II", "Teacher III"],
        "scale_min": 2,
        "scale_max": 6,
        "default_not_observed": 2,
    },
    "proficient": {
        "title": "Proficient",
        "positions": ["Teacher IV", "Teacher V", "Teacher VI", "Teacher VII"],
        "scale_min": 3,
        "scale_max": 7,
        "default_not_observed": 3,
    },
    "highly_proficient": {
        "title": "Highly Proficient",
        "positions": ["Master Teacher I", "Master Teacher II"],
        "scale_min": 4,
        "scale_max": 8,
        "default_not_observed": 4,
    },
    "distinguished": {
        "title": "Distinguished",
        "positions": [
            "Master Teacher III",
            "Master Teacher IV",
            "Master Teacher V",
        ],
        "scale_min": 5,
        "scale_max": 9,
        "default_not_observed": 5,
    },
}

COT_INDICATORS_BY_SY = {
    "2025-2026": [
        {
            "code": "1.1.2",
            "description": "Apply knowledge of content within and across curriculum teaching areas",
        },
        {
            "code": "1.4.2",
            "description": "Use a range of teaching strategies that enhance learner achievement in literacy and numeracy skills",
        },
        {
            "code": "1.5.2",
            "description": "Apply a range of teaching strategies to develop critical and creative thinking, as well as other higher-order thinking skills",
        },
        {
            "code": "2.3.2",
            "description": "Manage classroom structure to engage learners, individually or in groups, in meaningful exploration, discovery and hands-on activities within a range of physical learning environments",
        },
        {
            "code": "2.6.2",
            "description": "Manage learner behavior constructively by applying positive and non-violent discipline to ensure learning-focused environments",
        },
        {
            "code": "3.1.2",
            "description": "Use differentiated, developmentally appropriate learning experiences to address learners' gender, needs, strengths, interests and experiences",
        },
        {
            "code": "4.1.2",
            "description": "Plan, manage and implement developmentally sequenced teaching and learning process to meet curriculum requirements and varied teaching contexts",
        },
        {
            "code": "4.5.2",
            "description": "Select, develop, organize and use appropriate teaching and learning resources, including ICT, to address learning goals",
        },
        {
            "code": "5.1.2",
            "description": "Design, select, organize and use diagnostic, formative and summative assessment strategies consistent with curriculum requirements",
        },
    ],
    "2026-2027": [
        {
            "code": "1.1.2",
            "description": "Apply knowledge of content within and across curriculum teaching areas",
        },
        {
            "code": "1.4.2",
            "description": "Use a range of teaching strategies that enhance learner achievement in literacy and numeracy skills",
        },
        {
            "code": "1.5.2",
            "description": "Apply a range of teaching strategies to develop critical and creative thinking, as well as other higher-order thinking skills",
        },
        {
            "code": "1.6.2",
            "description": "Display proficient use of Mother Tongue, Filipino and English to facilitate teaching and learning",
        },
        {
            "code": "2.1.2",
            "description": "Establish safe and secure learning environments to enhance learning through the consistent implementation of policies, guidelines and procedures",
        },
        {
            "code": "2.2.2",
            "description": "Maintain learning environments that promote fairness, respect and care to encourage learning",
        },
        {
            "code": "3.2.2",
            "description": "Establish a learner-centered culture by using teaching strategies that respond to learners' linguistic, cultural, socio-economic and religious backgrounds",
        },
        {
            "code": "3.5.2",
            "description": "Adapt and use culturally appropriate teaching strategies to address the needs of learners from indigenous groups",
        },
        {
            "code": "5.3.2",
            "description": "Use strategies for providing timely, accurate and constructive feedback to improve learner performance",
        },
    ],
    "2027-2028": [
        {
            "code": "1.1.2",
            "description": "Apply knowledge of content within and across curriculum teaching areas",
        },
        {
            "code": "1.4.2",
            "description": "Use a range of teaching strategies that enhance learner achievement in literacy and numeracy skills",
        },
        {
            "code": "1.3.2",
            "description": "Ensure the positive use of ICT to facilitate the teaching and learning process",
        },
        {
            "code": "1.7.2",
            "description": "Use effective verbal and non-verbal classroom communication strategies to support learner understanding, participation, engagement and achievement",
        },
        {
            "code": "2.4.2",
            "description": "Maintain supportive learning environments that nurture and inspire learners to participate, cooperate and collaborate in continued learning",
        },
        {
            "code": "2.5.2",
            "description": "Apply a range of successful strategies that maintain learning environments that motivate learners to work productively by assuming responsibility for their own learning",
        },
        {
            "code": "3.3.2",
            "description": "Design, adapt and implement teaching strategies that are responsive to learners with disabilities, giftedness and talents",
        },
        {
            "code": "3.4.2",
            "description": "Plan and deliver teaching strategies that are responsive to the special educational needs of learners in difficult circumstances, including: geographic isolation; chronic illness; displacement due to armed conflict, urban resettlement or disasters; child abuse and child labor practices",
        },
    ],
}


# ==============================================================================
# HELPER FUNCTIONS
# ==============================================================================


def get_career_stage_by_position(position_name):
    """Finds the stage config matching a given teacher position name."""
    clean_position = position_name.strip().title()
    for key, stage in CAREER_STAGES.items():
        if clean_position in [p.title() for p in stage["positions"]]:
            return stage
    return None


# ==============================================================================
# API ENDPOINTS
# ==============================================================================


@app.route("/", methods=["GET"])
def home():
    """Root route providing system metadata and available API endpoints."""
    return jsonify(
        {
            "system": "Classroom Observation Tool (COT) Management API",
            "endpoints": {
                "get_career_stages": "/api/career-stages [GET]",
                "get_indicators": "/api/indicators/<school_year> [GET]",
                "evaluate_rating": "/api/evaluate-rating [POST]",
            },
        }
    )


@app.route("/api/career-stages", methods=["GET"])
def list_career_stages():
    """Returns rating scales and target positions across all career stages."""
    return jsonify({"status": "success", "data": CAREER_STAGES})


@app.route("/api/indicators/<school_year>", methods=["GET"])
def get_indicators(school_year):
    """Returns the COT indicators matrix for a specific School Year (e.g., '2025-2026')."""
    indicators = COT_INDICATORS_BY_SY.get(school_year)
    if not indicators:
        return (
            jsonify(
                {
                    "status": "error",
                    "message": f"School year '{school_year}' not found. Valid options: {list(COT_INDICATORS_BY_SY.keys())}",
                }
            ),
            404,
        )

    return jsonify(
        {
            "status": "success",
            "school_year": school_year,
            "total_indicators": len(indicators),
            "indicators": indicators,
        }
    )


@app.route("/api/evaluate-rating", methods=["POST"])
def evaluate_rating():
    """Evaluates an observation score for a specific position and indicator status."""
    payload = request.get_json() or {}
    position = payload.get("position")
    is_observed = payload.get("is_observed", True)
    given_score = payload.get("score")

    if not position:
        return (
            jsonify({"status": "error", "message": "Position is required."}),
            400,
        )

    stage_info = get_career_stage_by_position(position)
    if not stage_info:
        return (
            jsonify(
                {
                    "status": "error",
                    "message": f"Invalid position '{position}'.",
                }
            ),
            400,
        )

    if not is_observed:
        final_score = stage_info["default_not_observed"]
        note = f"Not Observed (NO) selected. Defaulted to minimum score level {final_score} for {stage_info['title']}."
    else:
        if (
            given_score is None
            or given_score < stage_info["scale_min"]
            or given_score > stage_info["scale_max"]
        ):
            return (
                jsonify(
                    {
                        "status": "error",
                        "message": f"Invalid score {given_score}. Valid rating scale for {position} ({stage_info['title']}) is {stage_info['scale_min']} to {stage_info['scale_max']}.",
                    }
                ),
                400,
            )
        final_score = given_score
        note = "Score accepted within standard scale bounds."

    return jsonify(
        {
            "status": "success",
            "position": position,
            "career_stage": stage_info["title"],
            "scale_range": f"{stage_info['scale_min']} - {stage_info['scale_max']}",
            "final_score": final_score,
            "note": note,
        }
    )


# ==============================================================================
# SERVER ENTRY POINT
# ==============================================================================

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)