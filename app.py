import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from fpdf import FPDF
import google.generativeai as genai
import io
import json

# Page Configuration
st.set_page_config(page_title="Universal Master Lesson Package Generator", page_icon="📚", layout="wide")

# Title & Description
st.title("📚 Universal Master Lesson Package Generator")
st.caption("Automated generation of data-dense Master Lesson Packages (Excel + PDF) across all DepEd Subject Areas.")
st.divider()

# Sidebar Setup
with st.sidebar:
    st.header("⚙️ App Authentication")
    st.markdown("Input your Gemini API Key to run the high-density content engine.")
    api_key = st.text_input("Gemini API Key", type="password")
    st.info("💡 Ensure `openpyxl` and `fpdf2` are added to your `requirements.txt` file.")

# Universal Inputs
col1, col2 = st.columns(2)
with col1:
    subject_area = st.selectbox(
        "Subject Area / Domain", 
        [
            "Core Languages (Oral Com, Komunikasyon, Reading & Writing)",
            "Mathematics (GenMath, Stat & Probability)",
            "Natural Sciences (Earth & Life, Physical Science, Bio, Chem, Physics)",
            "Social Sciences & Humanities (UCSP, Creative Writing, Philosophy)",
            "Applied / Specialized Track (Inquiries, Entrepreneurship, Media Literacy)",
            "TVL Strands (ICT, Home Economics, Industrial Arts, Agri-Fishery)",
            "Sports and Arts & Design Tracks"
        ]
    )
    grade_level = st.selectbox("Grade Level / Key Stage", ["Grade 11", "Grade 12", "Junior High School", "Elementary"])
    topic = st.text_input("Competency / Focus Topic", placeholder="e.g., Quadratic Equations, Media Literacy, or Food Selection")

with col2:
    st.subheader("📌 Universal Production Parameters")
    class_size = st.slider("Estimated Class Size (For Activity Scaffolding)", min_value=10, max_value=80, value=50)
    st.markdown("""
    * **Zero Placeholders:** Every paragraph, question, and lesson cell will be generated with complete text.
    * **Flat List Implementation:** Eliminates cell array errors during full 50-item workbook mappings.
    * **Adaptive Localization:** Automatically anchors context to relevant Philippine settings, local statutes, or national regulatory agencies matching the selected subject domain.
    """)

# Helper Function: Clean JSON String from Gemini
def clean_json_response(text):
    text = text.strip()
    if text.startswith("```json"):
        text = text[7:]
    if text.endswith("```"):
        text = text[:-3]
    return text.strip()

# Core Execution Button
if st.button("🚀 Generate Multi-Discipline Lesson Package Assets", type="primary"):
    if not api_key:
        st.error("⚠️ Please enter your Gemini API Key in the sidebar.")
    elif not topic:
        st.warning("⚠️ Please provide a Competency/Topic focus before generating.")
    else:
        with st.spinner("🧠 Engineering specialized content structures... This takes a moment due to text density..."):
            try:
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel(model_name='gemini-1.5-flash', generation_config={"response_mime_type": "application/json"})
                
                user_prompt = f"Subject: {subject_area}\nGrade Level: {grade_level}\nTopic: {topic}\nClass Size: {class_size}\n"
                user_prompt += """
                Generate a JSON object matching this exact structure:
                {
                    "overview": {
                        "references": "Write 4-5 complete sentences citing official academic curricula frameworks, reference materials, or DepEd teacher resource guides valid for this discipline.",
                        "ai_use": "Write 4-5 complete sentences explaining the collaborative and transparent deployment of generative AI tools to configure learner-centered lesson scaffolding.",
                        "content_standard": "Write 4-5 complete sentences detailing the precise conceptual standards or industry expectations targeted by this topic.",
                        "performance_standard": "Write 4-5 complete sentences outlining the specific skill proficiency, practical output, or demonstration required from the learners.",
                        "learning_competency": "Write 4-5 complete sentences unpacking the specific targeted learning competency from the official curriculum map.",
                        "objectives": "Write 4-5 complete sentences establishing three explicit, measurable behavior objectives spanning cognitive domain and performance skills.",
                        "context": "Write 4-5 complete sentences tracking class size realities, regional inclusion factors, and localized learner profiles.",
                        "pre_lesson": "Write 4-5 complete sentences laying out diagnostic check methods or bridging activities to activate prior knowledge.",
                        "resources": "Write 4-5 complete sentences listing physical, print, or alternative improvised instructional materials appropriate for this subject area given limited local funds.",
                        "integration": "Write 4-5 complete sentences discussing cross-curricular linkages (e.g., tying this topic to numeracy, language proficiency, national heritage, or daily life skills).",
                        "exit_ticket": "Write 4-5 complete sentences illustrating the terminal formative assessment or quick performance metrics applied at the lesson's close.",
                        "extended_learning": "Write 4-5 complete sentences establishing reinforcement steps for mastery extensions alongside direct remediation strategies."
                    },
                    "sessions": [
                        {"session_num": 1, "teacher_action": "Explicit instructional delivery or lecture setup in 3 full sentences.", "student_task": "Active learning task, peer discussion, or drill for session 1 in 3 full sentences."},
                        {"session_num": 2, "teacher_action": "Explicit instructional delivery, lesson coaching, or practical demonstration in 3 full sentences.", "student_task": "Active learning task, workshop action, or practice sheet task for session 2 in 3 full sentences."},
                        {"session_num": 3, "teacher_action": "Instructional monitoring, diagnostic review, or scaffolding for session 3 in 3 full sentences.", "student_task": "Group execution, problem solving, or collaborative task for session 3 in 3 full sentences."},
                        {"session_num": 4, "teacher_action": "Scaffolding check, formatting support, or simulation guide for session 4 in 3 full sentences.", "student_task": "Performance defense, processing execution, or application activity for session 4 in 3 full sentences."},
                        {"session_num": 5, "teacher_action": "Assessment processing, evaluation tracking, or feedback routing for session 5 in 3 full sentences.", "student_task": "Summative assessment completion, reflection summary, or task synthesis for session 5 in 3 full sentences."}
                    ],
                    "cornell_notes": [
                        {"cue": "Core Idea 1", "content": "Three or more rigorous sentences outlining technical facts, functional paradigms, or rules governing this topic.", "summary": "One sentence synthesis statement."},
                        {"cue": "Core Idea 2", "content": "Three or more rigorous sentences outlining technical facts, functional paradigms, or rules governing this topic.", "summary": "One sentence synthesis statement."},
                        {"cue": "Core Idea 3", "content": "Three or more rigorous sentences outlining technical facts, functional paradigms, or rules governing this topic.", "summary": "One sentence synthesis statement."},
                        {"cue": "Core Idea 4", "content": "Three or more rigorous sentences outlining technical facts, functional paradigms, or rules governing this topic.", "summary": "One sentence synthesis statement."},
                        {"cue": "Core Idea 5", "content": "Three or more rigorous sentences outlining technical facts, functional paradigms, or rules governing this topic.", "summary": "One sentence synthesis statement."},
                        {"cue": "Core Idea 6", "content": "Three or more rigorous sentences outlining technical facts, functional paradigms, or rules governing this topic.", "summary": "One sentence synthesis statement."},
                        {"cue": "Core Idea 7", "content": "Three or more rigorous sentences outlining technical facts, functional paradigms, or rules governing this topic.", "summary": "One sentence synthesis statement."},
                        {"cue": "Core Idea 8", "content": "Three or more rigorous sentences outlining technical facts, functional paradigms, or rules governing this topic.", "summary": "One sentence synthesis statement."},
                        {"cue": "Core Idea 9", "content": "Three or more rigorous sentences outlining technical facts, functional paradigms, or rules governing
