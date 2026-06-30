import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from fpdf import FPDF
import google.generativeai as genai
import io
import json

# Page Configuration
st.set_page_config(
    page_title="Universal Master Lesson Package Generator",
    page_icon="📚",
    layout="wide"
)

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
                
                # Using a configuration model structure to separate code from text formatting constraints
                model = genai.GenerativeModel(
                    model_name='gemini-1.5-flash',
                    generation_config={"response_mime_type": "application/json"}
                )
                
                # System instructions separate from variables to resolve Python syntax clashes
                system_instruction = """
                You are a Master Teacher and Instructional Designer for the Philippine Department of Education. 
                Generate a data-dense master lesson package. You must output raw JSON format ONLY matching the requested schema. 
                Do not truncate text or use short summaries. Provide complete sentences for every item.
                """
                
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
                        {"cue": "Core Idea 9", "content": "Three or more rigorous sentences outlining technical facts, functional paradigms, or rules governing this topic.", "summary": "One sentence synthesis statement."},
                        {"cue": "Core Idea 10", "content": "Three or more rigorous sentences outlining technical facts, functional paradigms, or rules governing this topic.", "summary": "One sentence synthesis statement."}
                    ],
                    "glossary": [
                        {"term": "Term 1", "definition": "A high-density, precise definition spanning exactly two academic sentences."},
                        {"term": "Term 2", "definition": "A high-density, precise definition spanning exactly two academic sentences."},
                        {"term": "Term 3", "definition": "A high-density, precise definition spanning exactly two academic sentences."},
                        {"term": "Term 4", "definition": "A high-density, precise definition spanning exactly two academic sentences."},
                        {"term": "Term 5", "definition": "A high-density, precise definition spanning exactly two academic sentences."},
                        {"term": "Term 6", "definition": "A high-density, precise definition spanning exactly two academic sentences."},
                        {"term": "Term 7", "definition": "A high-density, precise definition spanning exactly two academic sentences."},
                        {"term": "Term 8", "definition": "A high-density, precise definition spanning exactly two academic sentences."},
                        {"term": "Term 9", "definition": "A high-density, precise definition spanning exactly two academic sentences."},
                        {"term": "Term 10", "definition": "A high-density, precise definition spanning exactly two academic sentences."},
                        {"term": "Term 11", "definition": "A high-density, precise definition spanning exactly two academic sentences."},
                        {"term": "Term 12", "definition": "A high-density, precise definition spanning exactly two academic sentences."},
                        {"term": "Term 13", "definition": "A high-density, precise definition spanning exactly two academic sentences."},
                        {"term": "Term 14", "definition": "A high-density, precise definition spanning exactly two academic sentences."},
                        {"term": "Term 15", "definition": "A high-density, precise definition spanning exactly two academic sentences."}
                    ],
                    "assessments": [
                        {"day": 1, "q": "A unique full sentence identification question testing factual knowledge item 1.", "a": "Answer1"},
                        {"day": 1, "q": "A unique full sentence identification question testing factual knowledge item 2.", "a": "Answer2"},
                        {"day": 1, "q": "A unique full sentence identification question testing factual knowledge item 3.", "a": "Answer3"},
                        {"day": 1, "q": "A unique full sentence identification question testing factual knowledge item 4.", "a": "Answer4"},
                        {"day": 1, "q": "A unique full sentence identification question testing factual knowledge item 5.", "a": "Answer5"},
                        {"day": 1, "q": "A unique full sentence identification question testing factual knowledge item 6.", "a": "Answer6"},
                        {"day": 1, "q": "A unique full sentence identification question testing factual knowledge item 7.", "a": "Answer7"},
                        {"day": 1, "q": "A unique full sentence identification question testing factual knowledge item 8.", "a": "Answer8"},
                        {"day": 1, "q": "A unique full sentence identification question testing factual knowledge item 9.", "a": "Answer9"},
                        {"day": 1, "q": "A unique full sentence identification question testing factual knowledge item 10.", "a": "Answer10"},
                        {"day": 2, "q": "A unique full sentence identification question testing factual knowledge item 11.", "a": "Answer11"},
                        {"day": 2, "q": "A unique full sentence identification question testing factual knowledge item 12.", "a": "Answer12"},
                        {"day": 2, "q": "A unique full sentence identification question testing factual knowledge item 13.", "a": "Answer13"},
                        {"day": 2, "q": "A unique full sentence identification question testing factual knowledge item 14.", "a": "Answer14"},
                        {"day": 2, "q": "A unique full sentence identification question testing factual knowledge item 15.", "a": "Answer15"},
                        {"day": 2, "q": "A unique full sentence identification question testing factual knowledge item 16.", "a": "Answer16"},
                        {"day": 2, "q": "A unique full sentence identification question testing factual knowledge item 17.", "a": "Answer17"},
                        {"day": 2, "q": "A unique full sentence identification question testing factual knowledge item 18.", "a": "Answer18"},
                        {"day": 2, "q": "A unique full sentence identification question testing factual knowledge item 19.", "a": "Answer19"},
                        {"day": 2, "q": "A unique full sentence identification question testing factual knowledge item 20.", "a": "Answer20"},
                        {"day": 3, "q": "A unique full sentence identification question testing factual knowledge item 21.", "a": "Answer21"},
                        {"day": 3, "q": "A unique full sentence identification question testing factual knowledge item 22.", "a": "Answer22"},
                        {"day": 3, "q": "A unique full sentence identification question testing factual knowledge item 23.", "a": "Answer23"},
                        {"day": 3, "q": "A unique full sentence identification question testing factual knowledge item 24.", "a": "Answer24"},
                        {"day": 3, "q": "A unique full sentence identification question testing factual knowledge item 25.", "a": "Answer25"},
                        {"day": 3, "q": "A unique full sentence identification question testing factual knowledge item 26.", "a": "Answer26"},
                        {"day": 3, "q": "A unique full sentence identification question testing factual knowledge item 27.", "a": "Answer27"},
                        {"day": 3, "q": "A unique full sentence identification question testing factual knowledge item 28.", "a": "Answer28"},
                        {"day": 3, "q": "A unique full sentence identification question testing factual knowledge item 29.", "a": "Answer29"},
                        {"day": 3, "q": "A unique full sentence identification question testing factual knowledge item 30.", "a": "Answer30"},
                        {"day": 4, "q": "A unique full sentence identification question testing factual knowledge item 31.", "a": "Answer31"},
                        {"day": 4, "q": "A unique full sentence identification question testing factual knowledge item 32.", "a": "Answer32"},
                        {"day": 4, "q": "A unique full sentence identification question testing factual knowledge item 33.", "a": "Answer33"},
                        {"day": 4, "q": "A unique full sentence identification question testing factual knowledge item 34.", "a": "Answer34"},
                        {"day": 4, "q": "A unique full sentence identification question testing factual knowledge item 35.", "a": "Answer35"},
                        {"day": 4, "q": "A unique full sentence identification question testing factual knowledge item 36.", "a": "Answer36"},
                        {"day": 4, "q": "A unique full sentence identification question testing factual knowledge item 37.", "a": "Answer37"},
                        {"day": 4, "q": "A unique full sentence identification question testing factual knowledge item 38.", "a": "Answer38"},
                        {"day": 4, "q": "A unique full sentence identification question testing factual knowledge item 39.", "a": "Answer39"},
                        {"day": 4, "q": "A unique full sentence identification question testing factual knowledge item 40.", "a": "Answer40"},
                        {"day": 5, "q": "A unique full sentence identification question testing factual knowledge item 41.", "a": "Answer41"},
                        {"day": 5, "q": "A unique full sentence identification question testing factual knowledge item 42.", "a": "Answer42"},
                        {"day": 5, "q": "A unique full sentence identification question testing factual knowledge item 43.", "a": "Answer43"},
                        {"day": 5, "q": "A unique full sentence identification question testing factual knowledge item 44.", "a": "Answer44"},
                        {"day": 5, "q": "A unique full sentence identification question testing factual knowledge item 45.", "a": "Answer45"},
                        {"day": 5, "q": "A unique full sentence identification question testing factual knowledge item 46.", "a": "Answer46"},
                        {"day": 5, "q": "A unique full sentence identification question testing factual knowledge item 47.", "a": "Answer47"},
                        {"day": 5, "q": "A unique full sentence identification question testing factual knowledge item 48.", "a": "Answer48"},
                        {"day": 5, "q": "A unique full sentence identification question testing factual knowledge item 49.", "a": "Answer49"},
                        {"day": 5, "q": "A unique full sentence identification question testing factual knowledge item 50.", "a": "Answer50"}
                    ],
                    "article_1": {
                        "title": "Contextualized title exploring this topic inside Philippine societal framework, communities, or relevant national agencies",
                        "body": "A complete academic narrative of 200-300 words analyzing the real-world utility or cultural connections of this topic in the Philippines."
                    },
                    "article_2": {
                        "title": "A second title highlighting localized case examples, national policies, or community impacts across the archipelago",
                        "body": "A complete contextual essay of 200-300 words devoid of placeholders detailing the practical application of this topic nationally."
                    }
                }
                """
                
                response = model.generate_content(
                    contents=user_prompt,
                    generation_config={"response_mime_type": "application/json"}
                )
                
                raw_data = json.loads(clean_json_response(response.text))
                
                # --- FILE 1: MASTER_LESSON_PACKAGE.XLSX GENERATION ---
                wb = openpyxl.Workbook()
                
                # Styling configurations
                font_title = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
                font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                font_bold = Font(name="Calibri", size=11, bold=True)
                font_regular = Font(name="Calibri", size=11)
                fill_primary = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                fill_accent = PatternFill(start_color="F2DCDB", end_color="F2DCDB", fill_type="solid")
                align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
                align_left = Alignment(horizontal="left", vertical="top", wrap_text=True)

                # Tab 1: Lesson_Plan_Overview
                ws1 = wb.active
                ws1.title = "Lesson_Plan_Overview"
                ws1.views.sheetView[0].showGridLines = True
                
                ws1.merge_cells("A1:C1")
                ws1["A1"] = f"UNIVERSAL LESSON OVERVIEW: {topic.upper()} ({grade_level})"
                ws1["A1"].font = font_title
                ws1["A1"].fill = fill_primary
                ws1["A1"].alignment = align_center
                ws1.row_dimensions[1].height = 40
                
                overview_fields = [
                    ("A. References", "references"), ("B. Declaration of AI use", "ai_use"),
                    ("C. Content Standard", "content_standard"), ("D. Performance Standard", "performance_standard"),
                    ("E. Learning Competency", "learning_competency"), ("F. Learning Objectives", "objectives"),
                    ("G. Learning Context", "context"), ("H. Pre-Lesson", "pre_lesson"),
                    ("J. Learning Resources", "resources"), ("K. Opportunities for Integration", "integration"),
                    ("L. Formative Assessment", "exit_ticket"), ("M. Extended Learning Activities", "extended_learning")
                ]
                
                current_row = 3
                for label, key in overview_fields:
                    ws1.cell(row=current_row, column=1, value=label).font = font_bold
                    val = raw_data.get("overview", {}).get(key, "Data generation incomplete.")
                    ws1.cell(row=current_row, column=2, value=val).font = font_regular
                    ws1.cell(
