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
                
                # Defending against string interpretation bugs by using a clean native dictionary
                json_schema = {
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
                        {"session_num": 5, "teacher_action": "Assessment processing, evaluation tracking, or feedback routing for session 5 in 3 full sentences.", "student_task": "Summative assessment completion, reflection summary, or task synthesis for session 5 in 3 full sentences."}
                    ],
                    "cornell_notes": [
                        {"cue": "Core Idea 1", "content": "Three or more rigorous sentences outlining technical facts, functional paradigms, or rules governing this topic.", "summary": "One sentence synthesis statement."}
                    ],
                    "glossary": [
                        {"term": "Term 1", "definition": "A high-density, precise definition spanning exactly two academic sentences."}
                    ],
                    "assessments": [
                        {"day": 1, "q": "A unique full sentence identification question testing factual knowledge item 1.", "a": "Answer1"}
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
                
                # Safely convert structural dictionary into string without quote collisions
                schema_instruction = json.dumps(json_schema, indent=2)
                user_prompt = f"Subject: {subject_area}\nGrade Level: {grade_level}\nTopic: {topic}\nClass Size: {class_size}\n\nGenerate a JSON object matching this exact structure structure:\n{schema_instruction}\n"
                user_prompt += "\nCRITICAL: Please ensure to scale out the response to populate a complete array structure containing 5 total sessions, 10 distinct cornell_notes entries, 15 complete glossary technical terms, and 50 unique daily assessment items across days 1 to 5 as structurally represented by the template layout blocks."

                response = model.generate_content(contents=user_prompt)
                raw_data = json.loads(clean_json_response(response.text))
                
                # --- FILE 1: MASTER_LESSON_PACKAGE.XLSX ---
                wb = openpyxl.Workbook()
                
                f_title = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
                f_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                f_bold = Font(name="Calibri", size=11, bold=True)
                f_regular = Font(name="Calibri", size=11)
                fill_p = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                fill_a = PatternFill(start_color="F2DCDB", end_color="F2DCDB", fill_type="solid")
                a_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
                a_left = Alignment(horizontal="left", vertical="top", wrap_text=True)

                # Tab 1: Plan Overview
                ws1 = wb.active
                ws1.title = "Lesson_Plan_Overview"
                ws1.views.sheetView[0].showGridLines = True
                
                ws1.merge_cells("A1:C1")
                ws1["A1"] = f"LESSON OVERVIEW: {topic.upper()}"
                ws1["A1"].font = f_title; ws1["A1"].fill = fill_p; ws1["A1"].alignment = a_center
                ws1.row_dimensions[1].height = 40
                
                overview_fields = [
                    ("A. References", "references"), ("B. Declaration of AI use", "ai_use"),
                    ("C. Content Standard", "content_standard"), ("D. Performance Standard", "performance_standard"),
                    ("E. Learning Competency", "learning_competency"), ("F. Learning Objectives", "objectives"),
                    ("G. Learning Context", "context"), ("H. Pre-Lesson", "pre_lesson"),
                    ("J. Learning Resources", "resources"), ("K. Opportunities for Integration", "integration"),
                    ("L. Formative Assessment", "exit_ticket"), ("M. Extended Learning Activities", "extended_learning")
                ]
                
                curr_row = 3
                for label, key in overview_fields:
                    ws1.cell(row=curr_row, column=1, value=label).font = f_bold
                    val = raw_data.get("overview", {}).get(key, "Data incomplete.")
                    ws1.cell(row=curr_row, column=2, value=val).font = f_regular
                    ws1.cell(row=curr_row, column=2).alignment = a_left
                    ws1.row_dimensions[curr_row].height = 65
                    curr_row += 1
                
                curr_row += 2
                ws1.cell(row=curr_row, column=1, value="I. Universal Lesson Flow").font = f_bold
                curr_row += 1
                
                headers_s = ["Session Mapping", "Teacher Execution Plan", "Student Expected Task"]
                for col_idx, h in enumerate(headers_s, 1):
                    cell = ws1.cell(row=curr_row, column=col_idx, value=h)
                    cell.font = f_header; cell.fill = fill_p; cell.alignment = a_center
                ws1.row_dimensions[curr_row].height = 25
                
                for session in raw_data.get("sessions", []):
                    curr_row += 1
                    ws1.cell(row=curr_row, column=1, value=f"Session {session.get('session_num', '')}").font = f_bold
                    ws1.cell(row=curr_row, column=2, value=session.get('teacher_action', '')).font = f_regular
                    ws1.cell(row=curr_row, column=3, value=session.get('student_task', '')).font = f_regular
                    ws1.cell(row=curr_row, column=2).alignment = a_left
                    ws1.cell(row=curr_row, column=3).alignment = a_left
                    ws1.row_dimensions[curr_row].height = 60
                
                # Tab 2: Cornell Notes
                ws2 = wb.create_sheet(title="Cornell_Notes")
                ws2.views.sheetView[0].showGridLines = True
                headers_c = ["Conceptual Cue", "Factual Technical Discussion", "Core Synthesis Sentence"]
                for col_idx, h in enumerate(headers_c, 1):
                    cell = ws2.cell(row=1, column=col_idx, value=h)
                    cell.font = f_header; cell.fill = fill_p; cell.alignment = a_center
                
                for r_idx, note in enumerate(raw_data.get("cornell_notes", []), 2):
                    ws2.cell(row=r_idx, column=1, value=note.get('cue', '')).font = f_bold
                    ws2.cell(row=r_idx, column=2, value=note.get('content', '')).font = f_regular
                    ws2.cell(row=r_idx, column=3, value=note.get('summary', '')).font = f_regular
                    ws2.cell(row=r_idx, column=1).alignment = a_left
                    ws2.cell(row=r_idx, column=2).alignment = a_left
                    ws2.cell(row=r_idx, column=3).alignment = a_left
                    ws2.row_dimensions[r_idx].height = 65

                # Tab 3: Glossary
                ws3 = wb.create_sheet(title="Glossary")
                ws3.views.sheetView[0].showGridLines = True
                headers_g = ["Technical Term", "Expanded Definition"]
                for col_idx, h in enumerate(headers_g, 1):
                    cell = ws3.cell(row=1, column=col_idx, value=h)
                    cell.font = f_header; cell.fill = fill_p; cell.alignment = a_center
                
                for r_idx, g_item in enumerate(raw_data.get("glossary", []), 2):
                    ws3.cell(row=r_idx, column=1, value=g_item.get('term', '')).font = f_bold
                    ws3.cell(row=r_idx, column=2, value=g_item.get('definition', '')).font = f_regular
                    ws3.cell(row=r_idx, column=1).alignment = a_left
                    ws3.cell(row=r_idx, column=2).alignment = a_left
                    ws3.row_dimensions[r_idx].height = 40

                # Tab 4: Student Assessment
                ws4 = wb.create_sheet(title="Student_Assessment")
                ws4.views.sheetView[0].showGridLines = True
                headers_a = ["Item ID", "Tracking Window", "Identification Question Item"]
                for col_idx, h in enumerate(headers_a, 1):
                    cell = ws4.cell(row=1, column=col_idx, value=h)
                    cell.font = f_header; cell.fill = fill_p; cell.alignment = a_center
                
                for r_idx, ass in enumerate(raw_data.get("assessments", []), 2):
                    ws4.cell(row=r_idx, column=1, value=f"Item-{r_idx-1:02d}").font = f_bold
                    ws4.cell(row=r_idx, column=2, value=f"Day {ass.get('day', '')}").font = f_regular
                    ws4.cell(row=r_idx, column=3, value=ass.get('q', '')).font = f_regular
                    ws4.cell(row=r_idx, column=3).alignment = a_left
                    ws4.row_dimensions[r_idx].height = 35

                # Tab 5: Answer Key
                ws5 = wb.create_sheet(title="Teacher_Answer_Key")
                ws5.views.sheetView[0].showGridLines = True
                headers_ak = ["Item ID", "Target Answer Value"]
                for col_idx, h in enumerate(headers_ak, 1):
                    cell = ws5.cell(row=1, column=col_
