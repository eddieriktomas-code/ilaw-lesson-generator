import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
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

# Helper Function: Clean JSON String safely without relying on literal triple backtick strings
def clean_json_response(text):
    text = text.strip()
    prefix = "```" + "json"
    suffix = "```"
    if text.startswith(prefix):
        text = text[len(prefix):]
    if text.endswith(suffix):
        text = text[:-len(suffix)]
    return text.strip()

# Helper Function: Sanitize strings to prevent Latin-1 encoding crashes in standard PDF fonts
def sanitize_for_pdf(text):
    if not text:
        return ""
    # Map common non-Latin-1/smart characters to safe equivalents
    replacements = {
        '\u201c': '"', '\u201d': '"',  # Smart double quotes
        '\u2018': "'", '\u2019': "'",  # Smart single quotes
        '\u2014': '-', '\u2013': '-',  # Em/En dashes
        '\u2022': '*',                  # Bullet points
        '\xf1': 'n', '\xd1': 'N',      # Spanish ñ and Ñ
        '\u200b': '',                  # Zero-width spaces
    }
    for uni_char, ascii_char in replacements.items():
        text = text.replace(uni_char, ascii_char)
    return text.encode('latin-1', 'replace').decode('latin-1')

# Helper Function: Create and Format Excel Sheets safely without code duplication
def create_styled_sheet(workbook, title, headers, fill, font):
    ws = workbook.active if title == "Lesson_Plan_Overview" else workbook.create_sheet(title=title)
    ws.views.sheetView[0].showGridLines = True
    
    # Define clean, professional light gray borders
    thin_border = Border(
        left=Side(style='thin', color='D3D3D3'),
        right=Side(style='thin', color='D3D3D3'),
        top=Side(style='thin', color='D3D3D3'),
        bottom=Side(style='thin', color='D3D3D3')
    )
    
    if title != "Lesson_Plan_Overview":
        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx, value=h)
            cell.font = font
            cell.fill = fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        ws.row_dimensions[1].height = 25
    return ws

# Core Execution Button
if st.button("🚀 Generate Multi-Discipline Lesson Package Assets", type="primary"):
    if not api_key or not topic:
        st.warning("⚠️ Please provide your Gemini API Key and Competency Topic focus.")
    else:
        with st.spinner("🧠 Engineering specialized content structures... Please wait..."):
            try:
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel(model_name='gemini-1.5-flash', generation_config={"response_mime_type": "application/json"})
                
                # Structural schema compressed to prevent string buffer overflows
                json_schema = {
                    "overview": {
                        "references": "4-5 complete sentences citing frameworks.", "ai_use": "4-5 complete sentences on AI transparency.",
                        "content_standard": "4-5 complete sentences on conceptual standards.", "performance_standard": "4-5 sentences outlining practical skills.",
                        "learning_competency": "4-5 sentences on targeted curriculum maps.", "objectives": "3 explicit measurable objectives.",
                        "context": "Class size, local profile constraints.", "pre_lesson": "Diagnostic check tracking.",
                        "resources": "Physical print material targets.", "integration": "Cross-curricular links.",
                        "exit_ticket": "Terminal assessment rules.", "extended_learning": "Remediation and reinforcement plans."
                    },
                    "sessions": [
                        {"session_num": 1, "teacher_action": "3 clear sentences.", "student_task": "3 clear sentences."},
                        {"session_num": 2, "teacher_action": "3 clear sentences.", "student_task": "3 clear sentences."},
                        {"session_num": 3, "teacher_action": "3 clear sentences.", "student_task": "3 clear sentences."},
                        {"session_num": 4, "teacher_action": "3 clear sentences.", "student_task": "3 clear sentences."},
                        {"session_num": 5, "teacher_action": "3 clear sentences.", "student_task": "3 clear sentences."}
                    ],
                    "cornell_notes": [{"cue": "Idea Point", "content": "3 rigorous context sentences.", "summary": "1 sentence summary."}],
                    "glossary": [{"term": "Term Anchor", "definition": "Exactly 2 detailed definitions sentences."}],
                    "assessments": [{"day": 1, "q": "Factual question item sentence.", "a": "Exact expected answer key response."}],
                    "article_1": {"title": "Philippine Context Title", "body": "200-300 word narrative framework."},
                    "article_2": {"title": "Localized Case Analysis Title", "body": "200-300 word contextual case study."}
                }
                
                # Safer concatenation pattern to prevent f-string bracket evaluation conflicts
                user_prompt = "Subject Area Domain: " + str(subject_area) + "\n"
                user_prompt += "Target Learner Grade Level: " + str(grade_level) + "\n"
                user_prompt += "Selected Instructional Topic: " + str(topic) + "\n"
                user_prompt += "Expected Student Count: " + str(class_size) + " learners\n\n"
                user_prompt += "Instructions: You must generate a single, dense JSON object matching the schema below. Fill out all properties fully. Do not truncate text:\n"
                user_prompt += json.dumps(json_schema, indent=2) + "\n\n"
                user_prompt += "REQUIRED SCALE OUT: Expand output array components to populate exactly 5 distinct sessions, 10 distinct cornell_notes entries, 15 complete glossary definitions, and 50 unique daily assessment items across days 1 to 5."

                response = model.generate_content(contents=user_prompt)
                raw_data = json.loads(clean_json_response(response.text))
                
                # --- EXCEL COMPILATION ---
                wb = openpyxl.Workbook()
                f_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                f_bold = Font(name="Calibri", size=11, bold=True)
                f_regular = Font(name="Calibri", size=11)
                fill_p = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                fill_a = PatternFill(start_color="F2DCDB", end_color="F2DCDB", fill_type="solid")
                a_left = Alignment(horizontal="left", vertical="top", wrap_text=True)
                thin_border = Border(
                    left=Side(style='thin', color='D3D3D3'),
                    right=Side(style='thin', color='D3D3D3'),
                    top=Side(style='thin', color='D3D3D3'),
                    bottom=Side(style='thin', color='D3D3D3')
                )

                # Tab 1: Plan Overview Setup
                ws1 = create_styled_sheet(wb, "Lesson_Plan_Overview", [], fill_p, f_header)
                ws1.merge_cells("A1:C1")
                ws1["A1"] = f"LESSON OVERVIEW: {topic.upper()}"
                ws1["A1"].font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
                ws1["A1"].fill = fill_p
                ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")
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
                    c1 = ws1.cell(row=curr_row, column=1, value=label)
                    c1.font = f_bold
                    c1.border = thin_border
                    
                    val = raw_data.get("overview", {}).get(key, "Data incomplete.")
                    c2 = ws1.cell(row=curr_row, column=2, value=val)
                    c2.font = f_regular
                    c2.alignment = a_left
                    c2.border = thin_border
                    
                    ws1.row_dimensions[curr_row].height = 65
                    curr_row += 1
                
                curr_row += 2
                ws1.cell(row=curr_row, column=1, value="I. Universal Lesson Flow").font = f_bold
                curr_row += 1
                
                for idx, h in enumerate(["Session Mapping", "Teacher Execution Plan", "Student Expected Task"], 1):
                    c = ws1.cell(row=curr_row, column=idx, value=h)
                    c.font = f_header
                    c.fill = fill_p
                    c.alignment = Alignment(horizontal="center", vertical="center")
                    c.border = thin_border
                ws1.row_dimensions[curr_row].height = 25
                
                for session in raw_data.get("sessions", []):
                    curr_row += 1
                    c1 = ws1.cell(row=curr_row, column=1, value=f"Session {session.get('session_num', '')}")
                    c1.font = f_bold
                    c1.border = thin_border
                    
                    c2 = ws1.cell(row=curr_row, column=2, value=session.get('teacher_action', ''))
                    c2.font = f_regular
                    c2.alignment = a_left
                    c2.border = thin_border
                    
                    c3 = ws1.cell(row=curr_row, column=3, value=session.get('student_task', ''))
                    c3.font = f_regular
                    c3.alignment = a_left
                    c3.border = thin_border
                    
                    ws1.row_dimensions[curr_row].height = 60
                
                # Tab 2: Cornell Notes
                ws2 = create_styled_sheet(wb, "Cornell_Notes", ["Conceptual Cue", "Factual Technical Discussion", "Core Synthesis Sentence"], fill_p, f_header)
                for r_idx, note in enumerate(raw_data.get("cornell_notes", []), 2):
                    c1 = ws2.cell(row=r_idx, column=1, value=note.get('cue', ''))
                    c1.font = f_bold; c1.alignment = a_left; c1.border = thin_border
                    
                    c2 = ws2.cell(row=r_idx, column=2, value=note.get('content', ''))
                    c2.font = f_regular; c2.alignment = a_left; c2.border = thin_border
                    
                    c3 = ws2.cell(row=r_idx, column=3, value=note.get('summary', ''))
                    c3.font = f_regular; c3.alignment
