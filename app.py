import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import google.generativeai as genai
import io
import json

# Page Configuration
st.set_page_config(page_title="Jenalyn's Master Lesson Plan Generator", page_icon="📚", layout="wide")

st.title("📚 Jenalyn's Master Lesson Plan Generator")
st.caption("Produced by Jenalyn, Junior & Senior High School Specialist. Formulates data-dense, compliant DepEd Master Lesson Plans.")
st.divider()

# Sidebar Configuration
with st.sidebar:
    st.header("⚙️ Credentials")
    api_key = st.text_input("Gemini API Key", type="password")
    st.info("💡 Make sure 'openpyxl' and 'google-generativeai' are listed in your requirements.txt file.")

# User Inputs
col1, col2 = st.columns(2)
with col1:
    subject = st.text_input("Subject Area", value="Core Languages (Reading and Writing Skills)")
    grade_level = st.selectbox("Grade Level", ["Grade 7", "Grade 8", "Grade 9", "Grade 10", "Grade 11", "Grade 12"])
    melc_topic = st.text_input("MELC / Focus Topic", placeholder="e.g., Formulating Evaluative Statements")

with col2:
    sessions_count = st.slider("Number of Sessions", min_value=1, max_value=5, value=5)
    st.subheader("💡 Specialization Protocol")
    st.caption("Ensures 100% data-dense contexts, localized examples, and text-friendly structural processing.")

# Helper function to remove potential code blocks from AI output safely
def clean_json_response(text):
    text = text.strip()
    prefix = "```" + "json"
    suffix = "```"
    if text.startswith(prefix):
        text = text[len(prefix):]
    if text.endswith(suffix):
        text = text[:-len(suffix)]
    return text.strip()

# Execution Process
if st.button("🚀 Compile Master Lesson Plan Asset", type="primary"):
    if not api_key or not melc_topic:
        st.warning("⚠️ Please fill in your Gemini API Key and your targeted MELC Focus Topic.")
    else:
        with st.spinner("🧠 Jenalyn is processing curriculum matrices and localizing contents..."):
            try:
                genai.configure(api_key=api_key)
                
                # FIXED: Updated model routing naming convention to prevent 404 API version errors
                model = genai.GenerativeModel(
                    model_name='gemini-2.5-flash',
                    generation_config={"response_mime_type": "application/json"}
                )
                
                # Simplified and direct strict target schema
                json_schema = {
                    "A_References": "4-5 descriptive sentences citing curriculum guides and local textbook resources.",
                    "B_AI_Declaration": "4-5 sentences detailing how generative AI tools assisted and structured this blueprint context safely.",
                    "C_Content_Standard": "4-5 deep sentences explaining key concepts learners must understand.",
                    "D_Performance_Standard": "4-5 specific sentences detailing the practical application of skill metrics.",
                    "E_Learning_Competency": "4-5 sentences linking the target task back to standard DepEd curriculum maps.",
                    "F_Learning_Objectives": "At least 3 highly measurable terminal performance conditions.",
                    "G_Learning_Context": "4-5 sentences regarding cultural backdrop, class setup, and prior skill mastery.",
                    "H_Pre_Lesson": "4-5 sentences mapping out diagnostic assessment steps or review warm-ups.",
                    "J_Learning_Resources": "4-5 sentences mapping printed, digital, or localized real-world instructional media targets.",
                    "K_Integration_Opportunities": "4-5 sentences connecting the topic to other subjects (e.g., Social Sciences, TVL) and real-world localized situations.",
                    "L_Formative_Assessment": "4-5 sentences establishing a clear Exit Ticket terminal tracking activity.",
                    "M_Extended_Learning": "4-5 sentences providing remediation pathways and specialized independent task expansion plans.",
                    "Sessions_Flow": [
                        {
                            "session_number": 1,
                            "teacher_action": "2-3 precise sentences detailing instructional methods and guidance.",
                            "student_task": "2-3 precise sentences mapping student activities and interactive localized deliverables."
                        }
                    ]
                }
                
                # Dynamic safe prompt building
                prompt = "Role: Act as Jenalyn, a Junior and Senior High School Specialist. Produce a 100% data-dense Master Lesson plan.\n"
                prompt += f"Subject: {subject}\n"
                prompt += f"Grade Level: {grade_level}\n"
                prompt += f"MELC Topic Focus: {melc_topic}\n"
                prompt += f"Target Lesson Duration: exactly {sessions_count} distinct sessions.\n\n"
                prompt += "Instructions: Generate a single JSON object strictly using the structural schema outline provided below. Each overview key (A through M) MUST contain 4-5 fully elaborated sentences. For 'Sessions_Flow', populate exactly " + str(sessions_count) + " objects, detailing specific teacher actions (2-3 sentences) and student tasks (2-3 sentences) per session item. Apply localized Philippine contexts when appropriate.\n\n"
                prompt += json.dumps(json_schema, indent=2)
                
                response = model.generate_content(contents=prompt)
                raw_data = json.loads(clean_json_response(response.text))
                
                # --- EXCEL COMPILATION ---
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Lesson_Plan_Overview"
                ws.views.sheetView[0].showGridLines = True
                
                # Typographic Design Tokens (Clean, friendly, high-contrast)
                f_title = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
                f_header = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
                f_bold = Font(name="Segoe UI", size=11, bold=True)
                f_regular = Font(name="Segoe UI", size=11)
                
                fill_navy = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
                fill_ice = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
                
                thin_border = Border(
                    left=Side(style='thin', color='BFBFBF'),
                    right=Side(style='thin', color='BFBFBF'),
                    top=Side(style='thin', color='BFBFBF'),
                    bottom=Side(style='thin', color='BFBFBF')
                )
                
                # Build Master Banner Block
                ws.merge_cells("A1:C1")
                ws["A1"] = f"DEPED MASTER LESSON PACKAGE: {melc_topic.upper()}"
                ws["A1"].font = f_title
                ws["A1"].fill = fill_navy
                ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
                ws.row_dimensions[1].height = 40
                
                # Row Fields Matrix Mapping
                overview_fields = [
                    ("A. References", "A_References"),
                    ("B. Declaration of AI Use", "B_AI_Declaration"),
                    ("C. Content Standard", "C_Content_Standard"),
                    ("D. Performance Standard", "D_Performance_Standard"),
                    ("E. Learning Competency", "E_Learning_Competency"),
                    ("F. Learning Objectives", "F_Learning_Objectives"),
                    ("G. Learning Context", "G_Learning_Context"),
                    ("H. Pre-Lesson Check", "H_Pre_Lesson"),
                    ("J. Learning Resources", "J_Learning_Resources"),
                    ("K. Integration Links", "K_Integration_Opportunities"),
                    ("L. Formative Assessment (Exit Ticket)", "L_Formative_Assessment"),
                    ("M. Extended Learning Activities", "M_Extended_Learning")
                ]
                
                curr_row = 3
                for label, data_key in overview_fields:
                    c1 = ws.cell(row=curr_row, column=1, value=label)
                    c1.font = f_bold
                    c1.border = thin_border
                    c1.fill = fill_ice
                    c1.alignment = Alignment(vertical="top")
                    
                    val_text = raw_data.get(data_key, "Content not populated.")
                    c2 = ws.cell(row=curr_row, column=2, value=val_text)
                    c2.font = f_regular
                    c2.border = thin_border
                    c2.alignment = Alignment(wrap_text=True, vertical="top")
                    
                    ws.row_dimensions[curr_row].height = 70
                    curr_row += 1
                
                # Build Flow Mapping Header Row
                curr_row += 2
                ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=3)
                ws.cell(row=curr_row, column=1, value="I. DETAILED LESSON SESSIONS FLOW").font = f_bold
                curr_row += 1
                
                headers = ["Session Index", "Teacher Execution Action Plan", "Student Localized Tasks"]
                for col_idx, text in enumerate(headers, 1):
                    c = ws.cell(row=curr_row, column=col_idx, value=text)
                    c.font = f_header
                    c.fill = fill_navy
                    c.border = thin_border
                    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                ws.row_dimensions[curr_row].height = 25
                
                # Append Sessions Rows Flow
                for session in raw_data.get("Sessions_Flow", []):
                    curr_row += 1
                    
                    s_id = f"Session {session.get('session_number', '')}"
                    c1 = ws.cell(row=curr_row, column=1, value=s_id)
                    c1.font = f_bold
                    c1.border = thin_border
                    c1.alignment = Alignment(horizontal="center", vertical="top")
                    
                    c2 = ws.cell(row=curr_row, column=2, value=session.get('teacher_action', ''))
                    c2.font = f_regular
                    c2.border = thin_border
                    c2.alignment = Alignment(wrap_text=True, vertical="top")
                    
                    c3 = ws.cell(row=curr_row, column=3, value=session.get('student_task', ''))
                    c3.font = f_regular
                    c3.border = thin_border
                    c3.alignment = Alignment(wrap_text=True, vertical="top")
                    
                    ws.row_dimensions[curr_row].height = 75
                
                # Autofit column configurations safely
                ws.column_dimensions['A'].width = 35
                ws.column_dimensions['B'].width = 55
                ws.column_dimensions['C'].width = 55
                
                # Save out data matrix
                excel_buffer = io.BytesIO()
                wb.save(excel_buffer)
                excel_buffer.seek(0)
                
                st.success("🎉 Master Lesson Plan Matrix Generated Successfully!")
                st.download_button(
                    label="📥 Download Master_Lesson_Plan.xlsx",
                    data=excel_buffer,
                    file_name=f"Master_Plan_{melc_topic.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"❌ Structural compiling failure encountered: {str(e)}")
