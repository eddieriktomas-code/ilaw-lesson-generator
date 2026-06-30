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

# Helper Function: Clean JSON String from Gemini
def clean_json_response(text):
    text = text.strip()
    if text.startswith("
http://googleusercontent.com/immersive_entry_chip/0
http://googleusercontent.com/immersive_entry_chip/1
