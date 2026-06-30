import streamlit as st
import google.generativeai as genai

# Page Configuration
st.set_page_config(
    page_title="DepEd ILAW Lesson Exemplar Generator", 
    page_icon="📝", 
    layout="wide"
)

# Header Section
st.title("📝 DepEd ILAW Lesson Exemplar Generator")
st.caption("A smart tool for Filipino Master Teachers to generate localized, inclusive lessons based on the latest DepEd ILAW Framework.")
st.divider()

# Sidebar: Configuration & API Key Input
with st.sidebar:
    st.header("⚙️ Configuration")
    st.markdown("To keep this app 100% free, input your own Gemini API Key below. *(You can get one for free at aistudio.google.com)*")
    api_key = st.text_input("Gemini API Key", type="password")
    
    st.info("💡 **Why ILAW?** It stands for Intentions, Learning Experience, Assessing Learning, and Ways Forward.")

# Main Form Components
col1, col2 = st.columns(2)

with col1:
    st.subheader("📋 Class Information")
    subject = st.text_input("Subject & Grade Level", placeholder="e.g., Grade 11 Earth and Life Science")
    melc = st.text_area("Most Essential Learning Competency (MELC) or Topic", placeholder="e.g., Explain how the movement of plates leads to the formation of folds and faults.")

with col2:
    st.subheader("🏫 Classroom Setting & Context")
    class_size = st.slider("Estimated Class Size", min_value=10, max_value=80, value=55)
    resources = st.multiselect(
        "Available Classroom Resources",
        ["Blackboard & Chalk/Markers", "Printed Worksheets/IMs", "Projector / TV", "Science Lab Equipment", "Learner Smartphones / Internet"],
        default=["Blackboard & Chalk/Markers", "Printed Worksheets/IMs"]
    )
    custom_resources = st.text_input("Other local materials available:", placeholder="e.g., recycled cardboard, plastic bottles, old newspapers")

# Combine inputs into contextual string
resource_string = ", ".join(resources)
if custom_resources:
    resource_string += f", plus localized items: {custom_resources}"

# Generate Button Logic
if st.button("✨ Generate ILAW Lesson Exemplar", type="primary"):
    if not api_key:
        st.error("⚠️ Please enter your Gemini API Key in the sidebar to run the generator!")
    elif not subject or not melc:
        st.warning("⚠️ Please provide both the Subject/Grade Level and the MELC/Topic.")
    else:
        with st.spinner("🔄 Master Teacher AI is crafting your lesson... Please wait..."):
            try:
                # Initialize Gemini Model
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel('gemini-pro')
                
                # Construct System Prompt with ILAW Framework instructions
                prompt = f"""
                You are an expert Master Teacher and Instructional Coach for the Philippine Department of Education (DepEd). 
                Generate a highly detailed, comprehensive, and contextualized Lesson Exemplar using the official ILAW framework.

                --- USER INPUTS ---
                Subject & Grade Level: {subject}
                MELC/Topic: {melc}
                Class Size: {class_size} students
                Available Resources: {resource_string}

                --- INSTRUCTIONS & OUTPUT FORMAT ---
                The lesson plan must be completely designed to handle the user's specific class size and resource limits. 
                Structure the output with professional formatting using clear markdown headers exactly as follows:

                ## 📝 ILAW LESSON EXEMPLAR FOR {subject.upper()}
                **Topic:** {melc}  
                **Context:** Class size of {class_size} learners with access to {resource_string}.

                ### 1. INTENTIONS (What will they learn?)
                * Define 2-3 explicit, realistic, and measurable learning objectives. Aligned to cognitive, psychomotor, and affective domains if appropriate.
                * Consider student well-being and gender-responsive teaching.

                ### 2. LEARNING EXPERIENCE (How will they learn?)
                * Design active learning activities optimized for {class_size} students using only {resource_string}.
                * If class size is large (>40), include clear grouping strategies (e.g., row-partners, triad checks) to manage noise and movement.
                * Emphasize the use of local, low-cost, or improvised materials to ensure inclusivity. No child must be excluded due to economic constraints.

                ### 3. ASSESSING LEARNING (How will we know they learned?)
                * Provide a concrete formative assessment task matching the intentions.
                * Explicitly state how the teacher will provide immediate, actionable feedback to a large group of students.
                * Provide accommodations for diverse learners (e.g., remedial scaffolding).

                ### 4. WAYS FORWARD (Where do we go from here?)
                * Provide reflection prompts for the teacher.
                * Include 1 extension or enrichment activity using zero-cost materials, and a clear remediation strategy for struggling learners.
                """
                
                # Call AI API
                response = model.generate_content(prompt)
                
                # Display Results
                st.success("🎉 Lesson Exemplar Generated Successfully!")
                st.markdown(response.text)
                
                # Download Button option
                st.download_button(
                    label="📥 Download Lesson Plan as Text",
                    data=response.text,
                    file_name=f"ILAW_Lesson_{subject.replace(' ', '_')}.txt",
                    mime="text/plain"
                )
                
            except Exception as e:
                st.error(f"❌ An error occurred: {str(e)}")
