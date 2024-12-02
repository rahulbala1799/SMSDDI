import streamlit as st
import os
import pandas as pd

# Page Configuration
st.set_page_config(
    page_title="SMS DDI and Plan Accrual Journals Generator",
    page_icon="📤",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Sidebar Navigation
st.sidebar.title("Navigation")
menu = st.sidebar.radio("Go to:", ["Home", "Upload File"])

# Initialize session state for file handling
if "uploaded_file_path" not in st.session_state:
    st.session_state["uploaded_file_path"] = None

# Custom CSS for header
st.markdown(
    """
    <style>
        .header-container {
            text-align: center;
            margin-bottom: 20px;
        }
        .phorest-title {
            font-size: 3em;
            font-weight: bold;
            color: #007bff; /* Blue */
        }
        .app-name {
            font-size: 1.5em;
            color: #343a40; /* Dark Grey */
        }
        .main-content {
            margin-top: 20px;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# Display Header
st.markdown(
    """
    <div class="header-container">
        <div class="phorest-title">PHOREST</div>
        <div class="app-name">SMS DDI AND PLAN ACCRUAL JOURNALS GENERATOR</div>
    </div>
    """,
    unsafe_allow_html=True,
)

# Home Page
if menu == "Home":
    st.write("### Welcome to the File Upload App!")
    st.write(
        "This tool helps you generate SMS DDI and Plan Accrual Journals with ease. "
        "Navigate to 'Upload File' to get started."
    )

# Upload File Page
elif menu == "Upload File":
    st.write("## Upload Your Excel File")
    uploaded_file = st.file_uploader("Upload an Excel file for processing:", type=["xlsx"])

    if uploaded_file:
        # Save the uploaded file to a temporary location
        file_path = os.path.join(os.getcwd(), "uploaded_file.xlsx")
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.session_state["uploaded_file_path"] = file_path

        # Confirmation and preview
        st.success("File uploaded successfully and saved for processing!")
        st.write("### File Preview:")
        try:
            # Preview first few rows of the uploaded file
            uploaded_df = pd.read_excel(file_path)
            st.dataframe(uploaded_df.head())
        except Exception as e:
            st.error(f"Error reading the file: {e}")

    # Proceed Button
    if st.session_state["uploaded_file_path"]:
        st.write("### Ready to Process?")
        if st.button("Proceed to Processing"):
            st.write("Navigate to '1Processing.py' in the Pages menu to process your file.")
