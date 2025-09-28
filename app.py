import streamlit as st
import requests

BACKEND_URL = "http://localhost:8000/processing/process-image/"

st.set_page_config(page_title="Exam Evaluation", layout="centered")
st.title("📄 Exam Evaluation System")

uploaded_file = st.file_uploader("Upload a PDF file", type=["pdf"])

if uploaded_file:
    st.success(f"Uploaded file: {uploaded_file.name}")

    if st.button("Process File"):
        with st.spinner("Processing..."):
            files = {"file": (uploaded_file.name, uploaded_file, "application/pdf")}
            try:
                response = requests.post(BACKEND_URL, files=files)
                if response.status_code == 200:
                    result = response.json()
                    st.success(f"✅ {result.get('message', 'Files processed successfully!')}")

                else:
                    st.error(f"❌ Error {response.status_code}: {response.text}")
            except Exception as e:
                st.error(f"⚠️ Failed to connect to backend: {e}")

