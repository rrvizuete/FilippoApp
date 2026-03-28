import streamlit as st
import tempfile
import os
from processor import process_folder

st.title("MTR ↔ CMC BOL Mapper")

st.write("Upload MSG files to process")

uploaded_files = st.file_uploader(
    "Select .msg files",
    type=["msg"],
    accept_multiple_files=True
)

if uploaded_files:
    with tempfile.TemporaryDirectory() as tmpdir:
        input_folder = os.path.join(tmpdir, "input")
        os.makedirs(input_folder, exist_ok=True)

        # Save uploaded files
        for file in uploaded_files:
            file_path = os.path.join(input_folder, file.name)
            with open(file_path, "wb") as f:
                f.write(file.getbuffer())

        st.info("Processing files...")

        try:
            map_path, issues_path = process_folder(input_folder)

            st.success("Processing complete!")

            with open(map_path, "rb") as f:
                st.download_button(
                    label="Download BOL Map",
                    data=f,
                    file_name="MTR_to_CMC_BOL_Map.xlsx"
                )

            with open(issues_path, "rb") as f:
                st.download_button(
                    label="Download Issues",
                    data=f,
                    file_name="MTR_to_CMC_BOL_Issues.xlsx"
                )

        except Exception as e:
            st.error(f"Error: {e}")
        