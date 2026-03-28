import streamlit as st
import tempfile
import os
from processor import process_folder

st.set_page_config(page_title="MTR ↔ CMC BOL Mapper", layout="centered")

st.title("MTR ↔ CMC BOL Mapper")
st.write("Upload a folder containing `.msg` files, then download the Excel outputs.")

uploaded_files = st.file_uploader(
    "Choose a folder with MSG files",
    type=["msg"],
    accept_multiple_files="directory",
)

if uploaded_files:
    st.write(f"Files received: {len(uploaded_files)}")

    with tempfile.TemporaryDirectory() as tmpdir:
        input_folder = os.path.join(tmpdir, "input")
        os.makedirs(input_folder, exist_ok=True)

        for file in uploaded_files:
            file_path = os.path.join(input_folder, file.name)
            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            with open(file_path, "wb") as f:
                f.write(file.getbuffer())

        with st.spinner("Processing files..."):
            try:
                map_path, issues_path = process_folder(input_folder)
                st.success("Processing complete.")

                with open(map_path, "rb") as f:
                    st.download_button(
                        label="Download BOL Map",
                        data=f,
                        file_name="MTR_to_CMC_BOL_Map.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                with open(issues_path, "rb") as f:
                    st.download_button(
                        label="Download Issues",
                        data=f,
                        file_name="MTR_to_CMC_BOL_Issues.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

            except Exception as e:
                st.error(f"Error: {e}")