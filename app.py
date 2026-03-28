import os
import tempfile
import streamlit as st
from processor import process_folder

st.set_page_config(page_title="MTR ↔ CMC BOL Mapper", layout="centered")

st.title("MTR ↔ CMC BOL Mapper")
st.caption("Upload a folder containing .msg files, process them, and download the Excel outputs.")

uploaded_files = st.file_uploader(
    "Choose a folder with MSG files",
    type=["msg"],
    accept_multiple_files="directory",
)

if uploaded_files:
    st.info(f"{len(uploaded_files)} .msg file(s) ready to process.")

    progress_bar = st.progress(0, text="Waiting to start...")
    status_placeholder = st.empty()
    summary_placeholder = st.empty()

    with tempfile.TemporaryDirectory() as tmpdir:
        input_folder = os.path.join(tmpdir, "input")
        os.makedirs(input_folder, exist_ok=True)

        for file in uploaded_files:
            file_path = os.path.join(input_folder, file.name)
            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            with open(file_path, "wb") as f:
                f.write(file.getbuffer())

        def progress_callback(current, total, message):
            fraction = 0 if total == 0 else current / total
            progress_bar.progress(min(fraction, 1.0), text=message)
            status_placeholder.write(message)

        try:
            with st.spinner("Processing files..."):
                map_path, issues_path, summary = process_folder(
                    input_folder,
                    progress_callback=progress_callback
                )

            progress_bar.progress(1.0, text="Completed")
            status_placeholder.success("Processing complete.")

            c1, c2, c3 = st.columns(3)
            c1.metric("MSG files", summary["total_msg_files"])
            c2.metric("Matches found", summary["matches"])
            c3.metric("Issues logged", summary["issues"])

            with open(map_path, "rb") as f:
                st.download_button(
                    label="Download BOL Map",
                    data=f,
                    file_name="MTR_to_CMC_BOL_Map.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            with open(issues_path, "rb") as f:
                st.download_button(
                    label="Download Issues",
                    data=f,
                    file_name="MTR_to_CMC_BOL_Issues.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        except Exception as e:
            progress_bar.progress(0, text="Failed")
            status_placeholder.error(f"Error: {e}")