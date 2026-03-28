import os
import tempfile
import pandas as pd
import streamlit as st
from processor import process_folder

st.set_page_config(
    page_title="MTR ↔ CMC BOL Mapper",
    page_icon="📄",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
        .block-container {padding-top: 2rem; padding-bottom: 2rem; max-width: 900px;}
        .stDownloadButton button, .stButton button {width: 100%;}
        .small-note {color: #6b7280; font-size: 0.92rem;}
        .section-card {
            border: 1px solid rgba(120,120,120,.18);
            border-radius: 14px;
            padding: 1rem 1rem .25rem 1rem;
            margin-bottom: 1rem;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("📄 MTR ↔ CMC BOL Mapper")
st.caption("Upload a folder of Outlook .msg files, extract the ZIP attachments, scan the PDFs, and download the BOL mapping.")

with st.container():
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    uploaded_files = st.file_uploader(
        "Choose a folder with MSG files",
        type=["msg"],
        accept_multiple_files="directory",
        help="Pick the folder that contains your .msg files. Subfolders are supported.",
    )
    st.markdown(
        '<p class="small-note">Tip: the app copies the uploaded files to a temporary working folder before processing.</p>',
        unsafe_allow_html=True,
    )
    st.markdown("</div>", unsafe_allow_html=True)

if uploaded_files:
    st.info(f"{len(uploaded_files)} .msg file(s) ready to process.")

    progress_bar = st.progress(0, text="Waiting to start...")
    status_placeholder = st.empty()

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

            st.divider()
            st.subheader("Summary")

            c1, c2, c3 = st.columns(3)
            c1.metric("MSG files", summary["total_msg_files"])
            c2.metric("Matches found", summary["matches"])
            c3.metric("Issues logged", summary["issues"])

            if summary["total_msg_files"] == 0:
                st.warning("No .msg files were found in the uploaded folder.")
            elif summary["matches"] == 0:
                st.warning("The app processed the uploaded files but did not find any MTR ↔ CMC matches.")
            if summary["issues"] > 0:
                st.warning("Some files had issues. Review the Issues export for details.")

            st.divider()
            st.subheader("Preview")

            map_df = pd.read_excel(map_path)
            issues_df = pd.read_excel(issues_path)

            preview_tab1, preview_tab2 = st.tabs(["Matches preview", "Issues preview"])

            with preview_tab1:
                if map_df.empty:
                    st.info("No matches to preview.")
                else:
                    st.dataframe(map_df.head(20), use_container_width=True, hide_index=True)

            with preview_tab2:
                if issues_df.empty:
                    st.success("No issues logged.")
                else:
                    st.dataframe(issues_df.head(20), use_container_width=True, hide_index=True)

            st.divider()
            st.subheader("Downloads")

            d1, d2 = st.columns(2)
            with d1:
                with open(map_path, "rb") as f:
                    st.download_button(
                        label="Download BOL Map",
                        data=f,
                        file_name="MTR_to_CMC_BOL_Map.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
            with d2:
                with open(issues_path, "rb") as f:
                    st.download_button(
                        label="Download Issues",
                        data=f,
                        file_name="MTR_to_CMC_BOL_Issues.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

            if st.button("Reset app", use_container_width=True):
                st.rerun()

        except Exception as e:
            progress_bar.progress(0, text="Failed")
            status_placeholder.error(f"Error: {e}")
else:
    st.markdown("### How it works")
    st.markdown(
        """
        1. Upload a folder containing `.msg` files.  
        2. The app reads each email, finds ZIP attachments, and scans the PDFs inside.  
        3. It extracts the MTR BOL from the filename and the CMC BOL from the PDF text.  
        4. Download the Excel outputs.
        """
    )
