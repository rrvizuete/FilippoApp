import os
import zipfile
import tempfile
import re
import extract_msg
import fitz  # pymupdf
from openpyxl import Workbook

def extract_cmc_from_pdf(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    text = ""
    for page in doc:
        text += page.get_text()

    match = re.search(r"BOL#:\s*(\d+)", text)
    if match:
        return match.group(1)
    return None

def process_folder(folder_path):
    results = []
    issues = []

    for file_name in os.listdir(folder_path):
        if not file_name.endswith(".msg"):
            continue

        mtr_match = re.search(r"BOL#\s*(\d+)", file_name)
        if not mtr_match:
            issues.append([file_name, "MTR BOL not found"])
            continue

        mtr_bol = mtr_match.group(1)
        msg_path = os.path.join(folder_path, file_name)

        try:
            msg = extract_msg.Message(msg_path)
            attachments = msg.attachments

            zip_found = False

            for att in attachments:
                if att.longFilename.endswith(".zip"):
                    zip_found = True

                    with tempfile.NamedTemporaryFile(delete=False, suffix=".zip") as tmp_zip:
                        tmp_zip.write(att.data)
                        tmp_zip_path = tmp_zip.name

                    with zipfile.ZipFile(tmp_zip_path, 'r') as zip_ref:
                        for pdf_name in zip_ref.namelist():
                            if pdf_name.endswith(".pdf"):
                                pdf_bytes = zip_ref.read(pdf_name)
                                cmc_bol = extract_cmc_from_pdf(pdf_bytes)

                                if cmc_bol:
                                    results.append([mtr_bol, cmc_bol])
                                else:
                                    issues.append([file_name, f"No CMC in {pdf_name}"])

        except Exception as e:
            issues.append([file_name, str(e)])

    # Create Excel outputs
    map_path = os.path.join(folder_path, "MTR_to_CMC_BOL_Map.xlsx")
    issues_path = os.path.join(folder_path, "MTR_to_CMC_BOL_Issues.xlsx")

    wb = Workbook()
    ws = wb.active
    ws.append(["MTR BOL#", "CMC BOL#"])
    for row in results:
        ws.append(row)
    wb.save(map_path)

    wb2 = Workbook()
    ws2 = wb2.active
    ws2.append(["File", "Issue"])
    for row in issues:
        ws2.append(row)
    wb2.save(issues_path)

    return map_path, issues_path