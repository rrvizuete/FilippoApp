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


def process_folder(folder_path, progress_callback=None):
    results = []
    issues = []

    msg_files = []
    for root, _, files in os.walk(folder_path):
        for f in files:
            if f.lower().endswith(".msg"):
                msg_files.append(os.path.join(root, f))

    total_files = len(msg_files)

    if progress_callback:
        progress_callback(0, total_files, "Starting...")

    for idx, msg_path in enumerate(msg_files, start=1):
        file_name = os.path.basename(msg_path)

        if progress_callback:
            progress_callback(idx - 1, total_files, f"Processing {file_name}")

        mtr_match = re.search(r"BOL#\s*(\d+)", file_name)
        if not mtr_match:
            issues.append([file_name, "MTR BOL not found"])
            continue

        mtr_bol = mtr_match.group(1)

        try:
            msg = extract_msg.Message(msg_path)
            attachments = msg.attachments
            zip_found = False

            for att in attachments:
                name = (att.longFilename or att.shortFilename or "")
                if not name.lower().endswith(".zip"):
                    continue

                zip_found = True
                tmp_zip_path = None

                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".zip") as tmp_zip:
                        tmp_zip.write(att.data)
                        tmp_zip_path = tmp_zip.name

                    with zipfile.ZipFile(tmp_zip_path, "r") as zip_ref:
                        for pdf_name in zip_ref.namelist():
                            if not pdf_name.lower().endswith(".pdf"):
                                continue

                            try:
                                pdf_bytes = zip_ref.read(pdf_name)
                                cmc_bol = extract_cmc_from_pdf(pdf_bytes)

                                if cmc_bol:
                                    results.append([mtr_bol, cmc_bol, file_name, pdf_name])
                                else:
                                    issues.append([file_name, f"No CMC in {pdf_name}"])
                            except Exception as e:
                                issues.append([file_name, f"Error reading {pdf_name}: {e}"])

                except Exception as e:
                    issues.append([file_name, f"ZIP error ({name}): {e}"])

                finally:
                    if tmp_zip_path and os.path.exists(tmp_zip_path):
                        try:
                            os.remove(tmp_zip_path)
                        except Exception:
                            pass

            if not zip_found:
                issues.append([file_name, "No ZIP attachment found"])

        except Exception as e:
            issues.append([file_name, str(e)])

        if progress_callback:
            progress_callback(idx, total_files, f"Processed {file_name}")

    map_path = os.path.join(folder_path, "MTR_to_CMC_BOL_Map.xlsx")
    issues_path = os.path.join(folder_path, "MTR_to_CMC_BOL_Issues.xlsx")

    wb = Workbook()
    ws = wb.active
    ws.title = "BOL Map"
    ws.append(["MTR BOL#", "CMC BOL#", "MSG File", "PDF File"])
    for row in results:
        ws.append(row)
    wb.save(map_path)

    wb2 = Workbook()
    ws2 = wb2.active
    ws2.title = "Issues"
    ws2.append(["File", "Issue"])
    for row in issues:
        ws2.append(row)
    wb2.save(issues_path)

    summary = {
        "total_msg_files": total_files,
        "matches": len(results),
        "issues": len(issues),
    }

    return map_path, issues_path, summary
