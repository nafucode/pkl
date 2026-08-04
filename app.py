from __future__ import annotations

import io
import base64
import os
import re
import tempfile
import zipfile
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.utils import secure_filename

from translator import ConversionError, convert_workbook


BASE_DIR = Path(__file__).resolve().parent
ALLOWED_EXTENSIONS = {".xlsx"}

app = Flask(
    __name__,
    template_folder=str(BASE_DIR / "templates"),
    static_folder=str(BASE_DIR / "static"),
)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024


def is_allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


def package_code(filename: str) -> str:
    safe_name = secure_filename(filename) or "packing-list.xlsx"
    match = re.search(r"XFJ\d{4}-\d+", filename, re.IGNORECASE)
    return match.group(0).upper() if match else (Path(safe_name).stem or "packing-list")


def make_archive(result, code: str) -> io.BytesIO:
    archive = io.BytesIO()
    with zipfile.ZipFile(archive, "w", compression=zipfile.ZIP_DEFLATED) as zipped:
        zipped.write(result.pdf_path, f"{code}-English.pdf")
        zipped.write(result.excel_path, f"{code}-English.xlsx")
        zipped.write(result.report_path, f"{code}-translation-check-report.txt")
    archive.seek(0)
    return archive


def data_url(mimetype: str, content: bytes) -> str:
    return f"data:{mimetype};base64," + base64.b64encode(content).decode("ascii")


def report_summary(report_text: str, residual_chinese_count: int) -> dict[str, int]:
    quality_match = re.search(r"WARNING:\s*(\d+)\s+data quality issue", report_text)
    return {
        "residualChinese": residual_chinese_count,
        "qualityIssues": int(quality_match.group(1)) if quality_match else 0,
    }


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/convert")
def convert():
    uploaded = request.files.get("packing_list")
    if not uploaded or not uploaded.filename:
        return jsonify({"error": "请选择一个 Excel 装箱单文件。"}), 400
    if not is_allowed_file(uploaded.filename):
        return jsonify({"error": "目前只支持 .xlsx 文件。"}), 400

    safe_name = secure_filename(uploaded.filename) or "packing-list.xlsx"
    code = package_code(uploaded.filename)

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_root = Path(temp_dir)
        input_path = temp_root / safe_name
        output_dir = temp_root / "generated"
        uploaded.save(input_path)

        try:
            result = convert_workbook(input_path, output_dir)
        except ConversionError as exc:
            return jsonify({"error": str(exc)}), 400

        archive = make_archive(result, code)

    response = send_file(
        archive,
        mimetype="application/zip",
        as_attachment=True,
        download_name=f"{code}-English.zip",
    )
    response.headers["X-Package-Count"] = str(result.package_count)
    response.headers["X-Page-Count"] = str(result.page_count)
    response.headers["X-Residual-Chinese-Count"] = str(result.residual_chinese_count)
    return response


@app.post("/preview")
def preview():
    uploaded = request.files.get("packing_list")
    if not uploaded or not uploaded.filename:
        return jsonify({"error": "请选择一个 Excel 装箱单文件。"}), 400
    if not is_allowed_file(uploaded.filename):
        return jsonify({"error": "目前只支持 .xlsx 文件。"}), 400

    safe_name = secure_filename(uploaded.filename) or "packing-list.xlsx"
    code = package_code(uploaded.filename)

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_root = Path(temp_dir)
        input_path = temp_root / safe_name
        output_dir = temp_root / "generated"
        uploaded.save(input_path)

        try:
            result = convert_workbook(input_path, output_dir)
        except ConversionError as exc:
            return jsonify({"error": str(exc)}), 400

        archive = make_archive(result, code)
        pdf_bytes = result.pdf_path.read_bytes()
        excel_bytes = result.excel_path.read_bytes()
        report_text = result.report_path.read_text(encoding="utf-8")

        return jsonify(
            {
                "filename": f"{code}-English.zip",
                "pdfFilename": f"{code}-English.pdf",
                "excelFilename": f"{code}-English.xlsx",
                "packageCount": result.package_count,
                "pageCount": result.page_count,
                "residualChineseCount": result.residual_chinese_count,
                "summary": report_summary(report_text, result.residual_chinese_count),
                "reportText": report_text,
                "pdfDataUrl": data_url("application/pdf", pdf_bytes),
                "excelDataUrl": data_url(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    excel_bytes,
                ),
                "zipDataUrl": data_url("application/zip", archive.getvalue()),
            }
        )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=True)
