from flask import Flask, request, jsonify, send_file
import requests # ต้องมี requests ใน requirements.txt สำหรับการดาวน์โหลด URL
import io
import json as _json
import tempfile
import os

# ตรวจสอบให้แน่ใจว่า import path ถูกต้องตามโครงสร้าง utils/
from utils.merge_export import merge_and_export_docx
from utils.read_doc import read_docx_content 

app = Flask(__name__)

def find_url(obj):
    """Recursive helper to find a http/https URL inside nested JSON/structures."""
    if isinstance(obj, str):
        if obj.startswith('http://') or obj.startswith('https://'):
            return obj
        return None
    if isinstance(obj, dict):
        for key in ('file_url', 'fileUrl', 'url', 'docx_url', 'document_url', 'template_url', 'templateUrl'):
            if key in obj and isinstance(obj[key], str):
                if obj[key].startswith('http://') or obj[key].startswith('https://'):
                    return obj[key]
        for v in obj.values():
            res = find_url(v)
            if res:
                return res
    if isinstance(obj, list):
        for item in obj:
            res = find_url(item)
            if res:
                return res
    return None

def download_url_bytes(file_url: str, timeout: int = 15) -> bytes:
    """Download URL and return bytes; raise ValueError on non-200 or wrong scheme."""
    if not (file_url.startswith('http://') or file_url.startswith('https://')):
        raise ValueError("Invalid URL scheme. Must be http or https")
    resp = requests.get(file_url, timeout=timeout)
    resp.raise_for_status()
    return resp.content

# --- ENDPOINT 1: อ่านเนื้อหาจาก DOCX ---
@app.route('/read-docx', methods=['POST'])
def read_docx():
    """
    Endpoint สำหรับรับไฟล์ DOCX (อัพโหลดหรือเป็น URL) และส่งคืนเนื้อหาข้อความทั้งหมดในรูปแบบ JSON
    
    รองรับ:
    1. multipart/form-data: field name = 'file' (ไฟล์อัปโหลดตรง)
    2. JSON/Form Data: ค้นหา URL ลึกๆ ใน body (รองรับ input ที่ซับซ้อนจาก Workflow Node)
    """
    docx_bytes = None

    # 1. ตรวจสอบไฟล์ที่ถูกอัปโหลดโดยตรง (multipart/form-data)
    if 'file' in request.files:
        file = request.files['file']
        if file.filename == '':
            return jsonify({"status": "error", "message": "No selected file"}), 400
        if not file.filename.lower().endswith('.docx'):
            return jsonify({"status": "error", "message": "Invalid file format. Must be .docx"}), 400
        docx_bytes = file.read()
    else:
        # 2. ค้นหา URL ใน Request Body (JSON/Form/Query)
        file_url = None
        raw_body = request.get_data(as_text=True)

        # 2a) check query string
        file_url = request.args.get('file_url') or request.args.get('url')

        # 2b) check form
        if not file_url:
            file_url = request.form.get('file_url') or request.form.get('url')

        # 2c) check JSON payload robustly
        if not file_url and request.is_json:
            try:
                j = request.get_json(silent=True, force=True)
                file_url = find_url(j)
            except Exception:
                file_url = None

        # 2d) Fallback: try to parse raw body as JSON 
        if not file_url:
            try:
                parsed = _json.loads(raw_body) if raw_body else None
                file_url = find_url(parsed) if parsed else None
            except Exception:
                file_url = None

        if not file_url:
            return jsonify({"status": "error", "message": "No file part or file_url found in the request."}), 400

        try:
            resp = requests.get(file_url, timeout=15)
            if resp.status_code != 200:
                return jsonify({"status": "error", "message": f"Failed to fetch file. HTTP {resp.status_code}"}), 400

            content_type = resp.headers.get('Content-Type', '')
            if not ('wordprocessingml' in content_type or 'officedocument' in content_type or file_url.lower().endswith('.docx')):
                return jsonify({"status": "error", "message": "Fetched file does not appear to be a .docx"}), 400

            docx_bytes = resp.content

        except requests.RequestException as e:
            return jsonify({"status": "error", "message": f"Failed to fetch file: {str(e)}"}), 400

    if not docx_bytes:
        return jsonify({"status": "error", "message": "Document bytes could not be determined."}), 500

    result_json = read_docx_content(docx_bytes)
    return result_json, 200, {'Content-Type': 'application/json; charset=utf-8'}


# --- ENDPOINT 2: Merge ข้อมูลและ Export DOCX ---
@app.route('/merge-docx', methods=['POST'])
def merge_and_export():
    """
    Endpoint สำหรับรับข้อมูล JSON และ Merge ลงใน Template
    รองรับ:
      - body JSON (user data) และ optional template upload field 'template'
      - หรือ template_url ใน JSON/form/query
    """
    try:
        # รับข้อมูล JSON จาก Request Body (body มี JSON string ของข้อมูล merge)
        # รองรับทั้ง application/json และ raw body JSON
        user_data_json = None
        template_bytes = None
        # 1) ถ้ามีการอัพโหลด template มาเป็น multipart file 'template'
        if 'template' in request.files:
            f = request.files['template']
            if f.filename and f.filename.lower().endswith('.docx'):
                template_bytes = f.read()
            else:
                return jsonify({"status": "error", "message": "Template must be a .docx file"}), 400

        # 2) หา template_url ใน query/form/json
        if not template_bytes:
            # search in query/form
            template_url = request.args.get('template_url') or request.args.get('templateUrl') or request.args.get('template')
            if not template_url:
                template_url = request.form.get('template_url') or request.form.get('templateUrl') or request.form.get('template')

            # check JSON body for template_url and also capture user data
            if request.is_json:
                try:
                    body_json = request.get_json(silent=True, force=True) or {}
                    # If body contains template url nested, find it
                    if not template_url:
                        template_url = find_url(body_json)
                    # If body contains user data as object, serialize back to string for merge function
                    # If the whole body is the user data, use it
                    user_data_json = _json.dumps(body_json, ensure_ascii=False)
                except Exception:
                    body_json = None

            # Fallback: raw body text (if not parsed above)
            if not user_data_json:
                raw_body = request.get_data(as_text=True)
                user_data_json = raw_body

            if template_url and not template_bytes:
                try:
                    template_bytes = download_url_bytes(template_url, timeout=15)
                except Exception as e:
                    return jsonify({"status": "error", "message": f"Failed to fetch template: {str(e)}"}), 400

        else:
            # template provided as file; now get user JSON content
            if request.is_json:
                user_data_json = _json.dumps(request.get_json(silent=True, force=True) or {}, ensure_ascii=False)
            else:
                user_data_json = request.get_data(as_text=True)

        if not user_data_json:
            return jsonify({"status": "error", "message": "No JSON data provided for merge"}), 400

        # เรียกใช้ฟังก์ชัน Merge ข้อมูล โดยส่ง template_bytes (อาจเป็น None ถ้าใช้ TEMPLATE_PATH ใน utils)
        docx_bytes = merge_and_export_docx(user_data_json, template_bytes=template_bytes)

        # ส่งไฟล์ DOCX ที่ Generate แล้วกลับไปเป็น Byte Stream
        return send_file(
            io.BytesIO(docx_bytes),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name='generated_document.docx'
        )

    except Exception as e:
        print(f"Merge and Export Error: {str(e)}")
        return jsonify({"status": "error", "message": f"Internal server error: {str(e)}"}), 500

if __name__ == '__main__':
    # รัน Flask server ที่ Port 3000
    app.run(debug=True, host='0.0.0.0', port=3000)