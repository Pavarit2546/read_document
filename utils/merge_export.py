from docxtpl import DocxTemplate
import json
import io
import tempfile
import os
from typing import Optional

# กำหนดเส้นทางไปยังไฟล์ Template DOCX (ต้องอยู่ในระดับเดียวกับ app.py หรือ Docker build)
TEMPLATE_PATH = "ISMS-F-INF-036_template.docx"

def merge_and_export_docx(user_data_json: str, template_bytes: Optional[bytes] = None) -> bytes:
    """
    ฟังก์ชันสำหรับ Merge ข้อมูล JSON ลงใน Template DOCX และส่งเป็น Byte Stream
    :param user_data_json: ข้อมูลสำหรับ Merge ในรูปแบบ JSON string
    :param template_bytes: (Optional) ถ้าให้มา จะใช้เป็น template แทน TEMPLATE_PATH
    :return: Byte Stream ของไฟล์ DOCX ที่สมบูรณ์
    """
    temp_template_path = None
    try:
        data = json.loads(user_data_json)
        context = data

        # ถ้ามี template_bytes ให้สร้าง temp file และใช้แทน TEMPLATE_PATH
        template_path_to_use = TEMPLATE_PATH
        if template_bytes:
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
            tmp.write(template_bytes)
            tmp.close()
            temp_template_path = tmp.name
            template_path_to_use = temp_template_path

        # โหลด Template
        doc = DocxTemplate(template_path_to_use)

        # ทำการ Merge ข้อมูล
        doc.render(context)

        # บันทึกไฟล์ที่ Merge แล้วลงใน Memory (Buffer)
        output_buffer = io.BytesIO()
        doc.save(output_buffer)
        output_buffer.seek(0)

        return output_buffer.getvalue()

    except Exception as e:
        print(f"Merge Error: {str(e)}")
        raise Exception(f"Merge failed: {str(e)}")
    finally:
        # cleanup temp template if created
        if temp_template_path and os.path.exists(temp_template_path):
            try:
                os.unlink(temp_template_path)
            except Exception:
                pass