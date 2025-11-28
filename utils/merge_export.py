from docxtpl import DocxTemplate
import json
import io
import tempfile
import os
import requests
from typing import Optional, Dict, Any


# กำหนด Default Template Path ในกรณีที่ไม่มี template มาใน request
DEFAULT_TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), '..', 'templates', 'default_template.docx')

def merge_and_export_docx(user_data_json: str, template_bytes: Optional[bytes] = None) -> bytes:
    """
    ฟังก์ชันสำหรับ Merge ข้อมูล JSON ลงใน Template DOCX และส่งเป็น Byte Stream
    :param user_data_json: ข้อมูลสำหรับ Merge ในรูปแบบ JSON string
    :param template_bytes: (Optional) ถ้าให้มา จะใช้เป็น template แทน TEMPLATE_PATH
    :return: Byte Stream ของไฟล์ DOCX ที่สมบูรณ์
    """
    try:
        data = json.loads(user_data_json)
        context = data
        print(f"Merge Context: {context}")

        template_to_use = None
        if 'template_url' in data:
            print(f"Downloading template from: {data['template_url']}")
            response = requests.get(data['template_url'])
            response.raise_for_status()  # Raise an exception for bad status codes
            template_to_use = io.BytesIO(response.content)
        elif template_bytes:
            template_to_use = io.BytesIO(template_bytes)

        # หากไม่มี template_url หรือ template_bytes ให้ใช้ default template
        if not template_to_use:
            print(f"No template provided, using default template at: {DEFAULT_TEMPLATE_PATH}")
            template_to_use = DEFAULT_TEMPLATE_PATH

        # ลบ template_url ออกจาก context ก่อนทำการ render เพื่อไม่ให้แสดงในเอกสารผลลัพธ์
        # และป้องกันปัญหา lazy rendering ของ docxtpl
        # if 'template_url' in context:
        #     del context['template_url']

        # --- เพิ่มส่วนจัดการ Checkbox ---
        # วนลูปใน context เพื่อแปลงค่า boolean เป็นสัญลักษณ์ checkbox
        for key, value in context.items():
            if isinstance(value, bool):
                # ถ้าเป็น True ให้ใช้ '☒', ถ้าเป็น False ให้ใช้ '☐'
                context[key] = '☒' if value else '☐'
        # --------------------------------

        # โหลด Template
        doc = DocxTemplate(template_to_use)

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
    