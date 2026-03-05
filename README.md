# DOCX Processing API

ระบบ API สำหรับจัดการเอกสาร Word (.docx) พัฒนาด้วย Python และ Flask รองรับการอ่านเนื้อหา, การดึงหัวข้อ และการ Merge ข้อมูลลงใน Template เอกสาร

## คุณสมบัติหลัก
- **Read DOCX**: อ่านข้อความทั้งหมดจากไฟล์ Word
- **Read Prefix**: ดึงหัวข้อหรือคำนำหน้า (Prefix) ออกจากเอกสารเพื่อใช้กับ LLM
- **Merge DOCX**: นำข้อมูล JSON ไปเติมลงใน Template Word (Mail Merge style)
- **Merge by Prefix**: เติมข้อมูลลงในเอกสารโดยอ้างอิงจากหัวข้อ (เช่น ชื่อ: .......)

## การติดตั้ง

1. ติดตั้ง Python (แนะนำเวอร์ชัน 3.8 ขึ้นไป)
2. ติดตั้ง Library ที่จำเป็น:
   ```bash
   pip install flask requests python-docx
   ```
   *(หมายเหตุ: ตรวจสอบให้แน่ใจว่าได้ติดตั้ง dependencies ที่จำเป็นสำหรับไฟล์ในโฟลเดอร์ `utils/` ด้วย)*

## การเริ่มใช้งาน

รันเซิร์ฟเวอร์ด้วยคำสั่ง:
```bash
python app.py
```
เซิร์ฟเวอร์จะทำงานที่ `http://localhost:3000`

---

## รายละเอียด API (API Documentation)

### 1. อ่านเนื้อหาจาก DOCX (`/read-docx`)
ใช้สำหรับดึงข้อความทั้งหมดจากไฟล์ Word

- **Method:** `POST`
- **Content-Type:** `multipart/form-data` หรือ `application/json`
- **Input:**
  - อัปโหลดไฟล์ผ่าน Field: `file`
  - หรือส่ง JSON: `{ "file_url": "https://example.com/doc.docx" }`
- **Response:** JSON ข้อมูลเนื้อหาในเอกสาร

### 2. ดึงหัวข้อจาก DOCX (`/read-docx-prefix`)
ใช้สำหรับดึงเฉพาะส่วนที่เป็นหัวข้อหรือคำนำหน้า เพื่อนำไปให้ AI/LLM ประมวลผลต่อ

- **Method:** `POST`
- **Input:** รองรับทั้งการอัปโหลดไฟล์ (`file`) หรือส่ง URL ผ่าน JSON
- **Response:** JSON รายการหัวข้อที่พบในเอกสาร

### 3. Merge ข้อมูลลง Template (`/merge-docx`)
ใช้สำหรับสร้างเอกสารใหม่จาก Template โดยแทนที่ตัวแปรด้วยข้อมูล JSON

- **Method:** `POST`
- **Input:**
  - **Template:** อัปโหลดไฟล์ผ่าน Field `template` หรือส่ง `template_url` ใน JSON
  - **Data:** ข้อมูล JSON ที่ต้องการ Merge (ส่งมาใน Body)
- **Response:** ไฟล์ `.docx` ที่ Generate เสร็จแล้ว (ดาวน์โหลดอัตโนมัติ)

### 4. Merge ข้อมูลแบบ Prefix (`/merge-prefix`)
ใช้สำหรับเติมข้อมูลต่อท้ายหัวข้อที่กำหนดไว้ในเอกสาร (เช่น "ชื่อ: [ข้อมูลจาก JSON]")

- **Method:** `POST`
- **Input:**
  - **Template:** อัปโหลดไฟล์ `template` หรือระบุ `filepath` ใน JSON
  - **Data:** ข้อมูล JSON ที่มี Key ตรงกับหัวข้อในเอกสาร
- **Response:** ไฟล์ `.docx` ที่เติมข้อมูลครบถ้วน

---

## ตัวอย่างการเรียกใช้งาน (cURL)

**การอ่านไฟล์ผ่าน URL:**
```bash
curl -X POST http://localhost:3000/read-docx \
     -H "Content-Type: application/json" \
     -d '{"url": "https://your-storage.com/document.docx"}'
```

**การ Merge ข้อมูลกับ Template (Multipart):**
```bash
curl -X POST http://localhost:3000/merge-docx \
     -F "template=@my_template.docx" \
     -F "data={\"name\": \"John Doe\", \"date\": \"2023-10-27\"}"
```

## โครงสร้างโปรเจกต์
```text
├── app.py                # ไฟล์หลักสำหรับรัน Flask API
├── utils/                # โฟลเดอร์เก็บ Logic การจัดการไฟล์
│   ├── read_doc.py       # Logic การอ่านข้อความ
│   ├── read_docx_prefix.py # Logic การดึงหัวข้อ
│   ├── merge_export.py   # Logic การ Merge Template
│   └── merge_by_prefix.py # Logic การ Merge แบบ Prefix
└── README.md             # คู่มือการใช้งาน
```

## หมายเหตุ
- ระบบปิดการตรวจสอบ SSL (`verify=False`) ในการดาวน์โหลดไฟล์จาก URL เพื่อความสะดวกในการทดสอบในสภาพแวดล้อม Development
- รองรับเฉพาะไฟล์นามสกุล `.docx` เท่านั้น