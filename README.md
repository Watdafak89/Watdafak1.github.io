# ระบบจัดทำโครงการสอน

เว็บแอป Flask สำหรับหน้าจัดเตรียมเอกสารและข้อมูลรายวิชา ตามภาพอ้างอิง

## เริ่มต้นใช้งาน

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```

จากนั้นเปิด `http://127.0.0.1:5000`

หมายเหตุ: ปุ่มวิเคราะห์ในเวอร์ชันนี้เป็น frontend prototype ยังไม่ได้เชื่อม Gemini หรือสร้างไฟล์ Word จริง
