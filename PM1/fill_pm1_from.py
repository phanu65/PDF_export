from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from io import BytesIO
import datetime
import os

# === กำหนดพาธต่าง ๆ ===
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FONT_DIR = os.path.abspath(os.path.join(BASE_DIR, "..", "font"))
input_pdf_path = os.path.join(BASE_DIR, "พ.ม.1.pdf")
output_pdf_path = os.path.join(BASE_DIR, "output.pdf")

# === ค่าคงที่ ===
max_height = 790
checkmark = "✔"

def draw_text(can, text, x, y, font="THSarabun", size=16):
    can.setFont(font, size)
    can.drawString(x, y, text)

def fill_pm1_form(input_pdf_path, output_pdf_path):
    # โหลดฟอนต์
    try:
        pdfmetrics.registerFont(TTFont('THSarabun', os.path.join(FONT_DIR, 'THSarabunNew.ttf')))
        pdfmetrics.registerFont(TTFont('seguisym', os.path.join(FONT_DIR, 'seguisym.ttf')))
    except Exception as e:
        print("ไม่สามารถโหลดฟอนต์ได้:", e)
        print(f"ตรวจสอบว่าไฟล์ .ttf อยู่ในโฟลเดอร์: {FONT_DIR}")
        return

    # === สร้างหน้า overlay สำหรับหน้าแรก ===
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=A4)

    # วันที่ (แบบไทย)
    today = datetime.datetime.now()
    thai_date = f"{today.day}/{today.month}/{today.year + 543}"

    # วาดข้อความเฉพาะหน้าแรก
    draw_text(can, thai_date, 430, max_height-120)
    draw_text(can, "นายสมชาย ใจดี", 430, max_height-140)

    can.save()
    packet.seek(0)
    new_pdf = PdfReader(packet)

    # === เปิด PDF ต้นฉบับ ===
    try:
        existing_pdf = PdfReader(input_pdf_path)
    except FileNotFoundError:
        print(f"ไม่พบไฟล์ PDF ต้นฉบับ: {input_pdf_path}")
        return

    output = PdfWriter()

    # === ผสาน overlay กับหน้าแรก ===
    first_page = existing_pdf.pages[0]
    first_page.merge_page(new_pdf.pages[0])
    output.add_page(first_page)

    # === เพิ่มหน้าที่เหลือ ===
    for i in range(1, len(existing_pdf.pages)):
        output.add_page(existing_pdf.pages[i])

    # === เขียนออกเป็น PDF ใหม่ ===
    try:
        with open(output_pdf_path, "wb") as f:
            output.write(f)
        print(f"รีเซ็ตไฟล์ PDF แล้ว: {output_pdf_path}")
    except Exception as e:
        print("ไม่สามารถเขียนไฟล์ผลลัพธ์ได้:", e)

# === เรียกใช้ฟังก์ชัน ===
fill_pm1_form(input_pdf_path, output_pdf_path)
# === จบ ===