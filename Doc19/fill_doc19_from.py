from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from io import BytesIO
import datetime
import os


max_height = 790
checkmark = "✔"

def draw_text(can, text, x, y, font="THSarabun", size=16):
    can.setFont(font, size)
    can.drawString(x, y, text)

# ดึง path ฟอนต์
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FONT_DIR = os.path.abspath(os.path.join(BASE_DIR, "..", "font"))
input_pdf_path = os.path.join(BASE_DIR, "พม.2.pdf")
output_pdf_path = os.path.join(BASE_DIR, "output.pdf")

def fill_doc19_form(input_pdf_path, output_pdf_path):
    try:
        pdfmetrics.registerFont(TTFont('THSarabun', os.path.join(FONT_DIR, 'THSarabunNew.ttf')))
        pdfmetrics.registerFont(TTFont('seguisym', os.path.join(FONT_DIR, 'seguisym.ttf')))
    except Exception as e:
        print("ไม่สามารถโหลดฟอนต์ได้:", e)
        print(f"ตรวจสอบว่าไฟล์ .ttf อยู่ในโฟลเดอร์: {FONT_DIR}")
        return

    # สร้าง overlay PDF
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=A4)

    # ตั้งค่าวันที่ปัจจุบัน (แบบไทย)
    today = datetime.datetime.now()
    thai_date = f"{today.day}/{today.month}/{today.year + 543}"

    # ===== ข้อมูลส่วนหัว =====
    draw_text(can, "สำนักงานเขตบางกะปิ", 430, max_height-99)
    draw_text(can, thai_date, 440, max_height-118)

    # ===== ข้อมูลผู้ให้ถ้อยคำ =====
    draw_text(can, "นายสมชาย ใจดี", 190, max_height-134)
    draw_text(can, "35", 500, max_height-134)
    draw_text(can, "ไทย", 120, max_height-154.2)
    draw_text(can, "ไทย", 215, max_height-154.2)
    draw_text(can, "พุทธ", 330, max_height-154.2)
    draw_text(can, "พนักงานบริษัท", 460, max_height-154.2)

    draw_text(can, "15,000", 110, max_height-171.2)
    draw_text(can, "123/45", 340, max_height-171.2)
    draw_text(can, "5", 410, max_height-171.2)
    draw_text(can, "สุขุมวิท 42", 500, max_height-171.2)

    draw_text(can, "เพชรบุรี", 110, max_height-189.2)
    draw_text(can, "ถนนเพชรบุรี", 270, max_height-189.2)
    draw_text(can, "ราชเทวี", 450, max_height-189.2)
    draw_text(can, "กรุงเทพมหานคร", 120, max_height-207)
    draw_text(can, "10400", 270, max_height-207)
    draw_text(can, "0812345678", 450, max_height-207)

    # ===== ข้อมูลบัตรประชาชน =====
    draw_text(can, checkmark, 68, max_height-227, font="seguisym")
    draw_text(can, checkmark, 188, max_height-227, font="seguisym")
    draw_text(can, checkmark, 305, max_height-227, font="seguisym")
    draw_text(can, "1234567890123", 390, max_height-227)
    draw_text(can, "1234567890123", 100, max_height-243.5)
    draw_text(can, "เขตบางกะปิ", 270, max_height-244)
    draw_text(can, "15", 379, max_height-244)
    draw_text(can, "ตุลาคม", 440, max_height-244)
    draw_text(can, "2566", 510, max_height-244)

    # ===== ข้อมูลเด็ก =====
    draw_text(can, "ดวงดี ใจดี", 250, max_height-261.5)

    # ===== ข้อมูลบิดา =====
    draw_text(can, checkmark, 68, max_height-279.5, font="seguisym")
    draw_text(can, checkmark, 255, max_height-279.5, font="seguisym")
    draw_text(can, "นายสมหมาย ใจดี", 140, max_height-280)

    # ===== ข้อมูลมารดา =====
    draw_text(can, "สมหญิง ใจดี", 290, max_height-298.5)
    draw_text(can, checkmark, 68, max_height-316.5, font="seguisym")
    draw_text(can, "นางสมศรี ใจดี", 140, max_height-316.5)
    draw_text(can, checkmark, 291, max_height-316.5, font="seguisym")
    draw_text(can, "นางสาวสมใจ ใจดี", 360, max_height-316.5)
    draw_text(can, checkmark, 68, max_height-334.5, font="seguisym")
    draw_text(can,"นายสมชาย ใจดี", 140, max_height-352.5)
    draw_text(can,"นักสังคมสงเคราะห์", 361, max_height-352.5)
    draw_text(can, checkmark, 177, max_height-370.5, font="seguisym")
    draw_text(can, checkmark, 357, max_height-370.5, font="seguisym")
    draw_text(can, checkmark, 177, max_height-387.5, font="seguisym")
    draw_text(can, checkmark, 357, max_height-387.5, font="seguisym")
    draw_text(can, "ดวงดี ใจดี", 369, max_height-407)
    draw_text(can, "สมชาย ", 197, max_height-424.5)
    draw_text(can, "ใจดี",330, max_height-424.5)
    draw_text(can, "18", 495, max_height-424.5)
    draw_text(can, "ไทย", 110, max_height-442.5)
    draw_text(can, "ไทย", 194, max_height-442.5)
    draw_text(can, "พุทธ", 347, max_height-442.5)
    draw_text(can, "123/1", 460, max_height-442.5)
    draw_text(can, "5", 527, max_height-442.5)
    draw_text(can, "สุขุมวิท", 118, max_height-460.5)
    draw_text(can, "42", 190, max_height-460.5)
    draw_text(can, "เพชรบุรี", 297, max_height-460.5)
    draw_text(can, "ราชา", 450, max_height-460.5)
    draw_text(can, "พิษณุโลก", 110, max_height-478.5)
    draw_text(can, "10400", 240, max_height-478.5)
    draw_text(can, "0812345678", 426, max_height-478.5)
    draw_text(can,"สมศรี", 210, max_height-496.5)
    draw_text(can,"ใจดี", 340, max_height-496.5)
    draw_text(can, "19", 495, max_height-496.5)
    draw_text(can, "ไทย", 110, max_height-514.5)
    draw_text(can, "ไทย", 194, max_height-514.5)
    draw_text(can, "พุทธ", 347, max_height-514.5)
    draw_text(can, "123/2", 460, max_height-514.5)
    draw_text(can, "5", 527, max_height-514.5)
    draw_text(can, "สุขุมวิท", 118, max_height-532.5)
    draw_text(can, "42", 190, max_height-532.5)
    draw_text(can, "เพชรบุรี", 297, max_height-532.5)
    draw_text(can, "ราชา", 450, max_height-532.5)
    draw_text(can, "พิษณุโลก", 110, max_height-550.5)
    draw_text(can, "10400", 240, max_height-550.5)
    draw_text(can, "0812345678", 426, max_height-550.5)
    draw_text(can, checkmark, 231, max_height-568.5, font="seguisym")
    draw_text(can, checkmark, 347, max_height-568.5, font="seguisym")
    draw_text(can, "สวยามดีวันดีกกกกกกกกกกกกกฃ", 200, max_height-586.5)
    can.save()

    # รวมกับ PDF ต้นฉบับ
    packet.seek(0)
    new_pdf = PdfReader(packet)

    try:
        existing_pdf = PdfReader(input_pdf_path)
    except:
        print(f"ไม่พบไฟล์ PDF ต้นฉบับ: {input_pdf_path}")
        return

    output = PdfWriter()

    page = existing_pdf.pages[0]
    page.merge_page(new_pdf.pages[0])
    output.add_page(page)

    for i in range(1, len(existing_pdf.pages)):
        output.add_page(existing_pdf.pages[i])

    try:
        with open(output_pdf_path, "wb") as f:
            output.write(f)
        print(f"รีเซ็ตไฟล์ PDF แล้ว : {output_pdf_path}")
    except:
        print("ไม่สามารถเขียนไฟล์ผลลัพธ์ได้ กรุณาตรวจสอบสิทธิ์การเขียน")

# วิธีใช้งาน
fill_doc19_form(input_pdf_path, output_pdf_path)
