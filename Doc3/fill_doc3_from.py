from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from io import BytesIO
import os

# === กำหนดพาธต่าง ๆ ===
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FONT_DIR = os.path.abspath(os.path.join(BASE_DIR, "..", "font"))
input_pdf_path = os.path.join(BASE_DIR, "พ.ม.1.pdf")
output_pdf_path = os.path.join(BASE_DIR, "output.pdf")

# === ค่าคงที่ ===
max_height = 790
checkmark = ("✔")

def draw_text(can, text, x, y, font="THSarabun", size=16):
    can.setFont(font, size)
    can.drawString(x, y, text)

def fill_doc3_form(input_pdf_path, output_pdf_path):
    # โหลดฟอนต์
    try:
        pdfmetrics.registerFont(TTFont('THSarabun', os.path.join(FONT_DIR, 'THSarabunNew.ttf')))
        pdfmetrics.registerFont(TTFont('seguisym', os.path.join(FONT_DIR, 'seguisym.ttf')))
    except Exception as e:
        print("ไม่สามารถโหลดฟอนต์ได้:", e)
        print(f"ตรวจสอบว่าไฟล์ .ttf อยู่ในโฟลเดอร์: {FONT_DIR}")
        return

    # === สร้าง overlay หน้า 1 ===
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=A4)

    draw_text(can, "บ้านพักเด็กและครอบครัว", 435, max_height-112)
    draw_text(can, "12", 391, max_height-131)
    draw_text(can, "พฤศจิกายน", 440, max_height-131)
    draw_text(can, "2566", 520, max_height-131)
    draw_text(can, "นายสมชาย ใจดี", 220, max_height-148)
    draw_text(can, "35", 435, max_height-148)
    draw_text(can, "ไทย", 520, max_height-148)
    draw_text(can, "ไทย", 120, max_height-167.2)
    draw_text(can, "พุทธ", 215, max_height-167.2)
    draw_text(can, "พนักงานบริษัท", 320, max_height-167.2)
    draw_text(can, "15,000", 435, max_height-167.2)
    draw_text(can, "123/45", 150, max_height-185.2)
    draw_text(can, "5", 234, max_height-185.2)
    draw_text(can, "คลองลาน", 320, max_height-185.2)
    draw_text(can, "เมือง", 455, max_height-185.2)
    draw_text(can, "เพชรบุรี", 110, max_height-203.2)
    draw_text(can, "65000", 290, max_height-203.2)
    draw_text(can, "097-1234567", 455, max_height-203.2)
    draw_text(can, "1234567890123", 230, max_height-221.2)
    draw_text(can, "1234567890123", 420, max_height-221.2)
    draw_text(can, "อำเภอ", 120, max_height-238)
    draw_text(can, "12", 360, max_height-238)
    draw_text(can, "พฤศจิกายน", 430, max_height-238)
    draw_text(can, "2566", 520, max_height-238)
    draw_text(can, checkmark, 169, max_height-256.2,font="seguisym", size=20)
    draw_text(can, checkmark, 169, max_height-274.2,font="seguisym", size=20)
    draw_text(can, checkmark, 349, max_height-274.2,font="seguisym", size=20)
    draw_text(can, checkmark, 169, max_height-292.2,font="seguisym", size=20)
    draw_text(can, checkmark, 349, max_height-292.2,font="seguisym", size=20)

    # บุคคลที่ 2
    draw_text(can, "เด็กหญิงสมหญิง ใจดี", 220, max_height-310.2)
    draw_text(can, "นางสาวสมศรี ใจดี", 220, max_height-328.2)
    draw_text(can, "25", 435, max_height-328.2)
    draw_text(can, "ไทย", 520, max_height-328.2)
    draw_text(can, "ไทย", 120, max_height-347.2)
    draw_text(can, "พุทธ", 215, max_height-347.2)
    draw_text(can, "พนักงานบริษัท", 320, max_height-347.2)
    draw_text(can, "15,000", 435, max_height-347.2)
    draw_text(can, "123/45", 150, max_height-365.2)
    draw_text(can, "5", 234, max_height-365.2)
    draw_text(can, "คลองลาน", 320, max_height-365.2)
    draw_text(can, "เมือง", 455, max_height-365.2)
    draw_text(can, "เพชรบุรี", 110, max_height-383.2)
    draw_text(can, "65000", 290, max_height-383.2)
    draw_text(can, "097-1234567", 455, max_height-383.2)
    draw_text(can, "1234567890123", 230, max_height-401.2)
    draw_text(can, "1234567890123", 420, max_height-401.2)
    draw_text(can, "อำเภอ", 120, max_height-418.2)
    draw_text(can, "12", 360, max_height-418.2)
    draw_text(can, "พฤศจิกายน", 430, max_height-418.2)
    draw_text(can, "2566", 520, max_height-418.2)
    draw_text(can, checkmark, 169, max_height-436.2,font="seguisym", size=20)
    draw_text(can, checkmark, 169, max_height-454.2,font="seguisym", size=20)
    draw_text(can, checkmark, 349, max_height-454.2,font="seguisym", size=20)
    draw_text(can, checkmark, 169, max_height-472.2,font="seguisym", size=20)
    draw_text(can, checkmark, 349, max_height-472.2,font="seguisym", size=20)

    draw_text(can, "นายสมชาย ใจดี", 170, max_height-490.2)
    draw_text(can, checkmark, 96, max_height-525.2,font="seguisym", size=20)
    draw_text(can, checkmark, 310, max_height-525.2,font="seguisym", size=20)
    draw_text(can, checkmark, 96, max_height-543.2,font="seguisym", size=20)
    draw_text(can, checkmark, 310, max_height-543.2,font="seguisym", size=20)
    draw_text(can, checkmark, 96, max_height-561.2,font="seguisym", size=20)
    draw_text(can, checkmark, 310, max_height-561.2,font="seguisym", size=20)
    draw_text(can, checkmark, 96, max_height-579.2,font="seguisym", size=20)
    draw_text(can, "ทดสอบระบบ", 170, max_height-579.2)
    draw_text(can, "นายสมชาย ใจดี", 370, max_height-597.2)
    draw_text(can, "นักสังคมสงเคราะห์", 120, max_height-617.2)

    can.save()
    packet.seek(0)
    new_pdf_page1 = PdfReader(packet)

    # === เปิด PDF ต้นฉบับ ===
    try:
        existing_pdf = PdfReader(input_pdf_path)
    except FileNotFoundError:
        print(f"ไม่พบไฟล์ PDF ต้นฉบับ: {input_pdf_path}")
        return

    output = PdfWriter()

    # === ผสานหน้าแรก ===
    first_page = existing_pdf.pages[0]
    first_page.merge_page(new_pdf_page1.pages[0])
    output.add_page(first_page)

    # === ผสานหน้า 2 ถ้ามี ===
    if len(existing_pdf.pages) > 1:
        packet2 = BytesIO()
        can2 = canvas.Canvas(packet2, pagesize=A4)
        draw_text(can2, "นาง สมใจ ใจคำวัง", 225, max_height - 88.5)
        draw_text(can2, checkmark, 98, max_height - 170.5, font="seguisym", size=20)
        draw_text(can2, "นายไก่ชา มาละนะ", 300, max_height - 203.5, )
        draw_text(can2, "5", 473, max_height - 235.5, )
        draw_text(can2, "2", 515, max_height - 235.5, )
        draw_text(can2, "นาง สมใน ใจคำวัง", 200, max_height - 331.5, )
        draw_text(can2, checkmark, 98, max_height - 365.5, font="seguisym", size=20)
        draw_text(can2, "นายไก่ชา มาละนะ", 300, max_height - 397.5, )
        draw_text(can2, "นายไก่ชา มาละนะ", 350, max_height - 461.5, )
        draw_text(can2, checkmark, 98, max_height - 510.5, font="seguisym", size=20)
        draw_text(can2, checkmark, 98, max_height - 527.5, font="seguisym", size=20)
        can2.save()
        packet2.seek(0)
        new_pdf_page2 = PdfReader(packet2)

        page2 = existing_pdf.pages[1]
        page2.merge_page(new_pdf_page2.pages[0])
        output.add_page(page2)

    # === เพิ่มหน้าถัดไป (ถ้ามีหน้า 3+) ===
    for i in range(2, len(existing_pdf.pages)):
        output.add_page(existing_pdf.pages[i])

    # === เขียนออกเป็น PDF ใหม่ ===
    try:
        with open(output_pdf_path, "wb") as f:
            output.write(f)
        print(f"รีเซ็ตไฟล์ PDF แล้ว: {output_pdf_path}")
    except Exception as e:
        print("ไม่สามารถเขียนไฟล์ผลลัพธ์ได้:", e)

# === เรียกใช้ฟังก์ชัน ===
fill_doc3_form(input_pdf_path, output_pdf_path)
