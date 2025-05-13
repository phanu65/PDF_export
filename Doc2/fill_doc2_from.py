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

def draw_text(can, text, x, y, font="THSarabun", size=14):
    can.setFont(font, size)
    can.drawString(x, y, text)

# ดึง path ฟอนต์
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FONT_DIR = os.path.abspath(os.path.join(BASE_DIR, "..", "font"))
input_pdf_path = os.path.join(BASE_DIR, "แบบแสดงความจำนงส่งผู้รับบริการ.pdf")
output_pdf_path = os.path.join(BASE_DIR, "output.pdf")

def wrap_text(text, max_width, canvas, font="THSarabun", font_size=14):
    words = text.split(" ")  # แยกคำจากช่องว่าง
    lines = []
    current_line = ""

    canvas.setFont(font, font_size)

    for word in words:
        # ทดสอบบรรทัดที่มีคำใหม่เพิ่ม
        test_line = current_line + " " + word if current_line else word
        
        # เช็คว่าความกว้างของบรรทัดไม่เกิน max_width
        if canvas.stringWidth(test_line, font, font_size) <= max_width:
            current_line = test_line
        else:
            if current_line:
                lines.append(current_line)  # เพิ่มบรรทัดเก่าที่ได้มา
            current_line = word  # เริ่มบรรทัดใหม่ด้วยคำนี้

    # เพิ่มบรรทัดสุดท้าย
    if current_line:
        lines.append(current_line)

    return lines


def draw_text_lines_custom_xy(canvas, text, x_list, y_list, max_width, font="THSarabun", font_size=14):
    lines = wrap_text(text, max_width, canvas, font, font_size)  # รับข้อมูลบรรทัดที่จัดการแล้ว
    for i, line in enumerate(lines):
        if i < len(x_list) and i < len(y_list):  # ตรวจสอบขนาดของ x_list และ y_list
            canvas.setFont(font, font_size)
            canvas.drawString(x_list[i], y_list[i], line)  # วาดข้อความตามตำแหน่งที่กำหนด
        else:
            break  # หากไม่สามารถวาดได้เนื่องจาก x_list หรือ y_list ไม่มีพิกัดพอ




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
    draw_text(can, "สำนักงานเขตบางกะปิ", 400, max_height-54)
    draw_text(can, "1", 330, max_height-75)
    draw_text(can, "พฤศจิกายน", 390, max_height-75)
    draw_text(can, "2566", 485, max_height-75)
    
    draw_text(can, "นายสมชาย ใจดี", 260, max_height-114)
    draw_text(can, "35", 485, max_height-114)
    draw_text(can, "นักสังคมสงเคราะห์", 120, max_height-134.2)
    draw_text(can, "บ้านพักเด็กและครอบครัว", 350, max_height-134.2)
    draw_text(can, "123", 130, max_height-154.2)
    draw_text(can, "5", 177, max_height-154.2)
    draw_text(can, "ลำธาน", 220, max_height-154.2)
    draw_text(can, "เมือง", 320, max_height-154.2)
    draw_text(can, "เพชรบุรี", 430, max_height-154.2)
    draw_text(can, "เพชรบุรี", 110, max_height-173.2)
    draw_text(can, "068-990-1123", 290, max_height-173.2)
    draw_text(can, checkmark, 75, max_height-193.2, font="seguisym", size=16)
    draw_text(can, checkmark, 182, max_height-193.2, font="seguisym", size=16)
    draw_text(can, checkmark, 360, max_height-193.2, font="seguisym", size=16)
    draw_text(can, "ทดสอบระบบ", 110, max_height-212.2,)
    draw_text(can, "1234567890123", 210, max_height-212.2)
    draw_text(can, "อำเภอ", 370, max_height-212.2)
    draw_text(can, thai_date, 475, max_height-212.2)
    draw_text(can, "นายสมชาย ใจดี", 320, max_height-232.2)
    draw_text(can, "12", 485, max_height-232.2)
    draw_text(can, "123/23", 130, max_height-252.2)
    draw_text(can, "5", 200, max_height-252.2)
    draw_text(can, "ลานยาง", 240, max_height-252.2)
    draw_text(can, "คลองลาน", 330, max_height-252.2)
    draw_text(can, "เพชรบูรณ์", 440, max_height-252.2)
    draw_text(can, "เพชรบูรณ์", 110, max_height-271.2)
    draw_text(can, "097-1234567", 290, max_height-271.2)
    draw_text(can, "1", 230, max_height-291.2)
    draw_text(can, "พฤศจิกายน", 285, max_height-291.2)
    draw_text(can, "2566", 380, max_height-291.2)
    draw_text(can, "16.00 ", 455, max_height-291.2, )
    # ถ้าข้อความถูกตัด 2 บรรทัด ให้ขึ้นที่ y พิกัดใหม่
    long_text = (
    "ข้าพเจ้าขอแสดงความจำนงในการส่งต่อผู้รับบริการรายนี้ไปยังหน่วยงานที่เกี่ยวข้อง "
    "เพื่อให้ได้รับการดูแล ฟื้นฟู และพัฒนาศักยภาพตามความเหมาะสม......................................sdasdasda......................................"
    "โดยพิจารณาจากข้อมูลพื้นฐาน สภาพปัญหา และความต้องการของผู้รับบริการ "
    "ซึ่งผ่านการประเมินร่วมกันระหว่างเจ้าหน้าที่ที่เกี่ยวข้องและผู้รับบริการแล้ว "
    "ทั้งนี้เพื่อประโยชน์สูงสุดต่อการดำรงชีวิตและการพัฒนาคุณภาพชีวิตของผู้รับบริการอย่างยั่งยืน"
    )
    
    x_list = [120, 75,75,75] # บรรทัดแรกและบรรทัดที่สอง
    y_list = [max_height-311.2, max_height-331.2,max_height-351.2,max_height-371.2] # บรรทัดแรกและบรรทัดที่สอง
    draw_text_lines_custom_xy(can, long_text, x_list, y_list, max_width=400)
    
    
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
