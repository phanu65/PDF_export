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
input_pdf_path = os.path.join(BASE_DIR, "หนังราชการภายนอก.pdf")
output_pdf_path = os.path.join(BASE_DIR, "output.pdf")


# === โหลดฟอนต์ ===
def wrap_text(text, max_width_list, canvas, font="THSarabun", font_size=14):
    lines = []
    current_line = ""

    canvas.setFont(font, font_size)

    words = text.split(" ")
    line_index = 0

    for word in words:
        word_width = canvas.stringWidth(word, font, font_size)

        # ตรวจสอบว่า max_width_list มีค่าเพียงพอสำหรับบรรทัดนี้หรือไม่
        if line_index >= len(max_width_list):
            break

        if word_width > max_width_list[line_index]:
            # ตัดคำที่ยาวเกินบรรทัดออกเป็นท่อน ๆ
            for char in word:
                test_line = current_line + char
                if canvas.stringWidth(test_line, font, font_size) <= max_width_list[line_index]:
                    current_line = test_line
                else:
                    lines.append(current_line)
                    current_line = char
        else:
            test_line = current_line + " " + word if current_line else word
            if canvas.stringWidth(test_line, font, font_size) <= max_width_list[line_index]:
                current_line = test_line
            else:
                lines.append(current_line)
                current_line = word
                line_index += 1

    if current_line:
        lines.append(current_line)

    return lines



def draw_text_lines_custom_width(canvas, text, x_list, y_list, max_width_list, font="THSarabun", font_size=14):
    words = text.split(" ")
    lines = []
    current_line = ""
    line_index = 0

    canvas.setFont(font, font_size)

    for word in words:
        if line_index >= len(max_width_list):
            break  # ถ้าจำนวนบรรทัดเกินมากกว่า max_width_list จะหยุดการทำงาน

        test_line = current_line + " " + word if current_line else word
        if canvas.stringWidth(test_line, font, font_size) <= max_width_list[line_index]:
            current_line = test_line
        else:
            lines.append(current_line)
            current_line = word
            line_index += 1

    if current_line and line_index < len(max_width_list):
        lines.append(current_line)

    for i, line in enumerate(lines):
        if i < len(x_list) and i < len(y_list):
            canvas.setFont(font, font_size)
            canvas.drawString(x_list[i], y_list[i], line)
        else:
            print(f"ตำแหน่งไม่พอสำหรับบรรทัดที่ {i+1}, ข้อความ: {line}")
            break



def fill_doc25_form(input_pdf_path, output_pdf_path):
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


    # ===== ข้อมูลส่วนหัว =====
    draw_text(can, "สำนักงานเขตบางกะปิ", 90, max_height-55)
    draw_text(can, "สำนักงานเขตบางกะปิ", 370, max_height-55)
    draw_text(can, "ตำบลคลองจั่น อำเภอคลองลาน", 370, max_height-84)
    draw_text(can, "12", 360, max_height-112)
    draw_text(can, "พฤศจิกายน", 420, max_height-112)
    draw_text(can, "2566", 490, max_height-112)
    
    draw_text(can, "ขอคำสั่งคุ้มครอง", 100, max_height-142)
    draw_text(can, "สำนักพุทรา", 100, max_height-170)
    draw_text(can, "นายสมชาย ใจดี", 100, max_height-198)
    draw_text(can, "ทดสอบ", 135, max_height-227)
    
    long_text_1 = (
    "ข้าพเจ้าขอแสดงความจำนงในการส่ง เพื่อให้ได้รับการดูแล ฟื้นฟู "
    "เพื่อให้ได้รับการดูแล ฟื้นฟูพื่อให้ได้รับการดูแล ฟื้นฟูพื่อให้ได้รับการดูแล ฟื้นฟูพื่อให้ได้รับการดูแล ฟื้นฟูและพัฒนาศักยภาพ  ฟื้นฟูพื่อให้ได้รับการดูแล ฟื้นฟูและพัฒนาศักยภ ฟื้นฟูพื่อให้ได้รับการดูแล ฟื้นฟูและพัฒนาศักยภ ฟื้นฟูพื่อให้ได้รับการดูแล ฟื้นฟูและพัฒนาศักยภ "
    )
    max_width_list = [330,360,400] # กำหนดความกว้างสูงสุดของแต่ละบรรทัด
    x_list = [190, 100,100] # บรรทัดแรกและบรรทัดที่สอง
    y_list = [max_height - 257, max_height - 285, max_height - 313] # บรรทัดแรกและบรรทัดที่สอง
    draw_text_lines_custom_width(can, long_text_1, x_list, y_list, max_width_list)

    long_text_2 = (
    "ข้าพเจ้าขอแสดงความจำนงในการส่ง เพื่อให้ได้รับการดูแล ฟื้นฟู "
    "เพื่อให้ได้รับการดูแล ฟื้นฟูพื่อให้ได้รับการดูแล ฟื้นฟูพื่อให้ได้รับการดูแล ฟื้นฟูพื่อให้ได้รับการดูแล ฟื้นฟูและพัฒนาศักยภาพ  ฟื้นฟูพื่อให้ได้รับการดูแล ฟื้นฟูและพัฒนาศักยภ ฟื้นฟูพื่อให้ได้รับการดูแล ฟื้นฟูและพัฒนาศักยภ ฟื้นฟูพื่อให้ได้รับการดูแล ฟื้นฟูและพัฒนาศักยภ "
    )
    max_width_list = [330,460,400] # กำหนดความกว้างสูงสุดของแต่ละบรรทัด
    x_list = [220, 100,100] # บรรทัดแรกและบรรทัดที่สอง
    y_list = [max_height - 343, max_height - 372, max_height - 400] # บรรทัดแรกและบรรทัดที่สอง
    draw_text_lines_custom_width(can, long_text_2, x_list, y_list, max_width_list)
    
    long_text_3 = (
    "ข้าพเจ้าขอแสดงความจำนงในการส่ง เพื่อให้ได้รับการดูแล ฟื้นฟู "
    "เพื่อให้ได้รับการดูแล ฟื้นฟูพื่อให้ได้รับการดูแล ฟื้นฟูพื่อให้ได้รับการดูแล ฟื้นฟูพื่อให้ได้รับการดูแล ฟื้นฟูและพัฒนาศักยภาพ  ฟื้นฟูพื่อให้ได้รับการดูแล ฟื้นฟูและพัฒนาศักยภ ฟื้นฟูพื่อให้ได้รับการดูแล ฟื้นฟูและพัฒนาศักยภ ฟื้นฟูพื่อให้ได้รับการดูแล ฟื้นฟูและพัฒนาศักยภ "
    )
    max_width_list = [330,400] # กำหนดความกว้างสูงสุดของแต่ละบรรทัด
    x_list = [190, 100] # บรรทัดแรกและบรรทัดที่สอง
    y_list = [max_height - 429, max_height - 457] # บรรทัดแรกและบรรทัดที่สอง
    draw_text_lines_custom_width(can, long_text_3, x_list, y_list, max_width_list)
    
    
    draw_text(can, "นายสมชาย ใจดี", 90, max_height-626)
    draw_text(can, "นายสมชาย ใจดี", 90, max_height-654)
    draw_text(can, "นายสมชาย ใจดี", 115, max_height-682)






   
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
fill_doc25_form(input_pdf_path, output_pdf_path)
