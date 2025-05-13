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
input_pdf_path = os.path.join(BASE_DIR, "บันทึกถ้อยคำ.pdf")
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



# === ค่าคงที่ ===
max_height = 790
checkmark = ("✔")

def draw_text(can, text, x, y, font="THSarabun", size=16):
    can.setFont(font, size)
    can.drawString(x, y, text)

def fill_doc21_form(input_pdf_path, output_pdf_path):
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
    draw_text(can, "บัานพักเด็กและครอบครัว", 370, max_height-23)
    draw_text(can, "ตำบล ท่าช้าง", 300, max_height-46)
    draw_text(can, "อำเภอเมือง", 380, max_height-46)
    draw_text(can, "จังหวัดสุโขทัย", 460, max_height-46)
    draw_text(can, "12 ",330, max_height-70)
    draw_text(can, "พฤจิกายยน ",410, max_height-70)
    draw_text(can, "2556 ",510, max_height-70) 
    draw_text(can, "นาย สมใน วงวัง", 190, max_height-94)
    draw_text(can, "12", 410, max_height-94)
    draw_text(can, "1234567890123", 55, max_height-117)
    draw_text(can, "1234", 270, max_height-117)
    draw_text(can, "1234567890123", 55, max_height-140)
    draw_text(can, "1234", 450, max_height-140)
    draw_text(can,"06-1234-5678", 115, max_height-163)
    
    long_text = (
    "ข้าพเจ้าขอแสดงความจำนงในการส่ง เพื่อให้ได้รับการดูแล ฟื้นฟู "
    "พื่อให้ได้รับการดูแล ฟื้นฟูพื่อให้ได้รับการดูแล"
    "โดยพิจารณาจากข้อมูลพื้นฐานสภาพปัญหาและความต้องการของผู้รับบริการซึ่งผ่านการประเมิน ร่วมกันระหว่างเจ้าหน้าที่"
    "และหน่วยงานที่เกี่ยวข้อง"
    "เพื่อการพัฒนาที่สอดคล้อง กับ ความสามารถและความต้องการที่แท้จริง"
    "ในการฟื้นฟูศักยภาพและการพัฒนาของผู้รับบริการ"
    "โดยคำนึงถึง การสร้างความยั่งยืนในระยะยาว"
    "ข้าพเจ้าขอแสดงความจำนงในการส่ง เพื่อให้ได้รับการดูแล ฟื้นฟู "
    "พื่อให้ได้รับการดูแล ฟื้นฟูพื่อให้ได้รับการดูแล"
    "โดยพิจารณาจากข้อมูลพื้นฐานสภาพปัญหาและความต้องการของผู้รับบริการซึ่งผ่านการประเมิน ร่วมกันระหว่างเจ้าหน้าที่"
    "และหน่วยงานที่เกี่ยวข้อง"
    "ข้าพเจ้าขอแสดงความจำนงในการส่ง เพื่อให้ได้รับการดูแล ฟื้นฟู "
    "พื่อให้ได้รับการดูแล ฟื้นฟูพื่อให้ได้รับการดูแล"
    "โดยพิจารณาจากข้อมูลพื้นฐานสภาพปัญหาและความ ต้องการของผู้รับบริการซึ่งผ่านการประเมิน ร่วมกันระหว่างเจ้าหน้าที่"
    "และหน่วยงานที่เกี่ยวข้อง"
    )
    max_width_list = [400,600,600,600] # กำหนดความกว้างสูงสุดของแต่ละบรรทัด
    x_list = [310, 45,45,45] # บรรทัดแรกและบรรทัดที่สอง
    y_list = [max_height-163.2, max_height-188.2,max_height-212.2,max_height-236.2] # บรรทัดแรกและบรรทัดที่สอง
    draw_text_lines_custom_width(can, long_text, x_list, y_list, max_width_list)

    max_width_list = [400,600,600] # กำหนดความกว้างสูงสุดของแต่ละบรรทัด
    x_list = [75, 45,45] # บรรทัดแรกและบรรทัดที่สอง
    y_list = [max_height-259.2, max_height-282.2,max_height-305.2] # บรรทัดแรกและบรรทัดที่สอง
    draw_text_lines_custom_width(can, long_text, x_list, y_list, max_width_list)
    
    max_width_list = [400,600,600,600,600]# กำหนดความกว้างสูงสุดของแต่ละบรรทัด
    x_list = [75, 45,45, 45,45] # บรรทัดแรกและบรรทัดที่สอง
    y_list = [max_height-354.2, max_height-377.2,max_height-400.2,max_height-423.2,max_height-446.2] # บรรทัดแรกและบรรทัดที่สอง
    draw_text_lines_custom_width(can, long_text, x_list, y_list, max_width_list)
    
    draw_text(can, " 1 ตุลาคม 2568", 115, max_height-471.2)
    max_width_list = [400,600,600,600,600] # กำหนดความกว้างสูงสุดของแต่ละบรรทัด
    x_list = [75, 45,45, 45,45] # บรรทัดแรกและบรรทัดที่สอง
    y_list = [max_height-517.2, max_height-541.2,max_height-565.2,max_height-587.2,max_height-611.2] # บรรทัดแรกและบรรทัดที่สอง
    draw_text_lines_custom_width(can, long_text, x_list, y_list, max_width_list)
    
    max_width_list = [400,600,600,600] # กำหนดความกว้างสูงสุดของแต่ละบรรทัด
    x_list = [75, 45,45, 45] # บรรทัดแรกและบรรทัดที่สอง
    y_list = [max_height-658.2, max_height-682.2,max_height-706.2,max_height-729.2] # บรรทัดแรกและบรรทัดที่สอง
    draw_text_lines_custom_width(can, long_text, x_list, y_list, max_width_list)
    
    
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
        
        max_width_list = [400,500,500,500,500,500,500] # กำหนดความกว้างสูงสุดของแต่ละบรรทัด
        x_list = [75, 45,45,45,45,45,45]
        y_list = [max_height-19.2, max_height-42.2,max_height-65,max_height-89,max_height-112,max_height-137,max_height-159] # บรรทัดแรกและบรรทัดที่สอง
        draw_text_lines_custom_width(can2, long_text, x_list, y_list, max_width_list)
        
        max_width_list = [400,500,500,500,500,500,500] # กำหนดความกว้างสูงสุดของแต่ละบรรทัด
        x_list = [75, 45,45,45,45,45,45]
        y_list = [max_height-207.2, max_height-230.2,max_height-253,max_height-277,max_height-300,max_height-325,max_height-349] # บรรทัดแรกและบรรทัดที่สอง
        draw_text_lines_custom_width(can2, long_text, x_list, y_list, max_width_list)
        
        max_width_list = [400,500,500,500,500,500] # กำหนดความกว้างสูงสุดของแต่ละบรรทัด
        x_list = [75, 45,45,45,45,45]
        y_list = [max_height-396.2, max_height-419.2,max_height-442,max_height-466,max_height-491,max_height-513] # บรรทัดแรกและบรรทัดที่สอง
        draw_text_lines_custom_width(can2, long_text, x_list, y_list, max_width_list)
        
        max_width_list = [400,500,500,500,500,500] # กำหนดความกว้างสูงสุดของแต่ละบรรทัด
        x_list = [75, 45,45,45,45,45]
        y_list = [max_height-561.2, max_height-583.2,max_height-606,max_height-630,max_height-653,max_height-678] # บรรทัดแรกและบรรทัดที่สอง
        draw_text_lines_custom_width(can2, long_text, x_list, y_list, max_width_list)
        
        can2.save()
        packet2.seek(0)
        new_pdf_page2 = PdfReader(packet2)

        page2 = existing_pdf.pages[1]
        page2.merge_page(new_pdf_page2.pages[0])
        output.add_page(page2)

    # === เพิ่มการกรอกหน้า 3 ===
    if len(existing_pdf.pages) > 2:
        packet3 = BytesIO()
        can3 = canvas.Canvas(packet3, pagesize=A4)
        
        max_width_list = [400,500,500,500,500,500] # กำหนดความกว้างสูงสุดของแต่ละบรรทัด
        x_list = [75, 45,45,45,45,45]
        y_list = [max_height-42.2, max_height-65.2,max_height-89,max_height-112,max_height-137,max_height-159]
        draw_text_lines_custom_width(can3, long_text, x_list, y_list, max_width_list)
        
        can3.save()
        packet3.seek(0)
        new_pdf_page3 = PdfReader(packet3)
        
        page3 = existing_pdf.pages[2]
        page3.merge_page(new_pdf_page3.pages[0])
        output.add_page(page3)
    
    # === เพิ่มหน้าถัดไป (ตั้งแต่หน้า 4+) ===
    for i in range(3, len(existing_pdf.pages)):
        output.add_page(existing_pdf.pages[i])

    # === เขียนออกเป็น PDF ใหม่ ===
    try:
        with open(output_pdf_path, "wb") as f:
            output.write(f)
        print(f"รีเซ็ตไฟล์ PDF แล้ว: {output_pdf_path}")
    except Exception as e:
        print("ไม่สามารถเขียนไฟล์ผลลัพธ์ได้:", e)

# === เรียกใช้ฟังก์ชัน ===
fill_doc21_form(input_pdf_path, output_pdf_path)