try:
    from PyPDF2 import PdfWriter, PdfReader
    from reportlab.lib.colors import HexColor
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfgen import canvas
except ImportError:
    print("This script requires the 'reportlab' and 'PyPDF2' modules.\nPlease install them using \'pip install reportlab PyPDF2\' and try again.")
    exit(1)
import os

def create_certificate(template_path, names_list, output_path, position, font_path, font_size, font_color):
    # Register the custom font
    pdfmetrics.registerFont(TTFont('CustomFont', font_path))

    for index, name in enumerate(names_list):
        # Create a canvas and set the custom font, size, and color
        c = canvas.Canvas(f"temp_certificates/temp_certificate_{index}.pdf", pagesize=landscape(A4))
        c.setFont('CustomFont', font_size)
        c.setFillColor(HexColor(font_color))
        
        # Calculate the width of the name text
        text_width = pdfmetrics.stringWidth(name, 'CustomFont', font_size)
        
        # Calculate the x position to center the text
        centered_x = position[0] - (text_width / 2) + 20
        
        # Draw the name centered at the given position
        c.drawString(centered_x, position[1], name)
        c.save()

        # Merge the canvas with the template
        packet = open(f"temp_certificates/temp_certificate_{index}.pdf", "rb")
        new_pdf = PdfReader(packet)
        existing_pdf = PdfReader(open(template_path, "rb"))
        output = PdfWriter()

        # Add the "watermark" (the new pdf) on the existing page
        page = existing_pdf.pages[0]
        page.merge_page(new_pdf.pages[0])
        output.add_page(page)

        with open(f"{output_path}/certificate_{index}.pdf", "wb") as outputStream:
            output.write(outputStream)

names = ["CyberElites Club"]   # List of names

template_path = "Certificate_Template.pdf"
tmp_directory_path = "tmp"
output_path = "Generated_Certificates"

font_path = "Fonts/GreatVibes-Regular.ttf"
font_size = 50
font_color = "#ffffff"

os.makedirs(output_path, exist_ok=True)
os.makedirs(tmp_directory_path, exist_ok=True)

position = (400, 273)

create_certificate(template_path, names, output_path, position, font_path, font_size, font_color)