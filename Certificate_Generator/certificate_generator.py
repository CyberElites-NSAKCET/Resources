import os
from reportlab.pdfgen import canvas
from PyPDF2 import PdfWriter, PdfReader

def create_certificate(template_path, names_list, output_path, position):
    for index, name in enumerate(names_list):
        # Create a canvas and draw the name at the given position
        c = canvas.Canvas(f"tmp/temp_certificate_{index}.pdf")
        c.drawString(position[0], position[1], name)
        c.save()

        # Merge the canvas with the template
        packet = open(f"tmp/temp_certificate_{index}.pdf", "rb")
        new_pdf = PdfReader(packet)
        existing_pdf = PdfReader(open(template_path, "rb"))
        output = PdfWriter()

        # Add the "watermark" (the new pdf) on the existing page
        page = existing_pdf.pages[0]
        page.merge_page(new_pdf.pages[0])
        output.add_page(page)

        # Write the new PDF to a file
        with open(f"{output_path}/certificate_{index}.pdf", "wb") as outputStream:
            output.write(outputStream)

names = ["CyberElites Club"]   # List of names

template_path = "Certificate_Generator/Certificate_Template.pdf"
tmp_directory_path = "tmp"
output_path = "output_certificates"

os.makedirs(output_path, exist_ok=True)
os.makedirs(tmp_directory_path, exist_ok=True)

position = (250, 250)

create_certificate(template_path, names, output_path, position)
