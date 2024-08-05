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

def select_font():
    font_dict = {}
    truetype_font_files = get_files(FONTS_DIRECTORY_PATH, 'TTF')

    print("\nSelect a font:")
    for index, font_name in enumerate(truetype_font_files):
        print(f"{index + 1}. {font_name[:-4]}")
        font_dict[index + 1] = font_name
    try:
        font = int(input("\n--> "))
        font_file = font_dict[font]
    except:
        print("\nInvalid Input! Please select correct font index.\n\nExiting...\n")
        exit(1)

    return font_file
    
def get_files(directory, extension):
    pdf_files = [file for file in os.listdir(directory) if file.endswith(f'.{extension.lower()}')]
    return pdf_files

def get_single_file(directory_name, directory, extension):
    files = get_files(directory, extension)
    if len(files) == 1:
        return files[0]
    elif len(files) > 1:
        print(f"\nCannot read multiple {extension.upper()} files.\nPlease provide a single {extension.upper()} file within the \"{directory_name}\" directory.\n\nExiting....\n")
        exit(1)
    else:
        print(f"\nFailed to read from {extension.upper()} file.\nPlease provide a single {extension.upper()} file within the \"{directory_name}\" directory.\n\nExiting....\n")
        exit(1)

# Function to read the contents of the file
def read_wordlist(file_path):
    try:
        with open(file_path, 'r') as file:
            contents = file.readlines()
        return contents
    except FileNotFoundError:
        print(f"The file {file_path} does not exist.")
        return []

def create_certificate(template_file_path, wordlist_file_path, position, font_path, font_size, font_color):
    # Register the custom font
    pdfmetrics.registerFont(TTFont('CustomFont', font_path))
    
    # Read and print the contents of the file
    wordlist_contents = read_wordlist(wordlist_file_path)
    
    for name in wordlist_contents:
        name = name.strip()
        if not name:
            continue
        
        filename = "_".join(name.split())
        tmp_file = os.path.join(TEMPORARY_DIRECTORY_PATH, f"temp_{filename}.pdf")

        # Create a canvas and set the custom font, size, and color
        new_canvas = canvas.Canvas(tmp_file, pagesize=landscape(A4))
        new_canvas.setFont('CustomFont', font_size)
        new_canvas.setFillColor(HexColor(font_color))
        
        # Calculate the width of the name text
        text_width = pdfmetrics.stringWidth(name, 'CustomFont', font_size)
        
        # Calculate the x position to center the text
        centered_x = position[0] - (text_width / 2)
        
        # Draw the name centered at the given position
        new_canvas.drawString(centered_x, position[1], name)
        new_canvas.save()

        # Merge the canvas with the template
        packet = open(tmp_file, "rb")
        new_pdf = PdfReader(packet)
        existing_pdf = PdfReader(open(template_file_path, "rb"))
        output = PdfWriter()

        # Add the "watermark" (the new pdf) on the existing page
        page = existing_pdf.pages[0]
        page.merge_page(new_pdf.pages[0])
        output.add_page(page)

        with open(f"{OUTPUT_DIRECTORY_PATH}/{filename}_certificate.pdf", "wb") as outputStream:
            output.write(outputStream)

if __name__ == "__main__":
    
    FONTS_DIRECTORY_PATH = os.path.join(os.getcwd(), 'FONTS')
    TEMPLATE_DIRECTORY_PATH = os.path.join(os.getcwd(), 'Certificate_Template')
    WORDLIST_DIRECTORY_PATH = os.path.join(os.getcwd(), 'Wordlist')
    TEMPORARY_DIRECTORY_PATH = os.path.join(os.getcwd(), 'tmp')
    OUTPUT_DIRECTORY_PATH = os.path.join(os.getcwd(), 'Generated_Certificates')

    font_file = select_font()
    font_file_path = os.path.join(FONTS_DIRECTORY_PATH, font_file)

    pdf_files = get_files(TEMPLATE_DIRECTORY_PATH, 'PDF')
    template_file = get_single_file('Certificate_Template', TEMPLATE_DIRECTORY_PATH, 'PDF')
    template_file_path = os.path.join(TEMPLATE_DIRECTORY_PATH, template_file)

    text_files = get_files(WORDLIST_DIRECTORY_PATH, 'TXT')
    wordlist_file = get_single_file('Wordlist', WORDLIST_DIRECTORY_PATH, 'TXT')
    wordlist_file_path = os.path.join(WORDLIST_DIRECTORY_PATH, wordlist_file)

    os.makedirs(TEMPLATE_DIRECTORY_PATH, exist_ok=True)
    os.makedirs(WORDLIST_DIRECTORY_PATH, exist_ok=True)
    os.makedirs(TEMPORARY_DIRECTORY_PATH, exist_ok=True)
    os.makedirs(OUTPUT_DIRECTORY_PATH, exist_ok=True)

    font_size = 50
    font_color = "#55D3E2"
    position = (420, 375)

    create_certificate(template_file_path, wordlist_file_path, position, font_file_path, font_size, font_color)
