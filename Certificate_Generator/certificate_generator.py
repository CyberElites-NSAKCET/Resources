import os
import sys
# Get the parent directory, add it to python path and import the modules
parent_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
sys.path.append(parent_dir)
from Utilities.utils import read_wordlist, select_font, get_files, get_single_file
try:
    from PyPDF2 import PdfWriter, PdfReader
    from reportlab.lib.colors import HexColor
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfgen import canvas
except ImportError:
    print("\nThis script requires the \'reportlab\' and \'PyPDF2\' modules.\n\nPlease install them using \'pip install reportlab PyPDF2\' and try again.\n")
    exit(1)


## --------------------------------------------------------------------------
# Function to generate the certificates with appropriate names
def generate_certificates(template_file_path, wordlist_contents):
    """
    Generates personalized certificates by combining a template PDF with names from a wordlist.

    Args:
        template_file_path (str): Path to the template PDF file.
        wordlist_contents (list): List of names to be included on the certificates.

    Returns:
        str: Path to the directory containing the generated certificates.

    Raises:
        KeyboardInterrupt: If the user interrupts the process.
        Exception: For any error occurring during certificate generation.
    """   
        
    font_file = select_font(FONTS_DIRECTORY_PATH)
    font_file_path = os.path.join(FONTS_DIRECTORY_PATH, font_file)
    
    try:
        # Register the custom font
        pdfmetrics.registerFont(TTFont('CustomFont', font_file_path))
    except:
        print("\nInvalid Font file!\nPlease ensure that you use a valid TTF file.\n\nExiting...\n")
        exit(1)
        
    # Define a single temporary file path
    tmp_file = os.path.join(TEMPORARY_DIRECTORY_PATH, "tmp_file.pdf")
    os.makedirs(TEMPORARY_DIRECTORY_PATH, exist_ok=True)
    
    counter = 0
    output_folder_path = OUTPUT_DIRECTORY_PATH
    while os.path.exists(output_folder_path):
        counter += 1
        output_folder_path = f"{OUTPUT_DIRECTORY_PATH}({counter})"
        
    os.makedirs(output_folder_path, exist_ok=True)
        
    print("\nGenerating the certificates......\n")
    try:
        for name in wordlist_contents:
        
            filename = "_".join(name.split())
            print(f"{filename}_certificate.pdf")

            # Create a canvas and set the custom font, size, and color
            new_canvas = canvas.Canvas(tmp_file, pagesize=landscape(A4))
            new_canvas.setFont('CustomFont', FONT_SIZE)
            new_canvas.setFillColor(HexColor(FONT_COLOR))
        
            # Calculate the width of the name text with character spacing
            total_text_width = sum(pdfmetrics.stringWidth(char, 'CustomFont', FONT_SIZE) + CHAR_SPACING for char in name) - CHAR_SPACING
        
            # Calculate the x position to center the text with character spacing
            centered_x = POSITION[0] - (total_text_width / 2)
        
            # Draw each character with the specified spacing
            x_offset = centered_x
            for char in name:
                new_canvas.drawString(x_offset, POSITION[1], char)
                x_offset += pdfmetrics.stringWidth(char, 'CustomFont', FONT_SIZE) + CHAR_SPACING
        
            new_canvas.save()

            # Merge the canvas with the template
            packet = open(tmp_file, "rb")
            new_pdf = PdfReader(packet)
            try:
                existing_pdf = PdfReader(open(template_file_path, "rb"))
            except:
                print("\nError in reading PDF template!\nPlease ensure that the file is in the correct directory and not corrupted.\n\nExiting...\n")
                exit(1)
            
            output = PdfWriter()

            # Add the "watermark" (the new pdf) on the existing page
            page = existing_pdf.pages[0]
            page.merge_page(new_pdf.pages[0])
            output.add_page(page)

            with open(f"{output_folder_path}/{filename}_certificate.pdf", "wb") as outputStream:
                output.write(outputStream)
        
            # Clear the canvas for the next iteration (Certificate)
            new_canvas = None
            
        return output_folder_path
            
    except KeyboardInterrupt:
        print("\n\nKeyboard Interrupt!\nAll certificates aren't generated!\n\nExiting...\n")
        exit(1)
    except Exception as e:
        print(f"\nAn error occured in certificate generation!\n{e}\n\nExiting....\n")
        exit(1)
        

### ===========================================================================
## Main 
#

if __name__ == "__main__":
    """
    Main entry point for the script to generate certificates.

    Workflow:
        1. Sets up directory paths for templates, wordlists, and fonts.
        2. Reads user inputs to select certificate type and customize parameters.
        3. Calls `generate_certificates` function to create personalized certificates.
        4. Handles automation script integration if specified in the command-line arguments.
    """
    
    print("\n" + " Certificate Generator ".center(35, "-"))
    
    if os.getcwd()[-9:] == "Resources":
        CERTIFICATE_GENERATOR_DIRECTORY_PATH = os.path.join(os.getcwd(), 'Certificate_Generator')
        ROOT_REPO_PATH = os.getcwd()
        FONTS_DIRECTORY_PATH = os.path.join(os.getcwd(), 'Fonts')
        
    elif os.getcwd()[-21:] == "Certificate_Generator":
        CERTIFICATE_GENERATOR_DIRECTORY_PATH = os.getcwd()
        ROOT_REPO_PATH = os.path.join(os.getcwd(), '..')
        FONTS_DIRECTORY_PATH = os.path.join(ROOT_REPO_PATH, 'Fonts')
        
    else:
        print("\nPlease change your working directory to the main repository.\n\nExiting...\n")
        exit(1)
    
    
    automation_script = len(sys.argv) > 1 and sys.argv[1] == "extract_certify_and_email_script"
    
    if automation_script:
        CERTIFICATE_EMAIL_AUTOMATION_DIRECTORY_PATH = os.path.join(ROOT_REPO_PATH, "Certificate_Email_Automation")
        DIR_PATH = CERTIFICATE_EMAIL_AUTOMATION_DIRECTORY_PATH
    else:
        DIR_PATH = CERTIFICATE_GENERATOR_DIRECTORY_PATH
        
    TEMPLATE_DIRECTORY_PATH = os.path.join(DIR_PATH, "Certificate_Template")
    WORDLIST_DIRECTORY_PATH = os.path.join(DIR_PATH, "Wordlist")
    TEMPORARY_DIRECTORY_PATH = os.path.join(DIR_PATH, "tmp")
    OUTPUT_DIRECTORY_PATH = os.path.join(DIR_PATH, "Generated_Certificates")

    os.makedirs(TEMPLATE_DIRECTORY_PATH, exist_ok=True)
    os.makedirs(WORDLIST_DIRECTORY_PATH, exist_ok=True)
    os.makedirs(FONTS_DIRECTORY_PATH, exist_ok=True)

    pdf_files = get_files(TEMPLATE_DIRECTORY_PATH, 'PDF')
    template_file = get_single_file('Certificate_Template', TEMPLATE_DIRECTORY_PATH, 'PDF')
    template_file_path = os.path.join(TEMPLATE_DIRECTORY_PATH, template_file)

    text_files = get_files(WORDLIST_DIRECTORY_PATH, 'TXT')
    wordlist_file = get_single_file('Wordlist', WORDLIST_DIRECTORY_PATH, 'TXT')
    wordlist_file_path = os.path.join(WORDLIST_DIRECTORY_PATH, wordlist_file)
    
    # Read and print the contents of the file
    wordlist_contents = read_wordlist(wordlist_file_path)
    
    try:
        certificate_type = int(input(f"\nSelect the type of certificate:\n  1. Membership Certificate\n  2. Event Certificate\n\n--> "))
    except KeyboardInterrupt:
        print("\n\nKeyboard Interrupt!\n\nExiting...\n")
        exit(1)
    except Exception as e:
        print("\n\nInvalid Input!\nPlease select correct certificate type.\n\nExiting...\n")
        exit(1)
        
    if certificate_type == 1:    # Membership_Certificate
        FONT_SIZE = 31.5
        FONT_COLOR = "#55D3E2"
        POSITION = (421, 264)
        CHAR_SPACING = 1.5
    
    elif certificate_type == 2:    # Event_Certificate
        FONT_SIZE = 34
        FONT_COLOR = "#ffffff"
        POSITION = (421, 242)
        CHAR_SPACING = 1.15
    else:
        print("\n\nInvalid Input!\nPlease select correct certificate type.\n\nExiting...\n")
        exit(1)

    certificates_dir = generate_certificates(template_file_path, wordlist_contents)
    
    # Check command-line arguments
    if automation_script:
        gen_certs_dir_path = os.path.join(os.path.join(ROOT_REPO_PATH, "Certificate_Email_Automation"), "gen_certs_dir_path.txt")
        with open(gen_certs_dir_path, "w") as file:
            file.write(certificates_dir)
            
    print("\n\nCertificates generation successfull!\n\nSaved all certificates to \"" + certificates_dir[certificates_dir.find("Resources")+10:] + "\" directory.\n")
