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
import os
import sys


## ===========================================================================
### Functions

# Function to get list of files with specific extension within a directory
def get_files(directory, extension):
    """
    Retrieves a list of files with a given extension from the specified directory.

    Args:
        directory (str): The directory path to search for files.
        extension (str): The file extension to filter by (e.g., 'pdf', 'txt').

    Returns:
        list: A list of filenames with the specified extension in the given directory.

    Raises:
        FileNotFoundError: If the specified directory does not exist.
    """
    
    try:
        files_list = [file for file in os.listdir(directory) if file.endswith(f'.{extension.lower()}')]
    except FileNotFoundError as e:
        print(f"\nError reading template file: {e}")
        exit(1)
        
    return files_list

## --------------------------------------------------------------------------
# Function to select desired font
def select_font():
    """
    Prompts the user to select a TrueType font from the available fonts in the FONTS_DIRECTORY_PATH.

    Returns:
        str: The filename of the selected font.

    Raises:
        SystemExit: If the user input is invalid or an exception occurs.
    """
    
    font_dict = {}
    truetype_font_files = sorted(get_files(FONTS_DIRECTORY_PATH, 'TTF'))
    
    if len(truetype_font_files) < 1:
        print(f"\nNo fonts available in \"Fonts\" directory.\nPlease add any valid TTF files to Fonts directory and try again\n\nExiting....\n")
        exit(1)

    print("\nSelect a font for the Names:")
    for index, font_name in enumerate(truetype_font_files):
        print(f"  {index + 1}. {font_name[:-4]}")
        font_dict[index + 1] = font_name
    try:
        font = int(input("\n--> "))
        font_file = font_dict[font]
    except KeyboardInterrupt:
        print("\n\nKeyboard Interrupt!\n\nExiting...\n")
        exit(1)
    except Exception as e:
        print("\n\nInvalid Input!\nPlease select correct font index.\n\nExiting...\n")
        exit(1)

    return font_file

## --------------------------------------------------------------------------
# Function to get the correct file for certificate genetation
def get_single_file(directory_name, directory, extension):
    """
    Ensures there is a single file with the specified extension in the directory.
    Exits the program if there are none or multiple files.

    Args:
        directory_name (str): The name of the directory (for error messages).
        directory (str): The directory to search for the file.
        extension (str): The file extension to search for.

    Returns:
        str: The single file name with the specified extension.
    """
    
    files = get_files(directory, extension)
    if len(files) == 1:
        return files[0]
    elif len(files) > 1:
        print(f"\nCannot read multiple {extension.upper()} files.")
    else:
        print(f"\nFailed to read from {extension.upper()} file.")
        
        if extension == "TXT":
            try:
                wordlist_file_path = os.path.join(WORDLIST_DIRECTORY_PATH, "wordlist.txt")
                with open(wordlist_file_path, 'w') as f:
                    pass
                print("\nCreating an empty 'wordlist.txt' file in 'Wordlist' directory, add contents to it execute the script again!\n")
            except IOError as e:
                print(f"Error creating 'wordlist.txt' file: {e}")
        
    print(f"Please provide a single {extension.upper()} file within the \"{directory_name}\" directory.\n")
    exit(1)

## --------------------------------------------------------------------------
# Function to sort the provided wordlist file
def sort_wordlist(file_path):
    """
    Sorts the wordlist file contents.

    Args:
        file_path (str): The path to the wordlist file.

    Returns:
        list: A list of sorted lines of the file
    """
    try:
        # Read the names from the file
        with open(file_path, 'r') as file:
            names = file.readlines()
    except:
        print("\nError in reading TXT wordlist!\nPlease ensure that the file is not corrupted.\n\nExiting...\n")
        exit(1)

    # Strip whitespace and sort the names
    sorted_names = sorted(name.title().strip() for name in names if name.strip())

    try:
        with open(file_path, 'w') as output_file:
            output_file.write('\n'.join(sorted_names))
    except:
        print("\nError sorting the wordlist file.\nMake sure that the file isn't open!\n\nExiting...\n")
        exit(1)

    return sorted_names 

## --------------------------------------------------------------------------
# Function to read the contents of the file
def read_wordlist(file_path):
    """
    Reads the contents of a wordlist file, returning a list of lines.

    Args:
        file_path (str): The path to the wordlist file.

    Returns:
        list: A list of lines read from the file, or an empty list if the file does not exist.
    """
    
    forbidden_chars = set('<>\"?|/\\:*')

    names = sort_wordlist(file_path)
    
    # Check each name for forbidden characters
    line_error = False
    print()
    for line_number, line in enumerate(names, start=1):
        if any(char in forbidden_chars for char in line):
            print(f"Error: The wordlist file contains forbidden characters on - line {line_number}: ('{line.strip()}').")
            line_error = True
                    
    if line_error:
        print("Please remove any of the following characters from the wordlist: < > \" ? | / \\ : *\n\nExiting...\n")
        exit(1)
            
    return names

## --------------------------------------------------------------------------
# Function to generate the certificates with appropriate names
def create_certificate(template_file_path, wordlist_file_path):
    """
    Creates certificates by merging names from a wordlist with a template PDF.

    Args:
        template_file_path (str): Path to the template PDF file.
        wordlist_file_path (str): Path to the wordlist file containing names.
    """
    
    # Read and print the contents of the file
    wordlist_contents = read_wordlist(wordlist_file_path)
    if not wordlist_contents:
        print("\nEmpty Wordlist file!\nEnsure that Wordlist TXT files has correct Names.\n\nExiting...\n")
        exit(1)
        
    font_file = select_font()
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
    
    TEMPLATE_DIRECTORY_PATH = os.path.join(CERTIFICATE_GENERATOR_DIRECTORY_PATH, 'Certificate_Template')
    WORDLIST_DIRECTORY_PATH = os.path.join(CERTIFICATE_GENERATOR_DIRECTORY_PATH, 'Wordlist')
    TEMPORARY_DIRECTORY_PATH = os.path.join(CERTIFICATE_GENERATOR_DIRECTORY_PATH, 'tmp')
    OUTPUT_DIRECTORY_PATH = os.path.join(CERTIFICATE_GENERATOR_DIRECTORY_PATH, 'Generated_Certificates')
    
    if len(sys.argv) > 1 and sys.argv[1] == "other_script":
        TEMPLATE_DIRECTORY_PATH = os.path.join(os.path.join(ROOT_REPO_PATH, "New_Folder"), "Certificate_Template")
        WORDLIST_DIRECTORY_PATH = os.path.join(os.path.join(ROOT_REPO_PATH, "New_Folder"), "Wordlist")
        TEMPORARY_DIRECTORY_PATH = os.path.join(os.path.join(ROOT_REPO_PATH, "New_Folder"), "tmp")
        OUTPUT_DIRECTORY_PATH = os.path.join(os.path.join(ROOT_REPO_PATH, "New_Folder"), "Generated_Certificates")
        
    print("\n" + " Certificate Generator ".center(35, "-"))

    os.makedirs(TEMPLATE_DIRECTORY_PATH, exist_ok=True)
    os.makedirs(WORDLIST_DIRECTORY_PATH, exist_ok=True)
    os.makedirs(FONTS_DIRECTORY_PATH, exist_ok=True)

    pdf_files = get_files(TEMPLATE_DIRECTORY_PATH, 'PDF')
    template_file = get_single_file('Certificate_Template', TEMPLATE_DIRECTORY_PATH, 'PDF')
    template_file_path = os.path.join(TEMPLATE_DIRECTORY_PATH, template_file)

    text_files = get_files(WORDLIST_DIRECTORY_PATH, 'TXT')
    wordlist_file = get_single_file('Wordlist', WORDLIST_DIRECTORY_PATH, 'TXT')
    wordlist_file_path = os.path.join(WORDLIST_DIRECTORY_PATH, wordlist_file)
    
    FONT_SIZE = 31.5
    FONT_COLOR = "#55D3E2"
    POSITION = (421, 264)
    CHAR_SPACING = 1.5

    certificates_dir = create_certificate(template_file_path, wordlist_file_path)
    
    # Check command-line arguments
    if len(sys.argv) > 1 and sys.argv[1] == "other_script":
        output_dir_path_file = os.path.join(os.path.join(ROOT_REPO_PATH, "New_Folder"), "output_dir_path.txt")
        with open(output_dir_path_file, "w") as file:
            file.write(certificates_dir)
            
    print("\n\nCertificates generation successfull!\n\nSaved all certificates to \"" + certificates_dir[certificates_dir.find("Resources")+10:] + "\" directory.\n")
