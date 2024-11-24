try:
    import qrcode
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    print("This script requires the 'qrcode' and 'pillow' modules.\nPlease install them using 'pip install qrcode pillow' and try again.")
    exit(1)
import os
import re
from qrcode.image.styledpil import StyledPilImage
from qrcode.image.styles.moduledrawers import CircleModuleDrawer
from qrcode.image.styles.colormasks import SolidFillColorMask


## ===========================================================================
### Functions

# Function to input QR text
def get_text():
    """
    Prompts the user to input text to encode in the QR code.
    The user can enter multiple lines of text. To finish input, the user must press Enter twice.

    Returns:
        str: The input text to encode in the QR code.

    Exits:
        Exits the program if no text is provided.
    """

    print("\nEnter the text to encode as QR (Press Enter \'Twice\' to finish):\n")

    text_lines = []
    enter_pressed = 0
    
    while True:
        line = input()
        if not line:
            enter_pressed += 1
            if enter_pressed > 1:
                break
            text_lines.append("")
            continue
        else:
            enter_pressed = 0
        text_lines.append(line.rstrip())

    input_text = '\n'.join(text_lines).strip()

    if input_text.strip() == "":
        print("No Input!!!\nCan't generate QR without input text\n")
        exit(1)

    return input_text

## --------------------------------------------------------------------------
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
    
    pdf_files = [file for file in os.listdir(directory) if file.endswith(f'.{extension.lower()}')]
    return pdf_files

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
    truetype_font_files = get_files(FONTS_DIRECTORY_PATH, 'TTF')
    
    if len(truetype_font_files) == 0:
        print(f"\nNo fonts available in \"Fonts\" directory.\nPlease add any TTF files to Fonts directory and try again\n\nExiting....\n")
        exit(1)

    print("\nSelect a font for the QR title:")
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

## --------------------------------------------------------------------------
# Function to generate QR

def qr_gen(input_text, error_correction):
    """
    Generates a QR code image from the provided text with a specified error correction level.

    Args:
        input_text (str): The text to encode in the QR code.
        error_correction (str): The error correction level ('L', 'M', 'Q', 'H').

    Returns:
        PIL.Image.Image: The generated QR code image.

    Exits:
        Exits the program if an invalid error correction level is provided.
    """

    # Create a QR Code instance
    qr = qrcode.QRCode(
        version=4,              # Version controls size of QR
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=10,
        border=2,
    )
    qr.add_data(input_text)
    qr.make(fit=True)

    # Create the QR code with dot modules
    qr_img = qr.make_image(
        image_factory=StyledPilImage,
        module_drawer=CircleModuleDrawer(),  # Dots for QR code modules
        color_mask=SolidFillColorMask(front_color=(0, 0, 0), back_color=(255, 255, 255)),
    )

    # Convert QR code to an editable PIL image
    qr_image = qr_img.convert("RGBA")

    return qr_image

## --------------------------------------------------------------------------
# Function to get QR Image extension type
def extension_menu():
    """
    Prompts the user to select an image file extension for the QR code image.

    Returns:
        tuple: A tuple containing the file extension and the image format.
            - str: The file extension (e.g., '.png', '.jpg').
            - str: The image format (e.g., "PNG", "JPEG").
    
    Exits:
        Exits the program if an invalid extension is selected.
    """
    
    extension_type = input("\nEnter the image file extension for your QRCode\n 1. JPEG    2. JPG    3. PNG\n 4. GIF     5. TIFF   6. BMP\n\n --> ").lower().strip()
    image_format = "JPEG"

    if extension_type in ['1','jpeg','']:
        extension = '.jpeg'
    
    elif extension_type in ['2','jpg']:
        extension = '.jpg'

    elif extension_type in ['3','png']:
        extension = '.png'
        image_format = "PNG"

    elif extension_type in ['4','gif']:
        extension = '.gif'
        image_format = "GIF"

    elif extension_type in ['5','tiff']:
        extension = '.tiff'
        image_format = "TIFF"

    elif extension_type in ['6','bmp']:
        extension = '.bmp'
        image_format = "BMP"

    else:
        print("\nInvalid extension!!! Please select from the above options\n")
        exit(1)
    
    return extension, image_format

## --------------------------------------------------------------------------
# Function to add an image to the center of the QR Code
def add_center_image(qr_image):
    """
    Adds a center image to the QR code. The image is scaled to fit the center of the QR code.

    Args:
        qr_image (PIL.Image.Image): The QR code image.

    Returns:
        PIL.Image.Image: The QR code image with the center image added.

    Exits:
        Exits the program if the center image cannot be added due to an error.
    """
    
    # center_image_path = input("\nEnter the Full path of the image to place at the center of the QR code (or press Enter to skip): ").strip()
    
    center_image = "White_border_circle.png"
    center_image_path = os.path.join(LOGOS_DIRECTORY_PATH, center_image)
        
    if not center_image_path or center_image_path.strip() == "":
        return qr_image

    try:
        center_image = Image.open(center_image_path)
        qr_width, qr_height = qr_image.size
        center_img_width, center_img_height = center_image.size

        # Scale the center image
        scale_factor = 4
        center_img_width = qr_width // scale_factor
        center_img_height = qr_height // scale_factor
        center_image = center_image.resize((center_img_width, center_img_height), Image.Resampling.LANCZOS)

        pos = ((qr_width - center_img_width) // 2, (qr_height - center_img_height) // 2)
        qr_image.paste(center_image, pos, mask=center_image if center_image.mode == 'RGBA' else None)

    except Exception as e:
        print(f"\nFailed to add center image: {e}\n")
    
    return qr_image

## --------------------------------------------------------------------------
# Function to add a title to the QR Code
def add_title(qr_image, title):
    """
    Adds a title text to the top of the QR code image.

    Args:
        qr_image (PIL.Image.Image): The QR code image.
        title (str): The title text to add.

    Returns:
        PIL.Image.Image: The QR code image with the title added.
    """
    
    try:
        font = ImageFont.truetype(font_file_path, FONT_SIZE)
    except IOError:
        font = ImageFont.load_default()

    draw = ImageDraw.Draw(qr_image)
    title_bbox = draw.textbbox((0, 0), title, font=font)
    title_width = title_bbox[2] - title_bbox[0]
    title_height = title_bbox[3] - title_bbox[1] + 5

    # Create a new image with extra space for the title
    new_image = Image.new("RGB", (qr_image.width, qr_image.height + title_height + 10), "white")
    draw = ImageDraw.Draw(new_image)

    # Draw the title at the top center
    text_position = ((new_image.width - title_width) // 2, 2)
    draw.text(text_position, title, fill="black", font=font)

    # Paste the QR code below the title
    new_image.paste(qr_image, (0, title_height + 10))
    
    return new_image

## --------------------------------------------------------------------------
# Function manage all QR creation tasks
def generate_qrcode():
    """
    Manages the entire QR code generation process including text input, QR code creation, adding a center image, 
    adding a title, and saving the image to a file.

    Returns:
        str: The file path of the saved QR code image.

    Exits:
        Exits the program if an error occurs during QR code creation or saving.
    """
    
    # Get input text from the user
    input_text = get_text()

    # Get the QR Error Correction Level
    # error_correction = input("\nSelect Error Correction level (Low-L, Medium-M, Quartile-Q, High-H): ").upper().strip()
    error_correction = "H"
    
    if error_correction not in ['L','M','Q','H']:
        print("\nInvalid Input! Please select from L, M, Q, H\nQR code creation failed.\n")
        exit(1)
    
    try:
        # Generate QR Image
        qr_image = qr_gen(input_text, error_correction)
        
        # Add center image to the QR code
        qr_image = add_center_image(qr_image)
        
        # Get the title to add
        title = input("Enter the title to add at the top of the QR code (or press Enter to skip): ").strip()
        if title:
            qr_image = add_title(qr_image, title)

        filename = input("\nEnter the filename for the QR code image: ").strip()
        if re.search(FORBIDDEN_CHARS, filename):
            print("\nInvalid filename! QR code creation failed.\nFilename can't contain \\/:*?\"<>| symbols\n")
            return
        
        # Get the extension type
        extension, image_format = extension_menu()
        
        os.makedirs(QRCODES_DIRECTORY_PATH, exist_ok=True)
        
        qr_image_path =  os.path.join(QRCODES_DIRECTORY_PATH, f"{filename}{extension}")

        # Handle if QR Code Image has existing filename
        counter = 1
        while os.path.exists(qr_image_path):
            qr_image_path = f"{QRCODES_DIRECTORY_PATH}{filename}_{counter}{extension}"
            counter += 1

        # Save the QR Code
        qr_image.save(qr_image_path, format=image_format)

        return qr_image_path

    except Exception as e:
        print(f"\nOops! There was an error in creating QR.\n{e}\n")
        exit(1)


### ===========================================================================
## Main 
#

if __name__ == "__main__":

    FORBIDDEN_CHARS = r'[\/:*?"<>|]'
    
    if os.getcwd()[-9:] == "Resources":
        QRCODES_GENERATOR_DIRECTORY_PATH = os.path.join(os.getcwd(), 'QRCode_Generator')

    elif os.getcwd()[-16:] == "QRCode_Generator":
        QRCODES_GENERATOR_DIRECTORY_PATH = os.getcwd()
        
    else:
        print("\nPlease change your working directory to the main repository.\n\nExiting...\n")
        exit(1)

    LOGOS_DIRECTORY_PATH = os.path.join(QRCODES_GENERATOR_DIRECTORY_PATH, 'Logos')
    FONTS_DIRECTORY_PATH = os.path.join(QRCODES_GENERATOR_DIRECTORY_PATH, 'Fonts')
    QRCODES_DIRECTORY_PATH = os.path.join(QRCODES_GENERATOR_DIRECTORY_PATH, 'QRCodes')

    print("\n" + " QR Code Generator ".center(29, "-"))

    font_file = select_font()
    FONT_SIZE = 40
    font_file_path = os.path.join(FONTS_DIRECTORY_PATH, font_file)

    qr_image_path = generate_qrcode()

    print(f"\nQR code \"{qr_image_path[len(QRCODES_GENERATOR_DIRECTORY_PATH) + 1:]}\" created successfully!\n")
