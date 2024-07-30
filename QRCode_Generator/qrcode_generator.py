try:
    import qrcode
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    print("This script requires the 'qrcode' and 'pillow' modules.\nPlease install them using 'pip install qrcode pillow' and try again.")
    exit(1)
import os
import re


## ===========================================================================
### Functions

# Function to input QR text
def get_text():

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
# Function to generate QR
def qr_gen(input_text, error_correction):
    qr = qrcode.QRCode(
        version=1,
        error_correction=getattr(qrcode.constants, f"ERROR_CORRECT_{error_correction}"),
        box_size=10,
        border=2,
    )

    qr.add_data(input_text)
    qr.make(fit=True)
    
    qr_image = qr.make_image(fill_color="black", back_color="white").convert('RGB')

    return qr_image


## --------------------------------------------------------------------------
# Function to get QR Image extension type
def extension_menu():
    
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
    font_file = "Open Sans Regular.ttf"
    font_file_path = os.path.join(FONTS_DIRECTORY_PATH, font_file)
    try:
        font = ImageFont.truetype(font_file_path, 30)
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
            print("\nInvalid filename! QR code creation failed.\nFilename can't contain \/:*?\"<>| symbols\n")
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
    QRCODES_GENERATOR_DIRECTORY = 'QRCode_Generator'
    LOGOS_DIRECTORY = 'Logos'
    QRCODES_DIRECTORY = 'QRCodes'
    FONTS_DIRECTORY = 'FONTS'
    
    if os.getcwd()[-9:] == "Resources":
        QRCODES_GENERATOR_DIRECTORY_PATH = os.path.join(os.getcwd(), QRCODES_GENERATOR_DIRECTORY)
        LOGOS_DIRECTORY_PATH = os.path.join(QRCODES_GENERATOR_DIRECTORY_PATH, LOGOS_DIRECTORY)
        FONTS_DIRECTORY_PATH = os.path.join(QRCODES_GENERATOR_DIRECTORY_PATH, FONTS_DIRECTORY)
        
    elif os.getcwd()[-16:] == "QRCode_Generator":
        QRCODES_GENERATOR_DIRECTORY_PATH = os.getcwd()
        LOGOS_DIRECTORY_PATH = os.path.join(os.getcwd(), LOGOS_DIRECTORY)
        FONTS_DIRECTORY_PATH = os.path.join(os.getcwd(), FONTS_DIRECTORY)
        
    else:
        print("\nPlease change your working directory to the main repository.\nExiting...\n")
        exit(1)
        
    QRCODES_DIRECTORY_PATH = os.path.join(QRCODES_GENERATOR_DIRECTORY_PATH, QRCODES_DIRECTORY)

    print("\n" + " QR Code Generator ".center(29, "-"))

    qr_image_path = generate_qrcode()

    print(f"\nQR code \"{qr_image_path[len(QRCODES_DIRECTORY_PATH):]}\" created successfully!\n")
