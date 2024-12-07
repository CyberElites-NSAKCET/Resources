# QR Code Generator

A Python-based tool to generate customizable QR codes with features such as different styles, center images, titles, and more. This script provides options for standard or dotted QR codes, custom fonts, and output formats.

---

## Features

- **Customizable QR Code Styles**:
  - Standard QR Code
  - Dotted-style QR Code
- **Flexible QR Code Attributes**:
  - Custom text input (multiline support).
  - Adjustable error correction levels (L, M, Q, H).
  - Background and foreground color selection (Black or White).
- **Enhanced Features**:
  - Add a title above the QR code.
  - Include a center image/logo.
- **Output Formats**:
  - Supported extensions: JPEG, JPG, PNG, GIF, TIFF, BMP.
  - Automatically avoids overwriting existing files.
- **User-Friendly Interface**:
  - Font selection from available TTF files.
  - Interactive prompts for all customizable options.

---

## Directory Structure

Ensure the following directory structure is set up before running the script:

```plaintext
QRCode_Generator/
│
├── Logos/
│   └── <optional_center_images>
│
├── Fonts/
│   └── <font_files>.ttf
│
├── QRCodes/ (created automatically for generated QR codes)
│
└── QRCode_Generator.py
```

- **Logos/**: Holds optional center images (default files provided for white/black backgrounds).
- **Fonts/**: Contains TrueType font files (.ttf) for title text rendering.
- **QRCodes/**: Automatically created for saving generated QR codes.

---

## How to Use

### Prerequisites
- Python 3.6 or higher
- Install the required modules:
  ```bash
  pip install qrcode pillow
  ```

### Running the Script

1. **Navigate to the main directory** containing the script.
2. **Execute the script** using:
   ```bash
   python QRCode_Generator.py
   ```

## Steps for QR Code Generation

1. **Input Text**:
   Enter the text to encode in the QR code (multiline input supported).

2. **Select Options**:
   - Choose QR code style (Standard or Dotted).
   - Select background color (White or Black).

3. **Optional**:
   - Add a center image from the `Logos/` directory.
   - Add a title above the QR code.

4. **Save File**:
   - Specify the filename and select the output format (e.g., JPEG, PNG, etc.).
   - The QR code is saved in the `QRCodes/` directory.

---

## Customization Options

### Title Font
- Add `.ttf` font files to the `Fonts/` directory.
- Select the desired font during the script prompts.

### Center Image
- Add custom images to the `Logos/` directory.
- Ensure the image fits well within the QR code.

### QR Code Styling
- Adjust `box_size`, `border`, or `version` settings in the `standard_qr_gen` and `dots_qr_gen` functions for finer control.

---

## Error Handling

### Missing Directories/Files
- The script automatically creates missing directories (e.g., `QRCodes/`).
- Informative error messages guide corrections for missing or invalid inputs.

### Invalid Filename
- Prohibits filenames with restricted characters (`\/:*?"<>|`).

### Keyboard Interrupts
- Gracefully handles interruptions during execution.

