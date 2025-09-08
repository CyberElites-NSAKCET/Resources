# Certificate Generator

A Python-based script designed to generate personalized certificates by combining a pre-designed template with a sorted list of names. The tool allows extensive customization for fonts, colors, and text positioning, ensuring professional-looking results.

## Features

- **Template-Based Design**: Utilizes a pre-designed PDF template for generating certificates.
- **Font Customization**: Choose from TrueType font files (TTF) in the `Fonts` directory.
- **Advanced Text Styling**: Adjust font size, color, and character spacing for precise rendering.
- **Automated Name Sorting**: Reads names from a text file (`wordlist.txt`) and sorts them alphabetically.
- **Robust Error Handling**: Provides clear error messages for missing files, invalid names, or misconfigurations.
- **Batch Processing**: Generates certificates for all names in the list at once.

---

## Requirements

- **Python Version**: Python 3.6 or higher.
- **Required Modules**:
  - [ReportLab](https://pypi.org/project/reportlab/) (for PDF generation)
  - [PyPDF2](https://pypi.org/project/PyPDF2/) (for handling PDF templates)

Install the dependencies using the command:

```bash
pip install reportlab PyPDF2
```

## Directory Structure

Ensure the following directory structure is set up before running the script:

```plaintext
Fonts/
│
└──<font_files>.ttf

Certificate_Generator/
│
├── certificate_generator.py
│
├── Certificate_Template/
│   └── <template_file>.pdf
│
├── Wordlist/
│   └── wordlist.txt
│
├── tmp/ (created automatically for temporary files)
│
└── Generated_Certificates/ (created automatically for output)
```

- **`Fonts/`**: (in root directory) Stores TrueType font files (.ttf) for text rendering.
- **`Certificate_Template/`**: Holds the certificate template PDF.  
- **`Wordlist/`**: Contains a text file (`wordlist.txt`) with names (one name per line).  
- **`tmp/`**: Temporary folder for intermediate files (auto-created).  
- **`Generated_Certificates/`**: Folder where the final certificates are saved (auto-created).  

---

## How to Use
> **💡Run the script once to create all necessary files and directories:**

1) Clone the repo
   ```bash
   git clone https://github.com/CyberElites-NSAKCET/Resources
   ```
2) Change the working directory to the Certificate Generator directory
   ```bash
   cd Resources
   cd Certificate_Generator 
   ```
3) Run the script once to create all the files and directories
    ```bash
    python certificate_generator.py
    ```

### Setup the Directory:
1. Place a pre-designed PDF template in the `Certificate_Template/` directory.  
2. Add (one per line) to `wordlist.txt` in the `Wordlist/` directory.  
3. Add any `.ttf` font files to the `Fonts/` directory.  

### Run the Script:
Execute the script using the command:  
   ```bash
   python certificate_generator.py
   ```

### Follow Prompts:
1. Select the font from the displayed list.  
2. The script will process the names and generate certificates in the `Generated_Certificates/` directory.  

---

## Customization Options

- **Font Settings**:  
  Modify `FONT_SIZE` and `FONT_COLOR` variables in the script for custom styling.  

- **Text Positioning**:  
  Adjust the `POSITION` tuple to fine-tune text placement on the template.  

- **Character Spacing**:  
  Change the `CHAR_SPACING` variable for precise character alignment.  

---

## Error Handling

- **Multiple/No Files in Directories**:  
  The script will notify and prompt corrections if there are no files or multiple files in the required directories.  

- **Invalid Characters in Names**:  
  Names containing invalid characters (`<>\"?|/\\:*`) will be flagged as errors.  

- **Missing Directories or Files**:  
  The script will automatically create missing directories and provide informative error messages.  

