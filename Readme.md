# OCR to Excel Converter

OCR to Excel Converter is a Python-based desktop application that processes images using different OCR engines (PaddleOCR, Tesseract, or EasyOCR) and converts the extracted data into an Excel file. The application features a graphical user interface (GUI) built using Tkinter and ttkbootstrap.

## Table of Contents

- [Features](#features)
- [Installation](#installation)
- [Setup](#setup)
- [Usage](#usage)
- [Building the Executable](#building-the-executable)
- [Requirements](#requirements)
- [Project Structure](#project-structure)
- [Troubleshooting](#troubleshooting)
- [License](#license)

## Features

- Support for multiple OCR engines:
  - PaddleOCR (downloaded automatically in the development environment)
  - EasyOCR (downloaded automatically in the development environment)
  - Tesseract via pytesseract (requires manual installation)
- Confidence threshold configuration for output highlighting
- GUI for image upload and screenshot capture
- Automatic processing and Excel output

## Installation

1. **Clone the repository:**

   ```sh
   git clone https://github.com/a-sajjad72/Medical_OCR_SS_Tool.git
   cd Medical_OCR_SS_Tool
   ```

2. **Set up a virtual environment (recommended):**

   ```sh
   python -m venv venv
   ```

3. **Activate the virtual environment:**

   - **Windows:**
     ```sh
     venv\Scripts\activate
     ```
   - **macOS/Linux:**
     ```sh
     source venv/bin/activate
     ```

4. **Install dependencies:**

   The project contains two requirement files:
   
   - `requirements.txt` – Contains a frozen environment with pinned versions.
   - `requirements_paddle.txt` – Lists the dependency names for development.
   
   Use the appropriate file for your needs:
   
   ```sh
   pip install -r requirements.txt
   ```

## Setup

- **Tesseract:**  
  - **Manual Installation Required:** The Tesseract OCR model must be installed manually on your machine.  
  - **Installation Guide:** For detailed installation instructions, please refer to the [Tesseract OCR Installation Guide](https://tesseract-ocr.github.io/tessdoc/Installation.html).  
  - **Custom Path:** If Tesseract is installed in a location different from the default specified in `utils.py` (i.e. `utils.TESSERACT_PATH`), update that path accordingly.

- **Models:**
  - All required models for PaddleOCR and EasyOCR are automatically downloaded when you run the application in the development environment.

## Usage

1. **Run the application:**

   ```sh
   python main.py
   ```

2. **Configure and Process:**
   - Choose the OCR engine using the dropdown menu.
   - Adjust the confidence thresholds (high and medium) as needed.
   - Upload an image or use the screenshot feature.
   - The extracted data will be processed and saved as an Excel file with annotations.

## Building the Executable

The project includes PyInstaller commands to bundle the application into a standalone executable. See the commands below:

### Using PyInstaller on Windows

   ```bat
   pyinstaller --noconfirm --onefile --windowed --add-data "icons:icons" --add-data "models:models" --add-data "simfang.ttf:." --collect-all "paddle" --collect-all "paddleocr" --hidden-import "paddle" --hidden-import "paddleocr" --hidden-import "easyocr" --hidden-import "pytesseract" --icon "icons/icon.png" --name "OCR to Excel Converter" main.py --clean 
   ```

### Using PyInstaller on macOS/Linux

   ```bat
   pyinstaller --noconfirm --onefile --windowed --add-data "icons:icons" --add-data "models:models" --add-data "simfang.ttf:." --add-binary "/usr/local/brew-master/bin/tesseract:models/tesseract" --add-data "/usr/local/brew-master/Cellar/tesseract/5.5.0/share/tessdata:models/tesseract/tessdata" --collect-all "paddle" --collect-all "paddleocr" --hidden-import "paddle" --hidden-import "paddleocr" --hidden-import "easyocr" --hidden-import "pytesseract" --icon "icons/icon.png" --name "OCR to Excel Converter" main.py --clean
   ```

   *Make sure to adjust file paths as necessary for your environment.*

## Requirements

- **Python Version:**  
  The project has been tested with Python **3.12.7**. All libraries including Torch and others are compatible with this version.
  
- **Tesseract:**  
  Ensure Tesseract is installed manually on your machine. Adjust `utils.TESSERACT_PATH` in the code if Tesseract is located in a non-default path.

- **Other Libraries:**  
  The required libraries for PaddleOCR, EasyOCR, and Tesseract integration are listed in the requirements files.

## Project Structure

```
Medical_OCR_SS_Tool/
│
├── __pycache__/             # Compiled Python files (created automatically)
├── build/                   # Build output directory (created by PyInstaller)
├── dist/                    # Distribution output directory (created by PyInstaller)
├── icons/                   # Application icons and images
├── models/                  # OCR model files for PaddleOCR, EasyOCR, and Tesseract (if installed in a custom path)
├── OCR_Modules/             # OCR engine modules (e.g. easyOCR.py, paddleOCR.py, tesseractOCR.py)
├── test/                    # Test files and output (e.g., test1.png, test1.xlsx)
├── test2/                   # Additional test files and output (e.g., test2.png, test2.xlsx)
├── LICENSE                  # License file for the project
├── main.py                  # Main application file containing the OCRApp class
├── README.md                # Project documentation
├── requirements.txt         # Python dependencies for the frozen environment
├── requirements_paddle.txt  # Dependency names for development
└── utils.py                 # Utility functions (e.g. resource_path, get_tesseract_path, etc.)

```

## Troubleshooting

- **Tesseract Installation:**  
  If Tesseract is not found or not working as expected, ensure it is installed manually and that the path in `utils.py` (i.e., `utils.TESSERACT_PATH`) is adjusted accordingly.

- **PyInstaller Build Issues:**  
  If you encounter issues while building with PyInstaller, refer to the PyInstaller commands provided above and verify that all file paths and dependencies are correct for your target OS.

## License

This project is licensed under the Creative Commons License. See the [LICENSE](LICENSE) file for details.
