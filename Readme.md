# OCR to Excel Converter

OCR to Excel Converter is a Python-based desktop application that processes images using different OCR engines (PaddleOCR, Tesseract, or EasyOCR) and converts the extracted data into an Excel file. The application features a graphical user interface (GUI) built using Tkinter and ttkbootstrap.

## Table of Contents

- [Features](#features)
- [Installation](#installation)
- [Setup](#setup)
- [Usage](#usage)
- [Building the Executable](#building-the-executable)
- [Project Structure](#project-structure)
- [Troubleshooting](#troubleshooting)
- [License](#license)

## Features

- Support for multiple OCR engines:
  - PaddleOCR
  - EasyOCR
  - Tesseract via pytesseract (requres manual installation)
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
  - The Tesseract OCR model must be installed manually. The application will download PaddleOCR, EasyOCR, and other necessary models automatically in the development environment.
  - If Tesseract is installed in a location different from the default specified in `utils.py` (i.e. `utils.TESSERACT_PATH`), make sure to update that path accordingly.

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

## Project Structure

```
Medical_OCR_SS_Tool/
│
├── build/                   # Build output directory (created by PyInstaller)
├── icons/                   # Application icons and images
├── models/                  # OCR model files for PaddleOCR, EasyOCR and Tesseract (if installed in the custom path)
├── OCR_Modules/             # OCR engine modules (e.g. [`easyOCR.py`](Medical_OCR_SS_Tool/OCR_Modules/easyOCR.py), [`paddleOCR.py`](Medical_OCR_SS_Tool/OCR_Modules/paddleOCR.py), [`tesseractOCR.py`](Medical_OCR_SS_Tool/OCR_Modules/tesseractOCR.py))
├── __pycache__/
├── main.py                  # Main application file containing [`OCRApp`](d:/paddle-ocr/Medical_OCR_SS_Tool/main.py) class
├── requirements.txt         # Python dependencies for the project
├── requirements_paddle.txt  # Additional dependencies for PaddleOCR
└── utils.py                 # Utility functions (e.g. [`resource_path`](Medical_OCR_SS_Tool/utils.py))
```

## Troubleshooting

- **Tesseract Installation:**  
  If Tesseract is not found or not working as expected, ensure it is installed manually and that the path in `utils.py` (i.e., `utils.TESSERACT_PATH`) is adjusted accordingly.

- **PyInstaller Build Issues:**  
  If you encounter issues while building with PyInstaller, refer to the PyInstaller commands provided above and verify that all file paths and dependencies are correct for your target OS.

## License

This project is licensed under the Creative Commons License. See the [LICENSE](LICENSE) file for details.
