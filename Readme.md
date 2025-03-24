# OCR to Excel Converter

OCR to Excel Converter is a Python-based desktop application that processes images using different OCR engines (PaddleOCR, Tesseract, or EasyOCR) and converts the extracted data into an Excel file. The application features a graphical user interface (GUI) built using Tkinter and ttkbootstrap.

## Table of Contents

- [Features](#features)
- [Installation](#installation)
- [Requirements](#requirements)
- [Tesseract Installation](#tesseract-installation)
- [Setup](#setup)
- [Usage](#usage)
- [Building the Executable](#building-the-executable)
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

## Requirements

### Python Interpreter

- **Requires Python 3.12** (3.12.6 or newer recommended)
- **Minimum Supported Version**: Python 3.12.0
- **Tested Version**: Python 3.12.6
- **Download**: [python.org/downloads](https://www.python.org/downloads/)

**Important Notes**:
⚠️ **PyTorch Compatibility**: The application requires PyTorch, which currently only has official support for Python 3.12. Earlier versions (3.11 or below) are **not supported** due to dependency conflicts.

💡 **Installation Tips**:

- For Windows users: Check "Add python.exe to PATH" during installation
- For macOS/Linux users: Consider using [pyenv](https://github.com/pyenv/pyenv) for version management
- Verify installation: `python --version`

🔗 **PyTorch Compatibility Reference**:  
[Official PyTorch Python Support Matrix](https://pytorch.org/get-started/previous-versions/#python-compatibility)

### Tesseract OCR

- **Required for Tesseract OCR engine**
- **Version 5.3.0** or newer
- Installation guides below

### Python Dependencies

All Python package requirements are listed in `requirements.txt`. Key dependencies include:

- PaddleOCR
- EasyOCR
- pytesseract
- OpenCV
- ttkbootstrap
- openpyxl
- pytorch

## Tesseract Installation

### macOS

```bash
# Install using Homebrew
brew install tesseract

# Verify installation
tesseract --version
```

### Windows

1. Download installer from [UB-Mannheim Tesseract](https://github.com/UB-Mannheim/tesseract/wiki)
2. Run the installer with default settings:
   - Use recommended installation path i.e. (`C:\Program Files\Tesseract-OCR`)

### Custom Installations

Create `.env` file in project root for custom paths:

```env
TESS_BINARY_PATH=/path/to/tesseract
TESSDATA_PREFIX=/path/to/tessdata
```

## Setup

1. **Configuration:**

   - For non-standard Tesseract installations, edit the `.env` file
   - Models for PaddleOCR/EasyOCR will auto-download on first run

2. **Verify Paths:**
   ```sh
   python -c "from utils import get_tessbin_path, get_tessdata_path; print(f'Tesseract: {get_tessbin_path()}\nTessdata: {get_tessdata_path()}')"
   ```


## Usage

1. **Run the application:**

   ```sh
   python main.py
   ```

2. **Application workflow:**
   - Select OCR engine from dropdown
   - Adjust confidence thresholds (High/Medium)
   - Upload image or capture screenshot
   - Processed Excel file saves automatically
   - Results shown with bounding box visualization

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

## Troubleshooting

### Tesseract Issues

- **Path not found**: Verify installation and check `.env` file
- **Missing languages**: Install tesseract-lang (macOS) or reinstall with additional languages (Windows)
- **Version mismatch**: Requires Tesseract 5.3.0+

### Common Errors

- `TESSDATA_PREFIX not set`: Verify tessdata directory exists
- `No module named...`: Reinstall requirements.txt dependencies
- `Permission denied`: Run as administrator (Windows) or use `sudo` (macOS)

## License

This project is licensed under the Creative Commons License. See the [LICENSE](LICENSE) file for details.
