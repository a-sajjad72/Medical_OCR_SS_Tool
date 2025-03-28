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
  - PaddleOCR (auto-downloads models)
  - EasyOCR (auto-downloads models)
  - Tesseract (requires manual installation)
- Confidence-based Excel highlighting
- Cross-platform support (Windows/macOS)
- GUI with image upload and screenshot capture
- Automatic path detection for Tesseract

## Installation

1. **Clone the repository:**

   ```sh
   git clone https://github.com/a-sajjad72/Medical_OCR_SS_Tool.git
   cd Medical_OCR_SS_Tool
   ```

2. **Set up a virtual environment:**

   ```sh
   python -m venv venv
   ```

3. **Activate the virtual environment:**

   - **Windows:**
     ```cmd
     venv\Scripts\activate
     ```
   - **macOS/Linux:**
     ```sh
     source venv/bin/activate
     ```

4. **Install dependencies:**

   The project contains two requirement files:

   - `requirements.txt` – List of all dependencies with their versions. (recommended to use)
   - `requirements_paddle.txt` – list of key dependencies for PaddleOCR.

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

### Automated Build Scripts

1. **Windows**:

   ```bat
   build_windows.bat
   ```

2. **macOS/Linux**:
   ```bash
   source build_mac.sh
   ```

**Script Features**:

- Automatically detects Tesseract installation paths
- Handles both system-wide and custom installations
- Maintains consistent resource paths between dev and prod
- Validates Tesseract presence before building

**Manual Build Requirements**:

```bash
# For reference - use scripts instead
python -c "from utils import get_tessbin_path, get_tessdata_path; print(f'--add-binary {get_tessbin_path()}:models/tesseract --add-data {get_tessdata_path()}:models/tesseract/tessdata')"
```

**Why This Works**:

1. Uses your existing path detection logic from `utils.py`
2. Maintains frozen application structure
3. Eliminates hardcoded paths
4. Automatically adapts to different installations
5. Preserves PyInstaller's resource bundling requirements

## Troubleshooting

### Tesseract Issues

- **Path not found**: Verify installation and check `.env` file
- **Missing languages**: Install tesseract-lang (macOS) or reinstall with additional languages (Windows)
- **Version mismatch**: Requires Tesseract 5.3.0+
  t

### Common Errors

- `TESSDATA_PREFIX not set`: Verify tessdata directory exists
- `No module named...`: Reinstall requirements.txt dependencies
- `Permission denied`: Run as administrator (Windows) or use `sudo` (macOS)

## License

This project is licensed under the Creative Commons License. See the [LICENSE](LICENSE) file for details.
