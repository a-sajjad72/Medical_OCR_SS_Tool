#!/bin/bash

# Get paths using Python utils
TESS_BIN=$(python -c "from utils import get_tessbin_path; print(get_tessbin_path())")
TESS_DATA=$(python -c "from utils import get_tessdata_path; print(get_tessdata_path())")

pyinstaller --noconfirm --onedir --windowed \
--add-data "icons:icons" \
--add-data "models:models" \
--add-data "simfang.ttf:." \
--add-binary "$TESS_BIN:models/tesseract" \
--add-data "$TESS_DATA:models/tesseract/tessdata" \
--collect-all paddle \
--collect-all paddleocr \
--hidden-import paddle \
--hidden-import paddleocr \
--hidden-import easyocr \
--hidden-import pytesseract \
--icon "icons/icon.icns" \
--name "OCR to Excel Converter" \
main.py --clean
