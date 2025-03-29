@echo off
setlocal enabledelayedexpansion

:: Get paths using Python utils
FOR /F "delims=" %%a IN ('python -c "from utils import get_tessbin_path, get_tessdata_path; import os; print(os.path.dirname(get_tessbin_path()))"') DO (
    SET "TESS_PATH=%%a"
)

:: Verify paths
echo Tesseract Binary: !TESS_PATH!

:: Build command
pyinstaller --noconfirm --onefile --windowed ^
--add-data "icons;icons" ^
--add-data "models;models" ^
--add-data "simfang.ttf;." ^
--add-data "!TESS_PATH!;models/tesseract" ^
--collect-all paddle ^
--collect-all paddleocr ^
--hidden-import paddle ^
--hidden-import paddleocr ^
--hidden-import easyocr ^
--hidden-import pytesseract ^
--icon "icons/icon.png" ^
--name "OCR to Excel Converter" ^
main.py --clean

endlocal