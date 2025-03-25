@echo off
setlocal enabledelayedexpansion

:: Get paths using Python utils
FOR /F "tokens=1,2 delims=;" %%a IN ('python -c "from utils import get_tessbin_path, get_tessdata_path; print(get_tessbin_path() + ';' + get_tessdata_path())"') DO (
    SET "TESS_BIN=%%a"
    SET "TESS_DATA=%%b"
)

:: Verify paths
echo Tesseract Binary: !TESS_BIN!
echo Tesseract Data: !TESS_DATA!

:: Build command
pyinstaller --noconfirm --onefile --windowed ^
--add-data "icons;icons" ^
--add-data "models;models" ^
--add-data "simfang.ttf;." ^
--add-binary "!TESS_BIN!;models/tesseract/tesseract.exe" ^
--add-data "!TESS_DATA!;models/tesseract/tessdata" ^
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