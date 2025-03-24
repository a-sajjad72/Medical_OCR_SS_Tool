@echo off
FOR /F "tokens=*" %%a IN ('python -c "from utils import get_tessbin_path, get_tessdata_path; print(get_tessbin_path() + ';' + get_tessdata_path())"') DO SET "PATHS=%%a"

pyinstaller --noconfirm --onefile --windowed ^
--add-data "icons;icons" ^
--add-data "models;models" ^
--add-data "simfang.ttf;." ^
--add-binary "%PATHS:;=;models/tesseract/tessdata;models/tesseract% ^
--collect-all paddle ^
--collect-all paddleocr ^
--hidden-import paddle ^
--hidden-import paddleocr ^
--hidden-import easyocr ^
--hidden-import pytesseract ^
--icon "icons/icon.png" ^
--name "OCR to Excel Converter" ^
main.py --clean
