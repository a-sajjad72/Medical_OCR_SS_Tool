import sys
import os

TESSERACT_PATH = {
    "win32": "./models/tesseract/tesseract.exe",  # adjust it to your Tesseract installation path
    "darwin": "/usr/local/brew-master/bin/tesseract",
}


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def get_tesseract_path():
    if getattr(sys, "frozen", False):
        # We're running in a bundle, so use the bundled Tesseract binary
        return resource_path("models/tesseract")
    else:
        # Development environment: use the system-installed Tesseract
        if sys.platform in TESSERACT_PATH:
            return TESSERACT_PATH[sys.platform]
        else:
            raise RuntimeError(f"Unsupported platform: {sys.platform}")
