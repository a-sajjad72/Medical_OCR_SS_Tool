import datetime
import locale
import logging
import logging.handlers
import os
import subprocess
import sys
from pathlib import Path
from shutil import which

from dotenv import load_dotenv

load_dotenv()  # Load .env file if exists
logger = logging.getLogger()


class ErrorSessionHandler(logging.handlers.TimedRotatingFileHandler):
    """Custom handler that adds error session header on first error"""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.session_header_logged = False  # Track session state

    def emit(self, record):
        # Add header before first error
        if record.levelno >= logging.ERROR and not self.session_header_logged:
            self.session_header_logged = True
            self._log_session_header()

        super().emit(record)

    def _log_session_header(self):
        header = (
            "\n"
            + "=" * 80
            + "\n"
            + f"NEW ERROR SESSION: {datetime.datetime.now()}\n"
            + "=" * 80
            + "\n"
        )
        header_record = logging.makeLogRecord(
            {
                "msg": header,
                "levelno": logging.INFO,  # Log as INFO to avoid infinite loop
                "levelname": "INFO",
                "name": "ErrorSessionHeader",
            }
        )
        super().emit(header_record)


def handle_uncaught_exception(exc_type, exc_value, exc_traceback):
    """Handles uncaught exceptions"""
    logger.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))
    sys.__excepthook__(exc_type, exc_value, exc_traceback)


def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def get_tessbin_path():
    # Frozen application path (PyInstaller)
    if getattr(sys, "frozen", False):
        return resource_path(
            "models/tesseract/tesseract.exe"
            if sys.platform == "win32"
            else "models/tesseract/tesseract"
        )

    # Development paths
    # Check .env first for manual override
    if env_path := os.getenv("TESS_BINARY_PATH"):
        return env_path

    # Platform-specific detection
    if sys.platform == "darwin":
        try:
            # Try Homebrew installation path first
            result = subprocess.run(
                ["brew", "--prefix", "tesseract"],
                capture_output=True,
                text=True,
                check=True,
            )
            brew_prefix = result.stdout.strip()
            tesseract_bin = Path(brew_prefix) / "bin" / "tesseract"
            if tesseract_bin.exists():
                return str(tesseract_bin)
        except (subprocess.CalledProcessError, FileNotFoundError):
            pass  # Continue to fallback methods

        # Fallback to which command
        if tess_bin := which("tesseract"):
            return tess_bin

        # Final fallback to common install locations
        common_paths = [
            "/usr/local/bin/tesseract",
            "/opt/homebrew/bin/tesseract",
            os.path.expanduser("~/homebrew/bin/tesseract"),
        ]
        for path in common_paths:
            if os.path.exists(path):
                return path

        raise FileNotFoundError(
            "Tesseract not found. Install with 'brew install tesseract' or "
            "set TESS_BINARY_PATH in .env file"
            "TESS_BINARY_PATH=/path/to/tesseract"
        )

    elif sys.platform == "win32":
        # Windows path detection
        win_paths = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            os.path.expanduser(r"~\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"),
        ]
        for path in win_paths:
            if os.path.exists(path):
                return path

        raise FileNotFoundError(
            "Tesseract not found. Install from https://github.com/UB-Mannheim/tesseract/wiki "
            "or set TESS_BINARY_PATH in .env file"
            "TESS_BINARY_PATH=C:\\path\\to\\tesseract.exe"
        )

    else:
        raise RuntimeError(f"Unsupported platform: {sys.platform}")


def get_tessdata_path():
    # Frozen application path (PyInstaller)
    if getattr(sys, "frozen", False):
        return resource_path("models/tesseract/tessdata")

    # Development paths
    if env_path := os.getenv("TESSDATA_PREFIX"):
        return env_path

    if sys.platform == "darwin":
        try:
            # Try Homebrew installation path first
            result = subprocess.run(
                ["brew", "--prefix", "tesseract"],
                capture_output=True,
                text=True,
                check=True,
            )
            brew_prefix = result.stdout.strip()
            tessdata_path = Path(brew_prefix) / "share" / "tessdata"
            if tessdata_path.exists():
                return str(tessdata_path)
        except (subprocess.CalledProcessError, FileNotFoundError):
            pass

        # Fallback to common locations
        common_paths = [
            "/usr/local/share/tessdata",
            "/opt/homebrew/share/tessdata",
            os.path.expanduser("~/homebrew/share/tessdata"),
        ]
        for path in common_paths:
            if os.path.exists(path):
                return path

        raise FileNotFoundError(
            "Tessdata not found. Install languages with 'brew install tesseract-lang' "
            "or set TESSDATA_PREFIX in .env file"
        )

    elif sys.platform == "win32":
        # Get Tesseract installation directory
        tess_bin = get_tessbin_path()
        tessdata_path = os.path.join(os.path.dirname(tess_bin), "tessdata")

        if os.path.exists(tessdata_path):
            return tessdata_path

        raise FileNotFoundError(
            f"Tessdata not found at {tessdata_path}. "
            "Reinstall Tesseract with language data or set TESSDATA_PREFIX in .env file"
        )

    else:
        raise RuntimeError(f"Unsupported platform: {sys.platform}")

def ensure_locale():
    """Nuclear option for ttkbootstrap locale conflicts"""
    try:
        # Completely override environment variables first
        os.environ["LC_ALL"] = "en_US.UTF-8"
        os.environ["LANG"] = "en_US.UTF-8"
        os.environ["LC_CTYPE"] = "en_US.UTF-8"
        os.environ["LC_TIME"] = "en_US.UTF-8"
        
        # Force reset Python's internal locale tracking
        locale.setlocale(locale.LC_ALL, "en_US.UTF-8")
        locale.setlocale(locale.LC_TIME, "en_US.UTF-8")
        
    except locale.Error:
        try:
            # Fallback to C.UTF-8 if available
            os.environ["LC_ALL"] = "C.UTF-8"
            locale.setlocale(locale.LC_ALL, "C.UTF-8")
        except:
            # Final desperate fallback
            os.environ["LC_ALL"] = "C"
            locale.setlocale(locale.LC_ALL, "")
    
    print("Final locale settings:", locale.getlocale())
    