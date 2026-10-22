"""
VTU Captcha Solver Module
Automatically preprocesses and solves VTU portal captchas using OCR.
Supports native macOS Apple Vision OCR (fast & highly accurate) with seamless
cross-platform fallback to Tesseract OCR via pytesseract.
"""

import os
import sys
import re
import platform
import subprocess
from io import BytesIO
from PIL import Image, ImageOps

try:
    import pytesseract
    HAS_PYTESSERACT = True
    
    # 1. Try pip-installed tesseract-bin if available
    try:
        import tesseract_bin
        if hasattr(tesseract_bin, "TESSERACT_PATH") and os.path.exists(str(tesseract_bin.TESSERACT_PATH)):
            pytesseract.pytesseract.tesseract_cmd = str(tesseract_bin.TESSERACT_PATH)
        if hasattr(tesseract_bin, "TESSDATA_PREFIX") and os.path.exists(str(tesseract_bin.TESSDATA_PREFIX)):
            os.environ["TESSDATA_PREFIX"] = str(tesseract_bin.TESSDATA_PREFIX)
    except ImportError:
        pass

    # 2. Auto-detect Tesseract binary path on Windows if not already in PATH
    if platform.system() == "Windows":
        common_win_paths = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe")
        ]
        for p in common_win_paths:
            if os.path.exists(p):
                pytesseract.pytesseract.tesseract_cmd = p
                break
except ImportError:
    HAS_PYTESSERACT = False

from logger import setup_logger

logger = setup_logger(__name__)

# Path to the compiled macOS Vision OCR helper
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
MAC_VISION_BIN = os.path.join(CURRENT_DIR, "bin", "vtu_vision_ocr")
MAC_VISION_SWIFT = os.path.join(CURRENT_DIR, "bin", "vtu_vision_ocr.swift")


class CaptchaSolver:
    """Solves VTU results portal CAPTCHA images automatically."""

    @staticmethod
    def preprocess_image(image_input, debug_save_path=None):
        """
        Extract the dark grey foreground letters and remove the colorful background word cloud.
        
        VTU captchas consist of:
        - Foreground text: 6 alphanumeric characters in dark grey (RGB ~ 100-110).
        - Background: Multi-colored words ('UNIVERSITY', 'PASS', 'IA', 'NOTES', etc.) in lighter/colored tones.
        
        Args:
            image_input: File path (str), PIL.Image, or raw bytes.
            debug_save_path: Optional path to save the filtered black-on-white image.
            
        Returns:
            Preprocessed PIL Image (mode 'L', black text on white background with border).
        """
        if isinstance(image_input, str):
            img = Image.open(image_input).convert("RGB")
        elif isinstance(image_input, (bytes, bytearray)):
            img = Image.open(BytesIO(image_input)).convert("RGB")
        elif isinstance(image_input, Image.Image):
            img = image_input.convert("RGB")
        else:
            raise ValueError(f"Unsupported image input type: {type(image_input)}")

        w, h = img.size
        # Create a white background image
        filtered = Image.new("L", (w, h), 255)
        
        raw_pixels = img.load()
        out_pixels = filtered.load()

        # Isolate foreground dark gray pixels
        # Foreground letters have low brightness (<130) and low saturation (|r-g| < 22)
        for y in range(h):
            for x in range(w):
                r, g, b = raw_pixels[x, y]
                if max(r, g, b) < 130 and max(abs(r - g), abs(r - b), abs(g - b)) < 22:
                    out_pixels[x, y] = 0  # Black text

        # Add white padding around the text to avoid boundary clipping during OCR
        padded = ImageOps.expand(filtered, border=25, fill=255)

        if debug_save_path:
            try:
                padded.save(debug_save_path)
            except Exception as e:
                logger.warning(f"Failed to save debug preprocessed image: {e}")

        return padded

    @classmethod
    def solve_with_vision(cls, preprocessed_img):
        """Solve using native macOS Apple Vision OCR binary (auto-compiled if needed)."""
        if platform.system() == "Darwin":
            if not os.path.exists(MAC_VISION_BIN) and os.path.exists(MAC_VISION_SWIFT):
                try:
                    subprocess.run(
                        ["swiftc", "-O", MAC_VISION_SWIFT, "-o", MAC_VISION_BIN],
                        capture_output=True,
                        timeout=15
                    )
                except Exception as e:
                    logger.debug(f"Could not auto-compile Vision helper: {e}")

        if not (platform.system() == "Darwin" and os.path.exists(MAC_VISION_BIN)):
            return None

        temp_path = os.path.join(CURRENT_DIR, ".temp_vision_input.png")
        try:
            preprocessed_img.save(temp_path)
            res = subprocess.run(
                [MAC_VISION_BIN, temp_path],
                capture_output=True,
                text=True,
                timeout=5
            )
            if res.returncode == 0:
                raw_text = res.stdout.strip()
                cleaned = re.sub(r"[^A-Za-z0-9]", "", raw_text)
                if cleaned:
                    return cleaned
        except Exception as e:
            logger.debug(f"Vision OCR failed: {e}")
        finally:
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except OSError:
                    pass

        return None

    @classmethod
    def solve_with_tesseract(cls, preprocessed_img):
        """Solve using Tesseract OCR (cross-platform)."""
        if not HAS_PYTESSERACT:
            return None

        # Upscale 2x for better character edge recognition in Tesseract
        scaled = preprocessed_img.resize(
            (preprocessed_img.width * 2, preprocessed_img.height * 2),
            Image.Resampling.LANCZOS
        )

        for psm in [7, 6, 11]:
            try:
                cfg = (
                    f"--psm {psm} "
                    "-c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
                )
                raw_text = pytesseract.image_to_string(scaled, config=cfg)
                cleaned = re.sub(r"[^A-Za-z0-9]", "", raw_text)
                if len(cleaned) == 6:
                    return cleaned
                if cleaned and len(cleaned) in [5, 7]:
                    return cleaned[:6]
            except Exception as e:
                logger.debug(f"Tesseract PSM {psm} failed: {e}")

        return None

    @classmethod
    def solve(cls, image_input, debug_save_path=None):
        """
        Main method to preprocess and solve a VTU captcha image.
        
        Args:
            image_input: File path, PIL Image, or raw bytes.
            debug_save_path: Optional path to save the preprocessed image.
            
        Returns:
            6-character captcha string, or None if detection failed.
        """
        preprocessed = cls.preprocess_image(image_input, debug_save_path=debug_save_path)

        # 1. Try macOS Vision first (fastest and highest accuracy on macOS)
        result = cls.solve_with_vision(preprocessed)
        if result and len(result) == 6:
            logger.info(f"Solved captcha using Apple Vision: {result}")
            return result

        # 2. Try Tesseract OCR
        tess_result = cls.solve_with_tesseract(preprocessed)
        if tess_result:
            logger.info(f"Solved captcha using Tesseract: {tess_result}")
            return tess_result

        # If Vision got a result (even if not strictly 6 chars), return truncated/padded or best effort
        if result:
            logger.warning(f"Using non-standard length Vision result: {result}")
            return result[:6]

        logger.error("All OCR engines failed to solve captcha.")
        return None
