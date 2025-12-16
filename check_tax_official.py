"""
Tax Checker - Official API Integration Module

This module handles checking CCCD (Citizen ID) against the official tax portal
using Selenium WebDriver and EasyOCR for captcha solving.

Resources (bundled when running as PyInstaller app):
- ChromeDriver: For browser automation
- EasyOCR Models: For captcha recognition
"""
import os
import sys
import time
import warnings
from io import BytesIO

import numpy as np
from PIL import Image
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException
import easyocr

# Suppress SSL warnings
warnings.filterwarnings("ignore")

# =============================================================================
# Global State
# =============================================================================
_ocr_reader = None
_ocr_error = None


# =============================================================================
# Resource Path Utilities
# =============================================================================
def get_bundle_dir():
    """Get the base directory for bundled resources."""
    if getattr(sys, 'frozen', False):
        # Running as PyInstaller bundle
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_bundled_chromedriver():
    """
    Get the path to bundled ChromeDriver if available.
    
    Returns:
        str or None: Path to chromedriver executable, or None if not bundled.
    """
    bundle_dir = get_bundle_dir()
    
    if sys.platform == 'win32':
        driver_name = 'chromedriver.exe'
    else:
        driver_name = 'chromedriver'
    
    # Possible locations in bundled app
    possible_paths = [
        os.path.join(bundle_dir, '_internal', 'chromedriver', driver_name),
        os.path.join(bundle_dir, 'chromedriver', driver_name),
    ]
    
    # For macOS .app bundle
    if sys.platform == 'darwin':
        possible_paths.insert(0, os.path.join(bundle_dir, '..', 'Frameworks', 'chromedriver', driver_name))
        possible_paths.insert(1, os.path.join(bundle_dir, '..', 'Resources', 'chromedriver', driver_name))
    
    for path in possible_paths:
        if os.path.isfile(path):
            return os.path.abspath(path)
    
    return None


def get_bundled_model_dir():
    """
    Get the path to bundled EasyOCR models directory if available.
    
    Returns:
        str or None: Path to model directory, or None if not bundled.
    """
    bundle_dir = get_bundle_dir()
    
    possible_paths = [
        os.path.join(bundle_dir, '_internal', 'easyocr_models'),
        os.path.join(bundle_dir, 'easyocr_models'),
    ]
    
    # For macOS .app bundle
    if sys.platform == 'darwin':
        possible_paths.insert(0, os.path.join(bundle_dir, '..', 'Frameworks', 'easyocr_models'))
        possible_paths.insert(1, os.path.join(bundle_dir, '..', 'Resources', 'easyocr_models'))
    
    for path in possible_paths:
        if os.path.isdir(path):
            return os.path.abspath(path)
    
    return None


# =============================================================================
# Logging
# =============================================================================
def log(message, callback=None):
    """Log a message to callback or stdout."""
    if callback:
        callback(message)
    else:
        print(message)


# =============================================================================
# OCR Functions
# =============================================================================
def get_ocr_reader():
    """
    Get or initialize the EasyOCR reader.
    
    Uses bundled models if available, otherwise downloads from web.
    
    Returns:
        easyocr.Reader: The OCR reader instance.
        
    Raises:
        RuntimeError: If OCR initialization fails.
    """
    global _ocr_reader, _ocr_error
    
    if _ocr_error:
        raise _ocr_error
    
    if _ocr_reader is None:
        try:
            model_dir = get_bundled_model_dir()
            
            if model_dir:
                # Use bundled models - disable download
                _ocr_reader = easyocr.Reader(
                    ['en'],
                    gpu=False,
                    model_storage_directory=model_dir,
                    download_enabled=False
                )
            else:
                # Fallback: allow download from internet
                _ocr_reader = easyocr.Reader(['en'], gpu=False)
                
        except Exception as e:
            error_msg = str(e)
            if "urlopen error" in error_msg or "connection" in error_msg.lower():
                _ocr_error = RuntimeError(
                    "Failed to initialize OCR. Models not bundled and download failed.\n"
                    "Please check your internet connection or use bundled app version."
                )
                raise _ocr_error
            raise
    
    return _ocr_reader


# =============================================================================
# WebDriver Setup
# =============================================================================
def setup_driver(open_browser=False, log_callback=None):
    """
    Set up Chrome WebDriver.
    
    Priority:
    1. Bundled ChromeDriver (for offline operation)
    2. webdriver-manager (downloads if needed)
    
    Args:
        open_browser: If True, run in visible mode; otherwise headless.
        log_callback: Optional callback for logging messages.
        
    Returns:
        webdriver.Chrome: The configured Chrome WebDriver instance.
        
    Raises:
        RuntimeError: If driver setup fails.
    """
    log(f"Setting up Chrome driver (Headless: {not open_browser})...", log_callback)
    
    options = webdriver.ChromeOptions()
    if not open_browser:
        options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
    
    # Method 1: Try bundled ChromeDriver
    bundled_driver = get_bundled_chromedriver()
    if bundled_driver:
        try:
            log(f"Using bundled ChromeDriver: {bundled_driver}", log_callback)
            service = Service(executable_path=bundled_driver)
            driver = webdriver.Chrome(service=service, options=options)
            log("Chrome driver setup complete (bundled).", log_callback)
            return driver
        except WebDriverException as e:
            log(f"Bundled ChromeDriver failed: {str(e)[:100]}...", log_callback)
    
    # Method 2: Fallback to webdriver-manager (requires internet)
    try:
        log("Falling back to webdriver-manager (may download)...", log_callback)
        from webdriver_manager.chrome import ChromeDriverManager
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        log("Chrome driver setup complete (webdriver-manager).", log_callback)
        return driver
    except Exception as e:
        error_msg = str(e)
        raise RuntimeError(
            f"Could not set up ChromeDriver.\n"
            f"Error: {error_msg}\n\n"
            "Please ensure Chrome browser is installed, or use bundled app version."
        ) from e


# =============================================================================
# Captcha Solving
# =============================================================================
def solve_captcha(driver, log_callback=None):
    """
    Solve captcha on the current page using OCR.
    
    Args:
        driver: The Selenium WebDriver instance.
        log_callback: Optional callback for logging messages.
        
    Returns:
        str or None: The solved captcha text, or None if failed.
    """
    try:
        log("Attempting to solve captcha...", log_callback)
        
        # Find captcha image
        images = driver.find_elements(By.TAG_NAME, "img")
        captcha_img = None
        for img in images:
            src = img.get_attribute("src")
            if src and ("captcha" in src.lower() or "jcaptcha" in src.lower()):
                captcha_img = img
                break
        
        if not captcha_img:
            log("Captcha image not found.", log_callback)
            return None
        
        # Capture screenshot as PNG bytes (in memory, no file I/O)
        png_bytes = captcha_img.screenshot_as_png
        image = Image.open(BytesIO(png_bytes))
        image = image.convert('L')
        
        # Resize to make it bigger (3x)
        image = image.resize((image.width * 3, image.height * 3), Image.Resampling.LANCZOS)
        
        # Thresholding
        threshold = 140
        image = image.point(lambda x: 0 if x < threshold else 255)
        
        # Convert to numpy array for EasyOCR
        image_array = np.array(image)
        
        # Use EasyOCR
        reader = get_ocr_reader()
        results = reader.readtext(image_array, allowlist='abcdefghijklmnopqrstuvwxyz0123456789')
        
        # Combine all detected text
        result_text = ''.join([text for (_, text, _) in results]).strip().lower()
        log(f"Captcha solved: '{result_text}'", log_callback)
        return result_text
        
    except Exception as e:
        log(f"Captcha error: {e}", log_callback)
        return None


# =============================================================================
# Main Tax Check Function
# =============================================================================
def check_cccd_official(cccd, open_browser=False, log_callback=None, max_retries=20, delay_seconds=2):
    """
    Check a CCCD number against the official tax portal.
    
    Args:
        cccd: The citizen ID number to check.
        open_browser: If True, show browser window; otherwise headless.
        log_callback: Optional callback for logging messages.
        max_retries: Maximum number of captcha retry attempts.
        delay_seconds: Delay between retries.
        
    Returns:
        dict: Result containing tax_id, name, place, and status.
    """
    driver = None
    result = {
        "cccd": cccd,
        "tax_id": None,
        "name": None,
        "place": None,
        "status": "Not Found",
    }
    
    try:
        driver = setup_driver(open_browser=open_browser, log_callback=log_callback)
        
        url = "https://dichvucong.gdt.gov.vn/p/home/lookup-tin.html"
        log(f"Opening {url}", log_callback)
        driver.get(url)
        
        # Wait for page and input field
        wait = WebDriverWait(driver, 10)
        input_field = wait.until(EC.presence_of_element_located((By.ID, "searchId")))
        
        # Clear + type CCCD
        input_field.clear()
        input_field.send_keys(cccd)
        log(f"Entered CCCD: {cccd}", log_callback)
        
        for attempt in range(max_retries):
            log(f"Attempt {attempt + 1}/{max_retries}...", log_callback)
            
            # Solve captcha
            captcha_answer = solve_captcha(driver, log_callback)
            if not captcha_answer:
                time.sleep(delay_seconds)
                driver.refresh()
                input_field = wait.until(EC.presence_of_element_located((By.ID, "searchId")))
                input_field.clear()
                input_field.send_keys(cccd)
                continue
            
            # Enter captcha
            captcha_input = driver.find_element(By.ID, "captchaAnswer")
            captcha_input.clear()
            captcha_input.send_keys(captcha_answer)
            
            # Click submit
            submit_btn = driver.find_element(By.XPATH, "//button[@type='submit']")
            submit_btn.click()
            
            time.sleep(2)
            
            # Check for result table
            try:
                table = driver.find_element(By.CLASS_NAME, "table-striped")
                rows = table.find_elements(By.TAG_NAME, "tr")
                
                for row in rows:
                    cells = row.find_elements(By.TAG_NAME, "td")
                    if len(cells) >= 4:
                        result["tax_id"] = cells[0].text.strip()
                        result["name"] = cells[1].text.strip()
                        result["place"] = cells[2].text.strip()
                        result["status"] = "Found"
                        log(f"Found: Tax ID={result['tax_id']}, Name={result['name']}", log_callback)
                        return result
                        
            except Exception:
                pass
            
            # Check for error message (wrong captcha)
            try:
                error = driver.find_element(By.CLASS_NAME, "text-danger")
                if error and error.text:
                    log(f"Error: {error.text}", log_callback)
            except Exception:
                pass
            
            time.sleep(delay_seconds)
            driver.refresh()
            input_field = wait.until(EC.presence_of_element_located((By.ID, "searchId")))
            input_field.clear()
            input_field.send_keys(cccd)
        
        log(f"Max retries reached for CCCD: {cccd}", log_callback)
        result["status"] = "Max Retries"
        
    except Exception as e:
        log(f"Error checking CCCD {cccd}: {e}", log_callback)
        result["status"] = f"Error: {str(e)}"
        
    finally:
        if driver:
            driver.quit()
    
    return result
