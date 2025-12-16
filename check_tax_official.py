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
import csv
import warnings

from PIL import Image
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException
from webdriver_manager.chrome import ChromeDriverManager
import easyocr

# Suppress SSL warnings
warnings.filterwarnings("ignore")

# =============================================================================
# Global State
# =============================================================================
_ocr_reader = None


# =============================================================================
# Resource Path Utilities
# =============================================================================
def get_bundle_dir():
    """Get the base directory for bundled resources."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_bundled_chromedriver():
    """Get the path to bundled ChromeDriver if available."""
    bundle_dir = get_bundle_dir()
    driver_name = 'chromedriver.exe' if sys.platform == 'win32' else 'chromedriver'
    
    possible_paths = [
        os.path.join(bundle_dir, '_internal', 'chromedriver', driver_name),
        os.path.join(bundle_dir, 'chromedriver', driver_name),
    ]
    
    if sys.platform == 'darwin':
        possible_paths.insert(0, os.path.join(bundle_dir, '..', 'Frameworks', 'chromedriver', driver_name))
        possible_paths.insert(1, os.path.join(bundle_dir, '..', 'Resources', 'chromedriver', driver_name))
    
    for path in possible_paths:
        if os.path.isfile(path):
            return os.path.abspath(path)
    return None


def get_bundled_model_dir():
    """Get the path to bundled EasyOCR models directory if available."""
    bundle_dir = get_bundle_dir()
    
    possible_paths = [
        os.path.join(bundle_dir, '_internal', 'easyocr_models'),
        os.path.join(bundle_dir, 'easyocr_models'),
    ]
    
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
    """Get or initialize the EasyOCR reader."""
    global _ocr_reader
    
    if _ocr_reader is None:
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
    
    return _ocr_reader


# =============================================================================
# WebDriver Setup
# =============================================================================
def setup_driver(open_browser=False, log_callback=None):
    """
    Set up Chrome WebDriver.
    
    Priority:
    1. Bundled ChromeDriver (for offline operation)
    2. webdriver-manager (downloads correct version if needed)
    """
    log(f"Setting up Chrome driver (Headless: {not open_browser})...", log_callback)
    
    options = webdriver.ChromeOptions()
    if not open_browser:
        options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    
    # Method 1: Try bundled ChromeDriver first (for offline/packaged app)
    bundled_driver = get_bundled_chromedriver()
    if bundled_driver:
        try:
            log(f"Trying bundled ChromeDriver: {bundled_driver}", log_callback)
            service = Service(executable_path=bundled_driver)
            driver = webdriver.Chrome(service=service, options=options)
            log("Chrome driver setup complete (bundled).", log_callback)
            return driver
        except WebDriverException as e:
            log(f"Bundled ChromeDriver failed: {str(e)[:100]}...", log_callback)
    
    # Method 2: Fallback to webdriver-manager (auto-downloads correct version)
    log("Using webdriver-manager (may download if needed)...", log_callback)
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    log("Chrome driver setup complete.", log_callback)
    return driver


# =============================================================================
# Captcha Solving
# =============================================================================
def solve_captcha(driver, log_callback=None):
    """Solve captcha on the current page using OCR."""
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
        
        # Capture screenshot of the captcha
        captcha_path = "current_captcha.png"
        captcha_img.screenshot(captcha_path)
        
        # Process image for better OCR
        image = Image.open(captcha_path)
        image = image.convert('L')
        
        # Resize to make it bigger (3x)
        image = image.resize((image.width * 3, image.height * 3), Image.Resampling.LANCZOS)
        
        # Thresholding
        threshold = 140
        image = image.point(lambda x: 0 if x < threshold else 255)
        
        # Save processed image for EasyOCR
        processed_path = "current_captcha_processed.png"
        image.save(processed_path)
        
        # Use EasyOCR
        reader = get_ocr_reader()
        results = reader.readtext(processed_path, allowlist='abcdefghijklmnopqrstuvwxyz0123456789')
        
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
    """Check a CCCD number against the official tax portal."""
    driver = None
    result = {
        "cccd": cccd,
        "tax_id": None,
        "name": None,
        "place": None,
        "status": "Not Found",
    }
    
    try:
        log(f"Starting check for CCCD: {cccd}", log_callback)
        driver = setup_driver(open_browser, log_callback)
        url = "https://tracuunnt.gdt.gov.vn/tcnnt/mstcn.jsp"
        log(f"Navigating to {url}", log_callback)
        driver.get(url)
        
        for attempt in range(max_retries):
            try:
                log(f"Attempt {attempt+1}/{max_retries}...", log_callback)
                time.sleep(delay_seconds)
                
                # Wait for form
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.NAME, "cmt"))
                )
                
                # Use the 'mst' field
                mst_input = driver.find_element(By.NAME, "mst")
                mst_input.clear()
                mst_input.send_keys(cccd)
                
                # Solve Captcha
                captcha_text = solve_captcha(driver, log_callback)
                
                # Use JS to set captcha
                captcha_input = driver.find_element(By.NAME, "captcha")
                driver.execute_script("arguments[0].value = '';", captcha_input)
                driver.execute_script("arguments[0].value = arguments[1];", captcha_input, captcha_text)
                
                # Submit
                log("Submitting form...", log_callback)
                try:
                    buttons = driver.find_elements(By.TAG_NAME, "button")
                    search_btn = None
                    for btn in buttons:
                        if "tra cứu" in btn.text.lower():
                            search_btn = btn
                            break
                    
                    if search_btn:
                        driver.execute_script("arguments[0].click();", search_btn)
                    else:
                        inputs = driver.find_elements(By.TAG_NAME, "input")
                        for inp in inputs:
                            if inp.get_attribute("type") == "button" and "tra cứu" in inp.get_attribute("value").lower():
                                inp.click()
                                break
                        else:
                            captcha_input.submit()
                except Exception as e:
                    log(f"Submit error: {e}", log_callback)
                    captcha_input.submit()
                
                # Wait for result
                time.sleep(2)
                
                page_source = driver.page_source
                
                # Check for rate limiting
                if "Too Many Requests" in page_source:
                    log("Rate limited. Waiting 60 seconds...", log_callback)
                    time.sleep(60)
                    driver.refresh()
                    continue
                
                # Check for captcha error
                if "Vui lòng nhập đúng mã xác nhận" in page_source or "Sai mã xác nhận" in page_source:
                    log("Incorrect captcha, retrying...", log_callback)
                    continue
                
                # Check for results
                try:
                    table = driver.find_element(By.CLASS_NAME, "ta_border")
                    rows = table.find_elements(By.TAG_NAME, "tr")
                    if len(rows) > 1:
                        data_row = rows[1]
                        cols = data_row.find_elements(By.TAG_NAME, "td")
                        if len(cols) >= 5:
                            result["tax_id"] = cols[1].text.strip()
                            result["name"] = cols[2].text.strip()
                            result["place"] = cols[3].text.strip()
                            
                            if len(rows) > 2:
                                result["status"] = "More than 1 record"
                            else:
                                result["status"] = cols[4].text.strip()
                            
                            log(f"Found result: {result}", log_callback)
                            return result
                except:
                    pass
                
                if "Không tìm thấy" in page_source:
                    result["status"] = "Not Found"
                    log("Result: Not Found", log_callback)
                    return result
                
                log("Unknown state, refreshing...", log_callback)
                
            except Exception as e:
                log(f"Error during attempt {attempt+1}: {e}", log_callback)
                time.sleep(2)
                
    except Exception as e:
        result["status"] = f"Error: {str(e)}"
        log(f"Critical error: {e}", log_callback)
    finally:
        if driver:
            log("Closing driver...", log_callback)
            driver.quit()
    
    return result


def main():
    """Test function for standalone execution."""
    cccd_list = ["001090000001", "0319287396", "079203027888"]
    results = []
    
    print(f"Checking {len(cccd_list)} CCCDs on Official Site...")
    
    for cccd in cccd_list:
        print(f"Checking {cccd}...")
        res = check_cccd_official(cccd, open_browser=True)
        print(f"Result: {res}")
        results.append(res)
    
    # Save to CSV
    with open("tax_check_official_results.csv", "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["cccd", "tax_id", "name", "place", "status"])
        writer.writeheader()
        writer.writerows(results)
    
    print("Done. Results saved to tax_check_official_results.csv")


if __name__ == "__main__":
    main()
