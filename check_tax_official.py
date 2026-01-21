"""
Tax Checker - Official API Integration Module

This module handles checking CCCD (Citizen ID) against the official tax portal
using Selenium WebDriver and RapidOCR for captcha solving.

Resources (bundled when running as PyInstaller app):
- EdgeDriver: Handled by webdriver-manager (auto-download)
- RapidOCR Models: ONNX models in models/ folder
"""
import os
import sys
import time
import csv
import warnings
import re

from PIL import Image, ImageOps
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException
from webdriver_manager.chrome import ChromeDriverManager
from rapidocr_onnxruntime import RapidOCR


# =============================================================================
# Resource Path Helper (PyInstaller compatible)
# =============================================================================
def get_resource_path(relative_path):
    """
    Get absolute path to resource, works for dev and PyInstaller bundle.
    
    In PyInstaller --onefile mode, files are extracted to sys._MEIPASS.
    In development mode, uses the current working directory.
    """
    if hasattr(sys, '_MEIPASS'):
        # Running in PyInstaller bundle
        base_path = sys._MEIPASS
    else:
        # Running in development mode
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def get_bundled_browser_path():
    """
    Get the path to bundled Chrome browser and ChromeDriver.
    Works for both dev mode and PyInstaller bundle.
    
    Returns:
        tuple: (chrome_exe_path, chromedriver_path) or (None, None) if not found
    """
    if hasattr(sys, '_MEIPASS'):
        # Running in PyInstaller bundle
        base_path = sys._MEIPASS
    else:
        # Running in development mode - use the script's directory
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    # Expected paths relative to base
    chrome_exe = os.path.join(base_path, "bin", "chrome-win64", "chrome.exe")
    chromedriver = os.path.join(base_path, "bin", "chromedriver.exe")
    
    if os.path.exists(chrome_exe) and os.path.exists(chromedriver):
        return chrome_exe, chromedriver
    
    return None, None


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
# OCR Functions (RapidOCR)
# =============================================================================
_ocr_reader = None

def get_ocr_reader():
    """Get or initialize the RapidOCR reader."""
    global _ocr_reader
    
    if _ocr_reader is None:
        det_model = get_resource_path("models/en_PP-OCRv3_det_infer.onnx")
        rec_model = get_resource_path("models/en_PP-OCRv3_rec_infer.onnx")
        
        _ocr_reader = RapidOCR(
            det_model_path=det_model,
            rec_model_path=rec_model,
            use_cls=False,  # Skip classification model for captcha
        )
    
    return _ocr_reader


# =============================================================================
# WebDriver Setup
# =============================================================================
def setup_driver(open_browser=False, log_callback=None):
    """
    Setup Chrome driver using bundled browser (offline-first).
    Falls back to webdriver-manager if bundled browser not found.
    """
    log(f"Setting up Chrome driver (Headless: {not open_browser})...", log_callback)
    
    options = webdriver.ChromeOptions()
    if not open_browser: 
        options.add_argument("--headless=new")  # New headless mode renders like real browser
    options.add_argument("--window-size=1920,1080")  # Set proper window size to prevent overlay issues
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")  # Disable GPU for headless stability
    options.add_argument("--disable-extensions")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    
    # Try to use bundled browser first (100% offline)
    chrome_exe, chromedriver = get_bundled_browser_path()
    
    if chrome_exe and chromedriver:
        log(f"Using bundled Chrome: {chrome_exe}", log_callback)
        log(f"Using bundled ChromeDriver: {chromedriver}", log_callback)
        options.binary_location = chrome_exe
        driver = webdriver.Chrome(service=Service(chromedriver), options=options)
    else:
        # Fallback to webdriver-manager (requires internet for first download)
        log("Bundled browser not found, falling back to webdriver-manager...", log_callback)
        log("This will use the system Chrome and auto-download ChromeDriver.", log_callback)
        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()), 
            options=options
        )
    
    log("Chrome driver setup complete.", log_callback)
    return driver


# =============================================================================
# Captcha Solving
# =============================================================================
def solve_captcha(driver, log_callback=None):
    try:
        log("Attempting to solve captcha...", log_callback)
        
        # --- 1. Robust Image Finding ---
        # (Your existing finding logic is okay, but brittle. 
        # Ideally, find by ID or specific CSS selector if possible.)
        images = driver.find_elements(By.TAG_NAME, "img")
        captcha_img = None
        for img in images:
            src = img.get_attribute("src") or ""
            # Check src AND alt text, or ID if possible
            if "captcha" in src.lower() or "jcaptcha" in src.lower():
                captcha_img = img
                break
        
        if not captcha_img:
            log("Captcha image not found.", log_callback)
            return None
            
        # --- 2. Improved Capture & Preprocessing ---
        captcha_path = "current_captcha.png"
        processed_path = "current_captcha_processed.png"
        
        # Take screenshot
        captcha_img.screenshot(captcha_path)
        
        image = Image.open(captcha_path)
        
        # A. Convert to Grayscale (Do NOT binarize/threshold)
        image = image.convert('L')
        
        # B. Upscale (3x is good, keep this)
        image = image.resize((image.width * 3, image.height * 3), Image.Resampling.LANCZOS)
        
        # C. ADD PADDING (Crucial for RapidOCR)
        # Add a 50px white border so text doesn't touch edges
        image = ImageOps.expand(image, border=50, fill='white')
        
        # D. Optional: Invert if text is white-on-black
        # (Check the top-left pixel. If it's black (0), the background is black.)
        if image.getpixel((0, 0)) < 128:
            image = ImageOps.invert(image)

        image.save(processed_path)
        
        # --- 3. Run RapidOCR ---
        reader = get_ocr_reader() # Ensure this is initialized somewhere
        
        # Run inference
        result, _ = reader(processed_path)
        
        if not result:
            log("RapidOCR found no text in image.", log_callback)
            return None

        # --- 4. Clean Results ---
        # Extract text from the result list: [[box], "text", confidence]
        raw_text = ''.join([line[1] for line in result])
        
        # Filter: Remove spaces and special chars (keep only alphanumeric)
        # CAPTCHAs usually don't have spaces or punctuation
        final_text = re.sub(r'[^A-Za-z0-9]', '', raw_text)
        
        log(f"Captcha solved: '{final_text}'", log_callback)
        return final_text
        
    except Exception as e:
        log(f"Captcha error: {e}", log_callback)
        return None


# =============================================================================
# Main Tax Check Function
# =============================================================================
def check_cccd_official(cccd, open_browser=False, log_callback=None, max_retries=10, delay_seconds=4, driver=None):
    """
    Check a CCCD number against the official tax portal.
    
    Args:
        cccd: The CCCD number to check
        open_browser: Whether to show the browser (only used if driver is None)
        log_callback: Callback function for logging
        max_retries: Maximum number of captcha retry attempts
        delay_seconds: Delay between attempts
        driver: Optional existing WebDriver instance to reuse. If provided,
                the driver will NOT be closed after the check completes.
    
    Returns:
        dict with cccd, tax_id, name, place, status, and needs_driver_recreation flag
    """
    owns_driver = driver is None  # Track if we created the driver (and should close it)
    result = {
        "cccd": cccd,
        "tax_id": None,
        "name": None,
        "place": None,
        "status": "Not Found",
        "needs_driver_recreation": False,  # Flag to signal driver should be recreated
    }
    
    try:
        log(f"Starting check for CCCD: {cccd}", log_callback)
        if driver is None:
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
                
                if "Mã số thuế không hợp lệ" in page_source:
                    result["status"] = "Invalid Tax ID"
                    log("Result: Invalid Tax ID", log_callback)
                    return result
                
                log("Unknown state, refreshing...", log_callback)
                
            except Exception as e:
                log(f"Error during attempt {attempt+1}: {e}", log_callback)
                time.sleep(2)
                
    except Exception as e:
        error_str = str(e)
        result["status"] = f"Error: {error_str}"
        log(f"Critical error: {e}", log_callback)
        
        # Check if this is a driver-related error that requires recreation
        if any(err in error_str.lower() for err in [
            "invalid session id",
            "session not created",
            "session deleted",
            "no such session",
            "httpconnectionpool",
            "read timed out",
            "connection refused",
            "target window already closed",
            "chrome not reachable",
            "browser has disconnected",
        ]):
            result["needs_driver_recreation"] = True
            log("Driver error detected - signaling need for driver recreation", log_callback)
    finally:
        # Only close driver if we created it (not reusing an existing one)
        if driver and owns_driver:
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
