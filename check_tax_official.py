import time
import csv
import warnings
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from PIL import Image
import pytesseract

# Suppress SSL warnings
warnings.filterwarnings("ignore")

def log(message, callback=None):
    if callback:
        callback(message)
    else:
        print(message)

def setup_driver(open_browser=False, log_callback=None):
    log(f"Setting up Chrome driver (Headless: {not open_browser})...", log_callback)
    options = webdriver.ChromeOptions()
    if not open_browser: 
        options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    log("Chrome driver setup complete.", log_callback)
    return driver

def solve_captcha(driver, log_callback=None):
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
        captcha_img.screenshot("current_captcha.png")
        
        # Process and OCR
        image = Image.open("current_captcha.png")
        image = image.convert('L')
        
        # Resize to make it bigger (3x)
        image = image.resize((image.width * 3, image.height * 3), Image.Resampling.LANCZOS)
        
        # Thresholding
        threshold = 140
        image = image.point(lambda x: 0 if x < threshold else 255, '1')
        
        # Tesseract config
        custom_config = r'--oem 3 --psm 8 -c tessedit_char_whitelist=abcdefghijklmnopqrstuvwxyz0123456789'
        text = pytesseract.image_to_string(image, config=custom_config)
        result_text = text.strip()
        log(f"Captcha solved: '{result_text}'", log_callback)
        return result_text
    except Exception as e:
        log(f"Captcha error: {e}", log_callback)
        return None

def check_cccd_official(cccd, open_browser=False, log_callback=None):
    # Start a fresh driver for each CCCD to ensure stability
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
        max_retries = 20
        url = "https://tracuunnt.gdt.gov.vn/tcnnt/mstcn.jsp"
        log(f"Navigating to {url}", log_callback)
        driver.get(url)
        
        for attempt in range(max_retries):
            try:
                log(f"Attempt {attempt+1}/{max_retries}...", log_callback)
                # time.sleep(10)
                
                # Wait for form
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.NAME, "cmt"))
                )
                
                # Use the 'mst' field as requested by the user
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
                # Try multiple ways to submit
                try:
                    # Try finding the button by onclick attribute or generic button tag
                    buttons = driver.find_elements(By.TAG_NAME, "button")
                    search_btn = None
                    for btn in buttons:
                        if "tra cứu" in btn.text.lower():
                            search_btn = btn
                            break
                    
                    if search_btn:
                        driver.execute_script("arguments[0].click();", search_btn)
                    else:
                        # Try input[type=button]
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
                    log("Rate limited (Too Many Requests). Waiting 60 seconds...", log_callback)
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
                            result["status"] = cols[4].text.strip()
                            log(f"Found result: {result}", log_callback)
                            return result
                except:
                    pass
                
                if "Không tìm thấy" in page_source:
                    result["status"] = "Not Found"
                    log("Result: Not Found", log_callback)
                    return result
                
                log("Unknown state (no result/error found), refreshing...", log_callback)
                # Debug: print part of the page source to see what's happening
                # print(f"Page source snippet: {page_source[:500]}...")
                    
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
