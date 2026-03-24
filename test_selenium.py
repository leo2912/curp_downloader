from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
import time

curp = "SAHL140512MMNLRSA0"

options = Options()
options.add_argument("--headless")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")

print("Starting driver...")
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

try:
    print("Loading page...")
    driver.get("https://www.gob.mx/curp/")
    wait = WebDriverWait(driver, 20)
    
    print("Waiting for curpinput...")
    curp_input = wait.until(EC.presence_of_element_located((By.ID, "curpinput")))
    curp_input.clear()
    curp_input.send_keys(curp)
    
    print("Clicking search...")
    search_button = driver.find_element(By.ID, "searchButton")
    # Using JS click to avoid overlap issues
    driver.execute_script("arguments[0].click();", search_button)
    
    print("Waiting for download button...")
    try:
        download_button = wait.until(EC.element_to_be_clickable((By.ID, "download")))
        print("Found download button!")
    except Exception as e:
        print("Timeout or error finding download button:", type(e))
        with open("c:/Users/chrssh/GitHub/curp_downloader/page_source.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        driver.save_screenshot("c:/Users/chrssh/GitHub/curp_downloader/error_screenshot.png")
        print("Saved page source and screenshot")
finally:
    driver.quit()
