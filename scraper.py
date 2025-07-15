import os
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# === CONFIG ===
DOWNLOAD_DIR = os.path.abspath(".")  # project folder
CHROMEDRIVER_PATH = "/opt/homebrew/bin/chromedriver"  # change if needed

# === Chrome setup ===
options = Options()
options.add_experimental_option("prefs", {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})
# Optional: Headless mode
# options.add_argument("--headless=new")

# === Launch Chrome ===
service = Service(CHROMEDRIVER_PATH)
driver = webdriver.Chrome(service=service, options=options)
driver.get("https://www.mcxindia.com/market-data/bhavcopy")


try:
    print("⏳ Waiting for CSV icon via XPath...")

    # Use full XPath here
    download_link = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[4]/div[3]/div/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[5]/a[2]'))
    )

    # Scroll into view and click via JS
    driver.execute_script("arguments[0].scrollIntoView(true);", download_link)
    time.sleep(1)
    driver.execute_script("arguments[0].click();", download_link)

    print("✅ CSV download link clicked.")

except Exception as e:
    print(f"❌ Could not click download link: {str(e)}")


# === Wait for the file to appear ===
today = datetime.now().strftime("%d%m%Y")
expected_filename = f"BhavCopyDateWise_{today}.csv"
download_path = os.path.join(DOWNLOAD_DIR, expected_filename)

print("⏳ Waiting for CSV file...")
for _ in range(30):  # wait up to 30 seconds
    if os.path.exists(download_path):
        print(f"✅ File downloaded: {expected_filename}")
        break
    time.sleep(1)
else:
    print("❌ File did not download in time.")

driver.quit()
