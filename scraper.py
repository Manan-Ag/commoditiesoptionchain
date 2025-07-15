from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import sys

# 🚩 Change this to your download folder
download_dir = "/Users/manan/Documents/git projects/commoditiesoptionchain-1"

print("🔧 Configuring Chrome options...")
options = webdriver.ChromeOptions()
options.add_experimental_option("prefs", {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "safebrowsing.enabled": True
})

try:
    print("🚀 Launching Chrome driver...")
    driver = webdriver.Chrome(service=Service(), options=options)
except Exception as e:
    print(f"❌ Failed to launch Chrome: {e}")
    sys.exit(1)

try:
    print("🌐 Opening MCX Bhavcopy page...")
    driver.get("https://www.mcxindia.com/market-data/bhavcopy")
    time.sleep(3)

    print("🔍 Looking for CSV download link with ID `lnkExpToCSV`...")
    link = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "lnkExpToCSV"))
    )
    print("✅ Found download link!")

    try:
        print("📦 Scrolling into view...")
        driver.execute_script("arguments[0].scrollIntoView(true);", link)
        time.sleep(1)
    except Exception as e:
        print(f"⚠️ Could not scroll: {e}")

    try:
        print("🖱️ Attempting normal click...")
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "lnkExpToCSV")))
        link.click()
        print("✅ Normal click successful.")
    except Exception as e:
        print(f"⚠️ Normal click failed: {e}")
        print("🧪 Trying JavaScript click...")
        try:
            driver.execute_script("document.getElementById('lnkExpToCSV').click();")
            print("✅ JavaScript click successful.")
        except Exception as js_e:
            print(f"❌ JavaScript click also failed: {js_e}")
            raise

    print("⏳ Waiting for CSV file to appear in download folder...")
    found = False
    wait_seconds = 20
    for second in range(wait_seconds):
        files = [f for f in os.listdir(download_dir) if f.endswith(".csv")]
        if files:
            print(f"✅ CSV file downloaded: {files[0]}")
            found = True
            break
        time.sleep(1)

    if not found:
        print("❌ CSV file not found after timeout. Check browser settings or site behavior.")

except Exception as e:
    print(f"🔥 Script failed: {e}")

finally:
    print("🧹 Closing browser...")
    driver.quit()
