import os
import glob
import subprocess
import time
import urllib.request
import tkinter as tk
from tkinter import messagebox
from csvtoxlsm import update_xlsm_with_bhavcopy
from datevalidation import add_date_dropdown_to_option_chain
import logging

# === Check for internet connection ===
def is_connected():
    try:
        # Try multiple methods to check connectivity
        # Method 1: Try HTTP first (no SSL issues)
        urllib.request.urlopen('http://www.google.com', timeout=5)
        return True
    except:
        try:
            # Method 2: Try HTTPS
            urllib.request.urlopen('https://www.google.com', timeout=5)
            return True
        except:
            try:
                # Method 3: Use subprocess ping as fallback
                result = subprocess.run(['ping', '-c', '1', 'google.com'], 
                                      capture_output=True, timeout=10)
                return result.returncode == 0
            except:
                return False

# === GUI Prompt ===
def confirm_popup():
    root = tk.Tk()
    root.withdraw()  # Hide main window
    return messagebox.askyesno("Run Script", "Internet is available.\nDo you want to continue?")

# === STEP 1: DELETE OLD BhavCopy CSVs ===
def delete_old_bhavcopies(directory, prefix="BhavCopyDateWise_"):
    for file in glob.glob(os.path.join(directory, f"{prefix}*.csv")):
        log(f"🗑️ Deleting old file: {file}")
        os.remove(file)

# === STEP 2: RUN SCRAPER TO DOWNLOAD NEW CSV ===
def run_scraper():
    log("🔄 Running scraper.py...")
    result = subprocess.run(["python3", "scraper.py"], capture_output=True, text=True)
    print(result.stdout)
    if result.returncode != 0:
        print(result.stderr)
        raise RuntimeError("❌ scraper.py failed")

# === STEP 3: UPDATE XLSM FILES ===
def update_all_workbooks():
    files = [
        "GOLD option chain.xlsm",
        "CRUDEOIL option chain.xlsm",
        "SILVER option chain.xlsm"
    ]
    for file in files:
        log(f"📝 Updating: {file}")
        update_xlsm_with_bhavcopy(file)

# === STEP 4: ADD DROPDOWNS ===
def add_dropdowns():
    files = [
        "GOLD option chain.xlsm",
        "CRUDEOIL option chain.xlsm",
        "SILVER option chain.xlsm"
    ]
    for file in files:
        log(f"🎯 Adding dropdown to: {file}")
        add_date_dropdown_to_option_chain(file)

# === MAIN EXECUTION ===
if __name__ == "__main__":
    logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(message)s',
    handlers=[
        logging.FileHandler("/Users/manan/Documents/git projects/commoditiesoptionchain/log.txt"),
        logging.StreamHandler()
    ]
    )
    log = logging.info
    if not is_connected():
        log("❌ No internet connection. Retrying in 60 seconds...")
        time.sleep(60)
        if not is_connected():
            log("❌ Still no internet. Exiting.")
            exit(1)

    if not confirm_popup():
        log("❌ Cancelled by user.")
        exit(0)

    delete_old_bhavcopies(os.getcwd())
    run_scraper()
    update_all_workbooks()
    add_dropdowns()
    log("✅ All tasks completed successfully.")
