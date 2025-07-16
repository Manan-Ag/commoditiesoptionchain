import os
import glob
import subprocess
import time
import urllib.request
import tkinter as tk
from tkinter import messagebox
from csvtoxlsm import update_xlsm_with_bhavcopy
from datevalidation import add_date_dropdown_to_option_chain

# === Check for internet connection ===
def is_connected():
    try:
        urllib.request.urlopen('https://www.google.com', timeout=5)
        return True
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
        print(f"🗑️ Deleting old file: {file}")
        os.remove(file)

# === STEP 2: RUN SCRAPER TO DOWNLOAD NEW CSV ===
def run_scraper():
    print("🔄 Running scraper.py...")
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
        print(f"📝 Updating: {file}")
        update_xlsm_with_bhavcopy(file)

# === STEP 4: ADD DROPDOWNS ===
def add_dropdowns():
    files = [
        "GOLD option chain.xlsm",
        "CRUDEOIL option chain.xlsm",
        "SILVER option chain.xlsm"
    ]
    for file in files:
        print(f"🎯 Adding dropdown to: {file}")
        add_date_dropdown_to_option_chain(file)

# === MAIN EXECUTION ===
if __name__ == "__main__":
    if not is_connected():
        print("❌ No internet connection. Retrying in 60 seconds...")
        time.sleep(60)
        if not is_connected():
            print("❌ Still no internet. Exiting.")
            exit(1)

    if not confirm_popup():
        print("❌ Cancelled by user.")
        exit(0)

    delete_old_bhavcopies(os.getcwd())
    run_scraper()
    update_all_workbooks()
    add_dropdowns()
    print("✅ All tasks completed successfully.")
