import pyautogui
import time
import webbrowser
import os

# Open the MCX Bhavcopy page
webbrowser.open("https://www.mcxindia.com/market-data/bhavcopy")
print("🌐 Opening MCX Bhavcopy page...")
time.sleep(2)  # Wait for page to load

# Locate and click the CSV icon
print("🔍 Searching for CSV icon...")
button_location = pyautogui.locateCenterOnScreen("csv.png", confidence=0.8)

if button_location:
    pyautogui.moveTo(button_location)
    pyautogui.click()
    print("✅ Clicked CSV icon.")
else:
    print("❌ Could not find CSV icon. Try re-capturing the image.")
