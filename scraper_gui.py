import pyautogui
import time
import webbrowser

print("🌐 Opening MCX Bhavcopy page...")
webbrowser.open("https://www.mcxindia.com/market-data/bhavcopy")
time.sleep(5)  # Let page load

print("🔍 Searching for CSV icon...")
button_location = pyautogui.locateCenterOnScreen("image.png", confidence=0.9)

if button_location is not None:
    print("✅ Found CSV icon at:", button_location)
    pyautogui.moveTo(button_location)
    pyautogui.click()
    pyautogui.moveTo(button_location)
    pyautogui.click()
    pyautogui.alert(text='Clicked here', title='Debug', button='OK')

else:
    print("❌ Could not find CSV icon.")
