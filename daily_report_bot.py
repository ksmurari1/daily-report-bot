import pyautogui
import pyperclip
import openpyxl
from datetime import datetime
import time
import os
from copy import copy


# ============================================================
# INITIALIZATION
# ============================================================

current_datetime = datetime.now()
current_date = current_datetime.strftime("%Y-%m-%d")

# Output file names
excel_filename = f"daily_report_{current_date}.xlsx"
screenshot_filename = f"daily_report_{current_date}.png"

print("Starting Daily Report Automation...")
print(f"Report date: {current_date}")
print(f"Excel file: {excel_filename}")
print(f"Screenshot: {screenshot_filename}")
print("Daily Report Automation initialized.")


# ============================================================
# STEP 1: OPEN CHROME AND FETCH WEATHER
# ============================================================

print("Opening Chrome...")

pyautogui.hotkey("win", "r")
time.sleep(2)

pyautogui.write("chrome")
pyautogui.press("enter")

time.sleep(5)

# Open Chennai weather search
pyautogui.hotkey("ctrl", "l")
pyautogui.write("https://www.google.com/search?q=Chennai+weather")
pyautogui.press("enter")

time.sleep(5)

print("Weather page opened.")


# ============================================================
# STEP 2: COPY AND EXTRACT WEATHER INFORMATION
# ============================================================

print("Copying weather information...")

pyautogui.hotkey("ctrl", "a")
pyautogui.hotkey("ctrl", "c")

time.sleep(2)

# Return to VS Code
pyautogui.hotkey("alt", "tab")
time.sleep(2)

weather_data = pyperclip.paste()

# Convert copied webpage text into individual lines
lines = [line.strip() for line in weather_data.splitlines() if line.strip()]

# Extract useful weather details
temperature = next(
    (line for line in lines if "°C" in line),
    "N/A"
)
temperature = temperature.replace("°F", "").strip()

precipitation = next(
    (line for line in lines if "Precipitation:" in line),
    "N/A"
)
precipitation = precipitation.replace("Precipitation:", "").strip()

humidity = next(
    (line for line in lines if "Humidity:" in line),
    "N/A"
)
humidity = humidity.replace("Humidity:", "").strip()

wind = next(
    (line for line in lines if "Wind:" in line),
    "N/A"
)
wind = wind.replace("Wind:", "").strip()

print("Weather information captured:")
print(f"Temperature: {temperature}")
print(f"Precipitation: {precipitation}")
print(f"Humidity: {humidity}")
print(f"Wind: {wind}")


# ============================================================
# STEP 3: CREATE AND FORMAT EXCEL REPORT
# ============================================================

print("Creating Excel report...")

workbook = openpyxl.Workbook()
worksheet = workbook.active
worksheet.title = "Daily Weather Report"

# Add headers
worksheet.append([
    "Date",
    "Time",
    "Temperature",
    "Precipitation",
    "Humidity",
    "Wind",
    "Comment"
])

# Add weather data
worksheet.append([
    current_date,
    current_datetime.strftime("%H:%M:%S"),
    temperature,
    precipitation,
    humidity,
    wind,
    "Weather data captured automatically using PyAutoGUI"
])

# Format header row
for cell in worksheet[1]:

    # Copy existing font and make it bold
    new_font = copy(cell.font)
    new_font.bold = True
    cell.font = new_font

    # Copy existing alignment and center it
    new_alignment = copy(cell.alignment)
    new_alignment.horizontal = "center"
    cell.alignment = new_alignment


# Set column widths
column_widths = {
    "A": 15,
    "B": 12,
    "C": 15,
    "D": 18,
    "E": 15,
    "F": 15,
    "G": 55
}

for column, width in column_widths.items():
    worksheet.column_dimensions[column].width = width


# Freeze header row
worksheet.freeze_panes = "A2"

# Add filter to header
worksheet.auto_filter.ref = "A1:G2"


# Save Excel file
workbook.save(excel_filename)

print(f"Excel report created: {excel_filename}")


# ============================================================
# STEP 4: OPEN EXCEL AND CAPTURE SCREENSHOT
# ============================================================

print("Opening Excel report...")

os.startfile(excel_filename)

time.sleep(5)

# Maximize Excel window
pyautogui.hotkey("win", "up")

time.sleep(2)

# Capture screenshot
pyautogui.screenshot(screenshot_filename)

print(f"Screenshot captured: {screenshot_filename}")
print("Daily Report Automation completed successfully.")