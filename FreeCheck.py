# cancelled prject
import easyocr
import time
import pyautogui
import numpy as np
from PIL import ImageGrab
import openpyxl
import pyperclip
import tkinter as tk
from tkinter import filedialog

def capture_screenshot(region=None):
    """
    Capture screenshot of the screen or a specific region.
    """
    screenshot = ImageGrab.grab(bbox=region) if region else ImageGrab.grab()
    return screenshot

def extract_text_from_image(image):
    """
    Extract text from an image using EasyOCR.
    """
    # Convert PIL Image to NumPy array
    image_np = np.array(image)
    reader = easyocr.Reader(['en'])  # Initialize EasyOCR with English language
    results = reader.readtext(image_np)

    extracted_text = ""
    for bbox, text, confidence in results:
        extracted_text += f"{text}\n"
    
    all_extract = extracted_text.split('\n')
    if len(all_extract) > 23:
        serial_num = all_extract[5]
        #info = all_extract[23].split(" ")
        #free_location = info[2]

        #free_status = free_location[17]  # Assuming you're interested in a certain part of the string
        return f"{serial_num} " #{free_status}
    else:
        return "Unable to extract valid data"

def get_serial_numbers_from_excel(file_path):
    """
    Extract all serial numbers from the first column of the given Excel file.
    """
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    serial_numbers = []

    for row in sheet.iter_rows(min_row=1, max_col=1, values_only=True):  # Skipping header
        serial_numbers.append(row[0])

    return serial_numbers

def handle_serial(serial_number):
    """
    Handle each serial by pasting it, checking the account status using OCR, and deleting the entry.
    """
    # Use pyperclip to paste serial number into clipboard for pyautogui to paste
    pyperclip.copy(serial_number)
    
    # Paste the serial number (Ctrl+V)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.2)
    pyautogui.hotkey("tab")

    # Capture the screenshot after pasting
    screenshot = capture_screenshot()
    
    # Extract text from the screenshot
    extracted_text = extract_text_from_image(screenshot)
    print(f"Account Status for Serial {serial_number}: {extracted_text}")

    # Click the position to delete (742x, 1253y) and press delete for 2 seconds
    pyautogui.click()
    time.sleep(0.5)  # Wait for 721107004143   721106902584 lick action
    pyautogui.press('backspace')
    time.sleep(2)

    return extracted_text

def select_excel_file():
    """
    Prompt the user to select an Excel file using a file dialog.
    """
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    file_path = filedialog.askopenfilename(
        title="Select an Excel File",
        filetypes=(("Excel Files", "*.xlsx;*.xls"), ("All Files", "*.*"))
    )
    return file_path

def main():
    print("Starting the process...")

    # Prompt the user to select the Excel file
    excel_file = select_excel_file()

    if not excel_file:  # If the user cancels the file selection
        print("No file selected. Exiting.")
        return

    # Get the serial numbers from the selected Excel file
    serial_numbers = get_serial_numbers_from_excel(excel_file)

    all_results = []
    time.sleep(2)
    for serial_number in serial_numbers:
        print(f"Processing Serial: {serial_number}")
        result = handle_serial(serial_number)
        all_results.append(result)

    # Output all serials and their account statuses
    print("All Processed Serial Numbers and Their Account Statuses:")
    for result in all_results:
        print(result)

if __name__ == "__main__":
    main()
