#==============================================================================#
# Fundamentals #
#==============================================================================#

# Imports
from pathlib import Path
from tkinter import Tk, Canvas, Entry, Text, Button, PhotoImage
from pathlib import Path
from tkinter import Tk, Canvas, Entry, Text, Button, PhotoImage
import tkinter as tk
from tkinter import filedialog, messagebox, PhotoImage
import openpyxl
import clipboard
import pyautogui
import time
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
import xml.etree.ElementTree as ET
import pyodbc
import subprocess
import datetime
import getpass
import pandas as pd
import pyodbc
import subprocess
import datetime
import getpass

# Getting global variables
config_path = "c:\\Josh\\config.txt"
with open(config_path, 'r') as file:
    config = file.read().split("\n")

# Global Variables:
user = config[0]
import_speed = float(config[1])
asset_location = config[2]
database_path = config[3]
bartender_notepad = config[4]
serials_reversed = bool(config[5])
serials_list = []
remaining_serials = 0
file_path = None
error_serial = []
good_serial = []

# WANT TO REMOVE THIS OR AT LEAST UPDATE IT TO A FIREBASE THING
connection_string = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + database_path

# Formats
def RelativeToAssets(path: str) -> Path:
    return ASSETS_PATH / Path(path)

#  Lets you select an excel and then loads the serials for use.
def OpenExcel():
    global file_path
    file_path = filedialog.askopenfilename(title="Open an Excel for use", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    if file_path:
        LoadSerials(file_path)
    
# Gets the serials from the Excel file.
def LoadSerials(file_path):
    global serials_list, remaining_serials
    try:
        # loads at the active sheet 
        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active
        serials_list.clear() #clears the serials previously loaded if any.
        
        # Gets every serial in the first column and adds it into the serial list.
        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row):
            if row[0].value is not None:
                serials_list.append(str(row[0].value))

        # Gets the amount of serials loaded and updates the display with the function.  
        remaining_serials = len(serials_list)
        if serials_reversed == True:
            serials_list.reverse() # Reverse the serials to account for FlexiPro!
        UpdateSerialsDisplay()

    # If serials are not able to be loaded by the file, this handles the error.
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load serials: {e}")

# Configures the text to let the user know how many serials are loaded.
def UpdateSerialsDisplay():
    serials_text.delete(1.0, tk.END)
    serials_text.insert(tk.END, "\n".join(serials_list))
    count_label.config(text=f"{remaining_serials} Serials Loaded")

#==============================================================================#
# Import Stuff #
#==============================================================================#

# Helper function for PasteSerials now as FlexiPro requires you to wait for a loading bar to complete.
def CheckPixelFlexiPro():
        flexLocal = pyautogui.screenshot() # Takes a screenshot of the lower screen
        colorPixel = (flexLocal.getpixel((961,1016))) # Gets the color of the loading bar in that screenshot
        whitepixel = (250, 250, 250) # This is the color of the loading bar when you are able to import a new serial. As such it is the target.
        if colorPixel != whitepixel: 
            # If the pixel found is not the color of the target set, then we return false.
            return False
        else: 
            # If we got the target pixel from our check it will return true.
             return True
        
# Helper function for PasteSerials now as WMS sometimes creates errors.
def CheckPixelWMS():
        flexLocal = pyautogui.screenshot() # Takes a screenshot of the lower screen
        colorPixel = (flexLocal.getpixel((30,149))) # Gets the color of the loading bar in that screenshot
        errorColor = (255, 255, 255) # This is the color of the loading bar when you are able to import a new serial. As such it is the target.
        goodColor = (0, 0, 0)
        if colorPixel == errorColor:
            return True # If there is an error we return True for "isError"
        elif colorPixel == goodColor:
             return False # If there is not an error we return False for "isError"
        
# Pastes the serials automatically for WMS.
def PasteSerialsWMS():
    global remaining_serials
    # Handles if there are no serials in our serial list.
    if not serials_list:
        messagebox.showinfo("Info", "No serials loaded to paste.")
        return
    
    # Allows time for the user to focus on the target program
    time.sleep(5)  
    start = time.time() # This is the start of the timer for the import time.

    # Lists for filtering the serials.
    error_serial = [] # Stores the serials that were not able to be imported.
    good_serial = [] # Stores the serials that were able to be imported.
    
    for serial in serials_list[:]:
        if serial is not None:
            pyautogui.typewrite(serial) # Writes the serial.
            time.sleep(0.2) # WAS 0.5 BUT TRYING TO SPEED IT UP
            pyautogui.hotkey("tab") # Waits before pressing tab as to give the system a short breather.
            serials_list.remove(serial)
            remaining_serials -= 1
            UpdateSerialsDisplay()
            time.sleep(0.4) # WAS 1 BUT TRYING TO SPEED IT UP
            
            isError = CheckPixelWMS() # returns t/f

            if isError == False: # If there is no error, you can continue pasting. After sorting the serial
                print(str(serial) + " Has no error!")
                good_serial.append(serial)
            elif isError == True: # If there is an error you need to wait for the box to update, press ctrl + x and save the serial
                error_serial.append(serial)
                time.sleep(0.4) # WAS 0.75 BUT TRYING TO SPEED UP
                pyautogui.hotkey("ctrl", "x") # Gets out of the error screen
                print(str(serial) + " Had an error and will be saved")
            time.sleep(0.3) # WAS 0.5 BUT WANT TO SPEED IT UP.

    end = time.time() # Serials are all imported and as such we grab the ended time.
    
    # Prints the time the import took into the console
    print('Import completed in ' + str(round((end - start) / 60, 1)) + "minutes.") # Prints to the console how long the import took in mins
    
    # Prints the serials that did not have errors
    print("The good serials are:")
    for j in good_serial:
        print(j)
    
    # Prints the serials that did have errors
    print("The problematic serials are:")
    for i in (error_serial):
        print(i)
    
# Pastes the serials automatically for FlexiPro.
def PasteSerialsFlexi():
    global remaining_serials

    # Handles if there are no serials in our serial list.
    if not serials_list:
        messagebox.showinfo("Info", "No serials loaded to paste.")
        return
    
    # Allows time for the user to focus on the target program
    time.sleep(5)  
    start = time.time() # This is the start of the timer for the import time.

    for serial in serials_list[:]:
        if serial is not None:
            isPixelGood = CheckPixelFlexiPro()
            while isPixelGood == False: # If the pixel on the loading bar is not our target we run in this loop until the loading bar is percieved as the target color.
                time.sleep(0.75) # Add some time between screenshots as not to flood the program
                isPixelGood = CheckPixelFlexiPro() # If the pixel changed to target we will exit the loop and continue. 
            if CheckPixelFlexiPro() == True:    
                pyautogui.typewrite(serial) # Writes the serial.
                time.sleep(0.5)
                pyautogui.hotkey("tab") # Waits before pressing tab as to give the system a short breather.
                serials_list.remove(serial)
                remaining_serials -= 1
                UpdateSerialsDisplay()
                # Time between pastes, ideally rounded down as much as possible but FlexiPro is hard to estimate, you can customize this in the config.
                time.sleep(import_speed)

            # Shouldn't be possible to hit this else, but if it does happen, I aired on the side of not losing the serial and instead warning the user that some edge case was found.
            else: # Same as if true except has warning.
                print("Check the import for" + str(serial))
                time.sleep(2)
                pyautogui.typewrite(serial)
                time.sleep(0.5)
                pyautogui.hotkey("tab")
                serials_list.remove(serial)
                remaining_serials -= 1
                UpdateSerialsDisplay()
                time.sleep(import_speed)

    end = time.time() # Serials are all imported and as such we grab the ended time.
    print('Import completed in ' + str(round((end - start) / 60, 1)) + " minutes.") # Prints to the console how long the import took in mins

# Pastes the serials as fast as possible
def PasteSerialsNormal():
    global remaining_serials
    # Handles if there are no serials in our serial list.
    if not serials_list:
        messagebox.showinfo("Info", "No serials loaded to paste.")
        return
    # Allows time for the user to focus on the target program
    time.sleep(5)  
    start = time.time() # This is the start of the timer for the import time.

    for serial in serials_list[:]:
        if serial is not None:    
                pyautogui.typewrite(serial) # Writes the serial.
                pyautogui.hotkey("tab") # Waits before pressing tab as to give the system a short breather.
                serials_list.remove(serial)
                remaining_serials -= 1
                UpdateSerialsDisplay()
                # Time between pastes, ideally rounded down as much as possible but FlexiPro is hard to estimate, you can customize this in the config.
                time.sleep(0.3)

    end = time.time() # Serials are all imported and as such we grab the ended time.
    print('Import completed in ' + str(round((end - start) / 60, 1)) + " minutes.") # Prints to the console how long

#==============================================================================#
# Printing #
#==============================================================================#

# Formats TV devices by reversing every 10 serials and adding the found device name to the top.
def MakeTVSheet(device):
    total_strings = len(serials_list)
    formatted_list = []

    for i in range(0, total_strings, 10):
        formatted_list.append(device)
        chunk = serials_list[i:i + 10]
        chunk.reverse()
        formatted_list.extend(chunk)

    return '\n'.join(formatted_list)

# Helper function for Modem devices, it reverses every 5 serials.
def ReverseForModems():
    total_strings = len(serials_list)
    reversed_modems = []
    for i in range(0, total_strings, 5):
        chunk = serials_list[i:i + 5][::-1]
        reversed_modems.extend(chunk)
    return reversed_modems

# Formats Modem devices by reversing every 8 serials and adding the found device name to the top.
def MakeModemSheet(device):
    working_modems = ReverseForModems()
    total_strings = len(serials_list)
    formatted_list = []

    for i in range(0, total_strings, 8):
        formatted_list.append(device)
        chunk = working_modems[i:i + 8]
        formatted_list.extend(chunk)

    return '\n'.join(formatted_list)

# Takes the formatted string for the sheets, and executes a BAT file depending on the device that is found.
def CreatePurolatorSheet():
    global serials_list
    device = ""
    # Checks the first element in the serial list, and based on its prefix determines the device model name.
    if serials_list[0][0] == "T" and serials_list[0][1] == "M":
        device = "IPTVTCXI6HD"
    elif serials_list[0][0] == "M":
        device = "IPTVARXI6HD"
    elif serials_list[0][0] == "4" and serials_list[0][1] == "0" and serials_list[0][2] == "9":
        device = "CGM4981COM"
    elif serials_list[0][0] == "X" and serials_list[0][1] == "I" and serials_list[0][2] == "1":
        device = "SCXI11BEI"
    elif serials_list[0][0] == "3" and serials_list[0][1] == "3" and serials_list[0][2] == "6":
        device = "CGM4331COM"
    else:
        device = "TG4482A"

    # If a device is a TV device it will create the text file and execute a BAT file using it.
    if device in ["IPTVARXI6HD", "IPTVTCXI6HD", "SCXI11BEI"]:
        puro_sheet = MakeTVSheet(device)
        # Clears the stipulated notepad and then puts the purolator sheet made into it.
        with open(bartender_notepad, 'w') as file:
            file.write(f"{puro_sheet}\n")

        # CMD script changes the printer to the purolator sheet printer and then selects the modem sheet that prints 11 serials on each.
        cmd_script = """
        @echo off
        set "target_printer=55EXP_Purolator"
        powershell -Command "Get-WmiObject -Query \\"SELECT * FROM Win32_Printer WHERE ShareName='%target_printer%'\\" | Invoke-WmiMethod -Name SetDefaultPrinter"
        "C:\\Seagull\\BarTender 7.10\\Standard\\bartend.exe" /f=C:\\BTAutomation\\XI6.btw /p /x
        """
        # Updates the TEMP BAT and then executes it.
        with open("temp_cmd.bat", "w") as bat_file:
            bat_file.write(cmd_script)
        subprocess.run(["cmd.exe", "/c", "temp_cmd.bat"])
    
    # If a device is a Modem device it will create the text file and execute a BAT file using it.
    else:
        # Clears the stipulated notepad and then puts the purolator sheet made into it.
        puro_sheet = MakeModemSheet(device)
        with open(bartender_notepad, 'w') as file:
            file.write(f"{puro_sheet}\n")
        # CMD script changes the printer to the purolator sheet printer and then selects the modem sheet that prints 11 serials on each.
        cmd_script = """
        @echo off
        set "target_printer=55EXP_Purolator"
        powershell -Command "Get-WmiObject -Query \\"SELECT * FROM Win32_Printer WHERE ShareName='%target_printer%'\\" | Invoke-WmiMethod -Name SetDefaultPrinter"
        "C:\\Seagull\\BarTender 7.10\\Standard\\bartend.exe" /f=C:\\BTAutomation\\CODA.btw /p /x
        """
        # Updates the TEMP BAT and then executes it.
        with open("temp_cmd.bat", "w") as bat_file:
            bat_file.write(cmd_script)
        subprocess.run(["cmd.exe", "/c", "temp_cmd.bat"])

# Creates a Lot sheet based on the loaded serials.
def CreateLaser():
    # Create the barcode file for Bartender
    with open(bartender_notepad, 'w') as file:
        for serial in serials_list:
            file.write(f"{serial}\n")

    # Prepare and write the CMD script for printing barcodes
    cmd_script = """
    @echo off
    set "target_printer=55EXP_2"
    powershell -Command "Get-WmiObject -Query \\"SELECT * FROM Win32_Printer WHERE ShareName='%target_printer%'\\" | Invoke-WmiMethod -Name SetDefaultPrinter"
    "C:\\Seagull\\BarTender 7.10\\Standard\\bartend.exe" /f=C:\\BTAutomation\\NewPrintertest.btw /p /x
    """
    
    with open("temp_cmd.bat", "w") as bat_file:
        bat_file.write(cmd_script)

    # Run the script to print barcodes
    subprocess.run(["cmd.exe", "/c", "temp_cmd.bat"])

# Prints the loaded serials as barcodes.
def CreateBarcodes():
    # Create the barcode file for Bartender
    with open(bartender_notepad, 'w') as file:
        for serial in serials_list:
            file.write(f"{serial}\n")

    # Prepare and write the CMD script for printing barcodes
    cmd_script = """
    @echo off
    set "target_printer=55EXP_Barcode"
    powershell -Command "Get-WmiObject -Query \\"SELECT * FROM Win32_Printer WHERE ShareName='%target_printer%'\\" | Invoke-WmiMethod -Name SetDefaultPrinter"
    "C:\\Seagull\\BarTender 7.10\\Standard\\bartend.exe" /f=C:\\BTAutomation\\singlebar.btw /p /x
    """
    
    with open("temp_cmd.bat", "w") as bat_file:
        bat_file.write(cmd_script)

    # Run the script to print barcodes
    subprocess.run(["cmd.exe", "/c", "temp_cmd.bat"])

#==============================================================================#
# CTR Update Stuff #
#==============================================================================#

# Lets the GUI interact with Windows explorer and select the daily float snapshots for the use of the program.
def open_file_dialog():
    # Selects Multiple files to combine
    file_paths = filedialog.askopenfilenames(
        title="Select Excel Files to Combine",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )
    # Handles no selection
    if not file_paths:
        print("No files selected.")
        return

    # Initialize an empty list to store DataFrames
    excel_lst = []
    for file in file_paths:
        excel_lst.append(pd.read_excel(file))

    # Combine all DataFrames into one
    excel_merged = pd.concat(excel_lst, ignore_index=True)

    # Exports the combined DataFrame to a new Excel file
    output_file = 'DailyCombinedCTR.xlsx'
    excel_merged.to_excel(output_file, index=False)
    print(f"Combined Excel file saved as '{output_file}'.")
    
    # Asks for the file we created for use of the import.
    final_path = filedialog.askopenfilename(
        title="Select Excel File to use",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )
    if not final_path:
        print("No files selected.")
        return
    # Begin import with the file
    process_file(final_path)

# Using Filepath (the location of the chosen excel) if the file passes being possible, The first sheet in the workbook becomes our selection.
def process_file(file_path):
    try:
        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active
        process_sheet(sheet)
    except Exception as e:
        count_label.config(text=f"Error: {str(e)}")

# Main algorithmic processing for the update
def process_sheet(sheet):
    # Contractors, Robitaille, and Warehouse titles.
    robitaille_lst = ['8017', '8037', '8038', '8041', '8047', '8080', '8093']
    ctr_lst = ['8052', '8067', '8975', "8986", "8990", "8994", "8997"]
    warehouse_lst = ["NB1", "NF1"]
    
    data_lst = [] # Result List

    # Devices to count for each group
    robitaille_devices = [
        "XB8", "XB7", "XI6", "XIONE", "PODS", "ONTS",
        "SCHB1AEW", "SCHC2AEW", "SCHC3AEW", "SCXI11BEI-ENTOS",
        "MR36HW", "S5A134A", "CM8200A", "CODA5810"
    ]
    
    ctr_devices = [
        "XB8", "XB7", "XI6", "XIONE", "PODS", "ONTS",
        "SCHB1AEW", "SCHC2AEW", "SCHC3AEW", "SCXI11BEI-ENTOS",
        "CODA5810"
    ]
    warehouse_devices = [
        "XB8", "XB7", "XI6", "XIONE", "PODS", "ONTS",
        "SCHB1AEW", "SCHC2AEW", "SCHC3AEW", "SCXI11BEI-ENTOS",
        "MR36HW", "S5A134A", "CM8200A", "CODA5810"
    ]

    # Map item codes to device names
    device_mapping = {
        "CGM4981COM": "XB8",
        "CGM4331COM": "XB7", "TG4482A": "XB7",
        "IPTVARXI6HD": "XI6", "IPTVTCXI6HD": "XI6",
        "SCXI11BEI": "XIONE",
        "XE2SGROG1": "PODS",
        "XS010XB": "ONTS", "XS010XQ": "ONTS", "XS020XONT": "ONTS",
        "SCHB1AEW": "SCHB1AEW",
        "SCHC2AEW": "SCHC2AEW",
        "SCHC3AE0": "SCHC3AEW",
        "SCXI11BEI-ENTOS": "SCXI11BEI-ENTOS",
        "MR36HW": "MR36HW",
        "S5A134A": "S5A134A",
        "CM8200A": "CM8200A",
        "CODA5810": "CODA5810",
    }

    # Helper function to update totals based on the allowed device list
    def update_totals(totals, item_code, allowed_devices):
        if item_code in device_mapping:
            device = device_mapping[item_code]
            if device in allowed_devices:
                totals[device] += 1

    # Process Robitaille contractors (order as per your request)
    for contractor in robitaille_lst:
        contractor_totals = {device: 0 for device in robitaille_devices}
        for row in sheet.iter_rows(min_row=3, max_row=sheet.max_row):
            contractor_id = str(row[7].value)  # Column H
            item_code = row[5].value          # Column F
            inventory_type = str(row[9].value)  # Column J

            if contractor_id == contractor and inventory_type == f"CTR.Subready.{contractor}":
                update_totals(contractor_totals, item_code, robitaille_devices)

        data_lst.append(format_totals(contractor_totals, robitaille_devices))

    # Process CTR contractors 
    for contractor in ctr_lst:
        contractor_totals = {device: 0 for device in ctr_devices}
        for row in sheet.iter_rows(min_row=3, max_row=sheet.max_row):
            contractor_id = str(row[7].value)  # Column H
            item_code = row[5].value          # Column F
            inventory_type = str(row[9].value)  # Column J

            if contractor_id == contractor and inventory_type == f"CTR.Subready.{contractor}":
                update_totals(contractor_totals, item_code, ctr_devices)

        data_lst.append(format_totals(contractor_totals, ctr_devices))

    # Combine 8993 and 8982 contractors
    combined_contractor_totals = {device: 0 for device in ctr_devices}
    for row in sheet.iter_rows(min_row=3, max_row=sheet.max_row):
        contractor_id = str(row[7].value)  # Column H
        item_code = row[5].value          # Column F
        inventory_type = str(row[9].value)  # Column J

        if contractor_id in ["8993", "8982"] and inventory_type in [f"CTR.Subready.8993", f"CTR.Subready.8982"]:
            update_totals(combined_contractor_totals, item_code, ctr_devices)

    data_lst.append(format_totals(combined_contractor_totals, ctr_devices))

    # Get data for Warehouses
    for warehouse in warehouse_lst:
        warehouse_totals = {device: 0 for device in warehouse_devices}
        for row in sheet.iter_rows(min_row=3, max_row=sheet.max_row):
            contractor_id = str(row[1].value)  # Column B
            item_code = row[5].value          # Column F
            inventory_type = str(row[9].value)  # Column J

            if contractor_id == warehouse:
                update_totals(warehouse_totals, item_code, warehouse_devices)

        data_lst.append(format_totals(warehouse_totals, warehouse_devices))



    # Copy data to Excel with contractor/device names (order will be respected)
    full_list = robitaille_lst + ctr_lst  # Respect the order
    full_list = rearrange_lst(full_list, combined_contractor_totals)
    copy_data_to_excel(data_lst, full_list)

# Format totals for Excel output
def format_totals(totals, device_order):
    return '\n'.join(str(totals[device]) for device in device_order)

#Supposed to rearrange the list but I dont think it does.
def rearrange_lst(full_list, combined_contractor_totals):
    # Insert the combined contractor totals at the end
    full_list.append(combined_contractor_totals)

    # Dont even think this works or is importants...
    full_list.insert(5, full_list[7])
    full_list.insert(6, full_list[8])
    full_list.insert(12, full_list[14])

    return full_list

# Copies data to Excel based on all data gathered and helps navigate.
def copy_data_to_excel(data_lst, full_list):
    for index, data in enumerate(data_lst):
        clipboard.copy(data)
        time.sleep(6)  
        pyautogui.hotkey('ctrl', 'v')
        if index < len(data_lst) - 1:
            pyautogui.hotkey("ctrl", "alt", "pagedown")
            pyautogui.hotkey('ctrl','left')

# Converts XML files to serials
def open_xml_file():
    # Open a file dialog to choose the XML file
    file_path = filedialog.askopenfilename(
        title="Select an XML File",
        filetypes=[("XML files", "*.xml")]
    )
    
    # If the user selects a file, proceed to load it
    if file_path:
        try:
            # Parse the XML file
            tree = ET.parse(file_path)
            root = tree.getroot()
            
            # Extract all <SERIAL> values
            serial_numbers = [elem.text for elem in root.findall(".//SERIAL")]
            
            if serial_numbers:
                # Insert SERIAL numbers into the text box for easy copying
                serials_text.delete(1.0, tk.END)  # Clear the text box
                serials_text.insert(tk.END, "\n".join(serial_numbers))  # Insert serials into the text box
                clipboard.copy("\n".join(serial_numbers))
            else:
                messagebox.showinfo("No SERIAL Tags Found", "No <SERIAL> tags were found in the XML file.")
                
            return serial_numbers
        except Exception as e:
            # If there is an error, show it
            messagebox.showerror("Error", f"Failed to load XML file: {e}")
            return None
    else:
        messagebox.showwarning("No file selected", "Please select a valid XML file.")
        return None

#==============================================================================#
# GUI #
#==============================================================================#

OUTPUT_PATH = Path(__file__).parent
ASSETS_PATH = OUTPUT_PATH / Path(f"{asset_location}")
def relative_to_assets(path: str) -> Path:
    return ASSETS_PATH / Path(path)


window = Tk()
window.title("Rogers Toolbox 2.3")
window.geometry("600x350")
window.configure(bg = "#FFFFFF")


canvas = Canvas(
    window,
    bg = "#FFFFFF",
    height = 350,
    width = 600,
    bd = 0,
    highlightthickness = 0,
    relief = "ridge"
)
# The Background box.
canvas.place(x = 0, y = 0)
canvas.create_rectangle(
    0.0,
    0.0,
    600.0,
    350.0,
    fill="#D9D9D9",
    outline="")
# The Nav bar
canvas.create_rectangle(
    0.0,
    0.0,
    600.0,
    53.0,
    fill="#2D2D2D",
    outline="")
# The Print to lazer Printer button
button_image_1 = PhotoImage(
    file=relative_to_assets("button_1.png"))
button_1 = Button(
    image=button_image_1,
    borderwidth=0,
    highlightthickness=0,
    command=lambda: CreateLaser(),
    relief="flat"
)
button_1.place(
    x=554.0,
    y=9.0,
    width=30.0,
    height=35.0
)
# The CTR Update button
button_image_2 = PhotoImage(
    file=relative_to_assets("button_2.png"))
button_2 = Button(
    image=button_image_2,
    borderwidth=0,
    highlightthickness=0,
    command=lambda: open_file_dialog(),
    relief="flat"
)
button_2.place(
    x=487.0,
    y=3.0,
    width=52.0,
    height=46.0
)
# The convert XML Button
button_image_3 = PhotoImage(
    file=relative_to_assets("button_3.png"))
button_3 = Button(
    image=button_image_3,
    borderwidth=0,
    highlightthickness=0,
    command=lambda: open_xml_file(),
    relief="flat"
)
button_3.place(
    x=433.0,
    y=3.0,
    width=54.0,
    height=46.0
)
# The print Purolator Button
button_image_4 = PhotoImage(
    file=relative_to_assets("button_4.png"))
button_4 = Button(
    image=button_image_4,
    borderwidth=0,
    highlightthickness=0,
    command=lambda: CreatePurolatorSheet(),
    relief="flat"
)
button_4.place(
    x=389.0,
    y=5.0,
    width=44.0,
    height=42.0
)
#  The print Barcodes Button
button_image_5 = PhotoImage(
    file=relative_to_assets("button_5.png"))
button_5 = Button(
    image=button_image_5,
    borderwidth=0,
    highlightthickness=0,
    command=lambda: CreateBarcodes(),
    relief="flat"
)
button_5.place(
    x=331.0,
    y=0.0,
    width=46.0,
    height=53.0
)
# The open Excel button.
button_image_6 = PhotoImage(
    file=relative_to_assets("button_6.png"))
button_6 = Button(
    image=button_image_6,
    borderwidth=0,
    highlightthickness=0,
    command=lambda: OpenExcel(),
    relief="flat"
)
button_6.place(
    x=12.0,
    y=6.0,
    width=42.48554992675781,
    height=42.0
)
# The run button
button_image_8 = PhotoImage(
    file=relative_to_assets("button_8.png"))
button_8 = Button(
    image=button_image_8,
    borderwidth=0,
    highlightthickness=0,
    command=lambda: PasteSerialsFlexi(),
    relief="flat"
)
button_8.place(
    x=111.0,
    y=5.0,
    width=53.0,
    height=44.0
)

image_image_1 = PhotoImage(
    file=relative_to_assets("image_1.png"))
image_1 = canvas.create_image(
    300.0,
    71.0,
    image=image_image_1
)
# The entry Canvas
canvas.create_rectangle(
    12.0,
    90.0,
    588.0,
    335.0,
    fill="#B5B5B5",
    outline="")
# Places serials into the file.
serials_text = tk.Text(width=82, height=16, bg='#FFFFFF', fg='black',font=("Cabin", 12 * -1, "bold"))
serials_text.place(x=12,y=90)
canvas.create_text(
    255.0,
    64.0,
    anchor="nw",
    text="x Serials Loaded",
    fill="#2D2D2D",
    font=("Cabin", 12 * -1)
)

# Creates the label count for how many serials are loaded! Once again..
count_label = tk.Label(
    window,
    width=17,
    height=1,
    text="0 Serials Loaded",
    bg="#2D2D2D",
    fg="#FFFFFF",
    font=("Arial", 8, "bold")
)
count_label.place(x=224, y=60)

button_image_7 = PhotoImage(
    file=relative_to_assets("button_7.png"))
button_7 = Button(
    image=button_image_7,
    borderwidth=0,
    highlightthickness=0,
    command=lambda: PasteSerialsWMS(),
    relief="flat"
)
button_7.place(
    x=62.0,
    y=5.0,
    width=42.0,
    height=42.0
)
button_image_9 = PhotoImage(
    file=relative_to_assets("button_9.png"))
button_9 = Button(
    image=button_image_9,
    borderwidth=0,
    highlightthickness=0,
    command=lambda: PasteSerialsNormal(),
    relief="flat"
)
button_9.place(
    x=170.0,
    y=5.0,
    width=42.0,
    height=42.0
)
window.resizable(False, False)
window.mainloop()
