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

# Global Variables.
file_path = None
serials_list = []
remaining_serials = 0
notepad_path = "C:\\BTAutomation\\barcodes.txt"
count_label = str(0)
OUTPUT_PATH = Path(__file__).parent
ASSETS_PATH = OUTPUT_PATH / Path(r"C:\Users\jessi\OneDrive\Documents\Projects\Rogers Import Tool")


# Allows the user to easily and interchangeably apply their path by changing the global variable.
def RelativeToAssets(path: str) -> Path:
    return ASSETS_PATH / Path(path)

# Opens Excel file and puts the serials from the first column of the active sheet into our serial list with the LoadSerials function.
def OpenExcel():
    global file_path
    file_path = filedialog.askopenfilename(title="Select a File", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    if file_path:
        LoadSerials(file_path)

# Gets the serials from the Excel file.
def LoadSerials(file_path):
    global serials_list, remaining_serials
    try:
        # Opens the Excel, loads at the active sheet and clears the serials previously loaded if any.
        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active
        serials_list.clear()
        # Gets every serial in the first column and adds it into the serial list.
        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row):
            if row[0].value is not None:
                serials_list.append(str(row[0].value))
        # Gets the amount of serials loaded and updates the display with the function.
        remaining_serials = len(serials_list)
        UpdateSerialsDisplay()
    # If serials are not able to be loaded by the file, this handles the error.
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load serials: {e}")

# Deletes all serials in the display and then puts the new ones in seperating them by new lines and updates the GUI of the serial amounts.
def UpdateSerialsDisplay():
    serials_text.delete(1.0, tk.END)
    serials_text.insert(tk.END, "\n".join(serials_list))
    count_label.config(text=f"{remaining_serials} Serials Loaded")

# Pastes the serials automatically.
def PasteSerials():
    global remaining_serials
    # Handles if there are no serials in our serial list.
    if not serials_list:
        messagebox.showinfo("Info", "No serials loaded to paste.")
        return
    # Allows time for the user to focus on the target program
    time.sleep(5)  
    # For every serial in the list it is pasted and then down arrow is pressed. Serials are then removed and I attempt to make the Gui update to no avail.
    for serial in serials_list[:]:
        if serial is not None:
            clipboard.copy(serial)
            pyautogui.hotkey("ctrl", "v")
            pyautogui.hotkey("down")
            serials_list.remove(serial)
            remaining_serials -= 1
            UpdateSerialsDisplay()
            # Time between pastes, ideally rounded down as much as possible but oracle is finicy 
            time.sleep(0.5)

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

# Gets the reveresed serials from helper function and for every 8 serials it puts the device name at the top of the sheet.
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
        with open(notepad_path, 'w') as file:
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
        with open(notepad_path, 'w') as file:
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

# Takes the serials in the serial list and prints them with the appropriate printer.
def CreateBarcodes():
    with open(notepad_path, 'w') as file:
        for serial in serials_list:
            file.write(f"{serial}\n")

    cmd_script = """
    @echo off
    set "target_printer=55EXP_Barcode"
    powershell -Command "Get-WmiObject -Query \\"SELECT * FROM Win32_Printer WHERE ShareName='%target_printer%'\\" | Invoke-WmiMethod -Name SetDefaultPrinter"
    "C:\\Seagull\\BarTender 7.10\\Standard\\bartend.exe" /f=C:\\BTAutomation\\singlebar.btw /p /x
    """

    with open("temp_cmd.bat", "w") as bat_file:
        bat_file.write(cmd_script)

    subprocess.run(["cmd.exe", "/c", "temp_cmd.bat"])

# Creating GUI

# The window is created, named and proportioned.
window = Tk()
window.title("Import Tool")
window.geometry("350x500")
window.configure(bg = "#FFFFFF")


# This is the Red part that contains the buttons.
canvas = Canvas(
    window,
    bg = "#FFFFFF",
    height = 500,
    width = 350,
    bd = 0,
    highlightthickness = 0,
    relief = "ridge"
)
canvas.place(x = 0, y = 0)
canvas.create_rectangle(
    0.0,
    0.0,
    350.0,
    500.0,
    fill="#FF1B1F",
    outline="")

# Creates and places the Rogers Logo
image_image_1 = PhotoImage(
    file=RelativeToAssets("image_1.png"))
image_1 = canvas.create_image(
    81.26589965820312,
    87.0,
    image=image_image_1
)

# Creates OpenExcel icon
button_image_1 = PhotoImage(file=RelativeToAssets("button_1.png"))
button_1 = Button(
    image=button_image_1,
    borderwidth=0,
    highlightthickness=0,
    command=OpenExcel,  
    relief="flat"
)
button_1.place(
    x=10.1156005859375,
    y=270.0,
    width=148.6994171142578,
    height=42.0
)

# Creates CreateBarcodes icon
button_image_2 = PhotoImage(
    file=RelativeToAssets("button_2.png"))
button_2 = Button(
    image=button_image_2,
    borderwidth=0,
    highlightthickness=0,
    command=CreateBarcodes,
    relief="flat"
)
button_2.place(
    x=7.0809326171875,
    y=334.0,
    width=151.7340850830078,
    height=42.0
)

# Creates Run icon
button_image_3 = PhotoImage(
    file=RelativeToAssets("button_3.png"))
button_3 = Button(
    image=button_image_3,
    borderwidth=0,
    highlightthickness=0,
    command=PasteSerials,
    relief="flat"
)
button_3.place(
    x=10.1156005859375,
    y=206.0,
    width=144.65318298339844,
    height=42.0
)

# Creates CreatePuro icon
button_image_4 = PhotoImage(
    file=RelativeToAssets("button_4.png"))
button_4 = Button(
    image=button_image_4,
    borderwidth=0,
    highlightthickness=0,
    command=CreatePurolatorSheet,
    relief="flat"
)
button_4.place(
    x=10.1156005859375,
    y=398.0,
    width=148.6994171142578,
    height=42.0
)

# This is the Red part that contains the buttons.. Again?
canvas.create_rectangle(
    169.94219970703125,
    0.0,
    350.0,
    500.0,
    fill="#D9D9D9",
    outline="")

# creates the Loaded serials Contained
image_image_2 = PhotoImage(
    file=RelativeToAssets("image_2.png"))
image_2 = canvas.create_image(
    260.10406494140625,
    21.0,
    image=image_image_2
)

# Creates the label count for how many serials are loaded
canvas.create_text(
    212.0,
    14.0,
    anchor="nw",
    text=f"{count_label} Serials Loaded\n",
    fill="#D9D9D9",
    font=("Inter Black", 12 * -1)
)

# Creates the label count for how many serials are loaded! Once again..
count_label = tk.Label(
    window,
    text="0 Serials Loaded",
    bg="#FF1B1F",
    fg="#FFFFFF",
    font=("Arial", 8, "bold")
)
count_label.place(x=212, y=11)

# Places serials into the file.
serials_text = tk.Text(width=25, height=32, bg='gray90', fg='black',font=("Arial", 8, "bold"))
serials_text.place(x=183,y=41)

# Has to be the size its loaded as.
window.resizable(True, True)
window.mainloop()

