import tkinter as tk
from tkinter import filedialog, messagebox, PhotoImage
import openpyxl
import clipboard
import pyautogui
import time
import subprocess

# Global Variables
file_path = None
serials_list = []
remaining_serials = 0
notepad_path = "C:\\BTAutomation\\barcodes.txt"

def OpenExcel():
    global file_path
    file_path = filedialog.askopenfilename(title="Select a File", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    if file_path:
        selected_file_label.config(text=f"Selected File: {file_path}")
        LoadSerials(file_path)

def LoadSerials(file_path):
    global serials_list, remaining_serials
    # Tries opening Excel file and gets all data in the fist column and stores it in a list. 
    try:
        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active
        serials_list.clear()  # Clear existing serials

        # Start iterating from the first row
        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row):
            # Gets all data in the first column, as long as it is not empty
            if row[0].value is not None:
                serials_list.append(str(row[0].value))

        remaining_serials = int(len(serials_list))
        UpdateSerialsDisplay()
    
    # If the Excel fails to load you recieve an error.
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load serials: {e}")

# Displays the amount of serials the user has loaded.
def UpdateSerialsDisplay():
    serials_text.delete(1.0, tk.END)
    serials_text.insert(tk.END, "\n".join(serials_list))
    count_label.config(text=f"{remaining_serials} Serials Loaded")

# Pastes the serials in the order they were recieved from the excel until finished.
def PasteSerials():
    global remaining_serials
    if not serials_list:
        messagebox.showinfo("Info", "No serials loaded to paste.")
        return

    time.sleep(5)  # Allow time for the user to focus on the target program

    # Iterate over the serials list by creating a copy to avoid modification during iteration
    for serial in serials_list[:]:  
        if serial is not None:
            print(f"Pasting serial: {serial}")  # Debugging statement to track progress
            clipboard.copy(serial)
            pyautogui.hotkey("ctrl", "v")
            pyautogui.hotkey("down")
            serials_list.remove(serial)
            remaining_serials -= 1
            UpdateSerialsDisplay()
            time.sleep(0.5)  # Small delay between actions
        else:
            print("There are no serials loaded")  # Debugging statement to track None values

def MakeTVSheet(device):
    total_strings = len(serials_list)
    formatted_list = []

    for i in range(0, total_strings, 10):
        # Add device name at the beginning of each chunk.
        formatted_list.append(device)
        # Get the next 10 serials in our list of serials.
        chunk = serials_list[i:i+10]
        # Reverse the chunk
        chunk.reverse()
        # Add the reversed chunk to the formatted list
        formatted_list.extend(chunk)
    
    return '\n'.join(formatted_list)

# Helper Function for MakeModemSheet, it reverses the serials every 5 to compensate for box placement.
def ReverseForModems():
    total_strings = len(serials_list)
    reversed_modems = []
    for i in range(0, total_strings, 5):  
        chunk = serials_list[i:i + 5][::-1]  
        reversed_modems.extend(chunk)
    return reversed_modems

# This takes the reversed list from our helper function and adds the device name after every 8 devices.
def MakeModemSheet(device):
    working_modems = ReverseForModems()
    total_strings = len(serials_list)
    formatted_list = []

    for i in range(0, total_strings, 8):
        # Add device name at the beginning of each chunk
        formatted_list.append(device)
        
        # Get the next 8 serials
        chunk = working_modems[i:i+8]
        # Add the reversed chunk to the formatted list
        formatted_list.extend(chunk)
    
    return '\n'.join(formatted_list)

# Creates Purolator sheets formatted based on the device.
def CreatePurolatorSheet():
    # Finds the device name based on the initals of the first scanned item.
    global serials_list
    device = ""
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
    
    # Decides which way to print the purolator sheet based on the formatting of the devices.
    if device in ["IPTVARXI6HD", "IPTVTCXI6HD", "SCXI11BEI"]:
        puro_sheet = MakeTVSheet(device)
        # Clears barcode notepad, then puts the purolator sheet for our function into it.
        with open(notepad_path, 'w') as file:
            file.write(f"{puro_sheet}\n")
        
        # writes the script for CMD that will select the proper printer as default and then print the purolator papers for the barcodes.
        cmd_script = """
        @echo off

        rem Define the share name of the printer you want to switch to
        set "target_printer=55EXP_Purolator"

        rem Check if the printer with the target share name exists and set it as default
        powershell -Command "Get-WmiObject -Query \\"SELECT * FROM Win32_Printer WHERE ShareName='%target_printer%'\\" | Invoke-WmiMethod -Name SetDefaultPrinter"

        echo Hello! From the Purolator Printer.

        rem Execute BarTender command
        "C:\\Seagull\\BarTender 7.10\\Standard\\bartend.exe" /f=C:\\BTAutomation\\XI6.btw /p /x
        """
        
        # Writes the command to a temporary batch file
        with open("temp_cmd.bat", "w") as bat_file:
            bat_file.write(cmd_script)
            
        # Executes the batch file
        result = subprocess.run(["cmd.exe", "/c", "temp_cmd.bat"])


    else:
        puro_sheet = MakeModemSheet(device)
        with open(notepad_path, 'w') as file:
            file.write(f"{puro_sheet}\n")
        cmd_script = """
        @echo off

        rem Define the share name of the printer you want to switch to
        set "target_printer=55EXP_Purolator"

        rem Check if the printer with the target share name exists and set it as default
        powershell -Command "Get-WmiObject -Query \\"SELECT * FROM Win32_Printer WHERE ShareName='%target_printer%'\\" | Invoke-WmiMethod -Name SetDefaultPrinter"

        echo Hello! From the Purolator Printer.

        rem Execute BarTender command
        "C:\\Seagull\\BarTender 7.10\\Standard\\bartend.exe" /f=C:\\BTAutomation\\CODA.btw /p /x
        """
        # Writes the command to a temporary batch file
        with open("temp_cmd.bat", "w") as bat_file:
            bat_file.write(cmd_script)
            
        # Executes the batch file
        result = subprocess.run(["cmd.exe", "/c", "temp_cmd.bat"])

# Clears the notepad file and appends all serials, then runs CMD command
def CreateBarcodes():
    
    # Clears the notepad file
    with open(notepad_path, 'w') as file:
        for serial in serials_list:
            file.write(f"{serial}\n")

    # Execute CMD commands for changing to the right printer and printing barcodes for every serial.
    cmd_script = """
    @echo off
    rem Define the share name of the printer you want to switch to
    set "target_printer=55EXP_Barcode"
        
    rem Check if the printer with the target share name exists and set it as default
    powershell -Command "Get-WmiObject -Query \\"SELECT * FROM Win32_Printer WHERE ShareName='%target_printer%'\\" | Invoke-WmiMethod -Name SetDefaultPrinter"
    
    rem Execute BarTender command
    "C:\\Seagull\\BarTender 7.10\\Standard\\bartend.exe" /f=C:\\BTAutomation\\singlebar.btw /p /x
    """
        
    # Writes the command to a temporary batch file
    with open("temp_cmd.bat", "w") as bat_file:
        bat_file.write(cmd_script)
        
    # Executes the batch file
    result = subprocess.run(["cmd.exe", "/c", "temp_cmd.bat"])

# Creates the GUI for the application
def CreateGUI():
    root = tk.Tk()
    root.title("Rogers Import Tool")
    root.configure(bg='black')

    # Create a frame for the navigation bar
    nav_bar = tk.Frame(root, bg='red')
    nav_bar.pack(fill=tk.X, padx=10, pady=10)

    global selected_file_label, serials_text, count_label

    # Load icons
    select_icon = PhotoImage(file="Excel Icon.png")  # Replace with your icon path
    start_icon = PhotoImage(file="Rogers Icon.png")    # Replace with your icon path
    barcodes_icon = PhotoImage(file="Barcode Icon.png")  # Replace with your icon path
    purolator_icon = PhotoImage(file="Puro Icon.png")  # Replace with your icon path

    # File selection button with icon
    select_button = tk.Button(nav_bar, image=select_icon, command=OpenExcel, bg='red', borderwidth=0)
    select_button.pack(side=tk.LEFT, padx=5)

    # Start button with icon
    start_button = tk.Button(nav_bar, image=start_icon, command=PasteSerials, bg='red', borderwidth=0)
    start_button.pack(side=tk.LEFT, padx=5)

    # Clear and append button with icon
    clear_append_button = tk.Button(nav_bar, image=barcodes_icon, command=CreateBarcodes, bg='red', borderwidth=0)
    clear_append_button.pack(side=tk.LEFT, padx=5)

    # Purolator button with icon
    purolator_button = tk.Button(nav_bar, image=purolator_icon, command=CreatePurolatorSheet, bg='red', borderwidth=0)
    purolator_button.pack(side=tk.LEFT, padx=5)

    # Main content area
    content_frame = tk.Frame(root, bg='black')
    content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # Selected file label
    selected_file_label = tk.Label(content_frame, text="No file selected", bg='black', fg='white')
    selected_file_label.pack(pady=5)

    # Serial numbers display
    serials_text = tk.Text(content_frame, width=30, height=15, bg='black', fg='white')
    serials_text.pack(pady=5)

    # Remaining count label
    count_label = tk.Label(content_frame, text="0/0 remaining", bg='black', fg='white')
    count_label.pack(pady=5)

    # Start the Tkinter event loop
    root.mainloop()

CreateGUI()
