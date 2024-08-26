import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
import clipboard
import pyautogui
import time
import subprocess

# Global
file_path = None
serials_list = []
remaining_serials = 0
notepad_path = "C:\\BTAutomation\\barcodes.txt"

def open_excel():
    global file_path
    file_path = filedialog.askopenfilename(title="Select a File", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    if file_path:
        selected_file_label.config(text=f"Selected File: {file_path}")
        load_serials(file_path)

def load_serials(file_path):
    global serials_list, remaining_serials
    try:
        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active
        serials_list.clear()  # Clear existing serials

        # Start iterating from the first row
        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row):
            # Access the first column only and check if the cell is not empty
            if row[0].value is not None:
                serials_list.append(str(row[0].value))

        remaining_serials = len(serials_list)
        update_serials_display()

    except Exception as e:
        messagebox.showerror("Error", f"Failed to load serials: {e}")

def update_serials_display():
    serials_text.delete(1.0, tk.END)
    serials_text.insert(tk.END, "\n".join(serials_list))
    # Display the correct count: current remaining and total number loaded initially
    count_label.config(text=f"{remaining_serials}/{len(serials_list) + remaining_serials} remaining")

def paste_serials():
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
            update_serials_display()
            time.sleep(0.5)  # Small delay between actions
        else:
            print("Encountered None in serial list")  # Debugging statement to track None values

def MakeTVSheet(device):
    total_strings = len(serials_list)
    formatted_list = []

    for i in range(0, total_strings, 10):
        # Add device name and the number 10 at the beginning of each chunk
        formatted_list.append(device)
        
        # Get the next 10 strings (or fewer if at the end of the list)
        chunk = serials_list[i:i+10]
        # Reverse the chunk
        chunk.reverse()
        # Add the reversed chunk to the formatted list
        formatted_list.extend(chunk)
    
    return '\n'.join(formatted_list)


def MakeModemSheet(device):
    total_strings = len(serials_list)
    formatted_list = []

    for i in range(0, total_strings, 8):
        # Add device name and the number 10 at the beginning of each chunk
        formatted_list.append(device)
        
        # Get the next 10 strings (or fewer if at the end of the list)
        chunk = serials_list[i:i+5]
        # Reverse the chunk
        chunk.reverse()
        # Add the reversed chunk to the formatted list
        formatted_list.extend(chunk)
    
    return '\n'.join(formatted_list)

def CreatePurolatorSheet():
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
    
    if device in ["IPTVARXI6HD", "IPTVTCXI6HD", "SCXI11BEI"]:
        puro_sheet = MakeTVSheet(device)
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
        "C:\\Seagull\\BarTender 7.10\\Standard\\bartend.exe" /f=C:\\BTAutomation\\XI6.btw /p /x
        """
        
        # Write the command to a temporary batch file
        with open("temp_cmd.bat", "w") as bat_file:
            bat_file.write(cmd_script)
            
        # Execute the batch file and capture output and errors
        result = subprocess.run(["cmd.exe", "/c", "temp_cmd.bat"], capture_output=True, text=True)

        if result.returncode == 0:
            messagebox.showinfo("Info", "Serials appended to notepad and CMD commands executed successfully.")
        else:
            messagebox.showerror("Error", f"Failed to append serials or execute CMD command.\nReturn Code: {result.returncode}\nOutput: {result.stdout}\nError: {result.stderr}")

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

    

def clear_and_append_serials():
    """ Clears the notepad file and appends all serials, then runs CMD command. """
    
    # Clear the notepad file
    with open(notepad_path, 'w') as file:
        for serial in serials_list:
            file.write(f"{serial}\n")

    # Execute CMD commands
    cmd_script = """
    @echo off
    rem Define the share name of the printer you want to switch to
    set "target_printer=55EXP_Barcode"
        
    rem Check if the printer with the target share name exists and set it as default
    powershell -Command "Get-WmiObject -Query \\"SELECT * FROM Win32_Printer WHERE ShareName='%target_printer%'\\" | Invoke-WmiMethod -Name SetDefaultPrinter"
    
    rem Execute BarTender command
    "C:\\Seagull\\BarTender 7.10\\Standard\\bartend.exe" /f=C:\\BTAutomation\\singlebar.btw /p /x
    """
        
    # Write the command to a temporary batch file
    with open("temp_cmd.bat", "w") as bat_file:
        bat_file.write(cmd_script)
        
    # Execute the batch file and capture output and errors
    result = subprocess.run(["cmd.exe", "/c", "temp_cmd.bat"], capture_output=True, text=True)

    if result.returncode == 0:
        messagebox.showinfo("Info", "Serials appended to notepad and CMD commands executed successfully.")
    else:
        messagebox.showerror("Error", f"Failed to append serials or execute CMD command.\nReturn Code: {result.returncode}\nOutput: {result.stdout}\nError: {result.stderr}")

def create_gui():
    root = tk.Tk()
    root.title("Rogers Import Tool")
    root.configure(bg='black')

    global selected_file_label, serials_text, count_label

    # File selection button
    select_button = tk.Button(root, text="Select Excel File", command=open_excel, bg='gray', fg='white')
    select_button.grid(row=0, column=0, padx=10, pady=10)

    # Selected file label
    selected_file_label = tk.Label(root, text="No file selected", bg='black', fg='white')
    selected_file_label.grid(row=1, column=0, padx=10, pady=10)

    # Serial numbers display
    serials_text = tk.Text(root, width=30, height=15, bg='black', fg='white')
    serials_text.grid(row=2, column=0, padx=10, pady=10)

    # Remaining count label
    count_label = tk.Label(root, text="0/0 remaining", bg='black', fg='white')
    count_label.grid(row=3, column=0, padx=10, pady=10)

    # Start button
    start_button = tk.Button(root, text="Start", command=paste_serials, bg='red', fg='white')
    start_button.grid(row=4, column=0, padx=10, pady=10)

    # Clear and append button
    clear_append_button = tk.Button(root, text="Print Barcodes", command=clear_and_append_serials, bg='gray', fg='white')
    clear_append_button.grid(row=5, column=0, padx=10, pady=10)

    # Purolator button
    purolator_button = tk.Button(root, text="Print Purolator Sheets", command=CreatePurolatorSheet, bg='gray', fg='white')
    purolator_button.grid(row=6, column=0, padx=10, pady=10)

    root.mainloop()

create_gui()
