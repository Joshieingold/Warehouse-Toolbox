from pathlib import Path

# from tkinter import *
# Explicit imports to satisfy Flake8
from tkinter import Tk, Canvas, Entry, Text, Button, PhotoImage
import tkinter as tk
from tkinter import filedialog, messagebox, PhotoImage
import openpyxl
import clipboard
import pyautogui
import time
import subprocess
from pathlib import Path

file_path = None
serials_list = []
remaining_serials = 0
notepad_path = "C:\\BTAutomation\\barcodes.txt"
count_label = str(0)
OUTPUT_PATH = Path(__file__).parent
ASSETS_PATH = OUTPUT_PATH / Path(r"C:\Users\jessi\OneDrive\Documents\Projects\Rogers Import Tool")


def relative_to_assets(path: str) -> Path:
    return ASSETS_PATH / Path(path)


# Functions (no changes to the original function definitions)
def OpenExcel():
    global file_path
    file_path = filedialog.askopenfilename(title="Select a File", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    if file_path:
        LoadSerials(file_path)


def LoadSerials(file_path):
    global serials_list, remaining_serials
    try:
        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active
        serials_list.clear()

        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row):
            if row[0].value is not None:
                serials_list.append(str(row[0].value))

        remaining_serials = len(serials_list)
        UpdateSerialsDisplay()

    except Exception as e:
        messagebox.showerror("Error", f"Failed to load serials: {e}")

def UpdateSerialsDisplay():
    serials_text.delete(1.0, tk.END)
    serials_text.insert(tk.END, "\n".join(serials_list))
    count_label.config(text=f"{remaining_serials} Serials Loaded")

def PasteSerials():
    global remaining_serials
    if not serials_list:
        messagebox.showinfo("Info", "No serials loaded to paste.")
        return

    time.sleep(5)  # Allow time for the user to focus on the target program

    for serial in serials_list[:]:
        if serial is not None:
            clipboard.copy(serial)
            pyautogui.hotkey("ctrl", "v")
            pyautogui.hotkey("down")
            serials_list.remove(serial)
            remaining_serials -= 1
            UpdateSerialsDisplay()
            time.sleep(0.5)


def MakeTVSheet(device):
    total_strings = len(serials_list)
    formatted_list = []

    for i in range(0, total_strings, 10):
        formatted_list.append(device)
        chunk = serials_list[i:i + 10]
        chunk.reverse()
        formatted_list.extend(chunk)

    return '\n'.join(formatted_list)


def ReverseForModems():
    total_strings = len(serials_list)
    reversed_modems = []
    for i in range(0, total_strings, 5):
        chunk = serials_list[i:i + 5][::-1]
        reversed_modems.extend(chunk)
    return reversed_modems


def MakeModemSheet(device):
    working_modems = ReverseForModems()
    total_strings = len(serials_list)
    formatted_list = []

    for i in range(0, total_strings, 8):
        formatted_list.append(device)
        chunk = working_modems[i:i + 8]
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
        set "target_printer=55EXP_Purolator"
        powershell -Command "Get-WmiObject -Query \\"SELECT * FROM Win32_Printer WHERE ShareName='%target_printer%'\\" | Invoke-WmiMethod -Name SetDefaultPrinter"
        "C:\\Seagull\\BarTender 7.10\\Standard\\bartend.exe" /f=C:\\BTAutomation\\XI6.btw /p /x
        """

        with open("temp_cmd.bat", "w") as bat_file:
            bat_file.write(cmd_script)

        subprocess.run(["cmd.exe", "/c", "temp_cmd.bat"])

    else:
        puro_sheet = MakeModemSheet(device)
        with open(notepad_path, 'w') as file:
            file.write(f"{puro_sheet}\n")
        cmd_script = """
        @echo off
        set "target_printer=55EXP_Purolator"
        powershell -Command "Get-WmiObject -Query \\"SELECT * FROM Win32_Printer WHERE ShareName='%target_printer%'\\" | Invoke-WmiMethod -Name SetDefaultPrinter"
        "C:\\Seagull\\BarTender 7.10\\Standard\\bartend.exe" /f=C:\\BTAutomation\\CODA.btw /p /x
        """
        with open("temp_cmd.bat", "w") as bat_file:
            bat_file.write(cmd_script)

        subprocess.run(["cmd.exe", "/c", "temp_cmd.bat"])


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




OUTPUT_PATH = Path(__file__).parent
ASSETS_PATH = OUTPUT_PATH / Path(r"C:\Users\jessi\Music\build\assets\frame0")


def relative_to_assets(path: str) -> Path:
    return ASSETS_PATH / Path(path)


window = Tk()
window.title("Import Tool")
window.geometry("350x500")
window.configure(bg = "#FFFFFF")


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

image_image_1 = PhotoImage(
    file=relative_to_assets("image_1.png"))
image_1 = canvas.create_image(
    81.26589965820312,
    87.0,
    image=image_image_1
)

button_image_1 = PhotoImage(file=relative_to_assets("button_1.png"))
button_1 = Button(
    image=button_image_1,
    borderwidth=0,
    highlightthickness=0,
    command=OpenExcel,  # Connects to OpenExcel function
    relief="flat"
)
button_1.place(
    x=10.1156005859375,
    y=270.0,
    width=148.6994171142578,
    height=42.0
)

button_image_2 = PhotoImage(
    file=relative_to_assets("button_2.png"))
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

button_image_3 = PhotoImage(
    file=relative_to_assets("button_3.png"))
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

button_image_4 = PhotoImage(
    file=relative_to_assets("button_4.png"))
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

canvas.create_rectangle(
    169.94219970703125,
    0.0,
    350.0,
    500.0,
    fill="#D9D9D9",
    outline="")

image_image_2 = PhotoImage(
    file=relative_to_assets("image_2.png"))
image_2 = canvas.create_image(
    260.10406494140625,
    21.0,
    image=image_image_2
)

canvas.create_text(
    212.0,
    14.0,
    anchor="nw",
    text=f"{count_label} Serials Loaded\n",
    fill="#D9D9D9",
    font=("Inter Black", 12 * -1)
)

canvas.create_rectangle(
    184.10406494140625,
    40.0,
    336.8497314453125,
    491.0,
    fill="#B5B5B5",
    outline="")

serials_text = tk.Text(
    window,
    height=15,
    width=40,
    bg="#F0F0F0"
)
serials_text.place(x=336.8497314453125, y=491.0)

count_label = tk.Label(
    window,
    text="0 Serials Loaded",
    bg="#FF1B1F",
    fg="#FFFFFF",
    font=("Arial", 8, "bold")
)
count_label.place(x=212, y=11)
serials_text = tk.Text(width=25, height=32, bg='gray90', fg='black',font=("Arial", 8, "bold"))
serials_text.place(x=183,y=41)
window.resizable(False, False)
window.mainloop()

