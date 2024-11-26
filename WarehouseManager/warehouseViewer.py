import firebase_admin
from firebase_admin import credentials, firestore
import tkinter as tk
from tkinter import messagebox, simpledialog, Toplevel, Scrollbar, Text
mark_in_use_mode = False
# Initialize Firebase Admin SDK with Firestore
cred = credentials.Certificate('C:\\Users\\jessi\\source\\repos\\RogersQuery\\warehousemanager-37f64-firebase-adminsdk-8641r-3d99759038.json')
firebase_admin.initialize_app(cred)

# Initialize Firestore client
db = firestore.client()
def toggle_mark_in_use():
    """Toggle the 'Mark as In Use' mode."""
    global mark_in_use_mode
    mark_in_use_mode = not mark_in_use_mode
    mode_status = "ON" if mark_in_use_mode else "OFF"
    messagebox.showinfo("Mode Toggled", f"'Mark as In Use' mode is now {mode_status}.")

def fetch_data():
    """Fetch data from Firestore and update the warehouse grid."""
    warehouse_ref = db.collection('warehouse')
    docs = warehouse_ref.stream()
    data = {doc.id: doc.to_dict() for doc in docs}

    if data:
        update_grid(data)
    else:
        messagebox.showinfo("Info", "No data found in Firestore.")

def update_grid(data):
    """Update the grid based on Firestore data."""
    for row, rack in enumerate(warehouse_layout, start=1):
        for col, location in enumerate(rack, start=1):
            if location == "not a location":
                label = tk.Label(root, text="X", relief="solid", width=10, height=5, bg="gray")
            else:
                location_data = data.get(location, {})
                device = location_data.get('device', "Empty")
                serials = location_data.get('serials', [])
                inventory_number = location_data.get('inventory_number', "")
                bg_color = "#765341" if device != "Empty" else "white"
                fg_color = "white" if device != "Empty" else "black"

                label = tk.Label(root, text=f"{inventory_number}\n{device}\n{location}", relief="solid", width=10, height=5, bg=bg_color, fg=fg_color)
                label.grid(row=row, column=col, padx=2, pady=2)
                label.bind("<Button-1>", lambda e, loc=location: show_serials(data.get(loc, {})))

def show_serials(data):
    """Display the serials in a new dialog box."""
    if not data:
        messagebox.showinfo("Info", "No data available for this location.")
        return
    
    serials = data.get('serials', [])
    serial_text = "\n".join(serials) if serials else "No serials available."

    # Create a new top-level window
    serial_window = Toplevel(root)
    serial_window.title(data.get('inventory_number') + " " + data.get("device"))

    # Add scrollable text box
    scrollbar = Scrollbar(serial_window)
    scrollbar.pack(side="right", fill="y")

    text = Text(serial_window, wrap="word", yscrollcommand=scrollbar.set, width=50, height=15)
    text.insert("1.0", serial_text)
    text.configure(state="disabled")  # Make it read-only
    text.pack(side="left", fill="both", expand=True)

    scrollbar.config(command=text.yview)
def add_mass():
    """Add inventory to multiple locations using a single dialog box."""
    def submit():
        entries = text_area.get("1.0", tk.END).strip().split("\n")
        invalid_locations = []

        for entry in entries:
            try:
                # Split each line into components (e.g., location, device, inventory number, serials)
                location, device, inventory_number, *serials = entry.split(";")
                location, device, inventory_number = location.strip(), device.strip(), inventory_number.strip()
                serials = [s.strip() for s in serials if s.strip()]

                if not all([location, device, inventory_number]) or location not in valid_locations:
                    invalid_locations.append(location)
                    continue

                # Save to Firestore
                warehouse_ref = db.collection('warehouse').document(location)
                warehouse_ref.set({
                    "device": device,
                    "inventory_number": inventory_number,
                    "serials": serials
                })
            except ValueError:
                invalid_locations.append(entry)

        fetch_data()
        if invalid_locations:
            messagebox.showwarning("Warning", f"Some locations were invalid:\n{', '.join(invalid_locations)}")
        dialog.destroy()

    # Create the dialog window
    dialog = Toplevel(root)
    dialog.title("Add Mass Inventory")
    dialog.geometry("500x400")

    tk.Label(dialog, text="Enter data for multiple locations (one per line):\nFormat: Location;Device;Inventory Number;Serial1;Serial2;...").pack(pady=10)
    
    text_area = Text(dialog, wrap="word", width=60, height=20)
    text_area.pack(padx=10, pady=10)

    submit_button = tk.Button(dialog, text="Submit", command=submit)
    submit_button.pack(pady=10)

    # Make dialog modal
    dialog.transient(root)
    dialog.grab_set()
    root.wait_window(dialog)

    
def add_inventory():
    """Add inventory to a specific location with a single dialog box."""
    def submit():
        location = location_entry.get().strip()
        device = device_entry.get().strip()
        inventory_number = inventory_number_entry.get().strip()
        serials = serials_text.get("1.0", tk.END).strip().split("\n")

        if not all([location, device, inventory_number, serials]):
            messagebox.showerror("Error", "All fields are required.")
            return

        if location not in valid_locations:
            messagebox.showerror("Error", "Invalid location.")
            return

        # Save to Firestore
        warehouse_ref = db.collection('warehouse').document(location)
        warehouse_ref.set({
            "device": device,
            "inventory_number": inventory_number,
            "serials": serials
        })
        fetch_data()
        dialog.destroy()

    # Create the dialog window
    dialog = Toplevel(root)
    dialog.title("Add Inventory")
    dialog.geometry("400x300")

    # Location input
    tk.Label(dialog, text="Location:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
    location_entry = tk.Entry(dialog, width=30)
    location_entry.grid(row=0, column=1, padx=5, pady=5)

    # Device input
    tk.Label(dialog, text="Device:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
    device_entry = tk.Entry(dialog, width=30)
    device_entry.grid(row=1, column=1, padx=5, pady=5)

    # Inventory number input
    tk.Label(dialog, text="Inventory Number:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
    inventory_number_entry = tk.Entry(dialog, width=30)
    inventory_number_entry.grid(row=2, column=1, padx=5, pady=5)

    # Serials input
    tk.Label(dialog, text="Serials (one per line):").grid(row=3, column=0, sticky="nw", padx=5, pady=5)
    serials_text = Text(dialog, wrap="word", width=30, height=10)
    serials_text.grid(row=3, column=1, padx=5, pady=5)

    # Submit button
    submit_button = tk.Button(dialog, text="Submit", command=submit)
    submit_button.grid(row=4, column=1, pady=10, sticky="e")

    # Make dialog modal
    dialog.transient(root)
    dialog.grab_set()
    root.wait_window(dialog)

def move_inventory():
    """Move inventory from one location to another."""
    from_location = simpledialog.askstring("Input", "Enter the source location:")
    if from_location not in valid_locations:
        messagebox.showerror("Error", "Invalid source location.")
        return

    to_location = simpledialog.askstring("Input", "Enter the destination location:")
    if to_location not in valid_locations:
        messagebox.showerror("Error", "Invalid destination location.")
        return

    warehouse_ref = db.collection('warehouse').document(from_location)
    doc = warehouse_ref.get()

    if not doc.exists:
        messagebox.showerror("Error", "Source location is empty.")
        return

    data = doc.to_dict()
    # Move data to the new location
    db.collection('warehouse').document(to_location).set(data)
    # Clear the source location
    warehouse_ref.delete()
    fetch_data()

# Define warehouse layout
warehouse_layout = [
    ["A1.08.F3", "A1.08.F2", "A1.08.F1", "A1.07.J2", "A1.07.J1", "A1.06.J2", "A1.06.J1", 
     "A1.05.J2", "A1.05.J1", "A1.04.J2", "A1.04.J1", "A1.03.J2", "A1.03.J1", 
      "A1.02.J2", "A1.02.J1",  "A1.01.J2", "A1.01.J1"],
    ["A1.08.E3", "A1.08.E2", "A1.08.E1", "A1.07.H2", "A1.07.H1", "A1.06.H2", "A1.06.H1", 
      "A1.05.H2", "A1.05.H1",  "A1.04.H2", "A1.04.H1", "A1.03.H2", "A1.03.H1", 
      "A1.02.H2", "A1.02.H1",  "A1.01.H2", "A1.01.H1"],
    ["A1.08.D3", "A1.08.D2", "A1.08.D1", "A1.07.G2", "A1.07.G1", "A1.06.G2", "A1.06.G1", 
     "A1.05.G2", "A1.05.G1", "A1.04.G2", "A1.04.G1",  "A1.03.G2", "A1.03.G1", 
      "A1.02.G2", "A1.02.G1", "A1.01.G2", "A1.01.G1"],
    ["not a location", "not a location", "not a location", "A1.07.F2", "A1.07.F1", "A1.06.F2", "A1.06.F1", 
      "A1.05.F2", "A1.05.F1", "A1.04.F2", "A1.04.F1", "A1.03.F2", "A1.03.F1", 
      "A1.02.F2", "A1.02.F1", "A1.01.F2", "A1.01.F1"]
]

# Flatten the layout to get valid locations
valid_locations = {loc for row in warehouse_layout for loc in row if loc != "not a location"}

# Initialize the GUI
root = tk.Tk()
root.title("Warehouse Inventory")

# Buttons for adding and moving inventory
btn_add = tk.Button(root, text="Add Inventory", command=add_inventory)
btn_add.grid(row=0, column=0, columnspan=3, sticky="ew")

btn_move = tk.Button(root, text="Move Inventory", command=move_inventory)
btn_move.grid(row=0, column=3, columnspan=3, sticky="ew")

btn_add_mass = tk.Button(root, text="ADD", command=add_mass)
btn_add_mass.grid(row=0, column=6, columnspan=3, sticky="ew")
# Fetch initial data
fetch_data()
rows = 4  # Number of floors
columns = 10  # Number of boxes per floor
cell_size = 50  # Size of each box
pillar_width = 10  # Width of the blue pillars




# Start the GUI event loop
root.mainloop()
