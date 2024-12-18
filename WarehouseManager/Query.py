import sqlite3

# Sample data
data = [
    ('B1-705', 'O2.08.A3', 'DCX3510', 'NO ACCOUNT'),
    ('B1-892', 'O2.08.B3', 'Unknown', 'NO ACCOUNTS'),
    ('B1-915', 'A2.07.G1', 'DCX3510', 'NO ACCOUNTS'),
    ('B1-923', 'C1.07.G2', 'DCX3450', ''),
    ('B1-925', 'A2.07.F1', 'CODA4582', 'ONE LEFT'),
    ('B1-926', 'A2.03.F1', 'DCX3510', ''),
    ('C-001', 'C1.01.F2', 'CODA4582', ''),
    ('P-001', 'C1.07.F1', 'ARRIS602G', ''),
    ('Q-002', 'A1.04.H2', 'XB7', ''),
    ('Q-003', 'C1.03.F1', 'XI6', ''),
    ('Q-004', 'C1.02.H2', 'XB7', ''),
    ('Q-008', 'A1.01.F1', 'XB7', ''),
    ('Q-009', 'A1.03.J2', 'XB7', ''),
    ('Q-011', 'A1.04.J2', 'XI6', ''),
    ('Q-012', 'A1.02.H2', 'XB7', ''),
    ('R-001', 'C2.06.C1', 'XI6', 'IPTVTCXI6HD'),
    ('R-002', 'A2.07.F2', 'XB7', 'TG4482A '),
    ('R-003', 'C1.02.G1', 'XIONE', 'SCXI11BEI'),
    ('R-006', 'C1.01.H1', 'XIONE', 'SCXI11BEI'),
    ('R-007', 'C1.02.G2', 'XB7', 'Mix '),
    ('R-008', 'A2.04.G1', 'XB7', 'TG4482A'),
    ('R-009', 'C1.03.G1', 'XIONE', 'SCXI11BEI'),
    ('R-010', 'C2.07.D1', 'XI6', 'IPTVTCXI6HD'),
    ('R-011', 'C2.07.D2', 'XI6', 'IPTVTCXI6HD'),
    ('R-012', 'A2.04.G2', 'XB8', ''),
    ('R-024', 'C1.02.J1', 'XIONE', 'SCXI11BEI'),
    ('R-039', 'A2.03.J1', 'XIONE', ''),
    ('REPAIR 001', 'C1.03.H1', 'XB7', 'CGM4331COM'),
    ('REPAIR 002', 'C1.05.F2', 'XB7', 'CGM4331COM'),
    ('REPAIR 003', 'C1.03.B2', 'TG4482A', ''),
    ('REPAIR 004', 'C1.04.B1', 'XI6', 'IPTVTCXI6HD'),
    ('REPAIR 005', 'C1.05.G2', 'XB7', 'CGM4331COM'),
    ('REPAIR 007', 'C1-02-H1', 'XB7', 'CGM4331COM'),
    ('REPAIR-006', 'C1.05.F1', 'XB7', 'TG4482A'),
    ('T-001', 'A1.04.F2', 'Coda4582', ''),
    ('T-001', 'C1.01.G1', 'XB8', ''),
    ('T-002', 'C1.03.G2', 'XB8', ''),
    ('T-002', 'C2.03.C1', 'XB6', ''),
    ('T-003', 'A2.01.H2', 'XB8', ''),
    ('T-003', 'A2.02.G2', 'XB6', ''),
    ('T-004', 'A2.01.G2', 'XB8', ''),
    ('T-004', 'A2.03.F2', 'CODA4582', 'NO ACCOUNTS'),
    ('T-005', 'A1.02.G1', 'XB8', ''),
    ('T-005', 'A2.01.G1', 'CODA4582', ''),
    ('T-006', 'A1.02.H1', 'CODA4582', 'NO ACCOUNTS'),
    ('T-007', 'A1.01.G2', 'ARRIS602G', 'NO ACCOUNTS'),
    ('T-008', 'A1.03.G1', 'DCX700', ''),
    ('T-009', 'A2.01.F1', 'CODA4582', 'NO ACCOUNTS'),
    ('T-010', 'A1.01.F2', 'Arris602G', 'ONE LEFT'),
    ('T-011', 'A1.02.F2', 'CODA4582', 'NO ACCOUNTS'),
    ('T-012', 'A1.05.H1', 'CODA4582', 'NO ACCOUNTS'),
    ('T-012', 'C2.05.B1', 'XB6', ''),
    ('T-013', 'A1.04.G2', 'TM804G', 'NO ACCOUNTS'),
    ('T-014', 'C2.07.C1', 'CODA4582', 'NO ACCOUNTS'),
    ('T-015', 'C2.01.B1', 'CODA4582', 'NO ACCOUNTS'),
    ('T-016', 'A1.07.G2', 'XE1V2WNROG1', ''),
    ('T-017', 'A1.05.H2', 'CODA4582', 'NO ACCOUNTS'),
    ('T-020', 'A1.04.B2', 'AER1600LP6', 'ONE LEFT'),
    ('T-022', 'A1.05.F1', 'Arris602G', '2  LEFT'),
    ('T-023', 'A1.05.G2', 'DCX3510', 'NO ACCOUNTS'),
    ('T-024', 'A1.06.F2', 'XB6', '')
]

# Initialize SQLite database
sections = ["A1", "A2", "B1"]
rows = [f"{i:02}" for i in range(1, 9)]  # 01 to 08
columns = ["F", "J", "H", "E", "D", "G"]
levels = ["1", "2"]  # Only two levels

def generate_rack_positions():
    return [f"{section}.{row}.{column}{level}"
            for section in sections
            for row in rows
            for column in columns
            for level in levels]

# Initialize SQLite database
def init_db():
    conn = sqlite3.connect("inventory.db")
    cursor = conn.cursor()
    
    # Create inventory table
    cursor.execute('''CREATE TABLE IF NOT EXISTS inventory (
                        location TEXT,
                        shelf TEXT,
                        device TEXT,
                        notes TEXT
                      )''')
    cursor.execute("DELETE FROM inventory")  # Clear existing data
    cursor.executemany("INSERT INTO inventory (location, shelf, device, notes) VALUES (?, ?, ?, ?)", data)
    
    # Create racks table
    cursor.execute('''CREATE TABLE IF NOT EXISTS racks (
                        shelf TEXT PRIMARY KEY
                      )''')
    cursor.execute("DELETE FROM racks")  # Clear existing data
    racks = [(rack,) for rack in generate_rack_positions()]
    cursor.executemany("INSERT INTO racks (shelf) VALUES (?)", racks)
    
    conn.commit()
    conn.close()

# Query empty racks
def find_empty_racks():
    conn = sqlite3.connect("inventory.db")
    cursor = conn.cursor()
    
    query = '''
        SELECT racks.shelf
        FROM racks
        LEFT JOIN inventory ON racks.shelf = inventory.shelf
        WHERE inventory.shelf IS NULL
    '''
    cursor.execute(query)
    empty_racks = cursor.fetchall()
    conn.close()
    return empty_racks

# Query data
def query_data(criteria, value):
    conn = sqlite3.connect("inventory.db")
    cursor = conn.cursor()
    
    query = f"SELECT * FROM inventory WHERE {criteria} LIKE ?"
    cursor.execute(query, (f"%{value}%",))
    
    results = cursor.fetchall()
    conn.close()
    return results

# Display results
def display_results(results, headers):
    if not results:
        print("No matching results found.")
        return
    
    header_line = " | ".join(headers)
    print(header_line)
    print("-" * len(header_line))
    for row in results:
        print(" | ".join(str(cell) for cell in row))

# Main program
def main():
    init_db()
    
    print("Inventory Query Tool")
    print("Options: 1. Query by Shelf (e.g., A1, B2)")
    print("         2. Query by Device (e.g., XB6)")
    print("         3. Query by Location (e.g., T, Q)")
    print("         4. View Empty Racks")
    
    option = input("Enter your choice (1/2/3/4): ").strip()
    
    if option == "1":
        value = input("Enter shelf value: ").strip()
        results = query_data("shelf", value)
        display_results(results, ["Location", "Shelf", "Device", "Notes"])
    elif option == "2":
        value = input("Enter device value: ").strip()
        results = query_data("device", value)
        display_results(results, ["Location", "Shelf", "Device", "Notes"])
    elif option == "3":
        value = input("Enter location value: ").strip()
        results = query_data("location", value)
        display_results(results, ["Location", "Shelf", "Device", "Notes"])
    elif option == "4":
        empty_racks = find_empty_racks()
        display_results(empty_racks, ["Empty Shelf"])
    else:
        print("Invalid option selected.")

# Run the program
if __name__ == "__main__":
    main()