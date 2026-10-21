import datetime
import os
import csv
import pandas as pd
import matplotlib.pyplot as plt

# ==================== TOLL RATES ====================
TOLL_RATES = {
    "Bike": 20,
    "Car": 50,
    "Truck": 100,
    "Bus": 80,
    "Van": 60
}

# File to store vehicle entries
DATA_FILE = "toll_entries.csv"

def load_entries():
    """Load vehicle entries from CSV file"""
    entries = []
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, mode='r', newline='') as file:
            reader = csv.DictReader(file)
            for row in reader:
                row['toll'] = float(row['toll'])
                entries.append(row)
    return entries

def save_entry(entry):
    """Save a single entry to CSV file"""
    file_exists = os.path.exists(DATA_FILE)
    with open(DATA_FILE, mode='a', newline='') as file:
        writer = csv.DictWriter(file, fieldnames=['receipt_no', 'vehicle_no', 'vehicle_type', 'toll', 'timestamp'])
        if not file_exists:
            writer.writeheader()
        writer.writerow(entry)

def generate_receipt_number():
    """Generate unique receipt number"""
    now = datetime.datetime.now()
    date_str = now.strftime("%Y%m%d")
    return f"TOLL-{date_str}-{len(load_entries()) + 1:04d}"

def register_vehicle():
    """Register vehicle details and collect toll"""
    print("\n=== Register Vehicle Entry ===")
    vehicle_no = input("Enter Vehicle Number (e.g., KA01AB1234): ").strip().upper()
    
    print("\nVehicle Types:")
    for i, vtype in enumerate(TOLL_RATES.keys(), 1):
        print(f"{i}. {vtype} - ₹{TOLL_RATES[vtype]}")
    
    try:
        choice = int(input("\nSelect vehicle type (1-5): "))
        if choice < 1 or choice > len(TOLL_RATES):
            print("Invalid choice!")
            return
        vehicle_type = list(TOLL_RATES.keys())[choice-1]
    except ValueError:
        print("Invalid input!")
        return
    
    toll = TOLL_RATES[vehicle_type]
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    receipt_no = generate_receipt_number()
    
    entry = {
        'receipt_no': receipt_no,
        'vehicle_no': vehicle_no,
        'vehicle_type': vehicle_type,
        'toll': toll,
        'timestamp': timestamp
    }
    
    save_entry(entry)
    print(f"\n✅ Vehicle registered successfully!")
    generate_receipt(entry)

def generate_receipt(entry):
    """Generate payment receipt"""
    print("\n" + "="*55)
    print("          TOLL GATE RECEIPT")
    print("="*55)
    print(f"Receipt No   : {entry['receipt_no']}")
    print(f"Date/Time    : {entry['timestamp']}")
    print(f"Vehicle No   : {entry['vehicle_no']}")
    print(f"Vehicle Type : {entry['vehicle_type']}")
    print(f"Toll Amount  : ₹{entry['toll']:.2f}")
    print("-"*55)
    print("Thank you for using our toll service!")
    print("="*55)

def display_daily_report():
    """Display daily revenue report"""
    entries = load_entries()
    if not entries:
        print("\n=== DAILY REVENUE REPORT ===")
        print("No entries found yet.")
        return

    today = datetime.date.today()
    daily_entries = []
    total_revenue = 0.0

    for entry in entries:
        try:
            ts = entry['timestamp'].strip()
            if '/' in ts:
                dt = datetime.datetime.strptime(ts, "%m/%d/%Y %H:%M")
            else:
                dt = datetime.datetime.strptime(ts, "%Y-%m-%d %H:%M:%S")
            if dt.date() == today:
                daily_entries.append(entry)
                total_revenue += float(entry['toll'])
        except:
            continue

    print("\n" + "="*85)
    print("                    DAILY REVENUE REPORT")
    print("="*85)
    print(f"Date              : {today}")
    print(f"Total Vehicles    : {len(daily_entries)}")
    print(f"Total Revenue     : ₹{total_revenue:.2f}")
    print("="*85)

    if daily_entries:
        print(f"{'Time':<12} {'Receipt No':<20} {'Vehicle No':<15} {'Type':<8} {'Amount':>8}")
        print("-" * 85)
        for entry in daily_entries:
            ts = entry['timestamp'].strip()
            time_only = ts.split()[1][:8] if ' ' in ts else "N/A"
            print(f"{time_only:<12} {entry['receipt_no']:<20} {entry['vehicle_no']:<15} "
                  f"{entry['vehicle_type']:<8} ₹{float(entry['toll']):>7.2f}")
    else:
        print("No vehicles registered today yet.")
    print("="*85)

def view_all_history():
    """View complete history with summary"""
    entries = load_entries()
    print("\n" + "="*90)
    print("           COMPLETE VEHICLE ENTRY HISTORY")
    print("="*90)
    
    if not entries:
        print("No entries found.")
        return
    
    total_vehicles = len(entries)
    total_revenue = 0.0
    revenue_by_type = {"Bike": 0.0, "Car": 0.0, "Truck": 0.0, "Bus": 0.0, "Van": 0.0}
    
    for entry in entries:
        toll = float(entry['toll'])
        total_revenue += toll
        vtype = entry['vehicle_type']
        if vtype in revenue_by_type:
            revenue_by_type[vtype] += toll
    
    print(f"Total Registrations     : {total_vehicles}")
    print(f"Overall Total Revenue   : ₹{total_revenue:.2f}")
    print("-"*90)
    print("Revenue Breakdown:")
    for vtype in TOLL_RATES.keys():
        print(f"   {vtype:<8} : ₹{revenue_by_type[vtype]:.2f}")
    print("-"*90)
    
    print("Detailed Records:")
    print("-"*90)
    for entry in entries:
        print(f"{entry['timestamp']} | {entry['receipt_no']} | {entry['vehicle_no']:<12} | "
              f"{entry['vehicle_type']:<8} | ₹{entry['toll']:.2f}")
    print("="*90)

def export_to_excel():
    """Export to Excel and open automatically"""
    entries = load_entries()
    if not entries:
        print("No records to export!")
        return
    
    df = pd.DataFrame(entries)
    df = df[['receipt_no', 'timestamp', 'vehicle_no', 'vehicle_type', 'toll']]
    df.columns = ['Receipt No', 'Date & Time', 'Vehicle Number', 'Vehicle Type', 'Toll Amount (₹)']
    
    filename = f"Toll_Records_{datetime.date.today()}.xlsx"
    df.to_excel(filename, index=False)
    
    print(f"\n✅ Successfully exported to Excel!")
    print(f"File saved as: {filename}")
    
    try:
        os.startfile(filename)
        print("📂 Excel file opened automatically!")
    except:
        print("Please open the file manually.")

def generate_revenue_graph():
    """Generate Bar Graph"""
    entries = load_entries()
    if not entries:
        print("No records to generate graph!")
        return
    
    revenue_by_type = {"Bike": 0.0, "Car": 0.0, "Truck": 0.0, "Bus": 0.0, "Van": 0.0}
    for entry in entries:
        vtype = entry['vehicle_type']
        if vtype in revenue_by_type:
            revenue_by_type[vtype] += float(entry['toll'])
    
    types = list(revenue_by_type.keys())
    revenues = list(revenue_by_type.values())
    
    plt.figure(figsize=(10, 6))
    bars = plt.bar(types, revenues, color=['blue', 'green', 'red', 'purple', 'orange'])
    plt.title('Toll Revenue by Vehicle Type', fontsize=14, fontweight='bold')
    plt.xlabel('Vehicle Type')
    plt.ylabel('Revenue (₹)')
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    
    for bar in bars:
        height = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2., height + 5, f'₹{height:.0f}', ha='center')
    
    plt.tight_layout()
    plt.show()
    print("\n📊 Bar Graph displayed successfully!")

def generate_revenue_pie_chart():
    """Generate Pie Chart"""
    entries = load_entries()
    if not entries:
        print("No records to generate pie chart!")
        return
    
    revenue_by_type = {"Bike": 0.0, "Car": 0.0, "Truck": 0.0, "Bus": 0.0, "Van": 0.0}
    for entry in entries:
        vtype = entry['vehicle_type']
        if vtype in revenue_by_type:
            revenue_by_type[vtype] += float(entry['toll'])
    
    types = list(revenue_by_type.keys())
    revenues = list(revenue_by_type.values())
    total = sum(revenues)
    
    plt.figure(figsize=(10, 7))
    plt.pie(revenues, labels=types, autopct='%1.1f%%', 
            colors=['blue', 'green', 'red', 'purple', 'orange'], startangle=90, shadow=True)
    plt.title(f'Revenue Distribution by Vehicle Type\n(Total: ₹{total:.2f})', fontsize=14, fontweight='bold')
    plt.axis('equal')
    plt.show()
    print("\n📊 Pie Chart displayed successfully!")

# ==================== MAIN MENU ====================
def main_menu():
    while True:
        print("\n" + "="*70)
        print("TOLL GATE COLLECTION MANAGEMENT SYSTEM")
        print("="*70)
        print("1. Register New Vehicle Entry")
        print("2. Generate Daily Revenue Report")
        print("3. View Complete History")
        print("4. Export Records to Excel")
        print("5. Generate Revenue Bar Graph")
        print("6. Generate Revenue Pie Chart")
        print("7. Exit")
        print("="*70)
        
        choice = input("Enter your choice (1-7): ").strip()
        
        if choice == '1':
            register_vehicle()
        elif choice == '2':
            display_daily_report()
        elif choice == '3':
            view_all_history()
        elif choice == '4':
            export_to_excel()
        elif choice == '5':
            generate_revenue_graph()
        elif choice == '6':
            generate_revenue_pie_chart()
        elif choice == '7':
            print("\nThank you for using Toll Gate Management System!")
            break
        else:
            print("Invalid choice! Please try again.")

# Run the program
if __name__ == "__main__":
    print("🚦 Starting Toll Gate Management System...")
    main_menu()
