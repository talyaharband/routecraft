import pandas as pd
import math
import tkinter as tk
from tkinter import filedialog
from bidi.algorithm import get_display

# Warehouse coordinates
WAREHOUSE_LAT = 31.77927525
WAREHOUSE_LNG = 35.0105885


def calculate_haversine(lat1, lon1, lat2, lon2):
    """Calculate air distance in meters between two coordinates"""
    if pd.isna(lat1) or pd.isna(lon1) or pd.isna(lat2) or pd.isna(lon2):
        return float('inf')
    R = 6371000  # Earth radius in meters
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlambda / 2) ** 2
    return 2 * R * math.atan2(math.sqrt(a), math.sqrt(1 - a))


def find_closest_to_warehouse():
    # Open file dialog
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    print("Opening file selector... Please select your Excel file with cluster data.")
    file_path = filedialog.askopenfilename(
        title="Select Excel File with Cluster Data",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    
    if not file_path:
        print("❌ No file selected. Exiting.")
        return

    # Load Excel file
    df = pd.read_excel(file_path)

    # Calculate distance from warehouse to each address
    df['distance_to_warehouse'] = df.apply(
        lambda row: calculate_haversine(
            WAREHOUSE_LAT, WAREHOUSE_LNG,
            row['LAT'], row['LNG']
        ),
        axis=1
    )

    # Find the closest address
    closest_idx = df['distance_to_warehouse'].idxmin()
    closest_address = df.loc[closest_idx]

    # Build address from Street_Name and House_Number
    street_name = closest_address['Street_Name']
    house_number = closest_address['House_Number']
    address = f"{street_name} {house_number}"
    
    # Display with RTL support for Hebrew
    address_display = get_display(address)
    cluster_group = closest_address['cluster_group']

    # Display results
    print("\n" + "=" * 60)
    print(f"Closest Address: {address_display}")
    print(f"Cluster Group: {cluster_group}")
    print("=" * 60)


if __name__ == "__main__":
    find_closest_to_warehouse()

