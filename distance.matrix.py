import re
import time
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd
import requests

# Paste your Google API key here.
API_KEY = "my_api_key"


def clean_address(addr):
    """Strip apartment suffixes like 12/3 to improve geocoding."""
    return re.sub(r"(\d+)/\d+", r"\1", addr)


def get_coords(addr, key):
    """Resolve an address to coordinates, with a simplified-address fallback."""
    params = {"address": addr, "key": key}
    try:
        res = requests.get(
            "https://maps.googleapis.com/maps/api/geocode/json",
            params=params
        ).json()
        if res["status"] == "OK":
            return res["results"][0]["geometry"]["location"], "exact"

        cleaned = clean_address(addr)
        if cleaned != addr:
            params["address"] = cleaned
            res = requests.get(
                "https://maps.googleapis.com/maps/api/geocode/json",
                params=params
            ).json()
            if res["status"] == "OK":
                return res["results"][0]["geometry"]["location"], "building"
    except Exception as exc:
        print(f"Geocoding failed for '{addr}': {exc}")

    return None, "not found"


def build_full_addresses(df):
    if df.shape[1] >= 2:
        addresses = (df.iloc[:, 1].astype(str) + ", " + df.iloc[:, 0].astype(str)).tolist()
    else:
        addresses = [str(addr) + ", Lod" for addr in df.iloc[:, 0].tolist()]

    return [addr.replace("nan", "").strip(", ") for addr in addresses]


def sanitize_group_value(group_value):
    safe_value = str(group_value).strip()
    safe_value = re.sub(r'[<>:"/\\|?*]+', "_", safe_value)
    safe_value = re.sub(r"\s+", "_", safe_value)
    return safe_value or "unknown"


def build_matrix_for_addresses(full_addresses, coords_dict):
    matrix_data = []

    for i, origin in enumerate(full_addresses):
        row_minutes = []
        for j, dest in enumerate(full_addresses):
            if i == j:
                row_minutes.append(0)
                continue

            if not coords_dict[origin] or not coords_dict[dest]:
                row_minutes.append(999)
                continue

            dm_params = {
                "origins": f"{coords_dict[origin]['lat']},{coords_dict[origin]['lng']}",
                "destinations": f"{coords_dict[dest]['lat']},{coords_dict[dest]['lng']}",
                "mode": "driving",
                "departure_time": "now",
                "key": API_KEY,
            }

            try:
                dm_res = requests.get(
                    "https://maps.googleapis.com/maps/api/distancematrix/json",
                    params=dm_params
                ).json()
                if dm_res["status"] == "OK":
                    element = dm_res["rows"][0]["elements"][0]
                    if element["status"] == "OK":
                        seconds = element.get("duration_in_traffic", element["duration"])["value"]
                        row_minutes.append(round(seconds / 60))
                    else:
                        row_minutes.append(999)
                else:
                    row_minutes.append(999)
            except Exception:
                row_minutes.append(999)

            time.sleep(0.02)

        print(f"Completed row {i + 1}/{len(full_addresses)}")
        matrix_data.append(row_minutes)

    return pd.DataFrame(matrix_data, index=full_addresses, columns=full_addresses)


def run_distance_matrix_minutes():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    print("--- Step 1: Select input Excel file ---")
    input_path = filedialog.askopenfilename(title="Select input Excel file")
    if not input_path:
        return

    df = pd.read_excel(input_path)
    df.columns = df.columns.str.strip()

    if df.shape[1] < 9:
        messagebox.showerror(
            "Error",
            "The input file must include column I (the 9th column) with the group number."
        )
        return

    group_col = df.columns[8]
    grouped_frames = {
        group_value: group_df.reset_index(drop=True)
        for group_value, group_df in df.groupby(group_col, dropna=True)
    }

    if not grouped_frames:
        messagebox.showerror("Error", "No groups were found in column I.")
        return

    output_dir = Path(input_path).parent
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    all_addresses = []
    for group_df in grouped_frames.values():
        all_addresses.extend(build_full_addresses(group_df))
    unique_addresses = list(dict.fromkeys(all_addresses))

    print(f"Geocoding {len(unique_addresses)} unique addresses...")
    coords_dict = {}
    for addr in unique_addresses:
        coords, status = get_coords(addr, API_KEY)
        coords_dict[addr] = coords
        if status != "exact":
            print(f"Address '{addr}' resolved with status: {status}")

    print("\n--- Step 2: Build one matrix per group ---")
    saved_files = []

    for group_value, group_df in grouped_frames.items():
        full_addresses = build_full_addresses(group_df)
        print(f"\nProcessing group {group_value} with {len(full_addresses)} rows...")

        df_result = build_matrix_for_addresses(full_addresses, coords_dict)
        filename = f"matrix_{sanitize_group_value(group_value)}_{timestamp}.xlsx"
        save_path = output_dir / filename
        df_result.to_excel(save_path)
        saved_files.append(str(save_path))

        print(f"Saved: {save_path}")

    messagebox.showinfo(
        "Done",
        "Created one matrix file per group.\n\n"
        f"Files created: {len(saved_files)}\n"
        f"Output folder: {output_dir}"
    )


if __name__ == "__main__":
    run_distance_matrix_minutes()
