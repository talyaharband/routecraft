import pandas as pd
import requests
import time
import re
import tkinter as tk
from tkinter import filedialog, messagebox

# --- הגדרות ---
# הדביקי כאן את המפתח שלך
API_KEY = "API KEY"


def clean_address(addr):
    """מנקה לוכסנים ודירות כדי לעזור לגוגל למצוא את הבניין"""
    return re.sub(r'(\d+)/\d+', r'\1', addr)


def get_coords(addr, key):
    """מאתר קואורדינטות בצורה עקשנית - בדיוק כמו ב-Colab"""
    params = {'address': addr, 'key': key}
    try:
        res = requests.get("https://maps.googleapis.com/maps/api/geocode/json", params=params).json()
        if res['status'] == 'OK':
            return res['results'][0]['geometry']['location'], "מדויק"

        cleaned = clean_address(addr)
        if cleaned != addr:
            params['address'] = cleaned
            res = requests.get("https://maps.googleapis.com/maps/api/geocode/json", params=params).json()
            if res['status'] == 'OK':
                return res['results'][0]['geometry']['location'], "בניין"
    except Exception as e:
        print(f"שגיאה באיתור כתובת {addr}: {e}")

    return None, "לא נמצא"


def run_distance_matrix_minutes():
    # 1. בחירת הקובץ
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    print("--- שלב 1: בחירת קובץ אקסל ---")
    input_path = filedialog.askopenfilename(title="בחירת קובץ אקסל לנציגים")
    if not input_path: return

    df = pd.read_excel(input_path)
    df.columns = df.columns.str.strip()

    # יצירת רשימת כתובות
    if df.shape[1] >= 2:
        full_addresses = (df.iloc[:, 1].astype(str) + ", " + df.iloc[:, 0].astype(str)).tolist()
    else:
        full_addresses = [str(addr) + ", לוד" for addr in df.iloc[:, 0].tolist()]

    full_addresses = [addr.replace("nan", "").strip(", ") for addr in full_addresses]

    # --- 2. איתור מיקומים (Geocoding) ---
    print(f"מאתר {len(full_addresses)} מיקומים בגוגל...")
    coords_dict = {}
    for addr in full_addresses:
        coords, status = get_coords(addr, API_KEY)
        coords_dict[addr] = coords
        if status != "מדויק":
            print(f"📍 {addr} אותרה לפי: {status}")

    # --- 3. חישוב מטריצה בדקות (מספרים שלמים) ---
    print("\n--- שלב 2: מחשב מטריצת זמנים בדקות (זמן אמת) ---")
    matrix_data = []

    for i, origin in enumerate(full_addresses):
        row_minutes = []
        for j, dest in enumerate(full_addresses):
            # חוק האלכסון
            if i == j:
                row_minutes.append(0)
                continue

            if not coords_dict[origin] or not coords_dict[dest]:
                row_minutes.append(999)  # ערך גבוה לכתובת חסרה
                continue

            dm_params = {
                'origins': f"{coords_dict[origin]['lat']},{coords_dict[origin]['lng']}",
                'destinations': f"{coords_dict[dest]['lat']},{coords_dict[dest]['lng']}",
                'mode': 'driving',
                'departure_time': 'now',
                'key': API_KEY
            }

            try:
                dm_res = requests.get("https://maps.googleapis.com/maps/api/distancematrix/json",
                                      params=dm_params).json()
                if dm_res['status'] == 'OK':
                    el = dm_res['rows'][0]['elements'][0]
                    if el['status'] == 'OK':
                        # --- השינוי: הפיכה לשניות -> דקות -> עיגול לשלם ---
                        seconds = el.get('duration_in_traffic', el['duration'])['value']
                        minutes = round(seconds / 60)
                        row_minutes.append(minutes)
                    else:
                        row_minutes.append(999)
                else:
                    row_minutes.append(999)
            except:
                row_minutes.append(999)

            time.sleep(0.02)

        print(f"✅ הושלמה שורה {i + 1}/{len(full_addresses)}")
        matrix_data.append(row_minutes)

    # --- 4. שמירה ---
    df_result = pd.DataFrame(matrix_data, index=full_addresses, columns=full_addresses)

    save_path = filedialog.asksaveasfilename(
        title="שמירת מטריצת הדקות",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )

    if save_path:
        df_result.to_excel(save_path)
        print(f"\n✅ המטריצה נשמרה ב: {save_path}")
        messagebox.showinfo("סיום", "המטריצה מוכנה! הזמנים מוצגים בדקות שלמות.")


if __name__ == "__main__":
    run_distance_matrix_minutes()