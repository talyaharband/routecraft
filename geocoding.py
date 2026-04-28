import pandas as pd
import requests
import re
import time
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path


def load_env_file(env_path=".env"):
    """Load simple KEY=VALUE pairs from a local .env file."""
    env_file = Path(env_path)
    if not env_file.exists():
        return

    for raw_line in env_file.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


def get_api_key():
    load_env_file()
    return os.getenv("GOOGLE_MAPS_API_KEY") or os.getenv("GEOCODING_API_KEY")


def split_street_and_number(address_str):
    """
    מפרק כתובת בצורה חכמה: מוצא את הספרה הראשונה וחותך לפיה.
    מטפל גם ב'שבט יהודה 3' וגם ב'3 שבט יהודה'.
    """
    address_str = str(address_str).strip()
    if not address_str:
        return "", ""

    # חיפוש המיקום של הספרה הראשונה במחרוזת
    match = re.search(r'\d', address_str)

    if not match:
        # אם אין מספר בכלל, הכל זה רחוב
        return address_str, ""

    first_digit_idx = match.start()

    # אם הכתובת מתחילה במספר (למשל: "3 שבט יהודה")
    if first_digit_idx == 0:
        # מחפש איפה נגמרים המספרים ומתחילות האותיות
        end_of_num_match = re.search(r'^\d+[\s\-/]*', address_str)
        num_part = end_of_num_match.group(0).strip()
        street_part = address_str[len(num_part):].strip()
        return street_part, num_part

    # אם הכתובת מתחילה ברחוב (למשל: "שבט יהודה 3")
    street_part = address_str[:first_digit_idx].strip()
    num_part = address_str[first_digit_idx:].strip()

    # ניקוי פסיקים מיותרים שנשארים לפעמים בסוף שם הרחוב
    street_part = street_part.rstrip(',')

    return street_part, num_part


def clean_cell(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def get_lat_lng(address, key):
    """פונה ל-Google Geocoding API"""
    url = "https://maps.googleapis.com/maps/api/geocode/json"
    params = {'address': address, 'key': key}
    try:
        response = requests.get(url, params=params).json()
        if response['status'] == 'OK':
            location = response['results'][0]['geometry']['location']
            return location['lat'], location['lng']
    except Exception as e:
        print(f"שגיאה בכתובת {address}: {e}")
    return None, None


def run_geocoding_process():
    api_key = get_api_key()
    if not api_key:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Missing API key",
            "Add GOOGLE_MAPS_API_KEY=your_key to the local .env file.",
        )
        return

    # 1. פתיחת חלון לבחירת קובץ
    root = tk.Tk()
    root.withdraw()  # הסתרת החלון הראשי של tkinter

    print("אנא בחרי את קובץ האקסל המקורי...")
    input_path = filedialog.askopenfilename(
        title="בחירת קובץ אקסל",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    if not input_path:
        print("לא נבחר קובץ. היציאה מהתוכנית.")
        return

    # 2. טעינת הנתונים
    df = pd.read_excel(input_path)

    if df.shape[1] < 3:
        messagebox.showerror("שגיאה", "הקובץ חייב לכלול לפחות 3 עמודות: עיר, רחוב, מספר בית.")
        return

    city_col = df.columns[0]
    street_col = df.columns[1]
    house_number_col = df.columns[2]

    results = []
    print(f"מתחיל עיבוד של {len(df)} כתובות...")

    for index, row in df.iterrows():
        city = clean_cell(row[city_col])
        street_name = clean_cell(row[street_col])
        house_number = clean_cell(row[house_number_col])

        # 4. גיאוקודינג
        full_query = f"{street_name} {house_number}, {city}"
        lat, lng = get_lat_lng(full_query, api_key)

        results.append({
            'City': city,
            'Street_Name': street_name,
            'House_Number': house_number,
            'LAT': lat,
            'LNG': lng
        })

        print(f"[{index + 1}/{len(df)}] מעבד: {full_query}...")
        time.sleep(0.05)  # הגנה על ה-API

    # 5. שמירת התוצאה - בחירת מיקום לשמירה
    output_df = pd.DataFrame(results)
    save_path = filedialog.asksaveasfilename(
        title="שמירת הקובץ המעובד",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )

    if save_path:
        output_df.to_excel(save_path, index=False)
        print(f"✅ הקובץ נשמר בהצלחה בנתיב: {save_path}")
        messagebox.showinfo("סיום", f"העיבוד הסתיים! הקובץ נשמר ב:\n{save_path}")
    else:
        print("השמירה בוטלה על ידי המשתמש.")


if __name__ == "__main__":
    run_geocoding_process()
