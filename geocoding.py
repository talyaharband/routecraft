import pandas as pd
import requests
import re
import time
import tkinter as tk
from tkinter import filedialog, messagebox

# --- הגדרות ---
API_KEY = "API KEY"


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

    # וידוי שיש לפחות 2 עמודות
    if df.shape[1] < 2:
        messagebox.showerror("שגיאה", "בקובץ חייבות להיות לפחות 2 עמודות (עיר ורחוב+מספר)")
        return

    city_col = df.columns[0]
    address_col = df.columns[1]

    results = []
    print(f"מתחיל עיבוד של {len(df)} כתובות...")

    for index, row in df.iterrows():
        city = str(row[city_col])
        full_address_str = str(row[address_col])

        # 3. פירוק הכתובת
        street_name, house_number = split_street_and_number(full_address_str)

        # 4. גיאוקודינג
        full_query = f"{street_name} {house_number}, {city}"
        lat, lng = get_lat_lng(full_query, API_KEY)

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