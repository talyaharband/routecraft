import pandas as pd
import requests
import re
import time
import tkinter as tk
from tkinter import filedialog, messagebox

# --- הגדרות ---
API_KEY = "my_key"  # הזיני את המפתח שלך כאן


# --- פונקציות עזר ---

def split_street_and_number(address_str):
    """מפרק כתובת בצורה חכמה"""
    address_str = str(address_str).strip()
    if not address_str: return "", ""
    match = re.search(r'\d', address_str)
    if not match: return address_str, ""
    first_digit_idx = match.start()
    if first_digit_idx == 0:
        end_of_num_match = re.search(r'^\d+[\s\-/]*', address_str)
        num_part = end_of_num_match.group(0).strip()
        street_part = address_str[len(num_part):].strip()
        return street_part, num_part
    street_part = address_str[:first_digit_idx].strip().rstrip(',')
    num_part = address_str[first_digit_idx:].strip()
    return street_part, num_part


def get_lat_lng(address, key):
    """פונה ל-Google Geocoding API ומחזיר נ"צ"""
    url = "https://maps.googleapis.com/maps/api/geocode/json"
    params = {'address': address, 'key': key}
    try:
        response = requests.get(url, params=params).json()
        if response['status'] == 'OK':
            location = response['results'][0]['geometry']['location']
            return location['lat'], location['lng']
    except Exception as e:
        print(f"שגיאה בגיאוקודינג {address}: {e}")
    return None, None


# --- פונקציית האיחוד המרכזית ---

def run_combined_process():
    # 1. בחירת קובץ
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    print('--- שלב 1: טעינת קובץ כתובות מקורי ---')
    input_path = filedialog.askopenfilename(title="בחירת קובץ אקסל (עיר וכתובת)")
    if not input_path: return

    df = pd.read_excel(input_path)
    city_col = df.columns[0]
    address_col = df.columns[1]

    # 2. גיאוקודינג - המרה לקואורדינטות פעם אחת בלבד
    print(f'מבצע גיאוקודינג עבור {len(df)} כתובות...')
    locations_data = []

    for index, row in df.iterrows():
        city = str(row[city_col])
        full_addr_str = str(row[address_col])
        street, house_num = split_street_and_number(full_addr_str)

        query = f"{street} {house_num}, {city}"
        lat, lng = get_lat_lng(query, API_KEY)

        locations_data.append({
            'Label': f"{street} {house_num}",
            'Lat': lat,
            'Lng': lng,
            'Full_Query': query
        })
        # שימוש במירכאות בודדות כדי למנוע שגיאה עם ה-נ"צ
        print(f'[{index + 1}/{len(df)}] נמצא נ"צ עבור: {query}')
        time.sleep(0.02)

    # 3. בניית מטריצת המרחקים (בדקות) תוך שימוש בנ"צ שמצאנו
    print('\n--- שלב 2: מחשב מטריצת זמנים (שימוש בנ"צ קיים) ---')
    matrix_data = []
    labels = [loc['Label'] for loc in locations_data]

    for i, origin in enumerate(locations_data):
        row_minutes = []
        for j, dest in enumerate(locations_data):
            if i == j:
                row_minutes.append(0)
                continue

            if origin['Lat'] is None or dest['Lat'] is None:
                row_minutes.append(999)
                continue

            dm_params = {
                'origins': f"{origin['Lat']},{origin['Lng']}",
                'destinations': f"{dest['Lat']},{dest['Lng']}",
                'mode': 'driving',
                'departure_time': 'now',
                'key': API_KEY
            }

            try:
                res = requests.get("https://maps.googleapis.com/maps/api/distancematrix/json", params=dm_params).json()
                if res['status'] == 'OK':
                    el = res['rows'][0]['elements'][0]
                    if el['status'] == 'OK':
                        seconds = el.get('duration_in_traffic', el['duration'])['value']
                        row_minutes.append(round(seconds / 60))
                    else:
                        row_minutes.append(999)
                else:
                    row_minutes.append(999)
            except:
                row_minutes.append(999)

            time.sleep(0.01)

        print(f"✅ הושלמה שורה {i + 1}/{len(locations_data)} ({origin['Label']})")
        matrix_data.append(row_minutes)

    # 4. שמירת התוצאות
    df_matrix = pd.DataFrame(matrix_data, index=labels, columns=labels)

    save_path = filedialog.asksaveasfilename(
        title="שמירת מטריצת הדקות",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )

    if save_path:
        df_matrix.to_excel(save_path)
        print('\n--- המטריצה מוכנה ונשמרה בדקות שלמות! ---')
        messagebox.showinfo("סיום", "התהליך המאוחד הסתיים בהצלחה!")


if __name__ == "__main__":
    run_combined_process()