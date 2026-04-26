import pandas as pd
import math
import tkinter as tk
from tkinter import filedialog, messagebox


def calculate_haversine(lat1, lon1, lat2, lon2):
    """מחשב מרחק אווירי במטרים"""
    if pd.isna(lat1) or pd.isna(lon1) or pd.isna(lat2) or pd.isna(lon2):
        return float('inf')
    R = 6371000
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlambda / 2) ** 2
    return 2 * R * math.atan2(math.sqrt(a), math.sqrt(1 - a))


def run_smart_clustering():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    print("--- שלב 1: בחירת קובץ ---")
    input_path = filedialog.askopenfilename(title="בחירת קובץ קואורדינטות")
    if not input_path: return

    df = pd.read_excel(input_path)

    # --- שלב 2: תיקון המיון והצגת המספרים ---
    # הופך למספר לצורך המיון
    df['House_Number'] = pd.to_numeric(df['House_Number'], errors='coerce')

    # מסיר שורות ללא מספר בית (אם יש כאלו) כדי שלא יתקעו את ההפיכה למספר שלם
    df = df.dropna(subset=['House_Number']).reset_index(drop=True)

    # הופך למספר שלם (כדי להוריד את ה- .0)
    df['House_Number'] = df['House_Number'].astype(int)

    # מיון סופי ונקי
    df = df.sort_values(by=['Street_Name', 'House_Number']).reset_index(drop=True)

    all_clusters = []
    dist_threshold = 150

    anchor_address = df.iloc[0]
    current_cluster = [anchor_address.to_dict()]

    print(f"\n--- שלב 2: סריקה ואיחוד מול עוגן (סף: {dist_threshold} מטר) ---")

    for i in range(1, len(df)):
        curr = df.iloc[i]
        dist_from_anchor = calculate_haversine(
            anchor_address['LAT'], anchor_address['LNG'],
            curr['LAT'], curr['LNG']
        )

        if curr['Street_Name'] == anchor_address['Street_Name'] and dist_from_anchor <= dist_threshold:
            current_cluster.append(curr.to_dict())
        else:
            all_clusters.append(current_cluster)
            anchor_address = curr
            current_cluster = [anchor_address.to_dict()]

    all_clusters.append(current_cluster)

    # 3. יצירת רשימת הנציגים המפורטת
    reps_data = []
    for i, cluster in enumerate(all_clusters):
        rep = cluster[0].copy()
        rep['cluster_id'] = i
        rep['total_orders_in_cluster'] = len(cluster)

        # יצירת מחרוזת כתובות ללא נקודה עשרונית
        address_strings = [f"{item['Street_Name']} {int(item['House_Number'])}" for item in cluster]
        rep['detailed_addresses'] = ", ".join(address_strings)

        reps_data.append(rep)

    df_reps = pd.DataFrame(reps_data)

    save_path = filedialog.asksaveasfilename(
        title="שמירת קובץ הנציגים הסופי",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )

    if save_path:
        # כאן אנחנו מוודאים שגם בתוך האקסל העמודה תישמר כמספר שלם
        df_reps.to_excel(save_path, index=False)
        print(f"✅ הקובץ נשמר בפורמט נקי.")
        messagebox.showinfo("סיום", "הקובץ מוכן! מספרים שלמים, ללא נקודות עשרוניות.")


if __name__ == "__main__":
    run_smart_clustering()