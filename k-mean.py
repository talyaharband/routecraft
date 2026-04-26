import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from k_means_constrained import KMeansConstrained
import tkinter as tk
from tkinter import filedialog, messagebox


def run_constrained_clustering():
    # 1. בחירת הקובץ
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    print("אנא בחרי את קובץ האקסל המעודכן...")
    input_path = filedialog.askopenfilename(title="בחירת קובץ אקסל עם קואורדינטות")
    if not input_path: return

    # 2. טעינת הנתונים
    df = pd.read_excel(input_path)

    # עמודות 4 ו-5 (אינדקסים 3 ו-4 ב-Python)
    coords = df.iloc[:, [3, 4]].values

    # 3. בדיקה האם בכלל צריך חלוקה
    num_addresses = len(df)

    if num_addresses <= 30:
        # --- התיקון כאן: אם יש פחות מ-30, אין טעם להפעיל את האלגוריתם ---
        print(f"יש רק {num_addresses} כתובות. הן נכנסות לקבוצה אחת (0).")
        n_clusters = 1
        df['cluster_group'] = 0
    else:
        # אם יש יותר מ-30, מחשבים כמה קבוצות צריך ומפעילים את האלגוריתם
        n_clusters = int(np.ceil(num_addresses / 30))
        print(f"מבצע חלוקה ל-{n_clusters} קבוצות (מקסימום 30 כתובות לקבוצה)...")

        # 4. הרצת האלגוריתם
        clf = KMeansConstrained(
            n_clusters=n_clusters,
            size_min=1,
            size_max=30,
            random_state=42
        )
        df['cluster_group'] = clf.fit_predict(coords)

    # 5. ויזואליזציה - הצגת המפה
    plt.figure(figsize=(10, 7))
    scatter = plt.scatter(coords[:, 1], coords[:, 0], c=df['cluster_group'], cmap='viridis', s=50)

    plt.title(f'Delivery Clusters ({n_clusters} Groups)', fontsize=14)
    plt.xlabel('Longitude (LNG)')
    plt.ylabel('Latitude (LAT)')
    plt.colorbar(scatter, label='Group Number')
    plt.grid(True, linestyle='--', alpha=0.6)

    print("מציג את המפה... סגרי את חלון המפה כדי להמשיך לשמירת הקובץ.")
    plt.show()

    # 6. שמירת התוצאה
    save_path = filedialog.asksaveasfilename(
        title="שמירת קובץ החלוקה לקבוצות",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )

    if save_path:
        df.to_excel(save_path, index=False)
        messagebox.showinfo("סיום", f"הקובץ נשמר בהצלחה עם עמודת cluster_group!")
        print(f"✅ הקובץ נשמר ב: {save_path}")


if __name__ == "__main__":
    run_constrained_clustering()