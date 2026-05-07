import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from k_means_constrained import KMeansConstrained
import tkinter as tk
from tkinter import filedialog, messagebox

try:
    from bidi.algorithm import get_display
except ModuleNotFoundError:
    get_display = lambda value: value


def run_constrained_clustering():
    # 1. בחירת הקובץ
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    print(get_display("אנא בחרי את קובץ האקסל המעודכן..."))
    input_path = filedialog.askopenfilename(title="בחירת קובץ אקסל עם קואורדינטות")
    if not input_path:
        return

    # 2. טעינת הנתונים
    df = pd.read_excel(input_path)

    city_col = "City" if "City" in df.columns else df.columns[0]
    lat_col = "LAT" if "LAT" in df.columns else df.columns[3]
    lng_col = "LNG" if "LNG" in df.columns else df.columns[4]

    # 3. חלוקה לפי עיר
    df["cluster_group"] = -1
    global_cluster_id = 0
    total_clusters = 0

    print(get_display("\n--- שלב 3: חלוקה לקבוצות לפי ערים ---"))
    for city, city_df in df.groupby(city_col):
        coords = city_df[[lat_col, lng_col]].values
        num_addresses = len(city_df)

        print(get_display(f"\n🏙️ עיר: {city} | כתובות: {num_addresses}"))

        if num_addresses <= 30:
            print(get_display(f"   -> נכנס לקבוצה אחת (קבוצה {global_cluster_id})"))
            df.loc[city_df.index, "cluster_group"] = global_cluster_id
            global_cluster_id += 1
            total_clusters += 1
        else:
            n_clusters = int(np.ceil(num_addresses / 30))
            print(get_display(f"   -> מחלק ל-{n_clusters} קבוצות (מקסימום 30 לקבוצה)..."))

            # הרצת האלגוריתם לעיר הספציפית
            clf = KMeansConstrained(
                n_clusters=n_clusters,
                size_min=1,
                size_max=30,
                n_init=50,
                max_iter=500,
                random_state=42,
            )
            local_labels = clf.fit_predict(coords)
            # הוספת המזהה הגלובלי כדי שהמספרים ימשיכו לעלות
            df.loc[city_df.index, "cluster_group"] = local_labels + global_cluster_id

            global_cluster_id += n_clusters
            total_clusters += n_clusters

    print(get_display(f"\n✅ סך הכל נוצרו {total_clusters} קבוצות בכל הארץ."))

    # 5. ויזואליזציה - הצגת מפה לכל עיר בנפרד
    print(get_display("מציג מפות לכל עיר בנפרד... סגרי כל חלון מפה כדי להציג את העיר הבאה או כדי להמשיך לשמירת הקובץ."))
    for city, city_df in df.groupby(city_col):
        city_coords = city_df[[lat_col, lng_col]].values
        city_clusters = city_df["cluster_group"]

        plt.figure(figsize=(10, 7))
        scatter = plt.scatter(city_coords[:, 1], city_coords[:, 0], c=city_clusters, cmap="viridis", s=50)

        num_groups_in_city = city_clusters.nunique()
        display_city = get_display(str(city))
        plt.title(f"Delivery Clusters - {display_city} ({num_groups_in_city} Groups)", fontsize=14)
        plt.xlabel("Longitude (LNG)")
        plt.ylabel("Latitude (LAT)")
        plt.colorbar(scatter, label="Global Group Number")
        plt.grid(True, linestyle="--", alpha=0.6)

        plt.show()

    # 6. שמירת התוצאה
    save_path = filedialog.asksaveasfilename(
        title="שמירת קובץ החלוקה לקבוצות",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
    )

    if save_path:
        df.to_excel(save_path, index=False)
        messagebox.showinfo("סיום", f"הקובץ נשמר בהצלחה עם עמודת cluster_group!\nנוצרו {total_clusters} קבוצות.")
        print(get_display(f"✅ הקובץ נשמר ב: {save_path}"))


if __name__ == "__main__":
    run_constrained_clustering()
