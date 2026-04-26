import pandas as pd
from openai import OpenAI
import requests
import json
import time
import tkinter as tk
from tkinter import filedialog, messagebox


# --- חלונית מפתחות ---
class KeyRequester(tk.Toplevel):
    def __init__(self, title, prompt):
        super().__init__()
        self.title(title)
        self.geometry("450x200")
        self.attributes("-topmost", True)
        self.result = None
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - 225
        y = (self.winfo_screenheight() // 2) - 100
        self.geometry(f'+{x}+{y}')
        tk.Label(self, text=prompt, pady=15, font=("Arial", 10, "bold")).pack()
        self.entry = tk.Entry(self, width=50)
        self.entry.pack(pady=10, padx=20)
        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=15)
        tk.Button(btn_frame, text="📋 הדבק", command=lambda: self.entry.insert(0, self.clipboard_get())).pack(
            side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="אישור", command=self.on_submit, width=10, bg="#4CAF50", fg="white").pack(
            side=tk.LEFT, padx=10)
        self.wait_window()

    def on_submit(self):
        self.result = self.entry.get().strip()
        self.destroy()


# --- פונקציית העיבוד המרכזית ---
def process_address_clean(s1, s2, city, client, g_key):
    combined = f"{s1} {s2}".strip()

    # שלב 1: ה-AI מפרש את הכתובת לפי הבנתו הנקיה
    prompt = (
        f"Parse this Israeli address: '{combined}', City: '{city}'.\n"
        "Use your best judgment to extract: 'street', 'number', 'apt', 'floor', 'extra'.\n"
        "Fix typos and normalize street names (e.g. 'חזוניש' to 'חזון איש').\n"
        "Return ONLY JSON."
    )

    try:
        ai_res = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "system", "content": "You are a helpful Israeli address expert."},
                      {"role": "user", "content": prompt}],
            response_format={"type": "json_object"}
        )
        data = json.loads(ai_res.choices[0].message.content)

        street = data.get('street', '').strip()
        num = str(data.get('number', '')).strip()

        # שלב 2: אימות מול גוגל ובדיקת "כתובת אמת"
        google_status = "לא נבדק"
        final_street = street

        if street:
            url = f"https://maps.googleapis.com/maps/api/geocode/json?address={street} {num}, {city}&key={g_key}&language=he"
            res = requests.get(url).json()

            if res['status'] == 'OK':
                result = res['results'][0]
                # בדיקה האם זו כתובת מדויקת (בניין) או רק רחוב/עיר
                if 'street_address' in result['types'] or 'premise' in result['types']:
                    google_status = "✅ כתובת אומתה במלואה (קיימת במפה)"
                else:
                    # בודק אם לפחות הרחוב נמצא
                    has_route = any('route' in c['types'] for c in result['address_components'])
                    google_status = "⚠️ רחוב נמצא - מספר בית לא וודאי" if has_route else "❓ נמצאה רק העיר/שכונה"

                # עדכון שם הרחוב לשם הרשמי של גוגל
                for comp in result['address_components']:
                    if 'route' in comp['types']: final_street = comp['long_name']
            else:
                google_status = f"❌ כתובת לא נמצאה ({res['status']})"

        # שלב 3: הרכבת המחרוזת הסופית
        parts = [f"{final_street} {num}".strip()]
        if data.get('apt'): parts.append(f"דירה {data['apt']}")
        if data.get('floor'): parts.append(f"קומה {data['floor']}")

        final_address = " / ".join(parts)
        if data.get('extra'): final_address += f" ({data['extra']})"

        return final_address, google_status

    except Exception as e:
        return combined, f"שגיאה טכנית: {str(e)[:20]}"


# --- הפונקציה הראשית ---
def main():
    print("מתחיל עבודה...")
    root = tk.Tk();
    root.withdraw()

    o_key = KeyRequester("OpenAI Key", "הדביקי מפתח ChatGPT:").result
    g_key = KeyRequester("Google Key", "הדביקי מפתח Google Maps:").result
    if not o_key or not g_key: return

    client = OpenAI(api_key=o_key)
    path = filedialog.askopenfilename(title="בחרי אקסל")
    if not path: return

    df = pd.read_excel(path).fillna('')
    addresses, statuses = [], []

    for i, row in df.iterrows():
        print(f"מעבד שורה {i + 1}...")
        addr, stat = process_address_clean(row['ship_to_street1'], row['ship_to_street2'], row['site_name'], client,
                                           g_key)
        addresses.append(addr)
        statuses.append(stat)
        time.sleep(0.4)

    df['ship_to_street_final'] = addresses
    df['is_real_address'] = statuses  # עמודת הסטטוס שביקשת

    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="ניקוי_כתובות_בסיסי.xlsx")
    if save_path:
        df.to_excel(save_path, index=False)
        messagebox.showinfo("סיום", "העיבוד הושלם. בדקי את עמודת is_real_address")


# --- הסטרטר שמריץ את הכל ---
if __name__ == "__main__":
    main()