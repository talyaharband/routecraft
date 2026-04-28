import pandas as pd
from openai import OpenAI
import requests
import json
import tkinter as tk
from tkinter import filedialog, messagebox
from concurrent.futures import ThreadPoolExecutor, as_completed


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
def process_address_clean(index, s1, s2, city, client, g_key):
    combined = f"{s1} {s2}".strip()

    # שלב 1: פרומפט מדויק לאחידות ותיקונים מתקדמים
    system_prompt = (
        "You are an expert at parsing and cleaning Israeli addresses. "
        "Extract the address components into a strict JSON format. "
        "The JSON MUST have exactly these keys: 'street', 'number', 'apt', 'floor', 'extra', 'english_problem'.\n"
        "Rules:\n"
        "1. Fix typos and strictly UNIFY street names. Remove quotes or apostrophes from acronyms (e.g., 'חזון אי\"ש' -> 'חזון איש', 'הריטב\"א' -> 'הריטבא').\n"
        "2. Remove words like 'רחוב', 'שדרות', 'שד' from the street name entirely.\n"
        "3. Combine building number and letter (e.g. '15', 'א' -> '15א').\n"
        "4. Floor numbers with minus: standardize to '-1' (both '1-' and '-1' become '-1').\n"
        "5. If 'קומה' or 'דירה' appear WITHOUT any number nearby, completely ignore them and do not include them in the JSON.\n"
        "6. DO NOT include the city name anywhere in the output.\n"
        "7. Remove any dates or time markers from the text.\n"
        "8. If there are English words, translate them to Hebrew and match the closest known street name in the specific city. If you cannot translate/match it, set 'english_problem' to true. Otherwise, false.\n"
        "9. If a component is missing, return an empty string (\"\").\n\n"
        "Examples:\n"
        "Raw: 'רחוב הרצל 15 א קומה 2 כניסה ב', City: 'חיפה' -> {\"street\": \"הרצל\", \"number\": \"15א\", \"apt\": \"\", \"floor\": \"2\", \"extra\": \"כניסה ב\", \"english_problem\": false}\n"
        "Raw: 'הריטב\"א 12 דירה', City: 'אלעד' -> {\"street\": \"הריטבא\", \"number\": \"12\", \"apt\": \"\", \"floor\": \"\", \"extra\": \"\", \"english_problem\": false}\n"
        "Raw: 'חזון אי\"ש 10 קומה 1-', City: 'בני ברק' -> {\"street\": \"חזון איש\", \"number\": \"10\", \"apt\": \"\", \"floor\": \"-1\", \"extra\": \"\", \"english_problem\": false}\n"
        "Raw: 'הרקפת 4 בתאריך 12.05', City: 'תל אביב' -> {\"street\": \"הרקפת\", \"number\": \"4\", \"apt\": \"\", \"floor\": \"\", \"extra\": \"\", \"english_problem\": false}\n"
    )
    user_prompt = f"Raw Address: '{combined}', City: '{city}'"

    try:
        ai_res = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            response_format={"type": "json_object"},
            temperature=0.0
        )
        data = json.loads(ai_res.choices[0].message.content)

        street = data.get('street', '').strip()
        num = str(data.get('number', '')).strip()

        # שלב 2: אימות גיאוגרפי ושליפת קואורדינטות (על בסיס עיר, רחוב ומספר בלבד)
        google_status = "לא נבדק"
        coordinates = ""
        final_street = street

        if data.get('english_problem'):
            google_status = "❌ שגיאה: נמצאו מילים באנגלית ללא תרגום תואם לרחוב בעיר"
        elif street:
            # מגבילים את החיפוש לישראל כדי לדייק את גוגל
            url = f"https://maps.googleapis.com/maps/api/geocode/json?address={street} {num}, {city}&components=country:IL&key={g_key}&language=he"
            res = requests.get(url).json()

            if res.get('status') == 'OK':
                result = res['results'][0]
                types = result.get('types', [])
                
                # בודק אם זו כתובת אמיתית של בניין
                if 'street_address' in types or 'premise' in types or 'subpremise' in types:
                    google_status = "✅ כתובת אומתה במלואה (קיימת במפה)"
                    lat = result['geometry']['location']['lat']
                    lng = result['geometry']['location']['lng']
                    coordinates = f"{lat}, {lng}"
                else:
                    # מפרט את הבעיה במידה והיא לא מדוייקת
                    has_route = any('route' in c.get('types', []) for c in result.get('address_components', []))
                    if has_route:
                        google_status = "⚠️ בעיה: הרחוב נמצא, אך מספר הבית לא קיים במדויק במפה"
                    else:
                        google_status = "❓ בעיה: לא זוהה רחוב, גוגל מצא רק את האזור/עיר"
            else:
                reason = res.get('status', 'Unknown')
                if reason == 'ZERO_RESULTS':
                    google_status = "❌ שגיאה: הכתובת כלל לא נמצאה בישראל (Zero Results)"
                else:
                    google_status = f"❌ שגיאת מפות גוגל: {reason}"

        # שלב 3: הרכבת המחרוזת הסופית על בסיס הניקוי האחיד
        parts = []
        if final_street or num:
            parts.append(f"{final_street} {num}".strip())
            
        if data.get('apt'): parts.append(f"דירה {data['apt']}")
        if data.get('floor'): parts.append(f"קומה {data['floor']}")

        final_address = " / ".join(parts)
        if data.get('extra'): final_address += f" ({data['extra']})"
        
        if not final_address.strip():
            final_address = combined

        return index, final_address, google_status, coordinates

    except Exception as e:
        return index, combined, f"שגיאה טכנית: {str(e)[:20]}", ""


# --- הפונקציה הראשית ---
def main():
    print("מתחיל עבודה...")
    root = tk.Tk()
    root.withdraw()

    o_key = KeyRequester("OpenAI Key", "הדביקי מפתח ChatGPT:").result
    if not o_key: return
    g_key = KeyRequester("Google Key", "הדביקי מפתח Google Maps:").result
    if not g_key: return

    client = OpenAI(api_key=o_key)
    path = filedialog.askopenfilename(title="בחרי אקסל")
    if not path: return

    try:
        df = pd.read_excel(path).fillna('')
    except Exception as e:
        messagebox.showerror("שגיאה", f"שגיאה בטעינת הקובץ:\n{e}")
        return

    required_cols = ['ship_to_street1', 'ship_to_street2', 'site_name']
    for col in required_cols:
        if col not in df.columns:
            df[col] = '' 

    addresses = [""] * len(df)
    statuses = [""] * len(df)
    coords_list = [""] * len(df)

    print(f"מעבד {len(df)} שורות במקביל (משתמש ב-10 תהליכונים)...")

    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = {
            executor.submit(
                process_address_clean, 
                i, 
                row.get('ship_to_street1', ''), 
                row.get('ship_to_street2', ''), 
                row.get('site_name', ''), 
                client, 
                g_key
            ): i 
            for i, row in df.iterrows()
        }
        
        processed = 0
        for future in as_completed(futures):
            idx, addr, stat, coords = future.result()
            addresses[idx] = addr
            statuses[idx] = stat
            coords_list[idx] = coords
            
            processed += 1
            if processed % 10 == 0 or processed == len(df):
                print(f"עובדו {processed}/{len(df)} שורות...")

    df['ship_to_street_final'] = addresses
    df['is_real_address'] = statuses 
    df['coordinates'] = coords_list

    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="ניקוי_כתובות_מדויק.xlsx")
    if save_path:
        try:
            df.to_excel(save_path, index=False)
            messagebox.showinfo("סיום", "העיבוד הושלם בהצלחה. הקובץ נשמר.")
        except Exception as e:
             messagebox.showerror("שגיאה", f"לא ניתן לשמור את הקובץ. ייתכן שהוא פתוח בתוכנה אחרת:\n{e}")


# --- הסטרטר שמריץ את הכל ---
if __name__ == "__main__":
    main()