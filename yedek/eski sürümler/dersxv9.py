import sqlite3
import pandas as pd
import sys
import random
import os
import copy
import time

# BU VERSİYONDA KONTENJAN BİLGİSİ EKLENDİ,
# DERSLİKLER BELİRLENDİKTEN SONRA RASTGELE SEÇİM YAPIYOR, BUNUN YERİNE +1 +2 EN UYGUN SINIF SEÇİLEBİLİR AMA HEP AYNI SINIFI SEÇER, TEMİZLİK YAPILACAK MI

# --- AYARLAR ---
DAYS = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma"]
SLOTS = ["09:00-12:00", "13:00-16:00", "16:00-19:00", "19:00-21:00"]
DB_PATH = 'okul.db'
PREF_FILE = 'tercih.xlsx'
CONSTR_FILE = 'kisit_formu.xlsx'
OUTPUT_FILE = 'isletme_ders_programi.xlsx'
MAX_TRIALS = 30

# --- ESNEK AYARLAR ---
MAX_DAYS_PER_LECTURER = 3
MIN_SLOT_GAP = 2
HOCA_GUN_CEZASI = 500
TRIAL_TIMEOUT = 10

# --- ÖZEL KISIT AYARLARI ---
SPECIAL_CONSTRAINTS = {
    "Tezsiz": {
        "type": "ONLY",
        "days": ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma"],
        "slots": ["19:00-21:00"]
    },
    "!Tezsiz": {
        "type": "NEVER",
        "days": ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma"],
        "slots": ["19:00-21:00"]
    }
}


class Scheduler:
    def __init__(self, assignments, classrooms, preferences=[], constraints={}):
        self.assignments = assignments
        self.classrooms = classrooms
        self.initial_prefs = preferences
        self.schedule = copy.deepcopy(preferences)
        self.constraints = constraints
        self.start_time = 0

        counts = {}
        for a in assignments:
            counts[a['sinif']] = counts.get(a['sinif'], 0) + 1

        self.class_limits = {}
        for cls, count in counts.items():
            cls_str = str(cls)
            excluded_slots_count = 0
            for keyword, rules in SPECIAL_CONSTRAINTS.items():
                is_match = (keyword[1:] not in cls_str) if keyword.startswith("!") else (keyword in cls_str)
                if is_match and rules.get("type") == "NEVER":
                    excluded_slots_count = len(rules['days']) * len(rules['slots'])
                    break

            available_slots = (len(DAYS) * len(SLOTS)) - excluded_slots_count
            if "Tezsiz" in cls_str:
                self.class_limits[cls] = 10
            elif count > available_slots:
                self.class_limits[cls] = 2
            else:
                self.class_limits[cls] = 1

    def is_valid(self, assignment, day, slot, classroom_name):
        if time.time() - self.start_time > TRIAL_TIMEOUT: return False
        sinif_adi = str(assignment.get('sinif', ""))

        for keyword, rules in SPECIAL_CONSTRAINTS.items():
            is_match = (keyword[1:] not in sinif_adi) if keyword.startswith("!") else (keyword in sinif_adi)
            if is_match:
                ctype = rules.get("type", "ONLY")
                if ctype == "ONLY":
                    if day not in rules['days'] or slot not in rules['slots']: return False
                elif ctype == "NEVER":
                    if day in rules['days'] and slot in rules['slots']: return False

        hoca_adi = assignment.get('isim', "").strip()
        if self.constraints.get((hoca_adi, day, slot)) == 0: return False

        current_slot_idx = SLOTS.index(slot)
        for entry in self.schedule:
            if entry.get('isim', "").strip() == hoca_adi and entry['day'] == day:
                if abs(current_slot_idx - SLOTS.index(entry['slot'])) < MIN_SLOT_GAP: return False

        class_count_in_slot = 0
        for entry in self.schedule:
            if entry['day'] == day and entry['slot'] == slot:
                if entry['uye_id'] == assignment['uye_id']: return False
                if entry['classroom'] == classroom_name: return False
                if entry.get('sinif') == assignment.get('sinif'): class_count_in_slot += 1
        return class_count_in_slot < self.class_limits.get(assignment.get('sinif'), 1)

    def backtrack(self, index=0):
        if time.time() - self.start_time > TRIAL_TIMEOUT: return False
        if index == len(self.assignments): return True

        assignment = self.assignments[index]
        if any(d.get('ders_id') == assignment['ders_id'] and d.get('sinif') == assignment['sinif'] for d in
               self.initial_prefs):
            return self.backtrack(index + 1)

        sinif_adi = str(assignment.get('sinif', ""))
        potential_slots = []
        for d in DAYS:
            for s in SLOTS:
                skip_slot = False
                for keyword, rules in SPECIAL_CONSTRAINTS.items():
                    is_match = (keyword[1:] not in sinif_adi) if keyword.startswith("!") else (keyword in sinif_adi)
                    if is_match:
                        if rules.get("type") == "ONLY" and (d not in rules['days'] or s not in rules['slots']):
                            skip_slot = True;
                            break
                        if rules.get("type") == "NEVER" and (d in rules['days'] and s in rules['slots']):
                            skip_slot = True;
                            break
                if skip_slot: continue

                class_load = sum(1 for e in self.schedule if
                                 e['day'] == d and e['slot'] == s and e.get('sinif') == assignment.get('sinif'))
                hoca_o_gun_orada = any(e['day'] == d and e.get('isim') == assignment.get('isim') for e in self.schedule)
                global_load = sum(1 for e in self.schedule if e['day'] == d and e['slot'] == s)

                if class_load < self.class_limits.get(assignment.get('sinif'), 1):
                    potential_slots.append((d, s, class_load, not hoca_o_gun_orada, global_load))

        random.shuffle(potential_slots)
        potential_slots.sort(key=lambda x: (x[2], x[3], x[4]))

        # Dersin kontenjanına uygun derslikleri filtrele
        ders_kontenjan = assignment.get('kontenjan', 0)
        uygun_derslikler = [r for r in self.classrooms if r['kontenjan'] >= ders_kontenjan]
        random.shuffle(uygun_derslikler)

        for day, slot, _, _, _ in potential_slots:
            for room_info in uygun_derslikler:
                classroom_name = room_info['derslik_adi']
                if self.is_valid(assignment, day, slot, classroom_name):
                    self.schedule.append({**assignment, 'day': day, 'slot': slot, 'classroom': classroom_name})
                    if self.backtrack(index + 1): return True
                    self.schedule.pop()
                    if time.time() - self.start_time > TRIAL_TIMEOUT: return False
        return False

    def get_balance_score(self):
        slot_counts = {(d, s): 0 for d in DAYS for s in SLOTS}
        for entry in self.schedule: slot_counts[(entry['day'], entry['slot'])] += 1
        base_score = sum(v ** 2 for v in slot_counts.values())
        hoca_gunleri = {}
        for entry in self.schedule:
            hoca = entry.get('isim', "Bilinmiyor")
            hoca_gunleri.setdefault(hoca, set()).add(entry['day'])
        hoca_ceza = sum((len(g) - MAX_DAYS_PER_LECTURER) * HOCA_GUN_CEZASI for g in hoca_gunleri.values() if
                        len(g) > MAX_DAYS_PER_LECTURER)
        return base_score + hoca_ceza


def report_final(schedule, score):
    print(f"\n{'=' * 70}\n--- ÇAKIŞMA VE DAĞILIM RAPORU (Skor: {score}) ---\n{'=' * 70}")
    df = pd.DataFrame(schedule)
    if df.empty: return
    grouped = df.groupby(['sinif', 'day', 'slot'])
    for (sin_adi, gun, saat), group in grouped:
        if len(group) > 1:
            print(f"BİLGİ (Eşzamanlı Yerleşim): {sin_adi} -> {gun} ({saat})")
    print("-" * 30)
    hoca_yayilim = df.groupby('isim')['day'].nunique()
    fazla_gun = hoca_yayilim[hoca_yayilim > MAX_DAYS_PER_LECTURER]
    if not fazla_gun.empty:
        print(f"{MAX_DAYS_PER_LECTURER} günden fazla gelen hocalar:")
        for h, g in fazla_gun.items():
            print(f"   [!] {h}: {g} gün")
    else:
        print(f"Tüm hocalar {MAX_DAYS_PER_LECTURER} gün sınırına uyuyor.")
    print("=" * 70 + "\n")


def get_data():
    conn = sqlite3.connect(DB_PATH)
    # DÜZELTME: Kontenjan bilgisi artık OgretimUyeleriDersler (oud) tablosundan çekiliyor
    query = """
        SELECT oud.uye_id, ou.isim, oud.ders_id, d.ders_adi, oud.sinif, oud.kontenjan 
        FROM OgretimUyeleriDersler oud 
        JOIN OgretimUyeleri ou ON oud.uye_id = ou.uye_id 
        JOIN Dersler d ON oud.ders_id = d.ders_id
    """
    raw = pd.read_sql_query(query, conn).to_dict('records')
    final = []
    for r in raw:
        if not r['sinif']: continue
        for c in [s.strip() for s in str(r['sinif']).split(',') if s.strip()]:
            new_r = r.copy();
            new_r['sinif'] = c
            is_priority = any((k[1:] if k.startswith("!") else k) in c for k in SPECIAL_CONSTRAINTS.keys())
            new_r['priority'] = 100 if is_priority else 0
            final.append(new_r)
    final.sort(key=lambda x: x['priority'], reverse=True)

    rooms = pd.read_sql_query("SELECT derslik_adi, kontenjan FROM Derslikler", conn).to_dict('records')
    conn.close()
    return final, rooms


def save_to_master_excel(schedule_data, score):
    df = pd.DataFrame(schedule_data)
    master_df = pd.DataFrame(index=DAYS, columns=SLOTS).fillna("")
    for (day, slot), group in df.groupby(['day', 'slot']):
        lines = [f"{r['classroom']}: {r['ders_adi']} [{r['sinif']}] - {r['isim']}" for _, r in
                 group.sort_values(by='classroom').iterrows()]
        master_df.at[day, slot] = "\n".join(lines)
    with pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter') as writer:
        master_df.to_excel(writer, sheet_name='Genel Ders Programı')
        pd.DataFrame({"Kriter": ["Final Skor"], "Değer": [score]}).to_excel(writer, sheet_name='Rapor')
        ws = writer.sheets['Genel Ders Programı']
        ws.set_column('B:E', 85,
                      writer.book.add_format({'text_wrap': True, 'valign': 'top', 'border': 1, 'font_size': 8}))


def load_constraints():
    if not os.path.exists(CONSTR_FILE): return {}
    df = pd.read_excel(CONSTR_FILE, sheet_name='Ogretmen_Uygunluk')
    target_col = 'Uygun_mu (1=Evet, 0=Hayır)'
    return {(str(r['Ogretim_Uyesi']).strip(), str(r['Gun']).strip(), str(r['Saat']).strip()): (
        1 if pd.isna(r[target_col]) else int(r[target_col])) for _, r in df.iterrows()}


def load_preferences(all_assignments, classrooms):
    if not os.path.exists(PREF_FILE): return []
    pref_df = pd.read_excel(PREF_FILE, index_col=0)
    preferences = []
    for day in DAYS:
        for slot in SLOTS:
            cell = pref_df.at[day, slot]
            if pd.notna(cell) and str(cell).strip() != "":
                for entry in [e.strip() for e in str(cell).replace(',', '\n').split('\n') if e.strip()]:
                    if " - " in entry:
                        d_pref, h_pref = entry.split(" - ", 1)
                        match = next((a for a in all_assignments if
                                      a['ders_adi'].strip() == d_pref.strip() and a['isim'].strip() == h_pref.strip()),
                                     None)
                        if match and not any(
                                p['ders_id'] == match['ders_id'] and p['sinif'] == match['sinif'] for p in preferences):
                            uygun = [r['derslik_adi'] for r in classrooms if
                                     r['kontenjan'] >= match.get('kontenjan', 0)]
                            room = random.choice(uygun) if uygun else (
                                classrooms[0]['derslik_adi'] if classrooms else "Bilinmiyor")
                            preferences.append({**match, 'day': day, 'slot': slot, 'classroom': room})
    return preferences


if __name__ == "__main__":
    try:
        assignments, classrooms = get_data()
        prefs = load_preferences(assignments, classrooms);
        constraints = load_constraints()
        best_schedule, best_score = None, float('inf')
        for trial in range(1, MAX_TRIALS + 1):
            s = Scheduler(assignments, classrooms, preferences=prefs, constraints=constraints)
            s.start_time = time.time()
            if s.backtrack():
                cur_score = s.get_balance_score()
                if cur_score < best_score:
                    best_score = cur_score
                    best_schedule = copy.deepcopy(s.schedule)
                print(f"Deneme {trial}/{MAX_TRIALS} başarılı. (Skor: {cur_score})")
            else:
                print(f"Deneme {trial}/{MAX_TRIALS} başarısız.")

        if best_schedule:
            save_to_master_excel(best_schedule, best_score)
            report_final(best_schedule, best_score)
            print(f"BAŞARILI! Sonuç: {OUTPUT_FILE} (Skor: {best_score})")
        else:
            print("\nÇözüm bulunamadı. Lütfen kısıtları kontrol edin.")
    except Exception as e:
        print(f"\nHata: {e}")