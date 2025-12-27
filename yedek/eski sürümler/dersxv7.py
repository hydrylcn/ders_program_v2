import sqlite3
import pandas as pd
import sys
import random
import os
import copy
import time

# BU VERSİYONDA ÖZEL KISIT AYARLARINA ZORUNLU INCLUDE EKLENDİ

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
TRIAL_TIMEOUT = 5

# --- ÖZEL KISIT AYARLARI ---
# Bir dersin sınıf adında anahtar kelime geçiyorsa, o ders SADECE belirtilen gün ve saatlere yazılır.
SPECIAL_CONSTRAINTS = {
    "Tezsiz": {
        "days": ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma"],
        "slots": ["19:00-21:00"]
    },
    "1. Sınıf": {
        "days": ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma"],
        "slots": ["09:00-12:00", "13:00-16:00", "16:00-19:00", "19:00-21:00"]
    }
}



class Scheduler:
    def __init__(self, assignments, classrooms, preferences=[], constraints={}):
        self.assignments = assignments
        self.classrooms = classrooms
        self.initial_prefs = preferences
        self.schedule = copy.deepcopy(preferences)
        self.constraints = constraints
        self.total_lessons = len(assignments)
        self.class_limits = {}
        self.start_time = 0

        counts = {}
        for a in assignments:
            counts[a['sinif']] = counts.get(a['sinif'], 0) + 1

        for cls, count in counts.items():
            # Tezsiz dersler için otomatik bypass (Eşzamanlı ders izni)
            if "Tezsiz" in str(cls):
                self.class_limits[cls] = 10
            else:
                max_possible = len(DAYS) * len(SLOTS)
                self.class_limits[cls] = 2 if count > max_possible else 1

    def is_valid(self, assignment, day, slot, classroom):
        if time.time() - self.start_time > TRIAL_TIMEOUT:
            return False

        sinif_adi = str(assignment['sinif'])

        # DİNAMİK KISIT KONTROLÜ
        for keyword, rules in SPECIAL_CONSTRAINTS.items():
            if keyword in sinif_adi:
                if day not in rules['days'] or slot not in rules['slots']:
                    return False

        hoca_adi = assignment['isim'].strip()
        if self.constraints.get((hoca_adi, day, slot)) == 0:
            return False

        current_slot_idx = SLOTS.index(slot)
        for entry in self.schedule:
            if entry['isim'].strip() == hoca_adi and entry['day'] == day:
                existing_slot_idx = SLOTS.index(entry['slot'])
                if abs(current_slot_idx - existing_slot_idx) < MIN_SLOT_GAP:
                    return False

        class_count_in_slot = 0
        for entry in self.schedule:
            if entry['day'] == day and entry['slot'] == slot:
                if entry['uye_id'] == assignment['uye_id']: return False
                if entry['classroom'] == classroom: return False
                if entry['sinif'] == assignment['sinif']:
                    class_count_in_slot += 1
        return class_count_in_slot < self.class_limits[assignment['sinif']]

    def get_balance_score(self):
        slot_counts = {(d, s): 0 for d in DAYS for s in SLOTS}
        for entry in self.schedule:
            slot_counts[(entry['day'], entry['slot'])] += 1
        base_score = sum(v ** 2 for v in slot_counts.values())

        hoca_gunleri = {}
        for entry in self.schedule:
            hoca = entry['isim']
            hoca_gunleri.setdefault(hoca, set()).add(entry['day'])

        hoca_ceza = 0
        for gun_kumesi in hoca_gunleri.values():
            if len(gun_kumesi) > MAX_DAYS_PER_LECTURER:
                hoca_ceza += (len(gun_kumesi) - MAX_DAYS_PER_LECTURER) * HOCA_GUN_CEZASI
        return base_score + hoca_ceza

    def backtrack(self, index=0):
        if time.time() - self.start_time > TRIAL_TIMEOUT:
            return False

        if index < len(self.assignments):
            curr = self.assignments[index]
            if any(d['ders_id'] == curr['ders_id'] and d['sinif'] == curr['sinif'] for d in self.initial_prefs):
                return self.backtrack(index + 1)

        if index == len(self.assignments): return True

        assignment = self.assignments[index]
        sinif_adi = str(assignment['sinif'])

        potential_slots = []
        for d in DAYS:
            for s in SLOTS:
                # DİNAMİK BUDAMA: Özel kısıtlara uymayan slotları en baştan ele
                skip_slot = False
                for keyword, rules in SPECIAL_CONSTRAINTS.items():
                    if keyword in sinif_adi:
                        if d not in rules['days'] or s not in rules['slots']:
                            skip_slot = True
                            break
                if skip_slot:
                    continue

                class_load = sum(
                    1 for e in self.schedule if e['day'] == d and e['slot'] == s and e['sinif'] == assignment['sinif'])
                hoca_o_gun_orada = any(e['day'] == d and e['isim'] == assignment['isim'] for e in self.schedule)
                global_load = sum(1 for e in self.schedule if e['day'] == d and e['slot'] == s)

                if class_load < self.class_limits[assignment['sinif']]:
                    potential_slots.append((d, s, class_load, not hoca_o_gun_orada, global_load))

        random.shuffle(potential_slots)
        potential_slots.sort(key=lambda x: (x[2], x[3], x[4]))

        for day, slot, _, _, _ in potential_slots:
            sh_rooms = list(self.classrooms)
            random.shuffle(sh_rooms)
            for classroom in sh_rooms:
                if self.is_valid(assignment, day, slot, classroom):
                    self.schedule.append({**assignment, 'day': day, 'slot': slot, 'classroom': classroom})
                    if self.backtrack(index + 1): return True
                    self.schedule.pop()
                    if time.time() - self.start_time > TRIAL_TIMEOUT:
                        return False
        return False

    def report_conflicts(self, final_score):
        print(f"\n{'=' * 70}\n--- ÇAKIŞMA VE DAĞILIM RAPORU (Skor: {final_score}) ---\n{'=' * 70}")
        df = pd.DataFrame(self.schedule)
        grouped = df.groupby(['sinif', 'day', 'slot'])
        found = False
        for (sinif_adi, gun, saat), group in grouped:
            if len(group) > 1:
                found = True
                print(f"BİLGİ (Eşzamanlı Yerleşim): {sinif_adi} -> {gun} ({saat})")
        if not found: print("Sınıf çakışması bulunmuyor.")
        print("-" * 30)
        hoca_yayilim = df.groupby('isim')['day'].nunique()
        fazla_gun = hoca_yayilim[hoca_yayilim > MAX_DAYS_PER_LECTURER]
        if not fazla_gun.empty:
            print(f"{MAX_DAYS_PER_LECTURER} günden fazla gelen hocalar:")
            for h, g in fazla_gun.items():
                print(f"   [!] {h}: {g} gün")
        else:
            print(f"Tüm hocaların dersleri {MAX_DAYS_PER_LECTURER} gün veya altına toplandı.")
        print("=" * 70 + "\n")


# --- VERİ İŞLEME VE DİĞER FONKSİYONLAR (Öncekiyle Aynı) ---
def get_data():
    conn = sqlite3.connect(DB_PATH)
    query = "SELECT oud.uye_id, ou.isim, oud.ders_id, d.ders_adi, oud.sinif FROM OgretimUyeleriDersler oud JOIN OgretimUyeleri ou ON oud.uye_id = ou.uye_id JOIN Dersler d ON oud.ders_id = d.ders_id"
    raw = pd.read_sql_query(query, conn).to_dict('records')
    final = []
    for r in raw:
        for c in [s.strip() for s in r['sinif'].split(',') if s.strip()]:
            new_r = r.copy()
            new_r['sinif'] = c

            # Kısıtlı derslere öncelik ver (Çözümü kolaylaştırır)
            is_priority = any(keyword in c for keyword in SPECIAL_CONSTRAINTS.keys())
            new_r['priority'] = 100 if is_priority else 0
            final.append(new_r)
    final.sort(key=lambda x: x['priority'], reverse=True)
    rooms = pd.read_sql_query("SELECT derslik_adi FROM Derslikler", conn)['derslik_adi'].tolist()
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
        fmt = writer.book.add_format({'text_wrap': True, 'valign': 'top', 'border': 1, 'font_size': 8})
        ws.set_column('B:E', 85, fmt)


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
            if pd.notna(cell) and cell != "":
                content = str(cell).replace(',', '\n')
                for entry in [e.strip() for e in content.split('\n') if e.strip()]:
                    if " - " in entry:
                        d_pref, h_pref = entry.split(" - ", 1)
                        match = next((a for a in all_assignments if
                                      a['ders_adi'].strip() == d_pref.strip() and a['isim'].strip() == h_pref.strip()),
                                     None)
                        if match:
                            if not any(p['ders_id'] == match['ders_id'] and p['sinif'] == match['sinif'] for p in
                                       preferences):
                                preferences.append(
                                    {**match, 'day': day, 'slot': slot, 'classroom': random.choice(classrooms)})
    return preferences


if __name__ == "__main__":
    try:
        assignments, classrooms = get_data()
        prefs = load_preferences(assignments, classrooms)
        constraints = load_constraints()
        best_schedule, best_score = None, float('inf')

        print("Planlama başlatıldı...")
        for trial in range(1, MAX_TRIALS + 1):
            s = Scheduler(assignments, classrooms, preferences=prefs, constraints=constraints)
            s.start_time = time.time()
            if s.backtrack():
                score = s.get_balance_score()
                if score < best_score:
                    best_score = score
                    best_schedule = copy.deepcopy(s.schedule)
                print(f"Deneme {trial}/{MAX_TRIALS} başarılı. (Skor: {score})")
            else:
                print(f"Deneme {trial}/{MAX_TRIALS} başarısız.")

        if best_schedule:
            save_to_master_excel(best_schedule, best_score)
            final_rep = Scheduler(assignments, classrooms)
            final_rep.schedule = best_schedule
            final_rep.report_conflicts(best_score)
            print(f"BAŞARILI! Sonuç: {OUTPUT_FILE} (Skor: {best_score})")
    except Exception as e:
        print(f"\nHata: {e}")