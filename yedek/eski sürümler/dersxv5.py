import sqlite3
import pandas as pd
import sys
import random
import os
import copy

# --- AYARLAR ---
DAYS = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma"]
SLOTS = ["09:00-12:00", "13:00-16:00", "16:00-19:00", "19:00-21:00"]
DB_PATH = 'okul.db'
PREF_FILE = 'tercih.xlsx'
CONSTR_FILE = 'kisit_formu.xlsx'
OUTPUT_FILE = 'isletme_ders_programi.xlsx'
MAX_TRIALS = 30  # Gün kısıtı eklendiği için deneme sayısını biraz artırmak iyidir


class Scheduler:
    def __init__(self, assignments, classrooms, preferences=[], constraints={}):
        self.assignments = assignments
        self.classrooms = classrooms
        self.initial_prefs = preferences
        self.schedule = copy.deepcopy(preferences)
        self.constraints = constraints
        self.total_lessons = len(assignments)
        self.max_depth = 0
        self.class_limits = {}

        counts = {}
        for a in assignments:
            counts[a['sinif']] = counts.get(a['sinif'], 0) + 1

        max_possible = len(DAYS) * len(SLOTS)
        for cls, count in counts.items():
            self.class_limits[cls] = 2 if count > max_possible else 1

    def is_valid(self, assignment, day, slot, classroom):
        hoca_adi = assignment['isim'].strip()
        if self.constraints.get((hoca_adi, day, slot)) == 0:
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
        # 1. SLOT DENGE PUANI (Slotların eşit doluluğu)
        slot_counts = {(d, s): 0 for d in DAYS for s in SLOTS}
        for entry in self.schedule:
            slot_counts[(entry['day'], entry['slot'])] += 1
        base_score = sum(v ** 2 for v in slot_counts.values())

        # 2. HOCA GÜN SAYISI CEZASI (Hocaların günlere yayılımı)
        hoca_gunleri = {}
        for entry in self.schedule:
            hoca = entry['isim']
            hoca_gunleri.setdefault(hoca, set()).add(entry['day'])

        hoca_ceza = 0
        for gun_kumesi in hoca_gunleri.values():
            if len(gun_kumesi) > 3:
                # 3 günü aşan her gün için ağır ceza puanı (Örn: 500)
                hoca_ceza += (len(gun_kumesi) - 3) * 500

        return base_score + hoca_ceza

    def backtrack(self, index=0):
        if index < len(self.assignments):
            curr = self.assignments[index]
            if any(d['ders_id'] == curr['ders_id'] and d['sinif'] == curr['sinif'] for d in self.initial_prefs):
                return self.backtrack(index + 1)

        if index == len(self.assignments): return True

        assignment = self.assignments[index]
        potential_slots = []
        for d in DAYS:
            for s in SLOTS:
                class_load = sum(
                    1 for e in self.schedule if e['day'] == d and e['slot'] == s and e['sinif'] == assignment['sinif'])
                global_load = sum(1 for e in self.schedule if e['day'] == d and e['slot'] == s)

                # Hocanın o gün zaten dersi var mı kontrolü
                hoca_o_gun_orada = any(e['day'] == d and e['isim'] == assignment['isim'] for e in self.schedule)

                if class_load < self.class_limits[assignment['sinif']]:
                    # Sıralama kriteri:
                    # 1. Sınıf çakışması (x[2])
                    # 2. Hocanın mevcut günü mü? (x[3]: True(1) ise yeni gündür, False(0) ise mevcut gündür)
                    # 3. Genel slot yükü (x[4])
                    potential_slots.append((d, s, class_load, not hoca_o_gun_orada, global_load))

        random.shuffle(potential_slots)
        # Önce sınıf çakışmasını minimize et, sonra hocayı aynı güne topla, sonra genel yükü dengele
        potential_slots.sort(key=lambda x: (x[2], x[3], x[4]))

        for day, slot, _, _, _ in potential_slots:
            sh_rooms = list(self.classrooms)
            random.shuffle(sh_rooms)
            for classroom in sh_rooms:
                if self.is_valid(assignment, day, slot, classroom):
                    self.schedule.append({**assignment, 'day': day, 'slot': slot, 'classroom': classroom})
                    if self.backtrack(index + 1): return True
                    self.schedule.pop()
        return False

    def report_conflicts(self):
        print("\n" + "=" * 70 + "\n--- ÇAKIŞMA VE DAĞILIM RAPORU ---\n" + "=" * 70)
        df = pd.DataFrame(self.schedule)

        # Sınıf çakışmaları
        grouped = df.groupby(['sinif', 'day', 'slot'])
        found = False
        for (sinif_adi, gun, saat), group in grouped:
            if len(group) > 1:
                found = True
                print(f"DIKKAT: {sinif_adi} -> {gun} ({saat})")
                for _, row in group.iterrows():
                    print(f"   [!] {row['ders_adi']} - {row['isim']}")
        if not found: print("Sınıf çakışması bulunmuyor.")

        # Hoca gün sayıları
        print("-" * 30)
        hoca_yayilim = df.groupby('isim')['day'].nunique()
        fazla_gun = hoca_yayilim[hoca_yayilim > 3]
        if not fazla_gun.empty:
            print("3 günden fazla gelen hocalar:")
            for h, g in fazla_gun.items():
                print(f"   [!] {h}: {g} gün")
        else:
            print("Tüm hocaların dersleri 3 gün veya altına toplandı.")
        print("=" * 70 + "\n")


# --- VERİ İŞLEME FONKSİYONLARI ---

def load_constraints():
    if not os.path.exists(CONSTR_FILE): return {}
    df = pd.read_excel(CONSTR_FILE, sheet_name='Ogretmen_Uygunluk')
    constraints = {}
    target_col = 'Uygun_mu (1=Evet, 0=Hayır)'
    for _, row in df.iterrows():
        hoca, gun, saat = str(row['Ogretim_Uyesi']).strip(), str(row['Gun']).strip(), str(row['Saat']).strip()
        val = row[target_col]
        constraints[(hoca, gun, saat)] = 1 if pd.isna(val) else int(val)
    return constraints


def load_preferences(all_assignments, classrooms):
    if not os.path.exists(PREF_FILE): return []
    print("Tercih formu okunuyor...")
    pref_df = pd.read_excel(PREF_FILE, index_col=0)
    preferences = []

    for day in DAYS:
        for slot in SLOTS:
            cell_content = pref_df.at[day, slot]
            if pd.notna(cell_content) and cell_content != "":
                content = str(cell_content).replace(',', '\n')
                entries = [e.strip() for e in content.split('\n') if e.strip()]
                for entry in entries:
                    if " - " in entry:
                        d_pref, h_pref = entry.split(" - ", 1)
                        d_pref_clean, h_pref_clean = d_pref.strip(), h_pref.strip()
                        match = next((a for a in all_assignments if
                                      a['ders_adi'].strip() == d_pref_clean and a['isim'].strip() == h_pref_clean),
                                     None)
                        if match:
                            if not any(p['ders_id'] == match['ders_id'] and p['sinif'] == match['sinif'] for p in
                                       preferences):
                                preferences.append(
                                    {**match, 'day': day, 'slot': slot, 'classroom': random.choice(classrooms)})
                        else:
                            print(f"\nHATA: Tercih Bulunamadı -> {d_pref_clean} ({day} {slot})");
                            sys.exit(1)
    return preferences


def get_data():
    conn = sqlite3.connect(DB_PATH)
    query = "SELECT oud.uye_id, ou.isim, oud.ders_id, d.ders_adi, oud.sinif FROM OgretimUyeleriDersler oud JOIN OgretimUyeleri ou ON oud.uye_id = ou.uye_id JOIN Dersler d ON oud.ders_id = d.ders_id"
    raw = pd.read_sql_query(query, conn).to_dict('records')
    final = []
    counts = {}
    for r in raw:
        for c in [s.strip() for s in r['sinif'].split(',') if s.strip()]:
            new_r = r.copy();
            new_r['sinif'] = c;
            final.append(new_r)
            counts[c] = counts.get(c, 0) + 1
    for a in final: a['priority'] = counts[a['sinif']]
    final.sort(key=lambda x: x['priority'], reverse=True)
    rooms = pd.read_sql_query("SELECT derslik_adi FROM Derslikler", conn)['derslik_adi'].tolist()
    conn.close()
    return final, rooms


def save_to_master_excel(schedule_data):
    df = pd.DataFrame(schedule_data)
    master_df = pd.DataFrame(index=DAYS, columns=SLOTS).fillna("")
    for (day, slot), group in df.groupby(['day', 'slot']):
        lines = [f"{r['classroom']}: {r['ders_adi']} [{r['sinif']}] - {r['isim']}" for _, r in
                 group.sort_values(by='classroom').iterrows()]
        master_df.at[day, slot] = "\n".join(lines)
    with pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter') as writer:
        master_df.to_excel(writer, sheet_name='Genel Ders Programı')
        ws = writer.sheets['Genel Ders Programı']
        fmt = writer.book.add_format({'text_wrap': True, 'valign': 'top', 'border': 1, 'font_size': 8})
        ws.set_column('B:E', 85, fmt)


# --- ANA DÖNGÜ ---

if __name__ == "__main__":
    try:
        assignments, classrooms = get_data()
        prefs = load_preferences(assignments, classrooms)
        constraints = load_constraints()

        best_schedule = None
        best_score = float('inf')

        print(f"Hoca günlerini optimize ederek planlama başlatıldı. (Deneme: {MAX_TRIALS})")

        for trial in range(1, MAX_TRIALS + 1):
            scheduler = Scheduler(assignments, classrooms, preferences=prefs, constraints=constraints)
            if scheduler.backtrack():
                score = scheduler.get_balance_score()
                if score < best_score:
                    best_score = score
                    best_schedule = copy.deepcopy(scheduler.schedule)
                print(f"Deneme {trial}/{MAX_TRIALS} bitti. (Maliyet Puanı: {score})")
            else:
                print(f"Deneme {trial}/{MAX_TRIALS} başarısız.")

        if best_schedule:
            print(f"\n*** En iyi çözüm ({best_score}) kaydediliyor. ***")
            save_to_master_excel(best_schedule)
            final_rep = Scheduler(assignments, classrooms)
            final_rep.schedule = best_schedule
            final_rep.report_conflicts()
            print(f"Sonuç: {OUTPUT_FILE}")
    except Exception as e:
        print(f"\nSistem Hatası: {e}")