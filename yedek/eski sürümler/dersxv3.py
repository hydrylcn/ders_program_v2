import sqlite3
import pandas as pd
import sys
import random
import os

# --- AYARLAR ---
DAYS = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma"]
SLOTS = ["09:00-12:00", "13:00-16:00", "16:00-19:00", "19:00-21:00"]
DB_PATH = 'okul.db'
PREF_FILE = 'tercih.xlsx'
CONSTR_FILE = 'kisit_formu.xlsx'
OUTPUT_FILE = 'isletme_ders_programi.xlsx'


class Scheduler:
    def __init__(self, assignments, classrooms, preferences=[], constraints={}):
        self.assignments = assignments
        self.classrooms = classrooms
        self.schedule = preferences
        self.constraints = constraints
        self.total_lessons = len(assignments)
        self.max_depth = 0
        self.class_limits = {}
        self.soft_mode = False

        counts = {}
        for a in assignments:
            counts[a['sinif']] = counts.get(a['sinif'], 0) + 1

        max_possible = len(DAYS) * len(SLOTS)
        for cls, count in counts.items():
            self.class_limits[cls] = 2 if count > max_possible else 1

    def is_valid(self, assignment, day, slot, classroom):
        # 1. HOCA KISIT KONTROLÜ
        hoca_adi = assignment['isim'].strip()
        # Eğer hoca o slot için 0 olarak işaretlenmişse (meşgulse) geçersizdir
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

    def backtrack(self, index=0):
        if index < len(self.assignments) and any(
                d['ders_id'] == self.assignments[index]['ders_id'] and d['sinif'] == self.assignments[index]['sinif']
                for d in self.schedule):
            return self.backtrack(index + 1)

        if index > self.max_depth:
            self.max_depth = index
            mode_text = "ESNEK" if self.soft_mode else "KATI"
            print(
                f"-> [{mode_text} MOD] İlerleme: %{(self.max_depth / self.total_lessons) * 100:.1f} ({self.max_depth}/{self.total_lessons})")
            sys.stdout.flush()

        if index == len(self.assignments): return True

        assignment = self.assignments[index]
        potential_slots = []
        for d in DAYS:
            for s in SLOTS:
                current_count = sum(
                    1 for e in self.schedule if e['day'] == d and e['slot'] == s and e['sinif'] == assignment['sinif'])
                potential_slots.append((d, s, current_count))

        random.shuffle(potential_slots)
        potential_slots.sort(key=lambda x: x[2])

        for day, slot, _ in potential_slots:
            sh_rooms = list(self.classrooms)
            random.shuffle(sh_rooms)
            for classroom in sh_rooms:
                if self.is_valid(assignment, day, slot, classroom):
                    self.schedule.append({**assignment, 'day': day, 'slot': slot, 'classroom': classroom})
                    if self.backtrack(index + 1): return True
                    self.schedule.pop()
        return False

    def report_conflicts(self):
        print("\n" + "=" * 70 + "\n--- ÇAKIŞMA RAPORU ---\n" + "=" * 70)
        df = pd.DataFrame(self.schedule)
        grouped = df.groupby(['sinif', 'day', 'slot'])
        for (sinif_adi, gun, saat), group in grouped:
            if len(group) > 1:
                print(f"DIKKAT: {sinif_adi} -> {gun} ({saat})")
                for _, row in group.iterrows():
                    print(f"   [!] {row['ders_adi']} - {row['isim']}")
        print("=" * 70 + "\n")


def load_constraints():
    if not os.path.exists(CONSTR_FILE): return {}

    print("Hoca kısıtları yükleniyor...")
    # 'Ogretmen_Uygunluk' sayfasını oku
    df = pd.read_excel(CONSTR_FILE, sheet_name='Ogretmen_Uygunluk')
    constraints = {}

    # Gerçek sütun adı: 'Uygun_mu (1=Evet, 0=Hayır)'
    target_col = 'Uygun_mu (1=Evet, 0=Hayır)'

    for _, row in df.iterrows():
        hoca = str(row['Ogretim_Uyesi']).strip()
        gun = str(row['Gun']).strip()
        saat = str(row['Saat']).strip()

        # Boş hücreleri 1 (Uygun) kabul et
        val = row[target_col]
        uygunluk = 1 if pd.isna(val) else int(val)

        constraints[(hoca, gun, saat)] = uygunluk

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
                # ÖNCE VİRGÜL VE SATIR SONUNA GÖRE PARÇALARA AYIR
                content = str(cell_content).replace(',', '\n')
                entries = [e.strip() for e in content.split('\n') if e.strip()]

                for entry in entries:
                    if " - " in entry:
                        d_pref, h_pref = entry.split(" - ", 1)
                        d_pref_clean = d_pref.strip()
                        h_pref_clean = h_pref.strip()

                        match = next((a for a in all_assignments if
                                      a['ders_adi'].strip() == d_pref_clean
                                      and a['isim'].strip() == h_pref_clean), None)

                        if match:
                            # Aynı dersin aynı sınıfa tekrar eklenmesini önle
                            if not any(p['ders_id'] == match['ders_id'] and p['sinif'] == match['sinif'] for p in
                                       preferences):
                                preferences.append(
                                    {**match, 'day': day, 'slot': slot, 'classroom': random.choice(classrooms)})
                        else:
                            print("\n" + "!" * 60)
                            print(f"HATA: Tercih formundaki ders veritabanında bulunamadı!")
                            print(f"Aranan Ders: '{d_pref_clean}'")
                            print(f"Aranan Hoca: '{h_pref_clean}'")
                            print(f"Konum: {day} - {slot}")
                            print("İpucu: Virgül veya tire kullanımını kontrol edin.")
                            print("!" * 60 + "\n")
                            sys.exit(1)
    return preferences


def get_data():
    conn = sqlite3.connect(DB_PATH)
    query = """SELECT oud.uye_id, ou.isim, oud.ders_id, d.ders_adi, oud.sinif 
               FROM OgretimUyeleriDersler oud 
               JOIN OgretimUyeleri ou ON oud.uye_id = ou.uye_id 
               JOIN Dersler d ON oud.ders_id = d.ders_id"""
    raw_data = pd.read_sql_query(query, conn).to_dict('records')
    final_assignments = []
    class_counts = {}
    for row in raw_data:
        for c in [s.strip() for s in row['sinif'].split(',') if s.strip()]:
            new_row = row.copy()
            new_row['sinif'] = c
            final_assignments.append(new_row)
            class_counts[c] = class_counts.get(c, 0) + 1
    for a in final_assignments: a['priority'] = class_counts[a['sinif']]
    final_assignments.sort(key=lambda x: x['priority'], reverse=True)
    classrooms = pd.read_sql_query("SELECT derslik_adi FROM Derslikler", conn)['derslik_adi'].tolist()
    conn.close()
    return final_assignments, classrooms


def save_to_master_excel(schedule_data):
    df = pd.DataFrame(schedule_data)
    master_df = pd.DataFrame(index=DAYS, columns=SLOTS).fillna("")
    for (day, slot), group in df.groupby(['day', 'slot']):
        lines = [f"{r['classroom']}: {r['ders_adi']} ({r['sinif']}) - {r['isim']}" for _, r in
                 group.sort_values(by='classroom').iterrows()]
        master_df.at[day, slot] = "\n".join(lines)
    with pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter') as writer:
        master_df.to_excel(writer, sheet_name='Genel Ders Programı')
        ws = writer.sheets['Genel Ders Programı']
        fmt = writer.book.add_format({'text_wrap': True, 'valign': 'top', 'border': 1, 'font_size': 8})
        ws.set_column('B:E', 85, fmt)


if __name__ == "__main__":
    try:
        assignments, classrooms = get_data()
        prefs = load_preferences(assignments, classrooms)
        constraints = load_constraints()

        scheduler = Scheduler(assignments, classrooms, preferences=prefs, constraints=constraints)

        # Matematiksel kontrol: Eğer bir sınıfın dersi > 20 ise soft_mode başlat
        max_slots = len(DAYS) * len(SLOTS)
        for a in assignments:
            if assignments.count(a) > max_slots: scheduler.soft_mode = True

        if scheduler.backtrack():
            print("\n*** Planlama Başarıyla Tamamlandı! ***")
            save_to_master_excel(scheduler.schedule)
            scheduler.report_conflicts()
            print(f"Sonuç dosyası: {OUTPUT_FILE}")
        else:
            print("\n!!! HATA: Mevcut kısıtlarla çözüm üretilemedi. !!!")
    except Exception as e:
        print(f"\nSistem Hatası: {e}")