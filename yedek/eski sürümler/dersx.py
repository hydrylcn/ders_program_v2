import sqlite3
import pandas as pd
import sys
import random

# --- AYARLAR ---
DAYS = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma"]
SLOTS = ["09:00-12:00", "13:00-16:00", "16:00-19:00", "19:00-21:00"]
DB_PATH = 'okul.db'
OUTPUT_FILE = 'isletme_ders_programi.xlsx'


class Scheduler:
    def __init__(self, assignments, classrooms):
        self.assignments = assignments
        self.classrooms = classrooms
        self.schedule = []
        self.total_lessons = len(assignments)
        self.max_depth = 0
        self.class_limits = {}
        self.soft_mode = False

        # Sınıf bazlı ders sayılarını hesapla
        counts = {}
        for a in assignments:
            counts[a['sinif']] = counts.get(a['sinif'], 0) + 1

        # Eğer ders sayısı slot sayısından fazlaysa o sınıf için esnek modu (limit=2) aktif et
        max_possible = len(DAYS) * len(SLOTS)
        for cls, count in counts.items():
            self.class_limits[cls] = 2 if count > max_possible else 1

    def is_valid(self, assignment, day, slot, classroom):
        class_count_in_slot = 0
        for entry in self.schedule:
            if entry['day'] == day and entry['slot'] == slot:
                # Hoca ve Derslik çakışması KESİNLİKLE yasak
                if entry['uye_id'] == assignment['uye_id']: return False
                if entry['classroom'] == classroom: return False

                # Sınıf bazlı dinamik limit kontrolü
                if entry['sinif'] == assignment['sinif']:
                    class_count_in_slot += 1

        return class_count_in_slot < self.class_limits[assignment['sinif']]

    def backtrack(self, index=0):
        if index > self.max_depth:
            self.max_depth = index
            mode_text = "ESNEK" if self.soft_mode else "KATI"
            print(
                f"-> [{mode_text} MOD] İlerleme: %{(self.max_depth / self.total_lessons) * 100:.1f} ({self.max_depth}/{self.total_lessons} ders yerleştirildi)")
            sys.stdout.flush()

        if index == len(self.assignments): return True

        assignment = self.assignments[index]

        # ÖNCELİKLİ SLOT SEÇİMİ: Önce bu sınıfın hiç dersi olmayan slotları belirle
        possible_slots = []
        for d in DAYS:
            for s in SLOTS:
                current_count = sum(
                    1 for e in self.schedule if e['day'] == d and e['slot'] == s and e['sinif'] == assignment['sinif'])
                if current_count < self.class_limits[assignment['sinif']]:
                    possible_slots.append((d, s, current_count))

        # Dengeli dağıtım için karıştır, sonra çakışma sayısına göre sırala (0 olanlar başa gelir)
        random.shuffle(possible_slots)
        possible_slots.sort(key=lambda x: x[2])

        for day, slot, _ in possible_slots:
            sh_rooms = list(self.classrooms)
            random.shuffle(sh_rooms)
            for classroom in sh_rooms:
                if self.is_valid(assignment, day, slot, classroom):
                    self.schedule.append({**assignment, 'day': day, 'slot': slot, 'classroom': classroom})
                    if self.backtrack(index + 1): return True
                    self.schedule.pop()
        return False

    def report_conflicts(self):
        """Çakışan dersleri terminale bilgi olarak yazar."""
        print("\n" + "=" * 70)
        print("--- ÇAKIŞMA RAPORU (AYNI SINIFIN AYNI SAATTEKİ DERSLERİ) ---")
        print("=" * 70)

        df = pd.DataFrame(self.schedule)
        grouped = df.groupby(['sinif', 'day', 'slot'])

        has_conflict = False
        for (sinif_adi, gun, saat), group in grouped:
            if len(group) > 1:
                has_conflict = True
                print(f"DIKKAT: {sinif_adi} -> {gun} ({saat})")
                for _, row in group.iterrows():
                    print(f"   [!] {row['ders_adi']} - {row['isim']} (Derslik: {row['classroom']})")
                print("-" * 30)

        if not has_conflict:
            print("Tebrikler! Herhangi bir sınıf çakışması bulunmuyor.")
        print("=" * 70 + "\n")


def get_data():
    conn = sqlite3.connect(DB_PATH)
    query = """
    SELECT oud.uye_id, ou.isim, oud.ders_id, d.ders_adi, oud.sinif 
    FROM OgretimUyeleriDersler oud 
    JOIN OgretimUyeleri ou ON oud.uye_id = ou.uye_id 
    JOIN Dersler d ON oud.ders_id = d.ders_id
    """
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
        lines = []
        for _, r in group.sort_values(by='classroom').iterrows():
            lines.append(f"{r['classroom']}: {r['ders_adi']} ({r['sinif']}) - {r['isim']}")
        # Ayırıcı çizgiyi kaldırdık, sadece satır sonu ile birleştiriyoruz
        master_df.at[day, slot] = "\n".join(lines)

    with pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter') as writer:
        master_df.to_excel(writer, sheet_name='Genel Ders Programı')
        ws = writer.sheets['Genel Ders Programı']
        fmt = writer.book.add_format({'text_wrap': True, 'valign': 'top', 'border': 1, 'font_size': 8})
        ws.set_column('B:E', 80, fmt)


def check_feasibility(assignments):
    counts = {}
    for a in assignments:
        counts[a['sinif']] = counts.get(a['sinif'], 0) + 1
    max_slots = len(DAYS) * len(SLOTS)
    for cls, count in counts.items():
        if count > max_slots: return False, cls, count
    return True, None, None


if __name__ == "__main__":
    try:
        assignments, classrooms = get_data()
        scheduler = Scheduler(assignments, classrooms)

        feasible, b_cls, b_count = check_feasibility(assignments)
        if not feasible:
            scheduler.soft_mode = True
        else:
            scheduler.soft_mode = False

        print(f"Toplam {len(assignments)} ders oturumu planlanıyor...")

        if scheduler.backtrack():
            print("\n*** Planlama Başarıyla Tamamlandı! ***")
            save_to_master_excel(scheduler.schedule)
            scheduler.report_conflicts()
            print(f"Excel Dosyası: {OUTPUT_FILE}")
        else:
            print("\n<!> Katı modda çözüm bulunamadı, limitler esnetiliyor...")
            for cls in scheduler.class_limits:
                scheduler.class_limits[cls] = 2
            scheduler.schedule = [];
            scheduler.max_depth = 0
            scheduler.soft_mode = True
            if scheduler.backtrack():
                save_to_master_excel(scheduler.schedule)
                scheduler.report_conflicts()
                print("\n*** Esnek Modda Çözüm Üretildi. ***")
            else:
                print("\n!!! HATA: Çözüm üretilemedi. !!!")

    except Exception as e:
        print(f"\nHata oluştu: {e}")