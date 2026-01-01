import sqlite3
import pandas as pd
import sys
import random
import os
import copy
import time


# --- YOL YÖNETİMİ ---
def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


class Scheduler:
    def __init__(self, assignments, classrooms, preferences=[], constraints={}, ayarlar=None):
        self.assignments = assignments
        self.classrooms = classrooms
        self.initial_prefs = preferences
        self.schedule = copy.deepcopy(preferences)
        self.constraints = constraints
        self.start_time = 0
        self.ayarlar = ayarlar
        self.DAYS = ayarlar.get("DAYS", ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma"])
        self.SLOTS = ayarlar.get("SLOTS", ["09:00-12:00", "13:00-16:00", "16:00-19:00", "19:00-21:00"])
        self.MAX_DAYS_PER_LECTURER = ayarlar.get("MAX_DAYS_PER_LECTURER", 3)
        self.MIN_SLOT_GAP = ayarlar.get("MIN_SLOT_GAP", 1)
        self.TRIAL_TIMEOUT = ayarlar.get("TRIAL_TIMEOUT", 10)
        self.SPECIAL_CONSTRAINTS = ayarlar.get("SPECIAL_CONSTRAINTS", {})
        self.HOCA_GUN_CEZASI = 500
        self.class_limits = {}
        counts = {}
        for a in assignments:
            counts[a['sinif']] = counts.get(a['sinif'], 0) + 1
        for cls, count in counts.items():
            cls_str = str(cls)
            current_limit = 1
            is_affected_by_constraint = False
            for keyword, rules in self.SPECIAL_CONSTRAINTS.items():
                is_match = (keyword[1:] not in cls_str) if keyword.startswith("!") else (keyword in cls_str)
                if is_match:
                    is_affected_by_constraint = True
                    break
            if is_affected_by_constraint:
                current_limit = 10
            elif count > (len(self.DAYS) * len(self.SLOTS)):
                current_limit = 2
            self.class_limits[cls] = current_limit

    def is_valid(self, assignment, day, slot, classroom_name):
        if time.time() - self.start_time > self.TRIAL_TIMEOUT: return False
        sinif_adi = str(assignment.get('sinif', ""))
        for keyword, rules in self.SPECIAL_CONSTRAINTS.items():
            is_match = (keyword[1:] not in sinif_adi) if keyword.startswith("!") else (keyword in sinif_adi)
            if is_match:
                ctype = rules.get("type", "ONLY")
                if ctype == "ONLY":
                    if day not in rules['days'] or slot not in rules['slots']: return False
                elif ctype == "NEVER":
                    if day in rules['days'] and slot in rules['slots']: return False
        hoca_adi = assignment.get('isim', "").strip()
        if self.constraints.get((hoca_adi, day, slot)) == 0: return False
        current_slot_idx = self.SLOTS.index(slot)
        for entry in self.schedule:
            if entry.get('isim', "").strip() == hoca_adi and entry['day'] == day:
                if abs(current_slot_idx - self.SLOTS.index(entry['slot'])) < self.MIN_SLOT_GAP: return False
        class_count_in_slot = 0
        for entry in self.schedule:
            if entry['day'] == day and entry['slot'] == slot:
                if entry['uye_id'] == assignment['uye_id']: return False
                if entry['classroom'] == classroom_name: return False
                if entry.get('sinif') == assignment.get('sinif'): class_count_in_slot += 1
        return class_count_in_slot < self.class_limits.get(assignment.get('sinif'), 1)

    def backtrack(self, index=0):
        if time.time() - self.start_time > self.TRIAL_TIMEOUT: return False
        if index == len(self.assignments): return True
        assignment = self.assignments[index]
        if any(d.get('ders_id') == assignment['ders_id'] and d.get('sinif') == assignment['sinif'] for d in
               self.initial_prefs):
            return self.backtrack(index + 1)
        sinif_adi = str(assignment.get('sinif', ""))
        potential_slots = []
        for d in self.DAYS:
            for s in self.SLOTS:
                skip_slot = False
                for keyword, rules in self.SPECIAL_CONSTRAINTS.items():
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
        uygun_derslikler = [r for r in self.classrooms if r['kontenjan'] >= assignment.get('kontenjan', 0)]
        random.shuffle(uygun_derslikler)
        for day, slot, _, _, _ in potential_slots:
            for room_info in uygun_derslikler:
                classroom_name = room_info['derslik_adi']
                if self.is_valid(assignment, day, slot, classroom_name):
                    self.schedule.append({**assignment, 'day': day, 'slot': slot, 'classroom': classroom_name})
                    if self.backtrack(index + 1): return True
                    self.schedule.pop()
                    if time.time() - self.start_time > self.TRIAL_TIMEOUT: return False
        return False

    def get_balance_score(self):
        slot_counts = {(d, s): 0 for d in self.DAYS for s in self.SLOTS}
        for entry in self.schedule: slot_counts[(entry['day'], entry['slot'])] += 1
        base_score = sum(v ** 2 for v in slot_counts.values())
        hoca_gunleri = {}
        for entry in self.schedule:
            hoca = entry.get('isim', "Bilinmiyor")
            hoca_gunleri.setdefault(hoca, set()).add(entry['day'])
        hoca_ceza = sum((len(g) - self.MAX_DAYS_PER_LECTURER) * self.HOCA_GUN_CEZASI for g in hoca_gunleri.values() if
                        len(g) > self.MAX_DAYS_PER_LECTURER)
        return base_score + hoca_ceza


def get_data(db_path):
    conn = sqlite3.connect(db_path)
    query = "SELECT oud.uye_id, ou.isim, oud.ders_id, d.ders_adi, oud.sinif, oud.kontenjan FROM OgretimUyeleriDersler oud JOIN OgretimUyeleri ou ON oud.uye_id = ou.uye_id JOIN Dersler d ON oud.ders_id = d.ders_id"
    raw = pd.read_sql_query(query, conn).to_dict('records')
    final = []
    for r in raw:
        if not r['sinif']: continue
        for c in [s.strip() for s in str(r['sinif']).split(',') if s.strip()]:
            new_r = r.copy();
            new_r['sinif'] = c;
            final.append(new_r)
    rooms = pd.read_sql_query("SELECT derslik_adi, kontenjan FROM Derslikler", conn).to_dict('records')
    conn.close()
    return final, rooms


def load_constraints(constr_file):
    c_path = resource_path(constr_file)
    if not os.path.exists(c_path): return {}
    try:
        df = pd.read_excel(c_path, sheet_name='Ogretmen_Uygunluk')
        target_col = 'Uygun_mu (1=Evet, 0=Hayır)'
        return {(str(r['Ogretim_Uyesi']).strip(), str(r['Gun']).strip(), str(r['Saat']).strip()): (
            1 if pd.isna(r[target_col]) else int(r[target_col])) for _, r in df.iterrows()}
    except:
        return {}


def load_preferences(all_assignments, classrooms, pref_file, days, slots):
    p_path = resource_path(pref_file)
    if not os.path.exists(p_path): return []
    try:
        pref_df = pd.read_excel(p_path, index_col=0)
        preferences = []
        for d in days:
            for s in slots:
                if d in pref_df.index and s in pref_df.columns:
                    cell = pref_df.at[d, s]
                    if pd.notna(cell) and str(cell).strip() != "":
                        for entry in [e.strip() for e in str(cell).replace(',', '\n').split('\n') if e.strip()]:
                            if " - " in entry:
                                d_pref, h_pref = entry.split(" - ", 1)
                                match = next((a for a in all_assignments if
                                              a['ders_adi'].strip() == d_pref.strip() and a[
                                                  'isim'].strip() == h_pref.strip()), None)
                                if match:
                                    uygun = [r['derslik_adi'] for r in classrooms if
                                             r['kontenjan'] >= match.get('kontenjan', 0)]
                                    room = random.choice(uygun) if uygun else "Bilinmiyor"
                                    preferences.append({**match, 'day': d, 'slot': s, 'classroom': room})
        return preferences
    except:
        return []


def save_to_master_excel(schedule_data, score, output_file, days, slots):
    df = pd.DataFrame(schedule_data)
    master_df = pd.DataFrame(index=days, columns=slots).fillna("")
    for (day, slot), group in df.groupby(['day', 'slot']):
        lines = [f"{r['classroom']}: {r['ders_adi']} [{r['sinif']}] - {r['isim']}" for _, r in group.iterrows()]
        master_df.at[day, slot] = "\n".join(lines)
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        master_df.to_excel(writer, sheet_name='Genel Ders Programı')
        pd.DataFrame({"Kriter": ["Final Skor"], "Değer": [score]}).to_excel(writer, sheet_name='Rapor')


# --- SINAV TAKVİMİ (HATA DÜZELTİLDİ) ---

def save_exam_schedule(schedule_data, output_file, classrooms, days, slots):
    sub_slots_map = {"09:00-12:00": ["09:00-10:30", "10:30-12:00"], "13:00-16:00": ["13:00-14:30", "14:30-16:00"],
                     "16:00-19:00": ["16:00-17:30", "17:30-19:00"], "19:00-21:00": ["19:00-20:00", "20:00-21:00"]}
    all_exam_slots = []
    for s in slots: all_exam_slots.extend(sub_slots_map.get(s, [s]))

    exam_rooms_base = []
    for r in classrooms:
        new_room = r.copy()
        new_room['original_cap'] = r['kontenjan']
        if new_room['kontenjan'] > 50: new_room['kontenjan'] = 50
        exam_rooms_base.append(new_room)

    df = pd.DataFrame(schedule_data)
    exam_df = pd.DataFrame(index=days, columns=all_exam_slots).fillna("")

    for (day, slot), group in df.groupby(['day', 'slot']):
        sub_slots = sub_slots_map.get(slot, [slot])
        group_list = group.to_dict('records')
        mid = (len(group_list) + 1) // 2
        halves = [group_list[:mid], group_list[mid:]]

        for i, sub_group in enumerate(halves):
            if i >= len(sub_slots): break
            current_sub_slot = sub_slots[i]
            slot_entries = []
            used_rooms_in_slot = set()

            for row in sub_group:
                remaining_students = row['kontenjan']
                # DÜZELTME: Dersin TOPLAM kontenjanı 50'den büyükse, TÜM atamalar için sadece büyük sınıfları kullan.
                is_large_course = (row['kontenjan'] > 50)
                assigned_rooms = []

                available_rooms = exam_rooms_base.copy()
                random.shuffle(available_rooms)

                # Eğer ders bölünecekse, küçük sınıfları (Derslik 303 gibi) seçim listesinden kalıcı olarak çıkar.
                if is_large_course:
                    available_rooms = [r for r in available_rooms if r['original_cap'] >= 50]

                for room in available_rooms:
                    if room['derslik_adi'] not in used_rooms_in_slot:
                        assigned_rooms.append(room['derslik_adi'])
                        used_rooms_in_slot.add(room['derslik_adi'])
                        remaining_students -= room['kontenjan']
                        if remaining_students <= 0: break

                rooms_str = " + ".join(assigned_rooms) if assigned_rooms else "Uygun Büyük Derslik Yok"
                slot_entries.append(f"{rooms_str}: {row['ders_adi']} [{row['sinif']}] - {row['isim']}")
            exam_df.at[day, current_sub_slot] = "\n".join(slot_entries)

    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        exam_df.to_excel(writer, sheet_name='Sınav Takvimi')


def report_final(schedule, score, max_days):
    print(f"\n{'=' * 70}\n--- ÇAKIŞMA VE DAĞILIM RAPORU (Skor: {score}) ---\n{'=' * 70}")
    df = pd.DataFrame(schedule)
    if df.empty: return
    for (sin_adi, gun, saat), group in df.groupby(['sinif', 'day', 'slot']):
        if len(group) > 1: print(f"BİLGİ (Eşzamanlı Yerleşim): {sin_adi} -> {gun} ({saat})")
    hoca_gunleri = df.groupby('isim')['day'].nunique()
    fazla_gun = hoca_gunleri[hoca_gunleri > max_days]
    if not fazla_gun.empty:
        for hoca, gun_sayisi in fazla_gun.items(): print(f"   [!] {hoca}: {gun_sayisi} gün")
    else:
        print(f"Tüm hocalar {max_days} gün sınırına uyuyor.")


def arayuzden_baslat(ayarlar):
    try:
        assignments, classrooms = get_data(ayarlar["DB_PATH"])
        constraints = load_constraints(ayarlar["CONSTR_FILE"])
        prefs = load_preferences(assignments, classrooms, ayarlar["PREF_FILE"], ayarlar["DAYS"], ayarlar["SLOTS"])
        best_schedule, best_score = None, float('inf')
        for trial in range(1, ayarlar["MAX_TRIALS"] + 1):
            s = Scheduler(assignments, classrooms, preferences=prefs, constraints=constraints, ayarlar=ayarlar)
            s.start_time = time.time()
            if s.backtrack():
                cur_score = s.get_balance_score()
                if cur_score < best_score:
                    best_score = cur_score;
                    best_schedule = copy.deepcopy(s.schedule)
                print(f"Deneme {trial} başarılı. (Skor: {cur_score})")
            else:
                print(f"Deneme {trial} başarısız.")
        if best_schedule:
            save_to_master_excel(best_schedule, best_score, ayarlar["OUTPUT_FILE"], ayarlar["DAYS"], ayarlar["SLOTS"])
            save_exam_schedule(best_schedule,
                               os.path.join(os.path.dirname(ayarlar["OUTPUT_FILE"]), "isletme_sinav_takvimi.xlsx"),
                               classrooms, ayarlar["DAYS"], ayarlar["SLOTS"])
            report_final(best_schedule, best_score, ayarlar["MAX_DAYS_PER_LECTURER"])
            return True
        return False
    except Exception as e:
        print(f"\n❌ SİSTEM HATASI: {str(e)}");
        return False