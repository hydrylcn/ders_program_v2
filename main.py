# -*- coding: utf-8 -*-
import customtkinter as ctk
from tkinter import filedialog, messagebox
import sys
import os
import threading
import queue
import shutil

# Projedeki diğer modüller
import db
import ders
import html
import htmlxv2

# Grupların varsayılan gün ve saat ayarları (3. Sınıf için özel matris tanımı)
DEFAULT_PRESETS = {
    "Tezsiz": {
        "days": ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma"],
        "slots": ["19:00-21:00"]
    }
}


class ConsoleRedirector:
    def __init__(self, log_queue):
        self.log_queue = log_queue

    def write(self, string):
        self.log_queue.put(string)

    def flush(self):
        pass


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("light")
        self.title("Ders Programı Planlama Paneli")
        width, height = 950, 1000

        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        x = (screen_w - width) // 2
        y = (screen_h - height) // 2
        self.geometry(f"{width}x{height}+{x}+{y}")
        self.configure(fg_color="white")

        self.log_queue = queue.Queue()
        self.constraint_rows = []
        self.ALL_DAYS = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma"]
        self.ALL_SLOTS = ["09:00-12:00", "13:00-16:00", "16:00-19:00", "19:00-21:00"]

        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.pack(fill="both", expand=True, padx=10, pady=10)

        self.scroll_frame = ctk.CTkScrollableFrame(self.main_container, label_text="DERS PROGRAMI AYARLARI",
                                                   fg_color="#fcfcfc")
        self.scroll_frame.pack(padx=10, pady=(0, 10), fill="both", expand=True)

        self.setup_ui()

        self.run_btn = ctk.CTkButton(self, text="HESAPLAMAYI BAŞLAT", height=65, command=self.start_thread,
                                     fg_color="#27ae60", hover_color="#2ecc71", font=("Arial", 20, "bold"))
        self.run_btn.pack(pady=20, padx=25, fill="x", side="bottom")

        self.after(100, self.check_queue)

    def setup_ui(self):
        self.add_section("1. DOSYA YOLLARI")
        self.dersler_excel_ent = self.add_file_input("Dersler Excel:", "dersler.xlsx")
        self.pref_file_ent = self.add_file_input("Tercih Excel:", "tercih.xlsx")
        self.constr_file_ent = self.add_file_input("Kısıt Excel:", "kisit_formu.xlsx")
        self.output_file_ent = self.add_input("Çıktı Excel Adı:", "isletme_ders_programi.xlsx")

        self.add_section("2. ALGORİTMA AYARLARI")
        top_settings = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
        top_settings.pack(fill="x")
        self.max_trials_ent = self.add_input("Max Deneme:", "30", master=top_settings, side="left")
        self.timeout_ent = self.add_input("Zaman Aşımı (sn):", "30", master=top_settings, side="left")

        self.max_days_ent = self.add_input("Dersler kaç güne toplansın?:", "3")
        self.min_gap_ent = self.add_input("Ders aralarında boşluk (0,1,2):", "1")

        self.add_section("3. ÖZEL KISIT TANIMLARI")
        self.constraints_container = ctk.CTkFrame(self.scroll_frame, fg_color="#f0f0f0", corner_radius=10)
        self.constraints_container.pack(fill="x", padx=10, pady=10)

        ctk.CTkButton(self.scroll_frame, text="+ Yeni Kısıt Ekle", command=self.add_constraint_row,
                      fg_color="#3498db", hover_color="#2980b9").pack(pady=10)

        # --- GÜNCELLENMİŞ DEFAULT KISITLAR ---
        self.add_constraint_row("Tezsiz", "SADECE")
        self.add_constraint_row("Tezsiz", "ASLA", is_inverse=True)

        # 3. Sınıf Kısıtı: Sadece SADECE, Çarşamba 13-16 hariç tüm gündüz saatleri seçili
        s3_pairs = []
        for d in self.ALL_DAYS:
            s3_pairs.append([d, "09:00-12:00"])
            if d != "Çarşamba":  # Çarşamba günü 13:00-16:00 eklenmiyor
                s3_pairs.append([d, "13:00-16:00"])

        self.add_constraint_row("3. Sınıf", "SADECE", initial_pairs=s3_pairs)

    def add_section(self, text):
        ctk.CTkLabel(self.scroll_frame, text=text, font=("Arial", 14, "bold"), text_color="#34495e").pack(anchor="w",
                                                                                                          padx=10,
                                                                                                          pady=(15, 5))

    def add_input(self, label_text, default_val, master=None, side="top"):
        target = master if master else self.scroll_frame
        frame = ctk.CTkFrame(target, fg_color="transparent")
        frame.pack(fill="x", padx=10, pady=2, side=side)
        ctk.CTkLabel(frame, text=label_text, width=250, anchor="w").pack(side="left")
        ent = ctk.CTkEntry(frame, fg_color="white", width=150)
        ent.insert(0, default_val)
        ent.pack(side="left", fill="x", expand=True, padx=5)
        return ent

    def add_file_input(self, label_text, default_val):
        frame = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
        frame.pack(fill="x", padx=10, pady=2)
        ctk.CTkLabel(frame, text=label_text, width=250, anchor="w").pack(side="left")
        ent = ctk.CTkEntry(frame, fg_color="white")
        ent.insert(0, default_val)
        ent.pack(side="left", fill="x", expand=True, padx=5)
        ctk.CTkButton(frame, text="Seç", width=60, command=lambda: self.browse_file(ent)).pack(side="right")
        return ent

    def browse_file(self, entry):
        filename = filedialog.askopenfilename(filetypes=[("Excel Dosyaları", "*.xlsx")])
        if filename:
            entry.delete(0, "end")
            entry.insert(0, filename)

    def add_constraint_row(self, key="", t="SADECE", is_inverse=False, initial_pairs=None):
        row_frame = ctk.CTkFrame(self.constraints_container, fg_color="white", corner_radius=8, border_width=1,
                                 border_color="#ccc")
        row_frame.pack(fill="x", padx=10, pady=10)

        top_line = ctk.CTkFrame(row_frame, fg_color="transparent")
        top_line.pack(fill="x", padx=10, pady=5)

        options = ["Tezsiz", "Tezli", "Doktora", "1. Sınıf", "2. Sınıf", "3. Sınıf", "4. Sınıf", "Seçmeli", "Zorunlu",
                   "Özel..."]
        keyword_ent = ctk.CTkEntry(top_line, placeholder_text="Grup Adı", width=150)
        keyword_ent.insert(0, key)

        menu_val = key if key in options else "Özel..."
        group_menu = ctk.CTkOptionMenu(top_line, values=options, command=lambda c: (keyword_ent.delete(0, "end"),
                                                                                    keyword_ent.insert(0,
                                                                                                       c if c != "Özel..." else ""),
                                                                                    keyword_ent.focus() if c == "Özel..." else None),
                                       width=120)
        group_menu.set(menu_val);
        group_menu.pack(side="left", padx=5)
        keyword_ent.pack(side="left", padx=5)

        inverse_var = ctk.BooleanVar(value=is_inverse)
        ctk.CTkCheckBox(top_line, text="DIŞINDA", variable=inverse_var, width=40, font=("Arial", 12, "bold")).pack(
            side="left", padx=5)

        ctype_menu = ctk.CTkOptionMenu(top_line, values=["SADECE", "ASLA"], width=100)
        ctype_menu.set(t);
        ctype_menu.pack(side="left", padx=5)

        ctk.CTkButton(top_line, text="Sil", width=50, fg_color="#e74c3c",
                      command=lambda f=row_frame: self.remove_row(f)).pack(side="right")

        grid_container = ctk.CTkFrame(row_frame, fg_color="#f9f9f9", corner_radius=5)
        grid_container.pack(fill="x", padx=10, pady=5)

        for col_idx, slot in enumerate(self.ALL_SLOTS):
            ctk.CTkLabel(grid_container, text=slot.split('-')[0], font=("Arial", 10, "bold"), width=80).grid(row=0,
                                                                                                             column=col_idx + 1,
                                                                                                             pady=2)

        matrix_vars = {}
        matrix_widgets = []
        preset = DEFAULT_PRESETS.get(key, {})

        def update_matrix_style(*args):
            mode = ctype_menu.get()
            color = "#e74c3c" if mode == "ASLA" else "#27ae60"
            hover = "#c0392b" if mode == "ASLA" else "#2ecc71"
            text = "X" if mode == "ASLA" else ""
            for cb in matrix_widgets:
                cb.configure(fg_color=color, hover_color=hover, text=text)

        for row_idx, day in enumerate(self.ALL_DAYS):
            ctk.CTkLabel(grid_container, text=day[:3], font=("Arial", 10), width=40).grid(row=row_idx + 1, column=0,
                                                                                          padx=5)
            for col_idx, slot in enumerate(self.ALL_SLOTS):
                # Başlangıç seçimi
                is_sel = [day, slot] in initial_pairs if initial_pairs is not None else (
                            day in preset.get("days", []) and slot in preset.get("slots", []))

                v = ctk.BooleanVar(value=is_sel)
                cb = ctk.CTkCheckBox(grid_container, text="", variable=v, width=20, checkbox_width=18,
                                     checkbox_height=18)
                cb.grid(row=row_idx + 1, column=col_idx + 1, padx=2, pady=2)
                matrix_vars[(day, slot)] = v
                matrix_widgets.append(cb)

        ctype_menu.configure(command=lambda _: update_matrix_style())
        update_matrix_style()

        self.constraint_rows.append(
            {"frame": row_frame, "keyword": keyword_ent, "type": ctype_menu, "matrix_vars": matrix_vars,
             "inverse_var": inverse_var})

    def remove_row(self, frame):
        frame.destroy()
        self.constraint_rows = [r for r in self.constraint_rows if r["frame"] != frame]

    def open_log_popup(self):
        self.log_popup = ctk.CTkToplevel(self)
        self.log_popup.title("İşlem Raporu Takibi")
        self.log_popup.geometry("900x650")
        self.log_popup.attributes('-topmost', True)
        self.log_text_widget = ctk.CTkTextbox(self.log_popup, fg_color="#ffffff", text_color="#1a1a1a",
                                              font=("Courier New", 12))
        self.log_text_widget.pack(fill="both", expand=True, padx=10, pady=10)

    def check_queue(self):
        while not self.log_queue.empty():
            msg = self.log_queue.get()
            if hasattr(self, 'log_text_widget') and self.log_text_widget.winfo_exists():
                self.log_text_widget.insert("end", msg);
                self.log_text_widget.see("end")
        self.after(50, self.check_queue)

    def start_thread(self):
        self.open_log_popup()
        self.run_btn.configure(state="disabled", text="HESAPLANIYOR...")
        spec_cons_list = []
        for row in self.constraint_rows:
            k = row["keyword"].get().strip()
            if k:
                selected_pairs = [[d, s] for (d, s), v in row["matrix_vars"].items() if v.get()]
                spec_cons_list.append({
                    "keyword": "!" + k if row["inverse_var"].get() else k,
                    "type": row["type"].get(),
                    "selected_slots": selected_pairs
                })

        base_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.getcwd()
        db_path = os.path.join(base_dir, "okul.db")
        self.ayarlar = {
            "DAYS": self.ALL_DAYS, "SLOTS": self.ALL_SLOTS, "DB_PATH": db_path,
            "PREF_FILE": os.path.abspath(self.pref_file_ent.get()),
            "CONSTR_FILE": os.path.abspath(self.constr_file_ent.get()),
            "OUTPUT_FILE": os.path.abspath(self.output_file_ent.get()),
            "MAX_TRIALS": int(self.max_trials_ent.get() or 30),
            "TRIAL_TIMEOUT": int(self.timeout_ent.get() or 30),
            "SPECIAL_CONSTRAINTS": spec_cons_list,
            "MAX_DAYS_PER_LECTURER": int(self.max_days_ent.get() or 3),
            "MIN_SLOT_GAP": int(self.min_gap_ent.get() or 1) + 1
        }
        threading.Thread(target=self.run_logic, daemon=True).start()

    def run_logic(self):
        old_stdout = sys.stdout
        sys.stdout = ConsoleRedirector(self.log_queue)
        try:
            excel_yolu = os.path.abspath(self.dersler_excel_ent.get())
            if os.path.exists(excel_yolu): db.veritabanini_guncelle(excel_yolu)
            if ders.arayuzden_baslat(self.ayarlar):
                out_name = os.path.basename(self.ayarlar["OUTPUT_FILE"])
                exam_name = "isletme_sinav_takvimi.xlsx"
                import excel
                excel.tam_program_raporu(out_name, "ders_programi_tam_rapor.xlsx")
                excel.tam_program_raporu(exam_name, "sinav_takvimi_tam_rapor.xlsx")
                html.rapor_olustur(out_name, "ders_programi_takvim.html", "📅 Haftalık Ders Programı", "#1a73e8")
                html.rapor_olustur(exam_name, "sinav_takvimi_takvim.html", "✍️ Dönem Sonu Sınav Takvimi", "#d32f2f")
                htmlxv2.rapor_olustur_v2(out_name, "ders_programi_tablo.html",
                                         "📅 İktisadi İdari Bilimler Ders Programı", "#1a73e8")
                htmlxv2.rapor_olustur_v2(exam_name, "sinav_takvimi_tablo.html",
                                         "✍️ İktisadi İdari Bilimler Sınav Takvimi", "#d32f2f")
                print(f"\n✅ İŞLEM TAMAM: Seçili matris hücrelerine göre program hazırlandı.")
            else:
                print("\n⚠️ Çözüm bulunamadı.")
        except Exception as e:
            print(f"\n❌ SİSTEM HATASI: {str(e)}")
        finally:
            sys.stdout = old_stdout
            self.run_btn.configure(state="normal", text="HESAPLAMAYI BAŞLAT")


if __name__ == "__main__":
    app = App()
    app.mainloop()