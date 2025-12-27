import tkinter as tk
import subprocess
import sys

def calistir_db():
    # db.py dosyasını çalıştırır
    subprocess.run([sys.executable, "db.py"])

def calistir_ders():
    # dosyasını çalıştırır
    subprocess.run([sys.executable, "dersxv4.py"])

# Ana pencereyi oluştur
pencere = tk.Tk()
pencere.title("Ana Menü")
pencere.geometry("300x150")

# Veritabanı butonu
btn_db = tk.Button(pencere, text="Veritabanı", command=calistir_db)
btn_db.pack(pady=20)

# Ders Programı butonu
btn_ders = tk.Button(pencere, text="Ders Programı", command=calistir_ders)
btn_ders.pack(pady=20)

# Pencereyi göster
pencere.mainloop()
