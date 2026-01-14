import sqlite3
import pandas as pd
import os

def veritabanini_guncelle(excel_dosyasi_yolu):
    db_adi = os.path.join(os.getcwd(), "okul.db")
    if os.path.exists(db_adi):
        os.remove(db_adi)

    conn = sqlite3.connect(db_adi)
    cursor = conn.cursor()

    # Sadece gerekli iki tabloyu oluşturuyoruz
    cursor.execute("CREATE TABLE Derslikler (derslik_id INTEGER PRIMARY KEY AUTOINCREMENT, derslik_adi TEXT NOT NULL, kontenjan INTEGER)")
    cursor.execute("""
    CREATE TABLE OgretimUyeleriDersler (
        isim TEXT,
        ders_adi TEXT,
        sinif TEXT NOT NULL,
        kontenjan INTEGER,
        durum TEXT
    )
    """)
    conn.commit()

    try:
        df_derslikler = pd.read_excel(excel_dosyasi_yolu, sheet_name="Derslikler")
        df_uyeler_dersler = pd.read_excel(excel_dosyasi_yolu, sheet_name="OgretimUyeleriDersler")

        for _, row in df_derslikler.iterrows():
            cursor.execute("INSERT INTO Derslikler (derslik_adi, kontenjan) VALUES (?, ?)",
                           (row["Derslikler"].strip(), int(row["Kontenjan"])))

        for _, row in df_uyeler_dersler.iterrows():
            durum_bilgisi = str(row["Durum"]).strip() if pd.notna(row["Durum"]) else ""
            cursor.execute(
                "INSERT INTO OgretimUyeleriDersler (isim, ders_adi, sinif, kontenjan, durum) VALUES (?, ?, ?, ?, ?)",
                (str(row["OgretimUyesi"]).strip(), str(row["Ders"]).strip(), str(row["Sinif"]).strip(), int(row["Kontenjan"]), durum_bilgisi)
            )

        conn.commit()
        print(f"✅ Veritabanı tertemiz hale getirildi: {db_adi}")
    except Exception as e:
        print(f"❌ HATA: {e}")
    finally:
        conn.close()