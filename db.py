import sqlite3
import pandas as pd
import os


def veritabanini_guncelle(excel_dosyasi_yolu):
    """
    Excel'deki yeni 'Durum' sütununu destekleyen güncellenmiş fonksiyon.
    """

    db_adi = os.path.join(os.getcwd(), "okul.db")

    # 0. ESKİ VERİTABANINI SİL
    if os.path.exists(db_adi):
        try:
            os.remove(db_adi)
        except Exception as e:
            print(f"Uyarı: Eski veritabanı silinemedi. Hata: {e}")

    conn = sqlite3.connect(db_adi)
    cursor = conn.cursor()

    # 1. TABLOLARI OLUŞTUR (DURUM sütunu eklendi)
    cursor.execute("CREATE TABLE OgretimUyeleri (uye_id INTEGER PRIMARY KEY AUTOINCREMENT, isim TEXT NOT NULL)")
    cursor.execute(
        "CREATE TABLE Dersler (ders_id INTEGER PRIMARY KEY AUTOINCREMENT, ders_adi TEXT NOT NULL, sinif TEXT NOT NULL)")
    cursor.execute(
        "CREATE TABLE Derslikler (derslik_id INTEGER PRIMARY KEY AUTOINCREMENT, derslik_adi TEXT NOT NULL, kontenjan INTEGER)")

    # OgretimUyeleriDersler tablosuna 'durum' eklendi
    cursor.execute("""
    CREATE TABLE OgretimUyeleriDersler (
        uye_id INTEGER,
        ders_id INTEGER,
        sinif TEXT NOT NULL,
        kontenjan INTEGER,
        durum TEXT, 
        FOREIGN KEY (uye_id) REFERENCES OgretimUyeleri(uye_id),
        FOREIGN KEY (ders_id) REFERENCES Dersler(ders_id)
    )
    """)
    conn.commit()

    # 2. VERİ EKLEME İŞLEMİ
    try:
        df_uyeler = pd.read_excel(excel_dosyasi_yolu, sheet_name="OgretimUyeleri")
        df_dersler = pd.read_excel(excel_dosyasi_yolu, sheet_name="Dersler")
        df_derslikler = pd.read_excel(excel_dosyasi_yolu, sheet_name="Derslikler")
        df_uyeler_dersler = pd.read_excel(excel_dosyasi_yolu, sheet_name="OgretimUyeleriDersler")

        # Ana Tabloların Doldurulması
        for isim in df_uyeler["OgretimUyesi"]:
            cursor.execute("INSERT INTO OgretimUyeleri (isim) VALUES (?)", (isim.strip(),))

        for _, row in df_dersler.iterrows():
            cursor.execute("INSERT INTO Dersler (ders_adi, sinif) VALUES (?, ?)",
                           (row["Dersler"].strip(), row["Sinif"].strip()))

        for _, row in df_derslikler.iterrows():
            cursor.execute("INSERT INTO Derslikler (derslik_adi, kontenjan) VALUES (?, ?)",
                           (row["Derslikler"].strip(), int(row["Kontenjan"])))
        conn.commit()

        # ID Eşleştirme Sözlükleri
        cursor.execute("SELECT uye_id, isim FROM OgretimUyeleri")
        uye_adi_to_id = {isim.strip().upper(): uye_id for uye_id, isim in cursor.fetchall()}
        cursor.execute("SELECT ders_id, ders_adi FROM Dersler")
        ders_adi_to_id = {ders_adi.strip().upper(): ders_id for ders_id, ders_adi in cursor.fetchall()}

        # Bağlantı Tablosunun Doldurulması (DURUM sütunu burada işleniyor)
        for _, row in df_uyeler_dersler.iterrows():
            uye_id = uye_adi_to_id[row["OgretimUyesi"].strip().upper()]
            ders_id = ders_adi_to_id[row["Ders"].strip().upper()]

            # Excel'den gelen 'Durum' sütununu alıyoruz
            durum_bilgisi = str(row["Durum"]).strip() if pd.notna(row["Durum"]) else ""

            cursor.execute(
                "INSERT INTO OgretimUyeleriDersler (uye_id, ders_id, sinif, kontenjan, durum) VALUES (?, ?, ?, ?, ?)",
                (uye_id, ders_id, str(row["Sinif"]).strip(), int(row["Kontenjan"]), durum_bilgisi)
            )

        conn.commit()
        print(f"✅ Veritabanı başarıyla güncellendi: {db_adi}")

    except Exception as e:
        print(f"❌ HATA: {e}")
    finally:
        conn.close()


if __name__ == "__main__":
    # Dosya isminin dersler.xlsx olduğundan emin ol
    veritabanini_guncelle("dersler.xlsx")