import pandas as pd
import os
import sys

def rapor_olustur(excel_dosya="isletme_ders_programi.xlsx"):
    """ main.py üzerinden çağrılacak olan HTML oluşturma fonksiyonu """

    try:
        # EXE veya script konumunu belirle
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.abspath(os.path.dirname(__file__))

        # Excel dosyasının tam yolunu oluştur
        excel_yolu = os.path.join(base_dir, excel_dosya)

        if not os.path.exists(excel_yolu):
            print(f"❌ Hata: {excel_yolu} bulunamadı. HTML oluşturulamıyor.")
            return

        # 1. EXCEL DOSYASINI OKU
        df = pd.read_excel(excel_yolu, index_col=0)

        # 2. \n -> <br> DÖNÜŞÜMÜ
        df = df.astype(str).map(lambda x: x.replace("\n", "<br>") if x != "nan" else "")

        # 3. HTML TABLOYA DÖNÜŞTÜR
        html_tablo = df.to_html(
            border=1,
            justify="center",
            escape=False  # <br> etiketlerinin çalışması için
        )

        # 4. HTML ŞABLONU
        html_sablon = f"""
        <!DOCTYPE html>
        <html lang="tr">
        <head>
            <meta charset="UTF-8">
            <title>Ders Programı</title>
            <style>
                body {{ font-family: Arial, sans-serif; padding: 20px; }}
                table {{ border-collapse: collapse; width: 100%; }}
                th, td {{ border: 1px solid #333; padding: 8px; text-align: center; vertical-align: top; }}
                th {{ background-color: #f2f2f2; }}
            </style>
        </head>
        <body>
            <h2>Ders Programı</h2>
            {html_tablo}
        </body>
        </html>
        """

        # 5. HTML DOSYASINI EXE YANINA YAZ
        cikti_adi = os.path.join(base_dir, "ders_programi.html")
        with open(cikti_adi, "w", encoding="utf-8") as f:
            f.write(html_sablon)

        print(f"✅ HTML dosyası başarıyla oluşturuldu: {cikti_adi}")

    except Exception as e:
        print(f"❌ html.py hatası: {e}")

if __name__ == "__main__":
    rapor_olustur()