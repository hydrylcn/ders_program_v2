import pandas as pd
import re
import os
import sys

def rapor_olustur_v2(file_path='isletme_ders_programi.xlsx'):
    """ main.py üzerinden çağrılacak olan dinamik tablo fonksiyonu """

    try:
        # --- EXE ve YOL DÜZELTMESİ ---
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.abspath(os.path.dirname(__file__))

        excel_yolu = os.path.join(base_dir, file_path)
        # -----------------------------

        if not os.path.exists(excel_yolu):
            print(f"❌ Hata: {excel_yolu} bulunamadı. Dinamik tablo oluşturulamıyor.")
            return

        # 1. Excel dosyasını oku
        df = pd.read_excel(excel_yolu)

        # 2. Veriyi işle ve liste haline getir
        lessons_list = []
        time_slots = df.columns[1:]

        for index, row in df.iterrows():
            gun_adi = str(row.iloc[0]).strip()
            if gun_adi == "nan" or not gun_adi: continue

            for saat in time_slots:
                hucre = row[saat]
                if pd.isna(hucre) or str(hucre).strip() == "":
                    continue

                dersler = str(hucre).split('\n')

                for ders_satiri in dersler:
                    ders_satiri = ders_satiri.strip()
                    if not ders_satiri: continue

                    match = re.search(r"(.*?):\s*(.*?)\s*\[(.*?)\]\s*-\s*(.*)", ders_satiri)

                    if match:
                        lessons_list.append({
                            'Gün': gun_adi, 'Saat': saat,
                            'Derslik': match.group(1).strip(),
                            'Ders': match.group(2).strip(),
                            'Sınıf': match.group(3).strip(),
                            'Hoca': match.group(4).strip()
                        })
                    else:
                        match_simple = re.search(r"(.*?):\s*(.*)\s*-\s*(.*)", ders_satiri)
                        if match_simple:
                            lessons_list.append({
                                'Gün': gun_adi, 'Saat': saat,
                                'Derslik': match_simple.group(1).strip(),
                                'Ders': match_simple.group(2).strip(),
                                'Sınıf': 'Belirtilmemiş',
                                'Hoca': match_simple.group(3).strip()
                            })
                        else:
                            lessons_list.append({
                                'Gün': gun_adi, 'Saat': saat, 'Derslik': '-',
                                'Ders': ders_satiri, 'Sınıf': 'Belirtilmemiş', 'Hoca': 'Belirtilmemiş'
                            })

        if not lessons_list:
            print("⚠️ Uyarı: İşlenecek ders verisi bulunamadı.")
            return

        final_df = pd.DataFrame(lessons_list)

        all_teachers = sorted(final_df['Hoca'].unique())
        all_classes = sorted(final_df['Sınıf'].unique())
        all_days = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]
        existing_days = [d for d in all_days if d in final_df['Gün'].unique()]

        # 3. HTML ve JavaScript İçeriği (Şablonun aynı kalıyor)
        html_template = f"""
        <!DOCTYPE html>
        <html lang="tr">
        <head>
            <meta charset="UTF-8">
            <title>Dinamik Ders Programı</title>
            <style>
                body {{ font-family: 'Segoe UI', sans-serif; background-color: #f4f7f6; margin: 0; padding: 20px; }}
                .container {{ max-width: 1200px; margin: auto; background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.08); }}
                h1 {{ text-align: center; color: #1a73e8; margin-bottom: 30px; }}
                .filters {{ display: flex; gap: 15px; justify-content: center; background: #f8f9fa; padding: 20px; border-radius: 10px; margin-bottom: 30px; flex-wrap: wrap; border: 1px solid #eee; }}
                .filter-group {{ display: flex; flex-direction: column; }}
                label {{ font-weight: bold; margin-bottom: 8px; color: #555; font-size: 14px; }}
                select {{ padding: 10px; width: 250px; border-radius: 6px; border: 1px solid #ccc; font-size: 15px; background: white; cursor: pointer; outline: none; }}
                table {{ width: 100%; border-collapse: collapse; }}
                th {{ background-color: #1a73e8; color: white; padding: 15px; text-align: left; position: sticky; top: 0; }}
                td {{ padding: 12px; border-bottom: 1px solid #eee; font-size: 14px; }}
                tr:hover {{ background-color: #f1f7fd; }}
                .badge-sinif {{ background: #e3f2fd; color: #1565c0; padding: 4px 10px; border-radius: 15px; font-size: 12px; font-weight: bold; }}
                .hoca-adi {{ color: #2e7d32; font-weight: bold; }}
            </style>
        </head>
        <body>
        <div class="container">
            <h1>📅 İktisadi İdari Bilimler Ders Programı</h1>
            <div class="filters">
                <div class="filter-group"><label>Gün Seçin:</label><select id="daySelect" onchange="filterTable()"><option value="all">Tüm Günler</option>{''.join([f'<option value="{d}">{d}</option>' for d in existing_days])}</select></div>
                <div class="filter-group"><label>Öğretim Elemanı Seçin:</label><select id="teacherSelect" onchange="filterTable()"><option value="all">Tüm Hocalar</option>{''.join([f'<option value="{h}">{h}</option>' for h in all_teachers])}</select></div>
                <div class="filter-group"><label>Sınıf / Grup Seçin:</label><select id="classSelect" onchange="filterTable()"><option value="all">Tüm Sınıflar</option>{''.join([f'<option value="{c}">{c}</option>' for c in all_classes])}</select></div>
            </div>
            <table id="programTable">
                <thead><tr><th>Gün</th><th>Saat</th><th>Derslik</th><th>Ders Adı</th><th>Sınıf / Grup</th><th>Öğretim Elemanı</th></tr></thead>
                <tbody>
        """

        for _, row in final_df.iterrows():
            html_template += f"""
                    <tr data-day="{row['Gün']}" data-teacher="{row['Hoca']}" data-class="{row['Sınıf']}">
                        <td>{row['Gün']}</td><td>{row['Saat']}</td><td>{row['Derslik']}</td><td>{row['Ders']}</td><td><span class="badge-sinif">{row['Sınıf']}</span></td><td class="hoca-adi">{row['Hoca']}</td>
                    </tr>"""

        html_template += """
                </tbody></table></div>
        <script>
            function filterTable() {
                const dayVal = document.getElementById("daySelect").value;
                const teacherVal = document.getElementById("teacherSelect").value;
                const classVal = document.getElementById("classSelect").value;
                const rows = document.querySelectorAll("#programTable tbody tr");
                rows.forEach(row => {
                    const dMatch = (dayVal === "all" || row.getAttribute("data-day") === dayVal);
                    const tMatch = (teacherVal === "all" || row.getAttribute("data-teacher") === teacherVal);
                    const cMatch = (classVal === "all" || row.getAttribute("data-class") === classVal);
                    row.style.display = (dMatch && tMatch && cMatch) ? "" : "none";
                });
            }
        </script>
        </body></html>
        """

        # 4. HTML dosyasını EXE yanına kaydet
        cikti_adi = os.path.join(base_dir, "ders_programi_tablo.html")
        with open(cikti_adi, "w", encoding="utf-8") as f:
            f.write(html_template)

        print(f"✅ Filtreli HTML başarıyla oluşturuldu: {cikti_adi}")

    except Exception as e:
        print(f"❌ htmlxv2.py hatası: {e}")

if __name__ == "__main__":
    rapor_olustur_v2()