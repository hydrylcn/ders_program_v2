import pandas as pd
import re
import os
import sys


def rapor_olustur(file_path='isletme_ders_programi.xlsx'):
    try:
        # --- DOSYA YOLU AYARLARI ---
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.abspath(os.path.dirname(__file__))

        excel_yolu = os.path.join(base_dir, file_path)

        if not os.path.exists(excel_yolu):
            print(f"❌ Hata: {excel_yolu} bulunamadı.")
            return

        # 1. Excel'i Oku
        df = pd.read_excel(excel_yolu)

        # 2. Veriyi Ayrıştır
        lessons_list = []
        time_slots = [col for col in df.columns[1:] if not str(col).startswith('Unnamed')]
        gunler_sirali = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]

        for index, row in df.iterrows():
            gun_adi = str(row.iloc[0]).strip()
            if gun_adi == "nan" or not gun_adi or gun_adi not in gunler_sirali:
                continue

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
                                'Sınıf': 'Genel',
                                'Hoca': match_simple.group(3).strip()
                            })

        if not lessons_list:
            print("⚠️ Uyarı: İşlenecek veri bulunamadı.")
            return

        final_df = pd.DataFrame(lessons_list)

        all_teachers = sorted(final_df['Hoca'].unique())
        all_classes = sorted(final_df['Sınıf'].unique())
        existing_days = [d for d in gunler_sirali if d in final_df['Gün'].unique()]
        existing_hours = sorted(final_df['Saat'].unique())

        # 3. HTML Şablonu (Güncellenmiş CSS ile)
        html_content = f"""
        <!DOCTYPE html>
        <html lang="tr">
        <head>
            <meta charset="UTF-8">
            <title>Haftalık Ders Programı - Takvim</title>
            <style>
                body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f0f2f5; margin: 0; padding: 20px; }}
                .container {{ max-width: 1400px; margin: auto; background: white; padding: 20px; border-radius: 15px; box-shadow: 0 10px 30px rgba(0,0,0,0.1); }}
                h1 {{ text-align: center; color: #1a73e8; margin-bottom: 20px; }}

                .filters {{ display: flex; gap: 15px; justify-content: center; background: #fff; padding: 20px; border-radius: 10px; margin-bottom: 20px; border: 1px solid #e0e0e0; position: sticky; top: 10px; z-index: 1000; }}
                .filter-group {{ display: flex; flex-direction: column; }}
                label {{ font-weight: 600; margin-bottom: 5px; font-size: 13px; color: #666; }}
                select {{ padding: 8px 12px; border-radius: 5px; border: 1px solid #ddd; min-width: 200px; }}

                table {{ width: 100%; border-collapse: collapse; table-layout: fixed; }}
                th {{ background-color: #1a73e8; color: white; padding: 12px; border: 1px solid #1565c0; font-size: 14px; }}

                /* DEĞİŞİKLİK BURADA: td yüksekliği sabitlendi */
                td {{ border: 1px solid #e0e0e0; vertical-align: top; padding: 5px; background: #fafafa; height: 110px; }}

                .time-cell {{ background: #f8f9fa; font-weight: bold; text-align: center; width: 80px; color: #333; vertical-align: middle; height: auto; }}

                .lesson-card {{
                    background: #ffffff;
                    border-left: 4px solid #1a73e8;
                    margin-bottom: 8px;
                    padding: 8px;
                    border-radius: 4px;
                    box-shadow: 0 2px 5px rgba(0,0,0,0.05);
                    font-size: 12px;
                    transition: transform 0.2s;
                }}
                .lesson-card:hover {{ transform: scale(1.02); box-shadow: 0 4px 8px rgba(0,0,0,0.1); }}
                .lesson-name {{ font-weight: bold; color: #1a73e8; display: block; margin-bottom: 4px; border-bottom: 1px solid #eee; padding-bottom: 2px; }}
                .lesson-info {{ color: #555; line-height: 1.4; }}
                .teacher-name {{ color: #2e7d32; font-weight: 600; font-style: italic; }}
                .class-badge {{ display: inline-block; background: #e8f0fe; color: #1967d2; padding: 2px 6px; border-radius: 10px; font-size: 10px; margin-top: 4px; font-weight: bold; }}
            </style>
        </head>
        <body>
        <div class="container">
            <h1>📅 İşletme Bölümü Haftalık Ders Programı</h1>

            <div class="filters">
                <div class="filter-group">
                    <label>Öğretim Elemanı:</label>
                    <select id="teacherSelect" onchange="filterSchedule()">
                        <option value="all">Tüm Hocalar</option>
                        {"".join([f'<option value="{h}">{h}</option>' for h in all_teachers])}
                    </select>
                </div>
                <div class="filter-group">
                    <label>Sınıf / Grup:</label>
                    <select id="classSelect" onchange="filterSchedule()">
                        <option value="all">Tüm Sınıflar</option>
                        {"".join([f'<option value="{c}">{c}</option>' for c in all_classes])}
                    </select>
                </div>
            </div>

            <table id="scheduleTable">
                <thead>
                    <tr>
                        <th style="width: 80px;">Saat</th>
                        {"".join([f"<th>{gun}</th>" for gun in existing_days])}
                    </tr>
                </thead>
                <tbody>
        """

        for saat in existing_hours:
            html_content += f"<tr><td class='time-cell'>{saat}</td>"
            for gun in existing_days:
                html_content += "<td>"
                matching_lessons = final_df[(final_df['Gün'] == gun) & (final_df['Saat'] == saat)]
                for _, lesson in matching_lessons.iterrows():
                    html_content += f"""
                    <div class="lesson-card" data-teacher="{lesson['Hoca']}" data-class="{lesson['Sınıf']}">
                        <span class="lesson-name">{lesson['Ders']}</span>
                        <div class="lesson-info">
                            📍 {lesson['Derslik']}<br>
                            👨‍🏫 <span class="teacher-name">{lesson['Hoca']}</span><br>
                            <span class="class-badge">{lesson['Sınıf']}</span>
                        </div>
                    </div>
                    """
                html_content += "</td>"
            html_content += "</tr>"

        html_content += """
                </tbody>
            </table>
        </div>

        <script>
            function filterSchedule() {
                const teacherVal = document.getElementById("teacherSelect").value;
                const classVal = document.getElementById("classSelect").value;
                const cards = document.querySelectorAll(".lesson-card");

                cards.forEach(card => {
                    const tMatch = (teacherVal === "all" || card.getAttribute("data-teacher") === teacherVal);
                    const cMatch = (classVal === "all" || card.getAttribute("data-class") === classVal);
                    card.style.display = (tMatch && cMatch) ? "block" : "none";
                });
            }
        </script>
        </body>
        </html>
        """

        cikti_adi = os.path.join(base_dir, "ders_programi_takvim.html")
        with open(cikti_adi, "w", encoding="utf-8") as f:
            f.write(html_content)

        print(f"✅ Takvim formatlı HTML başarıyla oluşturuldu: {cikti_adi}")

    except Exception as e:
        print(f"❌ Hata oluştu: {e}")


if __name__ == "__main__":
    rapor_olustur_takvim()