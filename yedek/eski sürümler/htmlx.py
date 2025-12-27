import pandas as pd
import re

# 1. Excel dosyasını oku
file_path = 'isletme_ders_programi.xlsx'

try:
    df = pd.read_excel(file_path)
except Exception as e:
    print(f"Hata: {e}")
    exit()

# 2. Veriyi işle
lessons_list = []
time_slots = df.columns[1:]

for index, row in df.iterrows():
    gun_adi = str(row.iloc[0]).strip()
    if gun_adi == "nan": continue

    for saat in time_slots:
        hucre = row[saat]
        if pd.isna(hucre) or str(hucre).strip() == "":
            continue

        dersler = str(hucre).split('\n')

        for ders_satiri in dersler:
            ders_satiri = ders_satiri.strip()
            if not ders_satiri: continue

            # Regex: "Derslik (X): (Ders Adı) (Sınıf Bilgisi) - (Hoca)"
            # Örnek: Derslik 103: Enformetri (3. Sınıf) - Doç. Dr. Haydar YALÇIN
            match = re.search(r"(.*?):\s*(.*?)\s*\((.*?)\)\s*-\s*(.*)", ders_satiri)

            if match:
                lessons_list.append({
                    'Gün': gun_adi,
                    'Saat': saat,
                    'Derslik': match.group(1).strip(),
                    'Ders': match.group(2).strip(),
                    'Sınıf': match.group(3).strip(),
                    'Hoca': match.group(4).strip()
                })
            else:
                # Eğer parantez yoksa ama yine de format uygunsa (Fallback)
                match_simple = re.search(r"(.*?):\s*(.*)\s*-\s*(.*)", ders_satiri)
                if match_simple:
                    lessons_list.append({
                        'Gün': gun_adi, 'Saat': saat, 'Derslik': match_simple.group(1).strip(),
                        'Ders': match_simple.group(2).strip(), 'Sınıf': '-', 'Hoca': match_simple.group(3).strip()
                    })
                else:
                    lessons_list.append({
                        'Gün': gun_adi, 'Saat': saat, 'Derslik': '-', 'Ders': ders_satiri, 'Sınıf': '-',
                        'Hoca': 'Belirtilmemiş'
                    })

# 3. HTML Oluşturma
final_df = pd.DataFrame(lessons_list)
all_teachers = sorted(final_df['Hoca'].unique())

html_template = f"""
<!DOCTYPE html>
<html lang="tr">
<head>
    <meta charset="UTF-8">
    <title>Detaylı Ders Programı</title>
    <style>
        body {{ font-family: 'Segoe UI', Tahoma, sans-serif; background: #f0f2f5; padding: 20px; }}
        .container {{ max-width: 1200px; margin: auto; background: white; padding: 25px; border-radius: 12px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); }}
        h2 {{ text-align: center; color: #1a73e8; }}
        .filter-box {{ margin-bottom: 25px; text-align: center; background: #f8f9fa; padding: 15px; border-radius: 8px; }}
        select {{ padding: 12px; width: 400px; border-radius: 6px; border: 1px solid #ddd; font-size: 16px; cursor: pointer; }}
        table {{ width: 100%; border-collapse: collapse; }}
        th {{ background: #1a73e8; color: white; padding: 12px; text-align: left; }}
        td {{ padding: 10px; border-bottom: 1px solid #eee; font-size: 14px; color: #333; }}
        tr:hover {{ background: #e8f0fe; }}
        .badge-sinif {{ background: #e1f5fe; color: #01579b; padding: 4px 8px; border-radius: 4px; font-weight: bold; font-size: 12px; }}
        .hoca-adi {{ color: #2e7d32; font-weight: bold; }}
    </style>
</head>
<body>
    <div class="container">
        <h2>📊 Bölüm Ders Programı (Sınıf Detaylı)</h2>
        <div class="filter-box">
            <label><strong>Öğretim Elemanına Göre Filtrele:</strong></label><br><br>
            <select id="teacherSelect" onchange="filterTable()">
                <option value="all">Tüm Öğretmenler</option>
                {''.join([f'<option value="{h}">{h}</option>' for h in all_teachers])}
            </select>
        </div>
        <table id="programTable">
            <thead>
                <tr>
                    <th>Gün</th>
                    <th>Saat</th>
                    <th>Derslik</th>
                    <th>Ders</th>
                    <th>Sınıf</th>
                    <th>Hoca</th>
                </tr>
            </thead>
            <tbody>
"""

for _, row in final_df.iterrows():
    html_template += f"""
                <tr data-teacher="{row['Hoca']}">
                    <td>{row['Gün']}</td>
                    <td>{row['Saat']}</td>
                    <td>{row['Derslik']}</td>
                    <td>{row['Ders']}</td>
                    <td><span class="badge-sinif">{row['Sınıf']}</span></td>
                    <td class="hoca-adi">{row['Hoca']}</td>
                </tr>"""

html_template += """
            </tbody>
        </table>
    </div>
    <script>
        function filterTable() {
            var val = document.getElementById("teacherSelect").value;
            var rows = document.querySelectorAll("#programTable tbody tr");
            rows.forEach(r => {
                r.style.display = (val === "all" || r.getAttribute("data-teacher") === val) ? "" : "none";
            });
        }
    </script>
</body>
</html>
"""

with open("ders_programi.html", "w", encoding="utf-8") as f:
    f.write(html_template)

print("İşlem başarıyla tamamlandı. 'ders_programi.html' dosyasını kontrol edebilirsiniz.")