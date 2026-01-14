import pandas as pd
import re
import os
import sys

def rapor_olustur(file_path='isletme_ders_programi.xlsx', output_name="ders_programi_takvim.html",
                  baslik="📅 İşletme Bölümü Haftalık Ders Programı", ana_renk="#1a73e8"):
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
        time_slots = [str(col).strip() for col in df.columns[1:] if not str(col).startswith('Unnamed')]
        gunler_sirali = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]

        for index, row in df.iterrows():
            gun_adi = str(row.iloc[0]).strip()
            if gun_adi not in gunler_sirali:
                continue

            for i, saat in enumerate(time_slots):
                hucre = row.iloc[i + 1]
                if pd.isna(hucre) or str(hucre).strip() == "":
                    continue

                dersler = str(hucre).split('\n')
                for ders_satiri in dersler:
                    ders_satiri = ders_satiri.strip()
                    if not ders_satiri: continue

                    match = re.search(r"(.*?):\s*(.*?)\s*\[(.*?)\]\s*\"(.*?)\"\s*-\s*(.*)", ders_satiri)
                    if match:
                        derslik = match.group(1).strip()
                        ders_adi = match.group(2).strip()
                        sinif_grubu = match.group(3).strip()
                        durum = match.group(4).strip()
                        hoca = match.group(5).strip()

                        split_char = '\n' if '\n' in sinif_grubu else ','
                        alt_siniflar = [s.strip() for s in sinif_grubu.split(split_char) if s.strip()]

                        for tek_sinif in alt_siniflar:
                            lessons_list.append({
                                'Gün': gun_adi, 'Saat': saat,
                                'Derslik': derslik,
                                'Ders': ders_adi,
                                'Sınıf': tek_sinif,
                                'Durum': durum,
                                'Hoca': hoca
                            })
                    else:
                        match_simple = re.search(r"(.*?):\s*(.*)\s*-\s*(.*)", ders_satiri)
                        if match_simple:
                            lessons_list.append({
                                'Gün': gun_adi, 'Saat': saat,
                                'Derslik': match_simple.group(1).strip(),
                                'Ders': match_simple.group(2).strip(),
                                'Sınıf': 'Genel',
                                'Durum': '-',
                                'Hoca': match_simple.group(3).strip()
                            })

        if not lessons_list:
            print(f"⚠️ Uyarı: {file_path} için işlenecek veri bulunamadı.")
            return

        final_df = pd.DataFrame(lessons_list)
        all_teachers = sorted(final_df['Hoca'].unique())
        all_classes = sorted(final_df['Sınıf'].unique())
        all_status = sorted(final_df['Durum'].unique())
        existing_days = [d for d in gunler_sirali if d in final_df['Gün'].unique()]
        existing_hours = sorted(final_df['Saat'].unique(), key=lambda x: x.split('-')[0])

        badge_bg = "#ffebee" if ana_renk == "#d32f2f" else "#e8f0fe"
        badge_text = "#c62828" if ana_renk == "#d32f2f" else "#1967d2"

        # 3. HTML Şablonu
        html_content = f"""
        <!DOCTYPE html>
        <html lang="tr">
        <head>
            <meta charset="UTF-8">
            <title>{baslik}</title>
            <style>
                body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f0f2f5; margin: 0; padding: 20px; }}
                .container {{ max-width: 1400px; margin: auto; background: white; padding: 20px; border-radius: 15px; box-shadow: 0 10px 30px rgba(0,0,0,0.1); }}
                h1 {{ text-align: center; color: {ana_renk}; margin-bottom: 20px; }}
                .filters {{ display: flex; gap: 15px; justify-content: center; background: #fff; padding: 20px; border-radius: 10px; margin-bottom: 20px; border: 1px solid #e0e0e0; position: sticky; top: 10px; z-index: 1000; }}
                .filter-group {{ display: flex; flex-direction: column; }}
                label {{ font-weight: 600; margin-bottom: 5px; font-size: 13px; color: #666; }}
                select {{ padding: 8px 12px; border-radius: 5px; border: 1px solid #ddd; min-width: 180px; }}
                .btn-export {{ background-color: #2e7d32; color: white; border: none; padding: 9px 15px; border-radius: 5px; cursor: pointer; font-weight: bold; font-size: 13px; margin-top: auto; }}
                .btn-export:hover {{ background-color: #1b5e20; }}
                table {{ width: 100%; border-collapse: collapse; table-layout: fixed; }}
                th {{ background-color: {ana_renk}; color: white; padding: 12px; border: 1px solid #ddd; font-size: 14px; }}
                td {{ border: 1px solid #e0e0e0; vertical-align: top; padding: 5px; background: #fafafa; height: 110px; }}
                .time-cell {{ background: #f8f9fa; font-weight: bold; text-align: center; width: 80px; color: #333; vertical-align: middle; height: auto; font-size: 13px; }}
                .lesson-card {{ background: #ffffff; border-left: 4px solid {ana_renk}; margin-bottom: 8px; padding: 8px; border-radius: 4px; box-shadow: 0 2px 5px rgba(0,0,0,0.05); font-size: 12px; }}
                .lesson-name {{ font-weight: bold; color: {ana_renk}; display: block; margin-bottom: 4px; border-bottom: 1px solid #eee; padding-bottom: 2px; }}
                .teacher-name {{ color: #2e7d32; font-weight: 600; font-style: italic; }}
                .durum-info {{ color: #666; font-size: 11px; display: block; margin-top: 2px; font-weight: 500; }}
                .class-badge {{ display: inline-block; background: {badge_bg}; color: {badge_text}; padding: 2px 6px; border-radius: 10px; font-size: 10px; margin-top: 4px; font-weight: bold; }}
                .btn-single-add {{ margin-top: 6px; display: inline-block; background: #4285F4; color: white; padding: 3px 6px; border-radius: 3px; font-size: 9px; border: none; cursor: pointer; }}
            </style>
        </head>
        <body>
        <div class="container">
            <h1>{baslik}</h1>
            <div class="filters">
                <div class="filter-group"><label>Öğretim Üyesi:</label><select id="teacherSelect" onchange="filterSchedule()"><option value="all">Tüm Öğretim Üyeleri</option>{"".join([f'<option value="{h}">{h}</option>' for h in all_teachers])}</select></div>
                <div class="filter-group"><label>Sınıf / Grup:</label><select id="classSelect" onchange="filterSchedule()"><option value="all">Tüm Sınıflar</option>{"".join([f'<option value="{c}">{c}</option>' for c in all_classes])}</select></div>
                <div class="filter-group"><label>Durum:</label><select id="statusSelect" onchange="filterSchedule()"><option value="all">Tüm Durumlar</option>{"".join([f'<option value="{s}">{s}</option>' for s in all_status])}</select></div>
                <div class="filter-group"><label>&nbsp;</label><button class="btn-export" onclick="downloadICS()">📥 Seçili Programı İndir (.ics)</button></div>
            </div>

            <table id="scheduleTable">
                <thead><tr><th style="width: 80px;">Saat</th>{"".join([f"<th>{gun}</th>" for gun in existing_days])}</tr></thead>
                <tbody>
        """

        for saat in existing_hours:
            html_content += f"<tr><td class='time-cell'>{saat}</td>"
            for gun in existing_days:
                html_content += "<td>"
                matching_lessons = final_df[(final_df['Gün'] == gun) & (final_df['Saat'] == saat)]
                for _, lesson in matching_lessons.iterrows():
                    safe_title = lesson['Ders'].replace("'", "\\'")
                    js_data = f"{{title:'{safe_title}', teacher:'{lesson['Hoca']}', room:'{lesson['Derslik']}', day:'{lesson['Gün']}', hour:'{lesson['Saat']}', classInfo:'{lesson['Sınıf']}', status:'{lesson['Durum']}'}}"
                    html_content += f"""
                    <div class="lesson-card" data-teacher="{lesson['Hoca']}" data-class="{lesson['Sınıf']}" data-day="{lesson['Gün']}" data-hour="{lesson['Saat']}" data-room="{lesson['Derslik']}" data-status="{lesson['Durum']}">
                        <span class="lesson-name">{lesson['Ders']}</span>
                        <div class="lesson-info">
                            📍 {lesson['Derslik']}<br>
                            👨‍🏫 <span class="teacher-name">{lesson['Hoca']}</span><br>
                            <span class="durum-info">📄 {lesson['Durum']}</span>
                            <span class="class-badge">{lesson['Sınıf']}</span>
                        </div>
                        <button class="btn-single-add" onclick="addSingleToGoogle({js_data})">➕ Takvime Ekle</button>
                    </div>
                    """
                html_content += "</td>"
            html_content += "</tr>"

        html_content += """
                </tbody></table></div>
        <script>
            const gunMap = {'Pazartesi': '20260209', 'Salı': '20260210', 'Çarşamba': '20260211', 'Perşembe': '20260212', 'Cuma': '20260213', 'Cumartesi': '20260214', 'Pazar': '20260215'};

            function filterSchedule() {
                const t = document.getElementById("teacherSelect").value;
                const c = document.getElementById("classSelect").value;
                const s = document.getElementById("statusSelect").value;
                document.querySelectorAll(".lesson-card").forEach(card => {
                    const tM = (t === "all" || card.getAttribute("data-teacher") === t);
                    const cM = (c === "all" || card.getAttribute("data-class") === c);
                    const sM = (s === "all" || card.getAttribute("data-status") === s);
                    card.style.display = (tM && cM && sM) ? "block" : "none";
                });
            }

            function addSingleToGoogle(d) {
                let h = d.hour.split('-')[0].trim().replace(":","");
                if(h.length === 3) h = "0" + h;
                const start = gunMap[d.day] + "T" + h + "00";
                const end = gunMap[d.day] + "T" + (parseInt(h.substring(0,2))+1).toString().padStart(2, '0') + h.substring(2) + "00";
                const fullTitle = d.title + " [" + d.classInfo + "]";
                const details = `Hoca: ${d.teacher}\\nSınıf: ${d.classInfo}\\nDurum: ${d.status}`;
                const url = `https://www.google.com/calendar/render?action=TEMPLATE&text=${encodeURIComponent(fullTitle)}&details=${encodeURIComponent(details)}&location=${encodeURIComponent(d.room)}&dates=${start}/${end}&recur=RRULE:FREQ=WEEKLY;UNTIL=20260615T235959Z`;
                window.open(url, '_blank');
            }

            function downloadICS() {
                const tv = document.getElementById("teacherSelect").value;
                const cv = document.getElementById("classSelect").value;
                const sv = document.getElementById("statusSelect").value;
                if (tv === "all" && cv === "all" && sv === "all") {
                    alert("⚠️ Lütfen önce bir filtreleme yapın.");
                    return;
                }
                const visible = Array.from(document.querySelectorAll(".lesson-card")).filter(c => c.style.display !== "none");
                let ics = "BEGIN:VCALENDAR\\nVERSION:2.0\\nPRODID:-//Ege//TR\\nCALSCALE:GREGORIAN\\nMETHOD:PUBLISH\\n";
                visible.forEach(card => {
                    const title = card.querySelector(".lesson-name").innerText;
                    const hoca = card.getAttribute("data-teacher");
                    const classInfo = card.getAttribute("data-class");
                    const status = card.getAttribute("data-status");
                    const room = card.getAttribute("data-room");
                    const day = card.getAttribute("data-day");
                    let h = card.getAttribute("data-hour").split('-')[0].trim().replace(":","");
                    if(h.length === 3) h = "0" + h;
                    ics += "BEGIN:VEVENT\\n";
                    ics += `SUMMARY:${title} [${classInfo}]\\n`;
                    ics += `LOCATION:${room}\\n`;
                    ics += `DESCRIPTION:Hoca: ${hoca}\\nSınıf: ${classInfo}\\nDurum: ${status}\\n`;
                    ics += `DTSTART:${gunMap[day]}T${h}00\\n`;
                    ics += `DTEND:${gunMap[day]}T${(parseInt(h.substring(0,2))+1).toString().padStart(2, '0')}${h.substring(2)}00\\n`;
                    ics += `RRULE:FREQ=WEEKLY;UNTIL=20260615T235959Z\\n`;
                    ics += "END:VEVENT\\n";
                });
                ics += "END:VCALENDAR";
                const blob = new Blob([ics.replace(/\\\\n/g, "\\n")], { type: 'text/calendar;charset=utf-8' });
                const a = document.createElement("a");
                a.href = window.URL.createObjectURL(blob);
                a.download = "ders_programi.ics";
                a.click();
            }
        </script>
        </body></html>
        """

        with open(os.path.join(base_dir, output_name), "w", encoding="utf-8") as f:
            f.write(html_content)

        print(f"✅ Takvim formatlı HTML başarıyla oluşturuldu: {output_name}")

    except Exception as e:
        print(f"❌ Hata: {e}")

if __name__ == "__main__":
    rapor_olustur(file_path='isletme_ders_programi.xlsx', output_name="ders_programi_takvim.html", baslik="📅 İşletme Bölümü Haftalık Ders Programı", ana_renk="#1a73e8")
    rapor_olustur(file_path='isletme_sinav_takvimi.xlsx', output_name="sinav_takvimi_takvim.html", baslik="📝 İşletme Bölümü Sınav Takvimi", ana_renk="#d32f2f")