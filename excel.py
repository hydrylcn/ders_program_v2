import pandas as pd
import re
import os


def tam_program_raporu(input_file='ders_programi.xlsx', output_file='ders_programi.xlsx'):
    if not os.path.exists(input_file):
        print(f"❌ Hata: {input_file} bulunamadı.")
        return

    # 1. VERİ HAZIRLIĞI VE AYRIŞTIRMA
    df_raw = pd.read_excel(input_file)
    time_slots = [col for col in df_raw.columns[1:] if not str(col).startswith('Unnamed')]
    days = df_raw.iloc[:, 0].dropna().unique().tolist()

    lessons_list = []
    for _, row in df_raw.iterrows():
        gun_adi = str(row.iloc[0]).strip()
        for saat in time_slots:
            hucre = row[saat]
            if pd.notna(hucre) and str(hucre).strip() != "":
                dersler = str(hucre).split('\n')
                for ders_satiri in dersler:
                    match = re.search(r"(.*?):\s*(.*?)\s*\[(.*?)\]\s*\"(.*?)\"\s*-\s*(.*)", ders_satiri.strip())
                    if match:
                        lessons_list.append({
                            'Gün': gun_adi, 'Saat': saat, 'Derslik': match.group(1).strip(),
                            'Ders Adı': match.group(2).strip(), 'Sınıf': match.group(3).strip(),
                            'Durum': match.group(4).strip(), 'Öğretim Üyesi': match.group(5).strip()
                        })

    # 2. EXCEL YAZICI AYARLARI
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    workbook = writer.book

    # --- FORMATLAR ---
    header_fmt = workbook.add_format(
        {'bold': True, 'bg_color': '#1a73e8', 'font_color': 'white', 'border': 1, 'align': 'center',
         'valign': 'vcenter'})
    list_cell_fmt = workbook.add_format({'border': 1, 'valign': 'vcenter', 'font_size': 10})
    box_fmt = workbook.add_format(
        {'text_wrap': True, 'valign': 'top', 'align': 'left', 'border': 1, 'bg_color': '#e8f0fe', 'font_size': 9})
    day_fmt = workbook.add_format(
        {'bold': True, 'bg_color': '#f1f3f4', 'border': 1, 'align': 'center', 'valign': 'vcenter'})

    # ---------------------------------------------------------
    # SAYFA 1: LİSTE FORMATI (Her Satıra Bir Ders)
    # ---------------------------------------------------------
    list_df = pd.DataFrame(lessons_list)
    list_df.to_excel(writer, index=False, sheet_name='Ders Listesi')
    ws_list = writer.sheets['Ders Listesi']

    # Sütun genişlikleri ve başlık formatı
    col_widths = [12, 15, 15, 35, 15, 15, 25]
    for i, w in enumerate(col_widths):
        ws_list.set_column(i, i, w, list_cell_fmt)
    for col_num, value in enumerate(list_df.columns.values):
        ws_list.write(0, col_num, value, header_fmt)
    ws_list.freeze_panes(1, 0)
    ws_list.autofilter(0, 0, len(list_df), len(list_df.columns) - 1)

    # ---------------------------------------------------------
    # SAYFA 2: AYRIŞIK TAKVİM FORMATI (Görsel Pano)
    # ---------------------------------------------------------
    ws_calendar = workbook.add_worksheet('Görsel Takvim')
    ws_calendar.set_column(0, 0, 18)  # Gün sütunu
    for col_idx in range(1, len(time_slots) + 1):
        ws_calendar.set_column(col_idx, col_idx, 40)  # Saat sütunları

    # Saat Başlıklarını Yaz
    for i, slot in enumerate(time_slots):
        ws_calendar.write(0, i + 1, slot, header_fmt)

    current_row = 1
    for day in days:
        day_data = df_raw[df_raw.iloc[:, 0] == day]
        # Bu günde en çok kaç ders üst üste gelmiş bul (satır yüksekliği için)
        daily_max_lessons = 1
        for slot in time_slots:
            count = len(str(day_data[slot].values[0]).strip().split('\n')) if pd.notna(day_data[slot].values[0]) else 0
            if count > daily_max_lessons: daily_max_lessons = count

        # Dersleri hücre hücre yerleştir
        for slot_idx, slot in enumerate(time_slots):
            content = day_data[slot].values[0]
            if pd.notna(content):
                dersler = str(content).strip().split('\n')
                for i, ders in enumerate(dersler):
                    ws_calendar.write(current_row + i, slot_idx + 1, ders, box_fmt)

            # Boş kalan alt hücreleri boya (çizgiler bozulmasın)
            for i in range(len(dersler) if pd.notna(content) else 0, daily_max_lessons):
                ws_calendar.write(current_row + i, slot_idx + 1, "", list_cell_fmt)

        # Gün ismini birleştir
        if daily_max_lessons > 1:
            ws_calendar.merge_range(current_row, 0, current_row + daily_max_lessons - 1, 0, day, day_fmt)
        else:
            ws_calendar.write(current_row, 0, day, day_fmt)

        current_row += daily_max_lessons

    writer.close()
    print(f"✅ İşlem Tamamlandı! '{output_file}' dosyası oluşturuldu.")
    print("   - Sayfa 1: Ders Listesi (Filtrelenebilir)")
    print("   - Sayfa 2: Görsel Takvim (Pano formatı)")


if __name__ == "__main__":
    tam_program_raporu()