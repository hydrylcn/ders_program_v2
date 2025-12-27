import pandas as pd

# 1. Ders ve tercih bilgilerini tanımla
program = {
    "Pazartesi": {
        "09:00-12:00": [
            "Introduction to Microeconomics (Derslik 110)",
            "Sürdürülebilir Pazarlama (Derslik 204)"
        ],
        "13:00-16:00": [
            "İstatistik I (Derslik 209)",
            "Yönetim Geliştirme (Derslik 103)"
        ],
        "16:00-19:00": [
            "Muhasebe I (Derslik 302)",
            "Bilgi Yönetimi (Derslik 211)"
        ],
        "19:00-21:00": [
            "Matematik-I (Tek) (Derslik 105)",
            "Introduction to Business (Örgün + İ.Ö) (Derslik 201)"
        ]
    },
    "Salı": {
        "09:00-12:00": [
            "Mali Tablolar Analizi (Derslik 102)",
            "Consumer Behavior (Derslik 205)",
            "Operations Management I (Derslik 104)"
        ],
        "13:00-16:00": [
            "Sosyoloji (Derslik 105)",
            "Organizational Behavior (Derslik 210)"
        ],
        "16:00-19:00": [
            "Muhasebe I (Örgün + İ.Ö.) (Derslik 209)",
            "İşletmeye Giriş (Derslik 202)",
            "Reklamcılık Yönetimi (Derslik 302)"
        ],
        "19:00-21:00": [
            "Araştırma Yöntemleri (Derslik 108)",
            "E-İş ve Kurumsal Kaynak Planlama (Derslik 301)"
        ]
    }
}

# 2. Pandas DataFrame oluştur
gunler = list(program.keys())
saatler = ["09:00-12:00", "13:00-16:00", "16:00-19:00", "19:00-21:00"]

# 3. Verileri DataFrame formatına dönüştür
data = []

for gun in gunler:
    satir = [gun]  # Gün adı
    for saat in saatler:
        dersler = ", ".join([ders.split(' (')[0] + (' (' + ders.split('(')[1] if '(' in ders else '') for ders in program[gun].get(saat, [])])  # Ders ismini ve varsa derslik bilgisini ekle, virgülle ayır
        satir.append(dersler)
    data.append(satir)

# DataFrame oluştur
df = pd.DataFrame(data, columns=["Gün"] + saatler)

# 4. Excel dosyasına yaz
df.to_excel("tercih.xlsx", index=False)

print("✅ Tercih formu oluşturuldu: tercih.xlsx")
