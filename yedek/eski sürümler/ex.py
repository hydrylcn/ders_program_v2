import pandas as pd

# -----------------------------
# 1. VERİLER
# -----------------------------
ogretim_uyeleri = [
    "Öğretmen 1",
    "Öğretmen 2",
    "Öğretmen 3",
    "Öğretmen 4",
    "Öğretmen 5",
    "Dr. Öğr. Üyesi Miray BAYBARS",
    "Doç. Dr. Burcu ŞENTÜRK YILDIZ",
    "Araş. Gör. Dr. Özgür BABACAN",
    "Prof. Dr. G. Nazan GÜNAY",
    "Prof. Dr. Burcu ARACIOĞLU",
    "Doç. Dr. İnanç KABASAKAL",
    "Prof. Dr. A. Nazlı AYYILDIZ ÜNNÜ",
    "Dr. Öğr. Üyesi Hakan ERKAL",
    "Dr. Öğr. Üyesi A. Erhan ZALLUHOĞLU",
    "Prof. Dr. Dilek DEMİRHAN",
    "Dr. Öğr. Üyesi Esin GÜRBÜZ",
    "Prof. Dr. Ayla Özhan DEDEOĞLU",
    "Doç. Dr. Elif ÜSTÜNDAĞLI ERTEN",
    "Prof. Dr. Murat KOCAMAZ",
    "Doç. Dr. Haydar YALÇIN",
    "Prof. Dr. İpek KAZANÇOĞLU",
    "Dr. Öğr. Üyesi Ş. Sertaç ÇAKI",
    "Doç. Dr. Aydın KOÇAK",
    "Prof. Dr. Haluk SOYUER",
    "Prof. Dr. Türker SUSMUŞ",
    "Prof. Dr. Derya İLİC",
    "Doç. Dr. U. Gökay ÇİÇEKLİ",
    "Prof. Dr. Burak ÇAPRAZ",
    "Prof. Dr. Aykan CANDEMİR",
    "Prof. Dr. Keti VENTURA",
    "Araş. Gör. Dr. Begüm KANAT TİRYAKİ",
    "Dr. Öğr. Üyesi Sema AYDIN"
]


dersler = [
    "Matematik-I (Tek)",
    "Matematik-I (Çift)",
    "Business",
    "Sosyoloji",
    "Introduction to Microeconomics",
    "Kariyer Planlama",
    "Türk Dili I",
    "Atatürk İlkeleri ve İnkılap Tarihi I",
    "Muhasebe I",
    "Marketing Management I",
    "Operations Management I",
    "İstatistik I",
    "Organizational Behavior (Tek)",
    "Organizational Behavior (Çift)",
    "Araştırma Yöntemleri (Çift)",
    "Araştırma Yöntemleri (Tek)",
    "Financial Management I",
    "İşletme Hukuku",
    "Consumer Behavior (Tek)",
    "Consumer Behavior (Çift)",
    "Operations Research I",
    "Girişimcilik ve KOBİ Yönetimi",
    "Enformetri",
    "Sürdürülebilir Pazarlama",
    "Yönetim Muhasebesi",
    "Reklamcılık Yönetimi",
    "E-İş ve Kurumsal Kaynak Planlama",
    "Teknoloji ve Sanayi Dinamikleri",
    "Uygulamalı Finansal Piyasa İşlemleri",
    "Küresel Tedarik Zinciri ve Lojistik",
    "Bilgi Yönetimi",
    "Yatırım Yönetimi",
    "Management Consultancy",
    "Yönetim Geliştirme",
    "Borçlar Hukuku",
    "System Analysis and Design",
    "Human Resources Management",
    "Global Marketing (Tek)",
    "Global Marketing (Çift)",
    "Mali Tablolar Analizi",
    "Vestel İşletmecilik Seminerleri I",
    "Retailing I",
    "Muhasebe I (Örgün + İ.Ö.)",
    "Muhasebe II (Tasfiye)",
    "Uygulamalı Finansal Piyasa İşlemleri (Örgün)",
    "Business I (Tasfiye)",
    "Introduction to Business (Örgün + İ.Ö)",
    "Hukukun Temel Kavramları (Örgün + İ.Ö.)",
    "Ticaret Hukuku (Tasfiye)",
    "İşletmeye Giriş",
"Pazarlama Yönetimi",
    "Örgütler ve Yönetim",
    "Stratejik İşletme Finansı",
    "İşlemler Yönetimi",
    "Ticaret Hukuku",
    "Endüstriyel Pazarlama",
    "Tüketici Davranışları",
    "Stratejik Pazarlama Yönetimi",
    "Perakendecilik",
    "Hizmet Pazarlaması",
    "İnsan Kaynakları Yönetimi",
    "Bilimsel Araştırma Yöntemleri ve Etik",
    "Örgütsel Davranış",
    "Örgütler ve Yönetim",
    "Kurumsal Koçluk ve Mentorluk",
    "Yönetim Atölyesi",
    "Ağ Modelleri",
    "Malzeme ve Stok Yönetimi",
    "Lojistik ve Tedarik Zincirinde Güncel Konular",
    "Tedarik Zinciri ve Lojistik İçin Bilgi Sistemleri",
    "Dönem Projesi",
    "Stratejik Finans Yönetimi",
    "Finansal Muhasebe",
    "Sermaye Piyasaları ve Menkul Kıymetler Analizi",
    "Vadeli İşlem Piyasaları",
    "Uluslararası Finansman",
    "Veri Odaklı Üretim Planlama Stratejileri",
    "Proje Yönetimi",
    "Akıllı Karar Modelleri",
    "Stratejik Üretim Yönetimi ve Üretimde Dijitalleşme",
    "İş Analitiğinde Sayısal Yöntemler",
    "Stratejik Yatırım Kararları ve Planlama",
    "Hizmet Pazarlaması"
]


derslikler = [
    "Derslik 101",
    "Derslik 102",
    "Derslik 103",
    "Derslik 104",
    "Derslik 105",
    "Derslik 108",
    "Derslik 109",
    "Derslik 110",
    "Derslik 111",
    "Derslik 201",
    "Derslik 202",
    "Derslik 203",
    "Derslik 204",
    "Derslik 205",
    "Derslik 208",
    "Derslik 209",
    "Derslik 210",
    "Derslik 211",
    "Derslik 301",
    "Derslik 302",
    "Derslik 303"
]


kontenjan_bilgisi = {
    "Derslik 101": 32,
    "Derslik 102": 49,
    "Derslik 103": 20,
    "Derslik 104": 48,
    "Derslik 105": 49,
    "Derslik 108": 49,
    "Derslik 109": 20,
    "Derslik 110": 20,
    "Derslik 111": 20,
    "Derslik 201": 30,
    "Derslik 202": 49,
    "Derslik 203": 20,
    "Derslik 204": 48,
    "Derslik 205": 49,
    "Derslik 208": 49,
    "Derslik 209": 20,
    "Derslik 210": 20,
    "Derslik 211": 20,
    "Derslik 301": 41,
    "Derslik 302": 36,
    "Derslik 303": 64
}


# Derslerin sınıf bilgisi
sinif_bilgisi = {
    "1. Sınıf": [
        "Matematik-I (Tek)",
        "Matematik-I (Çift)",
        "Business",
        "Sosyoloji",
        "Introduction to Microeconomics",
        "Kariyer Planlama",
        "Türk Dili I",
        "Atatürk İlkeleri ve İnkılap Tarihi I"
    ],
    "2. Sınıf": [
        "Muhasebe I",
        "Marketing Management I",
        "Operations Management I",
        "İstatistik I",
        "Organizational Behavior (Tek)",
        "Organizational Behavior (Çift)"
    ],
    "3. Sınıf": [
        "Araştırma Yöntemleri (Çift)",
        "Araştırma Yöntemleri (Tek)",
        "Financial Management I",
        "İşletme Hukuku",
        "Consumer Behavior (Tek)",
        "Consumer Behavior (Çift)",
        "Operations Research I",
        "Girişimcilik ve KOBİ Yönetimi",
        "Enformetri",
        "Sürdürülebilir Pazarlama",
        "Yönetim Muhasebesi",
        "Reklamcılık Yönetimi",
        "E-İş ve Kurumsal Kaynak Planlama",
        "Teknoloji ve Sanayi Dinamikleri",
        "Uygulamalı Finansal Piyasa İşlemleri",
        "Küresel Tedarik Zinciri ve Lojistik",
        "Bilgi Yönetimi",
        "Yatırım Yönetimi",
        "Management Consultancy",
        "Yönetim Geliştirme",
        "Borçlar Hukuku"
    ],
    "4. Sınıf": [
        "System Analysis and Design",
        "Human Resources Management",
        "Global Marketing (Tek)",
        "Global Marketing (Çift)",
        "Mali Tablolar Analizi",
        "Vestel İşletmecilik Seminerleri I",
        "Retailing I"
    ],
    "0. Sınıf": [
        "Muhasebe I (Örgün + İ.Ö.)",
        "Muhasebe II (Tasfiye)",
        "Uygulamalı Finansal Piyasa İşlemleri (Örgün)",
        "Business I (Tasfiye)",
        "Introduction to Business (Örgün + İ.Ö)",
        "Hukukun Temel Kavramları (Örgün + İ.Ö.)",
        "Ticaret Hukuku (Tasfiye)",
        "İşletmeye Giriş"
    ],
    "İşletme İ.Ö. Tezsiz Yüksek Lisans Programı": [
        "Pazarlama Yönetimi",
        "Örgütler ve Yönetim",
        "Stratejik İşletme Finansı",
        "İşlemler Yönetimi",
        "Ticaret Hukuku"
    ],
    "Pazarlama ve Marka Yönetimi İ.Ö. Tezsiz Yüksek Lisans Programı": [
        "Endüstriyel Pazarlama",
        "Tüketici Davranışları",
        "Stratejik Pazarlama Yönetimi",
        "Perakendecilik",
        "Hizmet Pazarlaması"
    ],
    "İnsan Kaynakları Yönetimi ve Yönetim Geliştirme İ.Ö. Tezsiz Yüksek Lisans Programı": [
        "İnsan Kaynakları Yönetimi",
        "Bilimsel Araştırma Yöntemleri ve Etik",
        "Örgütsel Davranış",
        "Örgütler ve Yönetim"
    ],
    "Uzaktan Öğretim Lojistik Yönetimi İ.Ö. Tezsiz Yüksek Lisans Programı": [
        "Bilimsel Araştırma Yöntemleri ve Etik",
        "Ağ Modelleri",
        "Malzeme ve Stok Yönetimi",
        "Lojistik ve Tedarik Zincirinde Güncel Konular"
    ],
    "Muhasebe ve Finansman İ.Ö. Tezsiz Yüksek Lisans Programı": [
        "Stratejik Finans Yönetimi",
        "Finansal Muhasebe",
        "Sermaye Piyasaları ve Menkul Kıymetler Analizi",
        "Bilimsel Araştırma Yöntemleri ve Etik"
    ],
    "Dijital Teknolojilerle Üretim Yönetimi İ.Ö. Tezsiz Yüksek Lisans Programı": [
        "Veri Odaklı Üretim Planlama Stratejileri",
        "Proje Yönetimi",
        "Akıllı Karar Modelleri",
        "Stratejik Üretim Yönetimi ve Üretimde Dijitalleşme"
    ],
    "Uzaktan Öğretim İşletme Tezsiz Yüksek Lisans Programı": [
        "Örgütler ve Yönetim",
        "Tedarik Zinciri Yönetimi ve Lojistik",
        "Pazarlama Yönetimi",
        "İşlemler Yönetimi"
    ]
}


# Sabit ders atamaları
sabit_ders_atamasi = {
    "ÖĞRETMEN 1": [
        "Matematik-I (Tek)"
    ],
    "ÖĞRETMEN 2": [
        "Matematik-I (Çift)"
    ],
    "Dr. Öğr. Üyesi Miray BAYBARS": [
        "Business",
        "Kariyer Planlama",
        "Retailing I"
    ],
    "Doç. Dr. Burcu ŞENTÜRK YILDIZ": [
        "Sosyoloji"
    ],
    "ÖĞRETMEN 3": [
        "Introduction to Microeconomics"
    ],
    "ÖĞRETMEN 4": [
        "Türk Dili I"
    ],
    "ÖĞRETMEN 5": [
        "Atatürk İlkeleri ve İnkılap Tarihi I"
    ],
    "Araş. Gör. Dr. Özgür BABACAN": [
        "Muhasebe I",
        "Yatırım Yönetimi",
        "İşletmeye Giriş"
    ],
    "Prof. Dr. G. Nazan GÜNAY": [
        "Marketing Management I",
        "Vestel İşletmecilik Seminerleri I"
    ],
    "Prof. Dr. Burcu ARACIOĞLU": [
        "Operations Management I",
        "Küresel Tedarik Zinciri ve Lojistik"
    ],
    "Doç. Dr. İnanç KABASAKAL": [
        "İstatistik I",
        "Bilgi Yönetimi"
    ],
    "Prof. Dr. A. Nazlı AYYILDIZ ÜNNÜ": [
        "Organizational Behavior (Tek)"
    ],
    "Dr. Öğr. Üyesi Hakan ERKAL": [
        "Organizational Behavior (Çift)"
    ],
    "Dr. Öğr. Üyesi A. Erhan ZALLUHOĞLU": [
        "Araştırma Yöntemleri (Çift)",
        "Araştırma Yöntemleri (Tek)",
        "Girişimcilik ve KOBİ Yönetimi"
    ],
    "Prof. Dr. Dilek DEMİRHAN": [
        "Financial Management I"
    ],
    "Dr. Öğr. Üyesi Esin GÜRBÜZ": [
        "İşletme Hukuku",
        "Borçlar Hukuku"
    ],
    "Prof. Dr. Ayla Özhan DEDEOĞLU": [
        "Consumer Behavior (Tek)"
    ],
    "Doç. Dr. Elif ÜSTÜNDAĞLI ERTEN": [
        "Consumer Behavior (Çift)",
        "Reklamcılık Yönetimi"
    ],
    "Prof. Dr. Murat KOCAMAZ": [
        "Operations Research I"
    ],
    "Doç. Dr. Haydar YALÇIN": [
        "Enformetri"
    ],
    "Prof. Dr. İpek KAZANÇOĞLU": [
        "Sürdürülebilir Pazarlama"
    ],
    "Dr. Öğr. Üyesi Ş. Sertaç ÇAKI": [
        "Yönetim Muhasebesi",
        "Mali Tablolar Analizi"
    ],
    "Doç. Dr. Aydın KOÇAK": [
        "E-İş ve Kurumsal Kaynak Planlama"
    ],
    "Prof. Dr. Haluk SOYUER": [
        "Teknoloji ve Sanayi Dinamikleri"
    ],
    "Prof. Dr. Türker SUSMUŞ": [
        "Uygulamalı Finansal Piyasa İşlemleri",
        "Muhasebe I (Örgün + İ.Ö.)",
        "Muhasebe II (Tasfiye)",
        "Uygulamalı Finansal Piyasa İşlemleri (Örgün)"
    ],
    "Prof. Dr. Derya İLİC": [
        "Management Consultancy",
        "Yönetim Geliştirme"
    ],
    "Doç. Dr. U. Gökay ÇİÇEKLİ": [
        "System Analysis and Design"
    ],
    "Prof. Dr. Burak ÇAPRAZ": [
        "Human Resources Management"
    ],
    "Prof. Dr. Aykan CANDEMİR": [
        "Global Marketing (Tek)"
    ],
    "Prof. Dr. Keti VENTURA": [
        "Global Marketing (Çift)"
    ],
    "Araş. Gör. Dr. Begüm KANAT TİRYAKİ": [
        "Business I (Tasfiye)",
        "Introduction to Business (Örgün + İ.Ö)"
    ],
    "Dr. Öğr. Üyesi Sema AYDIN": [
        "Hukukun Temel Kavramları (Örgün + İ.Ö.)",
        "Ticaret Hukuku (Tasfiye)"
    ]
}


# -----------------------------
# 2. DERSLERİN SINIF BİLGİSİ
# -----------------------------
ders_sinif_dict = {}
for sinif, ders_listesi in sinif_bilgisi.items():
    for ders in ders_listesi:
        ders_sinif_dict[ders] = sinif

# -----------------------------
# 3. EXCEL SAYFALARI
# -----------------------------
df_uyeler = pd.DataFrame({"OgretimUyesi": ogretim_uyeleri})
df_dersler = pd.DataFrame({
    "Dersler": dersler,
    "Sinif": [ders_sinif_dict.get(ders, "Bilinmiyor") for ders in dersler]
})
df_derslikler = pd.DataFrame({
    "Derslikler": derslikler,
    "Kontenjan": [kontenjan_bilgisi.get(derslik, "Bilinmiyor") for derslik in derslikler]
})

data_uyeler_dersler = []
for uye, ders_listesi in sabit_ders_atamasi.items():
    for ders in ders_listesi:
        data_uyeler_dersler.append({
            "OgretimUyesi": uye,
            "Ders": ders,
            "Sinif": ders_sinif_dict.get(ders, "Bilinmiyor")
        })

df_uyeler_dersler = pd.DataFrame(data_uyeler_dersler)

# -----------------------------
# 4. EXCEL DOSYASINA YAZ
# -----------------------------
with pd.ExcelWriter("dersler.xlsx") as writer:
    df_uyeler.to_excel(writer, sheet_name="OgretimUyeleri", index=False)
    df_dersler.to_excel(writer, sheet_name="Dersler", index=False)
    df_derslikler.to_excel(writer, sheet_name="Derslikler", index=False)
    df_uyeler_dersler.to_excel(writer, sheet_name="OgretimUyeleriDersler", index=False)

print("✅ dersler.xlsx dosyası oluşturuldu!")
