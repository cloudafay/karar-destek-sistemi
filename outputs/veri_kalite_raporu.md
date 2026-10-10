# Veri kaynağı kaydı ve kalite raporu

Bu dosya `src/veri_hazirlama.py` tarafından üretilir; elle düzenlemeyin.

## Kaynak kaydı

- Dosya: `data/raw/WA_Fn-UseC_-Telco-Customer-Churn.csv`
- Boyut: 977.501 bayt
- Dosya tarihi (indirme): 10.10.2026 17:40
- SHA-256: `88be4b93fbe0cc83421af1c503794c97c342eca914c1576db7c276e61d61358a`
- Veri seti: IBM Telco Customer Churn (Kaggle: blastchar/telco-customer-churn)
- İndirilen kaynak: Kaggle (https://www.kaggle.com/datasets/blastchar/telco-customer-churn)

## Kontroller

- Boyut: 7.043 satır × 21 sütun; kaynak açıklamasıyla (7.043 × 21) uyumlu.
- Tekrar eden customerID: 0; tümüyle aynı satır: 0.
- Boş hücre içeren sütunlar: TotalCharges (11).
- TotalCharges ham dosyada metin olarak gelir; 11 kayıtta boştur. Bu kayıtların tenure değeri: 0.
- Churn dağılımı: Yes = 1.869, No = 5.174; ayrılma oranı %26,5. Sınıflar dengesiz olduğu için yalnızca doğruluk oranına bakılmaz.

## Uygulanan dönüşümler

- Tüm metin hücrelerinin baş ve sonundaki boşluklar silindi.
- tenure ve SeniorCitizen tam sayıya, MonthlyCharges ve TotalCharges ondalık sayıya çevrildi.
- Boş 11 TotalCharges değeri 0 ile dolduruldu (hepsi tenure = 0 olan, henüz faturası oluşmamış müşteriler). Bu kayıtlar `TotalCharges_dolduruldu` sütunuyla işaretlendi.
- Satır silinmedi; ham dosya değiştirilmedi.

## Sayısal alanların özeti (temiz veri)

| Alan | En küçük | Ortanca | Ortalama | En büyük |
|---|---|---|---|---|
| tenure | 0,00 | 29,00 | 32,37 | 72,00 |
| MonthlyCharges | 18,25 | 70,35 | 64,76 | 118,75 |
| TotalCharges | 0,00 | 1.394,55 | 2.279,73 | 8.684,80 |
