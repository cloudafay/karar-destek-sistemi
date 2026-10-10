# Veri sözlüğü

Bu dosya `src/veri_hazirlama.py` tarafından üretilir; elle düzenlemeyin.

| Sütun | Açıklama | Tip | Değerler | Modeldeki rolü |
|---|---|---|---|---|
| customerID | Müşteri kimliği | metin | 7.043 benzersiz kod | Girdi değil. Yalnızca kayıt eşleştirme ve tekrar kontrolü içindir. |
| gender | Cinsiyet | kategori | Female, Male | Girdi değil. Demografik alan; ayrımcılık riski nedeniyle girdi değildir, yalnızca grup bazlı kontrolde kullanılır. |
| SeniorCitizen | 65 yaş ve üzeri mi (1 = evet) | kategori | 0, 1 | Girdi değil. Demografik alan; ayrımcılık riski nedeniyle girdi değildir, yalnızca grup bazlı kontrolde kullanılır. |
| Partner | Eşi/partneri var mı | kategori | No, Yes | Girdi değil. Demografik alan; iş gerekçesi zayıf olduğu için kullanılmaz. |
| Dependents | Bakmakla yükümlü olduğu kişi var mı | kategori | No, Yes | Girdi değil. Demografik alan; iş gerekçesi zayıf olduğu için kullanılmaz. |
| tenure | Müşterinin şirkette bulunduğu ay sayısı | sayı | 0,00 – 72,00 | Girdi. |
| PhoneService | Telefon hizmeti var mı | kategori | No, Yes | Girdi. |
| MultipleLines | Birden fazla hattı var mı | kategori | No, No phone service, Yes | Girdi. |
| InternetService | İnternet hizmet türü | kategori | DSL, Fiber optic, No | Girdi. |
| OnlineSecurity | Çevrimiçi güvenlik ek hizmeti | kategori | No, No internet service, Yes | Girdi. |
| OnlineBackup | Çevrimiçi yedekleme ek hizmeti | kategori | No, No internet service, Yes | Girdi. |
| DeviceProtection | Cihaz koruma ek hizmeti | kategori | No, No internet service, Yes | Girdi. |
| TechSupport | Teknik destek ek hizmeti | kategori | No, No internet service, Yes | Girdi. |
| StreamingTV | TV yayın hizmeti | kategori | No, No internet service, Yes | Girdi. |
| StreamingMovies | Film yayın hizmeti | kategori | No, No internet service, Yes | Girdi. |
| Contract | Sözleşme türü | kategori | Month-to-month, One year, Two year | Girdi. |
| PaperlessBilling | Kâğıtsız fatura kullanıyor mu | kategori | No, Yes | Girdi. |
| PaymentMethod | Ödeme yöntemi | kategori | Bank transfer (automatic), Credit card (automatic), Electronic check, Mailed check | Girdi. |
| MonthlyCharges | Aylık ücret ($) | sayı | 18,25 – 118,75 | Girdi. |
| TotalCharges | Bugüne kadarki toplam ücret ($) | sayı | 0,00 – 8.684,80 | Girdi. |
| Churn | Müşteri ayrıldı mı (hedef değişken) | kategori | No, Yes | Hedef değişken; girdi olarak kullanılmaz. |

Temizleme sırasında eklenen sütun: `TotalCharges_dolduruldu` (ham dosyada TotalCharges boşsa True). Model girdisi değildir.
