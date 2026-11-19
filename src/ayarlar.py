"""Projenin yolları, sabitleri ve varsayımları.

Bir sayıyı (bütçe, teklif maliyeti, etki oranı...) değiştirmek için tek yer burasıdır.
"""
from pathlib import Path

# --- Klasörler ve dosyalar -------------------------------------------------
KOK = Path(__file__).resolve().parent.parent
HAM_VERI = KOK / "data" / "raw" / "WA_Fn-UseC_-Telco-Customer-Churn.csv"
TEMIZ_VERI = KOK / "data" / "processed" / "telco_temiz.csv"  # temizlenmiş, tüm kayıtlar
MODELLEME_VERISI = KOK / "data" / "processed" / "telco_modelleme.csv"  # tekrar eden kayıtlar çıkarılmış
TEKLIF_TABLOSU = KOK / "data" / "processed" / "teklif_senaryosu_SENTETIK.csv"
CIKTI = KOK / "outputs"
GRAFIK = CIKTI / "grafikler"
TABLO = CIKTI / "tablolar"
DOKUMAN = KOK / "docs"

RISK_TABLOSU = CIKTI / "musteri_risk_tablosu.csv"
MODEL_METRIKLERI = CIKTI / "model_metrikleri.json"
KAMPANYA_LISTESI = CIKTI / "kampanya_listesi.csv"
KARAR_KAYDI = CIKTI / "karar_kaydi.csv"
PLAN_OZETI = CIKTI / "plan_ozeti.json"
TEKRAR_OZETI = CIKTI / "tekrar_kayit_ozeti.json"
TEKRAR_KAYITLAR = TABLO / "tekrar_eden_kayitlar.csv"
CV_TABLOSU = TABLO / "capraz_dogrulama_katlar.csv"
SHAP_OZETI = CIKTI / "shap_ozeti.json"
RAPOR = DOKUMAN / "KDS_Proje_Raporu.docx"
EKRAN_GORUNTUSU = DOKUMAN / "yonetici_ekrani.png"

# --- Veri ve model ----------------------------------------------------------
# CSV'yi hangi siteden indirdiyseniz buraya yazın; kalite raporuna aynen geçer.
INDIRILEN_KAYNAK = "Kaggle (https://www.kaggle.com/datasets/blastchar/telco-customer-churn)"
BEKLENEN_SATIR, BEKLENEN_SUTUN = 7043, 21  # kaynak açıklamasındaki boyut
RASTGELE_TOHUM = 42
TEST_ORANI = 0.20  # %80 eğitim, %20 test
CV_KAT = 5         # eğitim kısmında (%80) uygulanan çapraz doğrulamanın kat sayısı
BOOTSTRAP_TEKRAR = 1000  # test kısmının (%20) güven aralıkları için yeniden örnekleme sayısı

HEDEF = "Churn"
SAYISAL = ["tenure", "MonthlyCharges", "TotalCharges"]
# Cinsiyet ve yaş grubu (65 yaş ve üzeri mi), ön sunumdaki öğretim elemanı geri bildirimi üzerine model girdisine
# eklendi. Ayrımcılık riskine karşı modelin bu gruplarda dengeli çalışıp çalışmadığı ayrıca denetlenir (GRUP_KONTROL).
DEMOGRAFIK = ["gender", "SeniorCitizen"]
KATEGORIK = DEMOGRAFIK + [
    "PhoneService", "MultipleLines", "InternetService", "OnlineSecurity",
    "OnlineBackup", "DeviceProtection", "TechSupport", "StreamingTV",
    "StreamingMovies", "Contract", "PaperlessBilling", "PaymentMethod",
]
# Model girdisi olmayan sütunlar ve gerekçeleri
MODEL_DISI = {
    "customerID": "Yalnızca kayıt eşleştirme ve tekrar kontrolü içindir.",
    "Partner": "Hane bilgisi; iş gerekçesi zayıf olduğu için kullanılmaz.",
    "Dependents": "Hane bilgisi; iş gerekçesi zayıf olduğu için kullanılmaz.",
}
GRUP_KONTROL = DEMOGRAFIK

# Tekrar eden kayıt: bu sütunlar dışındaki bütün özellikleri birebir aynı olan kayıtlar.
# (Kimlik her kayıtta farklıdır; etiket ise çelişip çelişmediğini görmek için ayrıca karşılaştırılır.)
TEKRAR_ANAHTARI_HARIC = ["customerID", HEDEF, "TotalCharges_dolduruldu"]

# Keşifsel analizde ve risk işaretlerinde kullanılan aralıklar
SURE_SINIR, SURE_ETIKET = [-1, 12, 24, 48, 72], ["0–12 ay", "13–24 ay", "25–48 ay", "49–72 ay"]
UCRET_SINIR = [0, 35, 55, 75, 95, float("inf")]
UCRET_ETIKET = ["35'ten az", "35–55", "55–75", "75–95", "95 ve üzeri"]

# --- SENTETİK VARSAYIMLAR ---------------------------------------------------
# Aşağıdaki her değer yöntemi göstermek için seçilmiş bir senaryo parametresidir.
# Gerçek bir işletmeden, sektör kaynağından veya kampanya deneyinden gelmez; yuvarlak ve
# makul görünen değerlerdir. Kime hangi teklifin uygun olduğu ve kimde daha etkili sayıldığı
# keşifsel analizdeki örüntülere dayanır; sayıların kendisi dayanmaz.
PARA = "$"            # veri setindeki ücretlerin birimi
UFUK_AY = 12          # elde tutulan müşterinin katkısının hesaplandığı dönem
KATKI_MARJI = 0.50    # aylık ücretin katkı kârına dönen payı
# Bütçe: sınırsız bütçeyle en iyi planın maliyetinin kabaca yarısı olacak biçimde seçilmiş yuvarlak bir değerdir.
# Portföy büyüklüğü değişirse (ör. TEST_ORANI) bu oran da değişir; güncel oran raporun 8.2. bölümünde
# hesaplanarak yazılır.
BUTCE = 10_000
DESTEK_KAPASITESI = 150  # teknik destek ekibinin dönemde bakabileceği müşteri sayısı (aynı not geçerli)

TEKLIF_YOK = "Teklif yok"
INDIRIM = "İndirim"
PAKET = "Uygun paket"
DESTEK = "Teknik destek"
TEKLIFLER = [INDIRIM, PAKET, DESTEK]

# İndirim: INDIRIM_AY boyunca aylık ücretten INDIRIM_ORANI kadar indirim.
# Taahhütsüz (aylık sözleşmeli) müşteride daha etkili sayılır; bu grupta kayıp oranı en yüksektir.
INDIRIM_ORANI = 0.15
INDIRIM_AY = 6
INDIRIM_ETKI = {"Month-to-month": 0.35, "One year": 0.15, "Two year": 0.15}

# Uygun paket: yüksek ücretli müşteriyi UFUK_AY boyunca biraz daha ucuz pakete geçirme.
# İndirim ve paket değerleri, her teklifin en az bir müşteri profilinde diğerlerinden iyi
# çıkacağı biçimde seçilmiştir; aksi hâlde tek bir teklif her zaman kazanır.
PAKET_ALT_SINIR = 70          # bu aylık ücretin altındakilere sunulmaz
PAKET_UCRET_DUSUSU = 0.05
PAKET_GECIS_MALIYETI = 20
PAKET_YUKSEK_UCRET = 90
PAKET_ETKI_YUKSEK, PAKET_ETKI_NORMAL = 0.35, 0.20

# Teknik destek: interneti olup teknik destek hizmeti almayan müşteriye ücretsiz destek.
# Fiber müşteride daha etkili sayılır; fiber kullananlarda kayıp oranı daha yüksektir.
DESTEK_MALIYETI = 35
DESTEK_ETKI = {"Fiber optic": 0.25, "DSL": 0.18}

KANIT_DUZEYI = "Düşük (sentetik varsayım)"

# Rapor kapağı: kendi bilgilerinizi buraya yazın; rapor yeniden üretildiğinde korunur.
RAPOR_HAZIRLAYAN = "[Ad Soyad] – [Öğrenci numarası]"
RAPOR_DERS = "[Dersin adı]"
RAPOR_OGRETIM_ELEMANI = "[Unvan Ad Soyad]"

SENARYO_A, SENARYO_B, SENARYO_C = "A - Mevcut yöntem", "B - Yalnızca risk modeli", "C - Karar destek sistemi"
ETKI_CARPANLARI = {"Kötümser": 0.5, "Temel": 1.0, "İyimser": 1.5}
