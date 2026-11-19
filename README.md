# Karar Destek Sistemi: müşteri kaybı riski ve kampanya bütçesi

Abonelik modeliyle çalışan bir işletme için öğrenci projesi. Sistem müşterilerin ayrılma riskini tahmin eder,
her müşteri için teklif seçeneklerini (indirim, uygun paket, teknik destek, teklif yok) karşılaştırır ve sınırlı
kampanya bütçesini beklenen net faydayı en yüksek yapacak biçimde dağıtır. Yol haritası `kds.docx` dosyasındadır;
proje raporu `docs/KDS_Proje_Raporu.docx` (ve `.pdf`) dosyasındadır.

**Önemli sınır:** Müşteri verisi gerçektir (IBM Telco Customer Churn, Kaggle). Teklif maliyetleri, tekliflerin etkisi,
bütçe ve kapasite ise **sentetik varsayımdır**; veri setinde kampanya bilgisi yoktur. Net fayda sayıları yöntemi
gösterir, gerçek bir kampanyanın sonucunu kanıtlamaz.

## Çalıştırma

Python 3.14 ile denenmiştir. Bu bilgisayarda Python `py` komutuyla çalışır.

```
py -m pip install -r requirements.txt
py src/hepsini_calistir.py
py -m streamlit run app.py
py src/rapor_olustur.py
py src/sunum_olustur.py
```

- `hepsini_calistir.py` ham veriden bütün çıktıları yeniden üretir (yarım dakikadan kısa sürer).
- `Ekrani_Ac.bat` dosyasına çift tıklamak yönetici ekranını tarayıcıda açar (komut yazmak gerekmez).
- `streamlit run app.py` yönetici ekranını tarayıcıda açar. Streamlit ilk çalıştırmada e-posta sorarsa boş bırakıp
  Enter'a basın.
- `rapor_olustur.py` Word raporunu güncel çıktılardan yeniden yazar; Word kuruluysa içindekiler tablosunu günceller
  ve PDF kopyasını da üretir.

## Varsayımları değiştirmek

Bütçe, teklif maliyetleri, etki oranları ve diğer bütün varsayımlar `src/ayarlar.py` dosyasında tek yerde durur;
her birinin neye dayandığı aynı dosyada ve raporun 7.3. bölümünde yazılıdır. Bir değeri değiştirdikten sonra
yukarıdaki ikinci ve dördüncü komutu yeniden çalıştırın.

Rapor yeniden üretilirken şunlara dikkat edin:

- Rapor dosyasının üzerine yazılır. Word'de elle yaptığınız düzenlemeler silinir. Kapaktaki ad, ders ve öğretim
  elemanı bilgisini `src/ayarlar.py` içindeki `RAPOR_` satırlarına yazarsanız korunur.
- Rapordaki sayılar ve tablolar çıktılardan otomatik gelir. Yorum cümleleri ise şimdiki sonuçlara göre yazılmıştır;
  varsayımları değiştirirseniz yorumların hâlâ doğru olup olmadığını okuyarak denetleyin.
- Rapordaki ekran görüntüsü (`docs/yonetici_ekrani.png`) elle alınmış durağan bir görseldir; ekranı değiştirirseniz
  yenisini aynı adla kaydedin.

## Aşamalar

| Betik | Yol haritasındaki aşama | Ürettiği çıktı |
|---|---|---|
| `src/veri_hazirlama.py` | 2-4: veri kaydı, sözlük, kalite kontrolü, tekrar eden kayıtların çıkarılması | `data/processed/telco_temiz.csv`, `telco_modelleme.csv`, `outputs/veri_kalite_raporu.md`, `outputs/tablolar/tekrar_eden_kayitlar.csv`, `docs/veri_sozlugu.md` |
| `src/kesifsel_analiz.py` | 5: keşifsel analiz | `outputs/grafikler/01`–`05`, `outputs/tablolar/kesif_kayip_oranlari.csv` |
| `src/model.py` | 6-7: model, çapraz doğrulama ve risk sıralaması | `outputs/model_metrikleri.json`, `outputs/musteri_risk_tablosu.csv`, grafik `06`–`07`, `outputs/tablolar/grup_bazli_kontrol.csv`, `capraz_dogrulama_katlar.csv` |
| `src/tekrar_kayit_etkisi.py` | 4: tekrar eden kayıtların doğrulamaya etkisi | `outputs/tablolar/tekrar_kayit_etkisi.csv`, `outputs/tekrar_kayit_ozeti.json` |
| `src/shap_analizi.py` | 6: model açıklaması (SHAP) | `outputs/shap_ozeti.json`, `outputs/tablolar/shap_degisken_onemi.csv`, `shap_grup_onemi.csv`, grafik `11`–`14` |
| `src/teklif_senaryosu.py` | 8: teklif senaryosu (sentetik) | `data/processed/teklif_senaryosu_SENTETIK.csv` |
| `src/optimizasyon.py` | 9: bütçe optimizasyonu | `outputs/kampanya_listesi.csv`, `outputs/plan_ozeti.json` |
| `app.py` | 10: yönetici ekranı | Kaydedilen kararlar `outputs/karar_kaydi.csv` dosyasına eklenir |
| `src/senaryo_karsilastirma.py` | 11: senaryo testi ve duyarlılık | `outputs/tablolar/senaryo_karsilastirma.csv`, `duyarlilik_etki_carpani.csv`, `butce_egrisi.csv`, grafik `08`–`10`, `outputs/yonetici_ozeti.md` |
| `src/rapor_olustur.py` | 12: proje raporu | `docs/KDS_Proje_Raporu.docx`, `docs/KDS_Proje_Raporu.pdf` |
| `src/sunum_olustur.py` | 12: sunum | `docs/KDS_Sunum.pptx` (16 slayt, konuşma notlu), `docs/KDS_Sunum.pdf` |

## Yönetici ekranı

- **Özet ve karar:** öneri ve beklenen net fayda, bütçenin teklif türlerine dağılımı, bugünkü yöntemle karşılaştırma,
  tekliflerin etkisi beklenenden az ya da çok olursa sonuç, risk dağılımı ve karar formu.
- **Müşteri listesi:** müşteri bazında risk, önerilen teklif ve tek bir müşteri için dört seçeneğin karşılaştırması.
- **Yöntem ve model:** model ölçütleri, kalibrasyon, grup bazlı kontrol ve senaryoların teknik karşılaştırması.

Sol paneldeki bütçe ve teknik destek kapasitesi değiştirildiğinde plan yeniden hesaplanır.

## Yöntem notları

- Önce kimlik ve etiket dışındaki bütün özellikleri birebir aynı olan kayıtlar çıkarılır (7.043 → 7.003 kayıt; ayrıntı
  raporun 4.2. bölümünde). Ardından veri %80 eğitim, %20 test olarak katmanlı bölünür (rastgelelik tohumu 42); oranı
  `src/ayarlar.py` içindeki `TEST_ORANI` belirler.
- Eğitim kısmında 5 katlı, iç içe (nested) çapraz doğrulama uygulanır: ayar seçimi her dış katın kendi eğitim kısmında
  yapılır. Olasılık eşiği de yalnızca eğitim kısmında seçilir. Test kısmı bir kez kullanılır ve çapraz doğrulamaya
  sokulmaz; belirsizliği bootstrap ile %95 güven aralığı olarak verilir.
- SHAP, lojistik regresyonun ve karar ağacının tahminlerini açıklar (`src/shap_analizi.py`). Birbirinden türetilebilen
  değişkenler (süre, aylık ücret, toplam ücret) grup olarak yorumlanır; nedenleri raporun 6.8. bölümündedir.
- Lojistik regresyon temel modeldir, karar ağacı karşılaştırma içindir. Cinsiyet ve yaş grubu, ön sunumdaki geri
  bildirim üzerine model girdisine eklenmiştir; modelin bu gruplarda dengeli çalışıp çalışmadığı ayrıca denetlenir.
  Eş ve bakmakla yükümlü olunan kişi bilgisi model girdisi değildir.
- Kampanya kararı, modelin eğitimde görmediği test müşterileri "güncel portföy" sayılarak verilir.
- Beklenen net fayda = ayrılma olasılığı × teklifin etki oranı × elde tutulan müşterinin 12 aylık katkısı − teklif maliyeti.
  Zaten kalacak müşteriye verilen teklif yalnızca maliyet yaratır.
- Optimizasyon ikili karar değişkenli doğrusal modeldir (PuLP + HiGHS): bütçe sınırı, müşteri başına en fazla bir
  teklif, teknik destek kapasitesi.

## Kaynaklar

- Kaggle, Telco Customer Churn: https://www.kaggle.com/datasets/blastchar/telco-customer-churn
- IBM örnek deposu: https://github.com/IBM/telco-customer-churn-on-icp4d
- scikit-learn model değerlendirme: https://scikit-learn.org/stable/modules/model_evaluation.html
