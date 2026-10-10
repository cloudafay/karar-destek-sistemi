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
```

- `hepsini_calistir.py` ham veriden bütün çıktıları yeniden üretir (yarım dakikadan kısa sürer).
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
| `src/veri_hazirlama.py` | 2-4: veri kaydı, sözlük, kalite kontrolü | `data/processed/telco_temiz.csv`, `outputs/veri_kalite_raporu.md`, `docs/veri_sozlugu.md` |
| `src/kesifsel_analiz.py` | 5: keşifsel analiz | `outputs/grafikler/01`–`05`, `outputs/tablolar/kesif_kayip_oranlari.csv` |
| `src/model.py` | 6-7: model ve risk sıralaması | `outputs/model_metrikleri.json`, `outputs/musteri_risk_tablosu.csv`, grafik `06`–`07`, `outputs/tablolar/grup_bazli_kontrol.csv` |
| `src/teklif_senaryosu.py` | 8: teklif senaryosu (sentetik) | `data/processed/teklif_senaryosu_SENTETIK.csv` |
| `src/optimizasyon.py` | 9: bütçe optimizasyonu | `outputs/kampanya_listesi.csv`, `outputs/plan_ozeti.json` |
| `app.py` | 10: yönetici ekranı | Kaydedilen kararlar `outputs/karar_kaydi.csv` dosyasına eklenir |
| `src/senaryo_karsilastirma.py` | 11: senaryo testi ve duyarlılık | `outputs/tablolar/senaryo_karsilastirma.csv`, `duyarlilik_etki_carpani.csv`, `butce_egrisi.csv`, grafik `08`–`10`, `outputs/yonetici_ozeti.md` |
| `src/rapor_olustur.py` | 12: proje raporu | `docs/KDS_Proje_Raporu.docx`, `docs/KDS_Proje_Raporu.pdf` |

## Yönetici ekranı

- **Özet ve karar:** öneri ve beklenen net fayda, bütçenin teklif türlerine dağılımı, bugünkü yöntemle karşılaştırma,
  tekliflerin etkisi beklenenden az ya da çok olursa sonuç, risk dağılımı ve karar formu.
- **Müşteri listesi:** müşteri bazında risk, önerilen teklif ve tek bir müşteri için dört seçeneğin karşılaştırması.
- **Yöntem ve model:** model ölçütleri, kalibrasyon, grup bazlı kontrol ve senaryoların teknik karşılaştırması.

Sol paneldeki bütçe ve teknik destek kapasitesi değiştirildiğinde plan yeniden hesaplanır.

## Yöntem notları

- Veri %80 eğitim, %20 test olarak katmanlı bölünür (rastgelelik tohumu 42). Model ayarları ve olasılık eşiği
  yalnızca eğitim verisinde çapraz doğrulamayla seçilir; test verisine en sonda bir kez bakılır.
- Lojistik regresyon temel modeldir, karar ağacı karşılaştırma içindir. Cinsiyet, yaş grubu, eş ve bakmakla
  yükümlü olunan kişi bilgisi model girdisi değildir; cinsiyet ve yaş grubu yalnızca sonuçların gruplar arasında
  denetlenmesinde kullanılır.
- Kampanya kararı, modelin eğitimde görmediği test müşterileri "güncel portföy" sayılarak verilir.
- Beklenen net fayda = ayrılma olasılığı × teklifin etki oranı × elde tutulan müşterinin 12 aylık katkısı − teklif maliyeti.
  Zaten kalacak müşteriye verilen teklif yalnızca maliyet yaratır.
- Optimizasyon ikili karar değişkenli doğrusal modeldir (PuLP + HiGHS): bütçe sınırı, müşteri başına en fazla bir
  teklif, teknik destek kapasitesi.

## Kaynaklar

- Kaggle, Telco Customer Churn: https://www.kaggle.com/datasets/blastchar/telco-customer-churn
- IBM örnek deposu: https://github.com/IBM/telco-customer-churn-on-icp4d
- scikit-learn model değerlendirme: https://scikit-learn.org/stable/modules/model_evaluation.html
