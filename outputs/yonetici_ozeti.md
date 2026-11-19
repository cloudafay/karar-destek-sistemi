# Yönetici özeti

Bu dosya `src/senaryo_karsilastirma.py` tarafından üretilir; elle düzenlemeyin.

İncelenen 1.401 müşteriden 512 tanesi (%36,5) yüksek riskli grupta (ayrılma olasılığı 0,31 ve üzeri).

Sistem, 10.000 $ bütçenin 9.999 $ tutarındaki bölümünü 224 müşteriye ayırmayı öneriyor (İndirim: 52, Uygun paket: 22, Teknik destek: 150). Varsayımlar altında bu planın beklenen ek katkısı 21.515 $, beklenen net faydası 11.515 $.

Aynı bütçeyle mevcut yöntemin (model yok, tek tip indirim) beklenen net faydası 5.679 $; yalnızca risk sırasına göre tek tip indirimin beklenen net faydası 6.978 $.

Bu sayılar kesin sonuç değildir: tekliflerin etkisi gerçek bir kampanyadan değil, varsayımlardan gelir. Gerçek etki varsayılanın yarısı kadar olursa önerilen planın net faydası 758 $ olur. Plan, gerçek etki varsayılan etkinin %46 düzeyine inene kadar maliyetini karşılar.

Önerilen karar: planı önce sınırlı bir pilotla ve kontrol grubuyla uygulamak, etkiyi ölçmek ve doğrulanan faydaya göre genişletmek.

## Senaryo karşılaştırması

| Senaryo | Teklif sunulan müşteri | Maliyet | Beklenen net fayda | Etiket bazlı net fayda | Gerçekte ayrılan hedef |
|---|---|---|---|---|---|
| A - Mevcut yöntem | 129 | 9.990 $ | 5.679 $ | 7.107 $ | 95 (%74) |
| B - Yalnızca risk modeli | 132 | 9.959 $ | 6.978 $ | 6.945 $ | 97 (%73) |
| C - Karar destek sistemi | 224 | 9.999 $ | 11.515 $ | 11.687 $ | 156 (%70) |

Beklenen net fayda modelin olasılıklarıyla, etiket bazlı net fayda ise test müşterilerinin gerçekte ayrılıp ayrılmadığı bilgisiyle hesaplanır. İkisinde de kampanya etkisi varsayımdır.

## Uyarılar

- Tekliflerin maliyeti ve etkisi sentetik varsayımdır; gerçek bir kampanyanın ya da deneyin sonucu değildir. Kanıt düzeyi düşüktür.
- Veri kesitsel bir örnektir. Tahminler belirli bir gelecek dönemi değil, geçmişte ayrılan müşterilere benzerliği gösterir.
- Önerilen 66 teklif, gerçek etki varsayılanın yarısı kadar olursa maliyetini karşılamıyor (listede “Varsayıma duyarlı” olarak işaretli).
- Beklenen net faydası pozitif olan 211 müşteri, bütçe veya teknik destek kapasitesi yetmediği için liste dışında kaldı.
- 5 müşterinin toplam ücret bilgisi eksikti ve 0 ile dolduruldu.
- 2 müşteri, modelin eğitim verisinde görmediği değerler taşıyor; bu tahminler daha az güvenilirdir.
