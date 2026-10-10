# Yönetici özeti

Bu dosya `src/senaryo_karsilastirma.py` tarafından üretilir; elle düzenlemeyin.

İncelenen 1.409 müşteriden 512 tanesi (%36,3) yüksek riskli grupta (ayrılma olasılığı 0,33 ve üzeri).

Sistem, 10.000 $ bütçenin 9.999 $ tutarındaki bölümünü 223 müşteriye ayırmayı öneriyor (İndirim: 57, Uygun paket: 16, Teknik destek: 150). Varsayımlar altında bu planın beklenen ek katkısı 21.584 $, beklenen net faydası 11.585 $.

Aynı bütçeyle mevcut yöntemin (model yok, tek tip indirim) beklenen net faydası 5.907 $; yalnızca risk sırasına göre tek tip indirimin beklenen net faydası 6.991 $.

Bu sayılar kesin sonuç değildir: tekliflerin etkisi gerçek bir kampanyadan değil, varsayımlardan gelir. Gerçek etki varsayılanın yarısı kadar olursa önerilen planın net faydası 793 $ olur. Plan, gerçek etki varsayılan etkinin %46 düzeyine inene kadar maliyetini karşılar.

Önerilen karar: planı önce sınırlı bir pilotla ve kontrol grubuyla uygulamak, etkiyi ölçmek ve doğrulanan faydaya göre genişletmek.

## Senaryo karşılaştırması

| Senaryo | Teklif sunulan müşteri | Maliyet | Beklenen net fayda | Etiket bazlı net fayda | Gerçekte ayrılan hedef |
|---|---|---|---|---|---|
| A - Mevcut yöntem | 130 | 9.964 $ | 5.907 $ | 5.632 $ | 87 (%67) |
| B - Yalnızca risk modeli | 135 | 9.995 $ | 6.991 $ | 7.562 $ | 102 (%76) |
| C - Karar destek sistemi | 223 | 9.999 $ | 11.585 $ | 11.165 $ | 150 (%67) |

Beklenen net fayda modelin olasılıklarıyla, etiket bazlı net fayda ise test müşterilerinin gerçekte ayrılıp ayrılmadığı bilgisiyle hesaplanır. İkisinde de kampanya etkisi varsayımdır.

## Uyarılar

- Tekliflerin maliyeti ve etkisi sentetik varsayımdır; gerçek bir kampanyanın ya da deneyin sonucu değildir. Kanıt düzeyi düşüktür.
- Veri kesitsel bir örnektir. Tahminler belirli bir gelecek dönemi değil, geçmişte ayrılan müşterilere benzerliği gösterir.
- Önerilen 68 teklif, gerçek etki varsayılanın yarısı kadar olursa maliyetini karşılamıyor (listede “Varsayıma duyarlı” olarak işaretli).
- Beklenen net faydası pozitif olan 248 müşteri, bütçe veya teknik destek kapasitesi yetmediği için liste dışında kaldı.
- 3 müşterinin toplam ücret bilgisi eksikti ve 0 ile dolduruldu.
- 1 müşteri, modelin eğitim verisinde görmediği değerler taşıyor; bu tahminler daha az güvenilirdir.
