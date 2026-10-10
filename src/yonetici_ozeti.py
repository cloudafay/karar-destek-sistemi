"""Yönetici özeti ve uyarılar. Metindeki her sayı hesaplanan çıktılardan gelir."""
import json

import ayarlar as A
from ortak import para, tr, yuzde


def secilen_esik():
    return json.loads(A.MODEL_METRIKLERI.read_text(encoding="utf-8"))["secilen_esik"]


def uyarilar(risk, ozet):
    """Yöneticinin kararı verirken bilmesi gereken sınırlılıklar."""
    liste = [
        "Tekliflerin maliyeti ve etkisi sentetik varsayımdır; gerçek bir kampanyanın ya da deneyin sonucu değildir. "
        "Kanıt düzeyi düşüktür.",
        "Veri kesitsel bir örnektir. Tahminler belirli bir gelecek dönemi değil, geçmişte ayrılan müşterilere benzerliği gösterir.",
    ]
    if ozet["varsayima_duyarli_teklif"]:
        liste.append(f"Önerilen {tr(ozet['varsayima_duyarli_teklif'])} teklif, gerçek etki varsayılanın yarısı kadar olursa "
                     "maliyetini karşılamıyor (listede “Varsayıma duyarlı” olarak işaretli).")
    if ozet["butce_disinda_kalan"]:
        liste.append(f"Beklenen net faydası pozitif olan {tr(ozet['butce_disinda_kalan'])} müşteri, bütçe veya teknik destek "
                     "kapasitesi yetmediği için liste dışında kaldı.")
    eksik, disinda = int(risk["TotalCharges_dolduruldu"].sum()), int(risk["egitim_araligi_disinda"].sum())
    if eksik:
        liste.append(f"{tr(eksik)} müşterinin toplam ücret bilgisi eksikti ve 0 ile dolduruldu.")
    if disinda:
        liste.append(f"{tr(disinda)} müşteri, modelin eğitim verisinde görmediği değerler taşıyor; bu tahminler daha az güvenilirdir.")
    return liste


def ozet_metni(risk, ozet, senaryolar, esik):
    """Karar için gereken bilgileri birkaç paragrafta toplar."""
    yuksek = int((risk["risk_grubu"] == "Yüksek").sum())
    a, b, c = (senaryolar.loc[ad] for ad in (A.SENARYO_A, A.SENARYO_B, A.SENARYO_C))
    dagilim = ", ".join(f"{teklif}: {tr(adet)}" for teklif, adet in ozet["teklif_dagilimi"].items() if adet)
    yari_etki = 0.5 * c["beklenen_ek_katki"] - c["maliyet"]
    return "\n\n".join([
        f"İncelenen {tr(len(risk))} müşteriden {tr(yuksek)} tanesi ({yuzde(yuksek / len(risk))}) yüksek riskli grupta "
        f"(ayrılma olasılığı {tr(esik, 2)} ve üzeri).",
        f"Sistem, {para(ozet['butce'])} bütçenin {para(ozet['harcama'])} tutarındaki bölümünü "
        f"{tr(ozet['teklif_sunulan_musteri'])} müşteriye ayırmayı öneriyor ({dagilim}). Varsayımlar altında bu planın "
        f"beklenen ek katkısı {para(ozet['beklenen_ek_katki'])}, beklenen net faydası {para(ozet['beklenen_net_fayda'])}.",
        f"Aynı bütçeyle mevcut yöntemin (model yok, tek tip indirim) beklenen net faydası {para(a['beklenen_net_fayda'])}; "
        f"yalnızca risk sırasına göre tek tip indirimin beklenen net faydası {para(b['beklenen_net_fayda'])}.",
        "Bu sayılar kesin sonuç değildir: tekliflerin etkisi gerçek bir kampanyadan değil, varsayımlardan gelir. "
        f"Gerçek etki varsayılanın yarısı kadar olursa önerilen planın net faydası {para(yari_etki)} olur. "
        f"Plan, gerçek etki varsayılan etkinin {yuzde(c['basabas_etki_carpani'], 0)} düzeyine inene kadar maliyetini karşılar.",
        "Önerilen karar: planı önce sınırlı bir pilotla ve kontrol grubuyla uygulamak, etkiyi ölçmek ve doğrulanan "
        "faydaya göre genişletmek.",
    ])


def ozet_dosyasi_yaz(risk, ozet, senaryolar):
    satirlar = [
        "# Yönetici özeti",
        "",
        "Bu dosya `src/senaryo_karsilastirma.py` tarafından üretilir; elle düzenlemeyin.",
        "",
        ozet_metni(risk, ozet, senaryolar, secilen_esik()),
        "",
        "## Senaryo karşılaştırması",
        "",
        "| Senaryo | Teklif sunulan müşteri | Maliyet | Beklenen net fayda | Etiket bazlı net fayda | Gerçekte ayrılan hedef |",
        "|---|---|---|---|---|---|",
    ]
    for ad, s in senaryolar.iterrows():
        satirlar.append(f"| {ad} | {tr(s['teklif_sunulan_musteri'])} | {para(s['maliyet'])} | {para(s['beklenen_net_fayda'])} | "
                        f"{para(s['etiket_bazli_net_fayda'])} | {tr(s['gercekte_ayrilan_hedef'])} ({yuzde(s['isabet_orani'], 0)}) |")
    satirlar += [
        "",
        "Beklenen net fayda modelin olasılıklarıyla, etiket bazlı net fayda ise test müşterilerinin gerçekte ayrılıp "
        "ayrılmadığı bilgisiyle hesaplanır. İkisinde de kampanya etkisi varsayımdır.",
        "",
        "## Uyarılar",
        "",
    ]
    satirlar += [f"- {uyari}" for uyari in uyarilar(risk, ozet)]
    (A.CIKTI / "yonetici_ozeti.md").write_text("\n".join(satirlar) + "\n", encoding="utf-8")
