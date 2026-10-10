"""Aşama 11: üç karar yaklaşımının aynı bütçe ve aynı varsayımlarla karşılaştırılması.

    A - Mevcut yöntem        : model yok; aylık sözleşmeli ve ilk 12 ayındaki müşterilere,
                               en yüksek ücretliden başlayarak bütçe bitene kadar aynı indirim
    B - Yalnızca risk modeli : yüksek riskli müşterilere, en riskliden başlayarak aynı indirim
    C - Karar destek sistemi : müşteri ve teklif bazında net fayda optimizasyonu

Sonuçlar sentetik kampanya varsayımlarına dayanır. C'nin kendi varsayımları altında daha iyi
çıkması, gerçek dünyada aynı iyileşmenin elde edileceğini kanıtlamaz.

Çalıştırma:  py src/senaryo_karsilastirma.py
"""
import pandas as pd
from matplotlib.ticker import FuncFormatter

import ayarlar as A
import grafik_stili as G
from optimizasyon import butce_ozeti, kampanya_listesi, optimize_et
from ortak import baslik_yaz, para, tr, yuzde
from teklif_senaryosu import teklif_tablosu
from yonetici_ozeti import ozet_dosyasi_yaz

BUTCE_DUZEYLERI = [2_500, 5_000, 7_500, 10_000, 15_000, 20_000, 30_000]
ETKI_DUZEYLERI = [0.5, 0.75, 1.0, 1.25, 1.5]


def tek_tip_indirim(teklifler, sirali_kimlikler, butce):
    """Verilen sıradaki müşterilere bütçe bitene kadar aynı indirimi verir."""
    indirim = teklifler[teklifler["offer_type"] == A.INDIRIM].set_index("customerID").loc[sirali_kimlikler].reset_index()
    return indirim[indirim["offer_cost"].cumsum() <= butce]


def senaryo_secimleri(risk, teklifler, butce, destek_kapasitesi=A.DESTEK_KAPASITESI):
    kural = risk[(risk["Contract"] == "Month-to-month") & (risk["tenure"] <= 12)]
    yuksek_risk = risk[risk["risk_grubu"] == "Yüksek"]
    return {
        A.SENARYO_A: tek_tip_indirim(teklifler, kural.sort_values("MonthlyCharges", ascending=False)["customerID"], butce),
        A.SENARYO_B: tek_tip_indirim(teklifler, yuksek_risk.sort_values("churn_olasiligi", ascending=False)["customerID"], butce),
        A.SENARYO_C: optimize_et(teklifler, butce, destek_kapasitesi),
    }


def degerlendir(secim, risk, butce):
    """Bir seçimin maliyetini ve faydasını iki bakışla ölçer.

    Beklenen net fayda modelin olasılıklarıyla hesaplanır. Etiket bazlı net fayda ise aynı etki
    varsayımını, test müşterilerinin gerçekte ayrılıp ayrılmadığı bilgisiyle birleştirir; modelin
    olasılıkları kararı yanlış yönlendiriyorsa burada görünür. Kampanya etkisi yine varsayımdır.
    """
    ayrildi = secim["customerID"].map(risk.set_index("customerID")["gercek_churn"])
    maliyet, ek_katki = secim["offer_cost"].sum(), secim["beklenen_ek_katki"].sum()
    etiket_katkisi = (ayrildi * secim["etki_orani"] * secim["retained_customer_contribution"]).sum()
    return {
        "teklif_sunulan_musteri": len(secim),
        "maliyet": maliyet,
        "butce_kullanimi": maliyet / butce if butce else 0.0,
        "beklenen_ek_katki": ek_katki,
        "beklenen_net_fayda": ek_katki - maliyet,
        "gercekte_ayrilan_hedef": int(ayrildi.sum()),
        "isabet_orani": ayrildi.mean() if len(secim) else 0.0,
        "etiket_bazli_net_fayda": etiket_katkisi - maliyet,
        # gerçek etki, varsayılanın bu katına inene kadar plan zarar etmez
        "basabas_etki_carpani": maliyet / ek_katki if ek_katki else float("nan"),
    }


def senaryo_tablosu(risk, teklifler, butce, destek_kapasitesi=A.DESTEK_KAPASITESI):
    secimler = senaryo_secimleri(risk, teklifler, butce, destek_kapasitesi)
    return pd.DataFrame({ad: degerlendir(secim, risk, butce) for ad, secim in secimler.items()}).T.rename_axis("senaryo")


def duyarlilik_tablosu(senaryolar):
    """Planlar sabitken gerçek etki varsayılandan farklı çıkarsa net fayda ne olur?"""
    return pd.DataFrame(
        {ad: [c * s["beklenen_ek_katki"] - s["maliyet"] for c in ETKI_DUZEYLERI] for ad, s in senaryolar.iterrows()},
        index=pd.Index(ETKI_DUZEYLERI, name="etki_carpani"))


def butce_egrisi(risk, teklifler):
    return pd.DataFrame(
        {b: senaryo_tablosu(risk, teklifler, b)["beklenen_net_fayda"] for b in BUTCE_DUZEYLERI}
    ).T.rename_axis("butce")


def main():
    baslik_yaz("Senaryo karşılaştırması")
    risk = pd.read_csv(A.RISK_TABLOSU)
    teklifler = teklif_tablosu(risk)

    senaryolar = senaryo_tablosu(risk, teklifler, A.BUTCE)
    duyarlilik = duyarlilik_tablosu(senaryolar)
    egri = butce_egrisi(risk, teklifler)
    A.TABLO.mkdir(parents=True, exist_ok=True)
    senaryolar.to_csv(A.TABLO / "senaryo_karsilastirma.csv", encoding="utf-8-sig")
    duyarlilik.to_csv(A.TABLO / "duyarlilik_etki_carpani.csv", encoding="utf-8-sig")
    egri.to_csv(A.TABLO / "butce_egrisi.csv", encoding="utf-8-sig")

    para_ekseni = FuncFormatter(lambda deger, _: tr(deger))
    varsayim = f"{para(A.BUTCE)} bütçe, sentetik kampanya varsayımları, {tr(len(risk))} müşteri"
    G.cubuk_grafigi(
        [f"{ad}\n{tr(s['teklif_sunulan_musteri'])} müşteri" for ad, s in senaryolar.iterrows()],
        senaryolar["beklenen_net_fayda"], "Senaryolara göre beklenen net fayda", f"{A.PARA}; {varsayim}",
        "08_senaryo_karsilastirma.png", bicim=para, eksen_bicimi=para_ekseni)
    G.cizgi_grafigi(
        egri.index.tolist(), {ad: egri[ad].tolist() for ad in egri.columns},
        "Bütçe arttıkça beklenen net fayda", f"{A.PARA}; sentetik kampanya varsayımları, {tr(len(risk))} müşteri",
        "09_butce_egrisi.png", f"Kampanya bütçesi ({A.PARA})", para_ekseni, para_ekseni, vurgu=A.SENARYO_C)
    G.cizgi_grafigi(
        ETKI_DUZEYLERI, {ad: duyarlilik[ad].tolist() for ad in duyarlilik.columns},
        "Kampanya etkisi varsayımdan saparsa net fayda", f"{A.PARA}; planlar sabit, {varsayim}",
        "10_duyarlilik.png", "Gerçekleşen etki / varsayılan etki",
        FuncFormatter(lambda deger, _: f"×{tr(deger, 2)}"), para_ekseni, vurgu=A.SENARYO_C)

    # Önerilen plan (C) ve yönetici özeti
    liste = kampanya_listesi(risk, teklifler, optimize_et(teklifler))
    ozet_dosyasi_yaz(risk, butce_ozeti(liste, A.BUTCE), senaryolar)

    for ad, s in senaryolar.iterrows():
        print(f"{ad}: {tr(s['teklif_sunulan_musteri'])} müşteri, maliyet {para(s['maliyet'])}, "
              f"beklenen net fayda {para(s['beklenen_net_fayda'])}, etiket bazlı net fayda {para(s['etiket_bazli_net_fayda'])}, "
              f"isabet {yuzde(s['isabet_orani'])}, başabaş etki çarpanı ×{tr(s['basabas_etki_carpani'], 2)}")
    print("Yazıldı: outputs/tablolar/senaryo_karsilastirma.csv, duyarlilik_etki_carpani.csv, butce_egrisi.csv, outputs/yonetici_ozeti.md")


if __name__ == "__main__":
    main()
