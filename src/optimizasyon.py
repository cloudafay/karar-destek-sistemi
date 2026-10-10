"""Aşama 9: bütçe kısıtı altında müşteri-teklif seçimi.

İkili karar değişkenli doğrusal model:
    x[m, t] = 1  ->  m müşterisine t teklifi sunulur
    amaç    : toplam beklenen net faydayı en büyük yap
    kısıtlar: toplam maliyet <= bütçe
              her müşteriye en fazla bir teklif
              teknik destek sayısı <= ekip kapasitesi
Beklenen net faydası pozitif olmayan seçenekler modele hiç alınmaz; müşteri riskli diye
zarar eden teklif sunulmaz.

Çalıştırma:  py src/optimizasyon.py
"""
import json

import pandas as pd
import pulp

import ayarlar as A
from ortak import baslik_yaz, para, tr, yuzde
from teklif_senaryosu import teklif_tablosu


def optimize_et(teklifler, butce=A.BUTCE, destek_kapasitesi=A.DESTEK_KAPASITESI):
    """Seçilen müşteri-teklif satırlarını döndürür."""
    aday = teklifler[(teklifler["offer_type"] != A.TEKLIF_YOK) & (teklifler["beklenen_net_fayda"] > 0)]
    if aday.empty or butce <= 0:
        return aday.iloc[0:0]
    fayda, maliyet = aday["beklenen_net_fayda"].to_dict(), aday["offer_cost"].to_dict()

    # PuLP 4 arayüzü: değişkenler problem üzerinden oluşturulur.
    problem = pulp.LpProblem("kampanya_butcesi", pulp.LpMaximize)
    x = problem.add_variable_dicts("sec", list(aday.index), cat="Binary")
    problem.setObjective(pulp.lpSum(fayda[i] * x[i] for i in aday.index))
    problem.addConstraint(pulp.lpSum(maliyet[i] * x[i] for i in aday.index) <= butce, "butce")
    for satirlar in aday.groupby("customerID").groups.values():
        if len(satirlar) > 1:
            problem.addConstraint(pulp.lpSum(x[i] for i in satirlar) <= 1)
    destek = aday.index[aday["offer_type"] == A.DESTEK]
    if len(destek) > destek_kapasitesi:
        problem.addConstraint(pulp.lpSum(x[i] for i in destek) <= destek_kapasitesi, "destek_kapasitesi")

    sonuc = problem.solve(pulp.HiGHS(msg=False))
    if sonuc.status != pulp.LpSolveStatus.Optimal:
        raise RuntimeError(f"Optimizasyon çözülemedi: {sonuc.status_str}")
    return aday.loc[[i for i in aday.index if x[i].value() > 0.5]]


def kampanya_listesi(risk, teklifler, secim, etki_carpani=1.0):
    """Tüm müşteriler için öneri listesi: önerilen teklif ya da neden teklif sunulmadığı."""
    liste = risk[["customerID", "churn_olasiligi", "risk_grubu", "risk_isaretleri"]].merge(
        secim[["customerID", "offer_type", "offer_cost", "beklenen_ek_katki", "beklenen_net_fayda",
               "confidence_or_evidence_level"]], on="customerID", how="left")
    secilmedi = liste["offer_type"].isna()

    gercek = teklifler[teklifler["offer_type"] != A.TEKLIF_YOK]
    en_iyi = gercek.groupby("customerID")["beklenen_net_fayda"].max()
    faydali = liste["customerID"].map(en_iyi) > 0
    liste["aciklama"] = ""
    liste.loc[secilmedi & faydali, "aciklama"] = "Bütçe veya kapasite sınırı nedeniyle listeye girmedi"
    liste.loc[secilmedi & ~faydali, "aciklama"] = "Hiçbir teklifin beklenen net faydası pozitif değil"

    # Etki varsayımı kötümser düzeye inerse teklif hâlâ kendini karşılıyor mu?
    kotumser_oran = min(A.ETKI_CARPANLARI.values()) / etki_carpani
    liste["kotumser_net_fayda"] = liste["beklenen_ek_katki"] * kotumser_oran - liste["offer_cost"]
    liste["dayaniklilik"] = (liste["kotumser_net_fayda"] > 0).map({True: "Dayanıklı", False: "Varsayıma duyarlı"})
    liste.loc[secilmedi, "dayaniklilik"] = ""

    liste["offer_type"] = liste["offer_type"].fillna(A.TEKLIF_YOK)
    liste["confidence_or_evidence_level"] = liste["confidence_or_evidence_level"].fillna("-")
    sayisal = ["offer_cost", "beklenen_ek_katki", "beklenen_net_fayda", "kotumser_net_fayda"]
    liste[sayisal] = liste[sayisal].fillna(0.0)
    return liste.sort_values(["beklenen_net_fayda", "churn_olasiligi"], ascending=False).reset_index(drop=True)


def butce_ozeti(liste, butce):
    onerilen = liste[liste["offer_type"] != A.TEKLIF_YOK]
    harcama = onerilen["offer_cost"].sum()
    return {
        "butce": butce,
        "harcama": harcama,
        "kalan": butce - harcama,
        "kullanim_orani": harcama / butce if butce else 0.0,
        "teklif_sunulan_musteri": len(onerilen),
        "teklif_dagilimi": onerilen["offer_type"].value_counts().reindex(A.TEKLIFLER, fill_value=0).to_dict(),
        "beklenen_ek_katki": onerilen["beklenen_ek_katki"].sum(),
        "beklenen_net_fayda": onerilen["beklenen_net_fayda"].sum(),
        "varsayima_duyarli_teklif": int((onerilen["dayaniklilik"] == "Varsayıma duyarlı").sum()),
        "butce_disinda_kalan": int((liste["aciklama"].str.startswith("Bütçe")).sum()),
    }


def kisitsiz_plan(teklifler):
    """Bütçe ve kapasite sınırı olmasaydı seçilecek plan; bütçe düzeyini yorumlamak için dayanak."""
    aday = teklifler[(teklifler["offer_type"] != A.TEKLIF_YOK) & (teklifler["beklenen_net_fayda"] > 0)]
    secim = aday.loc[aday.groupby("customerID")["beklenen_net_fayda"].idxmax()]  # sınır yoksa herkese en iyi teklifi
    return {
        "teklif_sunulan_musteri": len(secim),
        "maliyet": secim["offer_cost"].sum(),
        "beklenen_net_fayda": secim["beklenen_net_fayda"].sum(),
        "teklif_dagilimi": secim["offer_type"].value_counts().reindex(A.TEKLIFLER, fill_value=0).to_dict(),
    }


def main():
    baslik_yaz("Bütçe optimizasyonu")
    risk = pd.read_csv(A.RISK_TABLOSU)
    teklifler = teklif_tablosu(risk)
    secim = optimize_et(teklifler)
    liste = kampanya_listesi(risk, teklifler, secim)
    liste.to_csv(A.KAMPANYA_LISTESI, index=False, encoding="utf-8-sig")

    ozet = butce_ozeti(liste, A.BUTCE)
    kisitsiz = kisitsiz_plan(teklifler)
    yillik_gelir = risk["MonthlyCharges"].sum() * 12
    A.PLAN_OZETI.write_text(json.dumps(
        {"onerilen_plan": ozet, "kisitsiz_plan": kisitsiz, "portfoy_yillik_gelir": yillik_gelir},
        ensure_ascii=False, indent=2, default=float), encoding="utf-8")
    print(f"Bütçe {para(ozet['butce'])}; önerilen harcama {para(ozet['harcama'])} "
          f"({yuzde(ozet['kullanim_orani'])}), kalan {para(ozet['kalan'])}")
    print(f"Teklif sunulan müşteri: {tr(ozet['teklif_sunulan_musteri'])} / {tr(len(liste))}  "
          + ", ".join(f"{t}: {n}" for t, n in ozet["teklif_dagilimi"].items()))
    print(f"Beklenen ek katkı {para(ozet['beklenen_ek_katki'])}, beklenen net fayda {para(ozet['beklenen_net_fayda'])}")
    print(f"Etki varsayımına duyarlı teklif: {ozet['varsayima_duyarli_teklif']}; "
          f"faydalı olup bütçe/kapasite dışında kalan müşteri: {ozet['butce_disinda_kalan']}")
    print(f"Sınırsız bütçe ve kapasiteyle en iyi plan: {tr(kisitsiz['teklif_sunulan_musteri'])} müşteri, maliyet {para(kisitsiz['maliyet'])}, "
          f"beklenen net fayda {para(kisitsiz['beklenen_net_fayda'])}. Bütçe, portföyün yıllık gelirinin {yuzde(A.BUTCE / yillik_gelir)} kadarı.")
    print(f"Yazıldı: {A.KAMPANYA_LISTESI.relative_to(A.KOK)}, {A.PLAN_OZETI.relative_to(A.KOK)}")


if __name__ == "__main__":
    main()
