"""Aşama 8: teklif senaryosu tablosu (SENTETİK).

Veri setinde kampanya, teklif maliyeti veya kampanya etkisi yoktur. Buradaki maliyet ve etki
değerleri ayarlar.py içindeki varsayımlardan üretilir; gerçek bir kampanyanın sonucu değildir.

Hesap:
    ek elde tutma olasılığı = ayrılma olasılığı × teklifin etki oranı
    beklenen ek katkı       = ek elde tutma olasılığı × elde tutulan müşterinin dönemlik katkısı
    beklenen net fayda      = beklenen ek katkı − teklif maliyeti

Etki oranı, "teklif sunulmasaydı ayrılacak olan müşterilerin ne kadarı teklifle kalır" sorusunun
varsayılan yanıtıdır. Zaten kalacak müşteriye verilen teklif yalnızca maliyet yaratır.

Çalıştırma:  py src/teklif_senaryosu.py
"""
import numpy as np
import pandas as pd

import ayarlar as A
from ortak import baslik_yaz, para, tr


def teklif_tablosu(risk, etki_carpani=1.0, katki_marji=A.KATKI_MARJI):
    """Her müşteri için uygun olduğu tekliflerin maliyetini ve beklenen net faydasını hesaplar.

    Müşteriye uygun olmayan teklif tabloda yer almaz (ör. interneti olmayana teknik destek).
    """
    musteri = risk[["customerID", "churn_olasiligi", "Contract", "MonthlyCharges", "InternetService", "TechSupport"]]
    ucret = musteri["MonthlyCharges"]

    indirim = musteri.assign(
        offer_type=A.INDIRIM,
        offer_cost=A.INDIRIM_ORANI * A.INDIRIM_AY * ucret,
        etki_orani=musteri["Contract"].map(A.INDIRIM_ETKI),
    )
    paket = musteri[ucret >= A.PAKET_ALT_SINIR].assign(
        offer_type=A.PAKET,
        offer_cost=lambda d: A.PAKET_GECIS_MALIYETI + A.PAKET_UCRET_DUSUSU * A.UFUK_AY * d["MonthlyCharges"],
        etki_orani=lambda d: np.where(d["MonthlyCharges"] >= A.PAKET_YUKSEK_UCRET, A.PAKET_ETKI_YUKSEK, A.PAKET_ETKI_NORMAL),
    )
    destek = musteri[musteri["InternetService"].isin(A.DESTEK_ETKI) & musteri["TechSupport"].eq("No")].assign(
        offer_type=A.DESTEK,
        offer_cost=float(A.DESTEK_MALIYETI),
        etki_orani=lambda d: d["InternetService"].map(A.DESTEK_ETKI),
    )
    teklif_yok = musteri.assign(offer_type=A.TEKLIF_YOK, offer_cost=0.0, etki_orani=0.0)

    tablo = pd.concat([indirim, paket, destek, teklif_yok], ignore_index=True)
    tablo["etki_orani"] = (tablo["etki_orani"] * etki_carpani).clip(upper=1.0)
    tablo["estimated_incremental_retention"] = tablo["churn_olasiligi"] * tablo["etki_orani"]
    tablo["retained_customer_contribution"] = tablo["MonthlyCharges"] * katki_marji * A.UFUK_AY
    tablo["beklenen_ek_katki"] = tablo["estimated_incremental_retention"] * tablo["retained_customer_contribution"]
    tablo["beklenen_net_fayda"] = tablo["beklenen_ek_katki"] - tablo["offer_cost"]
    tablo["confidence_or_evidence_level"] = np.where(tablo["offer_type"] == A.TEKLIF_YOK, "-", A.KANIT_DUZEYI)
    return tablo[["customerID", "offer_type", "offer_cost", "etki_orani", "estimated_incremental_retention",
                  "retained_customer_contribution", "beklenen_ek_katki", "beklenen_net_fayda",
                  "confidence_or_evidence_level"]]


def main():
    baslik_yaz("Teklif senaryosu (SENTETİK)")
    risk = pd.read_csv(A.RISK_TABLOSU)
    tablo = teklif_tablosu(risk)
    tablo.to_csv(A.TEKLIF_TABLOSU, index=False, encoding="utf-8-sig")

    gercek = tablo[tablo["offer_type"] != A.TEKLIF_YOK]
    ozet = gercek.groupby("offer_type").agg(
        uygun_musteri=("customerID", "size"),
        ortalama_maliyet=("offer_cost", "mean"),
        net_faydasi_pozitif=("beklenen_net_fayda", lambda s: int((s > 0).sum())),
    ).reindex(A.TEKLIFLER)
    for teklif, satir in ozet.iterrows():
        print(f"{teklif}: {tr(satir['uygun_musteri'])} müşteriye uygun, ortalama maliyet {para(satir['ortalama_maliyet'], 2)}, "
              f"{tr(satir['net_faydasi_pozitif'])} müşteride beklenen net fayda pozitif")
    print(f"Yazıldı: {A.TEKLIF_TABLOSU.relative_to(A.KOK)}")


if __name__ == "__main__":
    main()
