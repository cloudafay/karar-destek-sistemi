"""Aşama 5: keşifsel analiz. Kayıp oranlarını gruplara göre hesaplar ve grafiklerini çizer.

Grafikler ilişki gösterir; neden-sonuç ilişkisini kanıtlamaz.

Çalıştırma:  py src/kesifsel_analiz.py
"""
import matplotlib.pyplot as plt
import pandas as pd

import ayarlar as A
import grafik_stili as G
from ortak import baslik_yaz, tr, yuzde

SOZLESME = {"Month-to-month": "Aylık (taahhütsüz)", "One year": "1 yıllık", "Two year": "2 yıllık"}
DESTEK = {"No": "Teknik destek yok", "Yes": "Teknik destek var", "No internet service": "İnternet hizmeti yok"}


def kayip_orani(df, grup, sira=None):
    """Bir gruplama (sütun adı ya da hazır gruplama serisi) için kayıp oranı tablosu."""
    ayrildi = df[A.HEDEF].eq("Yes")
    grup = df[grup] if isinstance(grup, str) else grup
    tablo = ayrildi.groupby(grup, observed=True).agg(musteri_sayisi="size", ayrilan="sum")
    tablo["kayip_orani"] = tablo["ayrilan"] / tablo["musteri_sayisi"]
    return tablo.loc[sira] if sira is not None else tablo


def genel_oran_grafigi(df, oran):
    fig = plt.figure(figsize=(7.2, 2.7))
    G.basliklar(fig, "Genel müşteri kayıp oranı", "Veri setindeki tüm müşteriler")
    fig.text(0.04, 0.50, yuzde(oran), fontsize=38, fontweight="bold", va="center")
    fig.text(0.31, 0.50, f"{tr(len(df))} müşteriden {tr(df[A.HEDEF].eq('Yes').sum())} kişi ayrıldı.",
             fontsize=11, color=G.IKINCIL, va="center")
    ax = fig.add_axes((0.05, 0.12, 0.90, 0.12))
    ax.plot([0, 1], [0, 0], color=G.ACIK_MAVI, linewidth=9, solid_capstyle="butt")
    ax.plot([0, oran], [0, 0], color=G.SERI[0], linewidth=9, solid_capstyle="butt")
    ax.set_xlim(0, 1)
    ax.axis("off")
    G.kaydet(fig, "01_genel_kayip_orani.png")


def etiketle(tablo, adlar=None):
    return [f"{adlar[g] if adlar else g}\n{tr(n)} müşteri" for g, n in zip(tablo.index, tablo["musteri_sayisi"])]


def main():
    baslik_yaz("Keşifsel analiz")
    df = pd.read_csv(A.TEMIZ_VERI)
    genel = df[A.HEDEF].eq("Yes").mean()
    referans = dict(referans=genel, referans_etiketi=f"Genel oran {yuzde(genel)}")

    tablolar = {
        "Sözleşme türü": kayip_orani(df, "Contract", list(SOZLESME)),
        "Abonelik süresi": kayip_orani(df, pd.cut(df["tenure"], A.SURE_SINIR, labels=A.SURE_ETIKET)),
        "Aylık ücret": kayip_orani(df, pd.cut(df["MonthlyCharges"], A.UCRET_SINIR, labels=A.UCRET_ETIKET, right=False)),
        "Teknik destek": kayip_orani(df, "TechSupport", list(DESTEK)),
    }

    genel_oran_grafigi(df, genel)
    G.cubuk_grafigi(etiketle(tablolar["Sözleşme türü"], SOZLESME), tablolar["Sözleşme türü"]["kayip_orani"],
                    "Sözleşme türüne göre kayıp oranı", "Ayrılan müşterilerin gruptaki payı",
                    "02_sozlesme_turune_gore.png", **referans)
    G.cubuk_grafigi(etiketle(tablolar["Abonelik süresi"]), tablolar["Abonelik süresi"]["kayip_orani"],
                    "Abonelik süresine göre kayıp oranı", "Ayrılan müşterilerin gruptaki payı",
                    "03_abonelik_suresine_gore.png", yatay=False, **referans)
    G.cubuk_grafigi(etiketle(tablolar["Aylık ücret"]), tablolar["Aylık ücret"]["kayip_orani"],
                    "Aylık ücrete göre kayıp oranı", f"Ayrılan müşterilerin gruptaki payı; ücret aralıkları {A.PARA}",
                    "04_aylik_ucrete_gore.png", yatay=False, **referans)
    G.cubuk_grafigi(etiketle(tablolar["Teknik destek"], DESTEK), tablolar["Teknik destek"]["kayip_orani"],
                    "Teknik destek durumuna göre kayıp oranı", "Ayrılan müşterilerin gruptaki payı",
                    "05_teknik_destege_gore.png", **referans)

    # Grafiklerin tablo karşılığı
    birlesik = pd.concat(
        [t.rename_axis("grup").reset_index().assign(degisken=ad) for ad, t in tablolar.items()], ignore_index=True
    )[["degisken", "grup", "musteri_sayisi", "ayrilan", "kayip_orani"]]
    A.TABLO.mkdir(parents=True, exist_ok=True)
    birlesik.to_csv(A.TABLO / "kesif_kayip_oranlari.csv", index=False, encoding="utf-8-sig")

    print(f"Genel kayıp oranı: {yuzde(genel)}")
    for ad, tablo in tablolar.items():
        print(f"{ad}: " + "; ".join(f"{g} {yuzde(o)}" for g, o in tablo["kayip_orani"].items()))
    print("Tablo: outputs/tablolar/kesif_kayip_oranlari.csv")


if __name__ == "__main__":
    main()
