"""Tekrar eden kayıtların doğrulamaya etkisini ölçer.

Üç veri düzeni aynı ayırma oranı, aynı tohum ve aynı yöntemle karşılaştırılır:
  1. Tüm kayıtlar (7.043)
  2. Tekrar eden kayıtlar çıkarılmış: kimlik ve etiket dışındaki bütün özellikleri aynı olanlar (asıl kural)
  3. Daha sıkı kural: modelin kullandığı 15 sütunu aynı olan kayıtlar da çıkarılmış (duyarlılık denemesi)

Her düzen için şunlar sayılır:
  - test kısmındaki kaç kaydın eğitim kısmında "ikizi" (aynı özellikli kaydı) var,
  - çapraz doğrulamada kaç doğrulama kaydının ikizi o katın eğitim kısmında kalıyor (sızıntı),
  - lojistik regresyon ve karar ağacının test başarısı.

Çalıştırma:  py src/tekrar_kayit_etkisi.py   (önce veri_hazirlama.py çalışmış olmalı)
"""
import json

import numpy as np
import pandas as pd
from sklearn.metrics import average_precision_score, roc_auc_score
from sklearn.model_selection import GridSearchCV, StratifiedKFold

import ayarlar as A
import model as M
from ortak import baslik_yaz, tr
from veri_hazirlama import tekrar_anahtari

DUZENLER = ["Tüm kayıtlar", "Tekrar eden kayıtlar çıkarıldı (asıl kural)", "Model girdisi aynı olanlar da çıkarıldı (sıkı kural)"]


def grup_numarasi(df, sutunlar):
    """Aynı özelliklere sahip kayıtlara aynı numarayı verir."""
    return df.groupby(sutunlar, sort=False, dropna=False).ngroup()


def ikiz_sayilari(df, grup, egitim_idx, test_idx):
    """Test kısmındaki kaç kaydın eğitim kısmında ikizi var; kaçında etiket çelişiyor."""
    egitim_etiketleri = df.loc[egitim_idx].groupby(grup.loc[egitim_idx])[A.HEDEF].agg(set)
    ikiz = grup.loc[test_idx].isin(egitim_etiketleri.index)
    celisen = [df.at[i, A.HEDEF] not in egitim_etiketleri[grup.at[i]] for i in test_idx[ikiz.to_numpy()]]
    return int(ikiz.sum()), int(sum(celisen))


def cv_sizintisi(grup, egitim_idx, y_egitim, katlar):
    """Çapraz doğrulamada ikizi aynı katın eğitim kısmında bulunan doğrulama kaydı sayısı (kat toplamı)."""
    g = grup.loc[egitim_idx].to_numpy()
    sayi = 0
    for egitim_konum, dogrulama_konum in katlar.split(np.zeros(len(g)), y_egitim):
        egitimdekiler = set(g[egitim_konum])
        sayi += sum(grup_no in egitimdekiler for grup_no in g[dogrulama_konum])
    return sayi


def duzeni_degerlendir(ad, veri):
    df, X_egitim, X_test, y_egitim, y_test = M.veri_bol(veri)
    katlar = StratifiedKFold(n_splits=A.CV_KAT, shuffle=True, random_state=A.RASTGELE_TOHUM)
    ozellik19, girdi15 = tekrar_anahtari(df), A.SAYISAL + A.KATEGORIK
    sonuc = {"duzen": ad, "kayit": len(df), "egitim": len(X_egitim), "test": len(X_test)}
    for etiket, sutunlar in (("19", ozellik19), ("15", girdi15)):
        grup = grup_numarasi(df, sutunlar)
        sonuc[f"test_ikizi_{etiket}"], sonuc[f"test_ikizi_celisen_{etiket}"] = ikiz_sayilari(df, grup, X_egitim.index, X_test.index)
        sonuc[f"cv_sizintisi_{etiket}"] = cv_sizintisi(grup, X_egitim.index, y_egitim, katlar)
    for model_adi, (model, izgara) in M.ADAYLAR.items():
        boru = GridSearchCV(M.boru_hatti(model), izgara, scoring="roc_auc", cv=katlar).fit(X_egitim, y_egitim).best_estimator_
        p = boru.predict_proba(X_test)[:, 1]
        anahtar = "lojistik" if model_adi == M.LOJISTIK else "agac"
        sonuc[f"{anahtar}_roc_auc"], sonuc[f"{anahtar}_pr_auc"] = roc_auc_score(y_test, p), average_precision_score(y_test, p)
    return sonuc


def main():
    baslik_yaz("Tekrar eden kayıtların etkisi")
    tum = pd.read_csv(A.TEMIZ_VERI)
    asil = pd.read_csv(A.MODELLEME_VERISI)
    siki = tum.drop_duplicates(subset=A.SAYISAL + A.KATEGORIK, keep="first")

    satirlar = [duzeni_degerlendir(ad, veri) for ad, veri in zip(DUZENLER, (tum, asil, siki))]
    tablo = pd.DataFrame(satirlar)
    A.TABLO.mkdir(parents=True, exist_ok=True)
    tablo.to_csv(A.TABLO / "tekrar_kayit_etkisi.csv", index=False, encoding="utf-8-sig")

    # Doğrulama: asıl kuralla çıkarılan veride 19 özellik düzeyinde ikiz ve sızıntı kalmamalı.
    assert satirlar[1]["test_ikizi_19"] == 0 and satirlar[1]["cv_sizintisi_19"] == 0
    assert satirlar[2]["test_ikizi_15"] == 0 and satirlar[2]["cv_sizintisi_15"] == 0

    ozet = json.loads(A.TEKRAR_OZETI.read_text(encoding="utf-8"))
    ozet["girdi_ayni_fazla_kayit"] = int(len(tum) - len(siki))
    ozet["etki"] = satirlar
    A.TEKRAR_OZETI.write_text(json.dumps(ozet, ensure_ascii=False, indent=2), encoding="utf-8")

    for s in satirlar:
        print(f"{s['duzen']}: {tr(s['kayit'])} kayıt (eğitim {tr(s['egitim'])}, test {tr(s['test'])})")
        print(f"  19 özellik düzeyinde: test kaydı ikizi {s['test_ikizi_19']} (etiketi çelişen {s['test_ikizi_celisen_19']}), "
              f"çapraz doğrulama sızıntısı {s['cv_sizintisi_19']}")
        print(f"  15 girdi düzeyinde  : test kaydı ikizi {s['test_ikizi_15']} (etiketi çelişen {s['test_ikizi_celisen_15']}), "
              f"çapraz doğrulama sızıntısı {s['cv_sizintisi_15']}")
        print(f"  test ROC-AUC lojistik {tr(s['lojistik_roc_auc'], 4)}, karar ağacı {tr(s['agac_roc_auc'], 4)}; "
              f"PR-AUC lojistik {tr(s['lojistik_pr_auc'], 4)}, karar ağacı {tr(s['agac_pr_auc'], 4)}")
    print("Yazıldı: outputs/tablolar/tekrar_kayit_etkisi.csv, outputs/tekrar_kayit_ozeti.json")


if __name__ == "__main__":
    main()
