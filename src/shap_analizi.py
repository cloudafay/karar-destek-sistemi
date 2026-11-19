"""SHAP analizi: tahmin modelinin kararlarını değişken bazında açıklar.

Ne yapar:
  1. Lojistik regresyon (temel model) için SHAP değerlerini hesaplar. Doğrusal modelde SHAP kesindir:
     katkı = katsayı × (müşterinin değeri − eğitim ortalaması), log-oran biriminde.
  2. Bir-sıcak kodlanmış sütunların katkılarını özgün değişkene toplar (17 değişken).
  3. Sonuçları iki bağımsız yolla denetler: karar ağacının TreeSHAP değerleri ve permütasyon önemi.
  4. Bağımlı değişken sorununu ölçer: aylık ücretin hizmet sütunlarıyla doğrusal ilişkisi vb.

SHAP modelin tahminini açıklar; ayrılmanın nedenini ya da bir müdahalenin etkisini kanıtlamaz.

Çalıştırma:  py src/shap_analizi.py   (önce model.py çalışmış olmalı)
"""
import json
import warnings

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import shap
from scipy.stats import spearmanr
from sklearn.inspection import permutation_importance
from sklearn.linear_model import LinearRegression, LogisticRegression
from sklearn.metrics import roc_auc_score
from sklearn.tree import DecisionTreeClassifier

import ayarlar as A
import grafik_stili as G
import model as M
from ortak import baslik_yaz, tr, yuzde

DEGISKEN_ADI = {"tenure": "Abonelik süresi", "MonthlyCharges": "Aylık ücret", "TotalCharges": "Toplam ücret", **M.OZELLIK_ADI}
HIZMET_SUTUNLARI = ["PhoneService", "MultipleLines", "InternetService", "OnlineSecurity", "OnlineBackup",
                    "DeviceProtection", "TechSupport", "StreamingTV", "StreamingMovies"]
EN_ONEMLI_SAYISI = 5     # değer yönü grafiğinde gösterilen değişken sayısı
# Birbirinden türetilebilen (bağımlı) değişkenler tek bir grupta birlikte değerlendirilir; Aas ve ark. (2021)
# bağımlı değişkenlerin Shapley değerlerinin gruplar hâlinde toplanmasını önerir.
GRUPLAR = {
    "Süre ve ücret": ["tenure", "MonthlyCharges", "TotalCharges"],
    "Hizmet paketi": HIZMET_SUTUNLARI,
    "Sözleşme": ["Contract"],
    "Faturalandırma ve ödeme": ["PaperlessBilling", "PaymentMethod"],
    "Demografik": A.DEMOGRAFIK,
}
# Başka değişkenlerden neredeyse tam hesaplanabildiği için tek başına yorumlanmayacak değişkenler.
TURETILMIS = ["MonthlyCharges", "TotalCharges"]


def grup_izin_onemi(boru, X, y, tekrar=30):
    """Bir grubun bütün sütunlarını aynı satır sırasıyla birlikte karıştırıp ROC-AUC'deki düşüşü ölçer."""
    rastgele = np.random.default_rng(A.RASTGELE_TOHUM)
    temel = roc_auc_score(y, boru.predict_proba(X)[:, 1])
    sonuc = {}
    for ad, sutunlar in GRUPLAR.items():
        dususler = []
        for _ in range(tekrar):
            karisik = X.copy()
            karisik[sutunlar] = X[sutunlar].to_numpy()[rastgele.permutation(len(X))]
            dususler.append(temel - roc_auc_score(y, boru.predict_proba(karisik)[:, 1]))
        sonuc[ad] = float(np.mean(dususler))
    return pd.Series(sonuc)


def yogun(matris):
    return matris.toarray() if hasattr(matris, "toarray") else np.asarray(matris)


def kaynak_degiskenler(on_isleme):
    """Ön işlemeden çıkan her sütunun hangi özgün değişkenden geldiği (sütun sırasıyla)."""
    kategoriler = on_isleme.named_transformers_["kategorik"].categories_
    return np.array(A.SAYISAL + [ozellik for ozellik, degerler in zip(A.KATEGORIK, kategoriler) for _ in degerler])


def degiskene_topla(degerler, kaynak, indeks):
    """Bir-sıcak sütunların SHAP değerlerini özgün değişkene toplar (SHAP toplamsal olduğu için geçerlidir)."""
    return pd.DataFrame(degerler, index=indeks).T.groupby(kaynak).sum().T


def modelleri_kur(X_egitim, y_egitim, parametreler):
    lojistik = M.boru_hatti(LogisticRegression(max_iter=2000, C=parametreler[M.LOJISTIK]["model__C"])).fit(X_egitim, y_egitim)
    p = parametreler[M.AGAC]
    agac = M.boru_hatti(DecisionTreeClassifier(
        random_state=A.RASTGELE_TOHUM, max_depth=p["model__max_depth"], min_samples_leaf=p["model__min_samples_leaf"],
    )).fit(X_egitim, y_egitim)
    return lojistik, agac


def deger_metni(degisken, deger):
    if degisken == "tenure":
        return f"{int(deger)} ay"
    if degisken == "MonthlyCharges":
        return f"{tr(deger, 2)} {A.PARA}"
    if degisken == "TotalCharges":
        return f"{tr(deger)} {A.PARA}"
    return M.DEGER_ADI.get(deger, deger)


def isaretli(deger):
    return ("+" if deger > 0 else "") + tr(deger, 2)


def deger_gruplari(X, degisken):
    """Bir değişkenin değerlerini grafikte gösterilecek gruplara ayırır (kategoriler aynen, sayılar aralıklara)."""
    if degisken == "tenure":
        return pd.cut(X[degisken], A.SURE_SINIR, labels=A.SURE_ETIKET).astype(str)
    if degisken == "MonthlyCharges":
        return pd.cut(X[degisken], A.UCRET_SINIR, labels=A.UCRET_ETIKET, right=False).astype(str)
    if degisken == "TotalCharges":
        return pd.qcut(X[degisken], 4, labels=["en düşük çeyrek", "2. çeyrek", "3. çeyrek", "en yüksek çeyrek"]).astype(str)
    return X[degisken].astype(str)


def grup_etiketi(degisken, grup):
    if degisken == "MonthlyCharges":
        return f"{DEGISKEN_ADI[degisken]}: {grup} {A.PARA}"
    if degisken in ("tenure", "TotalCharges"):
        return f"{DEGISKEN_ADI[degisken]}: {grup}"
    return f"{DEGISKEN_ADI[degisken]}: {M.DEGER_ADI.get(grup, grup)}"


def grup_karsilastirma_grafigi(grup_onemi):
    """Dört değişken grubunun payı, üç farklı ölçüye göre yan yana (SHAP lojistik, SHAP ağaç, permütasyon önemi)."""
    olcutler = [("lojistik_shap_pay", "SHAP (lojistik)"), ("agac_shap_pay", "SHAP (ağaç)"),
                ("izin_onemi_pay", "Permütasyon (lojistik)")]
    n = len(grup_onemi)
    fig, ax = plt.subplots(figsize=(7.4, 1.9 + 0.95 * n))
    fig.subplots_adjust(left=0.27, right=0.95, top=1 - 1.35 / fig.get_figheight(), bottom=0.5 / fig.get_figheight())
    yukseklik = 0.24
    for j, ((sutun, ad), renk) in enumerate(zip(olcutler, G.SERI)):
        konum = np.arange(n) + (j - 1) * (yukseklik + 0.04)
        ax.barh(konum, grup_onemi[sutun], height=yukseklik, color=renk, label=ad, edgecolor="none")
    ax.set_yticks(range(n), list(grup_onemi.index))
    ax.set_ylim(n - 0.5, -0.5)
    ax.xaxis.set_major_formatter(G.YUZDE_EKSENI)
    ax.grid(axis="x")
    ax.spines["bottom"].set_visible(False)
    ax.legend(loc="lower left", bbox_to_anchor=(0, 1.0), ncols=3, fontsize=9, borderaxespad=0.3, handlelength=1.4,
              columnspacing=1.2)
    G.basliklar(fig, "Değişken gruplarının önemi, üç ölçüye göre",
                "Her ölçüde toplam önemin payı; test verisi. Ölçüler aynı sırayı veriyorsa sonuç ölçüye bağlı değildir")
    G.kaydet(fig, "12_shap_grup_onemi.png")


def main():
    baslik_yaz("SHAP analizi")
    metrikler = json.loads(A.MODEL_METRIKLERI.read_text(encoding="utf-8"))
    df, X_egitim, X_test, y_egitim, y_test = M.veri_bol()
    lojistik, agac = modelleri_kur(X_egitim, y_egitim, metrikler["en_iyi_parametreler"])

    # Bütünlük denetimi: yeniden kurulan model, model.py'nin raporladığı modelle aynı sonucu vermeli.
    p_lojistik = lojistik.predict_proba(X_test)[:, 1]
    assert abs(roc_auc_score(y_test, p_lojistik) - metrikler["test"][M.LOJISTIK]["roc_auc"]) < 1e-9

    # --- 1. Lojistik regresyon: doğrusal SHAP -------------------------------------------------
    on, model = lojistik.named_steps["on_isleme"], lojistik.named_steps["model"]
    egitim_matrisi, test_matrisi = yogun(on.transform(X_egitim)), yogun(on.transform(X_test))
    kaynak = kaynak_degiskenler(on)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        masker = shap.maskers.Independent(egitim_matrisi, max_samples=len(egitim_matrisi))  # eğitim verisinin tamamı
        aciklama = shap.LinearExplainer(model, masker)(test_matrisi)
    taban = np.asarray(aciklama.base_values, dtype=float).reshape(-1)
    sutun_katkilari = np.asarray(aciklama.values, dtype=float)
    denetim = {
        "dogrusal_formul_farki": float(np.abs(sutun_katkilari - model.coef_[0] * (test_matrisi - egitim_matrisi.mean(0))).max()),
        "toplamsallik_farki": float(np.abs(taban + sutun_katkilari.sum(1) - model.decision_function(test_matrisi)).max()),
    }
    assert max(denetim.values()) < 1e-8, denetim
    lr_shap = degiskene_topla(sutun_katkilari, kaynak, X_test.index)

    # --- 2. Karar ağacı: TreeSHAP (ikinci görüş) ----------------------------------------------
    ag_on, ag_model = agac.named_steps["on_isleme"], agac.named_steps["model"]
    ag_matris = yogun(ag_on.transform(X_test))
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        agac_aciklayici = shap.TreeExplainer(ag_model)
        ham = agac_aciklayici.shap_values(ag_matris)
    ham = ham[1] if isinstance(ham, list) else (ham[..., 1] if np.ndim(ham) == 3 else ham)
    denetim["agac_toplamsallik_farki"] = float(np.abs(
        np.atleast_1d(agac_aciklayici.expected_value)[-1] + ham.sum(1) - ag_model.predict_proba(ag_matris)[:, 1]).max())
    assert denetim["agac_toplamsallik_farki"] < 1e-8, denetim
    agac_shap = degiskene_topla(ham, kaynak_degiskenler(ag_on), X_test.index)

    # --- 3. Permütasyon önemi: SHAP'ten bağımsız, başarıdaki düşüşe bakar ----------------
    izin = permutation_importance(lojistik, X_test, y_test, scoring="roc_auc", n_repeats=30, random_state=A.RASTGELE_TOHUM)
    izin_onemi = pd.Series(izin.importances_mean, index=X_test.columns)

    onem = pd.DataFrame({
        "lojistik_shap": lr_shap.abs().mean(), "agac_shap": agac_shap.abs().mean(), "izin_onemi": izin_onemi,
    }).sort_values("lojistik_shap", ascending=False)
    onem["lojistik_pay"] = onem["lojistik_shap"] / onem["lojistik_shap"].sum()
    onem["agac_pay"] = onem["agac_shap"] / onem["agac_shap"].sum()
    for sutun in ("lojistik_shap", "agac_shap", "izin_onemi"):
        onem[f"{sutun}_sira"] = onem[sutun].rank(ascending=False, method="min").astype(int)
    onem.insert(0, "degisken_adi", [DEGISKEN_ADI[d] for d in onem.index])
    onem.rename_axis("degisken").to_csv(A.TABLO / "shap_degisken_onemi.csv", encoding="utf-8-sig")
    korelasyon = {
        "lojistik_shap__agac_shap": float(spearmanr(onem["lojistik_shap"], onem["agac_shap"]).correlation),
        "lojistik_shap__izin_onemi": float(spearmanr(onem["lojistik_shap"], onem["izin_onemi"]).correlation),
        "agac_shap__izin_onemi": float(spearmanr(onem["agac_shap"], onem["izin_onemi"]).correlation),
    }
    ilk5 = {sutun: list(onem.sort_values(sutun, ascending=False).index[:5]) for sutun in ("lojistik_shap", "agac_shap", "izin_onemi")}

    # --- 3b. Bağımlı değişkenler gruplanınca: aynı üç ölçü grup düzeyinde ----------------------
    grup_toplam = lambda tablo: pd.DataFrame({ad: tablo[sutunlar].sum(axis=1) for ad, sutunlar in GRUPLAR.items()})
    lr_grup, agac_grup = grup_toplam(lr_shap), grup_toplam(agac_shap)
    grup_onemi = pd.DataFrame({
        "lojistik_shap": lr_grup.abs().mean(), "agac_shap": agac_grup.abs().mean(),
        "izin_onemi": grup_izin_onemi(lojistik, X_test, y_test),
    })
    for sutun in ("lojistik_shap", "agac_shap", "izin_onemi"):
        grup_onemi[f"{sutun}_pay"] = grup_onemi[sutun].clip(lower=0) / grup_onemi[sutun].clip(lower=0).sum()
        grup_onemi[f"{sutun}_sira"] = grup_onemi[sutun].rank(ascending=False, method="min").astype(int)
    grup_onemi = grup_onemi.sort_values("lojistik_shap", ascending=False)
    grup_onemi.rename_axis("grup").to_csv(A.TABLO / "shap_grup_onemi.csv", encoding="utf-8-sig")

    # Toplam ücretin SHAP işareti ile tek başına kayıp oranı karşılaştırması
    ceyrek = deger_gruplari(X_test, "TotalCharges")
    siralama = ["en düşük çeyrek", "2. çeyrek", "3. çeyrek", "en yüksek çeyrek"]
    celiski = pd.DataFrame({
        "musteri": ceyrek.value_counts(), "kayip_orani": y_test.groupby(ceyrek).mean(),
        "ortalama_shap": lr_shap["TotalCharges"].groupby(ceyrek).mean(),
    }).reindex(siralama).rename_axis("ceyrek").reset_index()

    # --- 4. Bağımlı değişken sorunu: nicel kanıt ---------------------------------------------
    hizmetler = pd.get_dummies(X_egitim[HIZMET_SUTUNLARI]).astype(float)
    ucret_r2 = LinearRegression().fit(hizmetler, X_egitim["MonthlyCharges"]).score(hizmetler, X_egitim["MonthlyCharges"])
    toplam_korelasyon = float(np.corrcoef(X_egitim["TotalCharges"], X_egitim["tenure"] * X_egitim["MonthlyCharges"])[0, 1])
    ayni_sutun = int(pd.DataFrame(egitim_matrisi).T.duplicated().sum())
    bagimlilik = {"aylik_ucret_r2_hizmetlerle": float(ucret_r2), "toplam_ucret_korelasyonu_sure_x_ucret": toplam_korelasyon,
                  "birebir_ayni_bir_sicak_sutun": ayni_sutun, "bir_sicak_sutun_sayisi": int(egitim_matrisi.shape[1])}

    # --- 5. Değerlere göre ortalama SHAP (yön) ------------------------------------------------
    en_onemliler = [d for d in onem.index if d not in TURETILMIS][:EN_ONEMLI_SAYISI]
    yon = []
    for degisken in en_onemliler:
        gruplar = deger_gruplari(X_test, degisken)
        for grup, katki in lr_shap[degisken].groupby(gruplar):
            yon.append({"degisken": degisken, "grup": grup, "etiket": grup_etiketi(degisken, grup),
                        "musteri": int((gruplar == grup).sum()), "ortalama_shap": float(katki.mean())})
    yon = pd.DataFrame(yon)

    # --- 6. Tek müşteri ----------------------------------------------------------------------
    en_riskli = int(np.argmax(p_lojistik))
    kimlik = df.at[X_test.index[en_riskli], "customerID"]
    satir, musteri = lr_shap.iloc[en_riskli], X_test.iloc[en_riskli]
    grup_katkisi = lr_grup.iloc[en_riskli].sort_values(key=np.abs, ascending=False)
    ozetler = {
        "Süre ve ücret": f"{int(musteri['tenure'])} ay · aylık {tr(musteri['MonthlyCharges'], 2)} {A.PARA} · toplam {tr(musteri['TotalCharges'])} {A.PARA}",
        "Hizmet paketi": f"{M.DEGER_ADI.get(musteri['InternetService'], musteri['InternetService'])} · "
                         f"güvenlik {M.DEGER_ADI.get(musteri['OnlineSecurity'], musteri['OnlineSecurity'])} · "
                         f"destek {M.DEGER_ADI.get(musteri['TechSupport'], musteri['TechSupport'])}",
        "Sözleşme": M.DEGER_ADI.get(musteri["Contract"], musteri["Contract"]),
        "Faturalandırma ve ödeme": M.DEGER_ADI.get(musteri["PaymentMethod"], musteri["PaymentMethod"]),
        "Demografik": f"{M.DEGER_ADI.get(musteri['gender'], musteri['gender'])} · {M.YAS_GRUBU[int(musteri['SeniorCitizen'])]}",
    }
    ornek_etiket = [f"{grup}\n{ozetler[grup]}" for grup in grup_katkisi.index]
    ornek_deger = [float(d) for d in grup_katkisi.values]
    ornek = {"kimlik": kimlik, "olasilik": float(p_lojistik[en_riskli]), "taban_log_oran": float(taban[en_riskli]),
             "toplam_katki": float(satir.sum()), "log_oran": float(taban[en_riskli] + satir.sum()),
             "grup_katkilari": [{"grup": g, "ozet": ozetler[g], "shap": float(d)} for g, d in grup_katkisi.items()],
             "degisken_katkilari": {DEGISKEN_ADI[d]: float(satir[d]) for d in GRUPLAR["Süre ve ücret"]}}

    # --- Grafikler ---------------------------------------------------------------------------
    ilk10 = onem.head(10)
    G.cubuk_grafigi(
        [DEGISKEN_ADI[d] + ("*" if d in TURETILMIS else "") for d in ilk10.index], ilk10["lojistik_shap"],
        "Tahmini en çok etkileyen değişkenler (SHAP)",
        "Test verisi, ortalama |SHAP| (log-oran); * süre ve ücretten türetilebilir (açık renk)",
        "11_shap_genel_onem.png", bicim=lambda d: tr(d, 2), eksen_bicimi=G.ONDALIK_EKSENI,
        renkler=[G.ACIK_RENK_TURETILMIS if d in TURETILMIS else G.SERI[0] for d in ilk10.index])
    grup_karsilastirma_grafigi(grup_onemi)

    yon_sirali = pd.concat([
        yon[yon["degisken"] == d].sort_values("ortalama_shap", ascending=False) for d in en_onemliler], ignore_index=True)
    G.cubuk_grafigi(
        list(yon_sirali["etiket"]), yon_sirali["ortalama_shap"],
        "Değerlere göre ortalama SHAP etkisi",
        "Kırmızı: ayrılma riskini artırır, mavi: azaltır (log-oran); türetilmiş ücret değişkenleri bu grafikte yok",
        "13_shap_deger_yonu.png", bicim=isaretli, eksen_bicimi=G.ONDALIK_EKSENI,
        renkler=[G.ARTIRAN if d > 0 else G.AZALTAN for d in yon_sirali["ortalama_shap"]])

    G.cubuk_grafigi(
        ornek_etiket, ornek_deger,
        f"Tek müşterinin riski nereden geliyor? ({kimlik})",
        f"Ortalama log-oran {tr(ornek['taban_log_oran'], 2)} + katkılar {isaretli(ornek['toplam_katki'])} = {tr(ornek['log_oran'], 2)} "
        f"(ayrılma olasılığı {yuzde(ornek['olasilik'], 0)}); kırmızı artırır, mavi azaltır",
        "14_shap_ornek_musteri.png", bicim=isaretli, eksen_bicimi=G.ONDALIK_EKSENI,
        renkler=[G.ARTIRAN if d > 0 else G.AZALTAN for d in ornek_deger])

    A.SHAP_OZETI.write_text(json.dumps({
        "shap_surumu": shap.__version__, "denetim": denetim, "korelasyon": korelasyon, "ilk5": ilk5,
        "bagimlilik": bagimlilik, "ornek_musteri": ornek,
        "gruplar": {ad: sutunlar for ad, sutunlar in GRUPLAR.items()}, "turetilmis": TURETILMIS,
        "grup_onemi": json.loads(grup_onemi.rename_axis("grup").reset_index().to_json(orient="records", force_ascii=False)),
        "toplam_ucret_celiskisi": json.loads(celiski.to_json(orient="records", force_ascii=False)),
        "degisken_onemi": json.loads(onem.rename_axis("degisken").reset_index().to_json(orient="records", force_ascii=False)),
        "deger_yonu": json.loads(yon_sirali.to_json(orient="records", force_ascii=False)),
    }, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"SHAP {shap.__version__}; doğrusal formül farkı {denetim['dogrusal_formul_farki']:.1e}, toplamsallık farkı {denetim['toplamsallik_farki']:.1e}")
    print("En önemli 6 değişken (lojistik SHAP): " + ", ".join(
        f"{DEGISKEN_ADI[d]} {tr(v, 2)}" for d, v in onem['lojistik_shap'].head(6).items()))
    print("Sıralama uyumu (Spearman): " + ", ".join(f"{k.replace('__', ' ↔ ')} {tr(v, 2)}" for k, v in korelasyon.items()))
    print(f"Aylık ücretin hizmet sütunlarıyla R² = {tr(ucret_r2, 2)}; toplam ücret ↔ süre × ücret korelasyonu {tr(toplam_korelasyon, 3)}; "
          f"birebir aynı bir-sıcak sütun: {ayni_sutun}/{egitim_matrisi.shape[1]}")
    print("Grup önemi (paylar): " + "; ".join(
        f"{g}: SHAP {yuzde(s['lojistik_shap_pay'], 0)} / ağaç {yuzde(s['agac_shap_pay'], 0)} / permütasyon {yuzde(s['izin_onemi_pay'], 0)}"
        for g, s in grup_onemi.iterrows()))
    print("Toplam ücret çeyreklerinde (kayıp oranı → ortalama SHAP): " + "; ".join(
        f"{s['ceyrek']} {yuzde(s['kayip_orani'], 0)} → {isaretli(s['ortalama_shap'])}" for _, s in celiski.iterrows()))
    print(f"Örnek müşteri {kimlik}: olasılık {yuzde(ornek['olasilik'], 0)}")
    print("Yazıldı: outputs/shap_ozeti.json, outputs/tablolar/shap_degisken_onemi.csv, shap_grup_onemi.csv, grafik 11–14")


if __name__ == "__main__":
    main()
