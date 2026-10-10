"""Aşama 6-7: tahmin modelleri, değerlendirme ve müşteri risk tablosu.

Lojistik regresyon temel modeldir; karar ağacı karşılaştırma içindir.
Test verisi model seçiminde kullanılmaz: ayarlar ve eşik eğitim verisi içinde
çapraz doğrulamayla belirlenir, test verisine yalnızca en sonda bir kez bakılır.

Çalıştırma:  py src/model.py
"""
import json

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import sklearn
from sklearn.calibration import calibration_curve
from sklearn.compose import ColumnTransformer
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import (accuracy_score, average_precision_score, brier_score_loss, confusion_matrix, f1_score,
                             precision_recall_curve, precision_score, recall_score, roc_auc_score, roc_curve)
from sklearn.model_selection import GridSearchCV, StratifiedKFold, cross_val_predict, cross_validate, train_test_split
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import OneHotEncoder, StandardScaler
from sklearn.tree import DecisionTreeClassifier

import ayarlar as A
import grafik_stili as G
from ortak import baslik_yaz, tr, yuzde

LOJISTIK, AGAC = "Lojistik regresyon", "Karar ağacı"
ADAYLAR = {
    LOJISTIK: (LogisticRegression(max_iter=2000), {"model__C": [0.01, 0.1, 1, 10]}),
    AGAC: (DecisionTreeClassifier(random_state=A.RASTGELE_TOHUM),
           {"model__max_depth": [3, 4, 5, 6, 8], "model__min_samples_leaf": [20, 50, 100]}),
}
CV_OLCUTLERI = {"roc_auc": "roc_auc", "pr_auc": "average_precision", "kesinlik": "precision",
                "duyarlilik": "recall", "f1": "f1", "dogruluk": "accuracy"}

# Risk işareti metinleri
OZELLIK_ADI = {
    "Contract": "Sözleşme", "InternetService": "İnternet", "PaymentMethod": "Ödeme",
    "OnlineSecurity": "Çevrimiçi güvenlik", "OnlineBackup": "Çevrimiçi yedekleme",
    "DeviceProtection": "Cihaz koruma", "TechSupport": "Teknik destek", "StreamingTV": "TV yayını",
    "StreamingMovies": "Film yayını", "PaperlessBilling": "Kâğıtsız fatura",
    "PhoneService": "Telefon hizmeti", "MultipleLines": "Çoklu hat",
}
DEGER_ADI = {
    "Month-to-month": "aylık (taahhütsüz)", "One year": "1 yıllık", "Two year": "2 yıllık",
    "Fiber optic": "fiber", "Yes": "var", "No": "yok",
    "No internet service": "internet yok", "No phone service": "telefon yok",
    "Electronic check": "elektronik çek", "Mailed check": "posta çeki",
    "Bank transfer (automatic)": "otomatik banka havalesi", "Credit card (automatic)": "otomatik kredi kartı",
}


def boru_hatti(model):
    """Ön işleme + model. Ölçekleme ve kodlama yalnızca eğitim verisinden öğrenilir."""
    on_isleme = ColumnTransformer([
        ("sayisal", StandardScaler(), A.SAYISAL),
        ("kategorik", OneHotEncoder(handle_unknown="ignore"), A.KATEGORIK),
    ])
    return Pipeline([("on_isleme", on_isleme), ("model", model)])


def esik_sec(y, olasilik):
    """F1 skorunu en yüksek yapan olasılık eşiği."""
    kesinlik, duyarlilik, esikler = precision_recall_curve(y, olasilik)
    f1 = 2 * kesinlik * duyarlilik / np.maximum(kesinlik + duyarlilik, 1e-12)
    return float(esikler[np.argmax(f1[:-1])])


def esik_olcutleri(y, olasilik, esik):
    tahmin = (olasilik >= esik).astype(int)
    return {
        "esik": round(float(esik), 4),
        "dogruluk": accuracy_score(y, tahmin),
        "kesinlik": precision_score(y, tahmin),
        "duyarlilik": recall_score(y, tahmin),
        "f1": f1_score(y, tahmin),
        "karisiklik_matrisi": confusion_matrix(y, tahmin).tolist(),  # [[TN, FP], [FN, TP]]
    }


def gruplara_ayir(X):
    """Kategorik sütunlar aynen, tenure ve aylık ücret aralıklara bölünmüş olarak."""
    gruplar = X[A.KATEGORIK].copy()
    gruplar["tenure"] = pd.cut(X["tenure"], A.SURE_SINIR, labels=A.SURE_ETIKET).astype(str)
    gruplar["MonthlyCharges"] = pd.cut(X["MonthlyCharges"], A.UCRET_SINIR, labels=A.UCRET_ETIKET, right=False).astype(str)
    return gruplar


def risk_isaretleri(X_egitim, y_egitim, X, en_fazla=3, kat=1.25):
    """Müşterinin içinde bulunduğu, eğitim verisinde kayıp oranı genel oranın en az `kat` katı olan gruplar.

    Betimleyici bir özettir: modelin iç hesabının dökümü değildir ve ayrılmanın nedenini kanıtlamaz.
    (Lojistik regresyon katsayıları bu iş için kullanılmadı; aylık ücret hizmet sütunlarının neredeyse
    doğrusal bileşimi olduğundan katsayılar tek tek yorumlanamıyor.)
    """
    egitim_gruplari, gruplar = gruplara_ayir(X_egitim), gruplara_ayir(X)
    oranlar = pd.DataFrame({s: gruplar[s].map(y_egitim.groupby(egitim_gruplari[s]).mean()) for s in gruplar.columns})
    sinir = kat * y_egitim.mean()

    def metin(ozellik, deger):
        if ozellik == "tenure":
            return f"Abonelik süresi: {deger}"
        if ozellik == "MonthlyCharges":
            return f"Aylık ücret: {deger} {A.PARA}"
        return f"{OZELLIK_ADI[ozellik]}: {DEGER_ADI.get(deger, deger)}"

    metinler = []
    for kimlik, satir in oranlar.iterrows():
        one_cikan = satir[satir >= sinir].sort_values(ascending=False).head(en_fazla)
        metinler.append("; ".join(metin(o, gruplar.at[kimlik, o]) for o in one_cikan.index) or "Belirgin risk işareti yok")
    return pd.Series(metinler, index=X.index)


def egitim_araligi_disinda(X_egitim, X):
    """Eğitim verisinde görülmemiş bir değere sahip müşterileri işaretler."""
    disinda = pd.Series(False, index=X.index)
    for sutun in A.SAYISAL:
        disinda |= (X[sutun] < X_egitim[sutun].min()) | (X[sutun] > X_egitim[sutun].max())
    for sutun in A.KATEGORIK:
        disinda |= ~X[sutun].isin(X_egitim[sutun].unique())
    return disinda


def egri_grafikleri(y, olasiliklar, taban_oran):
    fig, (sol, sag) = plt.subplots(1, 2, figsize=(9.6, 4.6))
    fig.subplots_adjust(left=0.07, right=0.97, top=0.74, bottom=0.13, wspace=0.24)
    for (ad, p), renk in zip(olasiliklar.items(), G.SERI):
        yanlis_pozitif, dogru_pozitif, _ = roc_curve(y, p)
        kesinlik, duyarlilik, _ = precision_recall_curve(y, p)
        sol.plot(yanlis_pozitif, dogru_pozitif, color=renk, label=f"{ad} (ROC-AUC {tr(roc_auc_score(y, p), 3)})")
        sag.plot(duyarlilik, kesinlik, color=renk, label=f"{ad} (PR-AUC {tr(average_precision_score(y, p), 3)})")
    sol.plot([0, 1], [0, 1], color=G.SOLUK, linewidth=1)
    sol.annotate("Rastgele sıralama", (0.62, 0.62), xytext=(6, -10), textcoords="offset points", fontsize=9, color=G.IKINCIL)
    sag.axhline(taban_oran, color=G.SOLUK, linewidth=1)
    sag.annotate(f"Kayıp oranı {yuzde(taban_oran)}", (0.0, taban_oran), xytext=(2, 4), textcoords="offset points",
                 fontsize=9, color=G.IKINCIL)
    sol.set(xlabel="Yanlış pozitif oranı", ylabel="Duyarlılık (recall)", xlim=(0, 1), ylim=(0, 1.02))
    sag.set(xlabel="Duyarlılık (recall)", ylabel="Kesinlik (precision)", xlim=(0, 1), ylim=(0, 1.02))
    for ax, ad in ((sol, "ROC eğrisi"), (sag, "Kesinlik–duyarlılık eğrisi")):
        ax.grid(True)
        ax.xaxis.set_major_formatter(G.ONDALIK_EKSENI)
        ax.yaxis.set_major_formatter(G.ONDALIK_EKSENI)
        ax.set_title(ad, loc="left", fontsize=10, color=G.IKINCIL)
        ax.legend(loc="lower right" if ax is sol else "upper right", fontsize=9)
    G.basliklar(fig, "Modellerin sıralama becerisi", "Test verisi; eğri sol üste (ROC) ve sağ üste (PR) yaklaştıkça model daha iyi")
    G.kaydet(fig, "06_roc_pr_egrileri.png")


def kalibrasyon_grafigi(y, olasiliklar):
    fig, ax = plt.subplots(figsize=(6.4, 5.2))
    fig.subplots_adjust(left=0.12, right=0.95, top=0.80, bottom=0.12)
    ax.plot([0, 1], [0, 1], color=G.SOLUK, linewidth=1)
    ax.annotate("Tam uyum", (0.84, 0.84), xytext=(6, -12), textcoords="offset points", fontsize=9, color=G.IKINCIL)
    for (ad, p), renk in zip(olasiliklar.items(), G.SERI):
        gerceklesen, tahmin = calibration_curve(y, p, n_bins=10, strategy="quantile")
        ax.plot(tahmin, gerceklesen, color=renk, marker="o", markersize=6, markeredgecolor=G.YUZEY, markeredgewidth=1.5,
                label=f"{ad} (Brier {tr(brier_score_loss(y, p), 3)})")
    ax.set(xlabel="Tahmin edilen ayrılma olasılığı", ylabel="Gerçekleşen ayrılma oranı", xlim=(0, 1), ylim=(0, 1))
    ax.xaxis.set_major_formatter(G.YUZDE_EKSENI)
    ax.yaxis.set_major_formatter(G.YUZDE_EKSENI)
    ax.grid(True)
    ax.legend(loc="upper left", fontsize=9)
    G.basliklar(fig, "Olasılıkların kalibrasyonu", "Test verisi, 10 eşit büyüklükte grup; çizgiye yakınlık olasılıkların güvenilirliğini gösterir")
    G.kaydet(fig, "07_kalibrasyon.png")


def grup_kontrolu(test, esik):
    """Model girdisi olmayan demografik gruplarda sonuçların benzer olup olmadığını denetler."""
    satirlar = []
    for sutun in A.GRUP_KONTROL:
        for deger, grup in test.groupby(sutun):
            tahmin = (grup["churn_olasiligi"] >= esik).astype(int)
            satirlar.append({
                "grup_degiskeni": sutun, "grup": deger, "musteri_sayisi": len(grup),
                "gercek_kayip_orani": grup["gercek_churn"].mean(),
                "ortalama_tahmin": grup["churn_olasiligi"].mean(),
                "yuksek_risk_payi": tahmin.mean(),
                "duyarlilik": recall_score(grup["gercek_churn"], tahmin),
                "kesinlik": precision_score(grup["gercek_churn"], tahmin),
            })
    return pd.DataFrame(satirlar)


def main():
    baslik_yaz("Model")
    df = pd.read_csv(A.TEMIZ_VERI)
    X, y = df[A.SAYISAL + A.KATEGORIK], df[A.HEDEF].eq("Yes").astype(int)
    X_egitim, X_test, y_egitim, y_test = train_test_split(
        X, y, test_size=A.TEST_ORANI, stratify=y, random_state=A.RASTGELE_TOHUM)
    katlar = StratifiedKFold(n_splits=A.CV_KAT, shuffle=True, random_state=A.RASTGELE_TOHUM)

    modeller, cv_sonuc, test_sonuc, esikler, parametreler, test_olasilik = {}, {}, {}, {}, {}, {}
    for ad, (model, izgara) in ADAYLAR.items():
        arama = GridSearchCV(boru_hatti(model), izgara, scoring="roc_auc", cv=katlar).fit(X_egitim, y_egitim)
        boru = arama.best_estimator_  # tüm eğitim verisiyle yeniden eğitilmiş hâli
        parametreler[ad] = arama.best_params_
        skorlar = cross_validate(boru, X_egitim, y_egitim, cv=katlar, scoring=CV_OLCUTLERI)
        cv_sonuc[ad] = {olcut: float(skorlar[f"test_{olcut}"].mean()) for olcut in CV_OLCUTLERI}
        # Eşik, eğitim verisindeki kat-dışı tahminlerden seçilir; test verisi görülmez.
        kat_disi = cross_val_predict(boru, X_egitim, y_egitim, cv=katlar, method="predict_proba")[:, 1]
        esikler[ad] = esik_sec(y_egitim, kat_disi)

        p = boru.predict_proba(X_test)[:, 1]
        modeller[ad], test_olasilik[ad] = boru, p
        test_sonuc[ad] = {
            "roc_auc": roc_auc_score(y_test, p),
            "pr_auc": average_precision_score(y_test, p),
            "brier": brier_score_loss(y_test, p),
            "esik_0_50": esik_olcutleri(y_test, p, 0.5),
            "secilen_esik": esik_olcutleri(y_test, p, esikler[ad]),
        }

    esik = esikler[LOJISTIK]
    if cv_sonuc[AGAC]["roc_auc"] > cv_sonuc[LOJISTIK]["roc_auc"] + 0.01:
        print("UYARI: Karar ağacı çapraz doğrulamada belirgin biçimde daha iyi; temel model seçimini gözden geçirin.")

    # --- Müşteri risk tablosu: modelin eğitimde görmediği test müşterileri "güncel portföy" sayılır
    test = df.loc[X_test.index].copy()
    test["churn_olasiligi"] = test_olasilik[LOJISTIK]
    test["gercek_churn"] = y_test
    test["risk_grubu"] = np.select(
        [test["churn_olasiligi"] >= esik, test["churn_olasiligi"] >= esik / 2], ["Yüksek", "Orta"], "Düşük")
    test["risk_isaretleri"] = risk_isaretleri(X_egitim, y_egitim, X_test)
    test["egitim_araligi_disinda"] = egitim_araligi_disinda(X_egitim, X_test)
    risk = test.sort_values("churn_olasiligi", ascending=False)[
        ["customerID", "churn_olasiligi", "risk_grubu", "risk_isaretleri", "tenure", "Contract", "MonthlyCharges",
         "InternetService", "TechSupport", "TotalCharges_dolduruldu", "egitim_araligi_disinda", "gercek_churn"]]
    risk.to_csv(A.RISK_TABLOSU, index=False, encoding="utf-8-sig")

    gruplar = grup_kontrolu(test, esik)
    A.TABLO.mkdir(parents=True, exist_ok=True)
    gruplar.to_csv(A.TABLO / "grup_bazli_kontrol.csv", index=False, encoding="utf-8-sig")

    egri_grafikleri(y_test, test_olasilik, y_test.mean())
    kalibrasyon_grafigi(y_test, test_olasilik)

    A.MODEL_METRIKLERI.write_text(json.dumps({
        "temel_model": LOJISTIK,
        "veri": {"egitim": len(X_egitim), "test": len(X_test),
                 "egitim_kayip_orani": float(y_egitim.mean()), "test_kayip_orani": float(y_test.mean())},
        "rastgele_tohum": A.RASTGELE_TOHUM,
        "secilen_esik": esik,
        "risk_gruplari": {"Yüksek": f">= {esik:.3f}", "Orta": f">= {esik / 2:.3f}", "Düşük": f"< {esik / 2:.3f}"},
        "en_iyi_parametreler": parametreler,
        "capraz_dogrulama": cv_sonuc,
        "test": test_sonuc,
        "surumler": {"scikit-learn": sklearn.__version__, "pandas": pd.__version__, "numpy": np.__version__},
    }, ensure_ascii=False, indent=2), encoding="utf-8")

    # --- Konsol özeti
    print(f"Eğitim: {tr(len(X_egitim))} müşteri, test: {tr(len(X_test))} müşteri (kayıp oranı {yuzde(y_test.mean())})")
    for ad in ADAYLAR:
        c, t, s = cv_sonuc[ad], test_sonuc[ad], test_sonuc[ad]["secilen_esik"]
        print(f"\n{ad}  {parametreler[ad]}")
        print(f"  Çapraz doğrulama: ROC-AUC {tr(c['roc_auc'], 3)}, PR-AUC {tr(c['pr_auc'], 3)}")
        print(f"  Test: ROC-AUC {tr(t['roc_auc'], 3)}, PR-AUC {tr(t['pr_auc'], 3)}, Brier {tr(t['brier'], 3)}")
        print(f"  Test, eşik 0,50: kesinlik {yuzde(t['esik_0_50']['kesinlik'])}, duyarlılık {yuzde(t['esik_0_50']['duyarlilik'])}, "
              f"F1 {tr(t['esik_0_50']['f1'], 3)}, doğruluk {yuzde(t['esik_0_50']['dogruluk'])}")
        print(f"  Test, eşik {tr(s['esik'], 2)}: kesinlik {yuzde(s['kesinlik'])}, duyarlılık {yuzde(s['duyarlilik'])}, "
              f"F1 {tr(s['f1'], 3)}, doğruluk {yuzde(s['dogruluk'])}")
    print(f"\nRisk grupları (eşik {tr(esik, 2)}): " + ", ".join(
        f"{g} {tr(n)}" for g, n in risk["risk_grubu"].value_counts().reindex(["Yüksek", "Orta", "Düşük"]).items()))
    print(f"Eğitim aralığı dışında değer taşıyan müşteri: {int(risk['egitim_araligi_disinda'].sum())}")
    print("Yazıldı: outputs/musteri_risk_tablosu.csv, outputs/model_metrikleri.json, outputs/tablolar/grup_bazli_kontrol.csv")


if __name__ == "__main__":
    main()
