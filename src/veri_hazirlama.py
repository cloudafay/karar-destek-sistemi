"""Aşama 2-4: ham veriyi denetler, temizler; kalite raporunu ve veri sözlüğünü yazar.

Çalıştırma:  py src/veri_hazirlama.py
"""
import hashlib
from datetime import datetime

import pandas as pd

import ayarlar as A
from ortak import baslik_yaz, tr, yuzde

# Sütun açıklamaları (veri sözlüğü için). Değer istatistikleri veriden hesaplanır.
ACIKLAMA = {
    "customerID": "Müşteri kimliği",
    "gender": "Cinsiyet",
    "SeniorCitizen": "65 yaş ve üzeri mi (1 = evet)",
    "Partner": "Eşi/partneri var mı",
    "Dependents": "Bakmakla yükümlü olduğu kişi var mı",
    "tenure": "Müşterinin şirkette bulunduğu ay sayısı",
    "PhoneService": "Telefon hizmeti var mı",
    "MultipleLines": "Birden fazla hattı var mı",
    "InternetService": "İnternet hizmet türü",
    "OnlineSecurity": "Çevrimiçi güvenlik ek hizmeti",
    "OnlineBackup": "Çevrimiçi yedekleme ek hizmeti",
    "DeviceProtection": "Cihaz koruma ek hizmeti",
    "TechSupport": "Teknik destek ek hizmeti",
    "StreamingTV": "TV yayın hizmeti",
    "StreamingMovies": "Film yayın hizmeti",
    "Contract": "Sözleşme türü",
    "PaperlessBilling": "Kâğıtsız fatura kullanıyor mu",
    "PaymentMethod": "Ödeme yöntemi",
    "MonthlyCharges": f"Aylık ücret ({A.PARA})",
    "TotalCharges": f"Bugüne kadarki toplam ücret ({A.PARA})",
    "Churn": "Müşteri ayrıldı mı (hedef değişken)",
}


def ham_veriyi_oku():
    # Boş TotalCharges hücreleri görünsün diye her sütun metin olarak okunur.
    return pd.read_csv(A.HAM_VERI, dtype=str, keep_default_na=False)


def kalite_kontrolu(ham):
    """Ham veri üzerindeki kontrollerin sonuçlarını sözlük olarak döndürür."""
    bos = ham.apply(lambda s: s.str.strip().eq("").sum())
    bos_toplam_ucret = ham[ham["TotalCharges"].str.strip() == ""]
    return {
        "satir": len(ham),
        "sutun": ham.shape[1],
        "tekrar_eden_kimlik": int(ham["customerID"].duplicated().sum()),
        "tekrar_eden_satir": int(ham.duplicated().sum()),
        "bos_hucreler": ", ".join(f"{sutun} ({adet})" for sutun, adet in bos[bos > 0].items()),
        "bos_toplam_ucret": len(bos_toplam_ucret),
        "bos_toplam_ucret_tenure": ", ".join(sorted(bos_toplam_ucret["tenure"].unique())),
        "churn_dagilimi": ham["Churn"].value_counts().to_dict(),
    }


def temizle(ham):
    """Tipleri düzeltir ve boş TotalCharges değerlerini doldurur. Ham dosyaya dokunmaz."""
    df = ham.copy()
    for sutun in df.columns:
        df[sutun] = df[sutun].str.strip()
    df["tenure"] = df["tenure"].astype(int)
    df["SeniorCitizen"] = df["SeniorCitizen"].astype(int)
    df["MonthlyCharges"] = df["MonthlyCharges"].astype(float)

    toplam = pd.to_numeric(df["TotalCharges"], errors="coerce")
    df["TotalCharges_dolduruldu"] = toplam.isna()
    # Boş değerlerin tamamı tenure = 0 olan, henüz fatura kesilmemiş yeni müşterilerdir.
    if not (df.loc[toplam.isna(), "tenure"] == 0).all():
        raise ValueError("Boş TotalCharges içeren ama tenure > 0 olan kayıt var; doldurma kuralını gözden geçirin.")
    df["TotalCharges"] = toplam.fillna(0.0)
    return df


def kalite_raporu_yaz(kontrol, temiz):
    dosya = A.HAM_VERI
    ozet = hashlib.sha256(dosya.read_bytes()).hexdigest()
    tarih = datetime.fromtimestamp(dosya.stat().st_mtime).strftime("%d.%m.%Y %H:%M")
    churn = kontrol["churn_dagilimi"]
    oran = churn["Yes"] / kontrol["satir"]
    boyut_uyumu = (kontrol["satir"], kontrol["sutun"]) == (A.BEKLENEN_SATIR, A.BEKLENEN_SUTUN)

    satirlar = [
        "# Veri kaynağı kaydı ve kalite raporu",
        "",
        "Bu dosya `src/veri_hazirlama.py` tarafından üretilir; elle düzenlemeyin.",
        "",
        "## Kaynak kaydı",
        "",
        f"- Dosya: `data/raw/{dosya.name}`",
        f"- Boyut: {tr(dosya.stat().st_size)} bayt",
        f"- Dosya tarihi (indirme): {tarih}",
        f"- SHA-256: `{ozet}`",
        "- Veri seti: IBM Telco Customer Churn (Kaggle: blastchar/telco-customer-churn)",
        f"- İndirilen kaynak: {A.INDIRILEN_KAYNAK}",
        "",
        "## Kontroller",
        "",
        f"- Boyut: {tr(kontrol['satir'])} satır × {kontrol['sutun']} sütun; kaynak açıklamasıyla "
        f"({tr(A.BEKLENEN_SATIR)} × {A.BEKLENEN_SUTUN}) {'uyumlu' if boyut_uyumu else 'UYUMSUZ'}.",
        f"- Tekrar eden customerID: {kontrol['tekrar_eden_kimlik']}; tümüyle aynı satır: {kontrol['tekrar_eden_satir']}.",
        f"- Boş hücre içeren sütunlar: {kontrol['bos_hucreler'] or 'yok'}.",
        f"- TotalCharges ham dosyada metin olarak gelir; {kontrol['bos_toplam_ucret']} kayıtta boştur. "
        f"Bu kayıtların tenure değeri: {kontrol['bos_toplam_ucret_tenure']}.",
        f"- Churn dağılımı: Yes = {tr(churn['Yes'])}, No = {tr(churn['No'])}; ayrılma oranı {yuzde(oran)}. "
        "Sınıflar dengesiz olduğu için yalnızca doğruluk oranına bakılmaz.",
        "",
        "## Uygulanan dönüşümler",
        "",
        "- Tüm metin hücrelerinin baş ve sonundaki boşluklar silindi.",
        "- tenure ve SeniorCitizen tam sayıya, MonthlyCharges ve TotalCharges ondalık sayıya çevrildi.",
        f"- Boş {kontrol['bos_toplam_ucret']} TotalCharges değeri 0 ile dolduruldu "
        "(hepsi tenure = 0 olan, henüz faturası oluşmamış müşteriler). "
        "Bu kayıtlar `TotalCharges_dolduruldu` sütunuyla işaretlendi.",
        "- Satır silinmedi; ham dosya değiştirilmedi.",
        "",
        "## Sayısal alanların özeti (temiz veri)",
        "",
        "| Alan | En küçük | Ortanca | Ortalama | En büyük |",
        "|---|---|---|---|---|",
    ]
    for sutun in A.SAYISAL:
        s = temiz[sutun]
        satirlar.append(f"| {sutun} | {tr(s.min(), 2)} | {tr(s.median(), 2)} | {tr(s.mean(), 2)} | {tr(s.max(), 2)} |")
    (A.CIKTI / "veri_kalite_raporu.md").write_text("\n".join(satirlar) + "\n", encoding="utf-8")


def veri_sozlugu_yaz(temiz):
    satirlar = [
        "# Veri sözlüğü",
        "",
        "Bu dosya `src/veri_hazirlama.py` tarafından üretilir; elle düzenlemeyin.",
        "",
        "| Sütun | Açıklama | Tip | Değerler | Modeldeki rolü |",
        "|---|---|---|---|---|",
    ]
    for sutun, aciklama in ACIKLAMA.items():
        s = temiz[sutun]
        if sutun == "customerID":
            tip, degerler = "metin", f"{tr(s.nunique())} benzersiz kod"
        elif sutun in A.SAYISAL:
            tip, degerler = "sayı", f"{tr(s.min(), 2)} – {tr(s.max(), 2)}"
        else:
            tip, degerler = "kategori", ", ".join(str(d) for d in sorted(s.unique()))
        if sutun == A.HEDEF:
            rol = "Hedef değişken; girdi olarak kullanılmaz."
        elif sutun in A.MODEL_DISI:
            rol = "Girdi değil. " + A.MODEL_DISI[sutun]
        else:
            rol = "Girdi."
        satirlar.append(f"| {sutun} | {aciklama} | {tip} | {degerler} | {rol} |")
    satirlar += [
        "",
        "Temizleme sırasında eklenen sütun: `TotalCharges_dolduruldu` "
        "(ham dosyada TotalCharges boşsa True). Model girdisi değildir.",
    ]
    (A.DOKUMAN / "veri_sozlugu.md").write_text("\n".join(satirlar) + "\n", encoding="utf-8")


def main():
    baslik_yaz("Veri hazırlama")
    ham = ham_veriyi_oku()
    kontrol = kalite_kontrolu(ham)
    temiz = temizle(ham)

    A.TEMIZ_VERI.parent.mkdir(parents=True, exist_ok=True)
    temiz.to_csv(A.TEMIZ_VERI, index=False)
    kalite_raporu_yaz(kontrol, temiz)
    veri_sozlugu_yaz(temiz)

    churn = kontrol["churn_dagilimi"]
    print(f"Boyut: {tr(kontrol['satir'])} satır × {kontrol['sutun']} sütun")
    print(f"Tekrar eden customerID: {kontrol['tekrar_eden_kimlik']}")
    print(f"Boş TotalCharges: {kontrol['bos_toplam_ucret']} (tenure değeri: {kontrol['bos_toplam_ucret_tenure']})")
    print(f"Churn: Yes = {tr(churn['Yes'])}, No = {tr(churn['No'])} ({yuzde(churn['Yes'] / kontrol['satir'])})")
    print(f"Yazıldı: {A.TEMIZ_VERI.relative_to(A.KOK)}, outputs/veri_kalite_raporu.md, docs/veri_sozlugu.md")


if __name__ == "__main__":
    main()
