"""Aşama 2-4: ham veriyi denetler, temizler; kalite raporunu ve veri sözlüğünü yazar.

Çalıştırma:  py src/veri_hazirlama.py
"""
import hashlib
import json
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


def tekrar_anahtari(df):
    """Bir kaydın "aynı" sayılması için karşılaştırılan sütunlar: kimlik ve etiket hariç bütün özellikler."""
    return [sutun for sutun in df.columns if sutun not in A.TEKRAR_ANAHTARI_HARIC]


def tekrar_kayitlari_ayikla(df):
    """Özellikleri birebir aynı olan kayıt gruplarından ilkini (dosya sırasına göre) tutar, kalanları çıkarır.

    Döndürür: (modelleme verisi, özet sözlüğü, çıkarılan kayıtlar tablosu).
    Etiketi çelişen gruplarda da ilk kayıt tutulur; çelişki ayrıca sayılır ve raporlanır.
    """
    anahtar = tekrar_anahtari(df)
    df = df.assign(_grup=df.groupby(anahtar, sort=False, dropna=False).ngroup())
    tekrar = df.duplicated(subset=anahtar, keep="first")
    tutulan = df[~tekrar].set_index("_grup")
    cikarilan = df[tekrar].assign(
        korunan_kimlik=lambda d: d["_grup"].map(tutulan["customerID"]),
        korunan_etiket=lambda d: d["_grup"].map(tutulan[A.HEDEF]),
    )
    cikarilan["etiket_ayni_mi"] = cikarilan[A.HEDEF] == cikarilan["korunan_etiket"]

    grup_boyutu = df.groupby("_grup").size()
    tekrar_gruplari = grup_boyutu[grup_boyutu > 1].index
    etiket_sayisi = df[df["_grup"].isin(tekrar_gruplari)].groupby("_grup")[A.HEDEF].nunique()
    modelleme = df[~tekrar].drop(columns="_grup").reset_index(drop=True)
    ozet = {
        "anahtar_sutunlari": anahtar,
        "kayit_oncesi": len(df),
        "kayit_sonrasi": len(modelleme),
        "tekrar_grubu": int(len(tekrar_gruplari)),
        "gruplardaki_toplam_kayit": int(grup_boyutu[tekrar_gruplari].sum()),
        "cikarilan_kayit": int(tekrar.sum()),
        "etiketi_celisen_grup": int((etiket_sayisi > 1).sum()),
        "etiket_dahil_birebir_ayni": int(df.duplicated(subset=anahtar + [A.HEDEF]).sum()),
        "cikarilan_etiketi_ayni": int(cikarilan["etiket_ayni_mi"].sum()),
        "cikarilan_etiketi_farkli": int((~cikarilan["etiket_ayni_mi"]).sum()),
        "cikarilan_tenure": sorted(int(t) for t in cikarilan["tenure"].unique()),
        "cikarilan_ayrilan": int(cikarilan[A.HEDEF].eq("Yes").sum()),
        "ayni_kimlik": int(df["customerID"].duplicated().sum()),
    }
    sutunlar = ["customerID", "korunan_kimlik", "tenure", "MonthlyCharges", "Contract", "InternetService", A.HEDEF,
                "korunan_etiket", "etiket_ayni_mi"]
    return modelleme, ozet, cikarilan[sutunlar].reset_index(drop=True)


def kalite_raporu_yaz(kontrol, temiz, tekrar):
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
        "- Temizleme sırasında satır silinmedi; ham dosya değiştirilmedi.",
        "",
        "## Tekrar eden kayıtlar",
        "",
        f"- Karşılaştırılan sütunlar: kimlik ve etiket dışındaki {len(tekrar['anahtar_sutunlari'])} özellik.",
        f"- {tr(tekrar['tekrar_grubu'])} grupta toplam {tr(tekrar['gruplardaki_toplam_kayit'])} kayıt birebir aynı özellikleri taşıyor; "
        f"her gruptan ilk kayıt tutulup {tr(tekrar['cikarilan_kayit'])} kayıt çıkarıldı "
        f"({tr(tekrar['cikarilan_etiketi_ayni'])} tanesinin etiketi tutulan kayıtla aynı, "
        f"{tr(tekrar['cikarilan_etiketi_farkli'])} tanesinin farklı; etiketi çelişen grup sayısı {tr(tekrar['etiketi_celisen_grup'])}).",
        f"- Çıkarılan kayıtların tenure değerleri: {', '.join(str(t) for t in tekrar['cikarilan_tenure'])} ay.",
        f"- Modelleme verisi: {tr(tekrar['kayit_sonrasi'])} kayıt (`data/processed/telco_modelleme.csv`). "
        "Çıkarılan kayıtların listesi `outputs/tablolar/tekrar_eden_kayitlar.csv` dosyasındadır.",
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

    modelleme, tekrar, cikarilan = tekrar_kayitlari_ayikla(temiz)
    # Doğrulama: çıkarma sonrasında aynı özelliklere sahip iki kayıt kalmamış olmalı.
    assert not modelleme.duplicated(subset=tekrar_anahtari(modelleme)).any()

    A.TEMIZ_VERI.parent.mkdir(parents=True, exist_ok=True)
    A.TABLO.mkdir(parents=True, exist_ok=True)
    temiz.to_csv(A.TEMIZ_VERI, index=False)
    modelleme.to_csv(A.MODELLEME_VERISI, index=False)
    cikarilan.to_csv(A.TEKRAR_KAYITLAR, index=False, encoding="utf-8-sig")
    A.TEKRAR_OZETI.write_text(json.dumps(tekrar, ensure_ascii=False, indent=2), encoding="utf-8")
    kalite_raporu_yaz(kontrol, temiz, tekrar)
    veri_sozlugu_yaz(temiz)

    churn = kontrol["churn_dagilimi"]
    print(f"Boyut: {tr(kontrol['satir'])} satır × {kontrol['sutun']} sütun")
    print(f"Tekrar eden customerID: {kontrol['tekrar_eden_kimlik']}")
    print(f"Boş TotalCharges: {kontrol['bos_toplam_ucret']} (tenure değeri: {kontrol['bos_toplam_ucret_tenure']})")
    print(f"Tekrar eden kayıt (kimlik ve etiket dışındaki {len(tekrar['anahtar_sutunlari'])} özellik aynı): "
          f"{tr(tekrar['tekrar_grubu'])} grup, {tr(tekrar['cikarilan_kayit'])} kayıt çıkarıldı "
          f"(etiketi çelişen grup: {tr(tekrar['etiketi_celisen_grup'])}); modelleme verisi {tr(tekrar['kayit_sonrasi'])} kayıt")
    print(f"Churn: Yes = {tr(churn['Yes'])}, No = {tr(churn['No'])} ({yuzde(churn['Yes'] / kontrol['satir'])})")
    print(f"Yazıldı: {A.TEMIZ_VERI.relative_to(A.KOK)}, outputs/veri_kalite_raporu.md, docs/veri_sozlugu.md")


if __name__ == "__main__":
    main()
