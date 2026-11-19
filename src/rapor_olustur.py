"""Aşama 12: proje raporunu (Word) üretir.

Rapordaki bütün sayılar outputs/ altındaki çıktılardan ve aynı hesap fonksiyonlarından okunur;
varsayımlar değişince betiği yeniden çalıştırmak raporu günceller. Word kuruluysa içindekiler
tablosu güncellenir ve raporun PDF kopyası da yazılır.

Çalıştırma:  py src/rapor_olustur.py   (önce hepsini_calistir.py çalışmış olmalı)
"""
import hashlib
import json
import subprocess
import sys
from datetime import datetime
from importlib.metadata import version

import numpy as np
import pandas as pd
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from sklearn.calibration import calibration_curve

import ayarlar as A
from optimizasyon import butce_ozeti, kampanya_listesi, kisitsiz_plan, optimize_et
from ortak import baslik_yaz, para, tr, yuzde
from senaryo_karsilastirma import ETKI_DUZEYLERI, duyarlilik_tablosu, senaryo_tablosu
from shap_analizi import DEGISKEN_ADI, isaretli
from teklif_senaryosu import teklif_tablosu
from veri_hazirlama import ACIKLAMA, ham_veriyi_oku, kalite_kontrolu
from yonetici_ozeti import ozet_metni

KOYU_MAVI, IKINCIL, CIZGI, BASLIK_DOLGU = "0D366B", "52514E", "C3C2B7", "EEF3FA"
YAZILIMLAR = ["pandas", "numpy", "scikit-learn", "scipy", "matplotlib", "shap", "PuLP", "highspy", "streamlit", "python-docx"]
CINSIYET = {"Female": "Kadın", "Male": "Erkek"}
YAS = {"0": "65 yaş altı", "1": "65 yaş ve üzeri"}


# --- Word yardımcıları -------------------------------------------------------------
class Rapor:
    """python-docx belgesini saran, başlık / paragraf / tablo / şekil ekleyen küçük yardımcı."""

    def __init__(self):
        self.doc = Document()
        self.sekil_no = self.tablo_no = 0
        self.no = {}  # metinde anılan tabloların numaraları
        bolum = self.doc.sections[0]
        bolum.page_width, bolum.page_height = Cm(21), Cm(29.7)  # A4
        bolum.left_margin = bolum.right_margin = bolum.top_margin = bolum.bottom_margin = Cm(2.5)

        normal = self.doc.styles["Normal"]
        normal.font.name, normal.font.size = "Calibri", Pt(11)
        normal.paragraph_format.space_after = Pt(7)
        normal.paragraph_format.line_spacing = 1.15
        dil = OxmlElement("w:lang")
        dil.set(qn("w:val"), "tr-TR")
        normal.element.get_or_add_rPr().append(dil)
        for ad, punto in (("Heading 1", 16), ("Heading 2", 12.5)):
            stil = self.doc.styles[ad]
            stil.font.name, stil.font.size, stil.font.bold = "Calibri", Pt(punto), True
            stil.font.color.rgb = RGBColor.from_string(KOYU_MAVI)
            for tema in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):  # şablonun tema yazı tipini bırak
                stil.element.rPr.rFonts.attrib.pop(qn(f"w:{tema}"), None)
        self.doc.styles["Heading 1"].paragraph_format.space_before = Pt(18)
        self.doc.styles["Heading 1"].paragraph_format.space_after = Pt(8)
        self.doc.styles["Heading 2"].paragraph_format.space_before = Pt(12)
        self.doc.styles["Heading 2"].paragraph_format.space_after = Pt(5)
        aciklama = self.doc.styles["Caption"]
        aciklama.font.size, aciklama.font.bold, aciklama.font.italic = Pt(9.5), False, True
        aciklama.font.color.rgb = RGBColor.from_string(IKINCIL)

        bolum.different_first_page_header_footer = True  # kapakta sayfa numarası yok
        alt_bilgi = bolum.footer.paragraphs[0]
        alt_bilgi.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self._alan(alt_bilgi, "PAGE")

    @staticmethod
    def _alan(paragraf, kod, yer_tutucu="1"):
        """Paragrafa bir Word alanı (sayfa numarası, içindekiler) ekler."""
        for tur, etiket, metin in (("begin", None, None), (None, "w:instrText", kod), ("separate", None, None),
                                   (None, "w:t", yer_tutucu), ("end", None, None)):
            oge = OxmlElement(etiket or "w:fldChar")
            if tur:
                oge.set(qn("w:fldCharType"), tur)
            else:
                oge.set(qn("xml:space"), "preserve")
                oge.text = metin
            paragraf.add_run()._r.append(oge)

    def baslik(self, metin, duzey=1):
        self.doc.add_heading(metin, level=duzey)

    def p(self, metin, giris=None, italik=False, girinti=False):
        """Paragraf. `giris` verilirse paragrafın başına kalın yazılır."""
        paragraf = self.doc.add_paragraph()
        if giris:
            paragraf.add_run(giris + " ").bold = True
        paragraf.add_run(metin).italic = italik
        if girinti:
            paragraf.paragraph_format.left_indent = Cm(1)
        return paragraf

    def maddeler(self, *metinler):
        for metin in metinler:
            paragraf = self.doc.add_paragraph(metin, style="List Bullet")
            paragraf.paragraph_format.space_after = Pt(3)

    def sayfa_sonu(self):
        self.doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    def sekil(self, dosya, aciklama, genislik_cm=15):
        self.sekil_no += 1
        paragraf = self.doc.add_paragraph()
        paragraf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraf.paragraph_format.keep_with_next = True
        paragraf.paragraph_format.space_after = Pt(2)
        paragraf.add_run().add_picture(str(dosya), width=Cm(genislik_cm))
        alt = self.doc.add_paragraph(f"Şekil {self.sekil_no}. {aciklama}", style="Caption")
        alt.alignment = WD_ALIGN_PARAGRAPH.CENTER
        return self.sekil_no

    def tablo(self, aciklama, basliklar, satirlar, genislikler_cm, saga=(), ad=None):
        """Başlık satırı gölgeli tablo. `saga`: sağa hizalanacak sütunlar; `ad`: metinden anmak için etiket."""
        self.tablo_no += 1
        if ad:
            self.no[ad] = self.tablo_no
        ust = self.doc.add_paragraph(f"Tablo {self.tablo_no}. {aciklama}", style="Caption")
        ust.paragraph_format.keep_with_next = True
        ust.paragraph_format.space_after = Pt(3)

        tablo = self.doc.add_table(rows=1 + len(satirlar), cols=len(basliklar))
        tablo.style, tablo.alignment, tablo.autofit = "Table Grid", WD_TABLE_ALIGNMENT.CENTER, False
        kenarlar = OxmlElement("w:tblBorders")
        for kenar in ("top", "left", "bottom", "right", "insideH", "insideV"):
            oge = OxmlElement(f"w:{kenar}")
            for ozellik, deger in (("val", "single"), ("sz", "4"), ("space", "0"), ("color", CIZGI)):
                oge.set(qn(f"w:{ozellik}"), deger)
            kenarlar.append(oge)
        tablo._tbl.tblPr.append(kenarlar)

        for i, satir in enumerate([basliklar] + [list(s) for s in satirlar]):
            for j, deger in enumerate(satir):
                hucre = tablo.cell(i, j)
                hucre.width = Cm(genislikler_cm[j])
                paragraf = hucre.paragraphs[0]
                paragraf.paragraph_format.space_after = Pt(0)
                paragraf.paragraph_format.line_spacing = 1.0
                if j in saga:
                    paragraf.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                yazi = paragraf.add_run(str(deger))
                yazi.font.size, yazi.bold = Pt(9.5), i == 0
                if i == 0:
                    dolgu = OxmlElement("w:shd")
                    for ozellik, d in (("val", "clear"), ("color", "auto"), ("fill", BASLIK_DOLGU)):
                        dolgu.set(qn(f"w:{ozellik}"), d)
                    hucre._tc.get_or_add_tcPr().append(dolgu)
        if len(satirlar) <= 14:  # küçük tablo sayfalar arasında bölünmesin
            for satir in tablo.rows[:-1]:
                for hucre in satir.cells:
                    for paragraf in hucre.paragraphs:
                        paragraf.paragraph_format.keep_with_next = True
        for satir in tablo.rows:  # satır sayfa sonunda ikiye bölünmesin
            satir._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
        tekrar = OxmlElement("w:tblHeader")  # başlık satırı her sayfada yinelensin
        tablo.rows[0]._tr.get_or_add_trPr().append(tekrar)
        self.doc.add_paragraph().paragraph_format.space_after = Pt(2)
        return self.tablo_no


def word_ile_tamamla(dosya):
    """Word kuruluysa içindekiler tablosunu günceller ve PDF kopyasını yazar; değilse sessizce geçer."""
    pdf = dosya.with_suffix(".pdf")
    # -20 ve -21: Word'ün yerleşik "İçindekiler 1" ve "İçindekiler 2" stilleri; satır aralığı sıkılaştırılır.
    betik = f"""
        $w = New-Object -ComObject Word.Application; $w.Visible = $false; $w.DisplayAlerts = 0
        try {{
            $d = $w.Documents.Open('{dosya}')
            foreach ($stil in -20, -21) {{
                $b = $d.Styles.Item($stil).ParagraphFormat; $b.SpaceBefore = 0; $b.SpaceAfter = 2; $b.LineSpacingRule = 0
            }}
            foreach ($t in $d.TablesOfContents) {{ $t.Update() }}
            $d.Save(); $d.ExportAsFixedFormat('{pdf}', 17); $d.Close($false)
        }} finally {{ if ($w.Documents.Count -eq 0) {{ $w.Quit() }} }}
    """
    try:
        sonuc = subprocess.run(["powershell", "-NoProfile", "-Command", betik], capture_output=True, timeout=180)
        return sonuc.returncode == 0 and pdf.exists()
    except (OSError, subprocess.TimeoutExpired):
        return False


# --- Veriler -----------------------------------------------------------------------
def butce_kompozisyonu(teklifler):
    """Farklı bütçe düzeylerinde önerilen planın teklif türlerine göre bileşimi (kapasite sabit)."""
    satirlar = []
    for butce in (2_500, 5_000, A.BUTCE, 20_000, 30_000):
        secim = optimize_et(teklifler, butce, A.DESTEK_KAPASITESI)
        sayilar = secim["offer_type"].value_counts().reindex(A.TEKLIFLER, fill_value=0)
        satirlar.append({"butce": butce, "musteri": len(secim), **{t: int(sayilar[t]) for t in A.TEKLIFLER}})
    return pd.DataFrame(satirlar)


def verileri_topla():
    """Raporda kullanılan bütün tabloları ve sayıları tek sözlükte toplar."""
    ham = ham_veriyi_oku()
    temiz = pd.read_csv(A.TEMIZ_VERI)
    risk = pd.read_csv(A.RISK_TABLOSU)
    teklifler = teklif_tablosu(risk)
    liste = kampanya_listesi(risk, teklifler, optimize_et(teklifler))
    senaryolar = senaryo_tablosu(risk, teklifler, A.BUTCE)
    ayrildi = temiz[A.HEDEF].eq("Yes")
    return {
        "kontrol": kalite_kontrolu(ham),
        "temiz": temiz,
        "kesif": pd.read_csv(A.TABLO / "kesif_kayip_oranlari.csv").set_index(["degisken", "grup"]),
        "internet": ayrildi.groupby(temiz["InternetService"]).mean(),
        "metrik": json.loads(A.MODEL_METRIKLERI.read_text(encoding="utf-8")),
        "risk": risk,
        "gruplar": pd.read_csv(A.TABLO / "grup_bazli_kontrol.csv", dtype={"grup": str}),
        "teklifler": teklifler,
        "liste": liste,
        "ozet": butce_ozeti(liste, A.BUTCE),
        "kisitsiz": kisitsiz_plan(teklifler),
        "yillik_gelir": risk["MonthlyCharges"].sum() * 12,
        "senaryolar": senaryolar,
        "duyarlilik": duyarlilik_tablosu(senaryolar),
        "egri": pd.read_csv(A.TABLO / "butce_egrisi.csv", index_col="butce"),
        "butce_kompozisyonu": butce_kompozisyonu(teklifler),
        "tekrar": json.loads(A.TEKRAR_OZETI.read_text(encoding="utf-8")),
        "tekrar_kayitlar": pd.read_csv(A.TEKRAR_KAYITLAR),
        "cv": pd.read_csv(A.CV_TABLOSU),
        "shap": json.loads(A.SHAP_OZETI.read_text(encoding="utf-8")),
    }


# --- Bölümler ----------------------------------------------------------------------
def kapak_ve_icindekiler(r):
    for _ in range(5):
        r.doc.add_paragraph()
    for metin, punto, kalin, renk in (
        ("KARAR DESTEK SİSTEMİ", 26, True, KOYU_MAVI),
        ("Müşteri Kaybı Riskini Tahmin Etme ve Kampanya Bütçesini Optimize Etme", 16, False, "0B0B0B"),
        ("Proje Raporu", 14, False, IKINCIL),
    ):
        paragraf = r.doc.add_paragraph()
        paragraf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        yazi = paragraf.add_run(metin)
        yazi.font.size, yazi.bold = Pt(punto), kalin
        yazi.font.color.rgb = RGBColor.from_string(renk)
    for _ in range(6):
        r.doc.add_paragraph()
    for metin in (f"Hazırlayan: {A.RAPOR_HAZIRLAYAN}", f"Ders: {A.RAPOR_DERS}",
                  f"Öğretim elemanı: {A.RAPOR_OGRETIM_ELEMANI}", datetime.now().strftime("%d.%m.%Y")):
        paragraf = r.doc.add_paragraph(metin)
        paragraf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraf.paragraph_format.space_after = Pt(2)
    r.sayfa_sonu()

    paragraf = r.doc.add_paragraph()
    yazi = paragraf.add_run("İçindekiler")
    yazi.font.size, yazi.bold = Pt(16), True
    yazi.font.color.rgb = RGBColor.from_string(KOYU_MAVI)
    r._alan(r.doc.add_paragraph(), 'TOC \\o "1-2" \\h \\z \\u',
            "İçindekiler tablosu için bu satıra sağ tıklayıp “Alanı Güncelleştir” seçeneğini kullanın.")
    r.sayfa_sonu()


def yonetici_ozeti(r, v):
    r.baslik("1. Yönetici özeti")
    r.p("Bu proje, abonelik modeliyle çalışan bir telekom işletmesinin iki sorusuna birlikte yanıt veren bir karar destek "
        "sistemi prototipidir: hangi müşterilerin ayrılma riski yüksektir ve sınırlı kampanya bütçesi hangi müşteriye, "
        "hangi teklifle harcanmalıdır?")
    r.p("Sistem dört katmandan oluşur: müşteri kaybı tahmini, her müşteri için teklif seçeneklerinin ekonomik "
        "değerlendirmesi, bütçe kısıtı altında optimizasyon ve yöneticiye öneriyi gerekçesi ve belirsizliğiyle birlikte "
        "sunan karar ekranı.")
    for paragraf in ozet_metni(v["risk"], v["ozet"], v["senaryolar"], v["metrik"]["secilen_esik"]).split("\n\n"):
        r.p(paragraf)
    m, t = v["metrik"], v["tekrar"]
    loj = "Lojistik regresyon"
    guven = m["test_guven_araligi"][loj]["roc_auc"]
    r.p(f"Veri %{round((1 - A.TEST_ORANI) * 100)} eğitim ve %{round(A.TEST_ORANI * 100)} test olarak ayrılmıştır. Eğitim kısmında "
        f"{A.CV_KAT} katlı, iç içe çapraz doğrulama uygulanmış (lojistik regresyon ROC-AUC {tr(m['capraz_dogrulama'][loj]['roc_auc'], 3)} "
        f"± {tr(m['capraz_dogrulama_std'][loj]['roc_auc'], 3)}); test kısmında aynı ölçüt {tr(m['test'][loj]['roc_auc'], 3)} "
        f"(%95 bootstrap aralığı {tr(guven[0], 3)}–{tr(guven[1], 3)}) bulunmuştur. Doğrulamadan önce, kimlik ve etiket dışındaki bütün "
        f"özellikleri aynı olan {tr(t['cikarilan_kayit'])} tekrar eden kayıt çıkarılmıştır. Modelin açıklaması için SHAP kullanılmış; "
        "sonuçların birbirinden türetilebilen değişkenler nedeniyle nasıl yorumlanması gerektiği Bölüm 6.8'de anlatılmıştır.",
        giris="Doğrulama ve veri işleme:")
    r.p("Müşteri verisi gerçektir (IBM Telco Customer Churn). Teklif maliyetleri, tekliflerin etkisi, bütçe ve kapasite "
        "ise sentetik varsayımdır; veri setinde kampanya bilgisi yoktur. Bu nedenle net fayda sayıları yöntemin nasıl "
        "çalıştığını gösterir, gerçek bir kampanyanın sonucunu kanıtlamaz.", giris="Temel sınır:")


def problem(r):
    r.baslik("2. Problem tanımı ve kapsam")
    r.p("Abonelik tabanlı işletmeler müşterilerinin bir bölümünü fiyat, rakip teklifleri, hizmet kalitesi veya müşteri "
        "hizmetleri deneyimi nedeniyle kaybeder. İşletme iptal başvurusundan sonra harekete geçerse müşteriyi geri kazanmak "
        "zorlaşır. Öte yandan herkese indirim vermek de doğru değildir: bazı müşteriler zaten kalacaktır, bazıları indirim "
        "verilse de ayrılacaktır. Bütçenin yanlış müşterilere harcanması kârlılığı düşürür.")
    r.p("Karar vericinin yanıt aradığı sorular şunlardır:")
    r.maddeler(
        "Hangi müşteriler için ayrılma riski yüksektir?",
        "Riskli müşteride öne çıkan işaretler nelerdir: sözleşme türü mü, ücret mi, hizmet mi?",
        "Hangi müdahale seçeneği bu müşteri için ekonomik olarak anlamlıdır?",
        "Kampanya bütçesi hangi müşteri ve tekliflere ayrılmalıdır?",
        "Önerinin beklenen faydası nedir ve ne kadar belirsizlik içerir?",
    )
    r.baslik("2.1. Kapsam ve sınırlar", 2)
    r.p("Prototip telekomünikasyon alanı için kurulmuştur. Müşteri kaybı tahmini açık bir eğitim veri setiyle yapılır. "
        "Kampanya etkisi ve teklif maliyetleri için gerçek veri bulunmadığından ayrı ve açıkça “sentetik” olarak "
        "işaretlenmiş bir senaryo tablosu kullanılır. Belirli bir teklifin müşteriyi kesin olarak elde tutacağı iddia "
        "edilmez; sistem nihai kararı vermez, yöneticiye öneri sunar.")
    r.baslik("2.2. Başarı ölçütleri", 2)
    r.maddeler(
        "Tahmin katmanı: görülmemiş test verisinde duyarlılık, kesinlik, F1, ROC-AUC, PR-AUC ve kalibrasyon.",
        "Karar katmanı: aynı bütçe ve aynı varsayımlar altında beklenen net faydanın mevcut yöntemle karşılaştırılması.",
        "Şeffaflık: sentetik varsayımların gerçek verilerden ayrılması ve sonuçların varsayımlara duyarlılığının gösterilmesi.",
    )


def veri(r, v):
    k, dosya = v["kontrol"], A.HAM_VERI
    r.baslik("3. Veri kaynağı ve veri sözlüğü")
    r.baslik("3.1. Kaynak kaydı", 2)
    r.p("Ana veri seti IBM Telco Customer Churn örnek verisidir. Her satır bir müşteriyi temsil eder; hizmetler, abonelik "
        "süresi, sözleşme ve faturalandırma bilgilerini ve müşterinin ayrılıp ayrılmadığını gösteren Churn sütununu içerir.")
    r.tablo("Veri kaynağı kaydı", ["Bilgi", "Değer"], [
        ("Veri seti", "IBM Telco Customer Churn"),
        ("İndirilen kaynak", A.INDIRILEN_KAYNAK),
        ("Dosya", f"data/raw/{dosya.name}"),
        ("Dosya tarihi (indirme)", datetime.fromtimestamp(dosya.stat().st_mtime).strftime("%d.%m.%Y %H:%M")),
        ("Boyut", f"{tr(dosya.stat().st_size)} bayt; {tr(k['satir'])} satır × {k['sutun']} sütun"),
        ("SHA-256", hashlib.sha256(dosya.read_bytes()).hexdigest()),
    ], [4.2, 11.8], ad="kaynak")
    r.p("Veri seti geçmiş müşteri özellikleri ile ayrılma durumu arasındaki örüntüleri incelemeye uygundur. Hangi müşteriye "
        "hangi kampanyanın sunulduğunu ve kampanyanın sonucu değiştirip değiştirmediğini göstermez. Bu nedenle yalnızca "
        "müşteri kaybı tahmininde kullanılmıştır; kampanya etkisini kanıtlamaz. Veri kesitsel bir örnektir ve gerçek bir "
        "işletmenin güncel müşterilerini temsil etmeyebilir.", giris="Veri seti neyi gösterir, neyi göstermez?")
    r.baslik("3.2. Veri sözlüğü", 2)
    satirlar = []
    for sutun, aciklama in ACIKLAMA.items():
        if sutun == A.HEDEF:
            rol = "Hedef değişken; girdi değildir."
        elif sutun in A.MODEL_DISI:
            rol = "Girdi değil. " + A.MODEL_DISI[sutun]
        else:
            rol = "Girdi."
        satirlar.append((sutun, aciklama, rol))
    r.tablo("Veri sözlüğü ve sütunların modeldeki rolü", ["Sütun", "Açıklama", "Modeldeki rolü"], satirlar, [3.3, 5.7, 7.0], ad="sozluk")
    r.p("Proje önerisinde cinsiyet ve yaş grubunun ayrımcılık riski nedeniyle modele girmemesi planlanmıştı. Ön sunumda alınan "
        "geri bildirim üzerine bu iki alan, müşteri davranışını açıklamada kullanılabilecek bilgiler olduğu için model girdisine "
        "eklenmiştir. Ayrımcılık riskine karşı modelin kadın ve erkek müşterilerde, 65 yaş altı ve üzerindeki müşterilerde nasıl "
        "çalıştığı ayrıca denetlenmektedir (Bölüm 6.7); iki alanın tahmine katkısı Bölüm 6.8'de ölçülmüştür. Eş ve bakmakla "
        "yükümlü olunan kişi bilgisi, iş gerekçesi zayıf olduğu için girdi yapılmamıştır.")


def kalite(r, v):
    k, temiz = v["kontrol"], v["temiz"]
    churn = k["churn_dagilimi"]
    r.baslik("4. Veri kalite kontrolü ve veri işleme")
    r.baslik("4.1. Kontroller ve temizleme", 2)
    r.p("İndirilen dosya modele doğrudan verilmemiş, önce aşağıdaki kontroller yapılmıştır. Kontroller ve dönüşümler "
        "src/veri_hazirlama.py betiğiyle tekrarlanabilir; ham dosya değiştirilmez.")
    r.maddeler(
        f"Boyut: {tr(k['satir'])} satır ve {k['sutun']} sütun; kaynak açıklamasındaki boyutla "
        f"({tr(A.BEKLENEN_SATIR)} × {A.BEKLENEN_SUTUN}) uyumludur.",
        f"Tekrar: customerID alanında tekrar eden kayıt sayısı {k['tekrar_eden_kimlik']}, tümüyle aynı satır sayısı {k['tekrar_eden_satir']}.",
        f"Eksik değer: TotalCharges ham dosyada metin olarak gelir ve {k['bos_toplam_ucret']} kayıtta boştur. Bu kayıtların "
        f"tümünde tenure değeri {k['bos_toplam_ucret_tenure']} olduğundan (henüz faturası oluşmamış yeni müşteriler) boş değerler "
        "0 ile doldurulmuş ve ayrı bir sütunla işaretlenmiştir. Başka boş hücre yoktur; satır silinmemiştir.",
        f"Sınıf dağılımı: {tr(churn['Yes'])} müşteri ayrılmış, {tr(churn['No'])} müşteri kalmıştır; ayrılma oranı "
        f"{yuzde(churn['Yes'] / k['satir'])}. Sınıflar dengesiz olduğu için değerlendirmede yalnızca doğruluk oranına bakılmamıştır.",
        "Tip dönüşümü: tenure ve SeniorCitizen tam sayıya, MonthlyCharges ve TotalCharges ondalık sayıya çevrilmiştir.",
    )
    r.tablo(f"Sayısal alanların özeti (temiz veri; ücretler {A.PARA})", ["Alan", "En küçük", "Ortanca", "Ortalama", "En büyük"],
            [(s, tr(temiz[s].min(), 2), tr(temiz[s].median(), 2), tr(temiz[s].mean(), 2), tr(temiz[s].max(), 2))
             for s in A.SAYISAL], [4.0, 3.0, 3.0, 3.0, 3.0], saga=(1, 2, 3, 4))


def kesif(r, v):
    k = v["kesif"]["kayip_orani"]
    genel = v["kontrol"]["churn_dagilimi"]["Yes"] / v["kontrol"]["satir"]
    r.baslik("5. Keşifsel analiz")
    r.p("Model kurulmadan önce kayıp oranı başlıca müşteri gruplarında karşılaştırılmıştır. Grafiklerdeki yatay ya da dikey "
        f"gri çizgi genel kayıp oranını ({yuzde(genel)}) gösterir. Keşifsel analiz temizlenmiş verinin tamamında "
        f"({tr(v['kontrol']['satir'])} kayıt) yapılmıştır; modelleme ve doğrulama ise tekrar eden kayıtlar çıkarıldıktan sonraki "
        f"{tr(v['tekrar']['kayit_sonrasi'])} kayıtla yapılır (Bölüm 4.2).")
    r.sekil(A.GRAFIK / "01_genel_kayip_orani.png", "Genel müşteri kayıp oranı", 13)
    r.sekil(A.GRAFIK / "02_sozlesme_turune_gore.png", "Sözleşme türüne göre kayıp oranı")
    r.p(f"En belirgin fark sözleşme türündedir: aylık (taahhütsüz) sözleşmede kayıp oranı {yuzde(k['Sözleşme türü', 'Month-to-month'])}, "
        f"bir yıllıkta {yuzde(k['Sözleşme türü', 'One year'])}, iki yıllıkta {yuzde(k['Sözleşme türü', 'Two year'])} düzeyindedir.")
    r.sekil(A.GRAFIK / "03_abonelik_suresine_gore.png", "Abonelik süresine göre kayıp oranı")
    r.p(f"Kayıp ilk yılda yoğunlaşır: ilk 12 ayındaki müşterilerde oran {yuzde(k['Abonelik süresi', '0–12 ay'])}, "
        f"49–72 aydır müşteri olanlarda {yuzde(k['Abonelik süresi', '49–72 ay'])} düzeyindedir.")
    r.sekil(A.GRAFIK / "04_aylik_ucrete_gore.png", "Aylık ücret aralıklarına göre kayıp oranı")
    r.p(f"Aylık ücreti 35 {A.PARA} altında olan müşterilerde kayıp oranı {yuzde(k['Aylık ücret', A.UCRET_ETIKET[0]])} ile en düşüktür; "
        f"75–95 {A.PARA} aralığında {yuzde(k['Aylık ücret', '75–95'])} ile en yüksek düzeye çıkar. İlişki tek yönlü artan değildir.")
    r.sekil(A.GRAFIK / "05_teknik_destege_gore.png", "Teknik destek durumuna göre kayıp oranı")
    r.p(f"Teknik destek hizmeti almayan internet müşterilerinde kayıp oranı {yuzde(k['Teknik destek', 'No'])}, alanlarda "
        f"{yuzde(k['Teknik destek', 'Yes'])} düzeyindedir.")
    r.p("Bu grafikler ilişki gösterir, neden-sonuç ilişkisini kanıtlamaz. Örneğin teknik destek almayanlarda kaybın yüksek "
        "olması, desteğin kendisinden değil bu müşterilerin başka özelliklerinden (sözleşme türü, abonelik süresi) "
        "kaynaklanıyor olabilir. Bu ayrım, teklif etkilerinin neden veriden türetilmeyip varsayım olarak bırakıldığını da açıklar.",
        giris="Yorum sınırı:")


def model(r, v):
    m, risk = v["metrik"], v["risk"]
    esik = m["secilen_esik"]
    loj, agac = "Lojistik regresyon", "Karar ağacı"
    r.baslik("6. Tahmin modeli")
    r.baslik("6.1. Girdiler", 2)
    r.p(f"Hedef değişken Churn sütunudur. Modelde {len(A.SAYISAL)} sayısal ({', '.join(A.SAYISAL)}) ve {len(A.KATEGORIK)} kategorik "
        "alan (cinsiyet ve yaş grubu, hizmetler, sözleşme türü, faturalandırma ve ödeme yöntemi) girdi olarak kullanılmıştır. "
        f"customerID ile eş ve bakmakla yükümlü olunan kişi bilgisi girdi değildir (Tablo {r.no['sozluk']}).")
    r.baslik("6.2. Yöntem", 2)
    r.maddeler(
        f"Tekrar eden kayıtlar çıkarıldıktan sonra kalan {tr(m['veri']['egitim'] + m['veri']['test'])} kayıt katmanlı (stratified) biçimde "
        f"%{round((1 - A.TEST_ORANI) * 100)} eğitim ve %{round(A.TEST_ORANI * 100)} test olarak bölünmüştür: "
        f"{tr(m['veri']['egitim'])} eğitim, {tr(m['veri']['test'])} test müşterisi. Rastgelelik tohumu {m['rastgele_tohum']} olarak kaydedilmiştir.",
        "Ölçekleme ve kategorik kodlama model ile aynı işlem hattındadır; dönüşümler yalnızca eğitim kısmından (çapraz doğrulamada her katın kendi eğitim kısmından) öğrenilir ve doğrulama ya da test kısmına aynen uygulanır.",
        "Lojistik regresyon temel model, karar ağacı karşılaştırma modelidir. Az sayıda modeli doğru değerlendirmek tercih edilmiştir.",
        f"Eğitim kısmında {A.CV_KAT} katlı katmanlı çapraz doğrulama uygulanmıştır. Model ayarları ROC-AUC ölçütüyle seçilmiştir. "
        "Rapor edilen çapraz doğrulama sonuçları iç içe (nested) çapraz doğrulamadan gelir: ayar seçimi de her dış katın kendi eğitim "
        "kısmında yapılır. Seçilmiş modeli aynı katlarla yeniden değerlendirmek hatayı olduğundan düşük gösterirdi (Varma ve Simon, 2006). Seçilen ayarlar: "
        f"lojistik regresyon C = {tr(m['en_iyi_parametreler'][loj]['model__C'], 2)}; karar ağacı en büyük derinlik "
        f"{m['en_iyi_parametreler'][agac]['model__max_depth']}, yaprak başına en az {m['en_iyi_parametreler'][agac]['model__min_samples_leaf']} müşteri.",
        f"Test kısmı ({tr(m['veri']['test'])} müşteri) çapraz doğrulamaya sokulmamıştır: model ve eşik seçiminde kullanılmamış, yalnızca en "
        f"sonda bir kez değerlendirme için kullanılmıştır. Bu kısımdaki ölçütlerin ne kadar oynadığı, {tr(m['bootstrap_tekrar'])} "
        "yeniden örneklemeli bootstrap ile %95 güven aralığı olarak verilmiştir.",
    )
    r.baslik("6.3. Sonuçlar", 2)
    cv_ort_std = lambda ad, olcut: f"{tr(m['capraz_dogrulama'][ad][olcut], 3)} ± {tr(m['capraz_dogrulama_std'][ad][olcut], 3)}"
    satirlar = [("ROC-AUC (iç içe çapraz doğrulama)",) + tuple(cv_ort_std(ad, "roc_auc") for ad in (loj, agac)),
                ("PR-AUC (iç içe çapraz doğrulama)",) + tuple(cv_ort_std(ad, "pr_auc") for ad in (loj, agac))]
    for etiket, anahtar in (("ROC-AUC (test)", "roc_auc"), ("PR-AUC (test)", "pr_auc"), ("Brier skoru (test)", "brier")):
        satirlar.append((etiket,) + tuple(tr(m["test"][ad][anahtar], 3) for ad in (loj, agac)))
    for blok, aciklama in (("secilen_esik", "seçilen eşik"), ("esik_0_50", "eşik 0,50")):
        satirlar.append((f"Eşik ({aciklama})",) + tuple(tr(m["test"][ad][blok]["esik"], 2) for ad in (loj, agac)))
        for etiket, anahtar in (("Kesinlik (precision)", "kesinlik"), ("Duyarlılık (recall)", "duyarlilik"), ("Doğruluk", "dogruluk")):
            satirlar.append((f"{etiket}, {aciklama}",) + tuple(yuzde(m["test"][ad][blok][anahtar]) for ad in (loj, agac)))
        satirlar.append((f"F1, {aciklama}",) + tuple(tr(m["test"][ad][blok]["f1"], 3) for ad in (loj, agac)))
    r.tablo("Model değerlendirme ölçütleri", ["Ölçüt", loj, agac], satirlar, [7.6, 4.2, 4.2], saga=(1, 2), ad="olcut")
    s = m["test"][loj]["secilen_esik"]
    (dn, yp), (yn, dp) = s["karisiklik_matrisi"]
    daha_iyi = m["test"][loj]["roc_auc"] > m["test"][agac]["roc_auc"]
    cv_ayni_yon = (m["capraz_dogrulama"][loj]["roc_auc"] > m["capraz_dogrulama"][agac]["roc_auc"]) == daha_iyi
    r.p(f"Lojistik regresyon test verisinde ROC-AUC {tr(m['test'][loj]['roc_auc'], 3)} ile karar ağacından "
        f"({tr(m['test'][agac]['roc_auc'], 3)}) {'biraz daha iyi' if daha_iyi else 'biraz daha kötü'} sıralama yapmaktadır; "
        f"çapraz doğrulama sonuçları {'da aynı yöndedir' if cv_ayni_yon else 'ise ters yöndedir'}. "
        f"Seçilen eşikte ayrılan müşterilerin {yuzde(s['duyarlilik'])} kadarı yakalanmakta, riskli olarak işaretlenenlerin "
        f"{yuzde(s['kesinlik'])} kadarı gerçekten ayrılmaktadır.")
    r.p(f"Doğruluk oranı tek başına yanıltıcıdır: hiçbir müşteriyi riskli saymayan bir kural test verisinde "
        f"{yuzde(1 - m['veri']['test_kayip_orani'])} doğruluk elde eder ama ayrılan hiçbir müşteriyi yakalamaz. Bu nedenle ölçütler birlikte raporlanmıştır.")
    cv = v["cv"]
    kat_satirlari = []
    for kat in range(1, A.CV_KAT + 1):
        l, a = (cv[(cv["model"] == ad) & (cv["kat"] == kat)].iloc[0] for ad in (loj, agac))
        kat_satirlari.append((str(kat), tr(l["roc_auc"], 3), tr(l["pr_auc"], 3), tr(a["roc_auc"], 3), tr(a["pr_auc"], 3)))
    kat_satirlari.append(("Ortalama ± std",) + tuple(cv_ort_std(ad, olcut) for ad in (loj, agac) for olcut in ("roc_auc", "pr_auc")))
    r.tablo("Çapraz doğrulama katlarında başarı (iç içe, eğitim kısmı)",
            ["Kat", "Lojistik ROC-AUC", "Lojistik PR-AUC", "Ağaç ROC-AUC", "Ağaç PR-AUC"], kat_satirlari,
            [3.4, 3.15, 3.15, 3.15, 3.15], saga=(1, 2, 3, 4), ad="cv_katlar")
    secilen_c = sorted({json.loads(x)["model__C"] for x in cv[cv["model"] == loj]["secilen_ayar"]})
    std_loj = m["capraz_dogrulama_std"][loj]["roc_auc"]
    r.p(f"Başarının katlar arasındaki oynaklığı {'küçüktür' if std_loj < 0.02 else 'belirgindir'} (lojistik regresyon ROC-AUC standart sapması "
        f"{tr(std_loj, 3)}). Lojistik regresyonda katların iç döngülerinde seçilen C değerleri: {', '.join(tr(c, 2) for c in secilen_c)}.")
    gv = m["test_guven_araligi"]
    aralik = lambda ad, olcut, nokta, bicim: f"{bicim(nokta)} [{bicim(gv[ad][olcut][0])}–{bicim(gv[ad][olcut][1])}]"
    uc = lambda x: tr(x, 3)
    r.tablo(f"Test kısmında %95 güven aralıkları ({tr(m['bootstrap_tekrar'])} yeniden örnekleme)",
            ["Model", "ROC-AUC", "PR-AUC", "Kesinlik (seçilen eşik)", "Duyarlılık (seçilen eşik)"],
            [(ad, aralik(ad, "roc_auc", m["test"][ad]["roc_auc"], uc), aralik(ad, "pr_auc", m["test"][ad]["pr_auc"], uc),
              aralik(ad, "kesinlik", m["test"][ad]["secilen_esik"]["kesinlik"], yuzde),
              aralik(ad, "duyarlilik", m["test"][ad]["secilen_esik"]["duyarlilik"], yuzde)) for ad in (loj, agac)],
            [3.0, 3.2, 3.2, 3.3, 3.3], saga=(1, 2, 3, 4), ad="guven")
    cv_loj, test_loj = m["capraz_dogrulama"][loj]["roc_auc"], m["test"][loj]["roc_auc"]
    iceride = gv[loj]["roc_auc"][0] <= cv_loj <= gv[loj]["roc_auc"][1]
    r.p(f"Lojistik regresyonda çapraz doğrulama ortalaması ({tr(cv_loj, 3)}) ile test sonucu ({tr(test_loj, 3)}) arasındaki fark "
        f"{tr(abs(cv_loj - test_loj), 3)} olup çapraz doğrulama ortalaması test kısmının güven aralığının {'içindedir' if iceride else 'dışındadır'}. "
        "Test kısmı model seçiminde kullanılmadığı için çapraz doğrulamaya sokulmamış; bunun yerine yeniden örnekleme ile belirsizliği "
        "gösterilmiştir. Çapraz doğrulama eğitim kısmının, güven aralıkları test kısmının değişkenliğini ölçer.")
    r.tablo(f"Lojistik regresyonun karışıklık matrisi (test verisi, eşik {tr(s['esik'], 2)})",
            ["", "Tahmin: kalır", "Tahmin: ayrılır"],
            [("Gerçek: kaldı", tr(dn), tr(yp)), ("Gerçek: ayrıldı", tr(yn), tr(dp))], [5.0, 4.0, 4.0], saga=(1, 2))
    r.sekil(A.GRAFIK / "06_roc_pr_egrileri.png", "Modellerin ROC ve kesinlik–duyarlılık eğrileri (test verisi)", 16)
    r.baslik("6.4. Kalibrasyon", 2)
    gerceklesen, tahmin = calibration_curve(risk["gercek_churn"], risk["churn_olasiligi"], n_bins=10, strategy="quantile")
    sapma = np.abs(gerceklesen - tahmin)
    en_buyuk = int(sapma.argmax())
    r.p("Ekonomik hesaplarda modelin ürettiği olasılık doğrudan kullanıldığı için olasılıkların gerçek oranlarla uyumu "
        "önemlidir. Test müşterileri tahmin edilen olasılığa göre on eşit gruba ayrılmış ve her grupta gerçekleşen kayıp "
        f"oranı tahminle karşılaştırılmıştır. Lojistik regresyonun Brier skoru {tr(m['test'][loj]['brier'], 3)} düzeyindedir. "
        f"On grubun ortalama mutlak sapması {tr(100 * sapma.mean(), 1)} puan, en büyük sapması {tr(100 * sapma[en_buyuk], 1)} puandır "
        f"(tahmin {yuzde(tahmin[en_buyuk], 0)}, gerçekleşen {yuzde(gerceklesen[en_buyuk], 0)}). Noktalar tam uyum çizgisine "
        f"{'yakındır' if sapma.mean() < 0.05 else 'belirgin biçimde uzaktır'}.")
    r.sekil(A.GRAFIK / "07_kalibrasyon.png", "Tahmin edilen olasılıkların kalibrasyonu (test verisi)", 11)
    r.baslik("6.5. Eşik seçimi ve risk grupları", 2)
    r.p(f"Sınıflandırma eşiği {tr(esik, 2)} olarak seçilmiştir. Bu değer, eğitim verisindeki çapraz doğrulama tahminlerinde F1 "
        "skorunu en yüksek yapan olasılıktır. Ayrılan bir müşteriyi kaçırmak gereksiz bir teklif sunmaktan genellikle daha "
        "maliyetli olduğu için 0,50 yerine daha düşük bir eşik tercih edilmiş, böylece duyarlılık artmıştır. Risk grupları bu eşiğe göre tanımlanmıştır.")
    gruplar = risk.groupby("risk_grubu")["gercek_churn"].agg(["size", "mean"])
    araliklar = {"Yüksek": f"{tr(esik, 2)} ve üzeri", "Orta": f"{tr(esik / 2, 2)} – {tr(esik, 2)}", "Düşük": f"{tr(esik / 2, 2)} altı"}
    r.tablo("Risk grupları ve gruplarda gerçekleşen kayıp oranı (test verisi)",
            ["Risk grubu", "Ayrılma olasılığı", "Müşteri", "Gerçekleşen kayıp oranı"],
            [(g, araliklar[g], tr(gruplar.loc[g, "size"]), yuzde(gruplar.loc[g, "mean"])) for g in ("Yüksek", "Orta", "Düşük")],
            [3.4, 4.6, 3.0, 5.0], saga=(2, 3))
    r.p("Eşik yalnızca raporlama ve risk grupları içindir. Kampanya kararında eşik kullanılmaz; her müşterinin olasılığı "
        "doğrudan ekonomik hesaba girer (Bölüm 7).")
    r.baslik("6.6. Risk işaretleri", 2)
    r.p("Yöneticiye her müşteri için en fazla üç risk işareti gösterilir. İşaretler, müşterinin içinde bulunduğu ve eğitim "
        "verisinde kayıp oranı genel oranın en az 1,25 katı olan gruplardır (örneğin aylık sözleşme, ilk 12 ay, elektronik "
        "çekle ödeme). Bu betimleyici bir özettir; modelin iç hesabının dökümü değildir ve ayrılmanın nedenini kanıtlamaz.")
    r.p("Lojistik regresyon katsayıları bu amaçla kullanılmamıştır: aylık ücret, hizmet sütunlarının neredeyse doğrusal bir "
        f"bileşimi olduğundan (R² = {tr(v['shap']['bagimlilik']['aylik_ucret_r2_hizmetlerle'], 3)}) katsayılar tek tek yorumlanamamakta, "
        "ilk denemede yanıltıcı gerekçeler üretmiştir. Aynı sorunun SHAP açısından sonuçları Bölüm 6.8'de ele alınmıştır.")
    r.baslik("6.7. Grup bazlı kontrol", 2)
    g = v["gruplar"]
    adlar = {"gender": ("Cinsiyet", CINSIYET), "SeniorCitizen": ("Yaş grubu", YAS)}
    r.tablo("Cinsiyet ve yaş grubuna göre model sonuçları (test verisi)",
            ["Değişken", "Grup", "Müşteri", "Gerçek kayıp", "Ortalama tahmin", "Yüksek risk payı", "Duyarlılık", "Kesinlik"],
            [(adlar[s["grup_degiskeni"]][0], adlar[s["grup_degiskeni"]][1][s["grup"]], tr(s["musteri_sayisi"]),
              yuzde(s["gercek_kayip_orani"]), yuzde(s["ortalama_tahmin"]), yuzde(s["yuksek_risk_payi"]),
              yuzde(s["duyarlilik"]), yuzde(s["kesinlik"])) for _, s in g.iterrows()],
            [2.0, 2.9, 1.6, 1.9, 2.0, 2.0, 1.9, 1.7], saga=(2, 3, 4, 5, 6, 7))
    yasli, genc = g[g["grup"] == "1"].iloc[0], g[g["grup"] == "0"].iloc[0]
    kadin, erkek = g[g["grup"] == "Female"].iloc[0], g[g["grup"] == "Male"].iloc[0]
    ayni_yon = (kadin["yuksek_risk_payi"] > erkek["yuksek_risk_payi"]) == (kadin["gercek_kayip_orani"] > erkek["gercek_kayip_orani"])
    r.p(f"Cinsiyete göre yüksek riskli grupta yer alan müşteri payı kadınlarda {yuzde(kadin['yuksek_risk_payi'])}, erkeklerde "
        f"{yuzde(erkek['yuksek_risk_payi'])} olup gerçek kayıp oranlarındaki fark ({yuzde(kadin['gercek_kayip_orani'])} ve "
        f"{yuzde(erkek['gercek_kayip_orani'])}) ile {'aynı yöndedir' if ayni_yon else 'ters yöndedir'}. "
        "65 yaş ve üzeri müşterilerin "
        f"{yuzde(yasli['yuksek_risk_payi'])} kadarı yüksek riskli grupta yer almaktadır; bu oran diğer müşterilerde "
        f"{yuzde(genc['yuksek_risk_payi'])} düzeyindedir. Fark, bu grubun gerçek kayıp oranının da yüksek olmasıyla "
        f"({yuzde(yasli['gercek_kayip_orani'])}) uyumludur. "
        + (f"Gruplardaki ortalama tahmin gerçek kayıp oranına yakındır (65 yaş ve üzeri: tahmin {yuzde(yasli['ortalama_tahmin'])}, gerçek "
           f"{yuzde(yasli['gercek_kayip_orani'])}; kadın: tahmin {yuzde(kadin['ortalama_tahmin'])}, gerçek {yuzde(kadin['gercek_kayip_orani'])}); "
           "model bu grupların riskini sistematik olarak abartmamakta ya da küçümsememektedir. "
           if max(abs(g["ortalama_tahmin"] - g["gercek_kayip_orani"])) < 0.05 else
           "Bazı gruplarda ortalama tahmin ile gerçek kayıp oranı arasında 5 puandan büyük fark vardır; bu gruplar ayrıca izlenmelidir. ")
        + "Yaş grubu model girdisi olduğu için bu farkın bir kısmı doğrudan yaş bilgisinden, bir kısmı sözleşme türü gibi diğer "
        "girdilerden gelir. Sistem gerçek müşterilerle kullanılacaksa yaşa ya da cinsiyete göre farklı teklif sunmanın hukuki ve etik "
        "yönü ayrıca değerlendirilmeli, bu grupların tekliflerden nasıl etkilendiği izlenmelidir.")


def teklif(r, v):
    k, ic, kisitsiz = v["kesif"]["kayip_orani"], v["internet"], v["kisitsiz"]
    r.baslik("7. Teklif senaryosu (sentetik)")
    r.baslik("7.1. Neden ikinci bir tablo?", 2)
    r.p("Ayrılma olasılığı tek başına kampanya kararı değildir. Karar için her teklifin maliyeti ve müşteri davranışına ek "
        "etkisi gerekir; veri setinde bunların hiçbiri yoktur. Bu nedenle her müşteri ve teklif için maliyet ile etki "
        "varsayımlarını içeren ikinci bir tablo üretilmiş ve dosya adında açıkça işaretlenmiştir "
        f"({A.TEKLIF_TABLOSU.relative_to(A.KOK).as_posix()}).")
    r.baslik("7.2. Ekonomik hesap", 2)
    r.p("Dört seçenek karşılaştırılır: indirim, uygun paket, teknik destek ve teklif sunmama. Her müşteri–teklif çifti için:")
    for satir in ("Ek elde tutma olasılığı = ayrılma olasılığı × teklifin etki oranı",
                  f"Elde tutulan müşterinin katkısı = aylık ücret × katkı marjı × {A.UFUK_AY} ay",
                  "Beklenen ek katkı = ek elde tutma olasılığı × elde tutulan müşterinin katkısı",
                  "Beklenen net fayda = beklenen ek katkı − teklif maliyeti"):
        r.p(satir, italik=True, girinti=True).paragraph_format.space_after = Pt(2)
    r.p("Etki oranı, “teklif sunulmasaydı ayrılacak olan müşterilerin ne kadarı teklifle kalır” sorusunun varsayılan "
        "yanıtıdır. Böylece müşterinin tüm geliri kampanyanın başarısı sayılmaz: zaten kalacak bir müşteriye verilen teklif "
        "yalnızca maliyet yaratır. Teklif maliyetinin, müşteri sonradan ayrılsa bile tamamının harcandığı kabul edilmiştir "
        "(ihtiyatlı varsayım).")
    r.baslik("7.3. Varsayımlar ve dayanakları", 2)
    r.p("Aşağıdaki değerlerin hiçbiri gerçek bir işletmeden, sektör kaynağından veya kampanya deneyinden alınmamıştır. "
        "Tekliflerin yapısı proje yol haritasından, kime uygun oldukları ve kimde daha etkili sayıldıkları keşifsel "
        "analizdeki örüntülerden gelir. Sayıların kendisi yuvarlak ve makul görünen senaryo parametreleridir.")
    etki = lambda sozluk, anahtar: yuzde(sozluk[anahtar], 0)
    r.tablo("Senaryo varsayımları ve dayanakları", ["Varsayım", "Değer", "Dayanak"], [
        ("Para birimi", A.PARA, "Veri setindeki ücretlerin birimi; kur varsayımı eklememek için çevrilmemiştir."),
        ("Değerlendirme dönemi", f"{A.UFUK_AY} ay", "Yuvarlak varsayım: elde tutulan müşterinin bir yıllık katkısı."),
        ("Katkı marjı", f"Aylık ücretin {yuzde(A.KATKI_MARJI, 0)} kadarı",
         "Yuvarlak varsayım. Veri setinde maliyet bilgisi olmadığı için gelir yerine katkı kârı bu oranla temsil edilmiştir."),
        ("Kampanya bütçesi", para(A.BUTCE),
         f"Sınırsız bütçeyle en iyi planın maliyeti {para(kisitsiz['maliyet'])}; bütçe bunun {yuzde(A.BUTCE / kisitsiz['maliyet'], 0)} kadarıdır ve portföyün yıllık "
         f"gelirinin {yuzde(A.BUTCE / v['yillik_gelir'])} kadarıdır. Bu düzeyde bütçe gerçekten kısıtlar."),
        ("Teknik destek kapasitesi", f"{A.DESTEK_KAPASITESI} müşteri",
         f"Sınır olmasa {tr(kisitsiz['teklif_dagilimi'][A.DESTEK])} müşteriye destek önerilir; kapasite bunun {yuzde(A.DESTEK_KAPASITESI / kisitsiz['teklif_dagilimi'][A.DESTEK], 0)} kadarıdır."),
        ("İndirim: maliyet", f"{A.INDIRIM_AY} ay boyunca {yuzde(A.INDIRIM_ORANI, 0)} indirim", "Yuvarlak varsayım."),
        ("İndirim: etki oranı", f"Aylık sözleşmede {etki(A.INDIRIM_ETKI, 'Month-to-month')}, taahhütlü sözleşmede {etki(A.INDIRIM_ETKI, 'One year')}",
         f"Yön keşifsel analizden: aylık sözleşmede kayıp {yuzde(k['Sözleşme türü', 'Month-to-month'])}. Sayılar varsayımdır."),
        ("Uygun paket: uygunluk", f"Aylık ücreti {A.PAKET_ALT_SINIR} {A.PARA} ve üzeri olanlar",
         "Düşük ücretli müşteri için daha ucuz bir paket olmadığı varsayılmıştır; yüksek ücret aralıklarında kayıp oranı daha yüksektir."),
        ("Uygun paket: maliyet", f"{A.PAKET_GECIS_MALIYETI} {A.PARA} geçiş + {A.UFUK_AY} ay {yuzde(A.PAKET_UCRET_DUSUSU, 0)} ücret düşüşü", "Yuvarlak varsayım."),
        ("Uygun paket: etki oranı", f"{A.PAKET_YUKSEK_UCRET} {A.PARA} ve üzerinde {yuzde(A.PAKET_ETKI_YUKSEK, 0)}, altında {yuzde(A.PAKET_ETKI_NORMAL, 0)}",
         "Varsayım. İndirim değerleriyle birlikte, her teklifin en az bir müşteri profilinde öne çıkacağı biçimde ayarlanmıştır."),
        ("Teknik destek: uygunluk", "İnterneti olup teknik destek almayanlar",
         f"Keşifsel analiz: teknik desteği olmayanlarda kayıp {yuzde(k['Teknik destek', 'No'])}, olanlarda {yuzde(k['Teknik destek', 'Yes'])} (ilişki, nedensellik değil)."),
        ("Teknik destek: maliyet", f"Müşteri başına {A.DESTEK_MALIYETI} {A.PARA}", "Yuvarlak varsayım."),
        ("Teknik destek: etki oranı", f"Fiberde {etki(A.DESTEK_ETKI, 'Fiber optic')}, DSL'de {etki(A.DESTEK_ETKI, 'DSL')}",
         f"Yön veriden: fiber kullananlarda kayıp {yuzde(ic['Fiber optic'])}, DSL kullananlarda {yuzde(ic['DSL'])}. Sayılar varsayımdır."),
        ("Kanıt düzeyi", "Tüm teklifler için düşük", "Hiçbir değer gerçek kampanya verisine dayanmamaktadır."),
    ], [3.6, 4.4, 8.0])
    r.p("İndirim ve uygun paket değerleri, her teklifin en az bir müşteri profilinde diğerlerinden iyi çıkacağı biçimde "
        "seçilmiştir; aksi hâlde tek bir teklif her müşteride kazanmakta ve teklif karşılaştırması anlamsızlaşmaktadır. "
        "Bu, sonuçların bu varsayımlara bağlı olduğunu açıkça gösterir. Bütün değerler src/ayarlar.py dosyasında tek yerde "
        "durur ve değiştirilip sistem yeniden çalıştırılabilir.", giris="Açık not:")
    r.baslik("7.4. Tablonun özeti", 2)
    gercek = v["teklifler"][v["teklifler"]["offer_type"] != A.TEKLIF_YOK]
    ozet = gercek.groupby("offer_type").agg(uygun=("customerID", "size"), maliyet=("offer_cost", "mean"),
                                            pozitif=("beklenen_net_fayda", lambda s: (s > 0).sum())).reindex(A.TEKLIFLER)
    r.tablo(f"Teklif türlerine göre uygunluk ve beklenen fayda ({tr(len(v['risk']))} müşteri)",
            ["Teklif", "Uygun olduğu müşteri", "Ortalama maliyet", "Beklenen net faydası pozitif olan müşteri"],
            [(t, tr(s["uygun"]), para(s["maliyet"], 2), tr(s["pozitif"])) for t, s in ozet.iterrows()],
            [3.4, 3.8, 3.4, 5.4], saga=(1, 2, 3))
    r.p("Müşteriye uygun olmayan teklif tabloda yer almaz ve önerilemez (örneğin interneti olmayan müşteriye teknik destek). "
        "Tablodaki alanlar: customerID, offer_type, offer_cost, estimated_incremental_retention, retained_customer_contribution, "
        "beklenen net fayda ve confidence_or_evidence_level.")


def optimizasyon(r, v):
    ozet, kisitsiz, liste = v["ozet"], v["kisitsiz"], v["liste"]
    onerilen = liste[liste["offer_type"] != A.TEKLIF_YOK]
    r.baslik("8. Bütçe optimizasyonu")
    r.baslik("8.1. Model", 2)
    r.p("Müşteri–teklif seçimi ikili karar değişkenli doğrusal bir optimizasyon modeliyle yapılır. Her müşteri–teklif çifti "
        "için bir karar değişkeni tanımlanır (1: teklif sunulur, 0: sunulmaz). Amaç toplam beklenen net faydayı en büyük "
        "yapmaktır. Kısıtlar:")
    r.maddeler(
        "Toplam teklif maliyeti kampanya bütçesini aşamaz.",
        "Bir müşteriye aynı dönemde en fazla bir teklif sunulur.",
        "Teknik destek sunulan müşteri sayısı ekip kapasitesini aşamaz.",
        "Müşteriye uygun olmayan teklif seçilemez.",
    )
    r.p("Beklenen net faydası pozitif olmayan seçenekler modele alınmaz: müşteri riskli diye zarar eden bir teklif "
        "sunulmaz. Model PuLP kütüphanesiyle kurulmuş, HiGHS çözücüsüyle en iyi çözüme kadar çözülmüştür.")
    r.baslik("8.2. Bütçe düzeyinin dayanağı", 2)
    kom = v["butce_kompozisyonu"]
    destek_payi = lambda s: s[A.DESTEK] / s["musteri"] if s["musteri"] else 0.0
    ilk, secilen = kom.iloc[0], kom[kom["butce"] == A.BUTCE].iloc[0]
    r.p(f"Bütçe ve kapasite sınırı olmasaydı en iyi plan {tr(kisitsiz['teklif_sunulan_musteri'])} müşteriye teklif sunar, "
        f"{para(kisitsiz['maliyet'])} harcar ve {para(kisitsiz['beklenen_net_fayda'])} beklenen net fayda üretirdi. "
        f"{para(A.BUTCE)} bütçe bu harcamanın {yuzde(A.BUTCE / kisitsiz['maliyet'], 0)} kadarıdır ve portföyün yıllık gelirinin "
        f"{yuzde(A.BUTCE / v['yillik_gelir'])} kadarına karşılık gelir. Bu düzeyde bütçe gerçekten kısıtlar: beklenen net faydası "
        f"pozitif olan {tr(ozet['butce_disinda_kalan'])} müşteri, bütçe veya teknik destek kapasitesi yetmediği için liste dışında kalır. "
        f"Planın bileşimi bütçeye göre değişir (Tablo {r.tablo_no + 1}): teknik destek payı {para(ilk['butce'])} bütçede "
        f"{yuzde(destek_payi(ilk), 0)}, {para(A.BUTCE)} bütçede {yuzde(destek_payi(secilen), 0)} düzeyindedir.")
    r.tablo("Bütçe düzeyine göre önerilen planın bileşimi (teknik destek kapasitesi sabit)",
            ["Bütçe", "Teklif sunulan müşteri", A.INDIRIM, A.PAKET, A.DESTEK],
            [(para(s["butce"]), tr(s["musteri"]), tr(s[A.INDIRIM]), tr(s[A.PAKET]), tr(s[A.DESTEK])) for _, s in kom.iterrows()]
            + [("Sınırsız bütçe ve kapasite", tr(kisitsiz["teklif_sunulan_musteri"]))
               + tuple(tr(kisitsiz["teklif_dagilimi"][t]) for t in A.TEKLIFLER)],
            [4.6, 3.2, 2.7, 2.7, 2.8], saga=(1, 2, 3, 4), ad="butce_kompozisyon")
    r.baslik("8.3. Önerilen plan", 2)
    r.p(f"Sistem {para(ozet['butce'])} bütçenin {para(ozet['harcama'])} tutarındaki bölümünü ({yuzde(ozet['kullanim_orani'])}) "
        f"{tr(ozet['teklif_sunulan_musteri'])} müşteriye ayırmayı önermektedir. Planın beklenen ek katkısı "
        f"{para(ozet['beklenen_ek_katki'])}, beklenen net faydası {para(ozet['beklenen_net_fayda'])} düzeyindedir.")
    dagilim = onerilen.groupby("offer_type").agg(n=("customerID", "size"), maliyet=("offer_cost", "sum"),
                                                net=("beklenen_net_fayda", "sum")).reindex(A.TEKLIFLER, fill_value=0)
    r.tablo("Önerilen planda teklif türlerinin dağılımı", ["Teklif", "Müşteri", "Toplam maliyet", "Beklenen net fayda"],
            [(t, tr(s["n"]), para(s["maliyet"]), para(s["net"])) for t, s in dagilim.iterrows()]
            + [("Toplam", tr(dagilim["n"].sum()), para(dagilim["maliyet"].sum()), para(dagilim["net"].sum()))],
            [4.0, 3.0, 4.0, 4.5], saga=(1, 2, 3))
    r.maddeler(
        f"Önerilen {tr(ozet['varsayima_duyarli_teklif'])} teklif, gerçek etki varsayılanın yarısı kadar olursa maliyetini "
        "karşılamamaktadır; bu teklifler listede “Varsayıma duyarlı” olarak işaretlenmiştir.",
        f"Beklenen net faydası pozitif olan {tr(ozet['butce_disinda_kalan'])} müşteri, bütçe veya teknik destek kapasitesi "
        "yetmediği için liste dışında kalmıştır.",
        f"Kalan {tr(len(liste) - len(onerilen) - ozet['butce_disinda_kalan'])} müşteride hiçbir teklifin beklenen net faydası pozitif değildir; "
        "bu müşterilere teklif sunulmaması önerilir.",
    )
    r.tablo("Önerilen kampanya listesinden ilk on müşteri (beklenen net faydaya göre sıralı)",
            ["Müşteri kodu", "Ayrılma olasılığı", "Risk işaretleri", "Önerilen teklif", "Maliyet", "Beklenen net fayda"],
            [(s["customerID"], yuzde(s["churn_olasiligi"]), s["risk_isaretleri"], s["offer_type"], para(s["offer_cost"]),
              para(s["beklenen_net_fayda"])) for _, s in onerilen.head(10).iterrows()],
            [2.3, 1.7, 6.6, 2.2, 1.3, 1.9], saga=(1, 4, 5))
    r.p(f"Listenin tamamı {A.KAMPANYA_LISTESI.relative_to(A.KOK).as_posix()} dosyasındadır.")


def senaryolar(r, v):
    s, d, egri, ozet = v["senaryolar"], v["duyarlilik"], v["egri"], v["ozet"]
    a, b, c = (s.loc[ad] for ad in (A.SENARYO_A, A.SENARYO_B, A.SENARYO_C))
    r.baslik("9. Senaryo karşılaştırması ve duyarlılık")
    r.baslik("9.1. Senaryolar", 2)
    r.p(f"Üç karar yaklaşımı aynı {tr(len(v['risk']))} müşteri, aynı {para(A.BUTCE)} bütçe ve aynı varsayımlarla karşılaştırılmıştır:")
    r.maddeler(
        "A – Mevcut yöntem: model yoktur. Aylık sözleşmeli ve ilk 12 ayındaki müşterilere, en yüksek ücretliden başlayarak bütçe bitene kadar aynı indirim sunulur.",
        "B – Yalnızca risk modeli: yüksek riskli müşterilere, en riskliden başlayarak bütçe bitene kadar aynı indirim sunulur.",
        "C – Karar destek sistemi: müşteri ve teklif bazında beklenen net fayda, bütçe ve kapasite kısıtlarıyla optimize edilir.",
    )
    r.baslik("9.2. Sonuçlar", 2)
    r.tablo("Senaryo karşılaştırması",
            ["Senaryo", "Teklif sunulan müşteri", "Maliyet", "Beklenen net fayda", "Etiket bazlı net fayda", "Gerçekte ayrılan hedef"],
            [(ad, tr(x["teklif_sunulan_musteri"]), para(x["maliyet"]), para(x["beklenen_net_fayda"]), para(x["etiket_bazli_net_fayda"]),
              f"{tr(x['gercekte_ayrilan_hedef'])} ({yuzde(x['isabet_orani'], 0)})") for ad, x in s.iterrows()],
            [4.3, 2.2, 2.0, 2.5, 2.5, 2.5], saga=(1, 2, 3, 4, 5))
    r.sekil(A.GRAFIK / "08_senaryo_karsilastirma.png", "Senaryolara göre beklenen net fayda")
    r.p(f"Aynı bütçeyle risk modeli tek başına mevcut yönteme göre beklenen net faydayı {para(a['beklenen_net_fayda'])} düzeyinden "
        f"{para(b['beklenen_net_fayda'])} düzeyine {'çıkarmaktadır' if b['beklenen_net_fayda'] > a['beklenen_net_fayda'] else 'indirmektedir'}. "
        f"Hedeflenen müşterilerden gerçekte ayrılanların payı A'da {yuzde(a['isabet_orani'], 0)}, B'de {yuzde(b['isabet_orani'], 0)}, "
        f"C'de {yuzde(c['isabet_orani'], 0)} düzeyindedir. Karar destek sistemi aynı bütçeyle {para(c['beklenen_net_fayda'])} "
        "beklenen net fayda üretmektedir.")
    ayni_sira = s["beklenen_net_fayda"].rank().equals(s["etiket_bazli_net_fayda"].rank())
    if ayni_sira:
        sira_notu = "Senaryoların sıralaması bu ölçümde de değişmemektedir."
    else:
        sira_notu = (f"Bu ölçümde A {para(a['etiket_bazli_net_fayda'])}, B {para(b['etiket_bazli_net_fayda'])}, "
                     f"C {para(c['etiket_bazli_net_fayda'])} düzeyindedir.")
        if s["etiket_bazli_net_fayda"].idxmax() == A.SENARYO_C:
            sira_notu += " C her iki ölçümde de en yüksektir."
        if b["etiket_bazli_net_fayda"] <= a["etiket_bazli_net_fayda"]:
            sira_notu += (f" Buna karşılık bu ölçümde B, A'nın {para(a['etiket_bazli_net_fayda'] - b['etiket_bazli_net_fayda'])} gerisinde kalmaktadır: "
                          "risk modelinin tek başına sağladığı üstünlük, modelin olasılıklarına dayanan beklenen değerde görünmekte ama gerçek "
                          "etiketlerle doğrulanmamaktadır. Fayda, riskle birlikte teklifin maliyetini ve etkisini de hesaba katan C'de ortaya çıkmaktadır.")
    r.p("Beklenen net fayda modelin kendi olasılıklarıyla hesaplanır. Modelin olasılıkları kararı yanlış yönlendiriyorsa bunu "
        "görebilmek için ikinci bir ölçüm yapılmıştır: aynı etki varsayımı, test müşterilerinin gerçekte ayrılıp ayrılmadığı "
        f"bilgisiyle birleştirilmiştir. {sira_notu} Her iki ölçümde de kampanya etkisi varsayımdır.",
        giris="Etiket bazlı net fayda:")
    destek_payi = ozet["teklif_dagilimi"][A.DESTEK] / ozet["teklif_sunulan_musteri"]
    r.p(f"C senaryosunun üstünlüğünün önemli bir bölümü, A ve B yalnızca indirim sunabilirken sistemin daha ucuz bir teklifi "
        f"seçebilmesinden gelir: önerilen tekliflerin {yuzde(destek_payi, 0)} kadarı teknik destektir ve bu teklifin maliyeti ile "
        "etkisi varsayımdır. Sentetik bir senaryoda sistemin kendi varsayımları altında daha iyi sonuç vermesi, gerçek dünyada "
        "aynı iyileşmenin elde edileceğini kanıtlamaz.", giris="Yorum sınırı:")
    r.baslik("9.3. Bütçe düzeyine duyarlılık", 2)
    r.sekil(A.GRAFIK / "09_butce_egrisi.png", "Bütçe arttıkça senaryoların beklenen net faydası", 14)
    c_egri, a_egri = egri[A.SENARYO_C], egri[A.SENARYO_A]
    a_dusuyor = a_egri.iloc[-1] < a_egri.max() * 0.99
    r.p(f"Bütçe arttıkça her ek harcamanın getirisi azalmaktadır: karar destek sisteminde beklenen net fayda {para(egri.index[0])} bütçede "
        f"{para(c_egri.iloc[0])}, {para(A.BUTCE)} bütçede {para(c_egri.loc[A.BUTCE])}, {para(egri.index[-1])} bütçede {para(c_egri.iloc[-1])} "
        f"düzeyindedir; harcanan her dolar başına net fayda {tr(c_egri.iloc[0] / egri.index[0], 2)} dolardan {tr(c_egri.iloc[-1] / egri.index[-1], 2)} dolara düşer. "
        f"Mevcut yöntemde beklenen net fayda {para(a_egri.idxmax())} bütçede en yüksek düzeyine ({para(a_egri.max())}) ulaşır ve sonrasında "
        f"{'azalmaktadır' if a_dusuyor else 'artmamaktadır'}: kural, beklenen faydası maliyetini karşılamayan müşterilere de indirim vermeyi sürdürür. "
        "Karar destek sistemi ise net faydası pozitif olmayan teklifi sunmaz.")
    r.baslik("9.4. Kampanya etkisine duyarlılık", 2)
    r.p("Tekliflerin gerçek etkisi bilinmediği için planlar sabit tutulmuş ve gerçek etkinin varsayılandan düşük ya da yüksek "
        "çıkması durumunda net faydanın ne olacağı hesaplanmıştır.")
    r.tablo(f"Gerçekleşen etki varsayılandan saparsa net fayda ({A.PARA})", ["Senaryo"] + [f"×{tr(x, 2)}" for x in ETKI_DUZEYLERI],
            [(ad,) + tuple(tr(d.loc[x, ad]) for x in ETKI_DUZEYLERI) for ad in d.columns],
            [4.5, 2.3, 2.3, 2.3, 2.3, 2.3], saga=(1, 2, 3, 4, 5))
    r.sekil(A.GRAFIK / "10_duyarlilik.png", "Kampanya etkisi varsayımdan saparsa senaryoların net faydası", 14)
    durum = lambda x: "zarar eder" if x < 0 else "maliyetini karşılar"
    r.p(f"Gerçek etki varsayılanın yarısı kadar olursa A senaryosu {para(d.loc[0.5, A.SENARYO_A])} ile {durum(d.loc[0.5, A.SENARYO_A])}, "
        f"B senaryosu {para(d.loc[0.5, A.SENARYO_B])} ile {durum(d.loc[0.5, A.SENARYO_B])}, "
        f"önerilen plan ise {para(d.loc[0.5, A.SENARYO_C])} ile {durum(d.loc[0.5, A.SENARYO_C])}. "
        f"Önerilen plan, gerçek etki varsayılan etkinin {yuzde(c['basabas_etki_carpani'], 0)} düzeyine inene kadar maliyetini karşılar; "
        f"bu eşik A için {yuzde(a['basabas_etki_carpani'], 0)}, B için {yuzde(b['basabas_etki_carpani'], 0)} düzeyindedir.")


def ekran(r):
    r.baslik("10. Yönetici karar ekranı")
    r.p("Yönetici ekranı Streamlit ile hazırlanmıştır ve modelin teknik ayrıntılarından çok karar için gereken bilgiyi öne "
        "çıkarır. Yönetici sol paneldeki bütçe ve teknik destek kapasitesi değerlerini değiştirdiğinde optimizasyon yeniden "
        "çalışır ve ekrandaki bütün sayılar güncellenir.")
    r.tablo("Karar ekranının bölümleri", ["Bölüm", "Gösterilen bilgi"], [
        ("Öneri", "Önerilen müşteri sayısı ve harcama tek cümleyle; yüksek riskli müşteri sayısı, bütçe kullanımı, beklenen net fayda ve mevcut yönteme göre fark."),
        ("Bütçe nereye gidiyor?", "Teklif türüne göre harcama, müşteri sayısı ve her teklifin kime sunulduğu."),
        ("Bugünkü yöntemle karşılaştırma", "Mevcut yöntem, yalnızca risk sıralaması ve önerilen planın aynı bütçeyle beklenen net faydası."),
        ("Ne kadar emin olabiliriz?", "Teklifler beklenenden az ya da çok işe yararsa planın net faydası; planın zarara geçtiği eşik ve uyarılar."),
        ("Kim risk altında?", "Risk gruplarının dağılımı ve yüksek riskli müşterilerde en sık görülen risk işaretleri."),
        ("Kararınız", "Onayla, değiştir ya da reddet; karar gerekçesiyle ve o andaki varsayımlarla birlikte kayıt dosyasına eklenir."),
        ("Müşteri listesi sekmesi", "Müşteri kodu, ayrılma olasılığı, risk işaretleri, önerilen teklif, maliyet, beklenen net fayda ve belirsizlik; tek bir müşteri için dört seçeneğin karşılaştırması."),
        ("Yöntem ve model sekmesi", "Model ölçütleri (çapraz doğrulama ortalaması ± standart sapma), kalibrasyon, SHAP açıklamaları, grup bazlı kontrol ve senaryoların teknik karşılaştırması."),
    ], [4.6, 11.4])
    if A.EKRAN_GORUNTUSU.exists():
        r.sekil(A.EKRAN_GORUNTUSU, "Yönetici karar ekranının özet sekmesi", 16)
    r.p("Sistem nihai kararı otomatik vermez. Yöneticinin kararı ve gerekçesi "
        f"{A.KARAR_KAYDI.relative_to(A.KOK).as_posix()} dosyasına, kararın verildiği bütçe ve varsayımlarla birlikte kaydedilir.")


def sinirliliklar(r):
    r.baslik("11. Sınırlılıklar, etik ve riskler")
    r.maddeler(
        "Kampanya etkisi doğrulanmamıştır. Teklif maliyetleri, etki oranları, katkı marjı, bütçe ve kapasite sentetik varsayımdır; net fayda sayıları gerçek bir kampanyanın sonucunu göstermez.",
        "Ayrılma riski, müdahalenin etkisi değildir. Yüksek riskli müşteriye teklif sunmak müşteri kaybını kendiliğinden önlemez; etkinin ölçülmesi için kontrol gruplu bir deney gerekir.",
        "Açık eğitim verisi gerçek bir işletmenin güncel müşterilerini temsil etmeyebilir; model başka bir işletmeye doğrudan taşınmamalıdır.",
        "Veri kesitseldir. Tahminler belirli bir gelecek dönem için değil, geçmişte ayrılan müşterilere benzerlik içindir; “gelecek ay ayrılacak müşteri” olarak sunulmamalıdır.",
        "Kampanya kararı test verisindeki müşteriler üzerinde gösterilmiştir. Gerçek kullanımda model, sonucu henüz bilinmeyen güncel müşterilere uygulanır.",
        "Cinsiyet ve yaş grubu model girdisidir. Gruplar arasında belirgin bir dengesizlik görülmemiştir (Bölüm 6.7), ancak gerçek kullanımda bu bilgilere dayanarak farklı teklif sunmanın hukuki ve etik yönü ayrıca değerlendirilmeli, sonuçlar gruplar bazında izlenmelidir.",
        "Gerçek kişisel veriler kullanılacaksa hukuki dayanak, erişim yetkisi, veri minimizasyonu, saklama süresi ve güvenlik ayrıca ele alınmalıdır.",
        "Sistem öneri sunar; karar yöneticinindir. Öneri gerekçesi, maliyeti ve belirsizliğiyle birlikte gösterilir.",
        "Tekrar eden kayıtlar kimlik ve etiket dışındaki bütün özelliklerin aynı olmasına göre tanımlanmıştır. Çıkarılan kayıtların hepsi "
        "abonelik süresi 1 ay olan müşterilerdir ve çoğu büyük olasılıkla aynı müşterinin kopyası değil, aynı paketi alan farklı müşterilerdir "
        "(Bölüm 4.2). Etiketi çelişen gruplarda tutulan kaydın etiketi dosya sırasına bağlıdır.",
        "Bütün sonuçlar tek bir rastgele eğitim–test ayırmasına (tohum 42) dayanır. Test kısmı bir kez kullanıldığı için çapraz doğrulamaya "
        "sokulmamış, belirsizlik bootstrap güven aralıklarıyla gösterilmiştir; farklı bir ayırmada küçük farklar beklenmelidir.",
        "SHAP açıklamaları modelin davranışını açıklar, nedensellik kanıtı değildir. Birbirinden türetilebilen değişkenlerde (abonelik "
        "süresi, aylık ücret, toplam ücret) tek değişken düzeyindeki SHAP değerleri yanıltıcı olabilir; bu değişkenler grup olarak yorumlanmalıdır (Bölüm 6.8).",
    )


def sonuc(r, v):
    c = v["senaryolar"].loc[A.SENARYO_C]
    r.baslik("12. Sonuç ve sonraki adımlar")
    r.p("Proje, müşteri kaybı tahmininin tek başına bir karar olmadığını; tahminin teklif maliyeti, beklenen ek etki ve bütçe "
        "kısıtıyla birleştirildiğinde karar desteğine dönüştüğünü göstermektedir. Kurulan sistem görülmemiş veride makul bir "
        "sıralama becerisine sahiptir ve aynı bütçeyle, kendi varsayımları altında, mevcut yönteme göre daha yüksek beklenen "
        f"net fayda ({para(c['beklenen_net_fayda'])}) üretmektedir. Bu sonucun gerçek dünyadaki karşılığı kampanya etkisi ölçülmeden bilinemez.")
    r.p("Önerilen sonraki adımlar:")
    r.maddeler(
        "Planı sınırlı bir pilotla ve rastgele seçilmiş bir kontrol grubuyla uygulamak; teklif türü bazında gerçek etkiyi ölçmek.",
        "Ölçülen etki ve gerçek maliyet verileriyle senaryo tablosunu güncellemek; sentetik değerleri doğrulanmış değerlerle değiştirmek.",
        "Yeterli kampanya verisi oluştuğunda müşteri bazında ek etkiyi doğrudan tahmin eden (uplift) bir model denemek.",
        "Zaman bilgisi içeren veriyle modeli ileriye dönük doğrulamak ve olasılıkların kalibrasyonunu dönemsel olarak izlemek.",
        "Yönetici kararlarının kaydını, önerilerin ne sıklıkla ve hangi gerekçeyle değiştirildiğini görmek için incelemek.",
    )


def kabul_kriterleri(r):
    r.baslik("13. Kabul kriterlerinin karşılanması")
    r.tablo("Yol haritasındaki kabul kriterleri ve karşılandıkları yer", ["Kriter", "Karşılandığı yer"], [
        ("Veri kaynağı ve veri setinin sınırlılıkları açıklanmış olmalı.", "Bölüm 3 ve 11."),
        ("Veri temizleme adımları tekrarlanabilir biçimde kaydedilmiş olmalı.", "Bölüm 4; src/veri_hazirlama.py; outputs/veri_kalite_raporu.md."),
        ("Model, görülmemiş test verisi üzerinde değerlendirilmiş olmalı.", "Bölüm 6.2 ve 6.3."),
        ("Kesinlik, duyarlılık ve en az bir sıralama ölçütü raporlanmalı; mümkünse kalibrasyon incelenmeli.", f"Bölüm 6.3 (Tablo {r.no['olcut']}) ve 6.4."),
        ("Müşteri kaybı tahmini ile kampanya etkisi birbirinden ayrılmış olmalı.", "Bölüm 7.2: etki oranı ayrı bir varsayımdır."),
        ("Teklif maliyeti ve bütçe kısıtı hesaplamaya dahil edilmeli.", "Bölüm 7 ve 8."),
        ("Yönetici ekranında önerinin gerekçesi, maliyeti ve belirsizliği görünmeli.", "Bölüm 10."),
        ("Mevcut yaklaşım ile önerilen yaklaşım aynı varsayımlar altında karşılaştırılmalı.", "Bölüm 9."),
        ("Sentetik veriler ve varsayımlar gerçek işletme sonuçlarından açıkça ayrılmalı.", "Bölüm 7.3; sentetik tablonun dosya adı; ekrandaki uyarılar."),
        ("Kaynakça, kullanılan yazılım ve veri sürümleri rapora eklenmeli.", f"Bölüm 14 ve Tablo {r.no['kaynak']}."),
    ], [9.0, 7.0])


def kaynakca(r):
    r.baslik("14. Kaynakça, yazılım ve veri sürümleri")
    r.baslik("14.1. Kaynakça", 2)
    r.maddeler(
        "Kaggle. Telco Customer Churn veri seti (blastchar). https://www.kaggle.com/datasets/blastchar/telco-customer-churn",
        "IBM. Telco Customer Churn örnek deposu. https://github.com/IBM/telco-customer-churn-on-icp4d",
        "IBM Cognos Analytics. Telco Customer Churn sample. https://www.ibm.com/docs/en/cognos-analytics/12.1.x?topic=samples-telco-customer-churn",
        "scikit-learn. Model evaluation. https://scikit-learn.org/stable/modules/model_evaluation.html",
        "scikit-learn. Metrics API reference. https://scikit-learn.org/stable/api/sklearn.metrics.html",
        "PuLP. Python doğrusal programlama kütüphanesi. https://coin-or.github.io/pulp/",
        "HiGHS. Doğrusal ve karma tam sayılı optimizasyon çözücüsü. https://highs.dev",
        "Streamlit. https://streamlit.io",
        "SHAP. Python kütüphanesi. https://github.com/shap/shap",
    )
    r.baslik("14.1.1. Makale kaynakları", 2)
    r.p("Aşağıdaki makalelerin künyeleri DOI kayıt sistemi Crossref üzerinden doğrulanmıştır. Metinde (Yazar, yıl) biçiminde anılırlar.")
    r.maddeler(
        "Aas, K., Jullum, M. ve Løland, A. (2021). Explaining individual predictions when features are dependent: More accurate approximations to Shapley values. Artificial Intelligence, 298, 103502. https://doi.org/10.1016/j.artint.2021.103502",
        "Arlot, S. ve Celisse, A. (2010). A survey of cross-validation procedures for model selection. Statistics Surveys, 4, 40–79. https://doi.org/10.1214/09-SS054",
        "Barz, B. ve Denzler, J. (2020). Do we train on test data? Purging CIFAR of near-duplicates. Journal of Imaging, 6(6), 41. https://doi.org/10.3390/jimaging6060041",
        "Chang, V., Hall, K., Xu, Q. A., Amao, F. O. vd. (2024). Prediction of customer churn behavior in the telecommunication industry using machine learning models. Algorithms, 17(6), 231. https://doi.org/10.3390/a17060231",
        "El Attar, El-Hajj (2026). Explainable AI-driven customer churn prediction: a multi-model ensemble approach with SHAP-based feature analysis. Frontiers in Artificial Intelligence, 9, 1748799. https://doi.org/10.3389/frai.2026.1748799",
        "Kapoor, S. ve Narayanan, A. (2023). Leakage and the reproducibility crisis in machine-learning-based science. Patterns, 4(9), 100804. https://doi.org/10.1016/j.patter.2023.100804",
        "Kumar, I. E., Venkatasubramanian, S., Scheidegger, C. ve Friedler, S. (2020). Problems with Shapley-value-based explanations as feature importance measures. Proceedings of the International Conference on Machine Learning (ICML 2020), PMLR 119. https://arxiv.org/abs/2002.11097",
        "Lundberg, S. M. ve Lee, S.-I. (2017). A unified approach to interpreting model predictions. Advances in Neural Information Processing Systems 30 (NIPS 2017). https://arxiv.org/abs/1705.07874",
        "Lundberg, S. M., Erion, G., Chen, H., DeGrave, A. vd. (2020). From local explanations to global understanding with explainable AI for trees. Nature Machine Intelligence, 2(1), 56–67. https://doi.org/10.1038/s42256-019-0138-9",
        "Poudel, Pokharel ve Timilsina (2024). Explaining customer churn prediction in telecom industry using tabular machine learning models. Machine Learning with Applications, 17, 100567. https://doi.org/10.1016/j.mlwa.2024.100567",
        "Varma, S. ve Simon, R. (2006). Bias in error estimation when using cross-validation for model selection. BMC Bioinformatics, 7, 91. https://doi.org/10.1186/1471-2105-7-91",
    )
    r.baslik("14.2. Yazılım sürümleri", 2)
    r.tablo("Kullanılan yazılım sürümleri", ["Yazılım", "Sürüm"],
            [("Python", sys.version.split()[0])] + [(ad, version(ad)) for ad in YAZILIMLAR], [6.0, 4.0])
    r.p(f"Veri sürümü, kaynak kaydındaki (Tablo {r.no['kaynak']}) dosya tarihi ve SHA-256 özetiyle kaydedilmiştir.")
    r.baslik("14.3. Çalıştırma ve dosya düzeni", 2)
    r.p("Bütün çıktılar ham veriden aşağıdaki komutlarla yeniden üretilebilir:")
    for komut in ("py -m pip install -r requirements.txt", "py src/hepsini_calistir.py", "py -m streamlit run app.py", "py src/rapor_olustur.py"):
        r.p(komut, girinti=True).paragraph_format.space_after = Pt(2)
    r.tablo("Proje klasör yapısı", ["Klasör / dosya", "İçerik"], [
        ("data/raw/", "Kaynaktan indirilen, değiştirilmemiş ham CSV."),
        ("data/processed/", "Temizlenmiş veri ve sentetik teklif senaryosu tablosu."),
        ("src/", "Veri hazırlama (tekrar eden kayıtların çıkarılması dahil), keşifsel analiz, model (çapraz doğrulama), tekrar kayıt etkisi, SHAP analizi, teklif senaryosu, optimizasyon, senaryo karşılaştırması ve rapor betikleri; varsayımlar ayarlar.py dosyasında."),
        ("app.py", "Yönetici karar ekranı."),
        ("outputs/", "Grafikler, tablolar, model ölçütleri, müşteri risk tablosu, kampanya listesi, yönetici özeti ve karar kaydı."),
        ("docs/", "Veri sözlüğü ve bu rapor."),
    ], [4.2, 11.8])


def tekrar_kayitlari(r, v):
    t, k = v["tekrar"], v["kontrol"]
    ozellik_sayisi, girdi_sayisi = len(t["anahtar_sutunlari"]), len(A.SAYISAL + A.KATEGORIK)
    r.baslik("4.2. Tekrar eden kayıtların kontrolü ve çıkarılması", 2)
    r.p("Aynı kaydın hem eğitim hem de test (ya da çapraz doğrulamadaki doğrulama) kısmında bulunması sonuçları gerçekte olduğundan "
        "iyi gösterebilir. Bu nedenle model kurulmadan önce tekrar eden kayıtlar araştırılmıştır. Veri setinde sözleşme yenilemeyi ya da "
        "bir müşterinin birden fazla satırla temsil edilmesini gösteren bir alan yoktur; bu yüzden “yenilenen / tekrar eden kayıt”, "
        "aynı özellikleri taşıyan kayıtlar olarak ele alınmış ve kontrol dört düzeyde yapılmıştır.")
    r.tablo("Tekrar kontrolünün düzeyleri", ["Düzey", "Karşılaştırılan sütunlar", "Sonuç"], [
        ("1", "customerID", f"{tr(t['ayni_kimlik'])} tekrar"),
        ("2", "Bütün sütunlar (kimlik dahil)", f"{tr(k['tekrar_eden_satir'])} tekrar"),
        ("3", f"Kimlik ve etiket dışındaki {ozellik_sayisi} özellik (asıl kural)",
         f"{tr(t['tekrar_grubu'])} grup, {tr(t['gruplardaki_toplam_kayit'])} kayıt; {tr(t['cikarilan_kayit'])} fazla kayıt çıkarıldı"),
        ("4", f"Modelin kullandığı {girdi_sayisi} sütun (sıkı kural)",
         f"{tr(t['girdi_ayni_fazla_kayit'])} fazla kayıt; yalnızca duyarlılık denemesinde çıkarıldı"),
    ], [1.6, 7.4, 7.0], ad="tekrar_duzey")
    r.p("Çıkarma işlemi src/veri_hazirlama.py betiğiyle, ham dosyaya dokunmadan ve şu kurallarla yapılmıştır:", giris="Nasıl çıkarıldı?")
    r.maddeler(
        f"Karşılaştırma anahtarı kimlik ve etiket dışındaki {ozellik_sayisi} sütundur (demografik alanlar dahil; MonthlyCharges ve "
        "TotalCharges tam sayısal eşitlikle karşılaştırılır).",
        "Anahtarı aynı olan kayıtlar bir grup sayılır. Her grupta dosyadaki ilk kayıt tutulur, diğerleri çıkarılır; seçimde etikete bakılmaz.",
        f"Çıkarılan {tr(t['cikarilan_kayit'])} kaydın {tr(t['cikarilan_etiketi_ayni'])} tanesinin etiketi tutulan kayıtla aynı, "
        f"{tr(t['cikarilan_etiketi_farkli'])} tanesinin farklıdır; {tr(t['etiketi_celisen_grup'])} grupta etiketler çelişir. "
        "Bu gruplarda tutulan kaydın etiketi dosya sırasına bağlıdır.",
        "Çıkarma, eğitim–test ayrımından ve çapraz doğrulama katlarının oluşturulmasından önce yapılır; böylece aynı özellikli iki kayıt "
        "hiçbir zaman farklı kısımlara düşmez. Kod, çıkarmadan sonra aynı özellikli kayıt kalmadığını otomatik olarak doğrular.",
        f"Çıkarılan kayıtların listesi ({tr(t['cikarilan_kayit'])} satır, hangi kaydın tutulduğuyla birlikte) "
        f"{A.TEKRAR_KAYITLAR.relative_to(A.KOK).as_posix()} dosyasındadır; modelleme verisi {A.MODELLEME_VERISI.relative_to(A.KOK).as_posix()} "
        "dosyasına yazılır. Temizlenmiş tam veri (telco_temiz.csv) değiştirilmez.",
    )
    sure = t["cikarilan_tenure"]
    if sure == [1]:
        r.p("Çıkarılan kayıtların hepsi abonelik süresi 1 ay olan müşterilerdir. Bu müşterilerde toplam ücret aylık ücrete eşittir; "
            "sayısal alanlar ayırt edici bilgi taşımaz ve aynı hizmet paketini alan iki yeni müşteri birebir aynı görünür. "
            f"Etiketi çelişen {tr(t['etiketi_celisen_grup'])} grubun varlığı (aynı müşterinin kopyası olsalardı etiketleri aynı olurdu) "
            "bu kayıtların çoğunun aynı müşterinin kopyası değil, aynı paketi alan farklı yeni müşteriler olduğunu düşündürür. "
            "Buna rağmen çıkarma ihtiyatlı bir seçimdir: olası bir kopyanın doğrulamaya sızmasını engeller ve veri kaybı küçüktür "
            f"({yuzde(t['cikarilan_kayit'] / t['kayit_oncesi'])}).", giris="Bu kayıtlar neler?")
    else:
        r.p(f"Çıkarılan kayıtların abonelik süreleri (ay): {', '.join(str(x) for x in sure)}.", giris="Bu kayıtlar neler?")

    etki = pd.DataFrame(t["etki"])
    r.tablo("Üç veri düzeninde doğrulama sızıntısı ve test başarısı (aynı ayırma oranı, tohum ve yöntem)",
            ["Veri düzeni", "Kayıt", f"Test ikizi ({ozellik_sayisi} özellik)", f"CV sızıntısı ({ozellik_sayisi} özellik)",
             f"Test ikizi ({girdi_sayisi} girdi)", "ROC-AUC lojistik", "ROC-AUC ağaç"],
            [(s["duzen"], tr(s["kayit"]), f"{s['test_ikizi_19']} (çelişen {s['test_ikizi_celisen_19']})", tr(s["cv_sizintisi_19"]),
              f"{s['test_ikizi_15']} (çelişen {s['test_ikizi_celisen_15']})", tr(s["lojistik_roc_auc"], 3), tr(s["agac_roc_auc"], 3))
             for _, s in etki.iterrows()],
            [4.0, 1.4, 2.2, 2.0, 2.2, 2.1, 2.1], saga=(1, 2, 3, 4, 5, 6), ad="tekrar_etki")
    t0, t1, t2 = etki.iloc[0], etki.iloc[1], etki.iloc[2]
    fark = t1["lojistik_roc_auc"] - t0["lojistik_roc_auc"]
    guven = v["metrik"]["test_guven_araligi"]["Lojistik regresyon"]["roc_auc"]
    r.p(f"Çıkarmadan önce {t0['test_ikizi_19']} test kaydının (etiketi çelişen {t0['test_ikizi_celisen_19']}) eğitim kısmında özellikleri "
        f"birebir aynı bir ikizi, çapraz doğrulamada ise {t0['cv_sizintisi_19']} doğrulama kaydının ikizi aynı katın eğitim kısmında bulunuyordu. "
        "Asıl kuralla çıkarma sonrasında bu sayıların ikisi de sıfırdır. Başarıya etkisi "
        f"{'çok küçüktür' if abs(fark) < 0.005 else 'gözle görülürdür'}: lojistik regresyonun test ROC-AUC değeri "
        f"{tr(t0['lojistik_roc_auc'], 3)} düzeyinden {tr(t1['lojistik_roc_auc'], 3)} düzeyine değişmiştir (fark {tr(abs(fark), 3)}); bu fark, "
        f"test kısmının %95 bootstrap aralığının genişliğinin ({tr(guven[1] - guven[0], 3)}) yanında "
        f"{'çok küçüktür' if abs(fark) < (guven[1] - guven[0]) / 4 else 'küçük değildir'}. Çıkarmanın asıl kazancı sonucu değiştirmesi değil, "
        "doğrulamanın sızıntıdan arınmış olmasının yapısal olarak güvence altına alınmasıdır. Veri sızıntısı, makine öğrenmesi "
        "çalışmalarında bildirilen başarıyı şişiren yaygın bir hata kaynağıdır (Kapoor ve Narayanan, 2023); test kümesindeki yakın "
        "kopyaların başarıyı olduğundan yüksek gösterebildiği görüntü sınıflandırmada da gösterilmiştir (Barz ve Denzler, 2020).")
    r.p(f"Modelin kullandığı {girdi_sayisi} sütunu aynı olan kayıtların da çıkarıldığı sıkı kuralda lojistik regresyonun test ROC-AUC değeri "
        f"{tr(t2['lojistik_roc_auc'], 3)} bulunmuştur; sonuç çıkarma kuralına "
        f"{'duyarlı değildir' if abs(t2['lojistik_roc_auc'] - t1['lojistik_roc_auc']) < 0.005 else 'bir miktar duyarlıdır'}. "
        "Sıkı kural asıl kural olarak seçilmemiştir: aynı hizmet paketini ve aynı ücreti alan farklı müşteriler gerçek hayatta olağandır; "
        "bu kayıtlar tekrar değil gerçek gözlemlerdir ve çıkarılmaları veriyi gereksiz yere daraltır.")
    sozlesme, etiket = {"Month-to-month": "Aylık", "One year": "1 yıllık", "Two year": "2 yıllık"}, {"Yes": "Ayrıldı", "No": "Kaldı"}
    r.tablo(f"Çıkarılan kayıtlardan ilk sekizi (tamamı {A.TEKRAR_KAYITLAR.name} dosyasında)",
            ["Çıkarılan kimlik", "Tutulan kimlik", "Abonelik süresi", f"Aylık ücret ({A.PARA})", "Sözleşme", "Etiket (çıkarılan / tutulan)"],
            [(s["customerID"], s["korunan_kimlik"], f"{s['tenure']} ay", tr(s["MonthlyCharges"], 2), sozlesme.get(s["Contract"], s["Contract"]),
              f"{etiket[s['Churn']]} / {etiket[s['korunan_etiket']]}") for _, s in v["tekrar_kayitlar"].head(8).iterrows()],
            [2.7, 2.7, 2.1, 2.3, 2.2, 4.0], saga=(3,))


def veri_akisi(r, v):
    t, m, k, s = v["tekrar"], v["metrik"], v["kontrol"], v["shap"]
    r.baslik("4.3. Veri işleme akışı", 2)
    r.p("Ham veriden raporlanan sonuçlara kadar verinin geçtiği adımlar aşağıdadır. Her adım src/ klasöründeki bir betikle yapılır; "
        "src/hepsini_calistir.py bütün akışı baştan çalıştırır.")
    r.tablo("Verinin işlenme adımları", ["Adım", "İşlem", "Sonuç"], [
        ("1. Ham veri", "Kaggle'dan indirilen CSV hiçbir değişiklik yapılmadan saklanır.", f"{tr(k['satir'])} kayıt × {k['sutun']} sütun"),
        ("2. Temizleme", "Metin boşlukları silinir, tipler dönüştürülür; boş TotalCharges değerleri 0 ile doldurulur ve işaretlenir.",
         f"{tr(k['satir'])} kayıt (telco_temiz.csv)"),
        ("3. Tekrar eden kayıtlar", f"Kimlik ve etiket dışındaki {len(t['anahtar_sutunlari'])} özelliği aynı olan kayıtlardan ilki tutulur.",
         f"{tr(t['cikarilan_kayit'])} kayıt çıkarılır; {tr(t['kayit_sonrasi'])} kayıt (telco_modelleme.csv)"),
        ("4. Girdi seçimi", f"{len(A.SAYISAL)} sayısal ve {len(A.KATEGORIK)} kategorik sütun modele girer (cinsiyet ve yaş grubu dahil); kimlik, eş ve bakmakla yükümlü olunan kişi bilgisi ve etiket girmez.",
         f"{len(A.SAYISAL) + len(A.KATEGORIK)} girdi"),
        ("5. Ayırma", f"Katmanlı (stratified) rastgele ayırma, tohum {m['rastgele_tohum']}.",
         f"%{round((1 - A.TEST_ORANI) * 100)} eğitim: {tr(m['veri']['egitim'])}; %{round(A.TEST_ORANI * 100)} test: {tr(m['veri']['test'])}"),
        ("6. Ön işleme", "Sayısal sütunlar standartlaştırılır, kategorik sütunlar bir-sıcak kodlanır. Dönüşümler yalnızca eğitim kısmından "
         "(çapraz doğrulamada her katın eğitim kısmından) öğrenilir.", f"{tr(s['bagimlilik']['bir_sicak_sutun_sayisi'])} model sütunu"),
        ("7. Çapraz doğrulama", f"Eğitim kısmında {A.CV_KAT} katlı, iç içe: dış döngü başarıyı, iç döngü model ayarını seçer.",
         "Ortalama ± standart sapma"),
        ("8. Eşik seçimi", "Eşik, eğitim kısmındaki kat-dışı tahminlerde F1 skorunu en yükseğe çıkaran olasılıktır.", f"Eşik {tr(m['secilen_esik'], 2)}"),
        ("9. Test", "Test kısmı yalnızca burada, bir kez kullanılır; güven aralıkları bootstrap ile hesaplanır.",
         f"{tr(m['bootstrap_tekrar'])} yeniden örnekleme"),
        ("10. Açıklama", "Eğitilmiş modellerin test kısmındaki tahminleri SHAP ve permütasyon önemiyle açıklanır.", "Bölüm 6.8"),
        ("11. Karar katmanı", "Test müşterileri portföy sayılır; sentetik teklif tablosu ve bütçe optimizasyonu bu müşterilere uygulanır.",
         f"{tr(len(v['risk']))} müşterilik portföy"),
    ], [3.2, 8.4, 4.4], ad="akis")


def shap_bolumu(r, v):
    s = v["shap"]
    onem = pd.DataFrame(s["degisken_onemi"]).set_index("degisken")
    grup = pd.DataFrame(s["grup_onemi"]).set_index("grup")
    bag, kor, ornek, ilk5 = s["bagimlilik"], s["korelasyon"], s["ornek_musteri"], s["ilk5"]
    celiski = pd.DataFrame(s["toplam_ucret_celiskisi"])
    yon = pd.DataFrame(s["deger_yonu"]).set_index("etiket")["ortalama_shap"]
    turetilmis = s["turetilmis"]
    sira = list(onem.index)
    kucuk = lambda x: "sıfır" if x == 0 else ("10⁻¹⁴'ün altında" if x < 1e-14 else f"{x:.1e}")

    r.baslik("6.8. SHAP analizi: kullanılmalı mı?", 2)
    r.p("SHAP kullanılmasını öneriyoruz; ancak yalnızca modelin tahminini açıklayan bir araç olarak ve birbirinden türetilebilen "
        "değişkenler gruplanarak. Ayrılmanın nedenini ya da bir teklifin etkisini kanıtlayan bir araç olarak kullanılmamalıdır. "
        "Gerekçeler aşağıdadır.", giris="Öneri:")
    r.p("Aşağıdaki nedenlerle:", giris="Neden öneriyoruz?")
    r.maddeler(
        "Gerekçe göstermek zorunludur. Yönetici ekranında önerinin gerekçesi görünmelidir (Bölüm 13). SHAP, bir müşterinin tahminini "
        "değişkenlerin katkılarına toplamsal olarak ayırır: ortalama tahmin ile katkıların toplamı modelin gerçek çıktısına eşittir "
        f"(Lundberg ve Lee, 2017). Bu raporda bu özellik sayısal olarak doğrulanmıştır (en büyük fark {kucuk(s['denetim']['toplamsallik_farki'])}; "
        "yani yalnızca kayan nokta yuvarlama hatası).",
        "Aynı çerçeve hem tek müşteriyi (yerel açıklama) hem bütün portföyü (küresel açıklama) açıklar. Doğrusal modelde SHAP değeri "
        "katsayı × (müşterinin değeri − eğitim ortalaması) formülüne eşittir; ağaç modellerinde TreeSHAP kesin sonuç verir (Lundberg vd., 2020). "
        f"Yani sonuçlar yaklaşık değil kesindir; doğrusal formülle kütüphane arasındaki fark {kucuk(s['denetim']['dogrusal_formul_farki'])} bulunmuştur.",
        "Telekom müşteri kaybı çalışmalarında yaygın olarak kullanılır (Chang vd., 2024; Poudel vd., 2024; El Attar ve El-Hajj, 2026); "
        "sonuçlar literatürle aynı dilde sunulur.",
    )
    r.p("Aşağıdaki sınırlar nedeniyle:", giris="Neden dikkatli olunmalı?")
    r.maddeler(
        "SHAP modeli açıklar, dünyayı değil. Shapley değerlerine dayanan açıklamalar insan merkezli açıklama hedefleri için doğal bir "
        "çözüm değildir ve nedensel akıl yürütme gerektirir (Kumar vd., 2020). Bu projede ayrılma riski ile teklifin etkisi bilerek "
        "ayrı tutulmuştur (Bölüm 7): SHAP'in “Teknik destek yok” için riski artıran bir katkı göstermesi, destek vermenin kaybı azalttığı "
        "anlamına gelmez.",
        "Bağımlı değişkenler. Shapley tabanlı yöntemler değişkenlerin bağımsız olduğunu varsayabilir; bağımlılık varsa açıklamalar "
        "doğrusal modelde bile yanıltıcı olabilir (Aas vd., 2021). Bu veride bağımlılık çok güçlüdür: aylık ücret hizmet sütunlarıyla "
        f"doğrusal olarak R² = {tr(bag['aylik_ucret_r2_hizmetlerle'], 4)} düzeyinde açıklanır (ücret hizmet tarifesinden gelir); toplam ücret ile "
        f"süre × aylık ücret arasındaki korelasyon {tr(bag['toplam_ucret_korelasyonu_sure_x_ucret'], 4)}'dür; {bag['bir_sicak_sutun_sayisi']} "
        f"model sütununun {bag['birebir_ayni_bir_sicak_sutun']} tanesi başka bir sütunun birebir kopyasıdır.",
        "Doğrusal modelde SHAP, katsayılardan yeni bilgi üretmez; asıl katkısı müşteri bazında ayrıştırma ve görselleştirmedir. "
        "Doğrusal olmayan modellerde (ör. ağaç toplulukları) değeri daha büyüktür.",
        "Açıklama modele bağlıdır: aynı veriden eğitilen iki modelin SHAP sıralamaları farklı olabilir. Bu nedenle sonuçlar bir "
        "ikinci modelle (karar ağacı, TreeSHAP) ve SHAP'ten bağımsız bir ölçüyle (permütasyon önemi) karşılaştırılmıştır.",
    )

    r.p("Lojistik regresyonun SHAP değerleri test kısmındaki müşteriler için hesaplanmış, bir-sıcak kodlanmış sütunların katkıları "
        f"özgün değişkene toplanmıştır ({len(sira)} değişken). Permütasyon önemi, bir değişkenin (ya da grubun) test kısmında rastgele "
        "karıştırılıp ROC-AUC'deki düşüşe bakılmasıyla hesaplanır; modelin iç yapısından bağımsızdır.", giris="Yöntem:")
    ilk3 = sira[:3]
    r.sekil(A.GRAFIK / "11_shap_genel_onem.png", "Lojistik regresyon: değişkenlerin ortalama |SHAP| etkisi (test verisi)", 12)
    isaretli_olanlar = [d for d in ilk3 if d in turetilmis]
    r.p(f"Tahmini en çok etkileyen üç değişken sırasıyla {DEGISKEN_ADI[ilk3[0]]}, {DEGISKEN_ADI[ilk3[1]]} ve {DEGISKEN_ADI[ilk3[2]]} olup "
        f"ortalama |SHAP| değerleri {', '.join(tr(onem.loc[d, 'lojistik_shap'], 2) for d in ilk3)} düzeyindedir (log-oran birimi). "
        + (f"Bunlardan {', '.join(DEGISKEN_ADI[d] for d in isaretli_olanlar)} birbirinden türetilebilen bir değişkendir (şekilde açık renkli "
           "çubuk); aşağıdaki bulgu bu değişkenin tek başına yorumlanmaması gerektiğini gösterir." if isaretli_olanlar else ""))

    ters = (celiski.iloc[0]["ortalama_shap"] < 0 < celiski.iloc[-1]["ortalama_shap"]
            and celiski.iloc[0]["kayip_orani"] > celiski.iloc[-1]["kayip_orani"])
    r.tablo("Toplam ücretin çeyreklerinde gerçek kayıp oranı ve ortalama SHAP katkısı (test verisi)",
            ["Toplam ücret çeyreği", "Müşteri", "Gerçek kayıp oranı", "Ortalama SHAP katkısı"],
            [(c["ceyrek"], tr(c["musteri"]), yuzde(c["kayip_orani"]), isaretli(c["ortalama_shap"])) for _, c in celiski.iterrows()],
            [5.0, 3.0, 4.0, 4.0], saga=(1, 2, 3), ad="shap_celiski")
    if ters:
        r.p(f"Bağımlılığın sonucu burada somut olarak görülür: toplam ücretin en düşük çeyreğinde gerçek kayıp oranı "
            f"{yuzde(celiski.iloc[0]['kayip_orani'], 0)} iken ortalama SHAP katkısı {isaretli(celiski.iloc[0]['ortalama_shap'])} (riski azaltıyor "
            f"görünür); en yüksek çeyrekte kayıp oranı {yuzde(celiski.iloc[-1]['kayip_orani'], 0)} iken katkı "
            f"{isaretli(celiski.iloc[-1]['ortalama_shap'])} (riski artırıyor görünür). Tek başına bakıldığında SHAP işareti gerçek "
            "örüntünün tersidir. Bunun büyük olasılıkla nedeni, modelin toplam ücreti abonelik süresi ve aylık ücretle birlikte "
            "(toplam ücret ≈ süre × aylık ücret) bir düzeltme terimi olarak kullanmasıdır; üç değişkenin katkıları birbirini dengeler. "
            "Bu yüzden toplam ücret ya da aylık ücret için tek başına yorum yapılmaz; üçü “Süre ve ücret” grubu olarak birlikte "
            "değerlendirilir. Bağımlı değişkenlerin Shapley değerlerini gruplar hâlinde toplamak, bu soruna önerilen çözümlerden biridir (Aas vd., 2021).")
    else:
        r.p("Bu veride toplam ücretin SHAP işareti gerçek örüntüyle çelişmemektedir; yine de üç ücret ve süre değişkeni birlikte yorumlanmıştır.")

    olcut_adi = {"lojistik_shap": "lojistik regresyon SHAP", "agac_shap": "karar ağacı SHAP", "izin_onemi": "permütasyon önemi"}
    r.tablo("Değişken gruplarının önemi: üç ölçüye göre pay ve sıra (test verisi)",
            ["Grup", "SHAP, lojistik", "SHAP, ağaç", "Permütasyon önemi"],
            [(g, *(f"{yuzde(x[f'{o}_pay'], 0)} ({int(x[f'{o}_sira'])}.)" for o in olcut_adi)) for g, x in grup.iterrows()],
            [5.2, 3.6, 3.6, 3.6], saga=(1, 2, 3), ad="shap_grup")
    r.sekil(A.GRAFIK / "12_shap_grup_onemi.png", "Değişken gruplarının önemi: üç ölçüye göre pay (test verisi)", 13)
    birinci = {o: grup[f"{o}_pay"].idxmax() for o in olcut_adi}
    sonuncu = {o: grup[f"{o}_pay"].idxmin() for o in olcut_adi}
    r.p("Birinci sıradaki grup " + ", ".join(f"{olcut_adi[o]} ölçüsünde “{birinci[o]}”" for o in olcut_adi) + " olarak bulunmuştur. "
        + (f"En az önemli grup üç ölçüde de “{next(iter(sonuncu.values()))}” grubudur. " if len(set(sonuncu.values())) == 1 else
           "En az önemli grup ölçüye göre değişmektedir. ")
        + ("Üç ölçü birinci sırada anlaşmaktadır." if len(set(birinci.values())) == 1 else
           "Üç ölçü birinci sırada anlaşmamaktadır; yani hangi grubun “en önemli” olduğu modele ve ölçüye bağlıdır. Bu yüzden sonuçlar "
           "tek bir modelin ya da tek bir ölçünün kesin hükmü olarak sunulmamalıdır."))

    ilk8 = onem.head(8)
    r.tablo("En önemli sekiz değişken: üç ölçü (test verisi)",
            ["Değişken", "Lojistik SHAP", "Sıra", "Ağaç SHAP payı", "Sıra", "Permütasyon önemi", "Sıra"],
            [(DEGISKEN_ADI[d] + ("*" if d in turetilmis else ""), tr(x["lojistik_shap"], 2), f"{x['lojistik_shap_sira']}.", yuzde(x["agac_pay"], 0),
              f"{x['agac_shap_sira']}.", tr(x["izin_onemi"], 3), f"{x['izin_onemi_sira']}.") for d, x in ilk8.iterrows()],
            [3.8, 2.4, 1.2, 2.6, 1.2, 3.5, 1.3], saga=(1, 2, 3, 4, 5, 6), ad="shap_degisken")
    r.p("* Süre ve ücretlerden türetilebilen değişken. Lojistik SHAP: ortalama |SHAP| (log-oran); ağaç SHAP payı: toplam önemin payı; "
        "permütasyon önemi: ROC-AUC'deki ortalama düşüş.", italik=True)
    ortak = len(set(ilk5["lojistik_shap"]) & set(ilk5["agac_shap"]))
    r.p(f"Değişken sıralamaları arasındaki Spearman korelasyonu lojistik SHAP ile permütasyon önemi için {tr(kor['lojistik_shap__izin_onemi'], 2)} "
        f"(aynı model, farklı ölçü), lojistik SHAP ile karar ağacı SHAP'i için {tr(kor['lojistik_shap__agac_shap'], 2)} (farklı model) bulunmuştur. "
        f"İlk beş değişkenin {ortak} tanesi iki modelde ortaktır. Aynı modelin iki ölçüsü büyük ölçüde birbirini doğrularken iki modelin "
        "açıklaması daha az örtüşmektedir; bu, SHAP açıklamasının modele bağlı olduğunu gösterir.")

    r.sekil(A.GRAFIK / "13_shap_deger_yonu.png", "Değişken değerlerine göre ortalama SHAP etkisi (test verisi)", 13)
    def yon_degeri(etiket):
        return yon.get(etiket)
    parcalar = [(e, yon_degeri(e)) for e in ("Abonelik süresi: 0–12 ay", "Abonelik süresi: 49–72 ay", "Sözleşme: aylık (taahhütsüz)",
                                              "Sözleşme: 2 yıllık", "İnternet: fiber")]
    parcalar = [f"{e.lower() if False else e} {isaretli(x)}" for e, x in parcalar if x is not None]
    g = yon.get
    anahtarlar = ("Abonelik süresi: 0–12 ay", "Abonelik süresi: 49–72 ay", "Sözleşme: aylık (taahhütsüz)", "Sözleşme: 2 yıllık", "İnternet: fiber")
    uyum = all(g(a) is not None for a in anahtarlar) and g(anahtarlar[0]) > 0 > g(anahtarlar[1]) and g(anahtarlar[2]) > 0 > g(anahtarlar[3]) and g(anahtarlar[4]) > 0
    r.p("Değerlere göre yön şöyledir (ortalama SHAP, log-oran): " + "; ".join(parcalar) + ". Abonelik süresi, sözleşme türü ve internet "
        "türü için bu yönler " + ("Bölüm 5'teki kayıp oranı örüntüleriyle uyumludur: kısa süreli, taahhütsüz ve fiber müşterilerde risk artmakta, "
                                 "uzun süreli ve iki yıllık sözleşmeli müşterilerde azalmaktadır." if uyum else
                                 "Bölüm 5'teki kayıp oranı örüntüleriyle kısmen farklıdır; ayrıntı yukarıdaki şekildedir."))

    r.sekil(A.GRAFIK / "14_shap_ornek_musteri.png", f"Tek müşterinin (en yüksek riskli test müşterisi, {ornek['kimlik']}) risk ayrıştırması", 13)
    dk = ornek["degisken_katkilari"]
    r.p(f"Örnek müşterinin ayrılma olasılığı {yuzde(ornek['olasilik'], 0)} olup log-oran biriminde ortalama {tr(ornek['taban_log_oran'], 2)} düzeyindeki "
        f"başlangıca {isaretli(ornek['toplam_katki'])} katkı eklenerek {tr(ornek['log_oran'], 2)} bulunur. Değişken düzeyinde bu müşteri için "
        + ", ".join(f"{ad.lower()} {isaretli(x)}" for ad, x in dk.items())
        + " katkı verir; aralarında işareti ters olan katkılar bulunması bağımlılığın bir başka görünümüdür. Üçünün toplamı "
        f"({isaretli(sum(dk.values()))}) “Süre ve ücret” grubunun katkısıdır ve şekilde bu toplam gösterilmiştir.")

    r.p("Bu nedenle SHAP bu projede şu biçimde kullanılmıştır:", giris="Karar:")
    r.maddeler(
        "SHAP, modelin tahminini açıklamak için kullanılır; yönetici ekranının teknik sekmesinde ve bu raporda gösterilir.",
        "Birbirinden türetilebilen değişkenler (süre, aylık ücret, toplam ücret) tek tek değil grup olarak yorumlanır.",
        "Sonuçlar ikinci bir modelle ve SHAP'ten bağımsız bir ölçüyle çapraz denetlenir; ölçüler ayrışan yerlerde kesin hüküm verilmez.",
        "Teklif seçimi SHAP'e dayanmaz: tekliflerin etkisi ayrı tutulan varsayımlardan gelir ve ancak kontrol gruplu bir deneyle doğrulanabilir.",
        "Müşteri başına gösterilen “risk işaretleri” (Bölüm 6.6) yönetici ekranında kalmaya devam eder; SHAP onların yerini almaz, onları tamamlar.",
    )


def main():
    baslik_yaz("Proje raporu")
    v = verileri_topla()
    r = Rapor()
    kapak_ve_icindekiler(r)
    yonetici_ozeti(r, v)
    problem(r)
    veri(r, v)
    kalite(r, v)
    tekrar_kayitlari(r, v)
    veri_akisi(r, v)
    kesif(r, v)
    model(r, v)
    shap_bolumu(r, v)
    teklif(r, v)
    optimizasyon(r, v)
    senaryolar(r, v)
    ekran(r)
    sinirliliklar(r)
    sonuc(r, v)
    kabul_kriterleri(r)
    kaynakca(r)

    A.RAPOR.parent.mkdir(parents=True, exist_ok=True)
    try:
        r.doc.save(str(A.RAPOR))
    except PermissionError:
        sys.exit(f"Rapor yazılamadı: {A.RAPOR.name} Word'de açık olabilir. Dosyayı kapatıp yeniden deneyin.")
    print(f"Yazıldı: {A.RAPOR.relative_to(A.KOK)}  ({r.sekil_no} şekil, {r.tablo_no} tablo)")
    if word_ile_tamamla(A.RAPOR):
        print(f"İçindekiler güncellendi; PDF kopyası: {A.RAPOR.with_suffix('.pdf').relative_to(A.KOK)}")
    else:
        print("Word bulunamadı ya da açılamadı: içindekiler tablosunu Word'de sağ tıklayıp güncelleyin.")


if __name__ == "__main__":
    main()
