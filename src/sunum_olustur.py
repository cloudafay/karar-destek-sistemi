"""Proje sunumunu (PowerPoint) üretir: docs/KDS_Sunum.pptx.

Sunumdaki bütün rakamlar outputs/ altındaki çıktılardan okunur; proje yeniden çalıştırılıp
sonuçlar değişirse bu betiği yeniden çalıştırmak sunumu günceller. Kapaktaki ad, ders ve
öğretim elemanı src/ayarlar.py içindeki RAPOR_ satırlarından gelir. Word kuruluysa PDF
kopyası da yazılır.

Çalıştırma:  py src/sunum_olustur.py   (önce hepsini_calistir.py çalışmış olmalı)
"""
import json
import subprocess
import sys

import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt
from lxml import etree

import ayarlar as A
from ortak import baslik_yaz, para, tr, yuzde

# --- Görünüm -----------------------------------------------------------------------
LACIVERT, LACIVERT2, LACIVERT_CIZGI = "0C1B2E", "14283F", "24405F"
KAGIT, KART, CIZGI = "F3F5F8", "FBFCFD", "DDE3EC"
METIN, SOLUK = "33415A", "56657B"
MAVI, MAVI_KOYU, MAVI_ACIK, MAVI_ACIK_ZEMIN = "2A78D6", "1C5CAB", "6DA7EC", "DCE9FA"
TURUNCU, TURUNCU_KOYU, TURUNCU_ZEMIN, KREM = "EB6834", "B84A1A", "FBE3D6", "FFE3D3"
ALA, KOYU_METIN, KOYU_SOLUK = "199E70", "C5D3E6", "8DA2BE"
BASLIK_YAZI, GOVDE_YAZI, SIMGE_YAZI = "Segoe UI Semibold", "Segoe UI", "Segoe UI Symbol"
SOL, GENIS = 0.9, 11.53          # sayfa kenar boşluğu ve içerik genişliği (inç)
TOPLAM_SLAYT = 16


def renk(hex_):
    return RGBColor.from_string(hex_)


def buyuk(metin):
    """Türkçe büyük harfe çevirir: i → İ, ı → I."""
    return metin.replace("i", "İ").replace("ı", "I").upper()


# --- Çizim yardımcıları ---------------------------------------------------------------
def alfa(renk_ogesi, yuzde_opaklik):
    """Bir çizgi/dolgu rengine saydamlık ekler (python-pptx doğrudan sunmaz)."""
    for onceki in renk_ogesi.findall(qn("a:alpha")):
        renk_ogesi.remove(onceki)
    etree.SubElement(renk_ogesi, qn("a:alpha")).set("val", str(int(yuzde_opaklik * 1000)))


def arka_plan(s, hex_):
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = renk(hex_)


def sekil(s, tur, x, y, w, h, dolgu=None, cizgi=None, cizgi_kalinlik=1.0, yaricap=None, ad=None):
    sk = s.shapes.add_shape(tur, Inches(x), Inches(y), Inches(w), Inches(h))
    if yaricap is not None and tur == MSO_SHAPE.ROUNDED_RECTANGLE:
        sk.adjustments[0] = yaricap
    if dolgu:
        sk.fill.solid()
        sk.fill.fore_color.rgb = renk(dolgu)
    else:
        sk.fill.background()
    if cizgi:
        sk.line.color.rgb = renk(cizgi)
        sk.line.width = Pt(cizgi_kalinlik)
    else:
        sk.line.fill.background()
    sk.shadow.inherit = False
    if ad:
        sk.name = ad
    return sk


def kart(s, x, y, w, h, dolgu=KART, cizgi=CIZGI, ad="Kart", yaricap=0.06):
    return sekil(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, dolgu, cizgi, 0.75, yaricap, ad)


def yaz(s, x, y, w, h, icerik, boyut=16, rengi=METIN, kalin=False, yazi=GOVDE_YAZI, hiza=PP_ALIGN.LEFT,
        dikey=MSO_ANCHOR.TOP, madde=False, aralik=None, satir=None, ad="Metin", harf_araligi=None, italik=False):
    """Metin kutusu. `icerik`: metin ya da paragraf listesi; paragraf metin ya da [(metin, {kalin, renk}), ...]."""
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tb.name = ad
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = dikey
    paragraflar = [icerik] if isinstance(icerik, str) else icerik
    for i, parca in enumerate(paragraflar):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = hiza
        if aralik is not None and i < len(paragraflar) - 1:
            p.space_after = Pt(aralik)
        if satir:
            p.line_spacing = satir
        if madde:
            ppr = p._p.get_or_add_pPr()
            ppr.set("marL", "285750")
            ppr.set("indent", "-285750")
            etree.SubElement(ppr, qn("a:buFont")).set("typeface", "Arial")
            etree.SubElement(ppr, qn("a:buChar")).set("char", "•")
        parcalar = [(parca, {})] if isinstance(parca, str) else parca
        for metin, bicim in parcalar:
            r = p.add_run()
            r.text = metin
            r.font.name = bicim.get("yazi", yazi)
            r.font.size = Pt(bicim.get("boyut", boyut))
            r.font.bold = bicim.get("kalin", kalin)
            r.font.italic = bicim.get("italik", italik)
            r.font.color.rgb = renk(bicim.get("renk", rengi))
            if harf_araligi:
                r._r.get_or_add_rPr().set("spc", str(harf_araligi))
    return tb


def hap(s, x, y, w, metin, dolgu, yazi_rengi, boyut=11):
    """Küçük etiket (hap biçimli)."""
    sekil(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, 0.34, dolgu, None, yaricap=0.5, ad="Etiket arka planı")
    yaz(s, x, y, w, 0.34, buyuk(metin), boyut, yazi_rengi, True, hiza=PP_ALIGN.CENTER, dikey=MSO_ANCHOR.MIDDLE,
        ad="Etiket", harf_araligi=100)


def numara_dairesi(s, x, y, boyut, metin, dolgu=MAVI, yazi_rengi="FFFFFF", yazi=BASLIK_YAZI, punto=16):
    sekil(s, MSO_SHAPE.OVAL, x, y, boyut, boyut, dolgu, None, ad="Simge dairesi")
    yaz(s, x, y, boyut, boyut, metin, punto, yazi_rengi, True, yazi=yazi, hiza=PP_ALIGN.CENTER,
        dikey=MSO_ANCHOR.MIDDLE, ad="Simge")


def sinyal_cubuklari(s, x, y, rengi):
    for i, yukseklik in enumerate((0.085, 0.14, 0.195)):
        sekil(s, MSO_SHAPE.ROUNDED_RECTANGLE, x + i * 0.085, y + 0.195 - yukseklik, 0.055, yukseklik, rengi, None,
              yaricap=0.3, ad="Sinyal çubuğu")


def resim(s, yol, x, y, w, h, aciklama, kirp_alt=0.0):
    pic = s.shapes.add_picture(str(yol), Inches(x), Inches(y))
    if kirp_alt:
        pic.crop_bottom = kirp_alt
    pic.left, pic.top, pic.width, pic.height = Inches(x), Inches(y), Inches(w), Inches(h)
    pic._element.nvPicPr.cNvPr.set("descr", aciklama)
    pic.name = aciklama[:60]
    return pic


def notlar(s, metin):
    s.notes_slide.notes_text_frame.text = metin


def yeni_slayt(prs, no, eyebrow, baslik, koyu=False, vurgu=False):
    """Boş bir slayt açar; zemin, sinyal motifi, küçük başlık, başlık ve alt bilgiyi koyar."""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    zemin = MAVI_KOYU if vurgu else (LACIVERT if koyu else KAGIT)
    arka_plan(s, zemin)
    ana = "F3F5F8" if (koyu or vurgu) else LACIVERT
    ek = KREM if vurgu else (MAVI_ACIK if koyu else MAVI_KOYU)
    alt = "CFE0F5" if vurgu else (KOYU_SOLUK if koyu else SOLUK)
    sinyal_cubuklari(s, SOL, 0.72, ek)
    yaz(s, SOL + 0.42, 0.72, 8, 0.3, buyuk(eyebrow), 12, ek, True, dikey=MSO_ANCHOR.MIDDLE, ad="Bölüm etiketi",
        harf_araligi=250)
    yaz(s, SOL, 1.1, GENIS, 0.75, baslik, 34, ana, False, BASLIK_YAZI, ad="Başlık")
    yaz(s, SOL, 7.0, 8, 0.25, "Karar Destek Sistemi · Müşteri kaybı ve kampanya bütçesi", 11, alt, ad="Alt bilgi")
    yaz(s, SOL + GENIS - 2, 7.0, 2, 0.25, f"{no} / {TOPLAM_SLAYT}", 11, alt, hiza=PP_ALIGN.RIGHT, ad="Sayfa numarası")
    return s


# --- Grafik yardımcıları ----------------------------------------------------------------
def grafik_yazi_ayari(ch, boyut, rengi):
    ch.font.size = Pt(boyut)
    ch.font.name = GOVDE_YAZI
    ch.font.color.rgb = renk(rengi)


def etiket_yaz(nokta, metin, boyut, rengi, konum=XL_LABEL_POSITION.OUTSIDE_END, kalin=True):
    dl = nokta.data_label
    dl.position = konum
    tf = dl.text_frame
    tf.text = metin
    f = tf.paragraphs[0].runs[0].font
    f.size, f.bold, f.name = Pt(boyut), kalin, GOVDE_YAZI
    f.color.rgb = renk(rengi)


def sutun_grafigi(s, x, y, w, h, kategoriler, degerler, etiketler, renkler, ust=None, boyut=12):
    cd = CategoryChartData()
    cd.categories = kategoriler
    cd.add_series("Değer", degerler)
    ch = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(x), Inches(y), Inches(w), Inches(h), cd).chart
    ch.has_legend, ch.has_title = False, False
    grafik_yazi_ayari(ch, boyut, METIN)
    plot = ch.plots[0]
    plot.gap_width, plot.vary_by_categories = 55, False
    va = ch.value_axis
    va.visible, va.has_major_gridlines, va.minimum_scale = False, False, 0
    if ust:
        va.maximum_scale = ust
    ca = ch.category_axis
    ca.has_major_gridlines = False
    ca.format.line.color.rgb = renk("C3C2B7")
    ca.tick_labels.font.size = Pt(boyut)
    seri = plot.series[0]
    for i, (r, e) in enumerate(zip(renkler, etiketler)):
        nokta = seri.points[i]
        nokta.format.fill.solid()
        nokta.format.fill.fore_color.rgb = renk(r)
        etiket_yaz(nokta, e, boyut + 2, LACIVERT)
    return ch


def gruplu_yatay_grafik(s, x, y, w, h, kategoriler, seriler, renkler, etiketler, boyut=12):
    """Yan yana çubuk grafik (yatay); `etiketler[seri][kategori]` her çubuğun yazısı."""
    cd = CategoryChartData()
    cd.categories = kategoriler
    for ad, degerler in seriler.items():
        cd.add_series(ad, degerler)
    ch = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(x), Inches(y), Inches(w), Inches(h), cd).chart
    ch.has_title = False
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.TOP
    ch.legend.include_in_layout = False
    grafik_yazi_ayari(ch, boyut, METIN)
    plot = ch.plots[0]
    plot.gap_width, plot.overlap = 60, -5
    va = ch.value_axis
    va.visible, va.has_major_gridlines, va.minimum_scale = False, False, 0
    va.maximum_scale = max(max(v) for v in seriler.values()) * 1.22
    ca = ch.category_axis
    ca.reverse_order = True
    ca.has_major_gridlines = False
    ca.format.line.color.rgb = renk("C3C2B7")
    ca.tick_labels.font.size = Pt(boyut)
    for i, (seri, r) in enumerate(zip(plot.series, renkler)):
        seri.format.fill.solid()
        seri.format.fill.fore_color.rgb = renk(r)
        for j in range(len(kategoriler)):
            etiket_yaz(seri.points[j], etiketler[i][j], boyut - 1, METIN, kalin=False)
    return ch


# --- Veriler -----------------------------------------------------------------------------
def veriler():
    m = json.loads(A.MODEL_METRIKLERI.read_text(encoding="utf-8"))
    plan = json.loads(A.PLAN_OZETI.read_text(encoding="utf-8"))
    liste = pd.read_csv(A.KAMPANYA_LISTESI)
    onerilen = liste[liste["offer_type"] != A.TEKLIF_YOK]
    risk = pd.read_csv(A.RISK_TABLOSU)
    sen = pd.read_csv(A.TABLO / "senaryo_karsilastirma.csv").set_index("senaryo")
    dua = pd.read_csv(A.TABLO / "duyarlilik_etki_carpani.csv", index_col="etki_carpani")
    dua.index = dua.index.round(2)
    kesif = pd.read_csv(A.TABLO / "kesif_kayip_oranlari.csv").set_index(["degisken", "grup"])["kayip_orani"]
    temiz = pd.read_csv(A.TEMIZ_VERI)
    return {
        "m": m, "plan": plan["onerilen_plan"], "kisitsiz": plan["kisitsiz_plan"], "tekrar": json.loads(A.TEKRAR_OZETI.read_text(encoding="utf-8")),
        "shap": json.loads(A.SHAP_OZETI.read_text(encoding="utf-8")), "liste": liste, "onerilen": onerilen, "risk": risk,
        "sen": sen, "dua": dua, "kesif": kesif, "temiz": temiz,
        "teklif_ozeti": onerilen.groupby("offer_type").agg(n=("customerID", "size"), maliyet=("offer_cost", "sum")),
    }


# --- Slaytlar ------------------------------------------------------------------------------
def kapak(prs, v):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    arka_plan(s, LACIVERT)
    # Sinyal dalgalarını andıran yaylar (sağ alt köşe) ve birbirine bağlı düğümler
    for yaricap, opaklik in ((1.4, 90), (2.4, 65), (3.4, 45), (4.4, 30), (5.4, 18)):
        daire = sekil(s, MSO_SHAPE.OVAL, 13.333 - yaricap, 7.5 - yaricap, yaricap * 2, yaricap * 2, None, MAVI, 2.25, ad="Sinyal yayı")
        alfa(daire.line._get_or_add_ln().find(qn("a:solidFill")).find(qn("a:srgbClr")), opaklik)
    for (x, y, r, dolgu) in ((9.9, 4.6, 0.09, MAVI_ACIK), (11.3, 3.95, 0.1, TURUNCU), (11.0, 6.1, 0.09, MAVI_ACIK)):
        sekil(s, MSO_SHAPE.OVAL, x - r, y - r, 2 * r, 2 * r, dolgu, None, ad="Düğüm")
    sekil(s, MSO_SHAPE.OVAL, 11.3 - 0.2, 3.95 - 0.2, 0.4, 0.4, None, TURUNCU, 1.5, ad="Düğüm halkası")
    sinyal_cubuklari(s, SOL, 0.9, MAVI_ACIK)
    yaz(s, SOL + 0.42, 0.9, 8, 0.3, "KARAR DESTEK SİSTEMİ PROJESİ", 12, MAVI_ACIK, True, dikey=MSO_ANCHOR.MIDDLE, harf_araligi=250, ad="Bölüm etiketi")
    yaz(s, SOL, 2.3, 10.5, 1.3, "Karar Destek Sistemi", 60, "F3F5F8", False, BASLIK_YAZI, ad="Başlık")
    yaz(s, SOL, 3.75, 8.4, 1.2, "Müşteri kaybı riskini tahmin etme ve kampanya bütçesini optimize etme", 24, KOYU_METIN, ad="Alt başlık")
    yaz(s, SOL, 5.85, 8, 0.9, [[(A.RAPOR_HAZIRLAYAN, {"renk": "F3F5F8", "kalin": True})], A.RAPOR_DERS + " · " + A.RAPOR_OGRETIM_ELEMANI],
        16, KOYU_SOLUK, aralik=4, ad="Sunan")
    notlar(s, "Merhaba, bu sunumda karar destek sistemi projemizi anlatacağım. Proje iki soruya birlikte yanıt veriyor: hangi müşterilerin "
              "ayrılma riski yüksek ve sınırlı bir kampanya bütçesi hangi müşteriye, hangi teklifle harcanmalı? Önce problemi ve sistemin "
              "yapısını göstereceğim, sonra veriyi, doğrulamayı ve sonuçları. Toplam süre yaklaşık 12–15 dakika.")


def problem(prs, v):
    s = yeni_slayt(prs, 2, "01 · Problem", "Herkese indirim vermek doğru değil")
    yaz(s, SOL, 2.2, 4.9, 1.6, "Abonelik işletmeleri müşteri kaybıyla karşılaşır; ama her teklifin bir maliyeti vardır. Bütçe sınırlıdır "
        "ve yanlış müşteriye harcanan her dolar kârlılığı düşürür.", 17, METIN, satir=1.1, ad="Giriş metni")
    kart(s, SOL, 3.95, 4.9, 2.45, MAVI_KOYU, None, "Proje sorusu kartı")
    yaz(s, SOL + 0.3, 4.2, 4.3, 0.3, "PROJENİN SORUSU", 11, KREM, True, harf_araligi=200, ad="Etiket")
    yaz(s, SOL + 0.3, 4.6, 4.3, 1.6, "Hangi müşteriye, hangi teklifi, ne kadar bütçeyle sunmalı?", 24, "F3F5F8", False, BASLIK_YAZI, ad="Soru")
    satirlar = (("−", SOLUK, "Zaten kalacak müşteri", "Teklif yalnızca maliyet yaratır."),
                ("!", TURUNCU_KOYU, "Ne olursa olsun ayrılacak müşteri", "Teklif işe yaramaz."),
                ("✓", MAVI_KOYU, "Teklifin işe yarayacağı müşteri", "Bütçe asıl buraya gitmeli."))
    for i, (simge, r, baslik, metin) in enumerate(satirlar):
        y = 2.2 + i * 1.45
        kart(s, 6.3, y, 6.13, 1.3, ad="Müşteri türü kartı")
        numara_dairesi(s, 6.6, y + 0.32, 0.66, simge, r, "FFFFFF", SIMGE_YAZI, 20)
        yaz(s, 7.5, y + 0.25, 4.7, 0.4, baslik, 17, LACIVERT, True, ad="Kart başlığı")
        yaz(s, 7.5, y + 0.7, 4.7, 0.4, metin, 14, METIN, ad="Kart metni")
    notlar(s, "Problemle başlayalım. Müşteri kaybını önlemek için herkese indirim vermek ilk akla gelen çözüm, ama doğru değil. Müşterilerin bir "
              "kısmı zaten kalacak; onlara verilen teklif yalnızca maliyet. Bir kısmı ne yapsak ayrılacak; onlara verilen teklif işe yaramıyor. "
              "Teklifin gerçekten fark yarattığı müşteriler bu ikisinin arasında. Projenin sorusu: hangi müşteriye, hangi teklifi, ne kadar bütçeyle sunmalıyız?")


def sistem(prs, v):
    s = yeni_slayt(prs, 3, "02 · Sistem", "Sistem dört katmandan oluşuyor", koyu=True)
    adimlar = (("1", "Tahmin", "Her müşteri için ayrılma olasılığı (lojistik regresyon)"),
               ("2", "Teklif değerlendirme", "Teklif başına beklenen net fayda: ek katkı − maliyet"),
               ("3", "Optimizasyon", "Bütçe ve kapasite altında toplam faydayı en yükseğe çıkarır"),
               ("4", "Karar ekranı", "Öneri, gerekçe ve belirsizlik; karar yöneticidedir"))
    genislik = (GENIS - 3 * 0.5) / 4
    for i, (no, baslik, metin) in enumerate(adimlar):
        x = SOL + i * (genislik + 0.5)
        kart(s, x, 2.2, genislik, 3.05, LACIVERT2, LACIVERT_CIZGI, "Katman kartı")
        numara_dairesi(s, x + 0.25, 2.45, 0.55, no, MAVI, "FFFFFF", BASLIK_YAZI, 18)
        yaz(s, x + 0.25, 3.2, genislik - 0.5, 0.7, baslik, 18, "F3F5F8", True, ad="Katman başlığı")
        yaz(s, x + 0.25, 3.95, genislik - 0.5, 1.2, metin, 14, KOYU_METIN, ad="Katman metni")
        if i < 3:
            sekil(s, MSO_SHAPE.RIGHT_ARROW, x + genislik + 0.1, 3.55, 0.3, 0.3, MAVI_ACIK, None, ad="Ok")
    kart(s, SOL, 5.55, GENIS, 1.0, LACIVERT2, LACIVERT_CIZGI, "Not kartı")
    yaz(s, SOL + 0.35, 5.55, GENIS - 0.7, 1.0, "Tahmin tek başına yetmez: kampanyanın ek etkisi ve maliyeti ayrıca hesaplanır; yönetici kararı "
        "gerekçesiyle birlikte sunulur.", 17, "F3F5F8", dikey=MSO_ANCHOR.MIDDLE, ad="Not")
    notlar(s, "Sistemi dört katman olarak kurduk. Birinci katman her müşteri için ayrılma olasılığını tahmin ediyor. İkinci katman her teklifin müşteriye "
              "göre beklenen net faydasını hesaplıyor: teklif sunulmasa da kalacak bir müşteriyi kampanyanın başarısı saymıyoruz. Üçüncü katman bütçe ve "
              "ekip kapasitesi altında en yüksek toplam faydayı sağlayan müşteri–teklif seçimini buluyor. Dördüncü katman bunu yöneticiye gerekçesi ve "
              "belirsizliğiyle gösteriyor; karar yöneticide.")


def veri(prs, v):
    t, temiz = v["tekrar"], v["temiz"]
    ayrilan = int(temiz[A.HEDEF].eq("Yes").sum())
    s = yeni_slayt(prs, 4, "03 · Veri", "Veri gerçek, kampanya etkisi varsayım")
    kart(s, SOL, 2.2, 5.6, 4.45, ad="Gerçek veri kartı")
    hap(s, SOL + 0.3, 2.45, 1.75, "Gerçek veri", MAVI_ACIK_ZEMIN, MAVI_KOYU)
    satirlar = ((tr(len(temiz)), "müşteri kaydı"), (str(temiz.shape[1] - 1), "sütun: hizmet, sözleşme, ücret"),
                (yuzde(ayrilan / len(temiz)), f"ayrıldı ({tr(ayrilan)} müşteri)"))
    for i, (sayi, etiket) in enumerate(satirlar):
        y = 3.0 + i * 0.78
        yaz(s, SOL + 0.3, y, 2.3, 0.7, sayi, 34, LACIVERT, False, BASLIK_YAZI, ad="Rakam", dikey=MSO_ANCHOR.MIDDLE)
        yaz(s, SOL + 2.55, y, 2.9, 0.7, etiket, 16, METIN, ad="Rakam etiketi", dikey=MSO_ANCHOR.MIDDLE)
    yaz(s, SOL + 0.3, 5.45, 5.0, 1.0, f"IBM Telco Customer Churn (Kaggle). Tekrar eden {tr(t['cikarilan_kayit'])} kayıt çıkarıldı; "
        f"modelleme verisi {tr(t['kayit_sonrasi'])} kayıt. Cinsiyet ve yaş grubu, ön sunumdaki geri bildirim üzerine modele eklendi.",
        12, SOLUK, ad="Kaynak notu")
    kart(s, SOL + 5.93, 2.2, 5.6, 4.45, ad="Sentetik varsayım kartı")
    hap(s, SOL + 6.23, 2.45, 2.2, "Sentetik varsayım", TURUNCU_ZEMIN, TURUNCU_KOYU)
    yaz(s, SOL + 6.23, 3.0, 5.0, 0.6, "Kampanya etkisi veride yok", 22, LACIVERT, False, BASLIK_YAZI, ad="Kart başlığı")
    yaz(s, SOL + 6.23, 3.65, 5.0, 0.4, "Bu yüzden şunlar senaryo parametresidir:", 16, METIN, ad="Giriş")
    yaz(s, SOL + 6.23, 4.15, 5.0, 1.4, ["Teklif maliyetleri ve etki oranları", f"Kampanya bütçesi: {para(A.BUTCE)}",
        f"Teknik destek kapasitesi: {A.DESTEK_KAPASITESI} müşteri"], 16, METIN, madde=True, aralik=4, ad="Parametreler")
    yaz(s, SOL + 6.23, 5.65, 5.0, 0.8, "Sonuçlar yöntemi gösterir; gerçek bir kampanyanın başarısını kanıtlamaz.", 13, TURUNCU_KOYU, ad="Uyarı")
    notlar(s, f"Veri iki parçadan oluşuyor ve bunları baştan ayırıyoruz. Müşteri verisi gerçek: IBM Telco Customer Churn, {tr(len(temiz))} müşteri, "
              "yaklaşık dörtte biri ayrılmış. Ama bu veride kampanya bilgisi yok; hangi müşteriye hangi teklifin sunulduğunu ya da işe yarayıp "
              "yaramadığını bilmiyoruz. Bu yüzden teklif maliyetleri, etki oranları, bütçe ve kapasite bizim seçtiğimiz senaryo parametreleri; raporda "
              "tek tek gerekçesiyle yazılı. Sonuçlar yöntemi gösteriyor, gerçek bir kampanyanın sonucunu kanıtlamıyor. Bir de ön sunumdan sonra yaptığımız "
              "değişiklik var: cinsiyet ve yaş grubunu başta ayrımcılık riski yüzünden modele koymamıştık; geri bildiriminiz üzerine ikisini de girdi olarak "
              "ekledik ve modelin kadın-erkek, 65 yaş altı-üstü gruplarında dengeli çalışıp çalışmadığını ayrıca kontrol ediyoruz.")


def kesif(prs, v):
    k = v["kesif"]
    s = yeni_slayt(prs, 5, "04 · Keşifsel analiz", "Kayıp ilk yılda ve taahhütsüzlükte")
    genel = v["temiz"][A.HEDEF].eq("Yes").mean()
    for yon, (x, baslik, kat, anahtarlar, alt) in enumerate((
        (SOL, "Sözleşme türüne göre kayıp oranı", ["Aylık (taahhütsüz)", "1 yıllık", "2 yıllık"],
         [("Sözleşme türü", g) for g in ("Month-to-month", "One year", "Two year")],
         f"Aylık sözleşmede kayıp {yuzde(k['Sözleşme türü', 'Month-to-month'])}; iki yıllıkta yalnızca {yuzde(k['Sözleşme türü', 'Two year'])}."),
        (SOL + 5.93, "Abonelik süresine göre kayıp oranı", ["0–12 ay", "13–24 ay", "25–48 ay", "49–72 ay"],
         [("Abonelik süresi", g) for g in ("0–12 ay", "13–24 ay", "25–48 ay", "49–72 ay")],
         f"İlk 12 ayda kayıp {yuzde(k['Abonelik süresi', '0–12 ay'])}; 49–72. ayda {yuzde(k['Abonelik süresi', '49–72 ay'])}."))):
        kart(s, x, 2.2, 5.6, 3.75, ad="Grafik kartı")
        yaz(s, x + 0.3, 2.4, 5.0, 0.35, baslik, 14, SOLUK, True, ad="Grafik başlığı")
        degerler = [float(k[a]) for a in anahtarlar]
        sutun_grafigi(s, x + 0.2, 2.8, 5.2, 3.0, kat, degerler, [yuzde(d) for d in degerler],
                      [MAVI_KOYU if d == max(degerler) else MAVI_ACIK for d in degerler], ust=max(degerler) * 1.25)
        yaz(s, x, 6.1, 5.6, 0.55, alt, 16, LACIVERT, True, ad="Grafik özeti")
    notlar(s, "Modeli kurmadan önce kaybın nerede yoğunlaştığına baktık. İki örüntü çok belirgin. Birincisi sözleşme türü: aylık, taahhütsüz sözleşmede "
              f"müşterilerin {yuzde(k['Sözleşme türü', 'Month-to-month'])}’si ayrılmış; iki yıllık sözleşmede yalnızca {yuzde(k['Sözleşme türü', 'Two year'])}. "
              f"İkincisi abonelik süresi: ilk on iki ayda {yuzde(k['Abonelik süresi', '0–12 ay'])}, dört ile altı yıl arasında {yuzde(k['Abonelik süresi', '49–72 ay'])}. "
              "Bu grafikler ilişkiyi gösteriyor, nedensellik kanıtlamıyor; bu yüzden teklif etkilerini veriden türetmedik, ayrı varsayım olarak tuttuk.")


def dogrulama(prs, v):
    m, t, L = v["m"], v["tekrar"], "Lojistik regresyon"
    etki = pd.DataFrame(t["etki"])
    onceki = etki.iloc[0]
    s = yeni_slayt(prs, 6, "05 · Doğrulama", "Doğrulama sızıntıya karşı korunuyor")
    egitim_yuzde = round((1 - A.TEST_ORANI) * 100)
    kartlar = ((f"%{egitim_yuzde} / %{100 - egitim_yuzde}", "Katmanlı ayırma",
                f"{tr(m['veri']['egitim'])} eğitim, {tr(m['veri']['test'])} test müşterisi. Test kısmı yalnızca bir kez kullanıldı."),
               (f"{A.CV_KAT} kat", "İç içe çapraz doğrulama",
                f"Ayar seçimi her katın kendi içinde yapılır; ROC-AUC {tr(m['capraz_dogrulama'][L]['roc_auc'], 3)} ± {tr(m['capraz_dogrulama_std'][L]['roc_auc'], 3)}."),
               (f"{tr(t['cikarilan_kayit'])} kayıt", "Tekrar eden kayıtlar çıkarıldı",
                f"Özellikleri birebir aynı kayıtlar. Sızıntı: test {onceki['test_ikizi_19']} → 0, çapraz doğrulama {onceki['cv_sizintisi_19']} → 0."))
    genislik = (GENIS - 2 * 0.3) / 3
    for i, (rakam, baslik, metin) in enumerate(kartlar):
        x = SOL + i * (genislik + 0.3)
        kart(s, x, 2.2, genislik, 3.2, ad="Önlem kartı")
        yaz(s, x + 0.3, 2.45, genislik - 0.6, 0.8, rakam, 32, MAVI_KOYU, False, BASLIK_YAZI, ad="Rakam")
        yaz(s, x + 0.3, 3.35, genislik - 0.6, 0.6, baslik, 18, LACIVERT, True, ad="Kart başlığı")
        yaz(s, x + 0.3, 4.05, genislik - 0.6, 1.3, metin, 14, METIN, ad="Kart metni")
    guven = m["test_guven_araligi"][L]["roc_auc"]
    kart(s, SOL, 5.7, GENIS, 0.95, MAVI_ACIK_ZEMIN, None, "Sonuç kartı")
    yaz(s, SOL + 0.35, 5.7, GENIS - 0.7, 0.95, [[("Test ROC-AUC ", {}), (tr(m["test"][L]["roc_auc"], 3), {"kalin": True}),
        (f" (%95 aralık {tr(guven[0], 3)}–{tr(guven[1], 3)}), çapraz doğrulamayla tutarlı.", {})]], 18, LACIVERT, dikey=MSO_ANCHOR.MIDDLE, ad="Sonuç")
    notlar(s, "Sonuçların iyimser çıkmaması için üç önlem aldık. Birincisi, veriyi katmanlı ayırdık; test kısmına yalnızca en sonda bir kez baktık. İkincisi, "
              "eğitim kısmında beş katlı iç içe çapraz doğrulama yaptık: model ayarı her katın kendi içinde seçiliyor, böylece sonuç şişmiyor. Üçüncüsü, "
              f"özellikleri birebir aynı olan {tr(t['cikarilan_kayit'])} tekrar eden kaydı ayırmadan önce çıkardık. Çıkarmadan önce {onceki['test_ikizi_19']} test kaydının "
              f"eğitimde ikizi vardı ve çapraz doğrulamada {onceki['cv_sizintisi_19']} kayıt sızıntı yapıyordu; şimdi ikisi de sıfır. Başarıya etkisi çok küçük, "
              "ama doğrulamanın temiz olduğu artık garanti.")


def model(prs, v):
    m, risk, L, T = v["m"], v["risk"], "Lojistik regresyon", "Karar ağacı"
    e = m["test"][L]["secilen_esik"]
    guven = m["test_guven_araligi"][L]["roc_auc"]
    s = yeni_slayt(prs, 7, "06 · Tahmin modeli", f"Model: ROC-AUC {tr(m['test'][L]['roc_auc'], 2)} ile iyi sıralıyor")
    kutular = ((tr(m["test"][L]["roc_auc"], 3), f"Test ROC-AUC (%95 aralık {tr(guven[0], 3)}–{tr(guven[1], 3)})"),
               (tr(m["capraz_dogrulama"][L]["roc_auc"], 3), f"Çapraz doğrulama ROC-AUC (± {tr(m['capraz_dogrulama_std'][L]['roc_auc'], 3)})"),
               (yuzde(e["duyarlilik"], 0), f"Ayrılanların yakalanma oranı (eşik {tr(m['secilen_esik'], 2)})"),
               (yuzde(e["kesinlik"], 0), "Riskli denenlerin gerçekten ayrılma oranı"))
    for i, (rakam, etiket) in enumerate(kutular):
        x, y = SOL + (i % 2) * 2.75, 2.2 + (i // 2) * 2.25
        kart(s, x, y, 2.6, 2.05, ad="Ölçüt kartı")
        yaz(s, x + 0.25, y + 0.2, 2.2, 0.8, rakam, 32, MAVI_KOYU, False, BASLIK_YAZI, ad="Rakam")
        yaz(s, x + 0.25, y + 1.05, 2.15, 0.9, etiket, 13, METIN, ad="Ölçüt açıklaması")
    kart(s, SOL + 5.75, 2.2, 5.78, 4.3, "FCFCFB", CIZGI, "Grafik kartı")
    resim(s, A.GRAFIK / "06_roc_pr_egrileri.png", SOL + 5.85, 2.45, 5.58, 2.675, "ROC ve kesinlik–duyarlılık eğrileri: lojistik regresyon ve karar ağacı")
    yaz(s, SOL + 5.95, 5.35, 5.3, 1.1, f"Lojistik regresyon (mavi), karar ağacından (test ROC-AUC {tr(m['test'][T]['roc_auc'], 3)}) biraz daha iyi sıralıyor.",
        14, METIN, ad="Grafik açıklaması")
    g = risk.groupby("risk_grubu")["gercek_churn"].mean()
    notlar(s, f"Temel model lojistik regresyon; karşılaştırma için bir karar ağacı da eğittik. Lojistik regresyon test verisinde {tr(m['test'][L]['roc_auc'], 3)} "
              f"ROC-AUC veriyor; yüzde doksan beş aralığı {tr(guven[0], 3)} ile {tr(guven[1], 3)}. Çapraz doğrulama ortalaması {tr(m['capraz_dogrulama'][L]['roc_auc'], 3)} ile "
              f"tutarlı. Olasılık eşiğini {tr(m['secilen_esik'], 2)} seçtik: bu eşikte gerçekten ayrılan müşterilerin {yuzde(e['duyarlilik'], 0)}’i yakalanıyor, riskli denenlerin "
              f"{yuzde(e['kesinlik'], 0)}’ü gerçekten ayrılıyor. Test verisinde yüksek riskli grupta gerçek kayıp {yuzde(g['Yüksek'], 0)}, düşük riskli grupta {yuzde(g['Düşük'], 0)} civarında. "
              "Doğrulukla yetinmedik: hiç kimseyi riskli saymayan bir kural bile yüksek doğruluk alır ama kimseyi yakalamaz.")


def teklifler(prs, v):
    s = yeni_slayt(prs, 8, "07 · Teklifler", "Dört seçenek, bir net fayda hesabı")
    kartlar = (("%", MAVI_KOYU, A.INDIRIM, f"{A.INDIRIM_AY} ay boyunca aylık ücrette {yuzde(A.INDIRIM_ORANI, 0)} indirim. Aylık sözleşmede etkisi daha yüksek.", KART),
               ("⇄", MAVI_KOYU, A.PAKET, f"Aylık ücreti {A.PAKET_ALT_SINIR} {A.PARA} ve üzeri olan müşteriyi daha uygun pakete geçirme.", KART),
               ("✚", MAVI_KOYU, A.DESTEK, f"İnterneti olup desteği olmayan müşteriye ücretsiz destek ({A.DESTEK_MALIYETI} {A.PARA}).", KART),
               ("○", SOLUK, A.TEKLIF_YOK, "Net fayda pozitif değilse en iyi karar budur.", "E8ECF2"))
    genislik = (GENIS - 3 * 0.25) / 4
    for i, (simge, r, baslik, metin, zemin) in enumerate(kartlar):
        x = SOL + i * (genislik + 0.25)
        kart(s, x, 2.2, genislik, 2.6, zemin, CIZGI, "Teklif kartı")
        numara_dairesi(s, x + 0.25, 2.45, 0.55, simge, r, "FFFFFF", SIMGE_YAZI, 18)
        yaz(s, x + 0.25, 3.15, genislik - 0.5, 0.4, baslik, 18, LACIVERT, True, ad="Teklif adı")
        yaz(s, x + 0.25, 3.65, genislik - 0.5, 1.1, metin, 14, METIN, ad="Teklif açıklaması")
    kart(s, SOL, 5.1, GENIS, 1.55, LACIVERT, None, "Formül kartı")
    hap(s, SOL + 0.35, 5.3, 3.6, "Maliyet ve etkiler sentetik varsayım", TURUNCU, LACIVERT)
    yaz(s, SOL + 0.35, 5.8, GENIS - 0.7, 0.7, f"Beklenen net fayda = ayrılma olasılığı × teklifin etkisi × {A.UFUK_AY} aylık katkı − teklif maliyeti",
        19, "F3F5F8", False, BASLIK_YAZI, dikey=MSO_ANCHOR.MIDDLE, ad="Formül")
    notlar(s, "Her müşteri için dört seçeneği karşılaştırıyoruz: indirim, daha uygun pakete geçiş, teknik destek ve hiç teklif sunmamak. Hepsini aynı formülle "
              "değerlendiriyoruz: ayrılma olasılığı çarpı teklifin etkisi çarpı on iki aylık katkı, eksi teklifin maliyeti. Etki oranı, teklif olmasaydı ayrılacak "
              "olanların ne kadarının teklifle kalacağını söylüyor; böylece zaten kalacak bir müşterinin gelirini kampanyanın başarısı saymıyoruz. Önemli: "
              "maliyetler ve etki oranları gerçek veriden gelmiyor, senaryo varsayımı. Hangi teklifin kime uygun olduğunu keşifsel analizdeki örüntülere "
              "dayandırdık, ama sayıların kendisi yuvarlak varsayımlar ve raporun yedinci bölümünde dayanaklarıyla listeli.")


def optimizasyon(prs, v):
    p, k, ozet = v["plan"], v["kisitsiz"], v["teklif_ozeti"]
    s = yeni_slayt(prs, 9, "08 · Optimizasyon", f"{para(A.BUTCE)} bütçe kime ayrılıyor?", koyu=True)
    yaz(s, SOL, 2.15, 4.2, 1.5, tr(p["teklif_sunulan_musteri"]), 66, "F3F5F8", False, BASLIK_YAZI, ad="Müşteri sayısı")
    yaz(s, SOL, 3.55, 4.2, 0.5, "müşteriye teklif öneriliyor", 18, KOYU_METIN, ad="Etiket")
    yaz(s, SOL, 4.3, 4.2, 0.3, "Beklenen net fayda", 12, KOYU_SOLUK, ad="Etiket")
    yaz(s, SOL, 4.6, 4.2, 0.9, para(p["beklenen_net_fayda"]), 40, MAVI_ACIK, False, BASLIK_YAZI, ad="Net fayda")
    yaz(s, SOL, 5.65, 4.2, 1.0, f"Sınırsız bütçeyle plan {tr(k['teklif_sunulan_musteri'])} müşteri ve {para(k['maliyet'])} olurdu; bütçe bunun {yuzde(A.BUTCE / k['maliyet'], 0)}’ü.",
        12, KOYU_SOLUK, ad="Not")
    x0, genis = SOL + 4.9, 6.6
    renkler = {A.DESTEK: "3987E5", A.INDIRIM: ALA, A.PAKET: "D95926"}
    toplam = float(ozet["maliyet"].sum())
    x = x0
    siralama = sorted(A.TEKLIFLER, key=lambda t: -ozet.loc[t, "maliyet"])
    for sira, teklif in enumerate(siralama):
        w = genis * ozet.loc[teklif, "maliyet"] / toplam
        sekil(s, MSO_SHAPE.RECTANGLE, x, 2.25, max(w - 0.04, 0.05), 0.55, renkler[teklif], None, ad=f"Bütçe dilimi: {teklif}")
        yaz(s, x, 2.25, max(w - 0.04, 0.05), 0.55, yuzde(ozet.loc[teklif, "maliyet"] / toplam), 13, LACIVERT, True, hiza=PP_ALIGN.CENTER, dikey=MSO_ANCHOR.MIDDLE, ad="Dilim yüzdesi")
        x += w
    for i, teklif in enumerate(siralama):
        y = 3.05 + i * 0.5
        sekil(s, MSO_SHAPE.ROUNDED_RECTANGLE, x0, y + 0.08, 0.2, 0.2, renkler[teklif], None, yaricap=0.2, ad="Gösterge")
        yaz(s, x0 + 0.4, y, 2.2, 0.36, teklif, 16, "F3F5F8", ad="Teklif adı", dikey=MSO_ANCHOR.MIDDLE)
        yaz(s, x0 + 2.55, y, 4.0, 0.36, f"{tr(ozet.loc[teklif, 'n'])} müşteri · {para(ozet.loc[teklif, 'maliyet'])}", 16, KOYU_SOLUK, ad="Teklif özeti", dikey=MSO_ANCHOR.MIDDLE)
    kart(s, x0, 4.7, genis, 1.95, LACIVERT2, LACIVERT_CIZGI, "Kısıt kartı")
    yaz(s, x0 + 0.3, 4.85, 5.0, 0.3, "KISITLAR", 11, MAVI_ACIK, True, harf_araligi=200, ad="Etiket")
    yaz(s, x0 + 0.3, 5.2, 6.0, 1.4, [f"Toplam maliyet ≤ {para(A.BUTCE)}", "Müşteriye en fazla bir teklif",
        f"Teknik destek ≤ {A.DESTEK_KAPASITESI} müşteri"], 16, "F3F5F8", madde=True, aralik=3, ad="Kısıtlar")
    notlar(s, "Optimizasyon, bütçe ve kapasite sınırları altında toplam beklenen net faydayı en yükseğe çıkaran müşteri–teklif seçimini buluyor. Kısıtlar üç tane: "
              f"toplam maliyet bütçeyi aşmaz, bir müşteriye en fazla bir teklif sunulur, teknik destek ekibi en fazla {A.DESTEK_KAPASITESI} müşteriye bakabilir. "
              f"Sonuç: {tr(p['teklif_sunulan_musteri'])} müşteriye teklif öneriliyor; beklenen net fayda {para(p['beklenen_net_fayda'])}. Teknik destek ucuz ve müşteri başına "
              f"net faydası yüksek olduğu için ilk sırada. Bütçe sınırsız olsaydı {tr(k['teklif_sunulan_musteri'])} müşteriye teklif sunar, {para(k['maliyet'])} harcardık; "
              f"{para(A.BUTCE)} bunun {yuzde(A.BUTCE / k['maliyet'], 0)}’ü, yani bütçe gerçekten kısıtlıyor ve seçim yapmak anlamlı. Net faydası pozitif olmayan teklif hiç sunulmuyor.")


def senaryo(prs, v):
    sen = v["sen"]
    a, b, c = (sen.loc[ad] for ad in (A.SENARYO_A, A.SENARYO_B, A.SENARYO_C))
    oran = c["beklenen_net_fayda"] / a["beklenen_net_fayda"]
    katsayi = "iki" if abs(oran - 2) < 0.15 else tr(oran, 1)
    s = yeni_slayt(prs, 10, "09 · Karşılaştırma", f"Önerilen plan {katsayi} katı fayda sağlıyor")
    kart(s, SOL, 2.2, 6.6, 3.45, ad="Grafik kartı")
    yaz(s, SOL + 0.3, 2.35, 6.0, 0.35, f"Beklenen net fayda ({A.PARA}), aynı bütçe: {para(A.BUTCE)}", 13, SOLUK, True, ad="Grafik başlığı")
    degerler = [float(a["beklenen_net_fayda"]), float(b["beklenen_net_fayda"]), float(c["beklenen_net_fayda"])]
    sutun_grafigi(s, SOL + 0.2, 2.7, 6.2, 2.85, ["A · Mevcut yöntem", "B · Yalnızca risk modeli", "C · Karar destek sistemi"], degerler,
                  [para(d) for d in degerler], [MAVI_ACIK, MAVI_ACIK, MAVI_KOYU], ust=max(degerler) * 1.2, boyut=12)
    for i, (ad_, x_) in enumerate((("Mevcut yöntem", a), ("Yalnızca risk modeli", b), ("Karar destek sistemi", c))):
        y = 2.2 + i * 1.18
        vurgulu = i == 2
        kart(s, SOL + 6.9, y, 4.63, 1.05, MAVI_KOYU if vurgulu else KART, None if vurgulu else CIZGI, "Senaryo özeti")
        yaz(s, SOL + 7.15, y, 2.2, 1.05, [[(("ABC"[i] + " · "), {"kalin": True}), (ad_, {})]], 14, "F3F5F8" if vurgulu else METIN, dikey=MSO_ANCHOR.MIDDLE, ad="Senaryo adı")
        yaz(s, SOL + 9.2, y, 2.15, 1.05, para(x_["beklenen_net_fayda"]), 22, "F3F5F8" if vurgulu else LACIVERT, False, BASLIK_YAZI, hiza=PP_ALIGN.RIGHT, dikey=MSO_ANCHOR.MIDDLE, ad="Senaryo faydası")
    kart(s, SOL, 5.85, GENIS, 0.85, TURUNCU_ZEMIN, None, "Uyarı kartı")
    b_geride = b["etiket_bazli_net_fayda"] < a["etiket_bazli_net_fayda"]
    yaz(s, SOL + 0.35, 5.85, GENIS - 0.7, 0.85, [[("Dikkat: ", {"kalin": True}), (f"gerçek etiketlerle ölçülünce {'B, A’yı geçemiyor' if b_geride else 'A ile B arasındaki fark kalkıyor'} "
        f"({para(a['etiket_bazli_net_fayda'])} ve {para(b['etiket_bazli_net_fayda'])}); C yine en yüksek ({para(c['etiket_bazli_net_fayda'])}).", {})]],
        15, LACIVERT, dikey=MSO_ANCHOR.MIDDLE, ad="Uyarı")
    notlar(s, "Üç yaklaşımı aynı bütçe ve aynı varsayımlarla karşılaştırdık. A, bugünkü yöntem: aylık sözleşmeli ve ilk on iki ayındaki müşterilere aynı indirimi vermek. "
              "B, yalnızca risk modeli: en riskli müşterilere aynı indirim. C, bizim sistem: müşteri ve teklif bazında optimizasyon. Beklenen net fayda A için "
              f"{para(a['beklenen_net_fayda'])}, B için {para(b['beklenen_net_fayda'])}, C için {para(c['beklenen_net_fayda'])}. İki not var. Birincisi, C’nin üstünlüğünün önemli "
              "bölümü daha ucuz bir teklifi, yani teknik desteği seçebilmesinden geliyor; A ve B yalnızca indirim sunabiliyor. İkincisi, beklenen fayda modelin kendi "
              "olasılıklarıyla hesaplanıyor; müşterilerin gerçekte ayrılıp ayrılmadığına bakan ikinci ölçümde "
              + ("risk modeli tek başına bugünkü yöntemi geçemiyor" if b_geride else "A ile B arasındaki fark kayboluyor")
              + ", C ise iki ölçümde de en yüksek. Hepsi sentetik varsayımlar altında, gerçek kampanya sonucu değil.")


def belirsizlik(prs, v):
    sen, dua = v["sen"], v["dua"]
    c, a, b = A.SENARYO_C, A.SENARYO_A, A.SENARYO_B
    s = yeni_slayt(prs, 11, "10 · Belirsizlik", "Ne kadar emin olabiliriz?")
    kartlar = ((0.5, "Kötümser · ×0,5", "Teklifler beklenenin yarısı kadar işe yararsa."),
               (1.0, "Beklenen · ×1,0", "Teklifler varsayıldığı gibi işe yararsa."),
               (1.5, "İyimser · ×1,5", "Teklifler beklenenden yarı yarıya daha iyi işe yararsa."))
    genislik = (GENIS - 2 * 0.3) / 3
    for i, (carpan, baslik, metin) in enumerate(kartlar):
        x = SOL + i * (genislik + 0.3)
        deger = float(dua.loc[carpan, c])
        vurgulu = carpan == 1.0
        kart(s, x, 2.2, genislik, 2.95, MAVI_KOYU if vurgulu else KART, None if vurgulu else CIZGI, "Belirsizlik kartı")
        yaz(s, x + 0.3, 2.4, genislik - 0.6, 0.4, baslik, 17, "F3F5F8" if vurgulu else LACIVERT, True, ad="Kart başlığı")
        yaz(s, x + 0.3, 2.9, genislik - 0.6, 0.8, para(deger), 36, "F3F5F8" if vurgulu else LACIVERT, False, BASLIK_YAZI, ad="Net fayda")
        yaz(s, x + 0.3, 3.8, genislik - 0.6, 0.7, metin, 14, "F3F5F8" if vurgulu else METIN, ad="Açıklama")
        yaz(s, x + 0.3, 4.6, genislik - 0.6, 0.35, "✓  Plan kendini karşılar" if deger > 0 else "✗  Plan zarar eder", 14, KREM if vurgulu else MAVI_KOYU, True, ad="Durum")
    kart(s, SOL, 5.4, GENIS, 1.2, TURUNCU_ZEMIN, None, "Uyarı kartı")
    yaz(s, SOL + 0.35, 5.4, GENIS - 0.7, 1.2, [[(f"Etki yarıya inerse A ({para(dua.loc[0.5, a])}) ve B ({para(dua.loc[0.5, b])}) "
        f"{'zarar eder' if dua.loc[0.5, a] < 0 and dua.loc[0.5, b] < 0 else 'daha az kazanır'}. Önerilen plan, etki varsayılanın ", {}),
        (yuzde(sen.loc[c, 'basabas_etki_carpani'], 0) + "’sine", {"kalin": True}), (" inene kadar maliyetini karşılar.", {})]],
        16, LACIVERT, dikey=MSO_ANCHOR.MIDDLE, ad="Uyarı")
    notlar(s, "Tekliflerin gerçek etkisini bilmediğimiz için planı sabit tutup etkinin varsayımdan sapması durumunda ne olacağına baktık. Etki varsayılanın yarısı "
              f"kadar çıkarsa önerilen planın net faydası {para(dua.loc[0.5, c])}; varsayıldığı gibi çıkarsa {para(dua.loc[1.0, c])}, yarı yarıya daha iyi çıkarsa {para(dua.loc[1.5, c])}. "
              f"Aynı kötümser durumda bugünkü yöntem ({para(dua.loc[0.5, a])}) ve yalnızca risk modeli ({para(dua.loc[0.5, b])}) için sonuç böyle. Plan, gerçek etki varsayılan "
              f"etkinin yaklaşık {yuzde(sen.loc[c, 'basabas_etki_carpani'], 0)}’sine inene kadar maliyetini karşılıyor. Bu yüzden önerimiz planı bir anda tüm portföye uygulamak "
              "değil, önce kontrol gruplu küçük bir pilotla gerçek etkiyi ölçmek.")


def shap_oneri(prs, v):
    s = yeni_slayt(prs, 12, "11 · Model açıklaması", "SHAP: kullanıyoruz, ama dikkatle")
    for i, (simge, r, baslik, maddeler) in enumerate((
            ("✓", MAVI_KOYU, "Neden öneriyoruz?", ["Tahmini müşteri bazında ayrıştırır", "Hesap kesin, tahminle birebir tutar", "Telekom çalışmalarında yaygın"]),
            ("!", TURUNCU_KOYU, "Neden dikkatli olmalıyız?", ["Modeli açıklar, nedensellik kanıtlamaz", "Bağımlı değişkenlerde yanıltabilir", "Model değişince sıralama değişir"]))):
        x = SOL + i * 5.93
        kart(s, x, 2.2, 5.6, 2.75, ad="Gerekçe kartı")
        numara_dairesi(s, x + 0.3, 2.45, 0.6, simge, r, "FFFFFF", SIMGE_YAZI, 20)
        yaz(s, x + 1.1, 2.45, 4.3, 0.6, baslik, 21, LACIVERT, False, BASLIK_YAZI, dikey=MSO_ANCHOR.MIDDLE, ad="Kart başlığı")
        yaz(s, x + 0.3, 3.3, 5.0, 1.5, maddeler, 17, METIN, madde=True, aralik=6, ad="Maddeler")
    kart(s, SOL, 5.2, GENIS, 1.45, LACIVERT, None, "Karar kartı")
    yaz(s, SOL + 0.4, 5.38, 4, 0.3, "KARARIMIZ", 11, MAVI_ACIK, True, harf_araligi=200, ad="Etiket")
    yaz(s, SOL + 0.4, 5.75, GENIS - 0.8, 0.8, "SHAP’i yalnızca açıklama aracı olarak kullanıyoruz; teklif seçimi SHAP’e dayanmıyor.", 21, "F3F5F8", False, BASLIK_YAZI, ad="Karar")
    notlar(s, "Dersin gereği olarak SHAP analizini projeye uyguladık ve kullanılıp kullanılmaması gerektiğini de tartıştık. Önerimiz evet, ama sınırlı kullanım. Neden "
              "önerdik: SHAP bir müşterinin tahminini değişkenlerin katkılarına ayırıyor, yönetici ekranında önerinin gerekçesi görünmeli; doğrusal modelde ve ağaç "
              "modellerinde hesap kesin, biz bunu sayısal olarak doğruladık; telekom müşteri kaybı çalışmalarında da yaygın. Neden dikkatli olduk: SHAP modeli açıklar, "
              "dünyayı değil. Örneğin teknik destek yok riski artırıyor görünse bile destek vermenin kaybı azalttığı anlamına gelmez; bu yüzden risk ile teklifin etkisini "
              "bilerek ayrı tuttuk. Ayrıca değişkenler birbirine bağlıysa açıklama yanıltabilir ve model değişince sıralama değişebilir. Sonraki slayt bunun somut örneğini gösteriyor.")


def shap_bulgu(prs, v):
    sh = v["shap"]
    celiski = pd.DataFrame(sh["toplam_ucret_celiskisi"])
    dusuk, yuksek = celiski.iloc[0], celiski.iloc[-1]
    isaret = lambda x: ("+" if x > 0 else "") + tr(x, 2)
    grup = pd.DataFrame(sh["grup_onemi"])
    s = yeni_slayt(prs, 13, "12 · Bulgu", "Bağımlı değişkenler SHAP’i yanıltıyor", vurgu=True)
    bloklar = ((dusuk, "En düşük toplam ücret çeyreğinde gerçek kayıp", "riski azaltıyor"), (yuksek, "En yüksek çeyrekte gerçek kayıp", "riski artırıyor"))
    for i, (satir, onek, yon) in enumerate(bloklar):
        y = 2.2 + i * 1.35
        kart(s, SOL, y, 6.4, 1.2, "154E8F", "3A78C0", "Bulgu kartı")
        yaz(s, SOL + 0.3, y, 1.8, 1.2, isaret(satir["ortalama_shap"]), 32, KREM, False, BASLIK_YAZI, dikey=MSO_ANCHOR.MIDDLE, ad="SHAP katkısı")
        yaz(s, SOL + 2.2, y, 4.0, 1.2, f"{onek} {yuzde(satir['kayip_orani'], 0)}; SHAP {yon} gösteriyor.", 15, "F3F5F8", dikey=MSO_ANCHOR.MIDDLE, ad="Açıklama")
    yaz(s, SOL, 4.95, 6.4, 1.6, f"Sebep: toplam ücret ≈ süre × aylık ücret (korelasyon {tr(sh['bagimlilik']['toplam_ucret_korelasyonu_sure_x_ucret'], 4)}); "
        f"aylık ücret hizmetlerden R² = {tr(sh['bagimlilik']['aylik_ucret_r2_hizmetlerle'], 4)} ile hesaplanıyor.", 16, KREM, ad="Sebep")
    kart(s, SOL + 6.75, 2.2, 4.78, 4.4, "FCFCFB", None, "Grafik kartı")
    yaz(s, SOL + 6.95, 2.35, 4.4, 0.35, "Değişken gruplarının önemi (pay)", 12, SOLUK, True, ad="Grafik başlığı")
    gruplar = list(grup["grup"])
    ch_seriler = {"SHAP (lojistik)": [float(x) for x in grup["lojistik_shap_pay"]], "SHAP (ağaç)": [float(x) for x in grup["agac_shap_pay"]],
                  "Permütasyon": [float(x) for x in grup["izin_onemi_pay"]]}
    etiketler = [[yuzde(x, 0) for x in d] for d in ch_seriler.values()]
    gruplu_yatay_grafik(s, SOL + 6.85, 2.7, 4.6, 3.55, gruplar, ch_seriler, [MAVI, TURUNCU, ALA], etiketler, boyut=11)
    sonuncular = {grup.loc[grup[f"{o}_pay"].idxmin(), "grup"] for o in ("lojistik_shap", "agac_shap", "izin_onemi")}
    yaz(s, SOL + 6.95, 6.25, 4.4, 0.3, f"“{next(iter(sonuncular))}” üç ölçüde de sonuncu." if len(sonuncular) == 1
        else "En az önemli grup ölçüye göre değişiyor.", 11, SOLUK, ad="Grafik notu")
    notlar(s, "Bu slayt SHAP’in neden dikkatle kullanılması gerektiğinin somut örneği. Toplam ücreti dört eşit gruba böldük. Toplam ücreti en düşük çeyrekte, yani yeni "
              f"müşterilerde, gerçek kayıp {yuzde(dusuk['kayip_orani'], 0)}; ama SHAP bu grup için {isaret(dusuk['ortalama_shap'])} katkı veriyor, yani riski azaltıyor görünüyor. En yüksek "
              f"çeyrekte gerçek kayıp {yuzde(yuksek['kayip_orani'], 0)}, SHAP ise {isaret(yuksek['ortalama_shap'])} diyor. Tek başına bakınca SHAP işareti gerçek örüntünün tersi. Sebep, toplam "
              "ücretin süre çarpı aylık ücretle neredeyse aynı olması; aylık ücret de hizmetlerden neredeyse tam hesaplanabiliyor. Model bu üç değişkeni birlikte kullandığı için "
              "katkıları birbirini dengeliyor. Çözüm olarak süre ve ücret değişkenlerini tek grup olarak yorumladık. Sağdaki grafik gruplar düzeyinde bile sıralamanın modele ve "
              "ölçüye göre değiştiğini gösteriyor.")


def ekran(prs, v):
    s = yeni_slayt(prs, 14, "13 · Yönetici ekranı", "Yönetici karar ekranı")
    satirlar = (("★", "Öneri ve beklenen net fayda tek bakışta"), ("⚙", "Bütçeyi değiştir, plan yeniden hesaplanır"),
                ("!", "Etki varsayımı düşerse ne olur?"), ("✓", "Karar ve gerekçe kayda geçer"))
    for i, (simge, metin) in enumerate(satirlar):
        y = 2.2 + i * 1.12
        kart(s, SOL, y, 4.6, 0.98, ad="Özellik kartı")
        numara_dairesi(s, SOL + 0.25, y + 0.21, 0.56, simge, TURUNCU_KOYU if simge == "!" else MAVI_KOYU, "FFFFFF", SIMGE_YAZI, 16)
        yaz(s, SOL + 1.0, y, 3.45, 0.98, metin, 15, LACIVERT, True, dikey=MSO_ANCHOR.MIDDLE, ad="Özellik")
    en, boy = 6.5, 4.0
    kirp = 1 - (boy / (en * 1478 / 1500))
    resim(s, A.EKRAN_GORUNTUSU, SOL + 5.03, 2.2, en, boy, "Yönetici karar ekranı: öneri, beklenen net fayda, bütçenin dağılımı ve yöntemlerin karşılaştırması", kirp_alt=kirp)
    notlar(s, "Sistemin çıktısı bir yönetici ekranı. İlk sekme kararı vermek için gereken her şeyi düz bir dille gösteriyor: önerilen müşteri sayısı ve beklenen net "
              "fayda, bütçenin teklif türlerine dağılımı, bugünkü yöntemle karşılaştırma ve tekliflerin etkisi beklenenden az ya da çok çıkarsa sonuç. Soldaki panelden bütçe ve "
              "teknik destek kapasitesini değiştirince optimizasyon yeniden çalışıyor ve bütün sayılar güncelleniyor. Yönetici onayla, değiştir ya da reddet seçeneklerinden birini "
              "gerekçesiyle kaydediyor. Teknik ayrıntılar, SHAP açıklamaları ve çapraz doğrulama sonuçları ayrı bir sekmede. Canlı demo için terminalde py -m streamlit run app.py yeterli.")


def sinirlar(prs, v):
    s = yeni_slayt(prs, 15, "14 · Sınırlılıklar", "Sınırlılıklar ve sonraki adımlar")
    for i, (simge, r, baslik, maddeler) in enumerate((
            ("!", TURUNCU_KOYU, "Sınırlılıklar", ["Kampanya etkisi doğrulanmadı (sentetik)", "Veri kesitsel: gelecek ay tahmini değil",
                                                  f"Tek rastgele ayırma (tohum {A.RASTGELE_TOHUM})", "SHAP nedensellik kanıtı değil"]),
            ("➜", MAVI_KOYU, "Sonraki adımlar", ["Kontrol gruplu pilot uygulama", "Gerçek etki ve maliyetle güncelleme", "Uplift modeli denemesi",
                                                 "Zamanlı veriyle ileriye dönük doğrulama"]))):
        x = SOL + i * 5.93
        kart(s, x, 2.2, 5.6, 3.35, ad="Liste kartı")
        numara_dairesi(s, x + 0.3, 2.45, 0.6, simge, r, "FFFFFF", SIMGE_YAZI, 20)
        yaz(s, x + 1.1, 2.45, 4.3, 0.6, baslik, 21, LACIVERT, False, BASLIK_YAZI, dikey=MSO_ANCHOR.MIDDLE, ad="Kart başlığı")
        yaz(s, x + 0.3, 3.3, 5.0, 2.1, maddeler, 17, METIN, madde=True, aralik=6, ad="Maddeler")
    kart(s, SOL, 5.8, GENIS, 0.85, MAVI_ACIK_ZEMIN, None, "Öneri kartı")
    yaz(s, SOL + 0.35, 5.8, GENIS - 0.7, 0.85, [[("Önerimiz: ", {"kalin": True}), ("önce küçük bir pilot ve kontrol grubu; etki ölçülünce genişletme.", {})]],
        18, LACIVERT, dikey=MSO_ANCHOR.MIDDLE, ad="Öneri")
    notlar(s, "Sınırlılıkları açıkça söylemek istiyorum. En önemlisi, kampanya etkisini doğrulamış değiliz: maliyetler ve etki oranları sentetik varsayım, dolayısıyla net "
              "fayda sayıları yöntemin nasıl çalıştığını gösteriyor, gerçek bir kampanyanın sonucunu değil. Veri kesitsel; model gelecek ay kimin ayrılacağını değil, geçmişte "
              "ayrılanlara benzerliği tahmin ediyor. Sonuçlar tek bir rastgele ayırmaya dayanıyor. SHAP açıklaması da nedensellik kanıtı değil. Sonraki adım olarak önce kontrol "
              "gruplu küçük bir pilot öneriyoruz; ölçülen gerçek etki ve maliyetle senaryo tablosunu güncelleyip, yeterli veri oluşunca bir uplift modeli denenebilir.")


def sonuc(prs, v):
    sen = v["sen"]
    oran = sen.loc[A.SENARYO_C, "beklenen_net_fayda"] / sen.loc[A.SENARYO_A, "beklenen_net_fayda"]
    katsayi = "iki" if abs(oran - 2) < 0.15 else tr(oran, 1)
    s = prs.slides.add_slide(prs.slide_layouts[6])
    arka_plan(s, LACIVERT)
    for yaricap, opaklik in ((1.1, 90), (1.9, 60), (2.7, 40), (3.5, 25), (4.3, 15)):
        daire = sekil(s, MSO_SHAPE.OVAL, 13.333 - yaricap, 7.5 - yaricap, yaricap * 2, yaricap * 2, None, MAVI, 2.25, ad="Sinyal yayı")
        alfa(daire.line._get_or_add_ln().find(qn("a:solidFill")).find(qn("a:srgbClr")), opaklik)
    sinyal_cubuklari(s, SOL, 0.72, MAVI_ACIK)
    yaz(s, SOL + 0.42, 0.72, 8, 0.3, "SONUÇ", 12, MAVI_ACIK, True, dikey=MSO_ANCHOR.MIDDLE, harf_araligi=250, ad="Bölüm etiketi")
    maddeler = ("Tahmin tek başına karar değildir; etki ve maliyet birlikte düşünülmeli.",
                f"Aynı bütçeyle optimize plan, bugünkü yöntemin yaklaşık {katsayi} katı beklenen fayda sağlıyor (varsayımlar altında).",
                "Önce pilot ve kontrol grubu; sonra genişletme.")
    genislik = (GENIS - 2 * 0.3) / 3
    for i, metin in enumerate(maddeler):
        x = SOL + i * (genislik + 0.3)
        kart(s, x, 1.55, genislik, 2.75, LACIVERT2, LACIVERT_CIZGI, "Çıkarım kartı")
        numara_dairesi(s, x + 0.3, 1.85, 0.55, str(i + 1), MAVI, "FFFFFF", BASLIK_YAZI, 18)
        yaz(s, x + 0.3, 2.65, genislik - 0.6, 1.55, metin, 17, "F3F5F8", ad="Çıkarım")
    yaz(s, SOL, 4.9, 9, 1.3, "Teşekkürler", 60, "F3F5F8", False, BASLIK_YAZI, ad="Teşekkür")
    yaz(s, SOL, 6.05, 6, 0.6, "Sorular?", 24, MAVI_ACIK, ad="Sorular")
    notlar(s, "Özetleyecek olursam üç şey söyleyebilirim. Birincisi, müşteri kaybını tahmin etmek tek başına karar vermeye yetmiyor; teklifin ek etkisi ve maliyeti ayrıca "
              f"hesaplanmalı. İkincisi, aynı bütçeyle optimize edilmiş plan, herkese aynı indirimi vermenin yaklaşık {katsayi} katı beklenen fayda sağlıyor; ama bu sentetik varsayımlar "
              "altında geçerli. Üçüncüsü, bunu gerçek bir kampanyada doğrulamak için önce kontrol gruplu küçük bir pilot gerekiyor. Dinlediğiniz için teşekkürler; sorularınızı alabilirim. "
              "Olası sorular: neden lojistik regresyon (açıklanabilir ve iyi kalibre), varsayımlar nereden geldi (raporun yedinci bölümü), tekrar eden kayıtlar neden çıkarıldı "
              "(rapor 4.2), SHAP neden dikkatle kullanıldı (rapor 6.8).")


def pdf_yaz(dosya):
    """PowerPoint kuruluysa sunumun PDF kopyasını yazar; değilse sessizce geçer."""
    pdf = dosya.with_suffix(".pdf")
    betik = f"""
        $p = New-Object -ComObject PowerPoint.Application
        try {{ $d = $p.Presentations.Open('{dosya}', $true, $false, $false); $d.SaveAs('{pdf}', 32); $d.Close() }} finally {{ $p.Quit() }}
    """
    try:
        sonuc_ = subprocess.run(["powershell", "-NoProfile", "-Command", betik], capture_output=True, timeout=240)
        return sonuc_.returncode == 0 and pdf.exists()
    except (OSError, subprocess.TimeoutExpired):
        return False


def main():
    baslik_yaz("Sunum")
    v = veriler()
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    for slayt in (kapak, problem, sistem, veri, kesif, dogrulama, model, teklifler, optimizasyon, senaryo, belirsizlik,
                  shap_oneri, shap_bulgu, ekran, sinirlar, sonuc):
        slayt(prs, v)
    assert len(prs.slides) == TOPLAM_SLAYT
    prs.core_properties.title = "Karar Destek Sistemi: müşteri kaybı ve kampanya bütçesi"
    prs.core_properties.author = A.RAPOR_HAZIRLAYAN
    yol = A.DOKUMAN / "KDS_Sunum.pptx"
    try:
        prs.save(str(yol))
    except PermissionError:
        sys.exit(f"Sunum yazılamadı: {yol.name} PowerPoint'te açık olabilir. Dosyayı kapatıp yeniden deneyin.")
    print(f"Yazıldı: {yol.relative_to(A.KOK)}  ({len(prs.slides)} slayt)")
    if pdf_yaz(yol):
        print(f"PDF kopyası: {yol.with_suffix('.pdf').relative_to(A.KOK)}")
    else:
        print("PowerPoint bulunamadı ya da açılamadı: PDF kopyası yazılmadı.")


if __name__ == "__main__":
    main()
