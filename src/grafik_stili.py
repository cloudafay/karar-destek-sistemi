"""Tüm grafiklerin ortak görünümü: renkler, yazı tipi ve çubuk biçimi."""
import matplotlib

matplotlib.use("Agg")  # ekran açmadan dosyaya çizer
import matplotlib.pyplot as plt
from matplotlib.patches import PathPatch
from matplotlib.path import Path as Yol
from matplotlib.ticker import FuncFormatter

import ayarlar as A
from ortak import tr, yuzde

YUZEY = "#fcfcfb"
MUREKKEP = "#0b0b0b"
IKINCIL = "#52514e"
SOLUK = "#898781"
IZGARA = "#e1e0d9"
EKSEN = "#c3c2b7"
# Renk körlüğü denetiminden geçmiş kategorik sıra; seriler bu sırayla renk alır.
SERI = ["#2a78d6", "#eb6834", "#1baf7a"]
ACIK_MAVI = "#cde2fb"

plt.rcParams.update({
    "font.family": ["Segoe UI", "DejaVu Sans"],
    "font.size": 10,
    "figure.dpi": 100,
    "savefig.dpi": 200,
    "figure.facecolor": YUZEY,
    "axes.facecolor": YUZEY,
    "axes.edgecolor": EKSEN,
    "axes.linewidth": 1,
    "axes.spines.top": False,
    "axes.spines.right": False,
    "axes.labelcolor": IKINCIL,
    "axes.axisbelow": True,
    "grid.color": IZGARA,
    "grid.linewidth": 1,
    "xtick.color": SOLUK,
    "ytick.color": SOLUK,
    "xtick.labelcolor": IKINCIL,
    "ytick.labelcolor": IKINCIL,
    "xtick.major.size": 0,
    "ytick.major.size": 0,
    "text.color": MUREKKEP,
    "lines.linewidth": 2,
    "lines.solid_capstyle": "round",
    "lines.solid_joinstyle": "round",
    "legend.frameon": False,
    "text.parse_math": False,  # "$" işareti para birimi olarak yazılsın
})

YUZDE_EKSENI = FuncFormatter(lambda deger, _: yuzde(deger, 0))
ONDALIK_EKSENI = FuncFormatter(lambda deger, _: tr(deger, 1))


def basliklar(fig, baslik, alt_baslik):
    fig.text(0.04, 0.96, baslik, fontsize=13, fontweight="bold", va="top")
    fig.text(0.04, 0.96 - 0.32 / fig.get_figheight(), alt_baslik, fontsize=10, color=IKINCIL, va="top")


def kaydet(fig, dosya_adi):
    A.GRAFIK.mkdir(parents=True, exist_ok=True)
    fig.savefig(A.GRAFIK / dosya_adi)
    plt.close(fig)
    print(f"Grafik: outputs/grafikler/{dosya_adi}")


def _cubuk_ciz(ax, konum, deger, yatay, renk, kalinlik_px=22, yaricap_px=4):
    """Taban çizgisinde düz, veri ucunda yuvarlatılmış tek bir çubuk çizer."""
    (x0, y0), (x1, y1) = ax.transData.transform([(0, 0), (1, 1)])
    birim_x, birim_y = abs(x1 - x0), abs(y1 - y0)  # bir veri biriminin piksel karşılığı
    if yatay:
        yari, rd, rk = kalinlik_px / 2 / birim_y, yaricap_px / birim_x, yaricap_px / birim_y
    else:
        yari, rd, rk = kalinlik_px / 2 / birim_x, yaricap_px / birim_y, yaricap_px / birim_x
    rd = min(rd, abs(deger)) * (1 if deger >= 0 else -1)
    # (d, k): d = değer ekseni, k = kategori ekseni
    noktalar = [
        (0, konum - yari), (deger - rd, konum - yari), (deger, konum - yari), (deger, konum - yari + rk),
        (deger, konum + yari - rk), (deger, konum + yari), (deger - rd, konum + yari), (0, konum + yari), (0, konum - yari),
    ]
    kodlar = [Yol.MOVETO, Yol.LINETO, Yol.CURVE3, Yol.CURVE3, Yol.LINETO, Yol.CURVE3, Yol.CURVE3, Yol.LINETO, Yol.CLOSEPOLY]
    if not yatay:
        noktalar = [(k, d) for d, k in noktalar]
    ax.add_patch(PathPatch(Yol(noktalar, kodlar), facecolor=renk, edgecolor="none"))


def cubuk_grafigi(etiketler, degerler, baslik, alt_baslik, dosya_adi, yatay=True,
                  bicim=yuzde, eksen_bicimi=YUZDE_EKSENI, referans=None, referans_etiketi=None):
    """Tek serili çubuk grafik. Sırasız kategoriler yatay, sıralı aralıklar dikey çizilir."""
    n = len(etiketler)
    degerler = [float(d) for d in degerler]
    en_kucuk, en_buyuk = min(0, *degerler), max(0, *degerler)
    pay = (en_buyuk - en_kucuk) * 0.16
    sinir = (en_kucuk - (pay if en_kucuk < 0 else 0), en_buyuk + pay)

    if yatay:
        fig, ax = plt.subplots(figsize=(7.2, 1.35 + 0.5 * n))
        fig.subplots_adjust(left=0.33, right=0.96, top=1 - 0.95 / fig.get_figheight(), bottom=0.42 / fig.get_figheight())
        ax.set_xlim(*sinir)
        ax.set_ylim(n - 0.5, -0.5)  # ilk kategori en üstte
        ax.set_yticks(range(n), etiketler)
        ax.xaxis.set_major_formatter(eksen_bicimi)
        ax.grid(axis="x")
        ax.spines["bottom"].set_visible(False)
    else:
        fig, ax = plt.subplots(figsize=(7.2, 4.0))
        # referans çizgisinin etiketi sağ kenarın dışına yazılır
        fig.subplots_adjust(left=0.09, right=0.96 if referans is None else 0.85, top=0.76, bottom=0.17)
        ax.set_ylim(*sinir)
        ax.set_xlim(-0.6, n - 0.4)
        ax.set_xticks(range(n), etiketler)
        ax.yaxis.set_major_formatter(eksen_bicimi)
        ax.grid(axis="y")
        ax.spines["left"].set_visible(False)

    fig.canvas.draw()  # piksel ölçülerinin kesinleşmesi için
    for i, deger in enumerate(degerler):
        _cubuk_ciz(ax, i, deger, yatay, SERI[0])
        if yatay:
            ax.annotate(bicim(deger), (deger, i), xytext=(6 if deger >= 0 else -6, 0), textcoords="offset points",
                        va="center", ha="left" if deger >= 0 else "right")
        else:
            ax.annotate(bicim(deger), (i, deger), xytext=(0, 5 if deger >= 0 else -5), textcoords="offset points",
                        ha="center", va="bottom" if deger >= 0 else "top")

    if referans is not None:
        if yatay:
            ax.axvline(referans, color=SOLUK, linewidth=1)
            ax.annotate(referans_etiketi, (referans, 1), xycoords=("data", "axes fraction"), xytext=(0, 4),
                        textcoords="offset points", ha="center", va="bottom", fontsize=9, color=IKINCIL)
        else:
            ax.axhline(referans, color=SOLUK, linewidth=1)
            ax.annotate(referans_etiketi.replace(" %", "\n%"), (1, referans), xycoords=("axes fraction", "data"),
                        xytext=(6, 0), textcoords="offset points", ha="left", va="center", fontsize=9, color=IKINCIL)
    if en_kucuk < 0:  # sıfır çizgisi
        (ax.axvline if yatay else ax.axhline)(0, color=EKSEN, linewidth=1)

    basliklar(fig, baslik, alt_baslik)
    kaydet(fig, dosya_adi)


def cizgi_grafigi(x, seriler, baslik, alt_baslik, dosya_adi, x_etiketi, x_bicimi, y_bicimi, vurgu=None):
    """Birkaç serili çizgi grafik. `seriler`: {ad: değerler}; `vurgu` serisinin son değeri yazılır."""
    fig, ax = plt.subplots(figsize=(7.6, 4.6))
    fig.subplots_adjust(left=0.11, right=0.88, top=0.74, bottom=0.13)
    for (ad, degerler), renk in zip(seriler.items(), SERI):
        ax.plot(x, degerler, color=renk, marker="o", markersize=6, markeredgecolor=YUZEY, markeredgewidth=1.5, label=ad)
        if ad == vurgu:
            ax.annotate(y_bicimi(degerler[-1], None), (x[-1], degerler[-1]), xytext=(8, 0), textcoords="offset points",
                        va="center", fontsize=9)
    if min(min(d) for d in seriler.values()) < 0:
        ax.axhline(0, color=EKSEN, linewidth=1)
    else:
        ax.set_ylim(bottom=0)
    ax.set_xlabel(x_etiketi)
    ax.set_xticks(x)
    ax.xaxis.set_major_formatter(x_bicimi)
    ax.yaxis.set_major_formatter(y_bicimi)
    ax.grid(True)
    ax.legend(loc="lower left", bbox_to_anchor=(0, 1.0), ncols=len(seriler), fontsize=9, borderaxespad=0.3,
              handlelength=1.6, columnspacing=1.4)
    basliklar(fig, baslik, alt_baslik)
    kaydet(fig, dosya_adi)
