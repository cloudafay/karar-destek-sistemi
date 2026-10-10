"""Betiklerin ortak kullandığı küçük yardımcılar."""
import sys

import ayarlar as A

# Türkçe karakterlerin her konsolda doğru yazılması için
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


def tr(x, ondalik=0):
    """1234567.891 -> '1.234.567,89' (Türkçe sayı biçimi)."""
    metin = f"{x:,.{ondalik}f}"
    return metin.replace(",", "#").replace(".", ",").replace("#", ".").replace("-", "−")  # gerçek eksi işareti


def yuzde(x, ondalik=1):
    """0.265 -> '%26,5'."""
    return "%" + tr(100 * x, ondalik)


def para(x, ondalik=0):
    return f"{tr(x, ondalik)} {A.PARA}"


def baslik_yaz(metin):
    print(f"\n=== {metin} ===")
