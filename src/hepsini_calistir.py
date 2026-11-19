"""Tüm aşamaları sırayla çalıştırır; ham veriden bütün çıktıları yeniden üretir.

Çalıştırma:  py src/hepsini_calistir.py
Raporu yeniden yazmak için ayrıca:  py src/rapor_olustur.py
"""
import kesifsel_analiz
import model
import optimizasyon
import senaryo_karsilastirma
import shap_analizi
import tekrar_kayit_etkisi
import teklif_senaryosu
import veri_hazirlama

ASAMALAR = (veri_hazirlama, kesifsel_analiz, model, tekrar_kayit_etkisi, shap_analizi,
            teklif_senaryosu, optimizasyon, senaryo_karsilastirma)

if __name__ == "__main__":
    for asama in ASAMALAR:
        asama.main()
    print("\nTamamlandı. Yönetici ekranı için:  py -m streamlit run app.py")
