"""Tüm aşamaları sırayla çalıştırır; ham veriden bütün çıktıları yeniden üretir.

Çalıştırma:  py src/hepsini_calistir.py
"""
import kesifsel_analiz
import model
import optimizasyon
import senaryo_karsilastirma
import teklif_senaryosu
import veri_hazirlama

for asama in (veri_hazirlama, kesifsel_analiz, model, teklif_senaryosu, optimizasyon, senaryo_karsilastirma):
    asama.main()
print("\nTamamlandı. Yönetici ekranı için:  py -m streamlit run app.py")
