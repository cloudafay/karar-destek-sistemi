"""Aşama 10: yönetici karar ekranı.

İlk sekme kararı vermek için gereken her şeyi düz bir dille gösterir: öneri, bütçenin nereye
gittiği, bugünkü yöntemle karşılaştırma, belirsizlik ve karar formu. Müşteri listesi ile teknik
ayrıntılar ayrı sekmelerdedir.

Çalıştırma:  py -m streamlit run app.py
Önce bir kez `py src/hepsini_calistir.py` çalıştırılmış olmalıdır.
"""
import json
import sys
from datetime import datetime
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st

sys.path.insert(0, str(Path(__file__).parent / "src"))
import ayarlar as A
from optimizasyon import butce_ozeti, kampanya_listesi, optimize_et
from ortak import para, tr, yuzde
from senaryo_karsilastirma import duyarlilik_tablosu, senaryo_tablosu
from teklif_senaryosu import teklif_tablosu
from yonetici_ozeti import ozet_metni, uyarilar

st.set_page_config(page_title="Kampanya karar ekranı", layout="wide")

# --- Görünüm ----------------------------------------------------------------------
# Renkler renk körlüğü denetiminden geçmiş paletten; açık ve koyu tema için ayrı adımlar.
ACIK = dict(yuzey="#ffffff", murekkep="#0b0b0b", ikincil="#52514e", cizgi="rgba(11,11,11,0.10)", vurgu="#2a78d6",
            iz="#cde2fb", soluk="#a3a199", iyi="#006300", kotu="#d03b3b",
            seri=["#2a78d6", "#eb6834", "#1baf7a"], risk=["#104281", "#2a78d6", "#86b6ef"])
KOYU = dict(yuzey="#1a1a19", murekkep="#f0efec", ikincil="#c3c2b7", cizgi="rgba(255,255,255,0.10)", vurgu="#3987e5",
            iz="#184f95", soluk="#6b6a65", iyi="#0ca30c", kotu="#e66767",
            seri=["#3987e5", "#d95926", "#199e70"], risk=["#9ec5f4", "#3987e5", "#184f95"])
try:
    T = KOYU if st.context.theme.type == "dark" else ACIK
except Exception:
    T = ACIK

st.html(f"""<style>
.kds-izgara {{display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 14px; margin-bottom: 6px;}}
.kds-kart {{background: {T['yuzey']}; border: 1px solid {T['cizgi']}; border-radius: 12px; padding: 16px 18px; color: {T['murekkep']};}}
.kds-etiket {{font-size: 0.88rem; color: {T['ikincil']}; margin-bottom: 4px;}}
.kds-deger {{font-size: 1.9rem; font-weight: 600; line-height: 1.2;}}
.kds-not {{font-size: 0.88rem; color: {T['ikincil']}; margin-top: 6px; line-height: 1.45;}}
.kds-ana {{display: flex; flex-wrap: wrap; gap: 28px; align-items: center; justify-content: space-between; padding: 22px 26px; margin-bottom: 14px;}}
.kds-ana .kds-sol {{flex: 1 1 420px;}}
.kds-ana .kds-deger {{font-size: 3.2rem;}}
.kds-oneri {{font-size: 1.45rem; line-height: 1.4; font-weight: 500; margin: 0;}}
.kds-secili {{border: 2px solid {T['vurgu']};}}
.kds-iyi {{color: {T['iyi']}; font-weight: 600;}}
.kds-kotu {{color: {T['kotu']}; font-weight: 600;}}
.kds-olcer {{height: 8px; border-radius: 4px; background: {T['iz']}; margin-top: 10px; overflow: hidden;}}
.kds-olcer div {{height: 100%; background: {T['vurgu']}; border-radius: 4px;}}
.kds-serit {{display: flex; gap: 2px; height: 16px; border-radius: 8px; overflow: hidden; margin: 6px 0 10px 0;}}
.kds-gosterge {{display: flex; flex-wrap: wrap; gap: 6px 22px; font-size: 0.92rem; color: {T['murekkep']};}}
.kds-gosterge i {{display: inline-block; width: 11px; height: 11px; border-radius: 3px; margin-right: 7px;}}
</style>""")


def kart(etiket, deger, notu="", sinif=""):
    return (f'<div class="kds-kart {sinif}"><div class="kds-etiket">{etiket}</div>'
            f'<div class="kds-deger">{deger}</div><div class="kds-not">{notu}</div></div>')


def izgara(*kartlar):
    st.html('<div class="kds-izgara">' + "".join(kartlar) + "</div>")


def yatay_cubuk(veri, kategori, deger, etiket, vurgu=None):
    """Değerleri çubuk ucunda yazan yatay çubuk grafik. `vurgu` verilirse yalnızca o satır renklidir."""
    en_buyuk = max(veri[deger].max(), 1)
    taban = alt.Chart(veri).encode(
        y=alt.Y(f"{kategori}:N", sort=None, title=None,
                axis=alt.Axis(labelLimit=240, labelFontSize=13, labelColor=T["murekkep"], ticks=False, domain=False)),
        x=alt.X(f"{deger}:Q", axis=None, scale=alt.Scale(domain=[0, en_buyuk * 1.55])),
        tooltip=[alt.Tooltip(f"{kategori}:N", title=" "), alt.Tooltip(f"{etiket}:N", title="Değer")],
    )
    renk = alt.value(T["vurgu"])
    if vurgu:
        renk = alt.condition(alt.datum[kategori] == vurgu, alt.value(T["vurgu"]), alt.value(T["soluk"]))
    cubuk = taban.mark_bar(cornerRadiusEnd=4, size=22).encode(color=renk)
    yazi = taban.mark_text(align="left", dx=7, fontSize=13, color=T["murekkep"]).encode(text=f"{etiket}:N")
    return (cubuk + yazi).properties(height=34 + 46 * len(veri)).configure_view(strokeWidth=0)


# --- Veri ve hesap ------------------------------------------------------------------
@st.cache_data
def verileri_oku():
    return pd.read_csv(A.RISK_TABLOSU), json.loads(A.MODEL_METRIKLERI.read_text(encoding="utf-8"))


@st.cache_data
def hesapla(butce, destek_kapasitesi, katki_marji):
    risk, _ = verileri_oku()
    teklifler = teklif_tablosu(risk, katki_marji=katki_marji)
    liste = kampanya_listesi(risk, teklifler, optimize_et(teklifler, butce, destek_kapasitesi))
    return teklifler, liste, butce_ozeti(liste, butce), senaryo_tablosu(risk, teklifler, butce, destek_kapasitesi)


if not A.RISK_TABLOSU.exists() or not A.MODEL_METRIKLERI.exists():
    st.error("Model çıktıları bulunamadı. Önce terminalde `py src/hepsini_calistir.py` komutunu çalıştırın.")
    st.stop()

risk, metrikler = verileri_oku()
esik = metrikler["secilen_esik"]

# --- Kenar çubuğu: "ya şöyle olsaydı?" -------------------------------------------
BASLANGIC = {"butce": A.BUTCE, "kapasite": A.DESTEK_KAPASITESI, "marj": A.KATKI_MARJI}
for anahtar, deger in BASLANGIC.items():
    st.session_state.setdefault(anahtar, deger)

with st.sidebar:
    st.header("Ya şöyle olsaydı?")
    st.caption("Değerleri değiştirin; plan ve ekrandaki bütün sayılar yeniden hesaplanır.")
    butce = st.slider(f"Kampanya bütçesi ({A.PARA})", 0, 30_000, step=500, key="butce")
    destek_kapasitesi = st.slider("Teknik destek ekibinin bakabileceği müşteri", 0, 400, step=10, key="kapasite")
    with st.expander("Gelişmiş varsayım"):
        katki_marji = st.slider("Katkı marjı (aylık ücretin kâra dönen payı)", 0.20, 0.80, step=0.05, key="marj")
    st.button("Başlangıç değerlerine dön", on_click=lambda: st.session_state.update(BASLANGIC))
    st.caption("Teklif maliyetleri ve etkileri sentetik varsayımdır; gerçek kampanya sonucu değildir. "
               "Tüm varsayımlar `src/ayarlar.py` dosyasındadır.")

teklifler, liste, ozet, senaryolar = hesapla(butce, destek_kapasitesi, katki_marji)
onerilen = liste[liste["offer_type"] != A.TEKLIF_YOK]
bugunku, plan = senaryolar.loc[A.SENARYO_A], senaryolar.loc[A.SENARYO_C]
yuksek_riskli = risk[risk["risk_grubu"] == "Yüksek"]

# --- Başlık -------------------------------------------------------------------------
st.title("Kampanya bütçesi karar ekranı")
veri_tarihi = datetime.fromtimestamp(A.HAM_VERI.stat().st_mtime).strftime("%d.%m.%Y")
st.caption(f"{tr(len(risk))} müşterilik portföy · veri tarihi {veri_tarihi} · tutarlar {A.PARA} · "
           "teklif maliyetleri ve etkileri sentetik varsayımdır")

ozet_sekmesi, liste_sekmesi, teknik_sekme = st.tabs(["Özet ve karar", "Müşteri listesi", "Yöntem ve model"])

# ======================= 1. sekme: özet ve karar ====================================
with ozet_sekmesi:
    # --- Öneri
    fark = plan["beklenen_net_fayda"] - bugunku["beklenen_net_fayda"]
    if ozet["teklif_sunulan_musteri"]:
        oneri = (f"{para(butce)} bütçeyle <b>{tr(ozet['teklif_sunulan_musteri'])} müşteriye</b> teklif sunulması öneriliyor. "
                 f"Diğer {tr(len(liste) - ozet['teklif_sunulan_musteri'])} müşteriye teklif önerilmiyor.")
    else:
        oneri = "Bu bütçeyle önerilebilecek bir teklif yok."
    st.html(f"""<div class="kds-kart kds-ana">
      <div class="kds-sol">
        <div class="kds-etiket">Öneri</div>
        <p class="kds-oneri">{oneri}</p>
        <div class="kds-not">Aynı bütçeyle bugünkü yöntemin (herkese aynı indirim) beklenen net faydası {para(bugunku['beklenen_net_fayda'])}.</div>
      </div>
      <div>
        <div class="kds-etiket">Beklenen net fayda</div>
        <div class="kds-deger">{para(ozet['beklenen_net_fayda'])}</div>
        <div class="kds-not"><span class="kds-iyi">▲ {para(fark)}</span> bugünkü yönteme göre<br>
        Elde tutulan müşterilerden beklenen ek kâr − teklif maliyeti</div>
      </div>
    </div>""")

    izgara(
        kart("Ayrılma riski yüksek müşteri", tr(len(yuksek_riskli)),
             f"İncelenen {tr(len(risk))} müşterinin {yuzde(len(yuksek_riskli) / len(risk), 0)} kadarı"),
        kart("Önerilen harcama", para(ozet["harcama"]),
             f"{para(butce)} bütçenin {yuzde(ozet['kullanim_orani'], 0)} kadarı · kalan {para(ozet['kalan'])}"
             f'<div class="kds-olcer"><div style="width:{min(ozet["kullanim_orani"], 1) * 100:.1f}%"></div></div>'),
        kart("Bütçe yetmediği için dışarıda kalan", f"{tr(ozet['butce_disinda_kalan'])} müşteri",
             "Teklif faydalı olurdu, ancak bütçe ya da destek kapasitesi yetmedi"),
    )

    # --- Bütçe nereye gidiyor? / karşılaştırma
    sol, sag = st.columns(2, gap="large")
    with sol:
        st.subheader("Bütçe nereye gidiyor?")
        dagilim = onerilen.groupby("offer_type").agg(
            musteri=("customerID", "size"), harcama=("offer_cost", "sum"), net_fayda=("beklenen_net_fayda", "sum"),
        ).reindex(A.TEKLIFLER, fill_value=0).rename_axis("teklif").reset_index().sort_values("harcama", ascending=False)
        dagilim["etiket"] = [f"{para(h)} · {tr(m)} müşteri" for h, m in zip(dagilim["harcama"], dagilim["musteri"])]
        st.altair_chart(yatay_cubuk(dagilim, "teklif", "harcama", "etiket"))
        st.table(pd.DataFrame({
            "Teklif": dagilim["teklif"],
            "Ne sunuluyor?": dagilim["teklif"].map({
                A.DESTEK: f"İnterneti olup teknik destek almayan müşteriye ücretsiz destek ({A.DESTEK_MALIYETI} {A.PARA})",
                A.INDIRIM: f"{A.INDIRIM_AY} ay boyunca aylık ücrette {yuzde(A.INDIRIM_ORANI, 0)} indirim",
                A.PAKET: f"Aylık ücreti {A.PAKET_ALT_SINIR} {A.PARA} ve üzeri olanı daha uygun pakete geçirme",
            }),
            "Beklenen net fayda": dagilim["net_fayda"].map(para),
        }).set_index("Teklif"))
    with sag:
        st.subheader("Bugünkü yöntemle karşılaştırma")
        adlar = {A.SENARYO_A: "Bugünkü yöntem", A.SENARYO_B: "Yalnızca risk sıralaması", A.SENARYO_C: "Önerilen plan"}
        kiyas = pd.DataFrame({"yaklasim": [adlar[ad] for ad in senaryolar.index],
                              "net_fayda": senaryolar["beklenen_net_fayda"].values})
        kiyas["etiket"] = kiyas["net_fayda"].map(para)
        st.altair_chart(yatay_cubuk(kiyas, "yaklasim", "net_fayda", "etiket", vurgu="Önerilen plan"))
        st.table(pd.DataFrame({
            "Yaklaşım": kiyas["yaklasim"],
            "Nasıl çalışır?": ["Model yok; aylık sözleşmeli yeni müşterilere aynı indirim",
                               "En riskli müşterilere aynı indirim",
                               "Her müşteriye kendisi için en uygun teklif; faydası maliyetini karşılamıyorsa teklif yok"],
            "Teklif sunulan müşteri": senaryolar["teklif_sunulan_musteri"].map(tr).values,
        }).set_index("Yaklaşım"))
    st.caption("Üç yaklaşım da aynı bütçeyi ve aynı varsayımları kullanır. Çubuklar beklenen net faydayı gösterir.")

    # --- Ne kadar emin olabiliriz?
    st.subheader("Ne kadar emin olabiliriz?")
    st.markdown("Tekliflerin müşteriyi ne ölçüde elde tuttuğu ölçülmedi; bu bir varsayım. "
                "Aşağıda aynı planın, teklifler beklenenden az ya da çok işe yararsa getireceği sonuç var.")

    def sonuc_karti(ad, aciklama, carpan, sinif=""):
        net = carpan * plan["beklenen_ek_katki"] - plan["maliyet"]
        durum = '<span class="kds-kotu">▼ Plan zarar eder</span>' if net < 0 else '<span class="kds-iyi">▲ Plan kendini karşılar</span>'
        return kart(ad, para(net), f"{aciklama}<br>{durum}", sinif)

    izgara(
        sonuc_karti("Kötümser", "Teklifler beklenenin yarısı kadar işe yararsa", 0.5),
        sonuc_karti("Beklenen", "Teklifler varsayıldığı gibi işe yararsa", 1.0, "kds-secili"),
        sonuc_karti("İyimser", "Teklifler beklenenden yarı yarıya daha iyi işe yararsa", 1.5),
    )
    tum_uyarilar = uyarilar(risk, ozet)
    basabas = (f"Plan, tekliflerin etkisi varsayılanın {yuzde(plan['basabas_etki_carpani'], 0)} düzeyine inene kadar maliyetini karşılar. "
               if ozet["teklif_sunulan_musteri"] else "")
    st.warning(f"{basabas}Bu nedenle planın önce küçük bir pilotla ve kontrol grubuyla denenmesi önerilir.\n\n"
               + "\n".join(f"- {uyari}" for uyari in tum_uyarilar[:1] + tum_uyarilar[2:4]))
    with st.expander("Diğer uyarılar ve yazılı özet"):
        st.markdown("\n".join(f"- {uyari}" for uyari in tum_uyarilar[1:2] + tum_uyarilar[4:]))
        st.markdown(ozet_metni(risk, ozet, senaryolar, esik).replace("$", "\\$"))

    # --- Kim risk altında?
    st.subheader("Kim risk altında?")
    sayilar = risk["risk_grubu"].value_counts().reindex(["Yüksek", "Orta", "Düşük"], fill_value=0)
    serit = "".join(f'<div style="flex:{n}; background:{renk}"></div>' for n, renk in zip(sayilar, T["risk"]))
    gosterge = "".join(f'<span><i style="background:{renk}"></i>{grup} risk: {tr(n)} müşteri ({yuzde(n / len(risk), 0)})</span>'
                       for (grup, n), renk in zip(sayilar.items(), T["risk"]))
    st.html(f'<div class="kds-serit">{serit}</div><div class="kds-gosterge">{gosterge}</div>')
    isaretler = yuksek_riskli["risk_isaretleri"].str.split("; ").explode().value_counts().head(3)
    st.markdown("Yüksek riskli müşterilerde en sık öne çıkan işaretler: "
                + "; ".join(f"**{isaret}** ({yuzde(n / len(yuksek_riskli), 0)})" for isaret, n in isaretler.items())
                + ". Bunlar ilişki gösterir, ayrılmanın nedenini kanıtlamaz.")

    # --- Karar
    st.subheader("Kararınız")
    st.caption("Sistem öneri sunar; karar sizindir. Kararınız, gerekçeniz ve o andaki bütçe ile varsayımlar kayıt dosyasına eklenir.")
    with st.form("karar_formu", clear_on_submit=True):
        st.markdown(f"Karara konu plan: **{tr(ozet['teklif_sunulan_musteri'])} müşteri**, **{para(ozet['harcama']).replace('$', chr(92) + '$')}** harcama.")
        karar = st.radio("Karar", ["Onayla", "Değiştir", "Reddet"], horizontal=True)
        karar_veren = st.text_input("Karar veren")
        gerekce = st.text_area("Gerekçe")
        gonderildi = st.form_submit_button("Kararı kaydet", type="primary")
    if gonderildi:
        if not gerekce.strip():
            st.error("Karar kaydedilmedi: gerekçe boş bırakılamaz.")
        else:
            kayit = pd.DataFrame([{
                "zaman": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "karar_veren": karar_veren.strip(), "karar": karar,
                "gerekce": gerekce.strip(), "butce": butce, "destek_kapasitesi": destek_kapasitesi, "katki_marji": katki_marji,
                "teklif_sunulan_musteri": ozet["teklif_sunulan_musteri"], "harcama": round(ozet["harcama"], 2),
                "beklenen_net_fayda": round(ozet["beklenen_net_fayda"], 2),
            }])
            if A.KARAR_KAYDI.exists():
                kayit = pd.concat([pd.read_csv(A.KARAR_KAYDI), kayit], ignore_index=True)
            kayit.to_csv(A.KARAR_KAYDI, index=False, encoding="utf-8-sig")
            st.success(f"Karar kaydedildi: {karar}.")
    if A.KARAR_KAYDI.exists():
        with st.expander("Önceki kararlar"):
            st.dataframe(pd.read_csv(A.KARAR_KAYDI).iloc[::-1], hide_index=True, width="stretch")

# ======================= 2. sekme: müşteri listesi ==================================
with liste_sekmesi:
    f1, f2, f3 = st.columns([2, 3, 2])
    risk_filtresi = f1.multiselect("Risk grubu", ["Yüksek", "Orta", "Düşük"], placeholder="Tümü")
    teklif_filtresi = f2.multiselect("Önerilen teklif", A.TEKLIFLER + [A.TEKLIF_YOK], placeholder="Tümü")
    yalniz_onerilen = f3.toggle("Yalnızca teklif önerilenler", value=True)

    tablo = liste
    if yalniz_onerilen:
        tablo = tablo[tablo["offer_type"] != A.TEKLIF_YOK]
    if risk_filtresi:
        tablo = tablo[tablo["risk_grubu"].isin(risk_filtresi)]
    if teklif_filtresi:
        tablo = tablo[tablo["offer_type"].isin(teklif_filtresi)]

    st.caption(f"{tr(len(tablo))} müşteri listeleniyor; en yüksek beklenen net faydadan başlar. Risk işaretleri, müşterinin içinde "
               "bulunduğu ve kayıp oranı yüksek olan grupları gösterir; ayrılmanın nedenini kanıtlamaz.")
    para_sutunu = dict(format=f"%.0f {A.PARA}")
    gorunen = tablo.assign(churn_olasiligi=tablo["churn_olasiligi"] * 100)  # yüzde olarak gösterilir
    st.dataframe(
        gorunen[["customerID", "churn_olasiligi", "risk_grubu", "risk_isaretleri", "offer_type", "offer_cost",
                 "beklenen_net_fayda", "dayaniklilik", "aciklama"]],
        hide_index=True, width="stretch", height=430,
        column_config={
            "customerID": "Müşteri kodu",
            "churn_olasiligi": st.column_config.ProgressColumn("Ayrılma olasılığı", min_value=0, max_value=100, format="%%%.0f"),
            "risk_grubu": "Risk grubu",
            "risk_isaretleri": st.column_config.TextColumn("Risk işaretleri", width="large"),
            "offer_type": "Önerilen teklif",
            "offer_cost": st.column_config.NumberColumn("Maliyet", **para_sutunu),
            "beklenen_net_fayda": st.column_config.NumberColumn("Beklenen net fayda", **para_sutunu),
            "dayaniklilik": st.column_config.TextColumn(
                "Belirsizlik", help="Varsayıma duyarlı: teklif beklenenin yarısı kadar işe yararsa maliyetini karşılamaz"),
            "aciklama": st.column_config.TextColumn("Teklif sunulmama nedeni", width="medium"),
        })
    st.download_button("Listeyi CSV olarak indir", tablo.to_csv(index=False).encode("utf-8-sig"),
                       file_name="kampanya_listesi.csv", mime="text/csv")

    st.subheader("Tek bir müşteri için seçenekler")
    kimlik = st.selectbox("Müşteri kodu", liste["customerID"], help="Liste beklenen net faydaya göre sıralıdır")
    musteri = liste.set_index("customerID").loc[kimlik]
    st.markdown(f"Ayrılma olasılığı **{yuzde(musteri['churn_olasiligi'])}** ({musteri['risk_grubu']} risk). "
                f"Risk işaretleri: {musteri['risk_isaretleri']}.")
    yaz = lambda deger, bicim: "-" if pd.isna(deger) else bicim(deger)
    secenekler = teklifler[teklifler["customerID"] == kimlik].set_index("offer_type").reindex(A.TEKLIFLER + [A.TEKLIF_YOK])
    st.dataframe(pd.DataFrame({
        "Teklif": secenekler.index,
        "Durum": ["Uygun değil" if pd.isna(m) else ("Önerilen" if t == musteri["offer_type"] else "")
                  for t, m in secenekler["offer_cost"].items()],
        "Maliyet": [yaz(m, para) for m in secenekler["offer_cost"]],
        "Beklenen net fayda": [yaz(f, para) for f in secenekler["beklenen_net_fayda"]],
        "Ek elde tutma olasılığı": [yaz(o, yuzde) for o in secenekler["estimated_incremental_retention"]],
        "Beklenen ek katkı": [yaz(k, para) for k in secenekler["beklenen_ek_katki"]],
    }), hide_index=True, width="stretch")
    if musteri["offer_type"] == A.TEKLIF_YOK:
        st.caption(musteri["aciklama"] + ".")
    st.caption(f"Kanıt düzeyi: {A.KANIT_DUZEYI}. Elde tutulan müşterinin {A.UFUK_AY} aylık katkısı aylık ücret × katkı marjı ile hesaplanır.")

# ======================= 3. sekme: yöntem ve model ==================================
with teknik_sekme:
    st.caption("Bu sekme analistler içindir; karar için gereken bilgi ilk sekmededir.")
    st.subheader("Tahmin modeli")
    st.markdown(f"Temel model: {metrikler['temel_model']}. Eğitim {tr(metrikler['veri']['egitim'])}, test {tr(metrikler['veri']['test'])} müşteri. "
                f"Eşik ({tr(esik, 2)}), eğitim verisindeki çapraz doğrulama tahminlerinde F1 skorunu en yüksek yapan değerdir.")
    satirlar = []
    for ad, t in metrikler["test"].items():
        secilen = t["secilen_esik"]
        satirlar.append({"Model": ad, "ROC-AUC (çapraz doğrulama)": tr(metrikler["capraz_dogrulama"][ad]["roc_auc"], 3),
                         "ROC-AUC (test)": tr(t["roc_auc"], 3), "PR-AUC (test)": tr(t["pr_auc"], 3), "Brier (test)": tr(t["brier"], 3),
                         "Eşik": tr(secilen["esik"], 2), "Kesinlik": yuzde(secilen["kesinlik"]),
                         "Duyarlılık": yuzde(secilen["duyarlilik"]), "F1": tr(secilen["f1"], 3)})
    st.dataframe(pd.DataFrame(satirlar), hide_index=True, width="stretch")
    g1, g2 = st.columns(2)
    for sutun, dosya in ((g1, "06_roc_pr_egrileri.png"), (g2, "07_kalibrasyon.png")):
        if (A.GRAFIK / dosya).exists():
            sutun.image(str(A.GRAFIK / dosya))

    st.subheader("Senaryoların teknik karşılaştırması")
    st.dataframe(pd.DataFrame({
        "Senaryo": senaryolar.index,
        "Teklif sunulan müşteri": senaryolar["teklif_sunulan_musteri"].map(tr),
        "Maliyet": senaryolar["maliyet"].map(para),
        "Beklenen net fayda": senaryolar["beklenen_net_fayda"].map(para),
        "Etiket bazlı net fayda": senaryolar["etiket_bazli_net_fayda"].map(para),
        "Gerçekte ayrılan hedef": senaryolar["gercekte_ayrilan_hedef"].map(tr),
        "İsabet oranı": senaryolar["isabet_orani"].map(yuzde),
        "Başabaş etki çarpanı": senaryolar["basabas_etki_carpani"].map(lambda c: f"×{tr(c, 2)}"),
    }), hide_index=True, width="stretch")
    st.caption("Etiket bazlı net fayda, test müşterilerinin gerçekte ayrılıp ayrılmadığı bilgisiyle hesaplanır; kampanya etkisi yine "
               "varsayımdır. Başabaş etki çarpanı: gerçek etki varsayılanın bu katına inene kadar plan zarar etmez.")
    duyarlilik = duyarlilik_tablosu(senaryolar).reset_index().melt("etki_carpani", var_name="senaryo", value_name="net_fayda")
    duyarlilik["Net fayda"] = duyarlilik["net_fayda"].map(para)
    duyarlilik["Gerçekleşen / varsayılan etki"] = duyarlilik["etki_carpani"].map(lambda c: f"×{tr(c, 2)}")
    st.altair_chart(alt.Chart(duyarlilik).mark_line(strokeWidth=2, point=alt.OverlayMarkDef(size=60)).encode(
        x=alt.X("etki_carpani:Q", title="Gerçekleşen etki / varsayılan etki",
                axis=alt.Axis(values=sorted(duyarlilik["etki_carpani"].unique()),
                              labelExpr="'×' + replace(format(datum.value, '.2f'), '.', ',')")),
        y=alt.Y("net_fayda:Q", title=f"Net fayda ({A.PARA})",
                axis=alt.Axis(labelExpr="replace(format(datum.value, ',.0f'), /,/g, '.')")),
        color=alt.Color("senaryo:N", sort=None, title=None, scale=alt.Scale(range=T["seri"]), legend=alt.Legend(orient="top")),
        tooltip=["senaryo:N", "Gerçekleşen / varsayılan etki:N", "Net fayda:N"],
    ).properties(height=340))

    grup_dosyasi = A.TABLO / "grup_bazli_kontrol.csv"
    if grup_dosyasi.exists():
        st.subheader("Grup bazlı kontrol")
        st.caption("Cinsiyet ve yaş grubu model girdisi değildir; burada yalnızca sonuçların gruplar arasında nasıl dağıldığı denetlenir.")
        gruplar = pd.read_csv(grup_dosyasi, dtype={"grup": str})
        grup_adi = {"Female": "Kadın", "Male": "Erkek", "0": "65 yaş altı", "1": "65 yaş ve üzeri"}
        st.dataframe(pd.DataFrame({
            "Değişken": gruplar["grup_degiskeni"].map({"gender": "Cinsiyet", "SeniorCitizen": "Yaş grubu"}),
            "Grup": gruplar["grup"].map(grup_adi),
            "Müşteri": gruplar["musteri_sayisi"].map(tr),
            "Gerçek kayıp oranı": gruplar["gercek_kayip_orani"].map(yuzde),
            "Ortalama tahmin": gruplar["ortalama_tahmin"].map(yuzde),
            "Yüksek risk payı": gruplar["yuksek_risk_payi"].map(yuzde),
            "Duyarlılık": gruplar["duyarlilik"].map(yuzde),
            "Kesinlik": gruplar["kesinlik"].map(yuzde),
        }), hide_index=True, width="stretch")
