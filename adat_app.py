import streamlit as st
import pandas as pd
import numpy as np
import json
import os
from io import BytesIO

# Sayfa ayarı
st.set_page_config(page_title="Ortaklar Cari Hesap Adat Hesaplama", layout="wide")
st.title("📊 Ortaklar Cari Hesap Adat Hesaplama Sistemi")
st.markdown("---")

# --- Dosya yükleme ---
muavin_file = st.file_uploader("Muavin Excel Dosyasını Yükle (.xlsx)", type=["xlsx"])

# Dönem tarihleri
donem_baslangic = st.date_input("Dönem Başlangıç Tarihi", value=pd.to_datetime("2025-01-01"))
donem_bitis = st.date_input("Dönem Bitiş Tarihi", value=pd.to_datetime("2025-09-30"))

# Aylık faiz oranları bölümü
st.markdown("### 📅 Aylık Adat Faiz Oranları (%)")
cols = st.columns(6)
faiz_oranlari = {}

# JSON dosyasını kontrol et / yükle
faiz_oranlari_path = "faiz_oranlari.json"
if os.path.exists(faiz_oranlari_path):
    with open(faiz_oranlari_path, "r") as f:
        saved_rates = json.load(f)
else:
    saved_rates = {}

# Aylık oranlar için giriş alanları
aylar = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]
for i, ay in enumerate(aylar):
    default_value = saved_rates.get(str(i + 1), 44.25 if i >= 2 else 49.25)
    faiz_oranlari[i + 1] = cols[i % 6].number_input(f"{ay}", min_value=0.0, value=default_value, step=0.01)

# Kaydet butonu
st.markdown("---")
if st.button("💾 Faiz Oranlarını Kaydet"):
    with open(faiz_oranlari_path, "w") as f:
        json.dump(faiz_oranlari, f)
    st.success("Faiz oranları başarıyla kaydedildi!")

st.markdown("---")

# --- Hesaplama bölümü ---
if muavin_file:
    df = pd.read_excel(muavin_file)
    df.columns = ["Tarih", "Borç", "Alacak"]
    df["Tarih"] = pd.to_datetime(df["Tarih"], dayfirst=True)

    df = df.sort_values("Tarih").reset_index(drop=True)
    df["Sonraki_Tarih"] = pd.to_datetime(donem_bitis)
    df["Gün_Sayısı"] = (df["Sonraki_Tarih"] - df["Tarih"]).dt.days

    # Adat hesaplama
    df["Borç_Adat"] = df["Borç"] * df["Gün_Sayısı"]
    df["Alacak_Adat"] = df["Alacak"] * df["Gün_Sayısı"]
    df["Ay"] = df["Tarih"].dt.month
    df["Faiz_Oranı"] = df["Ay"].map(faiz_oranlari)
    df["Borç_Faiz"] = df["Borç_Adat"] * df["Faiz_Oranı"] / (365 * 100)
    df["Alacak_Faiz"] = df["Alacak_Adat"] * df["Faiz_Oranı"] / (365 * 100)

    # Sıfırları gizle
    df.replace(0, np.nan, inplace=True)

    # Para biçimi
    def fmt(x):
        if pd.isna(x):
            return ""
        return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    for col in ["Borç", "Alacak", "Borç_Faiz", "Alacak_Faiz"]:
        df[col] = df[col].apply(fmt)

    # Toplamlar
    borc_toplam = df["Borç"].replace("", 0).astype(str).str.replace(".", "", regex=False).str.replace(",", ".").astype(float).sum()
    alacak_toplam = df["Alacak"].replace("", 0).astype(str).str.replace(".", "", regex=False).str.replace(",", ".").astype(float).sum()
    borc_faiz_toplam = df["Borç_Faiz"].replace("", 0).astype(str).str.replace(".", "", regex=False).str.replace(",", ".").astype(float).sum()
    alacak_faiz_toplam = df["Alacak_Faiz"].replace("", 0).astype(str).str.replace(".", "", regex=False).str.replace(",", ".").astype(float).sum()
    net_adat = borc_faiz_toplam - alacak_faiz_toplam

    # Görüntüleme
    st.markdown("### 📄 Ayrıntılı Hesaplama Tablosu")
    st.dataframe(df, use_container_width=True)

    st.markdown("### 📘 Dönem Özeti")
    ozet = pd.DataFrame({
        "Borç Toplamı": [fmt(borc_toplam)],
        "Alacak Toplamı": [fmt(alacak_toplam)],
        "Borç Faiz Toplamı": [fmt(borc_faiz_toplam)],
        "Alacak Faiz Toplamı": [fmt(alacak_faiz_toplam)],
        "Net Adat Tutarı (Borç - Alacak)": [fmt(net_adat)]
    })
    st.table(ozet)

    # Excel çıktısı
    # Tarih sütunundaki saatleri kaldır
    df["Tarih"] = df["Tarih"].dt.date
    df["Sonraki_Tarih"] = pd.to_datetime(df["Sonraki_Tarih"]).dt.date

    # İstenmeyen sütunları kaldır
    df_export = df.drop(columns=["Borç_Adat", "Alacak_Adat"], errors="ignore")

    # Profesyonel Excel çıktısı
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        # --- Sayfa 1: Ayrıntılı Hesaplama ---
        df_export.to_excel(writer, index=False, sheet_name="Ayrıntılı Hesaplama", startrow=6)
        sheet1 = writer.sheets["Ayrıntılı Hesaplama"]

        # Başlık
        sheet1.merge_range("A1:E1", "ORTAKLAR CARİ HESAP ADAT HESAPLAMA RAPORU", workbook.add_format({
            "bold": True, "font_size": 14, "align": "center", "valign": "vcenter"
        }))
        sheet1.write("A3", "Dönem Başlangıç:", workbook.add_format({"bold": True}))
        sheet1.write("B3", str(donem_baslangic))
        sheet1.write("A4", "Dönem Bitiş:", workbook.add_format({"bold": True}))
        sheet1.write("B4", str(donem_bitis))

        # Biçimler
        header_format = workbook.add_format({
            "bold": True, "text_wrap": True, "valign": "middle", "align": "center",
            "border": 1, "bg_color": "#D9E1F2"
        })
        money_format = workbook.add_format({"num_format": "#,##0.00", "border": 1})
        normal_format = workbook.add_format({"border": 1})
        date_format = workbook.add_format({"num_format": "dd.mm.yyyy", "border": 1})

        # Sütun başlıklarını biçimlendir
        for col_num, value in enumerate(df_export.columns.values):
            sheet1.write(6, col_num, value, header_format)

        # Sütun biçimleri
        for col_num, col_name in enumerate(df_export.columns):
            if "Tarih" in col_name:
                sheet1.set_column(col_num, col_num, 14, date_format)
            elif "Borç" in col_name or "Alacak" in col_name or "Faiz" in col_name:
                sheet1.set_column(col_num, col_num, 18, money_format)
            else:
                sheet1.set_column(col_num, col_num, 15, normal_format)

        # --- Sayfa 2: Dönem Özeti ---
        ozet.to_excel(writer, index=False, sheet_name="Dönem Özeti", startrow=2)
        sheet2 = writer.sheets["Dönem Özeti"]
        sheet2.merge_range("A1:E1", "DÖNEM ÖZETİ", workbook.add_format({
            "bold": True, "font_size": 13, "align": "center", "valign": "vcenter", "bg_color": "#BDD7EE"
        }))
        sheet2.set_column("A:E", 25)
        sheet2.write("A8", "Hazırlayan:", workbook.add_format({"italic": True}))
        sheet2.write("B8", "Murat Uluat")




    st.download_button(
        label="📥 Profesyonel Excel Çıktısını İndir",
        data=output.getvalue(),
        file_name="Adat_Hesaplama_Raporu.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


else:
    st.info("Lütfen bir muavin dosyası yükleyin ve hesaplama yapmak için ayarlamaları tamamlayın.")
