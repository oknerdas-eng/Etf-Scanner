import os
import streamlit as st

st.title("ABD Borsası ETF Tarayıcı")
st.write(
    "Butona basarak Yahoo Finance üzerinden en güncel verileri çekebilir ve"
    " Excel raporunu oluşturabilirsiniz."
)

if st.button("Verileri Güncelle ve Excel Oluştur"):
  with st.spinner("Veriler çekiliyor, lütfen bekleyin..."):
    # Mevcut bot dosyanı burada çalıştırabilir veya fonksiyonlaştırabilirsin
    os.system("python etf_bot.py")
  st.success(
      "İşlem tamamlandı! 'etf_canli_kutuphane.xlsx' dosyası güncellendi."
  )

  if os.path.exists("etf_canli_kutuphane.xlsx"):
    with open("etf_canli_kutuphane.xlsx", "rb") as f:
      st.download_button(
          label="Excel Raporunu İndir",
          data=f,
          file_name="etf_canli_kutuphane.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )