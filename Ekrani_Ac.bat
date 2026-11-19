@echo off
title Karar Destek Sistemi - Yonetici Ekrani
cd /d "%~dp0"
echo Yonetici ekrani aciliyor... Tarayici 6 saniye icinde kendiliginden acilir.
echo Acilmazsa adres: http://localhost:8501
echo Kapatmak icin bu pencereyi kapatin (veya Ctrl+C).
echo.
start "" /min cmd /c "ping -n 7 127.0.0.1 >nul & start http://localhost:8501"
py -m streamlit run app.py --server.headless true
echo.
echo Ekran kapandi veya bir hata olustu. Hata varsa yukaridaki mesaji okuyun.
pause
