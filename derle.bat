@echo off
chcp 65001 > nul
title LeadScout Tek Tıkla Derleme İstasyonu v3.5
color 0b

echo ============================================================
echo   Altunay Soft ^| LeadScout Otomatik EXE Derleme Sistemi 🛰️
echo ============================================================
echo.

echo [1/4] Derleme için gerekli kütüphaneler denetleniyor ve kuruluyor...
python -m pip install --upgrade pip --quiet
python -m pip install pyinstaller playwright groq pandas openpyxl --quiet
if %errorlevel% neq 0 (
    echo.
    echo ❌ HATA: Python yüklü değil veya internet bağlantısı yok!
    pause
    exit
)
echo      - Kütüphaneler hazır.
echo.

echo [2/4] PyInstaller motoru Python üzerinden ateşleniyor...
echo      - (Bu işlem bilgisayar hızına göre 1-2 dakika sürebilir, lütfen bekleyin...)
python -m PyInstaller --noconfirm --onefile --windowed --name="LeadScout_Enterprise" "leadscout_enterprise.py"
if %errorlevel% neq 0 (
    echo.
    echo ❌ HATA: Derleme esnasında bir sorun oluştu!
    pause
    exit
)
echo      - EXE başarıyla oluşturuldu.
echo.

echo [3/4] Klasör yapısı temizleniyor, gereksiz kalıntılar süpürülüyor...
rmdir /s /q build
del /q LeadScout_Enterprise.spec
echo      - Temizlik tamamlandı.
echo.

echo [4/4] Harika! İşlem başarıyla tamamlandı.
echo ============================================================
echo 🚀 Bağımsız EXE dosyanız "dist" klasörünün içine bırakıldı!
echo 🛰️ Bu EXE'yi her bilgisayara götürüp doğrudan çalıştırabilirsiniz.
echo ============================================================
echo.
pause