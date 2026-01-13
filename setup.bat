@echo off
REM Klima Tüketim Raporu - Kurulum Scriptı

setlocal enabledelayedexpansion

echo.
echo ================================================================================
echo                    KLIMA RAPORU - KURULUM ORTAMI
echo ================================================================================
echo.

REM Python var mı kontrol et
python --version >nul 2>&1
if errorlevel 1 (
    echo [HATA] Python yüklü değil!
    echo.
    echo Python 3.7+ kurmanız gerekiyor:
    echo https://www.python.org/downloads/
    echo.
    echo Kurulum sırasında "Add Python to PATH" seçeneğini İŞARETLEYİN!
    echo.
    pause
    exit /b 1
)

echo [✓] Python bulundu
python --version

REM pip var mı kontrol et
echo.
echo [*] Paket yöneticisi kontrol ediliyor...
pip --version >nul 2>&1
if errorlevel 1 (
    echo [HATA] pip bulunamadı
    pause
    exit /b 1
)

REM Gereken paketleri yükle
echo [*] Gerekli paketler yükleniyor...
echo     - pandas
echo     - openpyxl
echo.

pip install pandas openpyxl -q

if errorlevel 1 (
    echo [HATA] Paket kurulumu başarısız
    echo Lütfen internete bağlı olduğunuzdan emin olun.
    pause
    exit /b 1
)

echo [✓] Tüm paketler başarıyla yüklendi
echo.
echo ================================================================================
echo                         KURULUM TAMAMLANDI
echo ================================================================================
echo.
echo KULLANIMIN BAŞLA:
echo.
echo 1. KOMUT SATIRI VERSIYONU (Hızlı):
echo    python klima_converter.py
echo.
echo 2. ARAYÜZ VERSIYONU (Kolay):
echo    python klima_gui.py
echo.
echo 3. WINDOWS BAŞLATICI:
echo    run.bat (çift tıkla)
echo.
echo ================================================================================
echo.
pause
