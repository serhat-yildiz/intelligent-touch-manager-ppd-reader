@echo off
REM Klima Tüketim Raporu Oluşturucu - Windows Başlatma Scripti

setlocal enabledelayedexpansion

echo.
echo ================================================================================
echo        KLIMA AYLIK TÜKETIM RAPORU OLUŞTURUCU
echo ================================================================================
echo.

REM Python kontrol et
python --version >nul 2>&1
if errorlevel 1 (
    echo [HATA] Python yüklü değil. Lütfen Python 3.7+ kurun.
    echo https://www.python.org/downloads/
    pause
    exit /b 1
)

REM Gerekli paketleri kontrol et
python -c "import pandas" >nul 2>&1
if errorlevel 1 (
    echo [UYARI] pandas paketi yüklü değil. Yükleniyor...
    pip install pandas openpyxl
)

python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo [UYARI] openpyxl paketi yüklü değil. Yükleniyor...
    pip install openpyxl
)

echo [✓] Tüm gereksinim kontrol edildi
echo.
echo Klima Converter başlatılıyor...
echo.

REM Ana programı çalıştır
python "%~dp0klima_converter.py"

if errorlevel 1 (
    echo.
    echo [HATA] Program çalıştırılırken hata oluştu.
    pause
) else (
    echo.
    echo [✓] İşlem tamamlandı.
    pause
)
