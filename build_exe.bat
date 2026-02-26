@echo off
REM Klima Tüketim Raporu - EXE Oluşturma Script
REM PyInstaller gereklidir: pip install pyinstaller

echo Paketler yükleniyor...
pip install pyinstaller -q

echo.
echo Klima_TüketimRaporu.exe olusturuluyor...
echo.

REM PyInstaller komutu - single file, windowed, icon
pyinstaller --onefile ^
    --windowed ^
    --name "Klima_TuketimRaporu" ^
    --distpath . ^
    --buildpath build ^
    --specpath . ^
    --add-data "klima_final.py:." ^
    --add-data "daire_sirasi.txt:." ^
    klima_gui_v3.py

echo.
echo BITTI! EXE dosyasi olusturuldu: Klima_TuketimRaporu.exe
echo.
pause
