@echo off
REM Klima Tüketim Raporu - EXE Oluşturma Script
REM PyInstaller gereklidir: pip install pyinstaller

echo Paketler yükleniyor...
".\.venv\Scripts\python.exe" -m pip install --upgrade pip setuptools wheel >nul
".\.venv\Scripts\python.exe" -m pip install pyinstaller pillow -q

echo.
echo Klima_TüketimRaporu.exe olusturuluyor...
echo.

echo.
REM PyInstaller komutu - tek satır ve doğru opsiyonlar kullanılıyor
REM exe'nin ikonunu "klima.ico" olarak belirtiyoruz
".\.venv\Scripts\python.exe" -m PyInstaller --onefile --windowed --name "Klima_TuketimRaporu" ^
    --icon "klima.ico" ^
    --distpath . --workpath build --specpath . ^
    --add-data "klima_final.py:." --add-data "daire_sirasi.txt:." klima_gui_v3.py

echo.
echo BITTI! EXE dosyasi olusturuldu: Klima_TuketimRaporu.exe
echo.
pause
