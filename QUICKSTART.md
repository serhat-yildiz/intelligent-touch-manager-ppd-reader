# Klima Tüketim Raporu - Hızlı Başlangıç Rehberi

## 🚀 İlk Kullanım

### Adım 1: Kurulum
```bash
# setup.bat dosyasını çift tıklayın
# veya PowerShell'de çalıştırın:
powershell -ExecutionPolicy Bypass -File setup.bat
```

### Adım 2: Programı Çalıştırın

**Seçenek A - GUI Arayüzü (En Kolay)**
```bash
python klima_gui.py
# veya run.bat dosyasını çift tıklayın
```

**Seçenek B - Komut Satırı**
```bash
python klima_converter.py
# Dosya yolunu girin
```

---

## 📊 Dosya Formatı

### GIRIŞ: PPD Dosyası Örneği
```csv
ESKİ NUMARASI,YENİ NUMARASI,DURUM,İLK OKUMA,SON OKUMA
1,5,SÜİT,100.00,116.57
2,6,FOLKART,200.00,233.14
3,7,SÜİT,150.00,182.47
```

### ÇIKTI: Oluşturulan CSV
```
FOLKART BLU ÇEŞME YÖNETİMİ
ARALIK / 2025 DÖNEMİ
ISITMA/SOĞUTMA SAYAÇ TÜKETİMLERİ

ESKİ NUMARASI,YENİ NUMARASI,DURUM,İLK OKUMA,SON OKUMA,TÜKETİM
1,5,SÜİT,100.00,116.57,16.57
2,6,FOLKART,200.00,233.14,33.14
3,7,SÜİT,150.00,182.47,32.47
```

---

## 🛠️ Dosya Açıklaması

### Ana Dosyalar
- **klima_converter.py** - Ana program (komut satırı)
- **klima_gui.py** - Grafik arayüz (kolay kullanım)
- **run.bat** - Windows başlatıcısı

### Kurulum Dosyaları
- **setup.bat** - İlk kurulum (paket yükleme)
- **analyze_ppd.py** - PPD dosyası analiz aracı

### Belge Dosyaları
- **README.md** - Detaylı dokümantasyon
- **QUICKSTART.md** - Bu dosya

---

## 📋 Desteklenen Sütunlar

Program şu sütunları otomatik tanır:
- ✅ ESKİ NUMARASI
- ✅ YENİ NUMARASI
- ✅ DURUM (SÜİT, FOLKART, vb)
- ✅ İLK OKUMA
- ✅ SON OKUMA
- ✅ TÜKETİM (otomatik hesaplanır)

Ek sütunlar varsa, çıktıya da aktarılır.

---

## ❓ Sık Sorulan Sorular

### S: "Python bulunamadı" hatası alıyorum
**C:** Python yükleyin: https://www.python.org/downloads/
Kurulum sırasında "Add Python to PATH" seçeneğini işaretleyin.

### S: "ModuleNotFoundError: No module named 'pandas'" hatası
**C:** Paketleri yükleyin:
```bash
pip install pandas openpyxl
```

### S: Dosyamın sütun adları farklı
**C:** `analyze_ppd.py` çalıştırarak sütun adlarını kontrol edin:
```bash
python analyze_ppd.py "path/to/your/file.csv"
```

### S: Excel dosyası açılamıyor
**C:** Excel 2016+ gereklidir. Alternatif olarak CSV dosyasını açın.

### S: Tarih otomatik tanınmıyor
**C:** Dosya adında şu format olmalı:
```
PPD_DDMMYYYY_DDMMYYYY.csv
Örnek: PPD_01122025_30122025.csv
```

---

## 📈 Örnek İş Akışı

```
1. PPD_01122025_30122025.csv (klima programından)
        ↓
2. klima_converter.py veya klima_gui.py
        ↓
3. Klima_ARALIK_2025_Tüketim.csv
4. Klima_ARALIK_2025_Tüketim.xlsx
        ↓
5. Excel'de açabilir ve istediğiniz biçimde düzenleyebilirsiniz
```

---

## 🔧 İleri Seçenekler

### Encoding Sorunları
PPD dosyanız özel karakter sorunları varsa, `klima_converter.py` dosyasında satır 27'deki encoding listesini düzenleyin:

```python
for encoding in ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252', 'iso-8859-9']:
```

### Toplu İşleme
Birden çok dosya işlemek için:
```bash
for %f in (PPD*.csv) do python klima_converter.py "%f"
```

---

## 💡 İpuçları

1. **Dosya Adlandırması:** Dosya adını değiştirmeyin, tarihi otomatik olarak bulur
2. **Yedekleme:** İlk çalıştırmadan önce orijinal PPD dosyasının yedeğini alın
3. **Otomasyoon:** Scheduled Task'te `klima_converter.py` çalıştırabilirsiniz
4. **Denetim:** Çıktıyı Excel'de manuel olarak kontrol edin

---

## 📞 Destek

Sorunlar için:
1. `analyze_ppd.py` çalıştırarak dosya yapısını kontrol edin
2. Çıktı mesajlarını dikkatle okuyun
3. Dosya kodlamasını kontrol edin (UTF-8 tercihlidir)

---

## 📝 Değişiklikleri Takip Etme

Programı güncellerken:
- Yeni sütunlar otomatik olarak eklenir
- Tarih formatı otomatik tanınır
- Encoding sorunları otomatik çözülür

Başka sorular varsa, `README.md` dosyasını inceleyin.
