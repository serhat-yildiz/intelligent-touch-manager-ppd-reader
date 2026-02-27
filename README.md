# 🌡️ Klima Tüketim Raporu - Folkart Blu Çeşme

Folkart Blu Çeşme Yönetim sistemi için PPD (Power Page Display) verilerini analiz ederek aylık ısıtma/soğutma tüketim raporları oluşturan profesyonel Python uygulaması.

## ✨ Özellikler

✅ **Otomatik Daire Gruplandırması** - Alt birimler (1A, 1B, 1C) otomatik olarak ana dairelere birleştirilir  
✅ **Çoklu Formatta Çıktı** - CSV ve Excel formatlarında rapor oluşturma  
✅ **Sayaç Formatı Entegrasyonu** - Şubat sayaç verilerine eşleştirilmiş rapor  
✅ **Daire Sıralama** - Kullanıcı tarafından tanımlanmış okuma sırası  
✅ **Detaylı İstatistikler** - Toplam, ortalama, min-max tüketim analizi  
✅ **Modern GUI v3** - Sekmeli arayüz (Rapor Oluştur + Hakkında)  
✅ **EXE Desteği** - Standalone executable olarak çalıştırabilir  

## 🚀 Kurulum & Çalıştırma

### Seçenek 1: EXE (Önerilen - Kolay) ⭐

> Not: Artık exe dosyasına özel **klima.ico** ikonu dahil edilmiştir. Eğer ikonu kendiniz yeniden üretmek isterseniz `make_icon.py` scriptini çalıştırabilirsiniz (`pip install pillow` gerektirir).

```bash
# 1. build_exe.bat dosyasını çift tıklayın
# 2. Veya komut satırından çalıştırın:
build_exe.bat

# İşlem tamamlandıktan sonra:
Klima_TuketimRaporu.exe   ← Python işareti yerine kendi ikonunuz görünecek
```

**Avantajları:**
- Python yüklü olmasa da çalışır
- Taşınabilir (farklı bilgisayarda kullanabilir)
- Daha hızlı başlangıç

### Seçenek 2: Python Kaynağı (Geliştirici)

```bash
# Paketleri kur
pip install pandas openpyxl pdfplumber

# Programı çalıştır
python klima_gui_v3.py
```

## 📖 Kullanım Aşamaları

1. **📁 Dosya Seç**: PPD CSV dosyasını seçin (`PPD_01012026_25022026.csv`)
2. **▶ Rapor Türü Seç**:
   - "▶ Standart Rapor" - Detaylı CSV + Excel
   - "▶ Sayaç Formatı" - Şubat sayaç verilerine eşleştirilmiş
3. **✅ Tamamlandı**: Raporlar çalışma dizinine kaydedilir

## 📊 Çıktı Dosyaları

### Standart Rapor
- `Klima_01_2026_Tüketim.csv` - Tüm yazılımlarla uyumlu
- `Klima_01_2026_Tüketim.xlsx` - Grafik ve formüller için

**Örnek İçerik:**
| DAİRE_ADI | YENİ_NO | TİP | AYLIK_TUKETIM_KWH |
|-----------|---------|-----|-------------------|
| DAIRE 1 | 1 | SÜİT | 36.18 |
| DAIRE 2 | 2 | ORTAK | 138.24 |
| DAIRE 3 | 3 | SÜİT | 117.40 |

### Sayaç Formatı
- `Klima_01_2026_SAYAÇ_OKUMALARI.xlsx`

## 🔢 Hesaplama Mantığı

### Formül
```
Aylık Tüketim (kWh) = ∑(Saatlik Tüketim Wh) / 1000
```

### Adımlar
1. **PPD Dosyasını Oku** - 7. satırdan daire adlarını al
2. **Daire Sütunlarını Tespit Et** - DAIRE 1A, 1B, 1C vb.
3. **Saatleri Topla** - Her daire için 730 saat tüm değerler toplanır
4. **Gruplandır** - 1A + 1B + 1C = Daire 1
5. **Dönüştür** - Wh'ı kWh'a böl (÷ 1000)
6. **Sırala** - `daire_sirasi.txt`'ye göre düzenle

### Örnek Hesaplama (Daire 1)
```
DAIRE 1A:  18.092 kWh
DAIRE 1B:  18.092 kWh
────────────────────
TOPLAM:    36.184 kWh ← Otomatik birleştirilir
```

## 📁 Dosya Yapısı

```
intelligent-touch-manager-ppd-reader-main/
├── klima_gui_v3.py              ⭐ ANA PROGRAM (Modern UI)
├── klima_final.py               ← Veri işleme motoru
├── klima_gui.py                 ← Eski versiyon
├── daire_sirasi.txt             ← Daire okuma sırası (80 daire)
├── build_exe.bat                ← EXE oluşturmak için
├── Klima_TuketimRaporu.exe      📦 ÇALIŞTIRILACAK DOSYA
├── Klima_01_2026_Tüketim.csv    ← Çıktı (rapor)
├── Klima_01_2026_Tüketim.xlsx   ← Çıktı (rapor)
└── README.md                    ← Bu dosya
```

## 🔧 Konfigürasyon

### Daire Sırası (`daire_sirasi.txt`)
Raporların oluşturulacağı sıra:
```
5
6
7
8
...
80
```

### Eski-Yeni Numara Eşleştirmesi (`Ekim.csv`)
İsteğe bağlı. Varsa, eski numaralar raporlarda gösterilir.

## 🖥️ Sistem Gereksinimleri

| Sistem | Gereksinim |
|--------|-----------|
| İşletim Sistemi | Windows 10/11 |
| CPU | Herhangi bir işlemci |
| RAM | En az 512 MB |
| Disk | 500 MB (EXE için) |
| Python | 3.10+ (kaynak kodu çalıştırırken) |

## 🐛 Hata Giderme

### ❌ "PPD Dosyası Seçilmedi"
→ Lütfen CSV PPD dosyasını seçin

### ❌ "Sayaç Dosyası Bulunamadı"
→ "Şubat Klima Sayaç Okumaları.xlsx" aynı dizinde olmalı

### ❌ "ModuleNotFoundError: pandas"
```bash
pip install pandas openpyxl
```

### ❌ EXE oluşturma başarısız
```bash
pip install pyinstaller
build_exe.bat
```

## 📊 İstatistikler (Rapor Sonunda)

Her raporun sonunda:
- **Toplam Alan** - Kaç daire/alan var
- **Genel Toplam (kWh)** - Tüm dairelerin aylık tüketimi
- **Ortalama (kWh)** - Daire başı ortalama
- **En Yüksek / En Düşük** - Min-max tüketim
- **Tür Bazlı Toplam** - SÜİT ve ORTAK ayrı ayrı

## 👨‍💻 Teknik Bilgiler

| Bilgi | Değer |
|-------|-------|
| Dil | Python 3.10+ (tip açıklamaları eklendi) |
| GUI Framework | tkinter (standart Python) |
| Veri İşleme | pandas (vektörize edilmiş, parse hızı artırıldı) |
| Excel Yazma | openpyxl |
| Build Tool | PyInstaller (ikon desteği, onedir/onefile opsiyonları) |
| Version | 3.1 (kod refaktör, ikon, performans) |
| Geliştirici | Serhat Yıldız 
| Email | ssyldz04@gmail.com |

## 📝 Sürüm Tarihi

### v3.0 (Şubat 2026) ⭐ CURRENT
- ✅ Modern UI redesign (sekmeli arayüz)
- ✅ Detaylı "Hakkında" sayfası (program nasıl çalışıyor)
- ✅ EXE build sistemi (`build_exe.bat`)
- ✅ Ay isimleri kaldırıldı (01_2026 formatı)
- ✅ Daire hesaplama hatası düzeltildi (sütun indexing)
- ✅ Windows 10/11 optimizasyonu

### v2.0 (Ocak 2026)
- Standart rapor ve sayaç formatı
- Daire sıralama sistemi

### v1.0
- İlk versiyon

## 🔗 Linkler

- **GitHub**: https://github.com/serhat-yildiz/intelligent-touch-manager-ppd-reader
- **Email**: ssyldz04@gmail.com

## 📄 Lisans

Bu proje Folkart Blu Çeşme Yönetimi için özel olarak geliştirilmiştir.

---

**🏢 Folkart Blu Çeşme Yönetim Sistemi**  
**Profesyonel Klima Tüketim Raporlama Çözümü**  
*v3.0 | Şubat 2026*
