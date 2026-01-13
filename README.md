# Klima Aylık Tüketim Raporu Oluşturucu

## Nedir?
PPD klima okumaları dosyasını, anlamlı bir formata (CSV ve Excel) dönüştüren Python programıdır.

## Özellikler
- ✅ PPD dosyasını otomatik olarak parse eder
- ✅ Aylık tüketim hesaplamalarını yapıştırır (SON OKUMA - İLK OKUMA)
- ✅ CSV ve Excel formatında çıktı oluşturur
- ✅ Türkçe ay adlarını otomatik olarak ekler
- ✅ İstatistikler (ortalama, toplam, min, max) sağlar

## Kurulum

### 1. Python Gereksinimi
Python 3.7 veya üzeri gereklidir.

### 2. Gerekli Paketleri Yükleyin
```bash
pip install pandas openpyxl
```

## Kullanım

### Seçenek 1: Komut Satırında
```bash
python klima_converter.py
# Dosya yolunu girin ve Enter tuşuna basın
```

### Seçenek 2: Dosyayı Direct Açma (Windows)
PPD dosyasını şu şekilde işleyebilirsiniz:
- `klima_converter.py` dosyasını çift tıklayın
- Dosya yolunu girin

## PPD Dosya Formatı
Program şu formatı bekler:

```
ESKİ NUMARASI, YENİ NUMARASI, DURUM, İLK OKUMA, SON OKUMA, ...
1, 5, SÜİT, [başlangıç], [bitiş], ...
```

## Çıktı
Program şu dosyaları oluşturur:
- `Klima_AYEKIM_2025_Tüketim.csv` - CSV formatı
- `Klima_AYEKIM_2025_Tüketim.xlsx` - Excel formatı

## Örnek Çıktı
```
FOLKART BLU ÇEŞME YÖNETİMİ
EKİM / 2025 DÖNEMİ
ISITMA/SOĞUTMA SAYAÇ TÜKETİMLERİ

ESKİ NUMARASI | YENİ NUMARASI | DURUM  | İLK OKUMA | SON OKUMA | TÜKETİM
1             | 5             | SÜİT   | 100       | 120       | 20
2             | 6             | FOLKART| 200       | 250       | 50
```

## Sorun Giderme

### Encoding Hatası
Dosyanız farklı bir encoding kullanıyorsa:
- `klima_converter.py` dosyasında satır 27'deki encoding listesini düzenleyin
- `cp1252`, `iso-8859-9` vb. ekleyin

### Sütun Adlarının Tanınmaması
- `analyze_ppd.py` çalıştırarak dosyanızın yapısını kontrol edin
- Sütun adlarını programda güncelle

## Geliştiriciler & Yazılımcılar

### Orijinal Geliştirici
- **Serhat Yıldız** - Yazılım Geliştirme Uzmanı
  - E-mail: ssyldz04@gmail.com

## Lisans
MIT
