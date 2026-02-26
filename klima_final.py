#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Klima Aylık Tüketim Raporu - PPD Parser
Folkart Blu Çeşme Yönetimi İçin
"""

import re
import csv
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import pandas as pd

class PPDRawParser:
    def __init__(self):
        self.months_tr = {
            1: "OCAK", 2: "ŞUBAT", 3: "MART", 4: "NİSAN",
            5: "MAYIS", 6: "HAZİRAN", 7: "TEMMUZ", 8: "AĞUSTOS",
            9: "EYLÜL", 10: "EKİM", 11: "KASIM", 12: "ARALIK"
        }
        self.numara_mapping = {}  # YENİ -> ESKİ mapping
        self.daire_sirasi = []  # Daire okuma sırası
        self.load_daire_sirasi()
    
    def load_daire_sirasi(self):
        """Daire sırası dosyasını yükle"""
        try:
            sira_file = Path(__file__).parent / "daire_sirasi.txt"
            if sira_file.exists():
                with open(sira_file, 'r', encoding='utf-8') as f:
                    self.daire_sirasi = [int(line.strip()) for line in f if line.strip()]
                print(f"✓ Daire sırası yüklendi ({len(self.daire_sirasi)} daire)")
            else:
                print("⚠ daire_sirasi.txt dosyası bulunamadı (varsayılan sırası kullanılacak)")
        except Exception as e:
            print(f"⚠ Daire sırası yüklenemedi: {e}")
    
    def load_numara_mapping(self, ekim_file):
        """Ekim dosyasından ESKİ -> YENİ numara eşleşmesini yükle"""
        try:
            with open(ekim_file, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f, delimiter=';')
                data = list(reader)
            
            # Satır 10'dan itibaren veri
            for row in data[9:]:
                if not row or not row[0].strip():
                    continue
                
                eski_no = row[0].strip()
                yeni_no = row[1].strip() if len(row) > 1 else ""
                
                if eski_no and yeni_no:
                    self.numara_mapping[yeni_no] = eski_no
            
            print(f"✓ {len(self.numara_mapping)} numaralama eşleşmesi yüklendi")
            return True
        except Exception as e:
            print(f"⚠ Numara mapping yüklenemedi: {e}")
            return False
    
    def parse_ppd_file(self, file_path):
        """
        PPD dosyasını raw olarak parse et
        Format:
        - Satır 1-6: Başlık/metadata
        - Satır 7: Daire adları (DAIRE 5A;DAIRE 5B;...)
        - Satır 8+: Saatlik veriler
        """
        print(f"📂 PPD dosyası parslanıyor: {file_path}")
        
        with open(file_path, 'r', encoding='utf-8-sig') as f:
            lines = f.readlines()
        
        # Satır 7'den daire adlarını al - TÜM sütunları oku, sonra filtrele
        daire_line = lines[6]  # 0-indexed, satır 7 = satır 6
        all_columns = [x.strip() for x in daire_line.split(';')]
        
        # Tüm sütunlardan daire/alan adlarını al
        daire_names = []
        daire_column_indices = []  # Orijinal dosyadaki hangi sütun indexi
        
        for col_idx, col_name in enumerate(all_columns):  # Baştan başla, index 0'dan
            if col_name and any(x in col_name.upper() for x in ['DAIRE', 'LOBI', 'YONETIM', 'FITNESS', 'RES', 'BAYBAYAN', 'MUTFAK', 'P.O']):
                daire_names.append(col_name)
                daire_column_indices.append(col_idx)
        
        print(f"✓ {len(daire_names)} alan bulundu: {daire_names}")
        
        # Satır 8'den itibaren saatlik verileri al
        data_lines = lines[7:]  # Satır 8 ve sonrası
        
        # Daire bazlı toplam oluştur
        daire_totals = {name: 0 for name in daire_names}
        
        for i, line in enumerate(data_lines):
            values = line.strip().split(';')
            
            # Doğru sütunlardan değerleri topla
            for daire_idx, col_idx in enumerate(daire_column_indices):
                if col_idx < len(values):
                    try:
                        value = values[col_idx]
                        val = float(value) if value and value != '-' else 0
                        if val > 0:  # Negatif/hata değerleri atla
                            daire_totals[daire_names[daire_idx]] += val
                    except:
                        pass
        
        # Sonuçları DataFrame'e çevir
        results = []
        for name, total in daire_totals.items():
            daire_no = self.extract_daire_number(name)
            daire_type = self.get_daire_type(name)
            
            results.append({
                'DAİRE_ADI': name,
                'DAİRE_NO': daire_no,
                'TİP': daire_type,
                'AYLIK_TUKETIM_WH': total,
                'AYLIK_TUKETIM_KWH': total / 1000
            })
        
        df = pd.DataFrame(results)
        
        # Daire numarasına göre grupla ve topla
        # LOBI, YONETIM, FITNESS vb. ORTAK alanlar tek başına kalsın
        df_ortak = df[df['TİP'] == 'ORTAK'].copy()
        df_suit = df[df['TİP'] == 'SÜİT'].copy()
        
        # SÜİT'ler daire numarasına göre topla
        if len(df_suit) > 0:
            grouped = df_suit.groupby('DAİRE_NO').agg({
                'AYLIK_TUKETIM_WH': 'sum',
                'AYLIK_TUKETIM_KWH': 'sum'
            }).reset_index()
            grouped['DAİRE_ADI'] = 'DAIRE ' + grouped['DAİRE_NO'].astype(str)
            grouped['TİP'] = 'SÜİT'
            grouped = grouped[['DAİRE_ADI', 'DAİRE_NO', 'TİP', 'AYLIK_TUKETIM_WH', 'AYLIK_TUKETIM_KWH']]
            
            # ORTAK ve SÜİT'leri birleştir
            df = pd.concat([grouped, df_ortak], ignore_index=True)
        else:
            df = df_ortak
        
        # ESKİ_NUMARA mapping'i ekle (varsa)
        if self.numara_mapping:
            df['ESKİ_NUMARA'] = df['DAİRE_NO'].astype(str).map(self.numara_mapping)
            # ESKİ_NUMARA olmayanlara boş koy
            df['ESKİ_NUMARA'] = df['ESKİ_NUMARA'].fillna('')
        else:
            df['ESKİ_NUMARA'] = ''
        
        print(f"✓ {len(df)} alan verileri işlendi (daire bazlı toplandı)")
        
        return df
    
    def extract_daire_number(self, daire_name):
        """Daire adından numarası çıkar"""
        daire_name = daire_name.strip().upper()
        
        if 'LOBI' in daire_name:
            return 'LOBI'
        if 'YONETIM' in daire_name:
            return 'YONETIM'
        if 'FITNESS' in daire_name:
            return 'FITNESS'
        if 'MUTFAK' in daire_name or 'P.O' in daire_name or 'BAYBAYAN' in daire_name:
            return 'ORTAK'
        
        match = re.search(r'(\d+)', daire_name)
        if match:
            return int(match.group(1))
        
        return daire_name
    
    def get_daire_type(self, daire_name):
        """Daire tipi belirle"""
        daire_name = daire_name.strip().upper()
        
        if any(x in daire_name for x in ['LOBI', 'YONETIM', 'FITNESS', 'MUTFAK', 'P.O', 'BAYBAYAN', 'RES']):
            return 'ORTAK'
        
        return 'SÜİT'
    
    def create_summary(self, df):
        """İstatistikler oluştur"""
        summary = {
            'Toplam Alan': len(df),
            'Genel Aylık Toplam (kWh)': df['AYLIK_TUKETIM_KWH'].sum(),
            'Ortalama (kWh)': df['AYLIK_TUKETIM_KWH'].mean(),
            'En Yüksek (kWh)': df['AYLIK_TUKETIM_KWH'].max(),
            'En Düşük (kWh)': df['AYLIK_TUKETIM_KWH'].min(),
        }
        
        for dtype in df['TİP'].unique():
            subset = df[df['TİP'] == dtype]
            summary[f'{dtype} - Toplam (kWh)'] = subset['AYLIK_TUKETIM_KWH'].sum()
            summary[f'{dtype} - Sayı'] = len(subset)
        
        return summary
    
    def export_results(self, df, summary, month_year):
        """CSV ve Excel'e kaydet"""
        # Daire sırasını uygula
        if len(self.daire_sirasi) > 0:
            # Daire sırasına göre sort et (sadece integer daireler)
            df_sorted = pd.DataFrame()
            for daire_no in self.daire_sirasi:
                daire_match = df[df['DAİRE_NO'] == daire_no]
                if len(daire_match) > 0:
                    df_sorted = pd.concat([df_sorted, daire_match], ignore_index=True)
            
            # Kalanları (sırada olmayan - ORTAK alanlar) sonuna ekle
            used_daires = set(self.daire_sirasi)
            df_remaining = df[~df['DAİRE_NO'].isin(used_daires)]
            # ORTAK alanları isme göre sırala
            if len(df_remaining) > 0:
                df_remaining = df_remaining.sort_values('DAİRE_ADI').reset_index(drop=True)
            df = pd.concat([df_sorted, df_remaining], ignore_index=True)
        
        # Dosya adı - "/" karakterini "_" ile değiştir (Windows uyumluluğu)
        safe_filename = month_year.replace(' / ', '_')
        csv_file = f"Klima_{safe_filename}_Tüketim.csv"
        xlsx_file = f"Klima_{safe_filename}_Tüketim.xlsx"
        
        # CSV
        print(f"\n💾 CSV kaydediliyor: {csv_file}")
        with open(csv_file, 'w', encoding='utf-8-sig') as f:
            f.write("FOLKART BLU ÇEŞME YÖNETİMİ\n")
            f.write(f"{month_year} DÖNEMİ\n")
            f.write("ISITMA/SOĞUTMA - AYLLIK TÜKETİM RAPORU\n\n")
        
        # ESKİ_NUMARA sütununu varsa dahil et
        if 'ESKİ_NUMARA' in df.columns:
            df_export = df[['ESKİ_NUMARA', 'DAİRE_ADI', 'DAİRE_NO', 'TİP', 'AYLIK_TUKETIM_WH', 'AYLIK_TUKETIM_KWH']]
        else:
            df_export = df[['DAİRE_ADI', 'DAİRE_NO', 'TİP', 'AYLIK_TUKETIM_WH', 'AYLIK_TUKETIM_KWH']]
        
        df_export.to_csv(csv_file, mode='a', index=False, encoding='utf-8-sig')
        
        with open(csv_file, 'a', encoding='utf-8-sig') as f:
            f.write("\n\nÖZET İSTATİSTİKLERİ\n")
            for key, value in summary.items():
                if isinstance(value, float):
                    f.write(f"{key};{value:.2f}\n")
                else:
                    f.write(f"{key};{value}\n")
        
        # Excel
        print(f"💾 Excel kaydediliyor: {xlsx_file}")
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Tüketim"
        
        title_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        title_font = Font(color="FFFFFF", bold=True, size=14)
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=11)
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        row = 1
        # Başlık satırlarındaki sütun sayısını ayarla
        col_count = 6 if 'ESKİ_NUMARA' in df.columns else 5
        ws.merge_cells(f'A{row}:{"ABCDEF"[col_count-1]}{row}')
        cell = ws[f'A{row}']
        cell.value = "FOLKART BLU ÇEŞME YÖNETİMİ"
        cell.font = title_font
        cell.fill = title_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[row].height = 25
        row += 1
        
        ws.merge_cells(f'A{row}:{"ABCDEF"[col_count-1]}{row}')
        cell = ws[f'A{row}']
        cell.value = month_year
        cell.font = Font(color="FFFFFF", bold=True, size=12)
        cell.fill = title_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        row += 1
        
        ws.merge_cells(f'A{row}:{"ABCDEF"[col_count-1]}{row}')
        cell = ws[f'A{row}']
        cell.value = "ISITMA/SOĞUTMA - AYLLIK TÜKETİM RAPORU"
        cell.font = Font(color="FFFFFF", bold=True, size=11)
        cell.fill = title_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        row += 2
        
        # Başlık satırı
        if 'ESKİ_NUMARA' in df.columns:
            headers = ['ESKİ NO', 'DAİRE ADI', 'DAİRE NO', 'TİP', 'TÜKETİM (Wh)', 'TÜKETİM (kWh)']
        else:
            headers = ['DAİRE ADI', 'DAİRE NO', 'TİP', 'TÜKETİM (Wh)', 'TÜKETİM (kWh)']
        
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=row, column=col_idx)
            cell.value = header
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border
        
        row += 1
        
        # Veri satırları
        for _, data_row in df.iterrows():
            col = 1
            if 'ESKİ_NUMARA' in df.columns:
                ws.cell(row=row, column=col).value = data_row['ESKİ_NUMARA']
                ws.cell(row=row, column=col).border = border
                col += 1
            
            ws.cell(row=row, column=col).value = data_row['DAİRE_ADI']
            ws.cell(row=row, column=col).border = border
            col += 1
            
            ws.cell(row=row, column=col).value = data_row['DAİRE_NO']
            ws.cell(row=row, column=col).border = border
            ws.cell(row=row, column=col).alignment = Alignment(horizontal="center")
            col += 1
            
            ws.cell(row=row, column=col).value = data_row['TİP']
            ws.cell(row=row, column=col).border = border
            ws.cell(row=row, column=col).alignment = Alignment(horizontal="center")
            col += 1
            
            ws.cell(row=row, column=col).value = data_row['AYLIK_TUKETIM_WH']
            ws.cell(row=row, column=col).border = border
            ws.cell(row=row, column=col).number_format = '0'
            ws.cell(row=row, column=col).alignment = Alignment(horizontal="right")
            col += 1
            
            ws.cell(row=row, column=col).value = data_row['AYLIK_TUKETIM_KWH']
            ws.cell(row=row, column=col).border = border
            ws.cell(row=row, column=col).number_format = '0.00'
            ws.cell(row=row, column=col).alignment = Alignment(horizontal="right")
            
            row += 1
        
        # Özet
        row += 2
        ws.merge_cells(f'A{row}:{"ABCDEF"[col_count-1]}{row}')
        cell = ws[f'A{row}']
        cell.value = "ÖZET İSTATİSTİKLERİ"
        cell.font = Font(bold=True, size=11, color="FFFFFF")
        cell.fill = header_fill
        row += 1
        
        for key, value in summary.items():
            ws.cell(row=row, column=1).value = key
            ws.cell(row=row, column=1).font = Font(bold=True)
            ws.cell(row=row, column=1).border = border
            
            ws.cell(row=row, column=2).value = value
            ws.cell(row=row, column=2).border = border
            if isinstance(value, float):
                ws.cell(row=row, column=2).number_format = '0.00'
            ws.cell(row=row, column=2).alignment = Alignment(horizontal="right")
            
            row += 1
        
        ws.column_dimensions['A'].width = 12
        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['C'].width = 12
        ws.column_dimensions['D'].width = 12
        ws.column_dimensions['E'].width = 18
        ws.column_dimensions['F'].width = 18
        
        wb.save(xlsx_file)
        
        print(f"\n✅ TAMAMLANDI!")
        print(f"   • {csv_file}")
        print(f"   • {xlsx_file}")
        
        return csv_file, xlsx_file
    
    def load_subat_sayac_data(self, excel_file):
        """Şubat sayaç okumaları Excel dosyasından veri yükle ve formata dönüştür"""
        print(f"\n📊 Şubat Sayaç Okumaları yükleniyor: {excel_file}")
        
        try:
            wb = load_workbook(excel_file)
            ws = wb.active
            
            # Daire verilerini oku (Satır 10'dan başlıyor)
            sayac_data = {}
            for row_idx in range(10, 100):
                eski_no = ws.cell(row_idx, 2).value
                yeni_no = ws.cell(row_idx, 3).value
                durum = ws.cell(row_idx, 4).value
                tuketim = ws.cell(row_idx, 7).value
                
                # Eğer tüm veriler boş ise dur
                if eski_no is None and yeni_no is None:
                    break
                
                # Yeni numaraya göre depolamak daha iyi
                if yeni_no is not None:
                    try:
                        yeni_no = int(yeni_no) if isinstance(yeni_no, str) else yeni_no
                        tuketim = float(tuketim) if tuketim else 0
                        sayac_data[yeni_no] = {
                            'ESKİ_NO': eski_no,
                            'YENİ_NO': yeni_no,
                            'DURUM': durum,
                            'TUKETIM': tuketim
                        }
                    except:
                        pass
            
            print(f"✓ {len(sayac_data)} daire verisi yüklendi")
            return sayac_data
        
        except Exception as e:
            print(f"⚠ Şubat verileri yüklenemedi: {e}")
            return {}
    
    def export_sayac_format(self, df, sayac_data, month_year):
        """Sayaç formatında Excel raporu oluştur"""
        print(f"\n💾 Sayaç Formatı Excel kaydediliyor...")
        
        # "/" karakterini "_" ile değiştir (Windows uyumluluğu)
        safe_filename = month_year.replace(' / ', '_')
        xlsx_file = f"Klima_{safe_filename}_SAYAÇ_OKUMALARI.xlsx"
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Sayaç Okumaları"
        
        # Stiller
        title_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        title_font = Font(color="FFFFFF", bold=True, size=14)
        subtitle_font = Font(color="FFFFFF", bold=True, size=11)
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=10)
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Başlık
        row = 1
        ws.merge_cells(f'A{row}:F{row}')
        cell = ws[f'A{row}']
        cell.value = "FOLKART BLU ÇEŞME YÖNETİMİ"
        cell.font = title_font
        cell.fill = title_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[row].height = 22
        row += 1
        
        ws.merge_cells(f'A{row}:F{row}')
        cell = ws[f'A{row}']
        cell.value = month_year
        cell.font = subtitle_font
        cell.fill = title_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        row += 1
        
        ws.merge_cells(f'A{row}:F{row}')
        cell = ws[f'A{row}']
        cell.value = "ISITMA/SOĞUTMA SAYAÇ TÜKETİMLERİ"
        cell.font = subtitle_font
        cell.fill = title_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        row += 2
        
        # Başlık satırı
        headers = ['ESKİ NUMARASI', 'YENİ NUMARASI', 'DURUM', 'ISITMA/SOĞUTMA', 'İLK OKUMA', 'SON OKUMA', 'TÜKETİM']
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=row, column=col_idx)
            cell.value = header
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
        
        ws.row_dimensions[row].height = 25
        row += 1
        
        # Veri satırları - daire sırasına göre yerleştir
        # Sırası olan daireleri önce göster
        daire_order = self.daire_sirasi if len(self.daire_sirasi) > 0 else sorted(sayac_data.keys())
        
        for daire_no in daire_order:
            if daire_no not in sayac_data:
                continue
                
            sayac = sayac_data[daire_no]
            
            # PPD'den gelen tüketimi bul
            tuketim_kw = 0
            if len(df) > 0:
                # DAİRE_NO ile eşleştir
                daire_match = df[df['DAİRE_NO'] == daire_no]
                if len(daire_match) > 0:
                    tuketim_kw = daire_match.iloc[0]['AYLIK_TUKETIM_KWH']
            
            # Sayaç formatında kullan, eğer yoksa PPD verisi kullan
            tuketim_val = sayac.get('TUKETIM', tuketim_kw)
            
            ws.cell(row=row, column=1).value = sayac['ESKİ_NO']
            ws.cell(row=row, column=1).border = border
            ws.cell(row=row, column=1).alignment = Alignment(horizontal="center")
            
            ws.cell(row=row, column=2).value = sayac['YENİ_NO']
            ws.cell(row=row, column=2).border = border
            ws.cell(row=row, column=2).alignment = Alignment(horizontal="center")
            
            ws.cell(row=row, column=3).value = sayac['DURUM']
            ws.cell(row=row, column=3).border = border
            ws.cell(row=row, column=3).alignment = Alignment(horizontal="center")
            
            ws.cell(row=row, column=4).value = ""  # ISITMA/SOĞUTMA etiketi boş
            ws.cell(row=row, column=4).border = border
            
            ws.cell(row=row, column=5).value = ""  # İLK OKUMA
            ws.cell(row=row, column=5).border = border
            ws.cell(row=row, column=5).alignment = Alignment(horizontal="right")
            
            ws.cell(row=row, column=6).value = ""  # SON OKUMA
            ws.cell(row=row, column=6).border = border
            ws.cell(row=row, column=6).alignment = Alignment(horizontal="right")
            
            ws.cell(row=row, column=7).value = tuketim_val if tuketim_val else ""
            ws.cell(row=row, column=7).border = border
            ws.cell(row=row, column=7).number_format = '0.00'
            ws.cell(row=row, column=7).alignment = Alignment(horizontal="right")
            
            row += 1
        
        # Genişlikleri ayarla
        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 12
        ws.column_dimensions['D'].width = 20
        ws.column_dimensions['E'].width = 15
        ws.column_dimensions['F'].width = 15
        ws.column_dimensions['G'].width = 15
        
        wb.save(xlsx_file)
        print(f"✓ Sayaç formatı: {xlsx_file}")
        
        return xlsx_file
