#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Klima Aylık Tüketim Raporu - Professional GUI v3
Folkart Blu Çeşme Yönetimi İçin
"""

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import threading
from pathlib import Path
import sys
import os
from datetime import datetime

# Ana modülü import et
sys.path.insert(0, os.path.dirname(__file__))
from klima_final import PPDRawParser

class KlimaGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Klima Tüketim Raporu - Folkart Blu Çeşme")
        self.root.geometry("1000x750")
        self.root.minsize(800, 600)
        
        # Font ayarları (Tema ayarlarından önce tanımla)
        self.title_font = ("Segoe UI", 16, "bold")
        self.header_font = ("Segoe UI", 12, "bold")
        self.normal_font = ("Segoe UI", 10)
        self.mono_font = ("Consolas", 9)
        
        # Tema ayarları
        style = ttk.Style()
        style.theme_use('alt')  # Daha kontrol edilebilir theme
        
        # Modern Minimalist Renk Şeması - Siyah Beyaz
        self.bg_color = "#ffffff"          # Temiz beyaz arka plan
        self.header_color = "#000000"      # Siyah başlık
        self.accent_color = "#000000"      # Siyah vurgu
        self.success_color = "#000000"     # Siyah
        self.error_color = "#cc0000"       # Koyu kırmızı (sadece hata için)
        
        # TTK Style tanımlamaları - Siyah Beyaz (Minimalist)
        style.configure('TFrame', background=self.bg_color)
        style.configure('TLabel', background=self.bg_color, foreground="#000000")
        style.configure('TLabelframe', background=self.bg_color, foreground="#000000")
        style.configure('TLabelframe.Label', background=self.bg_color, foreground="#000000", font=self.header_font)
        style.configure('TButton', background="#f0f0f0", foreground="#000000")
        style.map('TButton', 
                  background=[('active', '#e0e0e0'), ('pressed', '#d0d0d0')])
        style.configure('TNotebook', background=self.bg_color)
        style.configure('TNotebook.Tab', background=self.bg_color)
        
        self.root.configure(bg=self.bg_color)
        
        self.parser = PPDRawParser()
        self.selected_file = None
        self.ppd_df = None
        self.output_dir = None  # kullanıcı seçimiyle belirlenecek kayıt dizini
        
        self.create_ui()
    
    def create_ui(self):
        """Modern UI oluştur"""
        # Notebook (Tabs)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Tab 1: Ana İşlem
        self.tab_main = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_main, text="Rapor Oluştur")
        self.create_main_tab()
        
        # Tab 2: Hakkında
        self.tab_about = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_about, text="Hakkında")
        self.create_about_tab()
    
    def create_main_tab(self):
        """Ana işlem sekmesi"""
        main_frame = ttk.Frame(self.tab_main)
        main_frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        # Başlık
        title_label = ttk.Label(main_frame, text="Klima Tüketim Raporu Oluşturma", font=self.title_font)
        title_label.pack(pady=(0, 20))
        
        # Dosya Seçimi Bölümü
        file_frame = ttk.LabelFrame(main_frame, text="1. Dosya Seçimi", padding=15)
        file_frame.pack(fill="x", pady=10)
        
        file_btn_frame = ttk.Frame(file_frame)
        file_btn_frame.pack(fill="x", pady=10)
        
        self.btn_browse = ttk.Button(file_btn_frame, text="Dosya Seç", 
                                      command=self.select_file)
        self.btn_browse.pack(side="left", padx=5)
        
        self.file_label = ttk.Label(file_btn_frame, text="Dosya seçilmedi", 
                                    foreground="red", font=self.normal_font)
        self.file_label.pack(side="left", padx=20)
        
        # İşlem Bölümü
        process_frame = ttk.LabelFrame(main_frame, text="2. İşlem", padding=15)
        process_frame.pack(fill="x", pady=10)
        
        btn_frame = ttk.Frame(process_frame)
        btn_frame.pack(fill="x", pady=10)
        
        self.btn_process = ttk.Button(btn_frame, text="Rapor Oluştur", 
                                       command=self.process_file, state="disabled")
        self.btn_process.pack(side="left", padx=5)
        
        # Durumu göster
        status_frame = ttk.Frame(process_frame)
        status_frame.pack(fill="x", pady=10)
        
        ttk.Label(status_frame, text="Durum:", font=self.header_font).pack(side="left")
        self.status_label = ttk.Label(status_frame, text="Hazır", 
                                      foreground="blue", font=self.normal_font)
        self.status_label.pack(side="left", padx=10)
        
        # İşlem Günlüğü
        log_frame = ttk.LabelFrame(main_frame, text="3. İşlem Günlüğü", padding=10)
        log_frame.pack(fill="both", expand=True, pady=10)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, 
                                                   font=self.mono_font, 
                                                   bg="#ffffff", wrap=tk.WORD)
        self.log_text.pack(fill="both", expand=True)
        
        # Altbilgi
        footer_frame = ttk.Frame(main_frame)
        footer_frame.pack(fill="x", pady=(10, 0))
        
        ttk.Separator(footer_frame, orient="horizontal").pack(fill="x", pady=5)
        
        footer_text = ttk.Label(footer_frame, 
                               text="v3.0 | Geliştiriciler: Serhat Yıldız | Folkart Blu Çeşme Yönetim Sistemi",
                               font=("Arial", 8), foreground="#666666")
        footer_text.pack(side="left")
    
    def create_about_tab(self):
        """Hakkında sekmesi"""
        about_frame = ttk.Frame(self.tab_about)
        about_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Başlık
        title = ttk.Label(about_frame, text="Klima Tüketim Raporu Hakkında", font=self.title_font)
        title.pack(pady=(0, 20))
        
        # ScrolledText ile açıklama
        text_frame = ttk.Frame(about_frame)
        text_frame.pack(fill="both", expand=True)
        
        about_text = scrolledtext.ScrolledText(text_frame, height=30, font=self.normal_font,
                                               wrap=tk.WORD, bg="#ffffff", relief="flat")
        about_text.pack(fill="both", expand=True)
        
        about_text.insert(tk.END, """📋 PROGRAM HAKKINDA

Klima Tüketim Raporu, Folkart Blu Çeşme Yönetim sistemi için PPD (Power Page Display) 
verilerini analiz ederek aylık ısıtma/soğutma tüketim raporları oluşturmak için 
tasarlanmıştır.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

⚙️ PROGRAM NASIL ÇALIŞIR?

1. PPD DOSYASINI OKUMA:
   • Program, intelligent Touch Manager cihazından PPD CSV formatında veri alır
   • Her sütun bir klimayı (örn: DAIRE 1A, DAIRE 1B, DAIRE 6A vs.) temsil eder
   • Her satır saat başı tüketim verilerini (Wh) içerir

2. DAIRE GRUPLANDIRMASI:
   • Alt birimler (1A, 1B, 1C vb.) otomatik olarak ana dairelere (1, 2, 3 vs.) 
     gruplandırılır
   • Örnek: DAIRE 1A + DAIRE 1B = DAIRE 1 (toplam tüketim hesaplanır)

3. HESAPLAMA MANTIGI:
   • Dikey toplama: Her daire için tüm saat verilerinin saati saatine toplanır
   • Yatay toplama: Tüm saatlerin toplamı hesaplanarak aylık tüketim bulunur
   • Formül: Aylık Tüketim (kWh) = ∑(Saatlik Tüketim Wh) / 1000

4. DAIRE SIRASI:
   • Raporlar daire_sirasi.txt dosyasında belirtilen sıraya göre düzenlenir
   • İçerisinde tüm dairelerin okuma sırası tanımlanmıştır

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

ÇIKTI DOSYALARI

Program iki formatta rapor oluşturur:

1. STANDART RAPOR:
   • Klima_01_2026_Tüketim.csv - Metin formatı (tüm yazılımlarda açılabilir)
   • Klima_01_2026_Tüketim.xlsx - Excel formatı (grafik ve analiz için)
   
   İçerik:
   - Daire ismi
   - Tüketim (Wh ve kWh cinsinden)
   - Daire türü (SÜİT / ORTAK)
   - İstatistikler (toplam, ortalama, min, max)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

ÖRNEK HESAPLAMA

Daire 1 (1A + 1B):
  • DAIRE 1A: 18.092 kWh
  • DAIRE 1B: 18.092 kWh
  • ─────────────────
  • TOPLAM:    36.184 kWh

Her saat için:
  Saat 01:00 → 5 Wh (1A) + 5 Wh (1B) = 10 Wh/saat
  Saat 02:00 → 7 Wh (1A) + 8 Wh (1B) = 15 Wh/saat
  ...
  [Tüm 730 saat toplanır]

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

TEMEL ÖZELLIKLER

- Otomatik daire gruplandırması
- Çoklu formatta çıktı (CSV + Excel)
- Daire sıralama desteği
- Detaylı istatistikler
- Hızlı ve güvenilir hesaplama

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

KULLANMA ADIMLARI

1. "Dosya Seç" butonuna tıklayın
2. PPD CSV dosyasını seçin (PPD_01012026_25022026.csv gibi)
3. "Rapor Oluştur" butonuna tıklayın
4. Raporlar çalışma dizinine kaydedilecektir
5. İstatistikleri günlükten kontrol edin

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

TEKNIK BİLGİLER

Yazılım: Python 3.10+
Kütüphaneler: pandas, openpyxl
Geliştirici: Serhat Yıldız
Version: 3.0
Tarih: Şubat 2026

GitHub: https://github.com/serhat-yildiz/intelligent-touch-manager-ppd-reader

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━""")
        
        about_text.config(state="disabled")
    
    def select_file(self):
        """Dosya seçici aç"""
        file_path = filedialog.askopenfilename(
            title="PPD Dosyasını Seçin",
            filetypes=[("CSV Dosyaları", "*.csv"), ("Tüm Dosyalar", "*.*")],
            initialdir=str(Path.home() / "Desktop")
        )
        
        if file_path:
            self.selected_file = file_path
            self.file_label.config(text=Path(file_path).name, foreground="green")
            self.btn_process.config(state="normal")
            self.log(f"[OK] Dosya seçildi: {Path(file_path).name}\n")
    
    def process_file(self):
        """PPD dosyasını standart formatta işle"""
        if not self.selected_file:
            messagebox.showwarning("Uyarı", "Lütfen bir dosya seçin!")
            return
        
        # Kaydedilecek klasörü seç
        self.output_dir = filedialog.askdirectory(
            title="Raporları kaydetmek için klasör seçin",
            initialdir=str(Path.home() / "Desktop")
        )
        if not self.output_dir:
            # kullanıcı iptal ettiyse işlemi durdur
            self.log("[WARNING] Kayıt dizini seçilmedi, işlem iptal edildi.\n")
            return
        
        self.btn_process.config(state="disabled")
        self.status_label.config(text="İşleniyor...", foreground="orange")
        self.log_text.delete("1.0", tk.END)
        
        thread = threading.Thread(target=self._process_standard)
        thread.daemon = True
        thread.start()
    
    def _process_standard(self):
        """Standart rapor işleme"""
        try:
            self.log("[*] PPD dosyası okunuyor...\n")
            
            import re
            
            # PPD parse et
            self.ppd_df = self.parser.parse_ppd_file(self.selected_file)
            self.log(f"[OK] {len(self.ppd_df)} alan verisi işlendi\n")
            
            # Tarih bilgisini filename'den çıkar
            filename = Path(self.selected_file).name
            match = re.search(r'(\d{2})(\d{2})(\d{4})_(\d{2})(\d{2})(\d{4})', filename)
            if match:
                end_month = int(match.groups()[1])
                end_year = match.groups()[2]
                month_year = f"{end_month}_{end_year}"
            else:
                month_year = "RAPOR"
            
            self.log("[*] Rapor oluşturuluyor...\n")
            
            # Özet oluştur ve export et
            summary = self.parser.create_summary(self.ppd_df)
            csv_file, xlsx_file = self.parser.export_results(
                self.ppd_df, summary, month_year, output_dir=self.output_dir
            )
            
            self.log("[OK] Standart rapor başarıyla oluşturuldu!\n")
            
            # İstatistikler
            self.log("\n📈 İSTATİSTİKLER:\n")
            for key, value in summary.items():
                if isinstance(value, float):
                    self.log(f"   • {key}: {value:.2f}\n")
                else:
                    self.log(f"   • {key}: {value}\n")
            
            self.log("\n[DONE] TAMAMLANDI!\n")
            self.status_label.config(text="Tamamlandı", foreground="black")
            
            # Dosya adları mesaj için tam yol olarak göster
            messagebox.showinfo("Başarılı",
                                f"Rapor oluşturuldu!\n\n- {csv_file}\n- {xlsx_file}")
            
        except Exception as e:
            self.log(f"\n[ERROR] HATA: {str(e)}\n")
            self.status_label.config(text="Hata!", foreground="#cc0000")
            messagebox.showerror("Hata", f"İşlem başarısız:\n{str(e)}")
        
        finally:
            self.btn_process.config(state="normal")
    
    def log(self, message):
        """Mesajı log alanına ekle"""
        self.log_text.insert(tk.END, message)
        self.log_text.see(tk.END)
        self.root.update()

def main():
    root = tk.Tk()
    app = KlimaGUI(root)
    root.mainloop()

if __name__ == "__main__":
    try:
        main()
    except ImportError as e:
        print("Hata: Gerekli paketler yüklü değil.")
        print("Lütfen şu komutu çalıştırın:")
        print("  pip install pandas openpyxl")
        print(f"\nDetay: {e}")
