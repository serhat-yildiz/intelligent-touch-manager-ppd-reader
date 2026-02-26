#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Klima Aylık Tüketim Raporu - Professional GUI
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
        self.root.title("Klima Aylık Tüketim Raporu")
        self.root.geometry("900x650")
        self.root.resizable(False, False)
        
        # Ikon ayarla (simgeli pencere)
        try:
            self.root.iconbitmap(default='')  # Windows icon
        except:
            pass
        
        # Tema renkleri
        self.bg_color = "#f0f0f0"
        self.header_color = "#1F4E78"
        self.accent_color = "#4472C4"
        self.success_color = "#70AD47"
        self.error_color = "#ED7D31"
        
        self.root.configure(bg=self.bg_color)
        
        # Font ayarları
        self.title_font = ("Segoe UI", 14, "bold")
        self.header_font = ("Segoe UI", 11, "bold")
        self.normal_font = ("Segoe UI", 10)
        self.mono_font = ("Consolas", 9)
        
        self.parser = PPDRawParser()
        self.selected_file = None
        
        self.create_widgets()
    
    def create_widgets(self):
        """Arayüz bileşenlerini oluştur"""
        
        # Dosya Seçimi Bölümü
        file_frame = ttk.LabelFrame(self.root, text="1. Dosya Seçimi", padding=10)
        file_frame.pack(padx=20, pady=10, fill="x")
        
        self.file_label = ttk.Label(file_frame, text="Dosya seçilmedi", 
                                    font=self.normal_font, foreground="red")
        self.file_label.pack(anchor="w", pady=5)
        
        btn_frame = ttk.Frame(file_frame)
        btn_frame.pack(fill="x", pady=5)
        
        self.btn_browse = ttk.Button(btn_frame, text="📁 PPD Dosyası Seç", 
                                      command=self.select_file)
        self.btn_browse.pack(side="left", padx=5)
        
        # İşlem Bölümü
        process_frame = ttk.LabelFrame(self.root, text="2. İşlem", padding=10)
        process_frame.pack(padx=20, pady=10, fill="x")
        
        self.btn_process = ttk.Button(process_frame, text="▶ Raporu Oluştur", 
                                       command=self.process_file, state="disabled")
        self.btn_process.pack(side="left", padx=5)
        
        # Durumu göster
        self.status_label = ttk.Label(process_frame, text="Hazır", 
                                      font=self.normal_font, foreground="blue")
        self.status_label.pack(side="right", padx=5)
        
        # Çıktı Log Bölümü
        log_frame = ttk.LabelFrame(self.root, text="3. İşlem Günlüğü", padding=10)
        log_frame.pack(padx=20, pady=10, fill="both", expand=True)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=12, 
                                                   font=self.mono_font, 
                                                   state="normal")
        self.log_text.pack(fill="both", expand=True)
        
        # Altbilgi
        footer_frame = ttk.Frame(self.root)
        footer_frame.pack(pady=10, fill="x", padx=20)
        
        # Geliştiriciler bilgisi
        dev_text = "Geliştirici: Serhat Yıldız (ssyldz04@gmail.com) | Yazılım Geliştirme Uzmanı"
        self.dev_label = ttk.Label(footer_frame, text=dev_text, 
                                   font=("Arial", 8), foreground="#666666")
        self.dev_label.pack(side="left")
        
        self.version_label = ttk.Label(footer_frame, text="v2.0 - Klima Yönetim Sistemi", 
                                       font=("Arial", 8))
        self.version_label.pack(side="right")
    
    def select_file(self):
        """Dosya seçici aç"""
        file_path = filedialog.askopenfilename(
            title="PPD Dosyasını Seçin",
            filetypes=[("CSV Dosyaları", "*.csv"), ("Tüm Dosyalar", "*.*")],
            initialdir=str(Path.home() / "Desktop")
        )
        
        if file_path:
            self.selected_file = file_path
            self.file_label.config(text=file_path, foreground="green")
            self.btn_process.config(state="normal")
            self.log(f"✓ Dosya seçildi: {Path(file_path).name}\n")
    
    def process_file(self):
        """Dosyayı işle (ayrı thread'de)"""
        if not self.selected_file:
            messagebox.showwarning("Uyarı", "Lütfen bir dosya seçin!")
            return
        
        self.btn_process.config(state="disabled")
        self.status_label.config(text="İşleniyor...", foreground="orange")
        self.log_text.delete("1.0", tk.END)
        
        # Ayrı thread'de çalıştır
        thread = threading.Thread(target=self._process_in_thread)
        thread.daemon = True
        thread.start()
    
    def _process_in_thread(self):
        """İşlemi thread'de yap"""
        try:
            self.log("📂 Dosya okunuyor...")
            
            # Dosyayı işle
            import re
            from pathlib import Path
            
            # Ekim dosyasından mapping yüklemeyi dene
            ekim_file = Path(self.selected_file).parent / "Ekim.csv"
            if ekim_file.exists():
                self.log("📌 Ekim dosyasından numara eşleşmesi yükleniyor...\n")
                if self.parser.load_numara_mapping(str(ekim_file)):
                    self.log("✓ Numara eşleşmesi yüklendi\n")
                else:
                    self.log("⚠ Numara eşleşmesi yüklenemedi\n")
            else:
                pass  # Ekim.csv zorunlu değil
            
            df = self.parser.parse_ppd_file(self.selected_file)
            self.log(f"✓ {len(df)} alan okumalı verisi bulundu\n")
            
            self.log("✓ Veriler işlendi\n")
            
            filename = Path(self.selected_file).name
            date_info = self.parser.parse_dates_from_filename(filename) if hasattr(self.parser, 'parse_dates_from_filename') else None
            
            # Tarih bilgisini al - sadece sayı formatında (ay_yıl)
            match = re.search(r'(\d{2})(\d{2})(\d{4})_(\d{2})(\d{2})(\d{4})', filename)
            if match:
                end_month, end_year = match.groups()[1], match.groups()[2]
                month_year = f"{end_month}_{end_year}"  # Sadece 01_2026 formatı
            else:
                month_year = "RAPOR"
            
            output_file = f"Klima_{month_year.replace(' / ', '_')}_Tüketim.csv"
            excel_file = output_file.replace('.csv', '.xlsx')
            
            self.log(f"📊 Rapor oluşturuluyor...\n")
            
            summary = self.parser.create_summary(df)
            self.parser.export_results(df, summary, month_year)
            
            self.log(f"✓ CSV kaydedildi: {output_file}\n")
            self.log(f"✓ Excel kaydedildi: {excel_file}\n")
            
            # İstatistikler
            self.log("\n📈 İSTATİSTİKLER:\n")
            for key, value in summary.items():
                if isinstance(value, float):
                    self.log(f"   {key}: {value:.2f}\n")
                else:
                    self.log(f"   {key}: {value}\n")
            
            self.log("\n✅ TAMAMLANDI!\n")
            self.status_label.config(text="Tamamlandı ✓", foreground="green")
            
            messagebox.showinfo("Başarılı", f"Rapor oluşturuldu!\n\n{output_file}\n{excel_file}")
            
        except Exception as e:
            self.log(f"\n❌ HATA: {str(e)}\n")
            self.status_label.config(text="Hata!", foreground="red")
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
    except ImportError:
        print("Hata: Gerekli paketler yüklü değil.")
        print("Lütfen şu komutu çalıştırın:")
        print("  pip install pandas openpyxl")
        
        # Dosya Seçimi Bölümü
        file_frame = ttk.LabelFrame(self.root, text="1. Dosya Seçimi", padding=10)
        file_frame.pack(padx=20, pady=10, fill="x")
        
        self.file_label = ttk.Label(file_frame, text="Dosya seçilmedi", 
                                    font=self.normal_font, foreground="red")
        self.file_label.pack(anchor="w", pady=5)
        
        btn_frame = ttk.Frame(file_frame)
        btn_frame.pack(fill="x", pady=5)
        
        self.btn_browse = ttk.Button(btn_frame, text="📁 PPD Dosyası Seç", 
                                      command=self.select_file)
        self.btn_browse.pack(side="left", padx=5)
        
        # İşlem Bölümü
        process_frame = ttk.LabelFrame(self.root, text="2. İşlem", padding=10)
        process_frame.pack(padx=20, pady=10, fill="x")
        
        self.btn_process = ttk.Button(process_frame, text="▶ Raporu Oluştur", 
                                       command=self.process_file, state="disabled")
        self.btn_process.pack(side="left", padx=5)
        
        # Durumu göster
        self.status_label = ttk.Label(process_frame, text="Hazır", 
                                      font=self.normal_font, foreground="blue")
        self.status_label.pack(side="right", padx=5)
        
        # Çıktı Log Bölümü
        log_frame = ttk.LabelFrame(self.root, text="3. İşlem Günlüğü", padding=10)
        log_frame.pack(padx=20, pady=10, fill="both", expand=True)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=12, 
                                                   font=self.mono_font, 
                                                   state="normal")
        self.log_text.pack(fill="both", expand=True)
        
        # Altbilgi
        footer_frame = ttk.Frame(self.root)
        footer_frame.pack(pady=10, fill="x", padx=20)
        
        self.version_label = ttk.Label(footer_frame, text="v1.0 - Klima Yönetim Sistemi", 
                                       font=("Arial", 8))
        self.version_label.pack(side="right")
    
    def select_file(self):
        """Dosya seçici aç"""
        file_path = filedialog.askopenfilename(
            title="PPD Dosyasını Seçin",
            filetypes=[("CSV Dosyaları", "*.csv"), ("Tüm Dosyalar", "*.*")],
            initialdir=str(Path.home() / "Desktop")
        )
        
        if file_path:
            self.selected_file = file_path
            self.file_label.config(text=file_path, foreground="green")
            self.btn_process.config(state="normal")
            self.log(f"✓ Dosya seçildi: {Path(file_path).name}")
    
    def process_file(self):
        """Dosyayı işle (ayrı thread'de)"""
        if not self.selected_file:
            messagebox.showwarning("Uyarı", "Lütfen bir dosya seçin!")
            return
        
        self.btn_process.config(state="disabled")
        self.status_label.config(text="İşleniyor...", foreground="orange")
        self.log_text.delete("1.0", tk.END)
        
        # Ayrı thread'de çalıştır
        thread = threading.Thread(target=self._process_in_thread)
        thread.daemon = True
        thread.start()
    
    def _process_in_thread(self):
        """İşlemi thread'de yap"""
        try:
            self.log("📂 Dosya okunuyor...")
            
            # Dosyayı işle (output yakalamak için custom sürüm)
            import pandas as pd
            
            df = self.rapor.read_ppd(self.selected_file)
            self.log(f"✓ {len(df)} satır okumalı verisi bulundu")
            
            df = self.rapor.clean_data(df)
            self.log("✓ Veriler temizlendi")
            
            filename = Path(self.selected_file).name
            date_info = self.rapor.parse_dates_from_filename(filename)
            
            if date_info:
                month_name = self.rapor.months_tr.get(date_info['month'], str(date_info['month']))
                month_year = f"{month_name} / {date_info['year']}"
            else:
                month_year = "AYLIK RAPOR"
            
            output_file = f"Klima_{month_year.replace(' / ', '_')}_Tüketim.csv"
            excel_file = output_file.replace('.csv', '.xlsx')
            
            self.log(f"📊 Rapor oluşturuluyor...")
            self.rapor.export_csv(df, output_file, month_year, "")
            self.rapor.export_excel(df, excel_file, month_year, "")
            
            self.log(f"✓ CSV kaydedildi: {output_file}")
            self.log(f"✓ Excel kaydedildi: {excel_file}")
            
            # İstatistikler
            if 'TÜKETİM' in df.columns:
                self.log("\n📈 İstatistikler:")
                self.log(f"   Toplam Tüketim: {df['TÜKETİM'].sum():.2f}")
                self.log(f"   Ortalama: {df['TÜKETİM'].mean():.2f}")
                self.log(f"   En Yüksek: {df['TÜKETİM'].max():.2f}")
                self.log(f"   En Düşük: {df['TÜKETİM'].min():.2f}")
            
            self.log("\n✅ TAMAMLANDI!")
            self.status_label.config(text="Tamamlandı ✓", foreground="green")
            
            messagebox.showinfo("Başarılı", f"Rapor oluşturuldu!\n\n{output_file}\n{excel_file}")
            
        except Exception as e:
            self.log(f"\n❌ HATA: {str(e)}")
            self.status_label.config(text="Hata!", foreground="red")
            messagebox.showerror("Hata", f"İşlem başarısız:\n{str(e)}")
        
        finally:
            self.btn_process.config(state="normal")
    
    def log(self, message):
        """Mesajı log alanına ekle"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update()

def main():
    root = tk.Tk()
    app = KlimaGUI(root)
    root.mainloop()

if __name__ == "__main__":
    try:
        main()
    except ImportError:
        print("Hata: Gerekli paketler yüklü değil.")
        print("Lütfen şu komutu çalıştırın:")
        print("  pip install pandas openpyxl")
