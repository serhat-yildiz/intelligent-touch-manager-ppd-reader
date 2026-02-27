"""Basit bir Python scripti ile proje ikonu (klima.ico) üretir.

Bu dosya Pillow kütüphanesini kullanır ve çalışma dizininde 256x256/64x64 ICO
dosyaları oluşturur. `build_exe.bat` zaten pillow'u yüklediği için ayrıca bir şey
yüklemenize gerek yok.

Çalıştırmak için:
    python make_icon.py
"""
from PIL import Image, ImageDraw, ImageFont

size = (256, 256)
img = Image.new('RGBA', size, '#1F4E78')
draw = ImageDraw.Draw(img)

# yüklenebilen bir font seçmeye çalış
try:
    font = ImageFont.truetype('arial.ttf', 200)
except IOError:
    font = ImageFont.load_default()

text = 'K'
# yazının boyutunu hesapla
if hasattr(font, 'getsize'):
    w, h = font.getsize(text)
else:
    bbox = draw.textbbox((0, 0), text, font=font)
    w = bbox[2] - bbox[0]
    h = bbox[3] - bbox[1]

draw.text(((size[0] - w) / 2, (size[1] - h) / 2), text, font=font, fill='white')
img.save('klima.ico', format='ICO', sizes=[(256, 256)])

# daha küçük bir sürüm
img.resize((64, 64), Image.LANCZOS).save('klima_64.ico', format='ICO', sizes=[(64, 64)])

print('İkonlar üretildi: klima.ico & klima_64.ico')