"""
Запусти этот скрипт один раз чтобы создать icon.ico
python make_icon.py
"""
from PIL import Image, ImageDraw, ImageFont
import os, sys, urllib.request

# Пути — от папки скрипта, а не от текущего каталога: build.py и app.py
# ищут icon.ico рядом со скриптами.
HERE = os.path.dirname(os.path.abspath(__file__))

# Скачиваем оригинальный логотип Фонтанки
LOGO_URL = "https://fontanka.ru/apple-touch-icon.png"
LOGO_PATH = os.path.join(HERE, "fontanka_logo.png")
ICO_PATH = os.path.join(HERE, "icon.ico")

downloaded = False
print("Скачиваю логотип...")
try:
    urllib.request.urlretrieve(LOGO_URL, LOGO_PATH)
    downloaded = True
    print("OK")
except Exception as e:
    print(f"Не удалось скачать: {e}")
    if not os.path.exists(LOGO_PATH):
        print("Положи логотип вручную как fontanka_logo.png рядом со скриптом")
        sys.exit(1)
    print("Использую локальный fontanka_logo.png")

src = Image.open(LOGO_PATH).convert("RGBA")

SIZE = 512
img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))

# Логотип на верхние 78%
logo_h = int(SIZE * 0.78)
logo = src.resize((SIZE, logo_h), Image.LANCZOS)
img.paste(logo, (0, 0), logo)

draw = ImageDraw.Draw(img)

# Нижняя полоса
draw.rectangle([0, logo_h, SIZE, SIZE], fill=(210, 95, 12))
# Разделитель
draw.rectangle([0, logo_h, SIZE, logo_h + 2], fill=(255, 255, 255, 60))

# Шрифт — ищем системный
font = None
font_candidates = [
    "C:/Windows/Fonts/segoeuib.ttf",   # Segoe UI Bold
    "C:/Windows/Fonts/arialbd.ttf",    # Arial Bold
    "C:/Windows/Fonts/calibrib.ttf",   # Calibri Bold
    "C:/Windows/Fonts/trebucbd.ttf",   # Trebuchet Bold
]
for fp in font_candidates:
    if os.path.exists(fp):
        font = ImageFont.truetype(fp, 54)
        print(f"Шрифт: {os.path.basename(fp)}")
        break

if not font:
    font = ImageFont.load_default()

text = "analytics"
bb = draw.textbbox((0, 0), text, font=font)
tw = bb[2] - bb[0]
th = bb[3] - bb[1]
tx = (SIZE - tw) // 2
ty = logo_h + (SIZE - logo_h - th) // 2 - 2

# Тень
draw.text((tx + 2, ty + 2), text, font=font, fill=(0, 0, 0, 70))
# Текст
draw.text((tx, ty), text, font=font, fill=(255, 255, 255))

# Сохраняем ICO всех размеров
sizes = [16, 32, 48, 64, 128, 256]
icons = [img.resize((s, s), Image.LANCZOS).convert("RGBA") for s in sizes]
icons[0].save(ICO_PATH, format="ICO",
              sizes=[(s, s) for s in sizes],
              append_images=icons[1:])

print("✅ icon.ico создан успешно!")
if downloaded:  # положенный вручную логотип не удаляем
    os.remove(LOGO_PATH)
