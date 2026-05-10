# build.py — сборка ajur_analytics в .exe
# Запуск: python build.py

import subprocess
import sys
import os

APP_DIR = os.path.dirname(os.path.abspath(__file__))

# Дополнительные файлы данных для включения в .exe
data_files = [
    ("external_income.json", "."),
    ("verified_figures.json", "."),
    ("config.json", ".") if os.path.exists(os.path.join(APP_DIR, "config.json")) else None,
]

add_data = []
for item in data_files:
    if item is None:
        continue
    src, dst = item
    full_src = os.path.join(APP_DIR, src)
    if os.path.exists(full_src):
        add_data.extend(["--add-data", f"{full_src};{dst}"])
    else:
        print(f"⚠ Файл не найден, не включён: {src}")

# Иконка
icon_path = os.path.join(APP_DIR, "icon.ico")
icon_args = ["--icon", icon_path] if os.path.exists(icon_path) else []
if not icon_args:
    print("⚠ icon.ico не найден — собираем без иконки")

cmd = [
    sys.executable, "-m", "PyInstaller",
    "--onefile",
    "--windowed",
    "--name", "АналитикаЗаказов",
    # pandas / numpy
    "--hidden-import=pandas",
    "--hidden-import=numpy",
    "--hidden-import=numpy.core._multiarray_umath",
    # openpyxl
    "--hidden-import=openpyxl",
    "--hidden-import=openpyxl.styles",
    "--hidden-import=openpyxl.styles.fills",
    "--hidden-import=openpyxl.chart",
    "--hidden-import=openpyxl.utils",
    # watchdog
    "--hidden-import=watchdog",
    "--hidden-import=watchdog.observers",
    "--hidden-import=watchdog.observers.winapi",
    "--hidden-import=watchdog.events",
    # модули приложения
    "--hidden-import=analytics",
    "--hidden-import=comparison",
    "--hidden-import=dashboard",
    "--hidden-import=watcher",
    "--hidden-import=verified_figures_editor",
    "--hidden-import=external_income_editor",
    *icon_args,
    *add_data,
    "app.py",
]

print("Запускаю PyInstaller...")
print("Команда:", " ".join(cmd))
print()

result = subprocess.run(cmd, cwd=APP_DIR)

if result.returncode == 0:
    exe_path = os.path.join(APP_DIR, "dist", "АналитикаЗаказов.exe")
    print()
    print("✅ Сборка успешна!")
    print(f"   Файл: {exe_path}")
else:
    print()
    print("❌ Ошибка сборки — смотри вывод выше")
