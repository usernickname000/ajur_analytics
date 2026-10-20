# ============================================================
# paths.py — где лежат настройки и JSON-файлы приложения
#
# При запуске из исходников это папка проекта. В собранном .exe
# (PyInstaller --onefile) модули распаковываются во временную папку,
# которая удаляется при закрытии программы, поэтому всё, что должно
# сохраняться между запусками (config.json, external_income.json,
# verified_figures.json, client_aliases.json), хранится РЯДОМ с .exe.
# ============================================================

import os
import shutil
import sys

_SRC_DIR = os.path.dirname(os.path.abspath(__file__))


def is_frozen() -> bool:
    return bool(getattr(sys, 'frozen', False))


def app_dir() -> str:
    """Папка для сохраняемых файлов: рядом с .exe или папка проекта."""
    if is_frozen():
        return os.path.dirname(os.path.abspath(sys.executable))
    return _SRC_DIR


def bundle_dir() -> str:
    """Папка с файлами, вшитыми в .exe (или папка проекта)."""
    return getattr(sys, '_MEIPASS', _SRC_DIR)


def data_path(name: str) -> str:
    """
    Путь к сохраняемому файлу данных. В .exe при первом запуске копирует
    вшитую версию файла рядом с .exe, дальше работает только с этой копией.
    """
    target = os.path.join(app_dir(), name)
    if is_frozen() and not os.path.exists(target):
        bundled = os.path.join(bundle_dir(), name)
        if os.path.exists(bundled):
            try:
                shutil.copy2(bundled, target)
            except Exception:
                return bundled
    return target
