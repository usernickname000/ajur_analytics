# ============================================================
# jsonio.py — безопасная запись JSON-справочников
# ============================================================

import json
import os
import shutil


def save_json_atomic(path, data, backup=True):
    """
    Пишет JSON через временный файл и os.replace: сбой посреди записи
    (OneDrive, нехватка места) не оставит полупустой файл. Предыдущая
    версия сохраняется рядом как <имя>.bak.
    """
    tmp_path = path + '.tmp'
    with open(tmp_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    if backup and os.path.exists(path):
        shutil.copy2(path, path + '.bak')
    os.replace(tmp_path, path)
