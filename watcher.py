# ============================================================
# watcher.py — папка-наблюдатель
# Следит за папкой: если появился новый .xlsx — запускает анализ
# через 5 минут после появления файла (чтобы файл успел дописаться)
# ============================================================

import os
import time
import threading
import logging
from datetime import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

from analytics import run_analytics

WATCH_DELAY_SECONDS = 300   # 5 минут
SUPPORTED_EXTENSIONS = ('.xlsx', '.xls')

logger = logging.getLogger(__name__)


class _XlsxHandler(FileSystemEventHandler):
    """Обрабатывает события файловой системы."""

    def __init__(self, output_dir: str, log_callback):
        super().__init__()
        self._output_dir = output_dir
        self._log = log_callback
        self._pending: dict[str, threading.Timer] = {}  # path -> Timer
        self._lock = threading.Lock()

    def on_created(self, event):
        if event.is_directory:
            return
        path = event.src_path
        if not path.lower().endswith(SUPPORTED_EXTENSIONS):
            return
        # Игнорируем файлы, которые сами создаём (отчёты)
        basename = os.path.basename(path)
        if basename.startswith("аналитика_заказов_"):
            return
        self._schedule(path)

    def on_moved(self, event):
        """Срабатывает когда файл переименовывается (finish copy in Windows)."""
        if event.is_directory:
            return
        path = event.dest_path
        if not path.lower().endswith(SUPPORTED_EXTENSIONS):
            return
        basename = os.path.basename(path)
        if basename.startswith("аналитика_заказов_"):
            return
        self._schedule(path)

    def _schedule(self, path: str):
        with self._lock:
            # Если уже запланирован — сбрасываем таймер (файл ещё пишется)
            if path in self._pending:
                self._pending[path].cancel()
                self._log(f"⏱ Файл обновляется, таймер сброшен: {os.path.basename(path)}")

            self._log(
                f"📂 Обнаружен файл: {os.path.basename(path)} — "
                f"анализ запустится через {WATCH_DELAY_SECONDS // 60} мин."
            )
            timer = threading.Timer(
                WATCH_DELAY_SECONDS,
                self._run_analysis,
                args=[path]
            )
            timer.daemon = True
            timer.start()
            self._pending[path] = timer

    def _run_analysis(self, input_path: str):
        with self._lock:
            self._pending.pop(input_path, None)

        if not os.path.exists(input_path):
            self._log(f"⚠ Файл исчез до запуска анализа: {os.path.basename(input_path)}")
            return

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(
            self._output_dir,
            f"аналитика_заказов_{timestamp}.xlsx"
        )

        self._log(f"▶ Запускаю автоанализ: {os.path.basename(input_path)}")
        try:
            run_analytics(input_path, output_path, log=self._log)
            self._log(f"✅ Автоанализ завершён: {os.path.basename(output_path)}")
        except Exception as e:
            self._log(f"❌ Ошибка автоанализа ({os.path.basename(input_path)}): {e}")

    def cancel_all(self):
        with self._lock:
            for timer in self._pending.values():
                timer.cancel()
            self._pending.clear()


class FolderWatcher:
    """
    Публичный класс для управления наблюдателем из app.py.

    Использование:
        watcher = FolderWatcher(watch_dir, output_dir, log_callback)
        watcher.start()
        ...
        watcher.stop()
    """

    def __init__(self, watch_dir: str, output_dir: str, log_callback):
        self._watch_dir   = watch_dir
        self._output_dir  = output_dir
        self._log         = log_callback
        self._observer    = None
        self._handler     = None

    def start(self):
        if self._observer and self._observer.is_alive():
            return  # уже запущен

        self._handler = _XlsxHandler(self._output_dir, self._log)
        self._observer = Observer()
        self._observer.schedule(self._handler, self._watch_dir, recursive=False)
        self._observer.start()
        self._log(f"👁 Наблюдатель запущен: {self._watch_dir}")

    def stop(self):
        if self._observer:
            self._handler.cancel_all()
            self._observer.stop()
            self._observer.join(timeout=5)
            self._observer = None
            self._log("⏹ Наблюдатель остановлен.")

    @property
    def is_running(self) -> bool:
        return self._observer is not None and self._observer.is_alive()
