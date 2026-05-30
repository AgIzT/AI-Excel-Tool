# -*- coding: utf-8 -*-
"""
workers.py
后台线程：把耗时的 AI 识别放到 QThread，避免阻塞 UI。
只负责「识别」，导出在用户复核确认后由主线程执行（导出很快）。
"""

from PySide6.QtCore import QThread, Signal

from ocr_to_excel import extract_batch


class ExtractWorker(QThread):
    progress_update = Signal(int, int)
    log_message = Signal(str)
    extract_finished = Signal(object)   # 传出 extract_batch 的结果 dict
    task_failed = Signal(str)

    def __init__(self, image_paths, handwriting, template_path, api_config):
        super().__init__()
        self._image_paths = list(image_paths)
        self._handwriting = bool(handwriting)
        self._template_path = template_path
        self._api_config = api_config

    def run(self):
        try:
            def log_cb(msg: str):
                self.log_message.emit(msg)

            def progress_cb(current: int, total: int):
                self.progress_update.emit(current + 1, total)

            extraction = extract_batch(
                image_paths=self._image_paths,
                log_callback=log_cb,
                handwriting=self._handwriting,
                progress_callback=progress_cb,
                template_path=self._template_path,
                api_config=self._api_config,
            )
            self.extract_finished.emit(extraction)
        except Exception as e:
            self.task_failed.emit(str(e))
