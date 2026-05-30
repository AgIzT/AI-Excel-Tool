# -*- coding: utf-8 -*-
"""
workers.py
后台线程：把批量识别放到 QThread，避免阻塞 UI。
"""

from PySide6.QtCore import QThread, Signal

from ocr_to_excel import process_images_batch


class BatchWorker(QThread):
    progress_update = Signal(int, int)
    log_message = Signal(str)
    task_finished = Signal(list)
    task_failed = Signal(str)

    def __init__(self, image_paths, output_dir, handwriting, merge_output, template_path, api_config):
        super().__init__()
        self._image_paths = list(image_paths)
        self._output_dir = output_dir
        self._handwriting = bool(handwriting)
        self._merge_output = bool(merge_output)
        self._template_path = template_path
        self._api_config = api_config

    def run(self):
        try:
            def log_cb(msg: str):
                self.log_message.emit(msg)

            def progress_cb(current: int, total: int):
                self.progress_update.emit(current + 1, total)

            result = process_images_batch(
                image_paths=self._image_paths,
                output_dir=self._output_dir,
                log_callback=log_cb,
                handwriting=self._handwriting,
                merge_output=self._merge_output,
                merged_output_path=None,
                progress_callback=progress_cb,
                template_path=self._template_path,
                api_config=self._api_config,
            )
            self.task_finished.emit(result)
        except Exception as e:
            self.task_failed.emit(str(e))
