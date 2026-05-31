# -*- coding: utf-8 -*-
"""
workers.py
后台线程：把耗时的 AI 识别放到 QThread，避免阻塞 UI。
只负责「识别」，导出在用户复核确认后由主线程执行（导出很快）。
"""

from PySide6.QtCore import QThread, Signal

from ocr_to_excel import extract_batch, test_connectivity


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


class ConnectivityWorker(QThread):
    """设置页「测试连通性」：后台发一次最小请求，避免网络等待卡住 UI。"""

    finished_ok = Signal(str)    # 成功，带模型回显文本
    finished_err = Signal(str)   # 失败，带错误信息

    def __init__(self, api_config):
        super().__init__()
        self._api_config = api_config

    def run(self):
        try:
            reply = test_connectivity(self._api_config)
            self.finished_ok.emit(reply)
        except Exception as e:
            self.finished_err.emit(str(e))
