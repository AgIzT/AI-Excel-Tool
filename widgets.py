# -*- coding: utf-8 -*-
"""
widgets.py
自定义控件：支持拖拽落入文件的列表。
"""

import os

from PySide6.QtCore import QMimeData, Signal

from qfluentwidgets import ListWidget


class FileDropListWidget(ListWidget):
    files_dropped = Signal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragEnabled(False)
        self.setAlternatingRowColors(False)
        self.setSelectionMode(ListWidget.SelectionMode.SingleSelection)

    def dragEnterEvent(self, event):
        if self._has_urls(event.mimeData()):
            event.acceptProposedAction()
            return
        event.ignore()

    def dragMoveEvent(self, event):
        if self._has_urls(event.mimeData()):
            event.acceptProposedAction()
            return
        event.ignore()

    def dropEvent(self, event):
        paths = self._extract_paths(event.mimeData())
        if paths:
            self.files_dropped.emit(paths)
            event.acceptProposedAction()
            return
        event.ignore()

    @staticmethod
    def _has_urls(mime: QMimeData):
        return bool(mime and mime.hasUrls())

    @staticmethod
    def _extract_paths(mime: QMimeData):
        if not mime or not mime.hasUrls():
            return []
        paths = []
        for url in mime.urls():
            if not url.isLocalFile():
                continue
            p = os.path.abspath(url.toLocalFile())
            if os.path.isfile(p):
                paths.append(p)
        return paths
