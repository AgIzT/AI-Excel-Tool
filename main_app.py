# -*- coding: utf-8 -*-
import os
import sys
from PySide6.QtCore import QDateTime, QSettings, QTimer, Qt
from PySide6.QtGui import QPixmap, QSurfaceFormat, QTextCursor
from PySide6.QtWidgets import QApplication, QVBoxLayout, QWidget

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
_meipass = getattr(sys, "_MEIPASS", None)
for _p in filter(None, [_meipass, BASE_DIR]):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from qfluentwidgets import (
    FluentWindow,
    FluentIcon,
    InfoBar,
    InfoBarPosition,
    NavigationItemPosition,
    SubtitleLabel,
    TextEdit,
    Theme,
    setTheme,
    setThemeColor,
)

from doc_page import DocPageMixin
from settings_page import SettingsPageMixin
from template_manager import TemplateManager

# QSettings 中保存的界面主题键；与设置页的下拉一致
THEME_KEY = "ui/theme"
THEME_MAP = {"AUTO": Theme.AUTO, "LIGHT": Theme.LIGHT, "DARK": Theme.DARK}
BRAND_COLOR = "#4f8ef7"


class MainWindow(DocPageMixin, SettingsPageMixin, FluentWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings("InvoiceAI", "InvoiceAIDesktop")
        self.template_manager = TemplateManager(BASE_DIR)
        self.selected_files = []
        self._selected_file_set = set()
        self.output_excel_paths = []
        self._worker = None
        self._processing = False
        self._log_buffer = []
        self._preview_path = ""
        self._preview_source = QPixmap()

        self._log_flush_timer = QTimer(self)
        self._log_flush_timer.setSingleShot(True)
        self._log_flush_timer.setInterval(90)
        self._log_flush_timer.timeout.connect(self._flush_log_buffer)

        self._resize_timer = QTimer(self)
        self._resize_timer.setSingleShot(True)
        self._resize_timer.setInterval(180)
        self._resize_timer.timeout.connect(self._on_resize_stable)

        self._build_window()
        self._init_navigation()
        self._append_log("PySide6 + QFluentWidgets 外壳已启动")

    def _build_window(self):
        self.setWindowTitle("全自动单据入库系统")
        self.resize(1360, 860)
        self.setMinimumSize(1080, 700)

    def _init_navigation(self):
        # FluentWindow 自带左侧导航 + 内容堆栈；每个子界面必须先设唯一 objectName
        self.doc_interface = self._build_doc_page()
        self.doc_interface.setObjectName("docInterface")
        self.log_interface = self._build_log_page()
        self.log_interface.setObjectName("logInterface")
        self.settings_interface = self._build_setting_page()
        self.settings_interface.setObjectName("settingsInterface")

        self.addSubInterface(self.doc_interface, FluentIcon.DOCUMENT, "单据处理")
        self.addSubInterface(
            self.log_interface, FluentIcon.HISTORY, "运行日志", NavigationItemPosition.BOTTOM
        )
        self.addSubInterface(
            self.settings_interface, FluentIcon.SETTING, "设置", NavigationItemPosition.BOTTOM
        )

    def _build_log_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 18, 24, 22)
        layout.setSpacing(12)
        title = SubtitleLabel("运行日志", page)
        layout.addWidget(title, 0)
        self.log_box = TextEdit(page)
        self.log_box.setReadOnly(True)
        layout.addWidget(self.log_box, 1)
        return page

    # ---- 日志 ----
    def _append_log(self, message: str):
        ts = QDateTime.currentDateTime().toString("HH:mm:ss")
        self._log_buffer.append(f"[{ts}] {message}")
        if not self._log_flush_timer.isActive():
            self._log_flush_timer.start()

    def _flush_log_buffer(self):
        if not self._log_buffer:
            return
        chunk = "\n".join(self._log_buffer) + "\n"
        self._log_buffer.clear()
        self.log_box.insertPlainText(chunk)
        self.log_box.moveCursor(QTextCursor.MoveOperation.End)

    # ---- 浮层提示 ----
    def _toast(self, kind: str, title: str, content: str = "", duration: int = 3000):
        fn = {
            "success": InfoBar.success,
            "error": InfoBar.error,
            "warning": InfoBar.warning,
            "info": InfoBar.info,
        }.get(kind, InfoBar.info)
        fn(
            title=title,
            content=content,
            orient=Qt.Orientation.Horizontal,
            isClosable=True,
            position=InfoBarPosition.TOP,
            duration=duration,
            parent=self,
        )

    # ---- 预览随窗口大小重渲染 ----
    def resizeEvent(self, event):
        super().resizeEvent(event)
        # FluentWindow 初始化期间可能触发 resize，此时计时器尚未就绪
        if getattr(self, "_resize_timer", None) is not None:
            self._resize_timer.start()

    def _on_resize_stable(self):
        self._render_preview()


def _configure_qt_acceleration():
    QApplication.setAttribute(Qt.ApplicationAttribute.AA_UseDesktopOpenGL, True)
    QApplication.setAttribute(Qt.ApplicationAttribute.AA_ShareOpenGLContexts, True)
    fmt = QSurfaceFormat()
    fmt.setRenderableType(QSurfaceFormat.RenderableType.OpenGL)
    fmt.setSwapBehavior(QSurfaceFormat.SwapBehavior.DoubleBuffer)
    fmt.setSwapInterval(1)
    fmt.setSamples(4)
    QSurfaceFormat.setDefaultFormat(fmt)


def _apply_saved_theme():
    settings = QSettings("InvoiceAI", "InvoiceAIDesktop")
    name = str(settings.value(THEME_KEY, "AUTO", str) or "AUTO").upper()
    setTheme(THEME_MAP.get(name, Theme.AUTO))


def main():
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )
    _configure_qt_acceleration()
    app = QApplication(sys.argv)
    setThemeColor(BRAND_COLOR)
    _apply_saved_theme()
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
