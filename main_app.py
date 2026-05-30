# -*- coding: utf-8 -*-
import os
import sys
from PySide6.QtCore import QDateTime, QSettings, QTimer, Qt
from PySide6.QtGui import QPixmap, QSurfaceFormat, QTextCursor
from PySide6.QtWidgets import (
    QApplication,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QPushButton,
    QSizePolicy,
    QSpacerItem,
    QStackedWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
_meipass = getattr(sys, "_MEIPASS", None)
for _p in filter(None, [_meipass, BASE_DIR]):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from theme import build_stylesheet
from doc_page import DocPageMixin
from settings_page import SettingsPageMixin
from template_manager import TemplateManager


class MainWindow(DocPageMixin, SettingsPageMixin, QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings("InvoiceAI", "InvoiceAIDesktop")
        self.template_manager = TemplateManager(BASE_DIR)
        self.selected_files = []
        self._selected_file_set = set()
        self.output_excel_paths = []
        self._active_route = ""
        self._pages = {}
        self._route_indexes = {}
        self._nav_buttons = {}
        self._worker = None
        self._processing = False
        self._log_buffer = []
        self._log_collapsed = True
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

        self._route_builders = {
            "单据处理": self._build_doc_page,
            "设置": self._build_setting_page,
        }

        self._build_window()
        self._build_layout()
        self._apply_style()
        self._switch_route("单据处理")
        self._append_log("PySide6 阶段三已启动")

    def _build_window(self):
        self.setWindowTitle("全自动单据入库系统 - PySide6")
        self.resize(1360, 860)
        self.setMinimumSize(1080, 700)

    def _build_layout(self):
        root = QWidget(self)
        self.setCentralWidget(root)
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        content_wrap = QWidget(root)
        content_layout = QHBoxLayout(content_wrap)
        content_layout.setContentsMargins(12, 12, 12, 8)
        content_layout.setSpacing(12)
        root_layout.addWidget(content_wrap, 1)

        self.sidebar = QFrame(content_wrap)
        self.sidebar.setObjectName("sidebar")
        self.sidebar.setFixedWidth(220)
        sidebar_layout = QVBoxLayout(self.sidebar)
        sidebar_layout.setContentsMargins(12, 12, 12, 12)
        sidebar_layout.setSpacing(8)
        content_layout.addWidget(self.sidebar, 0)

        self._build_sidebar(sidebar_layout)

        self.workspace = QWidget(content_wrap)
        workspace_layout = QVBoxLayout(self.workspace)
        workspace_layout.setContentsMargins(0, 0, 0, 0)
        workspace_layout.setSpacing(10)
        content_layout.addWidget(self.workspace, 1)

        top_bar = QFrame(self.workspace)
        top_bar.setObjectName("topBar")
        top_bar.setFixedHeight(60)
        top_layout = QHBoxLayout(top_bar)
        top_layout.setContentsMargins(18, 8, 18, 8)
        top_layout.setSpacing(8)
        self.page_title = QLabel("单据处理", top_bar)
        self.page_title.setObjectName("pageTitle")
        top_layout.addWidget(self.page_title, 1)
        self.window_size_label = QLabel("窗口: -", top_bar)
        self.window_size_label.setObjectName("dimLabel")
        top_layout.addWidget(self.window_size_label, 0, Qt.AlignmentFlag.AlignRight)
        workspace_layout.addWidget(top_bar, 0)

        self.page_stack = QStackedWidget(self.workspace)
        self.page_stack.setObjectName("pageStack")
        workspace_layout.addWidget(self.page_stack, 1)

        self.log_drawer = QFrame(root)
        self.log_drawer.setObjectName("logDrawer")
        log_layout = QVBoxLayout(self.log_drawer)
        log_layout.setContentsMargins(12, 8, 12, 12)
        log_layout.setSpacing(6)
        root_layout.addWidget(self.log_drawer, 0)

        log_header = QWidget(self.log_drawer)
        log_header_layout = QHBoxLayout(log_header)
        log_header_layout.setContentsMargins(0, 0, 0, 0)
        log_header_layout.setSpacing(8)
        log_title = QLabel("实时日志", log_header)
        log_title.setObjectName("sectionTitle")
        log_header_layout.addWidget(log_title, 1)
        self.log_toggle_btn = QPushButton("展开", log_header)
        self.log_toggle_btn.setObjectName("logToggleButton")
        self.log_toggle_btn.setFixedSize(76, 28)
        self.log_toggle_btn.clicked.connect(self._toggle_log_drawer)
        log_header_layout.addWidget(self.log_toggle_btn, 0, Qt.AlignmentFlag.AlignRight)
        log_layout.addWidget(log_header, 0)

        self.log_box = QTextEdit(self.log_drawer)
        self.log_box.setObjectName("logBox")
        self.log_box.setReadOnly(True)
        self.log_box.setFixedHeight(190)
        self.log_box.setVisible(False)
        log_layout.addWidget(self.log_box, 0)

    def _build_sidebar(self, layout: QVBoxLayout):
        brand = QFrame(self.sidebar)
        brand.setObjectName("brandCard")
        brand.setFixedHeight(72)
        brand_layout = QVBoxLayout(brand)
        brand_layout.setContentsMargins(12, 8, 12, 8)
        brand_layout.setSpacing(0)
        title = QLabel("Invoice AI", brand)
        title.setObjectName("brandTitle")
        subtitle = QLabel("PySide6 业务版", brand)
        subtitle.setObjectName("dimLabel")
        brand_layout.addWidget(title, 0)
        brand_layout.addWidget(subtitle, 0)
        layout.addWidget(brand, 0)

        for route in ["单据处理", "设置"]:
            btn = QPushButton(route, self.sidebar)
            btn.setObjectName("navButton")
            btn.setProperty("active", False)
            btn.setFixedHeight(42)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.clicked.connect(lambda _=False, r=route: self._switch_route(r))
            layout.addWidget(btn, 0)
            self._nav_buttons[route] = btn

        divider = QFrame(self.sidebar)
        divider.setObjectName("divider")
        divider.setFixedHeight(1)
        layout.addWidget(divider, 0)
        layout.addItem(QSpacerItem(10, 10, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding))

        self.sidebar_status = QLabel("状态：就绪", self.sidebar)
        self.sidebar_status.setObjectName("dimLabel")
        layout.addWidget(self.sidebar_status, 0)

    def _switch_route(self, route: str):
        if route == self._active_route:
            return
        self._active_route = route
        self.page_title.setText(route)
        self.sidebar_status.setText(f"状态：{route}")

        for name, btn in self._nav_buttons.items():
            active = name == route
            btn.setProperty("active", active)
            btn.style().unpolish(btn)
            btn.style().polish(btn)

        if route not in self._pages:
            page = self._route_builders[route]()
            idx = self.page_stack.addWidget(page)
            self._pages[route] = page
            self._route_indexes[route] = idx

        self.page_stack.setCurrentIndex(self._route_indexes[route])
        self._append_log(f"路由切换 -> {route}")

    def _toggle_log_drawer(self):
        self._log_collapsed = not self._log_collapsed
        self.log_box.setVisible(not self._log_collapsed)
        self.log_toggle_btn.setText("展开" if self._log_collapsed else "收起")

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

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._resize_timer.start()

    def _on_resize_stable(self):
        self.window_size_label.setText(f"窗口: {self.width()} × {self.height()}")
        self._render_preview()

    def _apply_style(self):
        self.setStyleSheet(build_stylesheet())


def _configure_qt_acceleration():
    QApplication.setAttribute(Qt.ApplicationAttribute.AA_UseDesktopOpenGL, True)
    QApplication.setAttribute(Qt.ApplicationAttribute.AA_ShareOpenGLContexts, True)
    fmt = QSurfaceFormat()
    fmt.setRenderableType(QSurfaceFormat.RenderableType.OpenGL)
    fmt.setSwapBehavior(QSurfaceFormat.SwapBehavior.DoubleBuffer)
    fmt.setSwapInterval(1)
    fmt.setSamples(4)
    QSurfaceFormat.setDefaultFormat(fmt)


def main():
    _configure_qt_acceleration()
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
