# -*- coding: utf-8 -*-
import os
import sys
from PySide6.QtCore import QDateTime, QMimeData, QSettings, QThread, QTimer, Qt, Signal
from PySide6.QtGui import QPixmap, QSurfaceFormat, QTextCursor
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QScrollArea,
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

from ai_chat_service import GLMChatAssistant
from ocr_to_excel import process_images_batch
from template_manager import TemplateManager

CLR_BG = "#0f1117"
CLR_CARD = "#1a1d27"
CLR_CARD2 = "#1e2235"
CLR_BORDER = "#2a2d3e"
CLR_ACCENT = "#4f8ef7"
CLR_ACCENT2 = "#7c6af7"
CLR_SUCCESS = "#22c55e"
CLR_WARNING = "#f59e0b"
CLR_ERROR = "#ef4444"
CLR_TEXT = "#e2e8f0"
CLR_TEXT_DIM = "#64748b"

IMG_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".webp", ".tiff"}
EXCEL_EXTS = {".xlsx", ".xls", ".csv"}
ALL_EXTS = IMG_EXTS | EXCEL_EXTS


class BatchWorker(QThread):
    progress_update = Signal(int, int)
    log_message = Signal(str)
    task_finished = Signal(list)
    task_failed = Signal(str)

    def __init__(self, image_paths, output_dir, handwriting, merge_output, template_path, glm_api_key, deepseek_api_key):
        super().__init__()
        self._image_paths = list(image_paths)
        self._output_dir = output_dir
        self._handwriting = bool(handwriting)
        self._merge_output = bool(merge_output)
        self._template_path = template_path
        self._glm_api_key = (glm_api_key or "").strip()
        self._deepseek_api_key = (deepseek_api_key or "").strip()

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
                zhipu_api_key=self._glm_api_key,
                deepseek_api_key=self._deepseek_api_key,
            )
            self.task_finished.emit(result)
        except Exception as e:
            self.task_failed.emit(str(e))


class AIChatWorker(QThread):
    response_ready = Signal(str)
    task_failed = Signal(str)

    def __init__(self, send_func, user_message, api_key):
        super().__init__()
        self._send_func = send_func
        self._user_message = user_message
        self._api_key = api_key

    def run(self):
        try:
            answer = self._send_func(self._user_message, self._api_key)
            self.response_ready.emit(answer)
        except Exception as e:
            self.task_failed.emit(str(e))


class FileDropListWidget(QListWidget):
    files_dropped = Signal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragEnabled(False)
        self.setAlternatingRowColors(False)
        self.setSelectionMode(QListWidget.SelectionMode.SingleSelection)

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


class ChatInputTextEdit(QTextEdit):
    send_requested = Signal()

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            if event.modifiers() & Qt.KeyboardModifier.ShiftModifier:
                super().keyPressEvent(event)
                return
            self.send_requested.emit()
            event.accept()
            return
        super().keyPressEvent(event)


class ChatBubbleWidget(QFrame):
    def __init__(self, role: str, text: str, loading: bool = False, parent=None):
        super().__init__(parent)
        self.role = role
        self.loading = loading
        self.setObjectName("chatBubbleUser" if role == "user" else "chatBubbleAI")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(4)
        self.label = QLabel(text, self)
        self.label.setObjectName("chatBubbleText")
        self.label.setWordWrap(True)
        self.label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.label.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Preferred)
        self.label.setMaximumWidth(720)
        layout.addWidget(self.label, 0)
        self.setSizePolicy(QSizePolicy.Policy.Maximum, QSizePolicy.Policy.Preferred)
        self.setMaximumWidth(760)

    def set_message(self, text: str):
        self.label.setText(text)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings("InvoiceAI", "InvoiceAIDesktop")
        self.template_manager = TemplateManager(BASE_DIR)
        self.selected_files = []
        self._selected_file_set = set()
        self.output_excel_paths = []
        self.chat_assistant = None
        self._chat_api_key = ""
        self._active_route = ""
        self._pages = {}
        self._route_indexes = {}
        self._nav_buttons = {}
        self._worker = None
        self._ai_worker = None
        self._processing = False
        self._ai_busy = False
        self._log_buffer = []
        self._log_collapsed = True
        self._preview_path = ""
        self._preview_source = QPixmap()
        self._pending_ai_bubble = None

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
            "AI助手": self._build_ai_page,
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

        for route in ["单据处理", "AI助手", "设置"]:
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

    def _build_doc_page(self):
        page = QWidget(self.page_stack)
        layout = QHBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        left = QFrame(page)
        left.setObjectName("card")
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(16, 14, 16, 14)
        left_layout.setSpacing(8)
        layout.addWidget(left, 7)

        left_title = QLabel("文件上传与任务执行", left)
        left_title.setObjectName("sectionTitle")
        left_layout.addWidget(left_title, 0)

        file_toolbar = QWidget(left)
        file_toolbar_layout = QHBoxLayout(file_toolbar)
        file_toolbar_layout.setContentsMargins(0, 0, 0, 0)
        file_toolbar_layout.setSpacing(8)
        self.pick_files_btn = QPushButton("📂 浏览文件", file_toolbar)
        self.pick_files_btn.setObjectName("secondaryButton")
        self.pick_files_btn.clicked.connect(self._on_pick_files)
        self.clear_files_btn = QPushButton("🗑 清空列表", file_toolbar)
        self.clear_files_btn.setObjectName("secondaryButton")
        self.clear_files_btn.clicked.connect(self._on_clear_files)
        self.remove_selected_btn = QPushButton("✕ 移除选中", file_toolbar)
        self.remove_selected_btn.setObjectName("secondaryButton")
        self.remove_selected_btn.clicked.connect(self._on_remove_selected_file)
        file_toolbar_layout.addWidget(self.pick_files_btn, 0)
        file_toolbar_layout.addWidget(self.clear_files_btn, 0)
        file_toolbar_layout.addWidget(self.remove_selected_btn, 0)
        file_toolbar_layout.addItem(QSpacerItem(10, 10, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum))
        left_layout.addWidget(file_toolbar, 0)

        drop_hint = QLabel("将图片/Excel 直接拖拽到下方文件列表", left)
        drop_hint.setObjectName("dimLabel")
        left_layout.addWidget(drop_hint, 0)

        self.file_list = FileDropListWidget(left)
        self.file_list.setObjectName("fileList")
        self.file_list.files_dropped.connect(self._add_files)
        self.file_list.itemSelectionChanged.connect(self._on_file_selection_changed)
        left_layout.addWidget(self.file_list, 1)

        self.file_stats_label = QLabel("尚未选择文件", left)
        self.file_stats_label.setObjectName("dimLabel")
        left_layout.addWidget(self.file_stats_label, 0)

        action_row = QWidget(left)
        action_layout = QHBoxLayout(action_row)
        action_layout.setContentsMargins(0, 0, 0, 0)
        action_layout.setSpacing(8)
        self.start_btn = QPushButton("🚀 开始识别", action_row)
        self.start_btn.setObjectName("accentButton")
        self.start_btn.clicked.connect(self._on_start_process)
        self.open_btn = QPushButton("📊 打开输出", action_row)
        self.open_btn.setObjectName("successButton")
        self.open_btn.clicked.connect(self._on_open_output)
        action_layout.addWidget(self.start_btn, 1)
        action_layout.addWidget(self.open_btn, 1)
        left_layout.addWidget(action_row, 0)

        self.progress_bar = QProgressBar(left)
        self.progress_bar.setObjectName("progressBar")
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        left_layout.addWidget(self.progress_bar, 0)

        self.progress_label = QLabel("", left)
        self.progress_label.setObjectName("dimLabel")
        left_layout.addWidget(self.progress_label, 0)

        right = QFrame(page)
        right.setObjectName("card")
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(16, 14, 16, 14)
        right_layout.setSpacing(8)
        layout.addWidget(right, 5)

        right_title = QLabel("预览与识别选项", right)
        right_title.setObjectName("sectionTitle")
        right_layout.addWidget(right_title, 0)

        preview_card = QFrame(right)
        preview_card.setObjectName("subCard")
        preview_layout = QVBoxLayout(preview_card)
        preview_layout.setContentsMargins(10, 10, 10, 10)
        preview_layout.setSpacing(6)
        self.preview_label = QLabel("请选择图片文件进行预览", preview_card)
        self.preview_label.setObjectName("previewLabel")
        self.preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_label.setMinimumHeight(260)
        right_layout.addWidget(preview_card, 1)
        preview_layout.addWidget(self.preview_label, 1)

        options = QFrame(right)
        options.setObjectName("subCard")
        options_layout = QVBoxLayout(options)
        options_layout.setContentsMargins(14, 10, 14, 10)
        options_layout.setSpacing(8)

        tpl_row = QWidget(options)
        tpl_row_layout = QHBoxLayout(tpl_row)
        tpl_row_layout.setContentsMargins(0, 0, 0, 0)
        tpl_row_layout.setSpacing(8)
        tpl_label = QLabel("模板", tpl_row)
        tpl_label.setObjectName("dimLabel")
        self.template_combo = QComboBox(tpl_row)
        self.template_combo.setObjectName("comboBox")
        self._load_templates()
        tpl_row_layout.addWidget(tpl_label, 0)
        tpl_row_layout.addWidget(self.template_combo, 1)
        options_layout.addWidget(tpl_row, 0)

        add_tpl_btn = QPushButton("＋ 添加自定义模板", options)
        add_tpl_btn.setObjectName("secondaryButton")
        add_tpl_btn.clicked.connect(self._on_add_custom_template)
        options_layout.addWidget(add_tpl_btn, 0)

        self.handwriting_check = QCheckBox("手写体识别", options)
        self.handwriting_check.setChecked(False)
        options_layout.addWidget(self.handwriting_check, 0)

        self.merge_check = QCheckBox("合并输出到一个 Excel", options)
        self.merge_check.setChecked(False)
        options_layout.addWidget(self.merge_check, 0)

        right_layout.addWidget(options, 0)
        return page

    def _build_ai_page(self):
        page = QWidget(self.page_stack)
        layout = QHBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        left = QFrame(page)
        left.setObjectName("card")
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(16, 14, 16, 14)
        left_layout.setSpacing(10)
        layout.addWidget(left, 8)

        top_row = QWidget(left)
        top_row_layout = QHBoxLayout(top_row)
        top_row_layout.setContentsMargins(0, 0, 0, 0)
        top_row_layout.setSpacing(8)
        title = QLabel("AI 对话助手", top_row)
        title.setObjectName("sectionTitle")
        top_row_layout.addWidget(title, 1)
        self.chat_reset_btn = QPushButton("清空对话", top_row)
        self.chat_reset_btn.setObjectName("secondaryButton")
        self.chat_reset_btn.clicked.connect(self._on_reset_chat)
        self.chat_reset_btn.setFixedHeight(32)
        top_row_layout.addWidget(self.chat_reset_btn, 0)
        left_layout.addWidget(top_row, 0)

        self.chat_scroll = QScrollArea(left)
        self.chat_scroll.setObjectName("chatScroll")
        self.chat_scroll.setWidgetResizable(True)
        self.chat_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.chat_stream_widget = QWidget(self.chat_scroll)
        self.chat_stream_layout = QVBoxLayout(self.chat_stream_widget)
        self.chat_stream_layout.setContentsMargins(12, 12, 12, 12)
        self.chat_stream_layout.setSpacing(10)
        self.chat_stream_layout.addStretch(1)
        self.chat_scroll.setWidget(self.chat_stream_widget)
        left_layout.addWidget(self.chat_scroll, 1)

        input_wrap = QFrame(left)
        input_wrap.setObjectName("subCard")
        input_layout = QHBoxLayout(input_wrap)
        input_layout.setContentsMargins(10, 10, 10, 10)
        input_layout.setSpacing(8)
        self.chat_input = ChatInputTextEdit(input_wrap)
        self.chat_input.setObjectName("chatInput")
        self.chat_input.setPlaceholderText("输入问题，Enter 发送，Shift+Enter 换行")
        self.chat_input.setMinimumHeight(78)
        self.chat_input.setMaximumHeight(140)
        self.chat_input.send_requested.connect(self._on_send_chat)
        input_layout.addWidget(self.chat_input, 1)
        self.chat_send_btn = QPushButton("发送", input_wrap)
        self.chat_send_btn.setObjectName("accentButton")
        self.chat_send_btn.setFixedWidth(92)
        self.chat_send_btn.clicked.connect(self._on_send_chat)
        input_layout.addWidget(self.chat_send_btn, 0, Qt.AlignmentFlag.AlignBottom)
        left_layout.addWidget(input_wrap, 0)

        right = QFrame(page)
        right.setObjectName("card")
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(16, 14, 16, 14)
        right_layout.setSpacing(8)
        layout.addWidget(right, 4)
        tips_title = QLabel("快捷提示", right)
        tips_title.setObjectName("sectionTitle")
        right_layout.addWidget(tips_title, 0)
        for tip in ["总结最新日志", "解释识别失败原因", "如何提升OCR效果", "生成排障清单"]:
            btn = QPushButton(f"⚡ {tip}", right)
            btn.setObjectName("secondaryButton")
            btn.clicked.connect(lambda _=False, t=tip: self._insert_prompt(t))
            right_layout.addWidget(btn, 0)
        right_layout.addItem(QSpacerItem(10, 10, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding))

        self._reset_chat_ui()
        return page

    def _build_setting_page(self):
        page = QWidget(self.page_stack)
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)
        card = QFrame(page)
        card.setObjectName("card")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(16, 14, 16, 14)
        card_layout.setSpacing(10)
        title = QLabel("凭证设置", card)
        title.setObjectName("sectionTitle")
        card_layout.addWidget(title, 0)

        form = QFrame(card)
        form.setObjectName("subCard")
        form_layout = QVBoxLayout(form)
        form_layout.setContentsMargins(14, 12, 14, 12)
        form_layout.setSpacing(10)

        glm_label = QLabel("GLM API Key", form)
        glm_label.setObjectName("dimLabel")
        form_layout.addWidget(glm_label, 0)
        self.glm_key_edit = QLineEdit(form)
        self.glm_key_edit.setObjectName("keyLineEdit")
        self.glm_key_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.glm_key_edit.setPlaceholderText("请输入 GLM API Key")
        form_layout.addWidget(self.glm_key_edit, 0)

        deepseek_label = QLabel("DeepSeek API Key", form)
        deepseek_label.setObjectName("dimLabel")
        form_layout.addWidget(deepseek_label, 0)
        self.deepseek_key_edit = QLineEdit(form)
        self.deepseek_key_edit.setObjectName("keyLineEdit")
        self.deepseek_key_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.deepseek_key_edit.setPlaceholderText("请输入 DeepSeek API Key")
        form_layout.addWidget(self.deepseek_key_edit, 0)

        save_btn = QPushButton("保存配置", form)
        save_btn.setObjectName("accentButton")
        save_btn.clicked.connect(self._on_save_api_settings)
        form_layout.addWidget(save_btn, 0)

        tip = QLabel("提示：保存后将用于单据处理与AI助手请求。", form)
        tip.setObjectName("dimLabel")
        form_layout.addWidget(tip, 0)

        card_layout.addWidget(form, 0)
        card_layout.addItem(QSpacerItem(10, 10, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding))
        layout.addWidget(card, 1)
        self._load_api_settings_to_inputs()
        return page

    def _get_saved_api_keys(self):
        glm_key = str(self.settings.value("api/glm_key", "", str) or "").strip()
        deepseek_key = str(self.settings.value("api/deepseek_key", "", str) or "").strip()
        return glm_key, deepseek_key

    def _load_api_settings_to_inputs(self):
        if not hasattr(self, "glm_key_edit") or not hasattr(self, "deepseek_key_edit"):
            return
        glm_key, deepseek_key = self._get_saved_api_keys()
        self.glm_key_edit.setText(glm_key)
        self.deepseek_key_edit.setText(deepseek_key)

    def _persist_api_keys(self, glm_key: str, deepseek_key: str):
        self.settings.setValue("api/glm_key", (glm_key or "").strip())
        self.settings.setValue("api/deepseek_key", (deepseek_key or "").strip())
        self.settings.sync()

    def _require_api_keys(self, need_deepseek: bool):
        glm_key, deepseek_key = self._get_saved_api_keys()
        if not glm_key:
            self._append_log("[配置缺失] 请先在设置页配置 GLM API Key")
            QMessageBox.warning(self, "缺少凭证", "请先到【设置】页面填写并保存 GLM API Key。")
            return None, None
        if need_deepseek and not deepseek_key:
            self._append_log("[配置缺失] 请先在设置页配置 DeepSeek API Key")
            QMessageBox.warning(self, "缺少凭证", "请先到【设置】页面填写并保存 DeepSeek API Key。")
            return None, None
        return glm_key, deepseek_key

    def _on_save_api_settings(self):
        if not hasattr(self, "glm_key_edit") or not hasattr(self, "deepseek_key_edit"):
            return
        glm_key = self.glm_key_edit.text().strip()
        deepseek_key = self.deepseek_key_edit.text().strip()
        self._persist_api_keys(glm_key, deepseek_key)
        self.chat_assistant = None
        self._chat_api_key = ""
        self._append_log("[设置] API 凭证已保存")
        QMessageBox.information(self, "保存成功", "API 凭证已保存。")

    def _load_templates(self):
        names = self.template_manager.get_template_names()
        self.template_combo.clear()
        self.template_combo.addItems(names)
        default_name = self.template_manager.get_default_name()
        if default_name:
            idx = self.template_combo.findText(default_name)
            if idx >= 0:
                self.template_combo.setCurrentIndex(idx)

    def _on_add_custom_template(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "选择自定义模板文件",
            "",
            "Excel 模板 (*.xls *.xlsx);;所有文件 (*.*)",
        )
        if not path:
            return
        name = os.path.splitext(os.path.basename(path))[0]
        try:
            self.template_manager.add_custom_template(name, path)
            self._load_templates()
            idx = self.template_combo.findText(name)
            if idx >= 0:
                self.template_combo.setCurrentIndex(idx)
            self._append_log(f"[模板] 已添加自定义模板：{name}")
        except Exception as e:
            QMessageBox.critical(self, "添加模板失败", str(e))

    def _on_pick_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "选择单据文件（可多选）",
            "",
            "支持的文件 (*.jpg *.jpeg *.png *.bmp *.webp *.tiff *.xlsx *.xls *.csv);;图片文件 (*.jpg *.jpeg *.png *.bmp *.webp *.tiff);;Excel 文件 (*.xlsx *.xls *.csv);;所有文件 (*.*)",
        )
        if paths:
            self._add_files(paths)

    def _add_files(self, paths):
        added = 0
        for p in paths:
            p = os.path.abspath(p)
            if not os.path.isfile(p):
                continue
            ext = os.path.splitext(p)[-1].lower()
            if ext not in ALL_EXTS:
                continue
            if p in self._selected_file_set:
                continue
            self._selected_file_set.add(p)
            self.selected_files.append(p)
            item = QListWidgetItem(self._format_file_item_text(p))
            item.setData(Qt.ItemDataRole.UserRole, p)
            self.file_list.addItem(item)
            added += 1
        if added:
            self._append_log(f"[选择文件] 新增 {added} 个文件，共 {len(self.selected_files)} 个")
            self._update_file_stats()
            if self.file_list.currentItem() is None and self.file_list.count() > 0:
                self.file_list.setCurrentRow(self.file_list.count() - 1)

    def _format_file_item_text(self, path: str):
        ext = os.path.splitext(path)[-1].lower()
        icon = "🧾" if ext in IMG_EXTS else "📊"
        return f"{icon} {os.path.basename(path)}"

    def _on_clear_files(self):
        if self._processing:
            return
        self.file_list.clear()
        self.selected_files.clear()
        self._selected_file_set.clear()
        self._preview_path = ""
        self._preview_source = QPixmap()
        self.preview_label.setText("请选择图片文件进行预览")
        self.preview_label.setPixmap(QPixmap())
        self._update_file_stats()
        self._append_log("[清空] 已清空文件列表")

    def _on_remove_selected_file(self):
        if self._processing:
            return
        item = self.file_list.currentItem()
        if not item:
            return
        path = item.data(Qt.ItemDataRole.UserRole)
        row = self.file_list.row(item)
        self.file_list.takeItem(row)
        if path in self._selected_file_set:
            self._selected_file_set.remove(path)
        if path in self.selected_files:
            self.selected_files.remove(path)
        if self._preview_path == path:
            self._preview_path = ""
            self._preview_source = QPixmap()
            self.preview_label.setText("请选择图片文件进行预览")
            self.preview_label.setPixmap(QPixmap())
        self._update_file_stats()
        self._append_log(f"[移除] {os.path.basename(path)}")

    def _update_file_stats(self):
        n = len(self.selected_files)
        if n == 0:
            self.file_stats_label.setText("尚未选择文件")
            return
        img_n = sum(1 for p in self.selected_files if os.path.splitext(p)[-1].lower() in IMG_EXTS)
        excel_n = n - img_n
        parts = []
        if img_n:
            parts.append(f"{img_n} 张图片")
        if excel_n:
            parts.append(f"{excel_n} 个表格")
        self.file_stats_label.setText(f"已选 {n} 个文件：{'、'.join(parts)}")

    def _on_file_selection_changed(self):
        item = self.file_list.currentItem()
        if not item:
            return
        path = item.data(Qt.ItemDataRole.UserRole)
        ext = os.path.splitext(path)[-1].lower()
        if ext not in IMG_EXTS:
            self._preview_path = ""
            self._preview_source = QPixmap()
            self.preview_label.setPixmap(QPixmap())
            self.preview_label.setText("当前选中为 Excel/CSV，无法图片预览")
            return
        self._load_preview_source(path)
        self._render_preview()

    def _load_preview_source(self, path: str):
        if path == self._preview_path and not self._preview_source.isNull():
            return
        pix = QPixmap(path)
        if pix.isNull():
            self._preview_path = ""
            self._preview_source = QPixmap()
            self.preview_label.setPixmap(QPixmap())
            self.preview_label.setText("图片加载失败")
            return
        self._preview_path = path
        self._preview_source = pix

    def _render_preview(self):
        if self._preview_source.isNull():
            return
        w = max(120, self.preview_label.width() - 12)
        h = max(120, self.preview_label.height() - 12)
        scaled = self._preview_source.scaled(
            w,
            h,
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation,
        )
        self.preview_label.setText("")
        self.preview_label.setPixmap(scaled)

    def _on_start_process(self):
        if self._processing:
            return
        if not self.selected_files:
            QMessageBox.warning(self, "未选择文件", "请先选择至少一个单据文件。")
            return
        glm_key, deepseek_key = self._require_api_keys(need_deepseek=True)
        if not glm_key or not deepseek_key:
            return
        tpl_name = self.template_combo.currentText().strip()
        if not tpl_name:
            QMessageBox.warning(self, "模板缺失", "未检测到可用模板，请先添加模板。")
            return
        try:
            tpl_path = self.template_manager.get_template_path(tpl_name)
        except Exception as e:
            QMessageBox.critical(self, "模板错误", str(e))
            return

        output_dir = os.path.dirname(os.path.abspath(self.selected_files[0]))
        self._worker = BatchWorker(
            image_paths=self.selected_files,
            output_dir=output_dir,
            handwriting=self.handwriting_check.isChecked(),
            merge_output=self.merge_check.isChecked(),
            template_path=tpl_path,
            glm_api_key=glm_key,
            deepseek_api_key=deepseek_key,
        )
        self._worker.progress_update.connect(self._on_worker_progress)
        self._worker.log_message.connect(self._append_log)
        self._worker.task_finished.connect(self._on_worker_finished)
        self._worker.task_failed.connect(self._on_worker_failed)
        self._set_processing(True)
        self.progress_bar.setValue(0)
        self.progress_label.setText("初始化中...")
        self._append_log(f"[任务] 开始处理，模板={tpl_name}")
        self._worker.start()

    def _set_processing(self, processing: bool):
        self._processing = processing
        self.start_btn.setEnabled(not processing)
        self.pick_files_btn.setEnabled(not processing)
        self.clear_files_btn.setEnabled(not processing)
        self.remove_selected_btn.setEnabled(not processing)
        self.template_combo.setEnabled(not processing)
        self.handwriting_check.setEnabled(not processing)
        self.merge_check.setEnabled(not processing)
        if not self._ai_busy:
            self.sidebar_status.setText("状态：处理中" if processing else f"状态：{self._active_route}")

    def _on_worker_progress(self, current: int, total: int):
        total = max(1, total)
        pct = int((current / total) * 100)
        self.progress_bar.setValue(max(0, min(100, pct)))
        self.progress_label.setText(f"正在处理第 {current}/{total} 个文件...")

    def _on_worker_finished(self, paths):
        self._set_processing(False)
        self.output_excel_paths = list(paths)
        self.progress_bar.setValue(100)
        self.progress_label.setText("处理完成")
        self._append_log(f"[完成] 生成 {len(self.output_excel_paths)} 个输出文件")
        names = "\n".join([os.path.basename(p) for p in self.output_excel_paths[:8]])
        if len(self.output_excel_paths) > 8:
            names += f"\n... 共 {len(self.output_excel_paths)} 个"
        QMessageBox.information(self, "处理完成", f"识别任务已完成。\n\n输出文件：\n{names}")
        self._worker = None

    def _on_worker_failed(self, msg: str):
        self._set_processing(False)
        self.progress_label.setText("处理失败")
        self._append_log(f"[失败] {msg}")
        QMessageBox.critical(self, "处理失败", msg)
        self._worker = None

    def _on_open_output(self):
        if not self.output_excel_paths:
            QMessageBox.information(self, "提示", "尚未生成输出文件，请先完成识别。")
            return
        try:
            if len(self.output_excel_paths) == 1:
                os.startfile(self.output_excel_paths[0])
                self._append_log(f"[打开] {self.output_excel_paths[0]}")
                return
            parent_dir = os.path.dirname(self.output_excel_paths[0])
            os.startfile(parent_dir)
            self._append_log(f"[打开目录] {parent_dir}")
            QMessageBox.information(self, "已打开目录", "已为你打开输出目录。")
        except Exception as e:
            QMessageBox.critical(self, "打开失败", str(e))

    def _get_chat_assistant(self, api_key: str):
        api_key = (api_key or "").strip()
        if self.chat_assistant is None or self._chat_api_key != api_key:
            self.chat_assistant = GLMChatAssistant(
                api_key=api_key,
                model="glm-5",
                system_prompt="你是本软件的AI对话助手。请使用中文，回答准确、简洁、可执行。",
            )
            self._chat_api_key = api_key
        return self.chat_assistant

    def send_ai_message(self, user_message: str, api_key: str) -> str:
        assistant = self._get_chat_assistant(api_key)
        return assistant.send_message(user_message)

    def send_ai_message_async(self, user_message: str, api_key: str, on_success=None, on_error=None):
        if self._ai_worker is not None and self._ai_worker.isRunning():
            raise RuntimeError("AI助手正在处理上一条消息")
        self._ai_worker = AIChatWorker(self.send_ai_message, user_message, api_key)
        if on_success:
            self._ai_worker.response_ready.connect(on_success)
        if on_error:
            self._ai_worker.task_failed.connect(on_error)
        self._ai_worker.finished.connect(self._on_ai_worker_finished)
        self._ai_worker.start()

    def reset_ai_conversation(self, api_key: str):
        assistant = self._get_chat_assistant(api_key)
        assistant.reset()

    def _insert_prompt(self, text: str):
        if not hasattr(self, "chat_input"):
            return
        if self._ai_busy:
            return
        current = self.chat_input.toPlainText().strip()
        self.chat_input.setPlainText(f"{current}\n{text}".strip())
        self.chat_input.moveCursor(QTextCursor.MoveOperation.End)
        self.chat_input.setFocus()

    def _append_chat_bubble(self, role: str, text: str, loading: bool = False):
        row = QWidget(self.chat_stream_widget)
        row_layout = QHBoxLayout(row)
        row_layout.setContentsMargins(0, 0, 0, 0)
        row_layout.setSpacing(10)
        bubble = ChatBubbleWidget(role, text, loading=loading, parent=row)
        if role == "user":
            row_layout.addItem(QSpacerItem(10, 10, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum))
            row_layout.addWidget(bubble, 0, Qt.AlignmentFlag.AlignRight)
        else:
            row_layout.addWidget(bubble, 0, Qt.AlignmentFlag.AlignLeft)
            row_layout.addItem(QSpacerItem(10, 10, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum))
        insert_index = max(0, self.chat_stream_layout.count() - 1)
        self.chat_stream_layout.insertWidget(insert_index, row)
        QTimer.singleShot(0, self._scroll_chat_to_bottom)
        return bubble

    def _scroll_chat_to_bottom(self):
        bar = self.chat_scroll.verticalScrollBar()
        bar.setValue(bar.maximum())

    def _reset_chat_ui(self):
        while self.chat_stream_layout.count() > 1:
            item = self.chat_stream_layout.takeAt(0)
            w = item.widget()
            if w is not None:
                w.deleteLater()
        self._pending_ai_bubble = None
        self._append_chat_bubble("ai", "你好，我是 AI 对话助手。你可以问我识别流程、模板配置和异常排查。")

    def _on_send_chat(self):
        if self._ai_busy:
            return
        text = self.chat_input.toPlainText().strip()
        if not text:
            return
        glm_key, _ = self._require_api_keys(need_deepseek=False)
        if not glm_key:
            return
        self.chat_input.clear()
        self._append_chat_bubble("user", text)
        self._pending_ai_bubble = self._append_chat_bubble("ai", "正在思考...")
        self._append_log(f"[AI助手][用户] {text}")
        try:
            self.send_ai_message_async(
                text,
                api_key=glm_key,
                on_success=self._on_ai_response_success,
                on_error=self._on_ai_response_error,
            )
            self._set_ai_busy(True)
        except Exception as e:
            self._on_ai_response_error(str(e))

    def _set_ai_busy(self, busy: bool):
        self._ai_busy = busy
        self.chat_send_btn.setEnabled(not busy)
        self.chat_input.setEnabled(not busy)
        self.chat_reset_btn.setEnabled(not busy)
        if not self._processing:
            self.sidebar_status.setText("状态：AI思考中" if busy else f"状态：{self._active_route}")

    def _on_ai_response_success(self, answer: str):
        if self._pending_ai_bubble is not None:
            self._pending_ai_bubble.set_message(answer)
        else:
            self._append_chat_bubble("ai", answer)
        self._append_log(f"[AI助手] {answer}")
        self._set_ai_busy(False)

    def _on_ai_response_error(self, msg):
        err = str(msg)
        text = f"调用失败：{err}"
        if self._pending_ai_bubble is not None:
            self._pending_ai_bubble.set_message(text)
        else:
            self._append_chat_bubble("ai", text)
        self._append_log(f"[AI助手][错误] {err}")
        self._set_ai_busy(False)
        QMessageBox.critical(self, "AI 调用失败", err)

    def _on_ai_worker_finished(self):
        if self._ai_busy:
            self._set_ai_busy(False)
        self._pending_ai_bubble = None
        self._ai_worker = None

    def _on_reset_chat(self):
        if self._ai_busy:
            return
        glm_key, _ = self._require_api_keys(need_deepseek=False)
        if not glm_key:
            self._reset_chat_ui()
            return
        try:
            self.reset_ai_conversation(glm_key)
            self._reset_chat_ui()
            self._append_log("[AI助手] 对话已重置")
        except Exception as e:
            QMessageBox.critical(self, "重置失败", str(e))

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
        self.setStyleSheet(
            f"""
            QWidget {{
                background-color: {CLR_BG};
                color: {CLR_TEXT};
                font-family: "Microsoft YaHei UI";
                font-size: 13px;
            }}
            #sidebar {{
                background-color: {CLR_CARD};
                border: none;
                border-radius: 12px;
            }}
            #topBar {{
                background-color: {CLR_CARD};
                border: none;
                border-radius: 12px;
            }}
            #logDrawer {{
                background-color: {CLR_CARD};
                border: none;
            }}
            #brandCard {{
                background-color: {CLR_CARD2};
                border: none;
                border-radius: 10px;
            }}
            #brandTitle {{
                color: #ffffff;
                font-size: 20px;
                font-weight: 700;
            }}
            #pageTitle {{
                color: #ffffff;
                font-size: 22px;
                font-weight: 700;
            }}
            #sectionTitle {{
                color: {CLR_TEXT};
                font-size: 14px;
                font-weight: 700;
            }}
            #dimLabel {{
                color: {CLR_TEXT_DIM};
                font-size: 12px;
            }}
            #divider {{
                background-color: {CLR_BORDER};
                border: none;
            }}
            #card {{
                background-color: {CLR_CARD};
                border: none;
                border-radius: 12px;
            }}
            #subCard {{
                background-color: {CLR_CARD2};
                border: none;
                border-radius: 10px;
            }}
            #pageStack {{
                background: transparent;
                border: none;
            }}
            #logBox, #fileList, #comboBox {{
                background-color: {CLR_CARD2};
                border: 1px solid {CLR_BORDER};
                border-radius: 10px;
                color: #a8b4c8;
                selection-background-color: {CLR_ACCENT};
            }}
            #chatScroll {{
                background-color: {CLR_CARD2};
                border: 1px solid {CLR_BORDER};
                border-radius: 10px;
            }}
            #chatInput {{
                background-color: {CLR_CARD2};
                border: 1px solid {CLR_BORDER};
                border-radius: 8px;
                color: {CLR_TEXT};
                padding: 6px;
            }}
            #chatBubbleUser {{
                background-color: {CLR_ACCENT};
                border: none;
                border-radius: 12px;
            }}
            #chatBubbleAI {{
                background-color: #2a2f45;
                border: none;
                border-radius: 12px;
            }}
            #chatBubbleText {{
                color: #ffffff;
                font-size: 13px;
                line-height: 1.35;
            }}
            #comboBox {{
                min-height: 32px;
                padding-left: 8px;
            }}
            #previewLabel {{
                background-color: {CLR_CARD2};
                border: 1px dashed {CLR_BORDER};
                border-radius: 8px;
                color: {CLR_TEXT_DIM};
            }}
            #logToggleButton {{
                background-color: {CLR_BORDER};
                border: none;
                border-radius: 6px;
                color: {CLR_TEXT};
                font-size: 12px;
                font-weight: 600;
            }}
            #logToggleButton:hover {{
                background-color: #3a3d50;
            }}
            #navButton {{
                text-align: left;
                padding-left: 12px;
                background-color: {CLR_CARD};
                border: none;
                border-radius: 8px;
                color: {CLR_TEXT_DIM};
                font-size: 14px;
                font-weight: 700;
            }}
            #navButton:hover {{
                background-color: {CLR_CARD2};
            }}
            #navButton[active="true"] {{
                background-color: {CLR_ACCENT};
                color: #ffffff;
            }}
            QTextEdit, QListWidget {{
                background-color: {CLR_CARD2};
                border: 1px solid {CLR_BORDER};
                border-radius: 8px;
                color: {CLR_TEXT};
            }}
            QLineEdit {{
                background-color: {CLR_CARD2};
                border: 1px solid {CLR_BORDER};
                border-radius: 8px;
                color: {CLR_TEXT};
                min-height: 34px;
                padding: 0 10px;
            }}
            QProgressBar {{
                background-color: {CLR_BORDER};
                border: none;
                border-radius: 4px;
                min-height: 8px;
                max-height: 8px;
            }}
            QProgressBar::chunk {{
                background-color: {CLR_ACCENT};
                border-radius: 4px;
            }}
            QPushButton {{
                background-color: {CLR_CARD2};
                border: none;
                border-radius: 8px;
                color: {CLR_TEXT};
                min-height: 34px;
                padding: 0 12px;
                font-weight: 600;
            }}
            QPushButton:hover {{
                background-color: {CLR_BORDER};
            }}
            #accentButton {{
                background-color: {CLR_ACCENT2};
                color: #ffffff;
            }}
            #accentButton:hover {{
                background-color: #6656d8;
            }}
            #successButton {{
                background-color: {CLR_SUCCESS};
                color: #ffffff;
            }}
            #successButton:hover {{
                background-color: #1ba950;
            }}
            #secondaryButton {{
                text-align: left;
                background-color: {CLR_CARD2};
                color: {CLR_TEXT};
            }}
            QCheckBox {{
                color: {CLR_TEXT};
                spacing: 8px;
            }}
            QCheckBox::indicator {{
                width: 16px;
                height: 16px;
                border-radius: 4px;
                border: 1px solid {CLR_BORDER};
                background: {CLR_CARD2};
            }}
            QCheckBox::indicator:checked {{
                background: {CLR_ACCENT};
                border: 1px solid {CLR_ACCENT};
            }}
            """
        )


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
