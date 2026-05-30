# -*- coding: utf-8 -*-
"""
doc_page.py
「单据处理」页：文件选择/拖拽、预览、模板选择、批量识别与进度。
以 Mixin 形式提供给 MainWindow，依赖 self 上由外壳初始化的状态与日志方法。
"""

import os

from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QListWidgetItem,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QSpacerItem,
    QVBoxLayout,
    QWidget,
)

from config import IMG_EXTS, ALL_EXTS
from widgets import FileDropListWidget
from workers import BatchWorker


class DocPageMixin:
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
        api_config = self._require_api_config()
        if api_config is None:
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
            api_config=api_config,
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
