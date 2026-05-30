# -*- coding: utf-8 -*-
"""
review_dialog.py
导出前人工复核：把识别结果放进可编辑表格，用户改完再导出。
数据要进会计软件，准确性第一，所以导出前强制过一道人工核对。

用法：
    dlg = ReviewDialog(extraction, parent)
    if dlg.exec() == QDialog.DialogCode.Accepted:
        edited = dlg.get_edited_extraction()
        export_batch(edited, ...)
"""

from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QSpacerItem,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from theme import build_stylesheet


class ReviewDialog(QDialog):
    """识别结果复核对话框。exec() 返回 Accepted 后调用 get_edited_extraction() 取回编辑后的数据。"""

    def __init__(self, extraction: dict, parent=None):
        super().__init__(parent)
        self._headers = list(extraction.get("headers", []))
        self._items = list(extraction.get("items", []))
        self._failed = list(extraction.get("failed", []))

        # 与 _items 一一对应，便于回读
        self._tables = []
        self._supplier_edits = []
        self._date_edits = []

        self.setWindowTitle("导出前复核")
        self.setMinimumSize(900, 600)
        self.setStyleSheet(build_stylesheet())
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        title = QLabel("请核对识别结果，确认无误后再导出", self)
        title.setObjectName("sectionTitle")
        layout.addWidget(title, 0)

        if self._failed:
            names = "、".join(n for n, _ in self._failed)
            warn = QLabel(f"⚠ {len(self._failed)} 个文件识别失败，不会导出：{names}", self)
            warn.setObjectName("dimLabel")
            warn.setWordWrap(True)
            layout.addWidget(warn, 0)

        switch_row = QWidget(self)
        switch_layout = QHBoxLayout(switch_row)
        switch_layout.setContentsMargins(0, 0, 0, 0)
        switch_layout.setSpacing(8)
        switch_label = QLabel("当前文件", switch_row)
        switch_label.setObjectName("dimLabel")
        self.file_combo = QComboBox(switch_row)
        self.file_combo.setObjectName("comboBox")
        for it in self._items:
            self.file_combo.addItem(f"{it.get('name', '(未命名)')}（{len(it.get('records', []))} 条）")
        self.file_combo.currentIndexChanged.connect(self._on_file_changed)
        switch_layout.addWidget(switch_label, 0)
        switch_layout.addWidget(self.file_combo, 1)
        layout.addWidget(switch_row, 0)
        switch_row.setVisible(len(self._items) > 1)

        self.stack = QStackedWidget(self)
        layout.addWidget(self.stack, 1)
        for it in self._items:
            self.stack.addWidget(self._build_item_page(it))

        if not self._items:
            empty = QLabel("没有可复核的记录（全部文件识别失败）。", self)
            empty.setObjectName("dimLabel")
            self.stack.addWidget(empty)

        btn_row = QWidget(self)
        btn_layout = QHBoxLayout(btn_row)
        btn_layout.setContentsMargins(0, 0, 0, 0)
        btn_layout.setSpacing(8)
        btn_layout.addItem(QSpacerItem(10, 10, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum))
        cancel_btn = QPushButton("取消", btn_row)
        cancel_btn.setObjectName("secondaryButton")
        cancel_btn.clicked.connect(self.reject)
        self.export_btn = QPushButton("✅ 确认导出", btn_row)
        self.export_btn.setObjectName("accentButton")
        self.export_btn.clicked.connect(self._on_confirm)
        self.export_btn.setEnabled(bool(self._items))
        btn_layout.addWidget(cancel_btn, 0)
        btn_layout.addWidget(self.export_btn, 0)
        layout.addWidget(btn_row, 0)

    def _build_item_page(self, item: dict) -> QWidget:
        page = QWidget(self.stack)
        v = QVBoxLayout(page)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(8)

        meta = item.get("meta", {}) or {}
        meta_row = QWidget(page)
        meta_layout = QHBoxLayout(meta_row)
        meta_layout.setContentsMargins(0, 0, 0, 0)
        meta_layout.setSpacing(8)
        sup_label = QLabel("供应商", meta_row)
        sup_label.setObjectName("dimLabel")
        sup_edit = QLineEdit(str(meta.get("supplier", "") or ""), meta_row)
        date_label = QLabel("日期", meta_row)
        date_label.setObjectName("dimLabel")
        date_edit = QLineEdit(str(meta.get("date", "") or ""), meta_row)
        meta_layout.addWidget(sup_label, 0)
        meta_layout.addWidget(sup_edit, 2)
        meta_layout.addWidget(date_label, 0)
        meta_layout.addWidget(date_edit, 1)
        v.addWidget(meta_row, 0)
        self._supplier_edits.append(sup_edit)
        self._date_edits.append(date_edit)

        table = QTableWidget(page)
        records = item.get("records", []) or []
        table.setColumnCount(len(self._headers))
        table.setHorizontalHeaderLabels(self._headers)
        table.setRowCount(len(records))
        for r, rec in enumerate(records):
            rec = rec if isinstance(rec, dict) else {}
            for c, h in enumerate(self._headers):
                val = rec.get(h, "")
                table.setItem(r, c, QTableWidgetItem("" if val is None else str(val)))
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        table.horizontalHeader().setStretchLastSection(True)
        table.resizeColumnsToContents()
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        v.addWidget(table, 1)
        self._tables.append(table)

        tool_row = QWidget(page)
        tool_layout = QHBoxLayout(tool_row)
        tool_layout.setContentsMargins(0, 0, 0, 0)
        tool_layout.setSpacing(8)
        add_btn = QPushButton("＋ 增加行", tool_row)
        add_btn.setObjectName("secondaryButton")
        add_btn.clicked.connect(lambda _=False, t=table: self._add_row(t))
        del_btn = QPushButton("✕ 删除选中行", tool_row)
        del_btn.setObjectName("secondaryButton")
        del_btn.clicked.connect(lambda _=False, t=table: self._del_rows(t))
        tool_layout.addWidget(add_btn, 0)
        tool_layout.addWidget(del_btn, 0)
        tool_layout.addItem(QSpacerItem(10, 10, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum))
        v.addWidget(tool_row, 0)

        return page

    def _on_file_changed(self, idx: int):
        if 0 <= idx < self.stack.count():
            self.stack.setCurrentIndex(idx)

    def _add_row(self, table: QTableWidget):
        r = table.rowCount()
        table.insertRow(r)
        for c in range(table.columnCount()):
            table.setItem(r, c, QTableWidgetItem(""))
        table.scrollToBottom()
        table.setCurrentCell(r, 0)

    def _del_rows(self, table: QTableWidget):
        rows = sorted({i.row() for i in table.selectedIndexes()}, reverse=True)
        if not rows:
            cur = table.currentRow()
            if cur >= 0:
                rows = [cur]
        for r in rows:
            table.removeRow(r)

    def _read_table(self, table: QTableWidget) -> list:
        records = []
        for r in range(table.rowCount()):
            rec = {}
            has_value = False
            for c, h in enumerate(self._headers):
                cell = table.item(r, c)
                text = cell.text().strip() if cell is not None else ""
                if text:
                    has_value = True
                rec[h] = text
            if has_value:
                records.append(rec)
        return records

    def get_edited_extraction(self) -> dict:
        """把表格与供应商/日期回读为与 extract_batch 同构的 extraction。"""
        items = []
        for idx, it in enumerate(self._items):
            records = self._read_table(self._tables[idx])
            meta = {
                "supplier": self._supplier_edits[idx].text().strip(),
                "date": self._date_edits[idx].text().strip(),
            }
            items.append({
                "name": it.get("name", ""),
                "path": it.get("path", ""),
                "records": records,
                "meta": meta,
            })
        return {"headers": list(self._headers), "items": items, "failed": list(self._failed)}

    def _on_confirm(self):
        total = sum(len(self._read_table(t)) for t in self._tables)
        if total == 0:
            QMessageBox.warning(self, "无可导出数据", "当前没有任何记录，请先补充或点击取消。")
            return
        self.accept()
