# -*- coding: utf-8 -*-
"""
theme.py
界面配色常量与 QSS 样式表。改主题只动这里一处。
"""

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


def build_stylesheet() -> str:
    """返回主窗口使用的完整 QSS。"""
    return f"""
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
