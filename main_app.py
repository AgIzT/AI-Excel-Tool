# -*- coding: utf-8 -*-
"""
main_app.py
全自动单据入库系统 —— 现代化 Windows 桌面客户端
新增功能：
  · 多文件选择 + 拖拽上传（图片 / Excel）
  · 手写体识别开关
  · Excel 直接上传跳过 OCR
  · 批量处理 + 合并/分别输出选项
"""

import os
import sys
import threading
import datetime
import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk

# ── 拖拽支持（可选，未安装时降级为仅点击选择）──────────
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False

# ─────────────────────────────────────────────
# 全局主题配置
# ─────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 将项目目录提前加入 sys.path，兼容"直接运行 .py"和"PyInstaller EXE"两种模式。
# PyInstaller 打包时需要在模块顶层看到静态引用才会将 ocr_to_excel 打进包。
_meipass = getattr(sys, "_MEIPASS", None)
for _p in filter(None, [_meipass, BASE_DIR]):
    if _p not in sys.path:
        sys.path.insert(0, _p)

# 静态引用：让 PyInstaller 依赖分析能发现并打包 ocr_to_excel / template_manager
from ocr_to_excel import process_images_batch as _ocr_process_batch      # noqa: E402
from template_manager import TemplateManager as _TemplateManager          # noqa: E402

# ─────────────────────────────────────────────
# 颜色 & 字体常量
# ─────────────────────────────────────────────
CLR_BG          = "#0f1117"
CLR_CARD        = "#1a1d27"
CLR_CARD2       = "#1e2235"
CLR_BORDER      = "#2a2d3e"
CLR_ACCENT      = "#4f8ef7"
CLR_ACCENT2     = "#7c6af7"
CLR_SUCCESS     = "#22c55e"
CLR_WARNING     = "#f59e0b"
CLR_ERROR       = "#ef4444"
CLR_TEXT        = "#e2e8f0"
CLR_TEXT_DIM    = "#64748b"
CLR_TEXT_BRIGHT = "#ffffff"

FONT_TITLE  = ("Microsoft YaHei UI", 22, "bold")
FONT_HEADER = ("Microsoft YaHei UI", 13, "bold")
FONT_BODY   = ("Microsoft YaHei UI", 12)
FONT_SMALL  = ("Microsoft YaHei UI", 10)
FONT_MONO   = ("Consolas", 11)

# 支持的文件类型
IMG_EXTS   = {".jpg", ".jpeg", ".png", ".bmp", ".webp", ".tiff"}
EXCEL_EXTS = {".xlsx", ".xls"}
ALL_EXTS   = IMG_EXTS | EXCEL_EXTS


# ─────────────────────────────────────────────
# 主应用窗口
# ─────────────────────────────────────────────
class App(ctk.CTk if not HAS_DND else TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        # ── 状态变量 ──────────────────────────
        self.selected_files: list  = []   # 已选文件路径列表
        self.output_excel_paths: list = []
        self.is_processing = False

        # 模板管理器
        self.template_manager = _TemplateManager(BASE_DIR)

        # 选项变量
        self.handwriting_var = ctk.BooleanVar(value=False)
        self.merge_var       = ctk.BooleanVar(value=False)
        self.template_var    = ctk.StringVar(value=self.template_manager.get_default_name())

        self._build_window()
        self._build_layout()

        # 注册拖拽（需要 tkinterdnd2）
        if HAS_DND:
            self._register_drop()

    # ── 窗口基础设置 ────────────────────────────
    def _build_window(self):
        self.title("全自动单据入库系统")
        self.geometry("1060x760")
        self.minsize(900, 640)
        self.configure(bg=CLR_BG)

        self.update_idletasks()
        w = self.winfo_screenwidth()
        h = self.winfo_screenheight()
        x = (w - 1060) // 2
        y = (h - 760) // 2
        self.geometry(f"1060x760+{x}+{y}")

        try:
            self.iconbitmap(os.path.join(BASE_DIR, "icon.ico"))
        except Exception:
            pass

    # ── 整体布局 ────────────────────────────────
    def _build_layout(self):
        self._build_header()

        content = ctk.CTkFrame(self, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        content.columnconfigure(0, weight=1)
        content.columnconfigure(1, weight=1)
        content.rowconfigure(0, weight=1)

        self._build_left_panel(content)
        self._build_right_panel(content)
        self._build_statusbar()

    # ── 顶部标题栏 ──────────────────────────────
    def _build_header(self):
        header = ctk.CTkFrame(self, fg_color=CLR_CARD, corner_radius=0, height=70)
        header.pack(fill="x")
        header.pack_propagate(False)

        logo_frame = ctk.CTkFrame(header, fg_color="transparent")
        logo_frame.pack(side="left", padx=25, pady=10)

        ctk.CTkLabel(
            logo_frame, text=" AI ",
            font=("Microsoft YaHei UI", 14, "bold"),
            fg_color=CLR_ACCENT, text_color=CLR_TEXT_BRIGHT,
            corner_radius=6, width=40, height=32,
        ).pack(side="left", padx=(0, 12))

        title_frame = ctk.CTkFrame(logo_frame, fg_color="transparent")
        title_frame.pack(side="left")

        ctk.CTkLabel(
            title_frame, text="全自动单据入库系统",
            font=FONT_TITLE, text_color=CLR_TEXT_BRIGHT,
        ).pack(anchor="w")

        ctk.CTkLabel(
            title_frame, text="GLM-OCR  ×  DeepSeek  →  Excel  |  支持批量 / 拖拽 / 手写体",
            font=FONT_SMALL, text_color=CLR_TEXT_DIM,
        ).pack(anchor="w")

        ctk.CTkLabel(
            header, text="v2.0",
            font=FONT_SMALL, text_color=CLR_TEXT_DIM,
        ).pack(side="right", padx=25)

    # ── 左列：操作面板 ──────────────────────────
    def _build_left_panel(self, parent):
        left = ctk.CTkFrame(parent, fg_color=CLR_CARD, corner_radius=16)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=10)

        self._section_label(left, "⚙  操作控制台")

        self.steps_frame = steps_frame = ctk.CTkScrollableFrame(left, fg_color="transparent")
        steps_frame.pack(fill="both", expand=True, padx=16, pady=(0, 10))

        # ── 步骤 1：选择文件（多选 + 拖拽）──────
        step1_card = ctk.CTkFrame(steps_frame, fg_color=CLR_CARD2, corner_radius=12)
        step1_card.pack(fill="x", pady=(0, 6))

        step1_inner = ctk.CTkFrame(step1_card, fg_color="transparent")
        step1_inner.pack(fill="x", padx=14, pady=12)

        # 步骤号
        ctk.CTkLabel(
            step1_inner, text="01",
            font=("Microsoft YaHei UI", 13, "bold"),
            fg_color=CLR_ACCENT, text_color=CLR_TEXT_BRIGHT,
            corner_radius=8, width=36, height=36,
        ).pack(side="left", padx=(0, 12))

        # 文字说明
        txt_f = ctk.CTkFrame(step1_inner, fg_color="transparent")
        txt_f.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(
            txt_f, text="选择单据文件（支持多选）",
            font=("Microsoft YaHei UI", 12, "bold"),
            text_color=CLR_TEXT, anchor="w",
        ).pack(anchor="w")
        ctk.CTkLabel(
            txt_f, text="图片(JPG/PNG/BMP/WEBP) 或 Excel(XLSX/XLS)，可多选或拖拽",
            font=FONT_SMALL, text_color=CLR_TEXT_DIM, anchor="w", wraplength=220,
        ).pack(anchor="w")

        # 按钮区
        btn_col = ctk.CTkFrame(step1_inner, fg_color="transparent")
        btn_col.pack(side="right")

        ctk.CTkButton(
            btn_col, text="📂  浏览文件",
            font=FONT_BODY, width=120, height=34,
            fg_color=CLR_ACCENT, hover_color=self._darken(CLR_ACCENT),
            text_color=CLR_TEXT_BRIGHT, corner_radius=8,
            command=self._on_select_files,
        ).pack(pady=(0, 4))

        ctk.CTkButton(
            btn_col, text="🗑  清空列表",
            font=FONT_SMALL, width=120, height=28,
            fg_color=CLR_BORDER, hover_color="#3a3d50",
            text_color=CLR_TEXT_DIM, corner_radius=6,
            command=self._on_clear_files,
        ).pack()

        # 拖拽提示区
        self.drop_zone = ctk.CTkFrame(
            steps_frame, fg_color=CLR_CARD2, corner_radius=10,
            border_width=2, border_color=CLR_BORDER,
        )
        self.drop_zone.pack(fill="x", pady=(0, 4))

        self.drop_label = ctk.CTkLabel(
            self.drop_zone,
            text="🖱  将图片或 Excel 文件拖拽到此处\n（或点击上方【浏览文件】多选）",
            font=FONT_SMALL, text_color=CLR_TEXT_DIM,
        )
        self.drop_label.pack(pady=14)

        # 已选文件列表
        self.file_list_frame = ctk.CTkScrollableFrame(
            steps_frame, fg_color=CLR_CARD2, corner_radius=8, height=100,
        )
        self.file_list_frame.pack(fill="x", pady=(0, 4))

        self.file_count_label = ctk.CTkLabel(
            steps_frame, text="尚未选择文件",
            font=FONT_SMALL, text_color=CLR_TEXT_DIM, anchor="w",
        )
        self.file_count_label.pack(fill="x", padx=4, pady=(0, 8))

        # 分割线
        ctk.CTkFrame(steps_frame, fg_color=CLR_BORDER, height=1).pack(fill="x", pady=6)

        # ── 识别选项区 ──────────────────────────
        self._section_label(steps_frame, "🔧  识别选项", font=FONT_SMALL, padx=4)

        opt_frame = ctk.CTkFrame(steps_frame, fg_color=CLR_CARD2, corner_radius=10)
        opt_frame.pack(fill="x", pady=(0, 8))

        # ── 模板选择 ──────────────────────────
        self._build_template_selector(opt_frame)

        # 手写体识别开关
        hw_row = ctk.CTkFrame(opt_frame, fg_color="transparent")
        hw_row.pack(fill="x", padx=14, pady=(10, 6))

        ctk.CTkLabel(
            hw_row, text="✍  手写体识别",
            font=FONT_BODY, text_color=CLR_TEXT, anchor="w",
        ).pack(side="left", fill="x", expand=True)

        self.hw_switch = ctk.CTkSwitch(
            hw_row, text="",
            variable=self.handwriting_var,
            onvalue=True, offvalue=False,
            fg_color=CLR_BORDER, progress_color=CLR_ACCENT2,
            command=self._on_hw_toggle,
        )
        self.hw_switch.pack(side="right")

        self.hw_desc_label = ctk.CTkLabel(
            opt_frame,
            text="关闭：仅识别印刷文字",
            font=FONT_SMALL, text_color=CLR_TEXT_DIM, anchor="w",
        )
        self.hw_desc_label.pack(fill="x", padx=14, pady=(0, 10))

        # 输出模式
        merge_row = ctk.CTkFrame(opt_frame, fg_color="transparent")
        merge_row.pack(fill="x", padx=14, pady=(0, 6))

        ctk.CTkLabel(
            merge_row, text="📦  合并输出到一个 Excel",
            font=FONT_BODY, text_color=CLR_TEXT, anchor="w",
        ).pack(side="left", fill="x", expand=True)

        ctk.CTkSwitch(
            merge_row, text="",
            variable=self.merge_var,
            onvalue=True, offvalue=False,
            fg_color=CLR_BORDER, progress_color=CLR_SUCCESS,
        ).pack(side="right")

        ctk.CTkLabel(
            opt_frame,
            text="关闭：每张图片/文件分别生成独立 Excel",
            font=FONT_SMALL, text_color=CLR_TEXT_DIM, anchor="w",
        ).pack(fill="x", padx=14, pady=(0, 10))

        # 分割线
        ctk.CTkFrame(steps_frame, fg_color=CLR_BORDER, height=1).pack(fill="x", pady=6)

        # ── 步骤 2：开始处理 ────────────────────
        self._build_step_card(
            steps_frame,
            step_num="02",
            title="AI 识别并生成 Excel",
            desc="GLM-OCR 识图 → DeepSeek 匹配 → 导出标准表格",
            color=CLR_ACCENT2,
            action=self._on_start_process,
            btn_text="🚀  开始识别",
        )

        # 进度条
        self.progress_bar = ctk.CTkProgressBar(
            steps_frame, height=6, corner_radius=3,
            fg_color=CLR_BORDER, progress_color=CLR_ACCENT,
        )
        self.progress_bar.set(0)
        self.progress_bar.pack(fill="x", padx=4, pady=(4, 0))

        self.progress_label = ctk.CTkLabel(
            steps_frame, text="",
            font=FONT_SMALL, text_color=CLR_TEXT_DIM, anchor="w",
        )
        self.progress_label.pack(fill="x", padx=4, pady=(2, 8))

        # 分割线
        ctk.CTkFrame(steps_frame, fg_color=CLR_BORDER, height=1).pack(fill="x", pady=6)

        # ── 步骤 3：打开文件 ────────────────────
        self._build_step_card(
            steps_frame,
            step_num="03",
            title="打开输出文件",
            desc="用 Excel 查看已生成的标准进货单",
            color=CLR_SUCCESS,
            action=self._on_open_output,
            btn_text="📊  打开 Excel",
        )

        # 处理历史
        ctk.CTkFrame(steps_frame, fg_color=CLR_BORDER, height=1).pack(fill="x", pady=6)
        self._section_label(steps_frame, "📋  处理历史", font=FONT_SMALL, padx=4)

        self.history_frame = ctk.CTkFrame(steps_frame, fg_color="transparent")
        self.history_frame.pack(fill="x", padx=4)

    # ── 右列：预览 + 日志 ──────────────────────
    def _build_right_panel(self, parent):
        right = ctk.CTkFrame(parent, fg_color="transparent")
        right.grid(row=0, column=1, sticky="nsew", padx=(8, 0), pady=10)
        right.rowconfigure(0, weight=2)
        right.rowconfigure(1, weight=3)
        right.columnconfigure(0, weight=1)

        # 上：图片预览
        preview_card = ctk.CTkFrame(right, fg_color=CLR_CARD, corner_radius=16)
        preview_card.grid(row=0, column=0, sticky="nsew", pady=(0, 8))

        self._section_label(preview_card, "🖼  图片预览（最新选中）")

        self.preview_frame = ctk.CTkFrame(
            preview_card, fg_color=CLR_CARD2, corner_radius=10
        )
        self.preview_frame.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        self.preview_label = ctk.CTkLabel(
            self.preview_frame,
            text="请先选择单据图片\n\n支持拖拽或点击按钮上传",
            font=FONT_BODY, text_color=CLR_TEXT_DIM,
        )
        self.preview_label.pack(expand=True)

        # 下：实时日志
        log_card = ctk.CTkFrame(right, fg_color=CLR_CARD, corner_radius=16)
        log_card.grid(row=1, column=0, sticky="nsew", pady=(8, 0))

        self._section_label(log_card, "📡  实时处理日志")

        self.log_box = ctk.CTkTextbox(
            log_card, font=FONT_MONO,
            fg_color=CLR_CARD2, text_color="#a8b4c8",
            corner_radius=10, wrap="word", state="disabled",
        )
        self.log_box.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        log_toolbar = ctk.CTkFrame(log_card, fg_color="transparent")
        log_toolbar.pack(fill="x", padx=16, pady=(0, 12))

        ctk.CTkButton(
            log_toolbar, text="清空日志",
            font=FONT_SMALL, width=90, height=28,
            fg_color=CLR_BORDER, hover_color="#3a3d50",
            text_color=CLR_TEXT_DIM, corner_radius=6,
            command=self._clear_log,
        ).pack(side="right")

    # ── 底部状态栏 ──────────────────────────────
    def _build_statusbar(self):
        bar = ctk.CTkFrame(self, fg_color=CLR_CARD, corner_radius=0, height=32)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)

        self.status_dot = ctk.CTkLabel(
            bar, text="●", font=FONT_SMALL, text_color=CLR_TEXT_DIM
        )
        self.status_dot.pack(side="left", padx=(16, 4))

        self.status_text = ctk.CTkLabel(
            bar, text="就绪", font=FONT_SMALL, text_color=CLR_TEXT_DIM
        )
        self.status_text.pack(side="left")

        dnd_hint = "  |  支持拖拽文件" if HAS_DND else "  |  提示：安装 tkinterdnd2 可启用拖拽"
        ctk.CTkLabel(
            bar, text=dnd_hint,
            font=FONT_SMALL, text_color=CLR_TEXT_DIM,
        ).pack(side="right", padx=16)

    # ── 通用辅助：节标题 ────────────────────────
    def _section_label(self, parent, text, font=FONT_HEADER, padx=16):
        ctk.CTkLabel(
            parent, text=text, font=font,
            text_color=CLR_TEXT, anchor="w",
        ).pack(fill="x", padx=padx, pady=(16, 8))

    # ── 通用辅助：步骤卡片 ──────────────────────
    def _build_step_card(self, parent, step_num, title, desc, color, action, btn_text):
        card = ctk.CTkFrame(parent, fg_color=CLR_CARD2, corner_radius=12)
        card.pack(fill="x", pady=(0, 6))

        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=14, pady=12)

        ctk.CTkLabel(
            inner, text=step_num,
            font=("Microsoft YaHei UI", 13, "bold"),
            fg_color=color, text_color=CLR_TEXT_BRIGHT,
            corner_radius=8, width=36, height=36,
        ).pack(side="left", padx=(0, 12))

        text_frame = ctk.CTkFrame(inner, fg_color="transparent")
        text_frame.pack(side="left", fill="x", expand=True)

        ctk.CTkLabel(
            text_frame, text=title,
            font=("Microsoft YaHei UI", 12, "bold"),
            text_color=CLR_TEXT, anchor="w",
        ).pack(anchor="w")

        ctk.CTkLabel(
            text_frame, text=desc,
            font=FONT_SMALL, text_color=CLR_TEXT_DIM,
            anchor="w", wraplength=200,
        ).pack(anchor="w")

        btn = ctk.CTkButton(
            inner, text=btn_text,
            font=FONT_BODY, width=130, height=36,
            fg_color=color, hover_color=self._darken(color),
            text_color=CLR_TEXT_BRIGHT, corner_radius=8,
            command=action,
        )
        btn.pack(side="right")
        return btn

    def _darken(self, hex_color, factor=0.75):
        hex_color = hex_color.lstrip("#")
        r, g, b = (int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        return "#{:02x}{:02x}{:02x}".format(
            int(r * factor), int(g * factor), int(b * factor)
        )

    # ── 模板选择器 UI ────────────────────────────
    def _build_template_selector(self, parent):
        """模板下拉选择行 + 添加自定义模板按钮"""
        tpl_row = ctk.CTkFrame(parent, fg_color="transparent")
        tpl_row.pack(fill="x", padx=14, pady=(10, 2))

        ctk.CTkLabel(
            tpl_row, text="📋  导出模板",
            font=FONT_BODY, text_color=CLR_TEXT, anchor="w",
        ).pack(side="left", fill="x", expand=True)

        names = self.template_manager.get_template_names()
        if not names:
            names = ["（无可用模板）"]

        self.template_menu = ctk.CTkOptionMenu(
            tpl_row,
            values=names,
            variable=self.template_var,
            font=FONT_SMALL,
            width=168, height=28,
            fg_color=CLR_CARD,
            button_color=CLR_ACCENT,
            button_hover_color=self._darken(CLR_ACCENT),
            dropdown_fg_color=CLR_CARD2,
            dropdown_text_color=CLR_TEXT,
            text_color=CLR_TEXT,
            command=self._on_template_change,
        )
        self.template_menu.pack(side="right")

        # 添加自定义模板按钮（右对齐，轻量样式）
        add_row = ctk.CTkFrame(parent, fg_color="transparent")
        add_row.pack(fill="x", padx=14, pady=(2, 8))

        ctk.CTkButton(
            add_row, text="＋  添加自定义模板",
            font=FONT_SMALL, height=22, width=140,
            fg_color="transparent", hover_color=CLR_CARD,
            text_color=CLR_ACCENT, corner_radius=4,
            command=self._on_add_custom_template,
        ).pack(side="right")

        # 分隔线
        ctk.CTkFrame(parent, fg_color=CLR_BORDER, height=1).pack(
            fill="x", padx=14, pady=(0, 4)
        )

    def _on_template_change(self, name: str):
        """模板下拉切换回调"""
        self._log(f"[模板] 已切换到：{name}", color="info")
        self._set_status(f"模板：{name}", color=CLR_TEXT_DIM)

    def _on_add_custom_template(self):
        """添加自定义模板：弹出文件对话框，保存到 TemplateManager"""
        path = filedialog.askopenfilename(
            title="选择自定义模板文件",
            filetypes=[
                ("Excel 模板", "*.xls *.xlsx"),
                ("所有文件", "*.*"),
            ],
        )
        if not path:
            return

        name = os.path.splitext(os.path.basename(path))[0]
        existing = self.template_manager.get_all_templates()
        if name in existing:
            if not messagebox.askyesno(
                "重复模板", f"已存在同名模板「{name}」，是否覆盖？"
            ):
                return

        try:
            self.template_manager.add_custom_template(name, path)
        except Exception as e:
            messagebox.showerror("添加失败", str(e))
            return

        names = self.template_manager.get_template_names()
        self.template_menu.configure(values=names)
        self.template_var.set(name)
        self._log(f"[模板] 已添加自定义模板：{name}", color="success")
        messagebox.showinfo("添加成功", f"自定义模板「{name}」已添加！\n下次启动后仍会保留。")

    # ─────────────────────────────────────────────
    # 拖拽注册
    # ─────────────────────────────────────────────
    def _register_drop(self):
        """注册整个窗口和拖拽区域的 drop 事件"""
        try:
            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self._on_drop)
            self.drop_zone.drop_target_register(DND_FILES)
            self.drop_zone.dnd_bind("<<Drop>>", self._on_drop)
            # 更新提示文字
            self.drop_label.configure(
                text="🖱  将图片或 Excel 文件拖拽到此处\n（支持多文件同时拖入）",
                text_color=CLR_ACCENT,
            )
        except Exception:
            pass

    def _on_drop(self, event):
        """处理拖拽事件"""
        raw = event.data
        # tkinterdnd2 返回的路径可能用 {} 包裹（含空格时）
        paths = self._parse_drop_paths(raw)
        valid = [p for p in paths if os.path.splitext(p)[-1].lower() in ALL_EXTS]
        if not valid:
            self._log("[拖拽] 未检测到支持的文件格式（图片或 Excel）", color="warn")
            return
        self._add_files(valid)

    @staticmethod
    def _parse_drop_paths(raw: str) -> list:
        """解析 tkinterdnd2 返回的路径字符串"""
        raw = raw.strip()
        paths = []
        # 处理 {path with spaces} 格式
        import re
        tokens = re.findall(r'\{([^}]+)\}|(\S+)', raw)
        for t in tokens:
            p = t[0] if t[0] else t[1]
            if p:
                paths.append(p)
        return paths

    # ─────────────────────────────────────────────
    # 文件管理
    # ─────────────────────────────────────────────
    def _on_select_files(self):
        """多选文件对话框"""
        paths = filedialog.askopenfilenames(
            title="选择单据文件（可多选）",
            filetypes=[
                ("支持的文件", "*.jpg *.jpeg *.png *.bmp *.webp *.tiff *.xlsx *.xls"),
                ("图片文件",   "*.jpg *.jpeg *.png *.bmp *.webp *.tiff"),
                ("Excel 文件", "*.xlsx *.xls"),
                ("所有文件",   "*.*"),
            ],
        )
        if paths:
            self._add_files(list(paths))

    def _add_files(self, paths: list):
        """将文件添加到列表（去重）"""
        added = 0
        for p in paths:
            if p not in self.selected_files:
                self.selected_files.append(p)
                added += 1

        self._refresh_file_list()
        self._log(f"[选择文件] 新增 {added} 个文件，共 {len(self.selected_files)} 个", color="info")
        self._set_status(f"已选 {len(self.selected_files)} 个文件", color=CLR_SUCCESS)

        # 预览最后一张图片
        for p in reversed(paths):
            if os.path.splitext(p)[-1].lower() in IMG_EXTS:
                self._show_preview(p)
                break

    def _on_clear_files(self):
        """清空文件列表"""
        self.selected_files.clear()
        self._refresh_file_list()
        self.file_count_label.configure(text="尚未选择文件", text_color=CLR_TEXT_DIM)
        self.preview_label.configure(
            text="请先选择单据图片\n\n支持拖拽或点击按钮上传",
            image=None,
        )
        self._set_status("就绪", color=CLR_TEXT_DIM)
        self._log("[清空] 已清空文件列表")

    def _refresh_file_list(self):
        """刷新文件列表 UI"""
        # 清空旧列表
        for w in self.file_list_frame.winfo_children():
            w.destroy()

        for i, path in enumerate(self.selected_files):
            fname = os.path.basename(path)
            ext   = os.path.splitext(path)[-1].lower()
            icon  = "📊" if ext in EXCEL_EXTS else "🖼"
            color = CLR_WARNING if ext in EXCEL_EXTS else CLR_TEXT

            row = ctk.CTkFrame(self.file_list_frame, fg_color="transparent")
            row.pack(fill="x", pady=1)

            ctk.CTkLabel(
                row, text=f"{icon} {i+1:02d}. {fname}",
                font=FONT_SMALL, text_color=color, anchor="w",
            ).pack(side="left", fill="x", expand=True)

            # 删除单个文件按钮
            p = path
            ctk.CTkButton(
                row, text="✕",
                font=FONT_SMALL, width=24, height=20,
                fg_color="transparent", hover_color=CLR_ERROR,
                text_color=CLR_TEXT_DIM, corner_radius=4,
                command=lambda fp=p: self._remove_file(fp),
            ).pack(side="right", padx=2)

        n = len(self.selected_files)
        if n == 0:
            self.file_count_label.configure(text="尚未选择文件", text_color=CLR_TEXT_DIM)
        else:
            img_cnt   = sum(1 for p in self.selected_files if os.path.splitext(p)[-1].lower() in IMG_EXTS)
            excel_cnt = n - img_cnt
            parts = []
            if img_cnt:   parts.append(f"{img_cnt} 张图片")
            if excel_cnt: parts.append(f"{excel_cnt} 个 Excel")
            self.file_count_label.configure(
                text=f"✅  已选 {n} 个文件：{'、'.join(parts)}",
                text_color=CLR_SUCCESS,
            )


    def _remove_file(self, path: str):
        """从列表中移除单个文件"""
        if path in self.selected_files:
            self.selected_files.remove(path)
        self._refresh_file_list()
        self._set_status(f"已选 {len(self.selected_files)} 个文件", color=CLR_SUCCESS if self.selected_files else CLR_TEXT_DIM)

    # ─────────────────────────────────────────────
    # 手写体开关回调
    # ─────────────────────────────────────────────
    def _on_hw_toggle(self):
        """手写体识别开关回调，更新描述文字"""
        if self.handwriting_var.get():
            self.hw_desc_label.configure(
                text="开启：优先识别手写修改内容，手写与印刷冲突时以手写为准",
                text_color=CLR_ACCENT2,
            )
            self._log("[选项] 手写体识别已开启", color="info")
        else:
            self.hw_desc_label.configure(
                text="关闭：仅识别印刷文字",
                text_color=CLR_TEXT_DIM,
            )
            self._log("[选项] 手写体识别已关闭")

    # ─────────────────────────────────────────────
    # 开始处理
    # ─────────────────────────────────────────────
    def _on_start_process(self):
        """开始批量 AI 处理"""
        if self.is_processing:
            self._log("[警告] 正在处理中，请等待...", color="warn")
            return

        if not self.selected_files:
            messagebox.showwarning(
                "未选择文件",
                "请先选择至少一个单据文件！\n（图片或 Excel 均可）",
            )
            return

        self.is_processing = True
        self.output_excel_paths = []
        self._set_status("正在处理...", color=CLR_ACCENT)
        self._progress(0.03, "初始化中...")

        thread = threading.Thread(target=self._run_batch_process, daemon=True)
        thread.start()

    def _run_batch_process(self):
        """后台批量处理线程"""
        try:
            process_images_batch = _ocr_process_batch

            files       = list(self.selected_files)
            total       = len(files)
            handwriting = self.handwriting_var.get()
            merge       = self.merge_var.get()

            # 获取当前选中的模板路径
            tpl_name = self.template_var.get()
            try:
                tpl_path = self.template_manager.get_template_path(tpl_name)
            except Exception:
                tpl_path = None   # 回退到 ocr_to_excel 内置默认路径

            # 输出目录：与第一个文件同目录
            output_dir = os.path.dirname(os.path.abspath(files[0]))
            # 合并路径传 None，由后端根据 AI 提取的供应商/日期智能命名
            merged_path = None

            def log_cb(msg: str):
                self.after(0, lambda m=msg: self._log(m))

            def progress_cb(current: int, total_n: int):
                """current = 0-based index of the file just finished"""
                pct = 0.05 + 0.90 * (current / max(total_n, 1))
                label = f"正在处理第 {current+1}/{total_n} 个文件..."
                self.after(0, lambda p=pct, l=label: self._progress(p, l))

            result_paths = process_images_batch(
                image_paths=files,
                output_dir=output_dir,
                log_callback=log_cb,
                handwriting=handwriting,
                merge_output=merge,
                merged_output_path=merged_path,
                progress_callback=progress_cb,
                template_path=tpl_path,
            )

            self.output_excel_paths = result_paths
            self.after(0, lambda: self._on_batch_success(result_paths, total))

        except Exception as e:
            err_msg = str(e)
            self.after(0, lambda: self._on_process_error(err_msg))

    def _on_batch_success(self, result_paths: list, total_files: int):
        """批量处理成功回调（主线程）"""
        self.is_processing = False
        self._progress(1.0, "完成！")

        n = len(result_paths)
        names = [os.path.basename(p) for p in result_paths]

        self._log(
            f"\n🎉  批量处理完成！共处理 {total_files} 个文件，生成 {n} 个输出文件:",
            color="success",
        )
        for p in result_paths:
            self._log(f"    📄 {p}", color="success")
            self._add_history(os.path.basename(p), p)

        self._set_status(f"✅ 完成！生成 {n} 个文件", color=CLR_SUCCESS)

        summary = "\n".join(f"  • {nm}" for nm in names[:10])
        if n > 10:
            summary += f"\n  ... 共 {n} 个文件"

        messagebox.showinfo(
            "处理完成",
            f"成功处理 {total_files} 个文件！\n\n输出文件：\n{summary}\n\n点击【打开 Excel】可查看最近生成的文件。",
        )

    def _on_process_error(self, err_msg: str):
        """处理失败回调（主线程）"""
        self.is_processing = False
        self._progress(0, "")
        self._log(f"\n❌  处理失败: {err_msg}", color="error")
        self._set_status(f"❌ 失败: {err_msg[:60]}", color=CLR_ERROR)
        messagebox.showerror("处理失败", f"发生错误：\n\n{err_msg}")

    def _on_open_output(self):
        """打开最后生成的 Excel 文件"""
        if not self.output_excel_paths:
            messagebox.showinfo("提示", "尚未生成 Excel 文件，请先完成识别步骤。")
            return

        if len(self.output_excel_paths) == 1:
            p = self.output_excel_paths[0]
            if os.path.exists(p):
                os.startfile(p)
            else:
                messagebox.showerror("错误", f"文件不存在:\n{p}")
        else:
            # 多个文件：弹出选择框
            names = [os.path.basename(p) for p in self.output_excel_paths]
            win = ctk.CTkToplevel(self)
            win.title("选择要打开的文件")
            win.geometry("460x380")
            win.grab_set()
            win.configure(fg_color=CLR_BG)

            ctk.CTkLabel(
                win, text="请选择要打开的输出文件：",
                font=FONT_BODY, text_color=CLR_TEXT,
            ).pack(pady=(20, 10), padx=20)

            scroll = ctk.CTkScrollableFrame(win, fg_color=CLR_CARD, corner_radius=10)
            scroll.pack(fill="both", expand=True, padx=20, pady=(0, 10))

            for p in self.output_excel_paths:
                nm = os.path.basename(p)
                fp = p
                ctk.CTkButton(
                    scroll, text=f"📊  {nm}",
                    font=FONT_SMALL, anchor="w",
                    fg_color=CLR_CARD2, hover_color=CLR_BORDER,
                    text_color=CLR_TEXT, corner_radius=6,
                    command=lambda path=fp: (
                        os.startfile(path) if os.path.exists(path) else None
                    ),
                ).pack(fill="x", padx=8, pady=3)

            ctk.CTkButton(
                win, text="关闭",
                font=FONT_BODY, width=100,
                fg_color=CLR_BORDER, hover_color="#3a3d50",
                text_color=CLR_TEXT_DIM,
                command=win.destroy,
            ).pack(pady=10)

    # ─────────────────────────────────────────────
    # UI 辅助方法
    # ─────────────────────────────────────────────
    def _show_preview(self, image_path: str):
        """在预览区显示图片缩略图"""
        try:
            img = Image.open(image_path)
            max_w, max_h = 400, 190
            img.thumbnail((max_w, max_h), Image.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            self.preview_label.configure(image=photo, text="")
            self.preview_label._image = photo  # 防止被 GC
        except Exception as e:
            self.preview_label.configure(text=f"预览失败: {e}", image=None)

    def _log(self, msg: str, color: str = "normal"):
        """向日志框追加一行文字"""
        self.log_box.configure(state="normal")
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] {msg}\n"
        self.log_box.insert("end", line)
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _clear_log(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    def _set_status(self, text: str, color: str = CLR_TEXT_DIM):
        self.status_dot.configure(text_color=color)
        self.status_text.configure(text=text, text_color=color)

    def _progress(self, value: float, label: str = ""):
        self.progress_bar.set(value)
        self.progress_label.configure(text=label)

    def _add_history(self, filename: str, filepath: str):
        """在历史区添加一条记录"""
        ts = datetime.datetime.now().strftime("%m/%d %H:%M")

        row = ctk.CTkFrame(self.history_frame, fg_color=CLR_CARD2, corner_radius=8)
        row.pack(fill="x", pady=2)

        ctk.CTkLabel(
            row,
            text=f"📄  {filename}",
            font=FONT_SMALL, text_color=CLR_TEXT, anchor="w",
        ).pack(side="left", padx=10, pady=6, fill="x", expand=True)

        ctk.CTkLabel(
            row, text=ts,
            font=FONT_SMALL, text_color=CLR_TEXT_DIM,
        ).pack(side="right", padx=10)

        fp = filepath
        row.bind("<Button-1>", lambda e: os.startfile(fp) if os.path.exists(fp) else None)
        row.configure(cursor="hand2")


# ─────────────────────────────────────────────
# 程序入口
# ─────────────────────────────────────────────
if __name__ == "__main__":
    # ── 高 DPI 适配（防止 Windows 自动位图缩放导致文字模糊）──
    try:
        import ctypes
        # Per-Monitor DPI Aware V2（Windows 10 1703+）
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

    app = App()
    app.mainloop()
