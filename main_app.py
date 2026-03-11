# -*- coding: utf-8 -*-
"""
main_app.py
全自动单据入库系统 —— 现代化 Windows 桌面客户端
基于 customtkinter 深色主题
"""

import os
import sys
import threading
import subprocess
import datetime
import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk

# ─────────────────────────────────────────────
# 全局主题配置
# ─────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ─────────────────────────────────────────────
# 颜色 & 字体常量
# ─────────────────────────────────────────────
CLR_BG          = "#0f1117"   # 最深背景
CLR_CARD        = "#1a1d27"   # 卡片背景
CLR_CARD2       = "#1e2235"   # 次级卡片
CLR_BORDER      = "#2a2d3e"   # 边框
CLR_ACCENT      = "#4f8ef7"   # 蓝色主色调
CLR_ACCENT2     = "#7c6af7"   # 紫色辅色
CLR_SUCCESS     = "#22c55e"   # 绿色
CLR_WARNING     = "#f59e0b"   # 橙黄色
CLR_ERROR       = "#ef4444"   # 红色
CLR_TEXT        = "#e2e8f0"   # 主文字
CLR_TEXT_DIM    = "#64748b"   # 暗文字
CLR_TEXT_BRIGHT = "#ffffff"   # 亮白

FONT_TITLE  = ("Microsoft YaHei UI", 22, "bold")
FONT_HEADER = ("Microsoft YaHei UI", 13, "bold")
FONT_BODY   = ("Microsoft YaHei UI", 12)
FONT_SMALL  = ("Microsoft YaHei UI", 10)
FONT_MONO   = ("Consolas", 11)


# ─────────────────────────────────────────────
# 主应用窗口
# ─────────────────────────────────────────────
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.selected_image_path = None
        self.output_excel_path   = None
        self.is_processing       = False

        self._build_window()
        self._build_layout()

    # ── 窗口基础设置 ────────────────────────────
    def _build_window(self):
        self.title("全自动单据入库系统")
        self.geometry("940x700")
        self.minsize(800, 600)
        self.configure(fg_color=CLR_BG)

        # 居中显示
        self.update_idletasks()
        w = self.winfo_screenwidth()
        h = self.winfo_screenheight()
        x = (w - 940) // 2
        y = (h - 700) // 2
        self.geometry(f"940x700+{x}+{y}")

        # 设置图标（可选，失败时忽略）
        try:
            self.iconbitmap(os.path.join(BASE_DIR, "icon.ico"))
        except Exception:
            pass

    # ── 整体布局 ────────────────────────────────
    def _build_layout(self):
        # 顶部标题栏
        self._build_header()

        # 主内容区（左+右两列）
        content = ctk.CTkFrame(self, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        content.columnconfigure(0, weight=1)
        content.columnconfigure(1, weight=1)
        content.rowconfigure(0, weight=1)

        # 左列：操作面板
        self._build_left_panel(content)
        # 右列：预览 + 日志
        self._build_right_panel(content)

        # 底部状态栏
        self._build_statusbar()

    # ── 顶部标题栏 ──────────────────────────────
    def _build_header(self):
        header = ctk.CTkFrame(self, fg_color=CLR_CARD, corner_radius=0, height=70)
        header.pack(fill="x")
        header.pack_propagate(False)

        # 左侧 LOGO 区
        logo_frame = ctk.CTkFrame(header, fg_color="transparent")
        logo_frame.pack(side="left", padx=25, pady=10)

        # 渐变色标志方块
        badge = ctk.CTkLabel(
            logo_frame,
            text=" AI ",
            font=("Microsoft YaHei UI", 14, "bold"),
            fg_color=CLR_ACCENT,
            text_color=CLR_TEXT_BRIGHT,
            corner_radius=6,
            width=40, height=32,
        )
        badge.pack(side="left", padx=(0, 12))

        title_frame = ctk.CTkFrame(logo_frame, fg_color="transparent")
        title_frame.pack(side="left")

        ctk.CTkLabel(
            title_frame,
            text="GLM-OCR  ×  DeepSeek  →  Excel",
            font=FONT_TITLE,
            text_color=CLR_TEXT_BRIGHT,
        ).pack(anchor="w")

        ctk.CTkLabel(
            title_frame,
            text="GLM-OCR  ×  DeepSeek  →  Excel",
            font=FONT_SMALL,
            text_color=CLR_TEXT_DIM,
        ).pack(anchor="w")

        # 右侧版本信息
        ctk.CTkLabel(
            header,
            text="v1.0",
            font=FONT_SMALL,
            text_color=CLR_TEXT_DIM,
        ).pack(side="right", padx=25)

    # ── 左列：操作面板 ──────────────────────────
    def _build_left_panel(self, parent):
        left = ctk.CTkFrame(parent, fg_color=CLR_CARD, corner_radius=16)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=10)

        # 节标题
        self._section_label(left, "⚙  操作控制台")

        # ── 步骤卡片 ────────────────────────────
        steps_frame = ctk.CTkScrollableFrame(left, fg_color="transparent")
        steps_frame.pack(fill="both", expand=True, padx=16, pady=(0, 10))

        # 步骤 1：选择图片
        self._build_step_card(
            steps_frame,
            step_num="01",
            title="选择单据图片",
            desc="支持 JPG / PNG / BMP / WEBP 格式",
            color=CLR_ACCENT,
            action=self._on_select_image,
            btn_text="📂  浏览文件",
        )

        # 当前选中文件展示
        self.selected_label = ctk.CTkLabel(
            steps_frame,
            text="尚未选择图片",
            font=FONT_SMALL,
            text_color=CLR_TEXT_DIM,
            anchor="w",
            wraplength=340,
        )
        self.selected_label.pack(fill="x", padx=4, pady=(0, 12))

        # 分割线
        ctk.CTkFrame(steps_frame, fg_color=CLR_BORDER, height=1).pack(fill="x", pady=8)

        # 步骤 2：开始处理
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
            steps_frame,
            height=6,
            corner_radius=3,
            fg_color=CLR_BORDER,
            progress_color=CLR_ACCENT,
        )
        self.progress_bar.set(0)
        self.progress_bar.pack(fill="x", padx=4, pady=(4, 0))

        self.progress_label = ctk.CTkLabel(
            steps_frame,
            text="",
            font=FONT_SMALL,
            text_color=CLR_TEXT_DIM,
            anchor="w",
        )
        self.progress_label.pack(fill="x", padx=4, pady=(2, 12))

        # 分割线
        ctk.CTkFrame(steps_frame, fg_color=CLR_BORDER, height=1).pack(fill="x", pady=8)

        # 步骤 3：打开文件
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
        ctk.CTkFrame(steps_frame, fg_color=CLR_BORDER, height=1).pack(fill="x", pady=8)
        self._section_label(steps_frame, "📋  处理历史", font=FONT_SMALL, padx=4)

        self.history_frame = ctk.CTkFrame(steps_frame, fg_color="transparent")
        self.history_frame.pack(fill="x", padx=4)

        self.history_labels = []

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

        self._section_label(preview_card, "🖼  图片预览")

        self.preview_frame = ctk.CTkFrame(
            preview_card, fg_color=CLR_CARD2, corner_radius=10
        )
        self.preview_frame.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        self.preview_label = ctk.CTkLabel(
            self.preview_frame,
            text="请先选择单据图片\n\n支持拖拽或点击按钮上传",
            font=FONT_BODY,
            text_color=CLR_TEXT_DIM,
        )
        self.preview_label.pack(expand=True)

        # 下：实时日志
        log_card = ctk.CTkFrame(right, fg_color=CLR_CARD, corner_radius=16)
        log_card.grid(row=1, column=0, sticky="nsew", pady=(8, 0))

        self._section_label(log_card, "📡  实时处理日志")

        self.log_box = ctk.CTkTextbox(
            log_card,
            font=FONT_MONO,
            fg_color=CLR_CARD2,
            text_color="#a8b4c8",
            corner_radius=10,
            wrap="word",
            state="disabled",
        )
        self.log_box.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        # 日志底部工具栏
        log_toolbar = ctk.CTkFrame(log_card, fg_color="transparent")
        log_toolbar.pack(fill="x", padx=16, pady=(0, 12))

        ctk.CTkButton(
            log_toolbar,
            text="清空日志",
            font=FONT_SMALL,
            width=90, height=28,
            fg_color=CLR_BORDER,
            hover_color="#3a3d50",
            text_color=CLR_TEXT_DIM,
            corner_radius=6,
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

        ctk.CTkLabel(
            bar,
            text=f"模板路径: {os.path.basename(BASE_DIR)}",
            font=FONT_SMALL,
            text_color=CLR_TEXT_DIM,
        ).pack(side="right", padx=16)

    # ── 通用辅助：节标题 ────────────────────────
    def _section_label(self, parent, text, font=FONT_HEADER, padx=16):
        ctk.CTkLabel(
            parent,
            text=text,
            font=font,
            text_color=CLR_TEXT,
            anchor="w",
        ).pack(fill="x", padx=padx, pady=(16, 8))

    # ── 通用辅助：步骤卡片 ──────────────────────
    def _build_step_card(self, parent, step_num, title, desc, color, action, btn_text):
        card = ctk.CTkFrame(parent, fg_color=CLR_CARD2, corner_radius=12)
        card.pack(fill="x", pady=(0, 10))

        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=14, pady=12)

        # 左侧步骤号徽章
        badge = ctk.CTkLabel(
            inner,
            text=step_num,
            font=("Microsoft YaHei UI", 13, "bold"),
            fg_color=color,
            text_color=CLR_TEXT_BRIGHT,
            corner_radius=8,
            width=36, height=36,
        )
        badge.pack(side="left", padx=(0, 12))

        # 中间文字
        text_frame = ctk.CTkFrame(inner, fg_color="transparent")
        text_frame.pack(side="left", fill="x", expand=True)

        ctk.CTkLabel(
            text_frame,
            text=title,
            font=("Microsoft YaHei UI", 12, "bold"),
            text_color=CLR_TEXT,
            anchor="w",
        ).pack(anchor="w")

        ctk.CTkLabel(
            text_frame,
            text=desc,
            font=FONT_SMALL,
            text_color=CLR_TEXT_DIM,
            anchor="w",
            wraplength=200,
        ).pack(anchor="w")

        # 右侧按钮
        btn = ctk.CTkButton(
            inner,
            text=btn_text,
            font=FONT_BODY,
            width=130, height=36,
            fg_color=color,
            hover_color=self._darken(color),
            text_color=CLR_TEXT_BRIGHT,
            corner_radius=8,
            command=action,
        )
        btn.pack(side="right")
        return btn

    def _darken(self, hex_color, factor=0.75):
        """将十六进制颜色变暗"""
        hex_color = hex_color.lstrip("#")
        r, g, b = (int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        return "#{:02x}{:02x}{:02x}".format(
            int(r * factor), int(g * factor), int(b * factor)
        )

    # ─────────────────────────────────────────────
    # 事件处理
    # ─────────────────────────────────────────────
    def _on_select_image(self):
        """选择单据图片"""
        path = filedialog.askopenfilename(
            title="选择单据图片",
            filetypes=[
                ("图片文件", "*.jpg *.jpeg *.png *.bmp *.webp *.tiff"),
                ("所有文件", "*.*"),
            ],
        )
        if not path:
            return

        self.selected_image_path = path
        short_name = os.path.basename(path)
        self.selected_label.configure(
            text=f"✅  {short_name}",
            text_color=CLR_SUCCESS,
        )
        self._set_status(f"已选择: {short_name}", color=CLR_SUCCESS)
        self._log(f"[选择图片] {path}", color="info")
        self._show_preview(path)

    def _on_start_process(self):
        """开始 AI 处理"""
        if self.is_processing:
            self._log("[警告] 正在处理中，请等待...", color="warn")
            return

        if not self.selected_image_path:
            messagebox.showwarning(
                "未选择图片", "请先选择一张单据图片！\n（点击【选择单据图片】按钮）"
            )
            return

        # 在后台线程运行，避免 UI 卡死
        self.is_processing = True
        self.output_excel_path = None
        self._set_status("正在处理...", color=CLR_ACCENT)
        self._progress(0.05, "初始化中...")
        thread = threading.Thread(target=self._run_process, daemon=True)
        thread.start()

    def _run_process(self):
        """后台处理线程"""
        try:
            # 动态导入（避免启动时就拉依赖）
            sys.path.insert(0, BASE_DIR)
            from ocr_to_excel import process_image

            # 确定输出路径
            img_dir  = os.path.dirname(os.path.abspath(self.selected_image_path))
            img_name = os.path.splitext(os.path.basename(self.selected_image_path))[0]
            ts       = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            out_path = os.path.join(img_dir, f"{img_name}_{ts}_入库.xlsx")

            self._progress(0.15, "读取模板表头...")

            # 步骤回调
            step_map = {
                "步骤 1/4": (0.20, "读取模板表头..."),
                "步骤 2/4": (0.40, "AI 识别图片文字（GLM-OCR）..."),
                "步骤 3/4": (0.70, "DeepSeek 语义匹配..."),
                "步骤 4/4": (0.90, "生成 Excel 文件..."),
            }

            def log_cb(msg: str):
                # 根据关键词自动推进进度条
                for key, (pct, label) in step_map.items():
                    if key in msg:
                        self.after(0, lambda p=pct, l=label: self._progress(p, l))
                        break
                self.after(0, lambda m=msg: self._log(m))

            result_path = process_image(
                self.selected_image_path,
                out_path,
                log_callback=log_cb,
            )

            self.output_excel_path = result_path
            self.after(0, self._on_process_success)

        except Exception as e:
            err_msg = str(e)
            self.after(0, lambda: self._on_process_error(err_msg))

    def _on_process_success(self):
        """处理成功回调（主线程）"""
        self.is_processing = False
        self._progress(1.0, "完成！")
        fname = os.path.basename(self.output_excel_path)
        self._log(f"\n🎉  成功！文件已生成:\n    {self.output_excel_path}", color="success")
        self._set_status(f"✅ 生成成功: {fname}", color=CLR_SUCCESS)
        self._add_history(fname, self.output_excel_path)
        messagebox.showinfo(
            "处理完成",
            f"Excel 文件已生成！\n\n文件名: {fname}\n\n点击【打开 Excel】按钮可直接查看。",
        )

    def _on_process_error(self, err_msg: str):
        """处理失败回调（主线程）"""
        self.is_processing = False
        self._progress(0, "")
        self._log(f"\n❌  处理失败: {err_msg}", color="error")
        self._set_status(f"❌ 失败: {err_msg[:50]}", color=CLR_ERROR)
        messagebox.showerror("处理失败", f"发生错误：\n\n{err_msg}")

    def _on_open_output(self):
        """打开输出的 Excel 文件"""
        if not self.output_excel_path:
            messagebox.showinfo("提示", "尚未生成 Excel 文件，请先完成识别步骤。")
            return
        if not os.path.exists(self.output_excel_path):
            messagebox.showerror("错误", f"文件不存在:\n{self.output_excel_path}")
            return
        os.startfile(self.output_excel_path)

    # ─────────────────────────────────────────────
    # UI 辅助方法
    # ─────────────────────────────────────────────
    def _show_preview(self, image_path: str):
        """在预览区显示图片缩略图"""
        try:
            img = Image.open(image_path)
            # 计算缩放比
            max_w, max_h = 380, 180
            img.thumbnail((max_w, max_h), Image.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            self.preview_label.configure(image=photo, text="")
            self.preview_label._image = photo  # 防止被 GC
        except Exception as e:
            self.preview_label.configure(
                text=f"预览失败: {e}", image=None
            )

    def _log(self, msg: str, color: str = "normal"):
        """向日志框追加一行文字"""
        color_map = {
            "normal":  "#a8b4c8",
            "info":    CLR_ACCENT,
            "success": CLR_SUCCESS,
            "warn":    CLR_WARNING,
            "error":   CLR_ERROR,
        }
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
            font=FONT_SMALL,
            text_color=CLR_TEXT,
            anchor="w",
        ).pack(side="left", padx=10, pady=6, fill="x", expand=True)

        ctk.CTkLabel(
            row,
            text=ts,
            font=FONT_SMALL,
            text_color=CLR_TEXT_DIM,
        ).pack(side="right", padx=10)

        # 点击整行打开文件
        fp = filepath
        row.bind("<Button-1>", lambda e: os.startfile(fp) if os.path.exists(fp) else None)
        row.configure(cursor="hand2")


# ─────────────────────────────────────────────
# 程序入口
# ─────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()
