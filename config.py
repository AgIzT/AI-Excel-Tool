# -*- coding: utf-8 -*-
"""
config.py
集中存放模型、接口地址、运行常量。改模型/改地址只动这里一处。
"""

import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ─────────────────────────────────────────────
# 智谱 GLM 接口（OpenAI 兼容端点，OCR 提取与对话共用一家）
# ─────────────────────────────────────────────
GLM_BASE_URL = "https://open.bigmodel.cn/api/paas/v4/"

# 单据提取所用视觉模型。
# 备选模型 id（若默认报「模型不存在」，把下面这行换成其一即可）：
#   glm-4.6v        —— 推荐，质量最好，输入¥1/输出¥3 每百万 token
#   glm-4.6v-flash  —— 免费、轻量、偏弱
#   glm-4v-flash    —— 老版兜底，确定可用
EXTRACT_MODEL = "glm-4.6v"
EXTRACT_TEMPERATURE = 0.0
EXTRACT_MAX_TOKENS = 4096

# AI 对话助手模型
CHAT_MODEL = "glm-5"

# 环境变量兜底名（显式传入的 Key 优先级最高）
GLM_ENV_NAMES = ("ZHIPU_API_KEY", "GLM_API_KEY")

# ─────────────────────────────────────────────
# 文件与模板
# ─────────────────────────────────────────────
IMG_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".webp", ".tiff"}
EXCEL_EXTS = {".xlsx", ".xls", ".csv"}
ALL_EXTS = IMG_EXTS | EXCEL_EXTS

# 模板表头所在行（0 基）。第 0 行为说明行，第 1 行为真正的字段表头。
TEMPLATE_HEADER_ROW = 1

# 提取提示词文件
PROMPT_DIR = os.path.join(BASE_DIR, "prompts")
EXTRACT_PROMPT_PATH = os.path.join(PROMPT_DIR, "extract.txt")
