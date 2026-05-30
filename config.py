# -*- coding: utf-8 -*-
"""
config.py
集中存放模型、接口地址、运行常量。改模型/改地址只动这里一处。
"""

import os
from dataclasses import dataclass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ─────────────────────────────────────────────
# 接口预设：base_url + 默认模型。
# 真正生效的值来自设置页（QSettings），这里只是各家「开箱即用」的默认。
# 任何 OpenAI 兼容的多模态（视觉）模型都能用，「自定义」一栏可手填。
# ─────────────────────────────────────────────
PROVIDER_PRESETS = {
    "智谱 GLM": {
        "base_url": "https://open.bigmodel.cn/api/paas/v4/",
        "model": "glm-4.6v",
    },
    "OpenAI": {
        "base_url": "https://api.openai.com/v1/",
        "model": "gpt-4o",
    },
    "通义千问": {
        "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1/",
        "model": "qwen-vl-max",
    },
    "自定义": {
        "base_url": "",
        "model": "",
    },
}

DEFAULT_PROVIDER = "智谱 GLM"

# 兜底默认（设置页留空时使用）。
# 备选 GLM 模型 id：glm-4.6v(推荐) / glm-4.6v-flash(免费偏弱) / glm-4v-flash(老版兜底)
GLM_BASE_URL = PROVIDER_PRESETS[DEFAULT_PROVIDER]["base_url"]
EXTRACT_MODEL = PROVIDER_PRESETS[DEFAULT_PROVIDER]["model"]
EXTRACT_TEMPERATURE = 0.0
EXTRACT_MAX_TOKENS = 4096

# 环境变量兜底名（设置页显式填写的 Key 优先级最高）
GLM_ENV_NAMES = ("ZHIPU_API_KEY", "GLM_API_KEY")


@dataclass(frozen=True)
class ApiConfig:
    """一次识别请求所需的接口三要素，始终一起传递。"""
    base_url: str
    model: str
    api_key: str = ""

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
