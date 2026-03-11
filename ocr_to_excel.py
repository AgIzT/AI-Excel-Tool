# -*- coding: utf-8 -*-
"""
ocr_to_excel.py
自动化单据处理工具：图片 → GLM OCR → DeepSeek 语义匹配 → Excel 标准模板
"""

import os
import sys
import base64
import io
import json
import re
import requests
import pandas as pd
import xlrd
from PIL import Image

# ─────────────────────────────────────────────
# 配置区
# ─────────────────────────────────────────────
ZHIPU_API_KEY   = ""
DEEPSEEK_API_KEY = ""
DEEPSEEK_BASE_URL = ""
DEEPSEEK_MODEL   = ""

BASE_DIR      = os.path.dirname(os.path.abspath(__file__))
IMAGE_PATH    = os.path.join(BASE_DIR, "test_receipt.jpg")
TEMPLATE_PATH = os.path.join(BASE_DIR, "进货单商品导入模板.xls")
OUTPUT_PATH   = os.path.join(BASE_DIR, "标准输出测试.xlsx")

# ─────────────────────────────────────────────
# 第一步：读取模板表头
# ─────────────────────────────────────────────
def get_template_headers(xls_path: str) -> list:
    """读取 xls 模板，返回第1行（index=1）的非空表头列表"""
    wb = xlrd.open_workbook(xls_path)
    sh = wb.sheet_by_index(0)
    # 第0行是说明，第1行是真正的表头
    raw_headers = sh.row_values(1)
    headers = [h for h in raw_headers if str(h).strip()]
    print(f"[模板表头] {headers}")
    return headers


# ─────────────────────────────────────────────
# 第二步：GLM OCR —— 图片转文本
# ─────────────────────────────────────────────
def image_to_data_url(image_path: str) -> str:
    """读取图片并转换为 data:image/...;base64,... 格式的 URL"""
    ext = os.path.splitext(image_path)[-1].lower().lstrip(".")
    fmt_map = {"jpg": "jpeg", "jpeg": "jpeg", "png": "png",
               "bmp": "bmp", "gif": "gif", "webp": "webp"}
    mime = fmt_map.get(ext, "jpeg")
    with open(image_path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("utf-8")
    return f"data:image/{mime};base64,{b64}"


def ocr_with_glm_ocr(image_path: str) -> str:
    """
    使用智谱 GLM-OCR (layout_parsing 端点) 识别图片文字。
    这是专用文档解析模型，不需要 max_tokens 参数，适合单据OCR。
    """
    print("[OCR] 正在调用智谱 GLM-OCR (layout_parsing)...")
    data_url = image_to_data_url(image_path)
    url = "https://open.bigmodel.cn/api/paas/v4/layout_parsing"
    headers = {
        "Authorization": ZHIPU_API_KEY,
        "Content-Type": "application/json"
    }
    payload = {
        "model": "glm-ocr",
        "file": data_url
    }
    resp = requests.post(url, headers=headers, json=payload, timeout=120)
    if resp.status_code != 200:
        raise RuntimeError(f"GLM-OCR 调用失败: {resp.status_code} - {resp.text}")
    result = resp.json()
    # 优先取 md_results（Markdown格式），也可取 layout_details
    md_text = result.get("md_results", "")
    if not md_text:
        # 备用：拼接 layout_details 中的文字
        details = result.get("layout_details", [])
        md_text = "\n".join(
            d.get("text", "") for d in details if d.get("text")
        )
    return md_text


def ocr_with_glm4v_fallback(image_path: str) -> str:
    """
    备用方案：使用 glm-4v-flash 视觉模型识别图片（max_tokens 限制在 1024 以内）。
    图片会先用 Pillow 压缩至 800x800。
    """
    print("[OCR] 正在调用智谱 GLM-4V-Flash (备用)...")
    # 压缩图片
    img = Image.open(image_path)
    img.thumbnail((800, 800), Image.LANCZOS)
    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=80)
    buf.seek(0)
    img_b64 = base64.b64encode(buf.read()).decode("utf-8")
    print(f"[OCR] 压缩后 Base64 长度: {len(img_b64)}")

    from zhipuai import ZhipuAI
    client = ZhipuAI(api_key=ZHIPU_API_KEY)
    response = client.chat.completions.create(
        model="glm-4v-flash",
        messages=[{
            "role": "user",
            "content": [
                {"type": "image_url", "image_url": {"url": img_b64}},
                {"type": "text", "text": (
                    "请仔细识别这张图片中的所有文字内容，"
                    "包括商品名称、条码、编号、数量、单价、金额、单位等所有信息，"
                    "原样输出，不要做任何总结或解释，保留原始格式。"
                )}
            ]
        }],
        max_tokens=1024,   # glm-4v-flash 限制
    )
    return response.choices[0].message.content


def ocr_image_with_glm(image_path: str) -> str:
    """
    主 OCR 入口：优先使用 GLM-OCR（layout_parsing），
    若失败则自动降级到 glm-4v-flash。
    """
    print(f"[OCR] 读取图片: {image_path}")
    ocr_text = ""

    # 优先：GLM-OCR 专用端点
    try:
        ocr_text = ocr_with_glm_ocr(image_path)
        print("[OCR] GLM-OCR 识别成功")
    except Exception as e:
        print(f"[OCR] GLM-OCR 失败 ({e})，切换到 GLM-4V-Flash 备用方案...")
        try:
            ocr_text = ocr_with_glm4v_fallback(image_path)
            print("[OCR] GLM-4V-Flash 识别成功")
        except Exception as e2:
            raise RuntimeError(f"所有OCR方案均失败: {e2}") from e2

    print("[OCR] 识别完成，原始文本：")
    print("─" * 60)
    print(ocr_text)
    print("─" * 60)
    return ocr_text


# ─────────────────────────────────────────────
# 第三步：DeepSeek 语义匹配 —— 文本 → JSON
# ─────────────────────────────────────────────
def match_to_template_with_deepseek(ocr_text: str, headers: list) -> list:
    """
    将 OCR 文本和模板表头发给 DeepSeek，
    要求返回 JSON 数组，每个元素对应模板中一行商品数据。
    """
    from openai import OpenAI
    client = OpenAI(
        api_key=DEEPSEEK_API_KEY,
        base_url=DEEPSEEK_BASE_URL,
    )

    headers_str = json.dumps(headers, ensure_ascii=False)

    system_prompt = (
        "你是一个专业的数据提取助手。"
        "用户会给你一段从单据图片中 OCR 识别出的原始文字，"
        "以及一个 Excel 模板的表头字段列表。\n"
        "你的任务是：严格按照表头字段，从 OCR 文本中提取所有商品数据，"
        "并以 JSON 数组格式输出，每个元素是一个对象，键为表头字段名，值为对应数据。\n"
        "规则：\n"
        "1. 只输出纯 JSON，不要有任何其他说明文字、markdown 代码块标记。\n"
        "2. 如果某个字段无法从文本中找到对应数据，该字段值设为空字符串 \"\"。\n"
        "3. 如果有多行商品，输出多个 JSON 对象。\n"
        "4. 数量、单价、折扣等数值字段请只保留数字（含小数点）。"
    )

    user_prompt = (
        f"模板表头字段：{headers_str}\n\n"
        f"OCR 识别到的单据文本：\n{ocr_text}\n\n"
        "请提取所有商品数据并严格按 JSON 数组格式输出。"
    )

    print("[DeepSeek] 正在进行语义匹配...")
    response = client.chat.completions.create(
        model=DEEPSEEK_MODEL,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user",   "content": user_prompt},
        ],
        temperature=0.0,
        max_tokens=4096,
    )

    raw_result = response.choices[0].message.content
    print("[DeepSeek] 返回结果：")
    print("─" * 60)
    print(raw_result)
    print("─" * 60)

    # ── 解析 JSON ──────────────────────────────
    # 去除可能的 markdown 代码块
    cleaned = re.sub(r"```(?:json)?", "", raw_result).strip().rstrip("`").strip()

    try:
        data = json.loads(cleaned)
    except json.JSONDecodeError:
        # 尝试提取第一个 [ ... ] 块
        match = re.search(r'\[.*\]', cleaned, re.DOTALL)
        if match:
            data = json.loads(match.group())
        else:
            print("[警告] 无法解析 JSON，将尝试用空数据生成文件")
            data = [{h: "" for h in headers}]

    # 确保是列表
    if isinstance(data, dict):
        data = [data]

    print(f"[DeepSeek] 解析到 {len(data)} 条商品记录")
    return data


# ─────────────────────────────────────────────
# 第四步：导出 Excel
# ─────────────────────────────────────────────
def export_to_excel(records: list, headers: list, output_path: str):
    """将匹配结果按模板表头顺序导出为 Excel"""
    rows = []
    for rec in records:
        row = {h: rec.get(h, "") for h in headers}
        rows.append(row)

    df = pd.DataFrame(rows, columns=headers)
    print("\n[导出] 数据预览：")
    print(df.to_string(index=False))

    df.to_excel(output_path, index=False, engine="openpyxl")
    print(f"\n[导出] 已成功保存至: {output_path}")


# ─────────────────────────────────────────────
# 核心处理函数（供 GUI 调用）
# ─────────────────────────────────────────────
def process_image(image_path: str, output_path: str = None, log_callback=None) -> str:
    """
    核心处理流程，供外部（GUI）调用。
    
    参数:
        image_path:    要处理的单据图片路径
        output_path:   输出 Excel 文件路径，默认与图片同目录
        log_callback:  可选的日志回调函数 log_callback(msg: str)
    
    返回:
        最终输出的 Excel 文件路径
    """
    def log(msg):
        print(msg)
        if log_callback:
            log_callback(msg)

    log("=" * 50)
    log("  自动化单据处理工具 —— OCR to Excel")
    log("=" * 50)

    # 确定输出路径
    if output_path is None:
        img_dir = os.path.dirname(os.path.abspath(image_path))
        img_name = os.path.splitext(os.path.basename(image_path))[0]
        output_path = os.path.join(img_dir, f"{img_name}_输出.xlsx")

    # 1. 读取模板表头
    log("\n[步骤 1/4] 读取模板表头...")
    headers = get_template_headers(TEMPLATE_PATH)
    log(f"  → 共 {len(headers)} 个字段")

    # 2. OCR 识图
    log("\n[步骤 2/4] 正在调用 AI 识别图片文字...")
    ocr_text = ocr_image_with_glm(image_path)
    log(f"  → 识别完成，共 {len(ocr_text)} 个字符")

    # 3. DeepSeek 语义匹配
    log("\n[步骤 3/4] 正在进行语义匹配与字段对齐...")
    records = match_to_template_with_deepseek(ocr_text, headers)
    log(f"  → 匹配到 {len(records)} 条商品记录")

    # 4. 导出 Excel
    log("\n[步骤 4/4] 正在生成 Excel 文件...")
    export_to_excel(records, headers, output_path)
    log(f"\n✅ 全部完成！文件已保存至:\n   {output_path}")

    return output_path


# ─────────────────────────────────────────────
# 命令行入口（直接运行时使用默认测试图片）
# ─────────────────────────────────────────────
if __name__ == "__main__":
    result = process_image(IMAGE_PATH, OUTPUT_PATH)
    print(f"\n输出文件: {result}")
