# -*- coding: utf-8 -*-
"""
ocr_to_excel.py
自动化单据处理工具：图片 → GLM OCR → DeepSeek 语义匹配 → Excel 标准模板
支持：多图批处理 / 手写体识别开关 / 直接上传 Excel 跳过 OCR
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
ZHIPU_API_KEY    = ""
DEEPSEEK_API_KEY = ""
DEEPSEEK_BASE_URL = "https://api.deepseek.com/v1"
DEEPSEEK_MODEL   = "deepseek-chat"

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


def ocr_with_glm_ocr(image_path: str, handwriting: bool = False) -> str:
    """
    使用智谱 GLM-OCR (layout_parsing 端点) 识别图片文字。
    handwriting=True 时使用支持手写体增强的提示模式。
    """
    print(f"[OCR] 正在调用智谱 GLM-OCR (layout_parsing)... 手写体识别={'开启' if handwriting else '关闭'}")
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
    md_text = result.get("md_results", "")
    if not md_text:
        details = result.get("layout_details", [])
        md_text = "\n".join(
            d.get("text", "") for d in details if d.get("text")
        )
    return md_text


def ocr_with_glm4v_fallback(image_path: str, handwriting: bool = False) -> str:
    """
    备用方案：使用 glm-4v-flash 视觉模型识别图片。
    handwriting=True 时在提示词中特别强调手写体识别。
    """
    print(f"[OCR] 正在调用智谱 GLM-4V-Flash (备用)... 手写体识别={'开启' if handwriting else '关闭'}")
    img = Image.open(image_path)
    img.thumbnail((800, 800), Image.LANCZOS)
    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=80)
    buf.seek(0)
    img_b64 = base64.b64encode(buf.read()).decode("utf-8")

    if handwriting:
        ocr_instruction = (
            "请仔细识别这张图片中的所有文字内容，包括商品名称、条码、编号、数量、单价、金额、单位等所有信息。"
            "请特别注意表格中手写修改的数字或备注，"
            "如果遇到打印文字与手写文字重叠或冲突，请以手写的实际修改内容为准进行提取。"
            "原样输出全部内容，不要做任何总结或解释，保留原始格式。"
        )
    else:
        ocr_instruction = (
            "请仔细识别这张图片中的所有文字内容，"
            "包括商品名称、条码、编号、数量、单价、金额、单位等所有信息，"
            "原样输出，不要做任何总结或解释，保留原始格式。"
        )

    from zhipuai import ZhipuAI
    client = ZhipuAI(api_key=ZHIPU_API_KEY)
    response = client.chat.completions.create(
        model="glm-4v-flash",
        messages=[{
            "role": "user",
            "content": [
                {"type": "image_url", "image_url": {"url": img_b64}},
                {"type": "text", "text": ocr_instruction}
            ]
        }],
        max_tokens=1024,
    )
    return response.choices[0].message.content


def ocr_image_with_glm(image_path: str, handwriting: bool = False) -> str:
    """
    主 OCR 入口：优先使用 GLM-OCR（layout_parsing），
    若失败则自动降级到 glm-4v-flash。
    handwriting=True 时启用手写体增强识别提示。
    """
    print(f"[OCR] 读取图片: {image_path}")
    ocr_text = ""

    try:
        ocr_text = ocr_with_glm_ocr(image_path, handwriting=handwriting)
        print("[OCR] GLM-OCR 识别成功")
    except Exception as e:
        print(f"[OCR] GLM-OCR 失败 ({e})，切换到 GLM-4V-Flash 备用方案...")
        try:
            ocr_text = ocr_with_glm4v_fallback(image_path, handwriting=handwriting)
            print("[OCR] GLM-4V-Flash 识别成功")
        except Exception as e2:
            raise RuntimeError(f"所有OCR方案均失败: {e2}") from e2

    print("[OCR] 识别完成，原始文本：")
    print("─" * 60)
    print(ocr_text)
    print("─" * 60)
    return ocr_text


# ─────────────────────────────────────────────
# Excel 直接读取（跳过 OCR 步骤）
# ─────────────────────────────────────────────
def read_excel_as_text(excel_path: str) -> str:
    """
    直接读取 Excel 文件，将内容转换为文本格式供 DeepSeek 处理。
    支持 .xlsx 和 .xls 格式。
    """
    print(f"[Excel读取] 直接读取 Excel 文件: {excel_path}")
    ext = os.path.splitext(excel_path)[-1].lower()

    try:
        if ext == ".xls":
            df_dict = pd.read_excel(excel_path, sheet_name=None, engine="xlrd", header=None)
        else:
            df_dict = pd.read_excel(excel_path, sheet_name=None, engine="openpyxl", header=None)
    except Exception as e:
        raise RuntimeError(f"读取 Excel 失败: {e}")

    all_text = []
    for sheet_name, df in df_dict.items():
        all_text.append(f"[工作表: {sheet_name}]")
        # 转为文本表格
        df = df.fillna("")
        for _, row in df.iterrows():
            row_str = "\t".join(str(v).strip() for v in row)
            if row_str.strip():
                all_text.append(row_str)
        all_text.append("")

    result = "\n".join(all_text)
    print(f"[Excel读取] 读取完成，共 {len(result)} 个字符")
    return result


# ─────────────────────────────────────────────
# 第三步：DeepSeek 语义匹配 —— 文本 → JSON
# ─────────────────────────────────────────────
def match_to_template_with_deepseek(
    ocr_text: str,
    headers: list,
    handwriting: bool = False
) -> list:
    """
    将 OCR 文本和模板表头发给 DeepSeek，
    要求返回 JSON 数组，每个元素对应模板中一行商品数据。
    handwriting=True 时在提示词中特别强调手写体优先。
    """
    from openai import OpenAI
    client = OpenAI(
        api_key=DEEPSEEK_API_KEY,
        base_url=DEEPSEEK_BASE_URL,
    )

    headers_str = json.dumps(headers, ensure_ascii=False)

    handwriting_note = ""
    if handwriting:
        handwriting_note = (
            "\n5. 特别注意：单据中可能存在手写修改的数字或备注。"
            "如果遇到打印文字与手写文字重叠或冲突，请以手写的实际修改数量/内容为准进行提取，"
            "不要使用被手写覆盖的印刷数字。"
        )

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
        + handwriting_note
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
    cleaned = re.sub(r"```(?:json)?", "", raw_result).strip().rstrip("`").strip()

    try:
        data = json.loads(cleaned)
    except json.JSONDecodeError:
        match = re.search(r'\[.*\]', cleaned, re.DOTALL)
        if match:
            data = json.loads(match.group())
        else:
            print("[警告] 无法解析 JSON，将尝试用空数据生成文件")
            data = [{h: "" for h in headers}]

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


def export_merged_excel(all_records: list, headers: list, output_path: str):
    """
    合并多张图片的结果到一个 Excel 文件（所有记录在同一 Sheet）。
    all_records: [ [records_of_img1], [records_of_img2], ... ]
    """
    merged = []
    for records in all_records:
        merged.extend(records)

    rows = [{h: rec.get(h, "") for h in headers} for rec in merged]
    df = pd.DataFrame(rows, columns=headers)
    df.to_excel(output_path, index=False, engine="openpyxl")
    print(f"[合并导出] 共 {len(merged)} 条记录 → {output_path}")


def export_separate_excel(all_records: list, headers: list, output_paths: list):
    """
    将多张图片的结果分别输出到各自的 Excel 文件。
    all_records:  每张图片的记录列表
    output_paths: 对应的输出文件路径列表
    """
    for records, output_path in zip(all_records, output_paths):
        export_to_excel(records, headers, output_path)


# ─────────────────────────────────────────────
# 核心处理函数（供 GUI 调用）
# ─────────────────────────────────────────────
def process_image(
    image_path: str,
    output_path: str = None,
    log_callback=None,
    handwriting: bool = False,
) -> str:
    """
    核心处理流程（单张图片），供外部（GUI）调用。

    参数:
        image_path:    要处理的单据图片路径（或 Excel 文件路径，届时跳过 OCR）
        output_path:   输出 Excel 文件路径，默认与图片同目录
        log_callback:  可选的日志回调函数 log_callback(msg: str)
        handwriting:   是否启用手写体增强识别

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

    if output_path is None:
        img_dir = os.path.dirname(os.path.abspath(image_path))
        img_name = os.path.splitext(os.path.basename(image_path))[0]
        output_path = os.path.join(img_dir, f"{img_name}_输出.xlsx")

    # 1. 读取模板表头
    log("\n[步骤 1/4] 读取模板表头...")
    headers = get_template_headers(TEMPLATE_PATH)
    log(f"  → 共 {len(headers)} 个字段")

    # 2. OCR 识图 / 直接读取 Excel
    ext = os.path.splitext(image_path)[-1].lower()
    is_excel = ext in (".xlsx", ".xls", ".csv")

    if is_excel:
        log("\n[步骤 2/4] 检测到 Excel 文件，直接读取（跳过 OCR）...")
        ocr_text = read_excel_as_text(image_path)
        log(f"  → 读取完成，共 {len(ocr_text)} 个字符")
    else:
        log(f"\n[步骤 2/4] 正在调用 AI 识别图片文字（手写体识别: {'开启' if handwriting else '关闭'}）...")
        ocr_text = ocr_image_with_glm(image_path, handwriting=handwriting)
        log(f"  → 识别完成，共 {len(ocr_text)} 个字符")

    # 3. DeepSeek 语义匹配
    log("\n[步骤 3/4] 正在进行语义匹配与字段对齐...")
    records = match_to_template_with_deepseek(ocr_text, headers, handwriting=handwriting)
    log(f"  → 匹配到 {len(records)} 条商品记录")

    # 4. 导出 Excel
    log("\n[步骤 4/4] 正在生成 Excel 文件...")
    export_to_excel(records, headers, output_path)
    log(f"\n✅ 全部完成！文件已保存至:\n   {output_path}")

    return output_path


def process_images_batch(
    image_paths: list,
    output_dir: str,
    log_callback=None,
    handwriting: bool = False,
    merge_output: bool = False,
    merged_output_path: str = None,
    progress_callback=None,
) -> list:
    """
    批量处理多张图片/Excel，供 GUI 调用。

    参数:
        image_paths:         图片或 Excel 文件路径列表
        output_dir:          输出目录
        log_callback:        日志回调 log_callback(msg)
        handwriting:         是否启用手写体增强识别
        merge_output:        True=合并到一个 Excel；False=分别输出
        merged_output_path:  merge_output=True 时的合并输出路径
        progress_callback:   进度回调 progress_callback(current, total)

    返回:
        输出文件路径列表
    """
    def log(msg):
        print(msg)
        if log_callback:
            log_callback(msg)

    total = len(image_paths)
    log(f"[批量处理] 共 {total} 个文件，模式={'合并输出' if merge_output else '分别输出'}")

    # 1. 读取模板表头（只读一次）
    log("\n[步骤 1] 读取模板表头...")
    headers = get_template_headers(TEMPLATE_PATH)
    log(f"  → 共 {len(headers)} 个字段")

    all_records = []
    all_output_paths = []
    ts = __import__("datetime").datetime.now().strftime("%Y%m%d_%H%M%S")

    for idx, image_path in enumerate(image_paths, start=1):
        fname = os.path.basename(image_path)
        log(f"\n{'='*50}")
        log(f"[{idx}/{total}] 正在处理: {fname}")

        if progress_callback:
            progress_callback(idx - 1, total)

        ext = os.path.splitext(image_path)[-1].lower()
        is_excel = ext in (".xlsx", ".xls", ".csv")

        # OCR / 直接读取
        if is_excel:
            log(f"  → 检测到 Excel 文件，直接读取（跳过 OCR）")
            ocr_text = read_excel_as_text(image_path)
        else:
            log(f"  → 调用 OCR（手写体识别: {'开启' if handwriting else '关闭'}）")
            ocr_text = ocr_image_with_glm(image_path, handwriting=handwriting)

        log(f"  → 文本长度: {len(ocr_text)} 字符")

        # DeepSeek 匹配
        log(f"  → DeepSeek 语义匹配中...")
        records = match_to_template_with_deepseek(ocr_text, headers, handwriting=handwriting)
        log(f"  → 匹配到 {len(records)} 条商品记录")

        all_records.append(records)

        # 确定单独输出路径
        img_name = os.path.splitext(fname)[0]
        out_path = os.path.join(output_dir, f"{img_name}_{ts}_入库.xlsx")
        all_output_paths.append(out_path)

    if progress_callback:
        progress_callback(total - 1, total)

    # 导出
    if merge_output:
        if not merged_output_path:
            merged_output_path = os.path.join(output_dir, f"合并输出_{ts}.xlsx")
        log(f"\n[合并导出] 正在将 {total} 份数据合并到一个文件...")
        export_merged_excel(all_records, headers, merged_output_path)
        log(f"  → 合并文件: {merged_output_path}")
        result_paths = [merged_output_path]
    else:
        log(f"\n[分别导出] 正在生成 {total} 个独立文件...")
        export_separate_excel(all_records, headers, all_output_paths)
        for p in all_output_paths:
            log(f"  → {p}")
        result_paths = all_output_paths

    log(f"\n✅ 批量处理完成！共处理 {total} 个文件")
    return result_paths


# ─────────────────────────────────────────────
# 命令行入口（直接运行时使用默认测试图片）
# ─────────────────────────────────────────────
if __name__ == "__main__":
    result = process_image(IMAGE_PATH, OUTPUT_PATH)
    print(f"\n输出文件: {result}")
