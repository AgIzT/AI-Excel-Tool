# -*- coding: utf-8 -*-
"""
ocr_to_excel.py
单据处理核心：图片/表格 → GLM 视觉模型（一步出结构化 JSON）→ Excel 标准模板
"""

import os
import base64
import json
import re
import datetime

import pandas as pd

from config import (
    BASE_DIR,
    ApiConfig,
    GLM_BASE_URL,
    GLM_ENV_NAMES,
    EXTRACT_MODEL,
    EXTRACT_TEMPERATURE,
    EXTRACT_MAX_TOKENS,
    EXTRACT_PROMPT_PATH,
    IMG_EXTS,
    TEMPLATE_HEADER_ROW,
)


def _default_api_config() -> ApiConfig:
    """设置页未传入时的兜底接口配置（默认走智谱 GLM）。"""
    return ApiConfig(base_url=GLM_BASE_URL, model=EXTRACT_MODEL)

TEMPLATE_PATH = os.path.join(BASE_DIR, "templates", "进货单商品导入模板.xls")


class ExtractionError(RuntimeError):
    """提取失败（模型未返回可解析的结构化数据）"""


# ─────────────────────────────────────────────
# 文件名辅助工具
# ─────────────────────────────────────────────
def _sanitize_filename(name: str) -> str:
    """移除 Windows/Linux 文件名中不允许的非法字符"""
    return re.sub(r'[\\/:*?"<>|\n\r\t]', '', name).strip()


def _build_smart_filename(meta: dict, output_dir: str, suffix: str = "_入库.xlsx") -> str:
    """
    根据元信息生成智能文件名：YYYY-MM-DD_供应商名称_入库.xlsx
    缺字段时用今天日期 / '未知供应商'；同名已存在时追加 _2、_3 …
    """
    date_str = (meta.get("date") or "").strip()
    supplier_str = (meta.get("supplier") or "").strip()

    if not date_str:
        date_str = datetime.datetime.now().strftime("%Y-%m-%d")
    if not supplier_str:
        supplier_str = "未知供应商"

    supplier_str = _sanitize_filename(supplier_str) or "未知供应商"

    path = os.path.join(output_dir, f"{date_str}_{supplier_str}{suffix}")
    if os.path.exists(path):
        stem = f"{date_str}_{supplier_str}"
        count = 2
        while os.path.exists(os.path.join(output_dir, f"{stem}_{count}{suffix}")):
            count += 1
        path = os.path.join(output_dir, f"{stem}_{count}{suffix}")
    return path


# ─────────────────────────────────────────────
# API Key 解析
# ─────────────────────────────────────────────
def _resolve_runtime_key(explicit_key: str = "", env_names: tuple = (), display_name: str = "") -> str:
    key = (explicit_key or "").strip()
    if key:
        return key
    for env_name in env_names:
        val = (os.getenv(env_name) or "").strip()
        if val:
            return val
    raise RuntimeError(f"未配置 {display_name}，请先在设置页填写并保存")


def _build_client(api_config: ApiConfig):
    from openai import OpenAI
    key = _resolve_runtime_key(api_config.api_key, GLM_ENV_NAMES, "API Key")
    base_url = (api_config.base_url or GLM_BASE_URL).strip()
    return OpenAI(api_key=key, base_url=base_url)


# ─────────────────────────────────────────────
# 模板表头
# ─────────────────────────────────────────────
def get_template_headers(tpl_path: str) -> list:
    """读取模板表头（默认第 2 行），同时支持 .xls 和 .xlsx。"""
    ext = os.path.splitext(tpl_path)[-1].lower()
    engine = "xlrd" if ext == ".xls" else "openpyxl"
    df = pd.read_excel(tpl_path, sheet_name=0, header=None, engine=engine)
    if len(df) <= TEMPLATE_HEADER_ROW:
        raise ValueError(f"模板行数不足，无法读取第 {TEMPLATE_HEADER_ROW + 1} 行表头：{tpl_path}")
    raw_headers = df.iloc[TEMPLATE_HEADER_ROW].tolist()
    headers = [str(h).strip() for h in raw_headers if str(h).strip() and str(h).strip().lower() != "nan"]
    if not headers:
        raise ValueError(f"模板第 {TEMPLATE_HEADER_ROW + 1} 行没有有效表头：{tpl_path}")
    print(f"[模板表头] {headers}")
    return headers


# ─────────────────────────────────────────────
# 输入读取
# ─────────────────────────────────────────────
def image_to_data_url(image_path: str) -> str:
    """读取图片并转换为 data:image/...;base64,... 格式"""
    ext = os.path.splitext(image_path)[-1].lower().lstrip(".")
    fmt_map = {"jpg": "jpeg", "jpeg": "jpeg", "png": "png",
               "bmp": "bmp", "gif": "gif", "webp": "webp", "tiff": "tiff"}
    mime = fmt_map.get(ext, "jpeg")
    with open(image_path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("utf-8")
    return f"data:image/{mime};base64,{b64}"


def read_excel_as_text(excel_path: str) -> str:
    """把表格内容转成文本，供模型对齐字段。支持 .xlsx / .xls / .csv。"""
    print(f"[表格读取] {excel_path}")
    ext = os.path.splitext(excel_path)[-1].lower()

    try:
        if ext == ".csv":
            try:
                df = pd.read_csv(excel_path, header=None, dtype=str, keep_default_na=False)
            except UnicodeDecodeError:
                df = pd.read_csv(excel_path, header=None, dtype=str, keep_default_na=False, encoding="gbk")
            df_dict = {"CSV": df}
        elif ext == ".xls":
            df_dict = pd.read_excel(excel_path, sheet_name=None, engine="xlrd", header=None)
        else:
            df_dict = pd.read_excel(excel_path, sheet_name=None, engine="openpyxl", header=None)
    except Exception as e:
        raise RuntimeError(f"读取表格失败: {e}") from e

    all_text = []
    for sheet_name, df in df_dict.items():
        all_text.append(f"[工作表: {sheet_name}]")
        df = df.fillna("")
        for _, row in df.iterrows():
            row_str = "\t".join(str(v).strip() for v in row)
            if row_str.strip():
                all_text.append(row_str)
        all_text.append("")

    result = "\n".join(all_text)
    print(f"[表格读取] 完成，共 {len(result)} 个字符")
    return result


# ─────────────────────────────────────────────
# 单视觉模型提取：图片/文本 → records + meta
# ─────────────────────────────────────────────
def _load_extract_system_prompt(handwriting: bool) -> str:
    with open(EXTRACT_PROMPT_PATH, encoding="utf-8") as f:
        tpl = f.read()
    note = ""
    if handwriting:
        note = (
            "\n7. 特别注意：单据中可能存在手写修改的数字或备注。"
            "若打印文字与手写文字冲突，以手写的实际修改内容为准，"
            "不要采用被划掉或覆盖的印刷数字。"
        )
    return tpl.replace("__HANDWRITING_NOTE__", note)


def _parse_extract_result(raw_result: str, headers: list) -> tuple:
    """解析模型输出为 (records, meta)。完全无法解析时抛 ExtractionError。"""
    cleaned = re.sub(r"```(?:json)?", "", raw_result or "").strip().rstrip("`").strip()

    meta = {"supplier": "", "date": ""}
    parsed = None
    try:
        parsed = json.loads(cleaned)
    except json.JSONDecodeError:
        m = re.search(r"\{.*\}", cleaned, re.DOTALL)
        if m:
            try:
                parsed = json.loads(m.group())
            except json.JSONDecodeError:
                pass
        if parsed is None:
            m = re.search(r"\[.*\]", cleaned, re.DOTALL)
            if m:
                try:
                    parsed = json.loads(m.group())
                except json.JSONDecodeError:
                    pass

    if parsed is None:
        raise ExtractionError("模型未返回可解析的 JSON")

    if isinstance(parsed, dict) and "records" in parsed:
        raw_meta = parsed.get("meta") or {}
        meta["supplier"] = str(raw_meta.get("supplier") or "").strip()
        meta["date"] = str(raw_meta.get("date") or "").strip()
        data = parsed["records"]
        if not isinstance(data, list):
            data = [data] if isinstance(data, dict) else []
    elif isinstance(parsed, list):
        data = parsed
    elif isinstance(parsed, dict):
        data = [parsed]
    else:
        data = []

    records = [r for r in data if isinstance(r, dict)]
    return records, meta


def extract_records(
    file_path: str,
    headers: list,
    handwriting: bool = False,
    api_config: ApiConfig = None,
    log=print,
) -> tuple:
    """
    单一入口：图片走视觉识别，表格走文本对齐，统一调用所配置的多模态模型，
    一步返回 (records, meta)。
    """
    api_config = api_config or _default_api_config()
    model = (api_config.model or EXTRACT_MODEL).strip()
    ext = os.path.splitext(file_path)[-1].lower()
    headers_str = json.dumps(headers, ensure_ascii=False)

    if ext in IMG_EXTS:
        log(f"  → 调用 {model} 识别图片（手写体：{'开' if handwriting else '关'}）")
        data_url = image_to_data_url(file_path)
        user_content = [
            {
                "type": "text",
                "text": (
                    f"模板表头字段：{headers_str}\n\n"
                    "请从这张单据图片中提取供应商名称、单据日期及所有商品数据，"
                    "按系统指定的 JSON 格式输出。"
                ),
            },
            {"type": "image_url", "image_url": {"url": data_url}},
        ]
    else:
        log(f"  → 读取表格并调用 {model} 对齐字段")
        text = read_excel_as_text(file_path)
        user_content = (
            f"模板表头字段：{headers_str}\n\n"
            f"表格文本：\n{text}\n\n"
            "请提取供应商名称、单据日期及所有商品数据，按系统指定的 JSON 格式输出。"
        )

    client = _build_client(api_config)
    system_prompt = _load_extract_system_prompt(handwriting)
    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_content},
        ],
        temperature=EXTRACT_TEMPERATURE,
        max_tokens=EXTRACT_MAX_TOKENS,
    )
    raw_result = response.choices[0].message.content
    records, meta = _parse_extract_result(raw_result, headers)
    log(
        f"  → 提取到 {len(records)} 条记录 | "
        f"供应商: {meta['supplier'] or '未识别'} | 日期: {meta['date'] or '未识别'}"
    )
    return records, meta


# ─────────────────────────────────────────────
# 导出 Excel
# ─────────────────────────────────────────────
def export_to_excel(records: list, headers: list, output_path: str):
    rows = [{h: rec.get(h, "") for h in headers} for rec in records]
    df = pd.DataFrame(rows, columns=headers)
    df.to_excel(output_path, index=False, engine="openpyxl")
    print(f"[导出] {len(rows)} 行 → {output_path}")


def export_merged_excel(all_records: list, headers: list, output_path: str):
    merged = []
    for records in all_records:
        merged.extend(records)
    rows = [{h: rec.get(h, "") for h in headers} for rec in merged]
    df = pd.DataFrame(rows, columns=headers)
    df.to_excel(output_path, index=False, engine="openpyxl")
    print(f"[合并导出] 共 {len(merged)} 条记录 → {output_path}")


def export_separate_excel(all_records: list, headers: list, output_paths: list):
    for records, output_path in zip(all_records, output_paths):
        export_to_excel(records, headers, output_path)


# ─────────────────────────────────────────────
# 核心处理（供 GUI 调用）
# ─────────────────────────────────────────────
def process_image(
    image_path: str,
    output_path: str = None,
    log_callback=None,
    handwriting: bool = False,
    template_path: str = None,
    api_config: ApiConfig = None,
) -> str:
    """处理单个文件（图片或表格），返回输出 Excel 路径。"""
    def log(msg):
        print(msg)
        if log_callback:
            log_callback(msg)

    img_dir = os.path.dirname(os.path.abspath(image_path))
    smart_name = output_path is None
    _tpl_path = template_path or TEMPLATE_PATH

    log("[步骤 1/3] 读取模板表头...")
    headers = get_template_headers(_tpl_path)
    log(f"  → 共 {len(headers)} 个字段")

    log("[步骤 2/3] AI 提取...")
    records, meta = extract_records(image_path, headers, handwriting=handwriting, api_config=api_config, log=log)
    if not records:
        raise ExtractionError("未提取到任何商品记录")

    log("[步骤 3/3] 生成 Excel...")
    if smart_name:
        output_path = _build_smart_filename(meta, img_dir)
    export_to_excel(records, headers, output_path)
    log(f"✅ 完成：{output_path}")
    return output_path


def extract_batch(
    image_paths: list,
    log_callback=None,
    handwriting: bool = False,
    progress_callback=None,
    template_path: str = None,
    api_config: ApiConfig = None,
) -> dict:
    """
    仅识别、不落盘：对每个文件调用模型提取 (records, meta)，供「导出前人工复核」使用。
    返回 {"headers": [...], "items": [{"name","path","records","meta"}...], "failed": [(name,err)...]}。
    单个文件提取失败会被跳过并记入 failed；全部失败时 items 为空（由调用方决定如何处理）。
    """
    def log(msg):
        print(msg)
        if log_callback:
            log_callback(msg)

    total = len(image_paths)
    log(f"[批量识别] 共 {total} 个文件")

    _tpl_path = template_path or TEMPLATE_PATH
    log("\n[步骤 1] 读取模板表头...")
    headers = get_template_headers(_tpl_path)
    log(f"  → 共 {len(headers)} 个字段")

    items = []
    failed = []

    for idx, image_path in enumerate(image_paths, start=1):
        fname = os.path.basename(image_path)
        log(f"\n{'=' * 50}")
        log(f"[{idx}/{total}] {fname}")
        if progress_callback:
            progress_callback(idx - 1, total)

        try:
            records, meta = extract_records(
                image_path, headers, handwriting=handwriting, api_config=api_config, log=log
            )
            if not records:
                raise ExtractionError("未提取到任何商品记录")
        except Exception as e:
            log(f"  [跳过] 提取失败：{e}")
            failed.append((fname, str(e)))
            continue

        items.append({"name": fname, "path": image_path, "records": records, "meta": meta})

    if progress_callback:
        progress_callback(total, total)

    return {"headers": headers, "items": items, "failed": failed}


def export_batch(
    extraction: dict,
    output_dir: str,
    merge_output: bool = False,
    merged_output_path: str = None,
    log_callback=None,
) -> list:
    """
    把（可能经人工复核修改过的）识别结果写成 Excel。
    extraction 结构同 extract_batch 的返回值；仅使用其中的 headers 与 items。
    输出文件名按各 item 的 meta 在导出时即时生成，因此用户改了供应商/日期会反映到文件名。
    """
    def log(msg):
        print(msg)
        if log_callback:
            log_callback(msg)

    headers = extraction["headers"]
    items = extraction["items"]
    if not items:
        raise RuntimeError("没有可导出的记录")

    all_records = [it["records"] for it in items]
    all_metas = [it["meta"] for it in items]

    if merge_output:
        if not merged_output_path:
            merged_output_path = _build_smart_filename(all_metas[0], output_dir, suffix="_合并入库.xlsx")
        log(f"\n[合并导出] 将 {len(all_records)} 份数据合并...")
        export_merged_excel(all_records, headers, merged_output_path)
        log(f"  → {merged_output_path}")
        return [merged_output_path]

    log(f"\n[分别导出] 生成 {len(items)} 个文件...")
    output_paths = [_build_smart_filename(it["meta"], output_dir) for it in items]
    export_separate_excel(all_records, headers, output_paths)
    for p in output_paths:
        log(f"  → {p}")
    return output_paths


def process_images_batch(
    image_paths: list,
    output_dir: str,
    log_callback=None,
    handwriting: bool = False,
    merge_output: bool = False,
    merged_output_path: str = None,
    progress_callback=None,
    template_path: str = None,
    api_config: ApiConfig = None,
) -> list:
    """
    一步到位（识别即导出，无人工复核）。GUI 现在改用 extract_batch + 复核 + export_batch，
    此入口保留给命令行/无人值守场景。全部失败时抛出异常。
    """
    def log(msg):
        print(msg)
        if log_callback:
            log_callback(msg)

    total = len(image_paths)
    log(f"[批量处理] 共 {total} 个文件，模式={'合并输出' if merge_output else '分别输出'}")

    extraction = extract_batch(
        image_paths,
        log_callback=log_callback,
        handwriting=handwriting,
        progress_callback=progress_callback,
        template_path=template_path,
        api_config=api_config,
    )

    if not extraction["items"]:
        raise RuntimeError(f"全部 {total} 个文件提取失败，未生成任何文件（详见日志）")

    result_paths = export_batch(
        extraction,
        output_dir,
        merge_output=merge_output,
        merged_output_path=merged_output_path,
        log_callback=log_callback,
    )

    ok = len(extraction["items"])
    failed = extraction["failed"]
    if failed:
        log(f"\n⚠ 完成 {ok}/{total}，{len(failed)} 个失败：")
        for fname, err in failed:
            log(f"   ✗ {fname}：{err}")
    else:
        log(f"\n✅ 批量处理完成！共 {ok} 个文件")
    return result_paths


# ─────────────────────────────────────────────
# 命令行测试入口
# ─────────────────────────────────────────────
if __name__ == "__main__":
    test_img = os.path.join(BASE_DIR, "test_receipt.jpg")
    if not os.path.exists(test_img):
        print(f"未找到测试图片 {test_img}，请改用 GUI（python main_app.py）测试。")
    else:
        print(process_image(test_img))
