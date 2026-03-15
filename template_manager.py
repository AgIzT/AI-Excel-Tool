# -*- coding: utf-8 -*-
"""
template_manager.py
模板管理器 —— 内置模板注册 + 用户自定义模板持久化

架构设计说明：
  · 内置模板：随程序打包，只需在 BUILTIN_TEMPLATES 字典中新增一条记录即可扩展
  · 自定义模板：用户通过文件对话框选择 .xls/.xlsx，持久化到 user_templates.json
  · 所有消费方只需调用 get_template_path(name) 即可获取模板绝对路径
  · 未来可扩展：模板分组、模板描述、自动发现 templates/ 子目录等
"""

import os
import json


class TemplateManager:
    """
    统一管理内置模板和用户自定义模板。

    内置模板注册表：修改 BUILTIN_TEMPLATES 即可新增内置模板，无需改动其他代码。
    自定义模板：持久化存储在 base_dir/user_templates.json，程序重启后仍保留。
    """

    # ── 内置模板注册表（显示名称 → 相对于 base_dir 的文件名）──────
    # 新增内置模板时，只需在此处添加一行即可
    BUILTIN_TEMPLATES: dict = {
        "进货单商品导入模板": "进货单商品导入模板.xls",
        "多规格商品模板":     "多规格商品模板.xls",
    }

    _CONFIG_FILE = "user_templates.json"

    def __init__(self, base_dir: str) -> None:
        self.base_dir = base_dir
        self._config_path = os.path.join(base_dir, self._CONFIG_FILE)
        self._custom: dict = {}   # 显示名称 → 绝对路径
        self._load_custom()

    # ── 持久化 ───────────────────────────────────────────────────

    def _load_custom(self) -> None:
        """从 JSON 文件读取已保存的自定义模板"""
        if not os.path.exists(self._config_path):
            return
        try:
            with open(self._config_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                self._custom = data
        except Exception:
            self._custom = {}

    def _save_custom(self) -> None:
        """将自定义模板写入 JSON 文件"""
        try:
            with open(self._config_path, "w", encoding="utf-8") as f:
                json.dump(self._custom, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    # ── 查询接口 ─────────────────────────────────────────────────

    def get_all_templates(self) -> dict:
        """
        返回所有可用模板：{显示名称: 绝对路径}。
        - 内置模板优先排列
        - 仅返回文件实际存在的条目
        - 自定义模板追加在内置模板之后
        """
        result: dict = {}

        for name, filename in self.BUILTIN_TEMPLATES.items():
            path = os.path.join(self.base_dir, filename)
            if os.path.exists(path):
                result[name] = path

        for name, path in self._custom.items():
            if os.path.exists(path):
                result[name] = path

        return result

    def get_template_names(self) -> list:
        """返回所有可用模板名称列表"""
        return list(self.get_all_templates().keys())

    def get_template_path(self, name: str) -> str:
        """
        根据显示名称获取模板绝对路径。
        找不到或文件已丢失时抛出 ValueError。
        """
        all_templates = self.get_all_templates()
        if name not in all_templates:
            raise ValueError(f"模板 '{name}' 不存在或文件已丢失")
        return all_templates[name]

    def get_default_name(self) -> str:
        """返回第一个可用模板的名称，通常为进货单模板"""
        names = self.get_template_names()
        return names[0] if names else ""

    # ── 自定义模板管理 ────────────────────────────────────────────

    def add_custom_template(self, name: str, path: str) -> None:
        """
        注册并持久化一个自定义模板。
        name: 显示名称（建议使用文件名去掉扩展名）
        path: 模板文件的绝对路径
        """
        path = os.path.abspath(path)
        if not os.path.exists(path):
            raise FileNotFoundError(f"模板文件不存在：{path}")
        self._custom[name] = path
        self._save_custom()

    def remove_custom_template(self, name: str) -> None:
        """删除一个自定义模板记录（不删除文件本身，且内置模板不可删除）"""
        if name in self._custom:
            del self._custom[name]
            self._save_custom()

    def is_builtin(self, name: str) -> bool:
        """判断指定名称是否为内置模板"""
        return name in self.BUILTIN_TEMPLATES
