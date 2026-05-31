# -*- coding: utf-8 -*-
"""
settings_page.py
「设置」页：接口预设选择 + base_url/模型/API Key 自定义，外加界面主题切换，
全部持久化到 QSettings。以 Mixin 形式提供给 MainWindow，依赖 self.settings 与 self._append_log。
"""

from PySide6.QtWidgets import (
    QHBoxLayout,
    QSizePolicy,
    QSpacerItem,
    QVBoxLayout,
    QWidget,
)

from qfluentwidgets import (
    BodyLabel,
    CaptionLabel,
    ComboBox,
    LineEdit,
    PasswordLineEdit,
    PrimaryPushButton,
    PushButton,
    SimpleCardWidget,
    SubtitleLabel,
    Theme,
    setTheme,
)

from config import PROVIDER_PRESETS, DEFAULT_PROVIDER, ApiConfig

# 界面主题：下拉显示文案 → QSettings 存储值 → qfluentwidgets 主题枚举
THEME_OPTIONS = [("跟随系统", "AUTO"), ("浅色", "LIGHT"), ("深色", "DARK")]
THEME_ENUM = {"AUTO": Theme.AUTO, "LIGHT": Theme.LIGHT, "DARK": Theme.DARK}
THEME_INDEX = {name: i for i, (_, name) in enumerate(THEME_OPTIONS)}


class SettingsPageMixin:
    def _build_setting_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 18, 24, 22)
        layout.setSpacing(14)

        # ---- 凭证设置卡片 ----
        cred_card = SimpleCardWidget(page)
        cred_layout = QVBoxLayout(cred_card)
        cred_layout.setContentsMargins(18, 16, 18, 16)
        cred_layout.setSpacing(10)
        cred_title = SubtitleLabel("凭证设置", cred_card)
        cred_layout.addWidget(cred_title, 0)

        provider_label = BodyLabel("接口预设", cred_card)
        cred_layout.addWidget(provider_label, 0)
        self.provider_combo = ComboBox(cred_card)
        self.provider_combo.addItems(list(PROVIDER_PRESETS.keys()))
        self.provider_combo.currentTextChanged.connect(self._on_provider_changed)
        cred_layout.addWidget(self.provider_combo, 0)

        base_url_label = BodyLabel("接口地址 (base_url)", cred_card)
        cred_layout.addWidget(base_url_label, 0)
        self.base_url_edit = LineEdit(cred_card)
        self.base_url_edit.setPlaceholderText("如 https://open.bigmodel.cn/api/paas/v4/")
        cred_layout.addWidget(self.base_url_edit, 0)

        model_label = BodyLabel("模型名称", cred_card)
        cred_layout.addWidget(model_label, 0)
        self.model_edit = LineEdit(cred_card)
        self.model_edit.setPlaceholderText("多模态/视觉模型，如 glm-4.6v、gpt-4o、qwen-vl-max")
        cred_layout.addWidget(self.model_edit, 0)

        key_label = BodyLabel("API Key", cred_card)
        cred_layout.addWidget(key_label, 0)
        self.api_key_edit = PasswordLineEdit(cred_card)
        self.api_key_edit.setPlaceholderText("请输入对应平台的 API Key")
        cred_layout.addWidget(self.api_key_edit, 0)

        self._conn_worker = None
        btn_row = QWidget(cred_card)
        btn_row_layout = QHBoxLayout(btn_row)
        btn_row_layout.setContentsMargins(0, 0, 0, 0)
        btn_row_layout.setSpacing(8)
        save_btn = PrimaryPushButton("保存配置", btn_row)
        save_btn.clicked.connect(self._on_save_api_settings)
        self.test_conn_btn = PushButton("测试连通性", btn_row)
        self.test_conn_btn.clicked.connect(self._on_test_connectivity)
        btn_row_layout.addWidget(save_btn, 0)
        btn_row_layout.addWidget(self.test_conn_btn, 0)
        btn_row_layout.addStretch(1)
        cred_layout.addWidget(btn_row, 0)

        tip = CaptionLabel(
            "提示：任何 OpenAI 兼容的多模态模型均可。切换预设会自动填入地址与模型，仍可手动修改。",
            cred_card,
        )
        tip.setWordWrap(True)
        cred_layout.addWidget(tip, 0)
        layout.addWidget(cred_card, 0)

        # ---- 界面主题卡片 ----
        theme_card = SimpleCardWidget(page)
        theme_layout = QVBoxLayout(theme_card)
        theme_layout.setContentsMargins(18, 16, 18, 16)
        theme_layout.setSpacing(10)
        theme_title = SubtitleLabel("界面主题", theme_card)
        theme_layout.addWidget(theme_title, 0)

        theme_label = BodyLabel("外观模式", theme_card)
        theme_layout.addWidget(theme_label, 0)
        self.theme_combo = ComboBox(theme_card)
        self.theme_combo.addItems([text for text, _ in THEME_OPTIONS])
        self.theme_combo.currentIndexChanged.connect(self._on_theme_changed)
        theme_layout.addWidget(self.theme_combo, 0)

        theme_tip = CaptionLabel("「跟随系统」会根据 Windows 的浅色/深色设置自动切换。", theme_card)
        theme_tip.setWordWrap(True)
        theme_layout.addWidget(theme_tip, 0)
        layout.addWidget(theme_card, 0)

        layout.addItem(QSpacerItem(10, 10, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding))

        self._load_api_settings_to_inputs()
        self._load_theme_to_input()
        return page

    def _get_saved_api_config(self) -> ApiConfig:
        base_url = str(self.settings.value("api/base_url", "", str) or "").strip()
        model = str(self.settings.value("api/model", "", str) or "").strip()
        api_key = str(self.settings.value("api/api_key", "", str) or "").strip()
        return ApiConfig(base_url=base_url, model=model, api_key=api_key)

    def _load_api_settings_to_inputs(self):
        if not hasattr(self, "provider_combo"):
            return
        provider = str(self.settings.value("api/provider", DEFAULT_PROVIDER, str) or DEFAULT_PROVIDER)
        if provider not in PROVIDER_PRESETS:
            provider = DEFAULT_PROVIDER
        preset = PROVIDER_PRESETS.get(provider, {})
        cfg = self._get_saved_api_config()
        # 首次使用（无保存值）时用预设默认回填
        base_url = cfg.base_url or preset.get("base_url", "")
        model = cfg.model or preset.get("model", "")
        self.provider_combo.blockSignals(True)
        self.provider_combo.setCurrentText(provider)
        self.provider_combo.blockSignals(False)
        self.base_url_edit.setText(base_url)
        self.model_edit.setText(model)
        self.api_key_edit.setText(cfg.api_key)

    def _on_provider_changed(self, name: str):
        if not hasattr(self, "base_url_edit"):
            return
        preset = PROVIDER_PRESETS.get(name)
        if preset is None:
            return
        self.base_url_edit.setText(preset.get("base_url", ""))
        self.model_edit.setText(preset.get("model", ""))

    def _persist_api_settings(self, provider: str, base_url: str, model: str, api_key: str):
        self.settings.setValue("api/provider", provider)
        self.settings.setValue("api/base_url", (base_url or "").strip())
        self.settings.setValue("api/model", (model or "").strip())
        self.settings.setValue("api/api_key", (api_key or "").strip())
        self.settings.sync()

    def _require_api_config(self):
        cfg = self._get_saved_api_config()
        missing = []
        if not cfg.base_url:
            missing.append("接口地址")
        if not cfg.model:
            missing.append("模型名称")
        if not cfg.api_key:
            missing.append("API Key")
        if missing:
            fields = "、".join(missing)
            self._append_log(f"[配置缺失] 请先在设置页填写：{fields}")
            self._toast("warning", "缺少配置", f"请先到【设置】页面填写并保存：{fields}。")
            return None
        return cfg

    def _on_save_api_settings(self):
        if not hasattr(self, "api_key_edit"):
            return
        provider = self.provider_combo.currentText().strip()
        base_url = self.base_url_edit.text().strip()
        model = self.model_edit.text().strip()
        api_key = self.api_key_edit.text().strip()
        self._persist_api_settings(provider, base_url, model, api_key)
        self._append_log(f"[设置] 接口配置已保存（{provider} | {model or '未填模型'}）")
        self._toast("success", "保存成功", "接口配置已保存。")

    # ---- 模型连通性测试 ----
    def _on_test_connectivity(self):
        if not hasattr(self, "api_key_edit"):
            return
        # 已有测试在跑则忽略，避免重复发请求
        if self._conn_worker is not None and self._conn_worker.isRunning():
            return
        base_url = self.base_url_edit.text().strip()
        model = self.model_edit.text().strip()
        api_key = self.api_key_edit.text().strip()
        # 地址/模型必填；API Key 允许留空走环境变量兜底，缺失时由后端给出清晰报错
        missing = []
        if not base_url:
            missing.append("接口地址")
        if not model:
            missing.append("模型名称")
        if missing:
            self._toast("warning", "无法测试", f"请先填写：{'、'.join(missing)}。")
            return

        cfg = ApiConfig(base_url=base_url, model=model, api_key=api_key)
        from workers import ConnectivityWorker

        self.test_conn_btn.setEnabled(False)
        self.test_conn_btn.setText("测试中…")
        self._append_log(f"[连通性] 正在测试 {model} @ {base_url} …")
        worker = ConnectivityWorker(cfg)
        worker.finished_ok.connect(self._on_conn_ok)
        worker.finished_err.connect(self._on_conn_err)
        worker.finished.connect(self._on_conn_done)
        self._conn_worker = worker
        worker.start()

    def _on_conn_ok(self, reply: str):
        short = reply if len(reply) <= 40 else reply[:40] + "…"
        self._append_log(f"[连通性] 成功，模型回显：{short}")
        self._toast("success", "连接成功", f"模型已响应：{short}")

    def _on_conn_err(self, err: str):
        short = err if len(err) <= 90 else err[:90] + "…"
        self._append_log(f"[连通性] 失败：{err}")
        self._toast("error", "连接失败", short, duration=5000)

    def _on_conn_done(self):
        self.test_conn_btn.setEnabled(True)
        self.test_conn_btn.setText("测试连通性")
        self._conn_worker = None

    # ---- 界面主题 ----
    def _load_theme_to_input(self):
        if not hasattr(self, "theme_combo"):
            return
        name = str(self.settings.value("ui/theme", "AUTO", str) or "AUTO").upper()
        index = THEME_INDEX.get(name, 0)
        self.theme_combo.blockSignals(True)
        self.theme_combo.setCurrentIndex(index)
        self.theme_combo.blockSignals(False)

    def _on_theme_changed(self, index: int):
        if not (0 <= index < len(THEME_OPTIONS)):
            return
        name = THEME_OPTIONS[index][1]
        setTheme(THEME_ENUM.get(name, Theme.AUTO))
        self.settings.setValue("ui/theme", name)
        self.settings.sync()
        self._append_log(f"[设置] 界面主题已切换为 {THEME_OPTIONS[index][0]}")
