# -*- coding: utf-8 -*-
"""
settings_page.py
「设置」页：接口预设选择 + base_url/模型/API Key 自定义，持久化到 QSettings。
以 Mixin 形式提供给 MainWindow，依赖 self.settings 与 self._append_log。
"""

from PySide6.QtWidgets import (
    QComboBox,
    QFrame,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QSpacerItem,
    QVBoxLayout,
    QWidget,
)

from config import PROVIDER_PRESETS, DEFAULT_PROVIDER, ApiConfig


class SettingsPageMixin:
    def _build_setting_page(self):
        page = QWidget(self.page_stack)
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)
        card = QFrame(page)
        card.setObjectName("card")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(16, 14, 16, 14)
        card_layout.setSpacing(10)
        title = QLabel("凭证设置", card)
        title.setObjectName("sectionTitle")
        card_layout.addWidget(title, 0)

        form = QFrame(card)
        form.setObjectName("subCard")
        form_layout = QVBoxLayout(form)
        form_layout.setContentsMargins(14, 12, 14, 12)
        form_layout.setSpacing(10)

        provider_label = QLabel("接口预设", form)
        provider_label.setObjectName("dimLabel")
        form_layout.addWidget(provider_label, 0)
        self.provider_combo = QComboBox(form)
        self.provider_combo.setObjectName("comboBox")
        self.provider_combo.addItems(list(PROVIDER_PRESETS.keys()))
        self.provider_combo.currentTextChanged.connect(self._on_provider_changed)
        form_layout.addWidget(self.provider_combo, 0)

        base_url_label = QLabel("接口地址 (base_url)", form)
        base_url_label.setObjectName("dimLabel")
        form_layout.addWidget(base_url_label, 0)
        self.base_url_edit = QLineEdit(form)
        self.base_url_edit.setObjectName("keyLineEdit")
        self.base_url_edit.setPlaceholderText("如 https://open.bigmodel.cn/api/paas/v4/")
        form_layout.addWidget(self.base_url_edit, 0)

        model_label = QLabel("模型名称", form)
        model_label.setObjectName("dimLabel")
        form_layout.addWidget(model_label, 0)
        self.model_edit = QLineEdit(form)
        self.model_edit.setObjectName("keyLineEdit")
        self.model_edit.setPlaceholderText("多模态/视觉模型，如 glm-4.6v、gpt-4o、qwen-vl-max")
        form_layout.addWidget(self.model_edit, 0)

        key_label = QLabel("API Key", form)
        key_label.setObjectName("dimLabel")
        form_layout.addWidget(key_label, 0)
        self.api_key_edit = QLineEdit(form)
        self.api_key_edit.setObjectName("keyLineEdit")
        self.api_key_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.api_key_edit.setPlaceholderText("请输入对应平台的 API Key")
        form_layout.addWidget(self.api_key_edit, 0)

        save_btn = QPushButton("保存配置", form)
        save_btn.setObjectName("accentButton")
        save_btn.clicked.connect(self._on_save_api_settings)
        form_layout.addWidget(save_btn, 0)

        tip = QLabel("提示：任何 OpenAI 兼容的多模态模型均可。切换预设会自动填入地址与模型，仍可手动修改。", form)
        tip.setObjectName("dimLabel")
        tip.setWordWrap(True)
        form_layout.addWidget(tip, 0)

        card_layout.addWidget(form, 0)
        card_layout.addItem(QSpacerItem(10, 10, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding))
        layout.addWidget(card, 1)
        self._load_api_settings_to_inputs()
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
            QMessageBox.warning(self, "缺少配置", f"请先到【设置】页面填写并保存：{fields}。")
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
        QMessageBox.information(self, "保存成功", "接口配置已保存。")
