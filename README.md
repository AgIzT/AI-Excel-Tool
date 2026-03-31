全自动单据入库系统 (AI-Excel-Tool) 更新日志.README

\---
### 基于 AI 的全自动单据识别入库系统

![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)
![PySide6](https://img.shields.io/badge/Framework-PySide6-green.svg)
![License](https://img.shields.io/badge/License-MIT-yellow.svg)

> 解决手工录单转自动化。支持多张单据一键拖入，AI 自动提取信息并匹配导出自定义模板 Excel 表格。


## 🛠️ 快速上手

### 1. 下载
前往 [Releases](https://github.com/AgIzT/AI-Excel-Tool/releases) 页面下载最新的 `default.exe`。

### 2. 配置
首次启动请进入左侧 **“设置”** 页面，填入您的 API 密钥：
* `GLM API Key` (智谱 AI)
* `DeepSeek API Key`

### 3. 录入
1.  切换至 **“单据处理”** 页面。
2.  将送货单图片直接拖入识别列表。
3.  点击 **“开始识别”**，等待 AI 自动导出标准 Excel 表格。

---
> 这个项目最初是为了帮家里厂子里减轻录单负担而开发的。这个工具或者这种‘AI + 传统行业简单重复工作’的思路对你有启发，欢迎点个 Star 支持一下！


---

## 📈 版本演进

* **v1.0.0 Reforged**：重构迁移至 PySide6，新增 AI 助手。
* **v0.3.0**：新增智能命名规则与多模版支持。
* **v0.2.0**：支持多文件拖拽与手写体识别增强。
* **v0.1.0 Beta**：GLM-OCR + DeepSeek + Excel 导出流程。


---

## ⚖️ 开源协议
本项目遵循 [MIT License](LICENSE) 协议。
