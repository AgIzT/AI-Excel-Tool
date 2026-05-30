### 基于 AI 的全自动单据识别入库系统

![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)
![PySide6](https://img.shields.io/badge/Framework-PySide6-green.svg)
![License](https://img.shields.io/badge/License-MIT-yellow.svg)

> 解决手工录单转自动化。多张单据一键拖入，AI 自动提取信息并匹配导出自定义模板 Excel；导出前可人工复核，确保数据准确。


## ✨ 特性

* **一步识别**：图片与表格统一交给一个多模态（视觉）模型，一步提取供应商、日期与商品明细。
* **接口可自定义**：内置「智谱 GLM / OpenAI / 通义千问 / 自定义」预设，任何 OpenAI 兼容的多模态模型都能用。
* **导出前人工复核**：识别后先在可编辑表格中核对、修改，确认无误再导出——数据要进会计软件，准确性第一。
* **批量与模板**：多文件拖拽、手写体识别增强、合并输出、智能文件命名、内置与自定义 Excel 模板。


## 🛠️ 快速上手

### 1. 下载
前往 [Releases](https://github.com/AgIzT/AI-Excel-Tool/releases) 页面下载最新的 `default.exe`。

### 2. 配置
首次启动请进入左侧 **“设置”** 页面：
1.  选择 **接口预设**（智谱 GLM / OpenAI / 通义千问 / 自定义），会自动填入接口地址与模型，也可手动修改。
2.  填写对应平台的 **API Key**，点击保存。

> 任何 OpenAI 兼容的多模态（视觉）模型均可使用。

### 3. 录入
1.  切换至 **“单据处理”** 页面。
2.  将送货单图片（或 Excel/CSV）直接拖入识别列表，选择模板。
3.  点击 **“开始识别”**。
4.  在弹出的 **复核窗口** 核对、修改识别结果，确认后导出标准 Excel 表格。

---

##  版本

* **v2.0.0**：管线合并为单个多模态模型（移除 DeepSeek 两段式）；接口可自定义多模型；新增导出前人工复核；删除 AI 对话助手；代码模块化拆分。
* **v1.0.0 Reforged**：重构迁移至 PySide6，新增 AI 助手。
* **v0.3.0**：新增智能命名规则与多模版支持。
* **v0.2.0**：支持多文件拖拽与手写体识别增强。
* **v0.1.0 Beta**：GLM-OCR + DeepSeek + Excel 导出流程。


---
本项目遵循 [MIT License](LICENSE) 协议。
