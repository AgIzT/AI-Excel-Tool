全自动单据入库系统 (AI-Excel-Tool) 更新日志.README

\---
### 基于 AI 的全自动单据识别入库系统

![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)
![PySide6](https://img.shields.io/badge/Framework-PySide6-green.svg)
![License](https://img.shields.io/badge/License-MIT-yellow.svg)


> 解决手工录单自动化。支持多张单据一键拖入，AI 自动提取信息并匹配导出自定义模板 Excel 表格。



## 🛠️ 快速上手

1. **下载运行**：前往 [Releases](https://github.com/AgIzT/AI-Excel-Tool/releases) 下载最新的 `default.exe`。
2. **配置密钥**：在软件左侧“设置”页填入你的 GLM 和 DeepSeek API Key。
3. **开始录入**：将送货单图片直接拖入识别区。



> 这个项目最初是为了帮家里厂子里减轻录单负担而开发的。这个工具或者这种‘AI + 传统行业简单重复工作’的思路对你有启发，欢迎点个 Star 支持一下！
> 以下为更新日志

***

#### 🏁 **v1.0.0 AI-Excel-Tool 更新日志 "Reforged"** - 架构重构

架构：迁移至 PySide6 (Qt) 架构，异步处理与现代化模块化 UI，修复原UI界面卡顿。

AI助手：新增 AI 对话助手。

界面：现代化风格布局，弃用旧版长滚动条设计。

视觉：统一全局 QSS 样式。

***

#### 🏁 **v0.3 AI-Excel-Tool 更新日志** - 智能命名与模板管理

智能命名：AI 自动提取单据日期与供应商，生成「日期\_供应商\_入库.xlsx」智能文件名，支持冲突自动编号。

模板：新增支持多模板切换与自定义模板导入。

***

#### 🏁 **v0.2 AI-Excel-Tool 更新日志** - 功能完善

交互：支持多文件（图片/Excel）拖拽上传，新增列表预览及单文件删除。

识别：新增"手写体"开关。开启时优先提取手写涂改/备注；关闭则仅提取印刷体。

***

#### 🏁 **v0.1 AI-Excel-Tool Beta (初代版本)** - 核心功能架构

初代版本跑通了业务全流程，奠定 AI 自动化处理基础。

【基础架构】 GLM-OCR 识图提取 → DeepSeek 逻辑洗选与匹配 → 自动导出标准 Excel 表格。

【操作流】 左侧 01（选文件）- 02（AI 识别）- 03（打开导出文件）的三步工作流。

【现代化 UI】 采用深色极简主题，图片预览面板。

【内置模板】底层直接暂时封装进货单导入模板。

***
