![Business Value Card](https://github.com/joanna-sym/Business-Solution-Sales-Tender-Compliance-Check/blob/main/assets/Github%20Banner-Joanna%20Shen.jpg)

# ❤️ 阿九Joanna · 标书合规排雷引擎 (MedOps Engine)

> **"将医疗器械标书初筛从 4 小时缩短至 30 秒。"**
> **"From 4 hours to 30 seconds for MedTech tender auditing."**

---

## 🌟 核心价值看板 (Business Value)

![Business Value Card](https://github.com/joanna-sym/Business-Solution-Sales-Tender-Compliance-Check/blob/main/assets/image01.png)

### 🇨🇳 中文介绍
在医疗器械招投标过程中，500+ 页的招标文件隐藏着无数“控标陷阱”与“废标条款（★项）”。本项目是 **MedOps (医疗运营)** 体系的首个自动化方案，旨在解决以下痛点：
* **效率瓶颈**：人工核对几百项参数，极其耗时且疲劳。
* **废标风险**：肉眼漏看一行小字，可能导致整个省份的市场丢失。
* **数据孤岛**：注册证参数与标书要求的比对往往依赖跨部门反复口头确认。

**核心逻辑**：利用 Python 结构化解析技术，自动对标，红色高亮风险，实现“秒级排雷”。

### 🇺🇸 English Introduction
This automated engine is designed for the MedTech industry to solve the high-risk, low-efficiency problem in tender document review. It automatically extracts technical parameters and audits them against internal product specifications.

---

## 🚀 演示与架构 (Demo & Architecture)

![Operation Demo GIF](https://github.com/joanna-sym/Business-Solution-Sales-Tender-Compliance-Check/blob/main/assets/cap%2020260130.gif)

![Tech Structure Card](https://github.com/joanna-sym/Business-Solution-Sales-Tender-Compliance-Check/blob/main/assets/image02.png)

### 📂 核心功能 (Key Features)
* **Step 00: Mock Data** - 一键生成测试用的模拟标书与产品库。
* **Step 01: Core Engine** - 纯 Python 算法层，处理复杂的参数比对逻辑。
* **Step 02: Content OS** - 基于 Streamlit 的可视化界面，蓝橙配色，专业感十足。
* **🚀 One-Click Start** - 提供 `.bat` 批处理脚本，双击即用，无需配置环境。

---

## 🛠 复盘与技术笔记 (Reflection & Post-Mortem)

在从 0 到 1 的开发中，我沉淀了以下 MedOps 数字化转型经验：

### 1. 踩过的坑 (The "Pits")
* **Windows 编码之坑**：最初终端无法显示 Emoji 导致崩溃，随后通过 Streamlit GUI 彻底绕过底层编码限制，提升了系统的稳定性。
* **路径管理**：为了实现“解压即用”，引入了 `os.path` 动态定位技术，解决了不同电脑环境下找不到头像资源的问题。

### 2. 成功的弯路 (The "Detours")
* **技术降级**：最初想做通用的 PDF 解析，但考虑到医疗数据的 **100% 准确性** 要求，最终决定让用户先利用 WPS/Adobe 将 PDF 转为 Word，将算法核心聚焦在“逻辑比对”而非“模糊识别”。

### 3. 未来路线 (Roadmap)
* [ ] 接入 Gemini  实现模糊语义理解（如：识别“钉仓”与“组件”的语义统一）。
* [ ] 增加多产品线批量比对模式。

---

## 🏁 如何开始 (Getting Started)

1. **环境安装**: `pip install -r requirements.txt`
2. **生成数据**: 运行 `python step00_generate_mock_data.py`
3. **一键启动**: 双击运行 **`阿九Joanna标书合规排雷引擎启动.bat`**

---

## 👤 作者 (Author)
**阿九 Joanna** (Medical Device Professional / Python Enthusiast)
*专注医械行业数字化效率提升 | MedTech Digital Matrix*

---
*MIT License © 2026 阿九Joanna*
