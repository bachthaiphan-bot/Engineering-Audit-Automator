# 🏗️ 工程竣工决算文档生成工具

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)

> **一句话介绍**：程序批量生成全套文档仅需数分钟；包含人工核对在内，单个项目的报告编制时间由约1–2小时缩短至30–40分钟（资料齐全条件下的个人实测）。

[English](README.md) | 简体中文

---

## 🎯 解决什么问题？

工程审计人员每次项目结束都要手工填写5-6份格式固定的Word/Excel报告：
- 📄 报告封面
- 📋 审核报告
- 📊 审核表
- 📑 报告目录
- 📝 说明书
- 📈 决算报表

**痛点**：
- ⏰ 耗时长：每个项目约1–2小时
- ❌ 易出错：手工复制粘贴容易遗漏
- 😫 重复劳动：格式固定但必须逐个填写

**解决方案**：
本工具通过**模板引擎 + 数据映射**，从Excel数据源一键生成所有文档。程序批量生成全套文档仅需数分钟；包含人工核对在内，单个项目的报告编制时间由约1–2小时缩短至30–40分钟（资料齐全条件下的个人实测）。

---

## ✨ 核心功能

### 🚀 批量生成
- 一次处理多个项目
- 自动创建项目文件夹
- 按序号命名输出文件

### 🧹 智能清洗
- **日期格式化**：Excel序列号 → `2025年01月15日`
- **数值格式化**：`28041.07` → `28,041.07`（千位分隔符）
- **空值处理**：自动删除空值/零值行

### 📐 格式保持
- 完全保留模板样式（字体、颜色、对齐）
- 智能标题断行
- 表格结构完整

### ⚡ 性能优化
- 多线程并行处理
- 断点续跑支持
- GUI进度显示（exe版本）

---

## 📦 安装

### 环境要求
- Python 3.8 或更高版本
- Windows / macOS / Linux

### 安装依赖
```bash
pip install -r requirements.txt
```

---

## 🚀 快速开始

### 1️⃣ 准备数据
将Excel数据文件放入`待处理数据/`文件夹：
```
待处理数据/
  ├── 项目A_数据.xlsx
  ├── 项目B_数据.xlsm
  └── 项目C_数据.xls
```

### 2️⃣ 运行程序
```bash
python generate_docs_final.py
```

### 3️⃣ 查看结果
生成的文档在`生成结果/`文件夹：
```
生成结果/
  ├── XX核字[20XX]0000XXXX_项目A/
  │   ├── 1_报告封面_项目A.docx
  │   ├── 2_工程竣工决算审核报告_项目A.docx
  │   ├── 3_审核表_项目A.xlsx
  │   ├── 4_报告目录_项目A.docx
  │   ├── 5_竣工决算报告说明书_项目A.docx
  │   └── 6_基建工程竣工决算报表_项目A.xlsx
  └── ...
```

---

## 📖 详细文档

- 📘 [快速开始指南](QUICKSTART.md) - 5分钟上手
- 🏗️ [系统架构文档](ARCHITECTURE.md) - 技术细节
- 🤝 [贡献指南](CONTRIBUTING.md) - 如何参与开发
- 📋 [更新日志](CHANGELOG.md) - 版本历史
- 🚀 [发布指南](RELEASE_GUIDE.md) - GitHub发布流程

---

## 🔧 高级用法

### 启用调试模式
```bash
python generate_docs_final.py --debug
```

### 并行处理（加速）
```bash
python generate_docs_final.py --workers 4
```

### 断点续跑
```bash
python generate_docs_final.py --skip-existing
```

### 使用自定义数据文件
```bash
python generate_docs_final.py --data-file custom_data.xlsx
```

---

## 📁 项目结构

```
工程竣工决算文档生成工具/
├── generate_docs_final.py      # 主程序
├── auto_mapping.py              # 自动映射模块
├── requirements.txt             # 依赖清单
├── README.md                    # 项目文档
├── LICENSE                      # MIT许可证
│
├── 模板文件/                    # Word/Excel模板
│   ├── template_报告封面.docx
│   ├── template_审核报告.docx
│   ├── template_审核表.xlsx
│   ├── template_报告目录.docx
│   └── template_说明书.docx
│
├── 待处理数据/                  # 输入数据
│   └── (放置Excel文件)
│
└── 生成结果/                    # 输出文档
    ├── 项目A/
    ├── 项目B/
    └── logs/
```

---

## 🎓 技术栈

- **Python 3.8+** - 核心语言
- **python-docx** - Word文档处理
- **docxtpl** - Jinja2模板引擎
- **openpyxl** - Excel文件操作
- **pandas** - 数据处理
- **xlwings** - Excel自动化（可选，仅Windows）

---

## 🤝 贡献

欢迎提交Issue和Pull Request！

详见[贡献指南](CONTRIBUTING.md)

---

## 📄 许可证

本项目采用 [MIT许可证](LICENSE)

---

## 👤 作者

**Ethon Wong**

---

## 🙏 致谢

感谢以下开源项目：
- [python-docx](https://github.com/python-openxml/python-docx)
- [docxtpl](https://github.com/elapouya/python-docx-template)
- [openpyxl](https://openpyxl.readthedocs.io/)
- [pandas](https://pandas.pydata.org/)

---

## ⭐ Star History

如果这个项目对你有帮助，请给个Star⭐支持一下！

---

**让工程审计工作更高效！** 🚀
