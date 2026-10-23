# 🏗️ 工程竣工决算文档生成工具

> **30秒电梯演讲**：工程审计人员每次项目结束都要手工填写5-6份格式固定的Word/Excel报告，重复劳动耗时且易错。本工具通过模板引擎 + 数据映射，**一键批量生成**所有竣工决算文档。程序批量生成全套文档仅需数分钟；包含人工核对在内，单个项目的报告编制时间由约1–2小时缩短至30–40分钟（资料齐全条件下的个人实测）。

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)

---

## ✨ 核心功能

- **🚀 批量生成**：从Excel数据源一次性生成多个项目的全套文档
- **📋 智能映射**：自动识别Excel中的字段并映射到Word/Excel模板
- **🧹 数据清洗**：
  - 日期自动格式化（支持Excel序列号 → `yyyy年mm月dd日`）
  - 数值千位分隔符（`28041.07` → `28,041.07`）
  - 空值/零值行自动删除（避免报告中出现无意义的空行）
- **📐 格式保持**：生成的文档完全保留模板样式（字体、颜色、对齐方式）
- **⚡ 并行处理**：支持多线程加速（可选）

---

## 📦 环境依赖

### 必需
- **Python**: 3.8 或更高版本
- **操作系统**: Windows / macOS / Linux

### Python 库
```bash
pip install -r requirements.txt
```

核心依赖：
- `python-docx` - Word文档处理
- `docxtpl` - Word模板渲染引擎
- `openpyxl` - Excel文件读写
- `pandas` - 数据处理

可选依赖：
- `xlwings` - Excel自动化（仅Windows，用于公式计算）
- `xlrd` - 旧版Excel格式支持（.xls）

---

## 🚀 快速开始

### 1. 准备数据文件
将Excel数据文件放入 `待处理数据/` 文件夹：
```
待处理数据/
  └── 项目数据_2025.xlsx
```

### 2. 运行生成程序
```bash
python generate_docs_final.py
```

### 3. 查看生成结果
生成的文档位于 `生成结果/` 文件夹，按项目分组：
```
生成结果/
  ├── XXXX核字[20XX]00001XX9号_项目A/
  │   ├── 1_报告封面_项目A.docx
  │   ├── 2_工程竣工决算审核报告_项目A.docx
  │   ├── 3_审核表_项目A.xlsx
  │   └── ...
  └── logs/
      └── app.log
```

---

## 📁 项目结构

```
工程竣工决算文档生成工具/
├── generate_docs_final.py      # 主程序：文档生成引擎
├── auto_mapping.py              # 自动映射：从Excel提取数据
├── analyze_cover.py             # 工具：分析模板结构
├── check_template.py            # 工具：检查模板占位符
├── requirements.txt             # Python依赖清单
├── README.md                    # 项目说明文档
├── LICENSE                      # MIT开源协议
│
├── 模板文件/                    # Word/Excel模板（含占位符）
│   ├── template_报告封面.docx
│   ├── template_审核报告.docx
│   ├── template_审核表.xlsx
│   ├── template_报告目录.docx
│   └── template_说明书.docx
│
├── 待处理数据/                  # 输入：Excel数据源
│   └── (放置项目数据文件)
│
└── 生成结果/                    # 输出：生成的文档
    ├── 项目A/
    ├── 项目B/
    └── logs/
```

---

## 🔧 高级用法

### 命令行参数
```bash
# 启用调试模式（详细日志）
python generate_docs_final.py --debug

# 使用自定义数据文件
python generate_docs_final.py --data-file custom_data.xlsx

# 并行处理（4线程）
python generate_docs_final.py --workers 4

# 断点续跑（跳过已生成的项目）
python generate_docs_final.py --skip-existing
```

### 自定义模板
1. 在Word/Excel模板中使用 `{{字段名}}` 作为占位符
2. 确保Excel数据源的列名与占位符一致
3. 运行程序自动替换

---

## 🤝 贡献指南

欢迎提交Issue和Pull Request！

1. Fork本仓库
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启Pull Request

---

## 📄 许可证

本项目采用 MIT 许可证 - 详见 [LICENSE](LICENSE) 文件

---

## 👤 作者

**Ethon Wong**

如有问题或建议，欢迎通过Issue联系。

---

## 🙏 致谢

感谢以下开源项目：
- [python-docx](https://github.com/python-openxml/python-docx) - Word文档处理
- [docxtpl](https://github.com/elapouya/python-docx-template) - Jinja2模板引擎
- [openpyxl](https://openpyxl.readthedocs.io/) - Excel文件操作
