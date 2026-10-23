# 快速开始指南

本指南将帮助你在5分钟内完成首次文档生成。

## 前置条件

- ✅ 已安装Python 3.8或更高版本
- ✅ 已安装pip包管理器

## 步骤1：安装依赖（2分钟）

打开命令行，进入项目目录：

```bash
# Windows
cd path\to\engineering-audit-doc-generator
pip install -r requirements.txt

# macOS/Linux
cd path/to/engineering-audit-doc-generator
pip3 install -r requirements.txt
```

## 步骤2：准备数据（1分钟）

### 方式A：使用自动映射（推荐）

将你的Excel项目数据文件放入`待处理数据/`文件夹：

```
待处理数据/
  ├── 项目A_数据.xlsx
  ├── 项目B_数据.xlsm
  └── 项目C_数据.xls
```

程序会自动从这些文件中提取数据。

### 方式B：手动准备date.xlsx

如果你已经有一个汇总的Excel文件，将其重命名为`date.xlsx`并放在项目根目录。

**Excel格式要求**：
- 第一行：字段名（如"项目名称"、"项目编码"等）
- 后续行：每行代表一个项目的数据

## 步骤3：运行程序（1分钟）

```bash
python generate_docs_final.py
```

你会看到类似的输出：

```
=== 工程竣工决算文档生成程序 ===
已读取 5 行数据
加载主模板: 模板文件/template_审核报告.docx
输出目录: 生成结果
开始处理第 1 行数据
主文档生成完成: 生成结果/项目A/2_工程竣工决算审核报告_项目A.docx
额外文档生成完成: 生成结果/项目A/1_报告封面_项目A.docx
...
```

## 步骤4：查看结果（1分钟）

打开`生成结果/`文件夹，你会看到按项目分组的文档：

```
生成结果/
  ├── XX核字[20XX]0000XXXX_项目A/
  │   ├── 1_报告封面_项目A.docx
  │   ├── 2_工程竣工决算审核报告_项目A.docx
  │   ├── 3_审核表_项目A.xlsx
  │   ├── 4_报告目录_项目A.docx
  │   ├── 5_竣工决算报告说明书_项目A.docx
  │   └── 6_基建工程竣工决算报表_项目A.xlsx
  │
  ├── XX核字[20XX]0000XXXX_项目B/
  │   └── ...
  │
  └── logs/
      └── app.log  ← 详细日志
```

## 常见问题

### Q1: 提示"未找到数据文件 date.xlsx"

**解决方案**：
- 确保`待处理数据/`文件夹中有Excel文件，或
- 手动创建`date.xlsx`文件

### Q2: 生成的文档中有些字段显示为空

**原因**：Excel数据源中对应字段为空

**解决方案**：
- 检查Excel数据源，确保必填字段有值
- 查看`生成结果/logs/app.log`日志，搜索"缺少数据字段"

### Q3: 日期显示不正确

**原因**：Excel中的日期格式不标准

**解决方案**：
- 确保日期字段使用Excel的日期格式（不是文本）
- 或使用标准格式：`2025-01-15`、`2025/01/15`、`2025年01月15日`

### Q4: 数值没有千位分隔符

**原因**：字段名不包含数值关键词（如"金额"、"费用"等）

**解决方案**：
- 修改Excel列名，添加"金额"、"费用"等关键词
- 或在代码中手动添加格式化规则

### Q5: 程序运行很慢

**解决方案**：
- 启用并行处理：`python generate_docs_final.py --workers 4`
- 跳过已生成项目：`python generate_docs_final.py --skip-existing`

## 下一步

- 📖 阅读[README.md](README.md)了解完整功能
- 🏗️ 阅读[ARCHITECTURE.md](ARCHITECTURE.md)了解系统架构
- 🛠️ 自定义模板：修改`模板文件/`中的Word/Excel文件
- 🚀 高级用法：查看命令行参数 `python generate_docs_final.py --help`

## 获取帮助

- 📝 查看日志：`生成结果/logs/app.log`
- 🐛 报告问题：在GitHub上创建Issue
- 💬 讨论交流：在GitHub Discussions中提问
