# 项目架构文档

## 系统概述

本工具是一个基于Python的文档自动化生成系统，专为工程审计场景设计。核心思想是**模板 + 数据 = 文档**。

## 核心模块

### 1. `generate_docs_final.py` - 主程序
**职责**：文档生成引擎，协调所有模块完成批量渲染

**核心流程**：
```
读取Excel数据源
  ↓
遍历每一行（每个项目）
  ↓
构建渲染上下文（数据清洗 + 格式化）
  ↓
加载Word/Excel模板
  ↓
替换占位符 → 生成文档
  ↓
后处理（删除空行、标题断行等）
  ↓
保存到输出目录
```

**关键函数**：
- `main()` - 程序入口，处理命令行参数
- `build_context()` - 数据清洗与格式化
- `render_excel_template()` - Excel模板渲染
- `cleanup_document_rows_by_context()` - 文档后处理

### 2. `auto_mapping.py` - 自动映射模块
**职责**：从`待处理数据/`文件夹中的Excel文件自动提取数据并生成`date.xlsx`

**工作原理**：
1. 扫描文件夹中的所有Excel文件（.xlsx/.xlsm/.xls）
2. 从每个文件的隐藏工作表`_calc`中读取P列（标题）和Q列（数据）
3. 合并所有项目数据到一个统一的Excel文件

**关键函数**：
- `read_excel_data_from_data_folder()` - 读取所有Excel文件
- `create_date_xlsx_from_data()` - 生成date.xlsx
- `safe_read_cell_value()` - 安全读取单元格（避免类型误判）

### 3. `analyze_cover.py` - 模板分析工具
**职责**：分析Word模板的内部结构，帮助开发者理解模板布局

**用途**：
- 调试模板占位符位置
- 检查段落和表格结构
- 定位样式问题

### 4. `check_template.py` - 模板检查工具
**职责**：检查模板中的占位符和runs结构

## 数据流图

```
┌─────────────────┐
│  待处理数据/     │
│  ├─ 项目A.xlsx  │
│  ├─ 项目B.xlsx  │
│  └─ 项目C.xlsx  │
└────────┬────────┘
         │ auto_mapping.py
         ↓
┌─────────────────┐
│   date.xlsx     │ ← 统一数据源
└────────┬────────┘
         │ generate_docs_final.py
         ↓
┌─────────────────┐
│   模板文件/      │
│  ├─ 封面.docx   │
│  ├─ 报告.docx   │
│  └─ 审核表.xlsx │
└────────┬────────┘
         │ docxtpl + openpyxl
         ↓
┌─────────────────┐
│   生成结果/      │
│  ├─ 项目A/      │
│  │  ├─ 封面.docx│
│  │  ├─ 报告.docx│
│  │  └─ 审核表.xlsx│
│  └─ 项目B/      │
└─────────────────┘
```

## 关键技术

### 1. 模板引擎
- **Word**: 使用`docxtpl`（基于Jinja2）
- **Excel**: 使用`openpyxl`进行占位符替换

### 2. 数据清洗
**日期识别**：
- 字段名启发式（包含"日期"、"时间"等关键词）
- Excel序列号范围检测（40179-47482，对应2010-2030）
- 避免金额字段被误识别

**数值格式化**：
- 千位分隔符：`28041.07` → `28,041.07`
- 统一两位小数

### 3. 并行处理
- 文档级并行：`--workers` 参数控制单个项目内的模板并行渲染
- 项目级并行：`--projects-workers` 参数控制多个项目同时处理
- 注意：并行模式下禁用Excel COM和GUI（避免冲突）

### 4. 错误处理
- 日志系统：控制台 + 文件（`生成结果/logs/app.log`）
- 优雅降级：Excel COM失败时回退到openpyxl
- 断点续跑：`--skip-existing`跳过已生成项目

## 扩展点

### 添加新模板
1. 在`模板文件/`中创建新模板（使用`{{字段名}}`占位符）
2. 在`main()`函数的`extra_templates`字典中注册
3. 运行程序自动渲染

### 添加新的数据清洗规则
在`build_context()`函数中添加自定义逻辑：
```python
# 示例：自定义字段处理
if key == "特殊字段":
    context[key] = custom_format(value)
```

### 添加新的后处理规则
在`cleanup_document_rows_by_context()`中添加规则：
```python
# 示例：删除特定条件的行
if _is_empty_or_zero(context.get("某字段")):
    _delete_table_row_by_first_cell_text(doc_path, "行标签", logger)
```

## 性能优化

### 当前性能
- 单项目处理时间：约10-15秒
- 批量处理（100个项目）：约20-30分钟
- 并行加速比：2-3x（取决于CPU核心数）

### 优化建议
1. **启用并行**：`--workers 4 --projects-workers 2`
2. **跳过已生成**：`--skip-existing`（断点续跑）
3. **禁用Excel COM**：如果不需要公式计算，可以移除xlwings依赖

## 安全考虑

### 数据隐私
- 所有数据文件应添加到`.gitignore`
- 不要将包含真实项目数据的Excel文件提交到版本控制

### 文件路径
- 使用相对路径，避免硬编码绝对路径
- 使用`os.path.join()`确保跨平台兼容性

### 依赖安全
- 定期更新依赖库（`pip install --upgrade -r requirements.txt`）
- 使用虚拟环境隔离项目依赖
