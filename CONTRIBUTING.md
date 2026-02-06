# 贡献指南

感谢你考虑为本项目做出贡献！

## 开发环境设置

### 1. 克隆仓库
```bash
git clone <REPOSITORY_URL>
cd engineering-audit-doc-generator
```

### 2. 创建虚拟环境
```bash
python -m venv .venv
source .venv/bin/activate  # Linux/macOS
# 或
.venv\Scripts\activate  # Windows
```

### 3. 安装依赖
```bash
pip install -r requirements.txt
```

## 代码规范

### Python代码风格
- 遵循 PEP 8 规范
- 使用类型提示（Type Hints）
- 函数必须包含 Docstring
- 变量命名使用英文且具有描述性

### 示例
```python
def format_date_value(value: Any) -> str:
    """
    将日期值格式化为 yyyy年mm月dd日
    
    Args:
        value: 日期值（可以是datetime对象、Excel序列号或字符串）
    
    Returns:
        格式化后的日期字符串
    
    Examples:
        >>> format_date_value(datetime(2025, 1, 15))
        '2025年01月15日'
    """
    # 实现代码...
```

### 提交信息规范
使用语义化提交信息：
- `feat:` 新功能
- `fix:` 修复bug
- `docs:` 文档更新
- `style:` 代码格式调整
- `refactor:` 代码重构
- `test:` 测试相关
- `chore:` 构建/工具链相关

示例：
```
feat: 添加并行处理支持
fix: 修复日期字段误识别问题
docs: 更新README安装说明
```

## 提交Pull Request流程

1. **Fork仓库**并创建你的分支
   ```bash
   git checkout -b feature/my-new-feature
   ```

2. **编写代码**并确保：
   - 代码通过所有测试
   - 添加必要的注释和文档
   - 遵循代码规范

3. **提交更改**
   ```bash
   git add .
   git commit -m "feat: 添加新功能描述"
   ```

4. **推送到你的Fork**
   ```bash
   git push origin feature/my-new-feature
   ```

5. **创建Pull Request**
   - 在GitHub上打开Pull Request
   - 清晰描述你的更改
   - 关联相关的Issue（如果有）

## 报告Bug

使用GitHub Issues报告bug时，请包含：
- **问题描述**：清晰简洁的描述
- **复现步骤**：如何触发这个bug
- **预期行为**：你期望发生什么
- **实际行为**：实际发生了什么
- **环境信息**：
  - 操作系统
  - Python版本
  - 相关依赖版本
- **日志/截图**：如果适用

## 功能请求

使用GitHub Issues提交功能请求时，请包含：
- **功能描述**：你希望添加什么功能
- **使用场景**：为什么需要这个功能
- **建议实现**：如果有想法，描述如何实现

## 代码审查

所有提交都需要经过代码审查。审查者会关注：
- 代码质量和可读性
- 是否遵循项目规范
- 是否有充分的测试
- 是否有必要的文档

## 许可证

提交代码即表示你同意你的贡献将在MIT许可证下发布。
