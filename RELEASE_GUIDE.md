# GitHub发布指南

本指南将帮助你将项目发布到GitHub。

## 准备工作

### 1. 创建GitHub账号
如果还没有GitHub账号，访问 https://github.com 注册。

### 2. 安装Git
- **Windows**: 下载 [Git for Windows](https://git-scm.com/download/win)
- **macOS**: `brew install git`
- **Linux**: `sudo apt-get install git` 或 `sudo yum install git`

### 3. 配置Git
```bash
git config --global user.name "Ethon Wong"
git config --global user.email "your-email@example.com"
```

## 发布步骤

### 步骤1：创建GitHub仓库

1. 登录GitHub
2. 点击右上角 `+` → `New repository`
3. 填写仓库信息：
   - **Repository name**: `engineering-audit-doc-generator`
   - **Description**: `🏗️ 工程竣工决算文档自动化生成工具 - 批量生成Word/Excel审计报告`
   - **Visibility**: Public（公开）或 Private（私有）
   - **不要**勾选 "Initialize this repository with a README"（我们已经有了）
4. 点击 `Create repository`

### 步骤2：初始化本地仓库

在项目根目录打开命令行：

```bash
# 初始化Git仓库
git init

# 添加所有文件到暂存区
git add .

# 查看将要提交的文件（可选）
git status

# 提交到本地仓库
git commit -m "feat: 初始提交 - 工程竣工决算文档生成工具 v2.2"
```

### 步骤3：关联远程仓库

将`<YOUR_USERNAME>`替换为你的GitHub用户名：

```bash
git remote add origin https://github.com/<YOUR_USERNAME>/engineering-audit-doc-generator.git

# 验证远程仓库
git remote -v
```

### 步骤4：推送到GitHub

```bash
# 推送到main分支
git branch -M main
git push -u origin main
```

如果提示输入用户名和密码：
- **用户名**: 你的GitHub用户名
- **密码**: 使用Personal Access Token（不是账号密码）

#### 创建Personal Access Token
1. GitHub → Settings → Developer settings → Personal access tokens → Tokens (classic)
2. Generate new token → 勾选 `repo` 权限
3. 复制生成的token（只显示一次！）
4. 在命令行中粘贴token作为密码

### 步骤5：验证发布

访问 `https://github.com/<YOUR_USERNAME>/engineering-audit-doc-generator`

你应该能看到：
- ✅ README.md自动显示在首页
- ✅ 所有代码文件
- ✅ LICENSE文件
- ✅ 完整的项目结构

## 后续维护

### 更新代码

```bash
# 修改代码后...

# 查看更改
git status

# 添加更改
git add .

# 提交更改
git commit -m "fix: 修复日期识别bug"

# 推送到GitHub
git push
```

### 创建Release版本

1. GitHub仓库页面 → Releases → Create a new release
2. 填写信息：
   - **Tag version**: `v2.2.0`
   - **Release title**: `v2.2.0 - 自动映射与并行处理`
   - **Description**: 从CHANGELOG.md复制更新内容
3. 上传打包的exe文件（可选）
4. 点击 `Publish release`

### 添加徽章（Badges）

在README.md顶部添加：

```markdown
[![GitHub release](https://img.shields.io/github/release/<YOUR_USERNAME>/engineering-audit-doc-generator.svg)](https://github.com/<YOUR_USERNAME>/engineering-audit-doc-generator/releases)
[![GitHub stars](https://img.shields.io/github/stars/<YOUR_USERNAME>/engineering-audit-doc-generator.svg)](https://github.com/<YOUR_USERNAME>/engineering-audit-doc-generator/stargazers)
[![GitHub issues](https://img.shields.io/github/issues/<YOUR_USERNAME>/engineering-audit-doc-generator.svg)](https://github.com/<YOUR_USERNAME>/engineering-audit-doc-generator/issues)
```

## 推广项目

### 1. 完善README
- 添加项目截图/演示GIF
- 添加使用案例
- 添加常见问题解答

### 2. 添加Topics
GitHub仓库页面 → About → Topics → 添加：
- `python`
- `document-generation`
- `automation`
- `excel`
- `word`
- `audit`
- `engineering`

### 3. 分享到社区
- 知乎：写一篇技术文章
- CSDN/博客园：发布使用教程
- Python中文社区：分享项目链接

### 4. 持续改进
- 及时回复Issues
- 接受Pull Requests
- 定期发布新版本
- 更新文档

## 常见问题

### Q: 推送时提示"Permission denied"

**解决方案**：
1. 检查Personal Access Token是否有`repo`权限
2. 使用SSH密钥代替HTTPS：
   ```bash
   git remote set-url origin git@github.com:<YOUR_USERNAME>/engineering-audit-doc-generator.git
   ```

### Q: 如何删除已提交的敏感文件？

**解决方案**：
```bash
# 从Git历史中删除文件
git filter-branch --force --index-filter \
  "git rm --cached --ignore-unmatch 敏感文件.xlsx" \
  --prune-empty --tag-name-filter cat -- --all

# 强制推送
git push origin --force --all
```

### Q: 如何回滚到之前的版本？

**解决方案**：
```bash
# 查看提交历史
git log --oneline

# 回滚到指定提交
git reset --hard <commit-hash>

# 强制推送
git push -f origin main
```

## 安全提示

⚠️ **发布前务必检查**：
- [ ] 没有硬编码的密码/API密钥
- [ ] 没有真实的项目数据
- [ ] 没有个人敏感信息
- [ ] .gitignore正确配置
- [ ] 所有敏感文件已被忽略

## 下一步

发布成功后，考虑：
1. 📝 撰写博客文章介绍项目
2. 🎥 录制使用演示视频
3. 📊 添加项目统计（下载量、Star数）
4. 🌐 创建项目官网（GitHub Pages）
5. 📦 发布到PyPI（`pip install`支持）

---

**祝贺你！项目已准备好与世界分享！** 🎉
