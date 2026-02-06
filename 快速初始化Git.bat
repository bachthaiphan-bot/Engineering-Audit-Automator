@echo off
chcp 65001 >nul
echo ================================================================
echo                   Git仓库快速初始化工具
echo ================================================================
echo.

echo [步骤1/3] 初始化Git仓库...
git init
if %errorlevel% neq 0 (
    echo ❌ Git初始化失败！请确保已安装Git。
    pause
    exit /b 1
)
echo ✅ Git仓库初始化成功
echo.

echo [步骤2/3] 添加所有文件到暂存区...
git add .
if %errorlevel% neq 0 (
    echo ❌ 添加文件失败！
    pause
    exit /b 1
)
echo ✅ 文件已添加到暂存区
echo.

echo [步骤3/3] 查看将要提交的文件...
git status
echo.

echo ================================================================
echo 准备提交！
echo ================================================================
echo.
echo 请输入提交信息（直接回车使用默认信息）：
set /p commit_msg="提交信息: "

if "%commit_msg%"=="" (
    set commit_msg=feat: 初始提交 - 工程竣工决算文档生成工具 v2.2
)

echo.
echo 正在提交: %commit_msg%
git commit -m "%commit_msg%"

if %errorlevel% neq 0 (
    echo ❌ 提交失败！
    pause
    exit /b 1
)

echo.
echo ================================================================
echo ✅ Git仓库初始化完成！
echo ================================================================
echo.
echo 下一步：
echo   1. 在GitHub上创建新仓库
echo   2. 运行以下命令关联远程仓库（替换YOUR_USERNAME）：
echo.
echo      git remote add origin https://github.com/YOUR_USERNAME/engineering-audit-doc-generator.git
echo      git branch -M main
echo      git push -u origin main
echo.
echo 详细步骤请查看: RELEASE_GUIDE.md
echo ================================================================
echo.
pause
