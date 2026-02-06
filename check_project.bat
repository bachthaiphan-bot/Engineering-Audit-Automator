@echo off
chcp 65001 >nul
echo ================================================================================
echo                    工程竣工决算文档生成工具
echo                      项目完整性检查工具
echo ================================================================================
echo.

echo [1/5] 检查核心文档...
if exist README.md (echo ✅ README.md) else (echo ❌ README.md 缺失)
if exist LICENSE (echo ✅ LICENSE) else (echo ❌ LICENSE 缺失)
if exist requirements.txt (echo ✅ requirements.txt) else (echo ❌ requirements.txt 缺失)
if exist .gitignore (echo ✅ .gitignore) else (echo ❌ .gitignore 缺失)
echo.

echo [2/5] 检查补充文档...
if exist CONTRIBUTING.md (echo ✅ CONTRIBUTING.md) else (echo ⚠️ CONTRIBUTING.md 缺失)
if exist CHANGELOG.md (echo ✅ CHANGELOG.md) else (echo ⚠️ CHANGELOG.md 缺失)
if exist QUICKSTART.md (echo ✅ QUICKSTART.md) else (echo ⚠️ QUICKSTART.md 缺失)
if exist ARCHITECTURE.md (echo ✅ ARCHITECTURE.md) else (echo ⚠️ ARCHITECTURE.md 缺失)
echo.

echo [3/5] 检查核心代码...
if exist generate_docs_final.py (echo ✅ generate_docs_final.py) else (echo ❌ generate_docs_final.py 缺失)
if exist auto_mapping.py (echo ✅ auto_mapping.py) else (echo ❌ auto_mapping.py 缺失)
echo.

echo [4/5] 检查目录结构...
if exist 模板文件 (echo ✅ 模板文件/) else (echo ❌ 模板文件/ 缺失)
if exist 待处理数据 (echo ✅ 待处理数据/) else (echo ❌ 待处理数据/ 缺失)
if exist 生成结果 (echo ✅ 生成结果/) else (echo ❌ 生成结果/ 缺失)
echo.

echo [5/5] 检查Python环境...
python --version >nul 2>&1
if %errorlevel% equ 0 (
    echo ✅ Python已安装
    python --version
) else (
    echo ❌ Python未安装或未添加到PATH
)
echo.

echo ================================================================================
echo 检查完成！
echo.
echo 下一步：
echo   1. 如果所有核心项目都显示✅，可以继续发布流程
echo   2. 如果有❌项目，请先解决缺失的文件
echo   3. 阅读 GITHUB_CHECKLIST.md 完成发布前检查
echo   4. 按照 RELEASE_GUIDE.md 发布到GitHub
echo ================================================================================
echo.
pause
