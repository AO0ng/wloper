@echo off
chcp 65001 >nul
echo ============================================
echo   水文周报抓取 - 首次配置
echo ============================================
echo.

cd /d "%~dp0"

echo [1/3] 安装 Playwright 浏览器内核 ^(Chromium^)...
echo   这一步需要下载约 150MB，请耐心等待...
call npx playwright install chromium
if %ERRORLEVEL% NEQ 0 (
    echo [错误] 浏览器安装失败，请检查网络连接后重试。
    pause
    exit /b 1
)
echo    [完成]
echo.

echo [2/3] 验证 Node.js 依赖...
call npm install
echo    [完成]
echo.

echo [3/3] 测试运行...
echo    （首次运行会启动无头浏览器抓取数据）
call node scraper.js
echo.
echo ============================================
echo   配置完成！
echo   请编辑 scraper.js 顶部的 CONFIG 配置区：
echo   - referenceDate: 基准日期
echo   - stations: 站点名称列表
echo   - stationTypes: 站类列表
echo   然后按 README 配置任务计划程序即可。
echo ============================================
pause
