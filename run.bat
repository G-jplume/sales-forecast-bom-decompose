@echo off
chcp 65001 >nul
title 销售预测捆绑SKU拆解工具
cd /d "%~dp0"
echo ========================================
echo   销售预测捆绑SKU拆解 Streamlit 小程序
echo ========================================
echo.
echo 正在启动 Web 服务...
echo 浏览器将自动打开 http://localhost:8501
echo.
echo 按 Ctrl+C 可停止服务
echo.
python -m streamlit run app.py --server.headless=false --browser.gatherUsage=false
pause
