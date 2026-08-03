@echo off
chcp 65001
echo ========================================
echo 🎯 HR个性化PD生成器 - 黑客松Demo版
echo ========================================
echo.
echo ✨ 特点:
echo   - 预加载示例数据
echo   - 无需API密钥
echo   - 30秒完整演示
echo   - 100%% 离线运行
echo.
echo ========================================
echo.
echo 正在启动程序...
echo 浏览器将自动打开 http://localhost:8501
echo.
echo 💡 演示流程:
echo    1. 点击"使用示例数据快速演示"
echo    2. 查看文件解析结果
echo    3. 点击"AI智能分析并生成"
echo    4. 查看完整的个性化计划
echo.
echo 按 Ctrl+C 可以停止程序
echo ========================================
echo.

python -m streamlit run app_demo_offline.py

pause
