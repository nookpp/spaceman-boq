@echo off
chcp 65001 > nul
echo ========================================================
echo       AI BOQ Pro - ระบบสร้างและถอด BOQ ด้วย AI
echo       (Streamlit + OpenCode Go)
echo ========================================================
echo.
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo [ERROR] Virtual Environment not found!
    echo Please install dependencies first.
    pause
    exit /b
)

echo กำลังเปิดโปรแกรมผ่านเว็บเบราว์เซอร์...
echo เพื่อให้เพื่อนในวงแลนเดียวกันเข้าใช้งานได้ ให้ใช้ URL: Network URL ด้านล่าง
echo.
".venv\Scripts\streamlit.exe" run app.py --server.address 0.0.0.0 --server.port 8501
pause
