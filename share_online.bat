@echo off
chcp 65001 > nul
echo =================================================================
echo       AI BOQ Pro - ส่งออกลิงก์ออนไลน์ให้เพื่อนทดสอบทันที
echo       (Secure HTTPS via Cloudflare Tunnel)
echo =================================================================
echo.
cd /d "%~dp0"

echo 1. กำลังเปิดโปรแกรม Streamlit ในเครื่อง...
start /B "" ".venv\Scripts\streamlit.exe" run app.py --server.headless true --server.port 8501

echo 2. รอระบบเริ่มทำงาน 3 วินาที...
timeout /t 3 /nobreak > nul

echo 3. กำลังสร้างลิงก์ออนไลน์สาธารณะ (HTTPS) ให้เพื่อนกดเข้า...
echo =================================================================
echo สังเกตบรรทัดที่มีข้อความ "...trycloudflare.com"
echo คัดลอกลิงก์นั้น แล้วส่งให้เพื่อนเปิดได้ทันทีจากมือถือหรือคอมพิวเตอร์!
echo (รหัสผ่านเข้าเว็บเริ่มต้นคือ: boq2026)
echo =================================================================
echo.
cloudflared.exe tunnel --url http://localhost:8501
pause
