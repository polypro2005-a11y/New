@echo off
cd /d D:\ai-stack\apps\hermes-webui
echo 🤖 Telegram bot polling daemon
echo ================================
start /B /MIN pythonw -c "import sys; sys.argv=['','--polling']; exec(open(r'C:\Users\Admin\workspace\email_analysis_automation.py','r',encoding='utf-8').read())"
echo ✅ Polling запущен (pythonw, скрытое окно)
echo.
echo Для остановки: taskkill /f /im pythonw.exe
