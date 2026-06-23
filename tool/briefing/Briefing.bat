@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo   KPC Morgenbriefing wird erstellt ...
echo ============================================
echo.
python briefing.py
echo.
pause
