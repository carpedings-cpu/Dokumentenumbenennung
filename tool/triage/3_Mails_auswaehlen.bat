@echo off
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul && (set "PY=py") || (set "PY=python")
echo ============================================
echo   KPC Triage - Mails auswaehlen
echo   Liste oeffnet sich: anhaken, was umbenannt
echo   werden soll. Outlook wird nur gelesen.
echo ============================================
echo.
%PY% triage_gui.py
