@echo off
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul && (set "PY=py") || (set "PY=python")
echo ============================================
echo   KPC Triage - HEUTE auswaehlen
echo   Zeigt ALLE heutigen Mails erneut zur Auswahl.
echo ============================================
echo.
%PY% triage_gui.py --heute
