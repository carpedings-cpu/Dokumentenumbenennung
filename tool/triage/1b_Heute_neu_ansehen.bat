@echo off
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul && (set "PY=py") || (set "PY=python")
echo ============================================
echo   KPC Triage - HEUTE erneut ansehen
echo   (nur Uebersicht, legt nichts ab)
echo ============================================
echo.
%PY% triage.py --heute
echo.
echo Fertig. Die Uebersicht sollte sich im Browser geoeffnet haben.
echo.
pause
