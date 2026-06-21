@echo off
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul && (set "PY=py") || (set "PY=python")
echo ============================================
echo   KPC Triage - PROBELAUF (nur ansehen)
echo   Es wird NICHTS abgelegt oder veraendert.
echo ============================================
echo.
%PY% triage.py
echo.
echo Probelauf fertig. Die Uebersicht sollte sich im Browser geoeffnet haben.
echo (Ablageordner: Triage-Berichte)
echo.
pause
