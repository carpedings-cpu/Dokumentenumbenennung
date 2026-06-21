@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo   KPC Triage - Einrichtung (nur EINMAL noetig)
echo ============================================
echo.
where py >nul 2>nul && (set "PY=py") || (set "PY=python")
echo Installiere die benoetigten Komponenten ...
%PY% -m pip install --upgrade pywin32 anthropic
echo.
echo Fertig. Du kannst jetzt "1_Probelauf.bat" doppelklicken.
echo.
pause
