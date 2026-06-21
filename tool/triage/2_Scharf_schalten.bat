@echo off
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul && (set "PY=py") || (set "PY=python")
echo ============================================
echo   KPC Triage - SCHARF SCHALTEN
echo   Relevante Mails werden als .msg ins
echo   Eingangs-Koerbchen (00_Posteingang) gelegt.
echo   Outlook wird weiterhin NUR gelesen.
echo ============================================
echo.
set /p OK=Wirklich scharf schalten? (J/N):
if /I not "%OK%"=="J" goto :abbruch
echo.
%PY% triage.py --scharf
echo.
echo Scharfer Lauf fertig.
echo.
pause
exit /b 0

:abbruch
echo.
echo Abgebrochen - es wurde nichts abgelegt.
echo.
pause
