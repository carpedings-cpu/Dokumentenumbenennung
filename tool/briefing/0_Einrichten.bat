@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo   KPC Morgenbriefing - Einrichten (einmalig)
echo ============================================
echo.
echo Installiert die noetigen Python-Teile (pywin32 fuer Outlook).
echo.
python -m pip install --upgrade pywin32
echo.
if not exist ".env" (
  copy ".env.example" ".env" >nul
  echo .env wurde angelegt - bitte oeffnen und den Gemini-Schluessel eintragen.
) else (
  echo .env ist bereits vorhanden.
)
echo.
echo Fertig. Jetzt .env oeffnen, Schluessel eintragen, dann "Briefing" starten.
pause
