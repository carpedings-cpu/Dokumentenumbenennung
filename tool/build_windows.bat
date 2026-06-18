@echo off
REM ===================================================================
REM  Baut den Dokumenten-Umbenenner als eigenstaendige Windows-.exe.
REM  Voraussetzung: Python 3.9+ installiert (von python.org).
REM  Doppelklick auf diese Datei genuegt.
REM  Ergebnis:  dist\Dokumenten-Umbenenner.exe
REM ===================================================================

cd /d "%~dp0\.."

echo.
echo [1/2] Installiere PyInstaller, anthropic, tkinterdnd2 und extract-msg ...
python -m pip install --upgrade pyinstaller anthropic tkinterdnd2 extract-msg || goto :fehler

echo.
echo [2/2] Baue die EXE ...
python -m PyInstaller --noconfirm dokumenten-umbenenner.spec || goto :fehler

echo.
echo Fertig! Die Datei liegt hier:
echo   %cd%\dist\Dokumenten-Umbenenner.exe
echo.
pause
exit /b 0

:fehler
echo.
echo Beim Bauen ist ein Fehler aufgetreten. Bitte die Meldungen oben pruefen.
echo.
pause
exit /b 1
