@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo   KPC Triage - Aktualisieren (neueste Version laden)
echo ============================================
echo.
echo Laedt die aktuelle Version von GitHub und ersetzt nur die Programmdateien.
echo Deine eigenen Dateien bleiben unangetastet:
echo   .env, projekte_mapping.json, triage_state.json
echo.
set "URL=https://github.com/carpedings-cpu/Dokumentenumbenennung/archive/refs/heads/claude/amazing-feynman-eimedk.zip"
powershell -NoProfile -ExecutionPolicy Bypass -Command "try{ $ErrorActionPreference='Stop'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; $z=Join-Path $env:TEMP 'kpc_triage.zip'; $d=Join-Path $env:TEMP 'kpc_triage_x'; Invoke-WebRequest -Uri $env:URL -OutFile $z; if(Test-Path $d){Remove-Item $d -Recurse -Force}; Expand-Archive $z $d -Force; $t=Join-Path ((Get-ChildItem $d -Directory)[0].FullName) 'tool\triage'; foreach($f in 'triage.py','triage_gui.py','0_Einrichten.bat','1_Probelauf.bat','1b_Heute_neu_ansehen.bat','2_Scharf_schalten.bat','3_Mails_auswaehlen.bat','3b_Heute_auswaehlen.bat','README.md','.env.example'){ Copy-Item (Join-Path $t $f) -Destination '.' -Force }; Write-Host 'OK: Programmdateien aktualisiert.' } catch { Write-Host ('FEHLER: ' + $_.Exception.Message) }"
echo.
pause
