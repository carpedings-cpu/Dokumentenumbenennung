@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo   KPC Morgenbriefing - Aktualisieren
echo ============================================
echo.
echo Laedt die aktuelle Version von GitHub und ersetzt nur die Programmdateien.
echo Deine eigenen Dateien bleiben: .env, briefing_state.json, briefing_log.md
echo.
set "URL=https://github.com/carpedings-cpu/Dokumentenumbenennung/archive/refs/heads/claude/amazing-feynman-eimedk.zip"
powershell -NoProfile -ExecutionPolicy Bypass -Command "try{ $ErrorActionPreference='Stop'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; $z=Join-Path $env:TEMP 'kpc_brief.zip'; $d=Join-Path $env:TEMP 'kpc_brief_x'; Invoke-WebRequest -Uri $env:URL -OutFile $z; if(Test-Path $d){Remove-Item $d -Recurse -Force}; Expand-Archive $z $d -Force; $t=Join-Path ((Get-ChildItem $d -Directory)[0].FullName) 'tool\briefing'; $skip=@('.env','briefing_state.json','briefing_log.md','briefing_config.json','9_Aktualisieren.bat'); Get-ChildItem $t -File | Where-Object { $skip -notcontains $_.Name } | ForEach-Object { Copy-Item $_.FullName -Destination '.' -Force }; Write-Host 'OK: Programmdateien aktualisiert.' } catch { Write-Host ('FEHLER: ' + $_.Exception.Message) }"
echo.
pause
