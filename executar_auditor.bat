@echo off
setlocal
cd /d "%~dp0"

if not exist "configuracao.bat" (
    echo Configuracao ausente.
    echo Copie configuracao.exemplo.bat para configuracao.bat e ajuste o site e os escopos.
    pause
    exit /b 2
)

call "configuracao.bat"
if errorlevel 1 exit /b %errorlevel%

if not exist "AuditorPlanilhas.exe" (
    echo Executavel AuditorPlanilhas.exe nao encontrado nesta pasta.
    pause
    exit /b 3
)

start "" "AuditorPlanilhas.exe"
