@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if errorlevel 1 (
    set "PYTHON=python"
) else (
    set "PYTHON=py -3"
)

%PYTHON% -m pip install -r requirements.txt -r requirements-build.txt
if errorlevel 1 exit /b %errorlevel%

%PYTHON% -m PyInstaller --clean --noconfirm auditor_planilhas.spec
if errorlevel 1 exit /b %errorlevel%

copy /y "executar_auditor.bat" "dist\AuditorPlanilhas\executar_auditor.bat" >nul
copy /y "configuracao.exemplo.bat" "dist\AuditorPlanilhas\configuracao.exemplo.bat" >nul

echo Pacote criado em dist\AuditorPlanilhas.
echo Distribua a pasta inteira, nao apenas o arquivo .exe.
