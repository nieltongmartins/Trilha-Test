@echo off
rem Copie este arquivo como configuracao.bat e ajuste apenas as localizacoes.
rem Nao inclua senha, token, cookie nem qualquer outra credencial.
set "SHAREPOINT_SITE_URL=https://tenant.sharepoint.com/sites/exemplo"
set "SHAREPOINT_SCOPE_PATHS=/sites/exemplo/Documentos Compartilhados"

rem Caminhos abaixo sao opcionais. O launcher usa a pasta da distribuicao.
rem set "AUDIT_DATABASE_PATH=data\database\auditoria.db"
rem set "AUDIT_LOG_PATH=logs\auditoria.log"
rem set "AUDIT_TEMP_DIRECTORY=data\temp"
rem set "AUDIT_REPORTS_DIRECTORY=data\reports"
rem set "AUDIT_LOG_LEVEL=INFO"
