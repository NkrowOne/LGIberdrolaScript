@echo off
:: ============================================================
:: Lanzador para CertificadosLG.ps1
:: Evita restricciones de ExecutionPolicy en equipos corporativos
:: ============================================================
title Certificados LG - Iniciando...

:: Ejecutar el script .ps1 del mismo directorio con bypass de politica
powershell.exe -ExecutionPolicy Bypass -NoProfile -File "%~dp0CertificadosLG.ps1"

:: Si powershell.exe no esta disponible, intentar con ruta completa
if %ERRORLEVEL% NEQ 0 (
    echo Reintentando con ruta completa de PowerShell...
    %SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -NoProfile -File "%~dp0CertificadosLG.ps1"
)

exit /b
