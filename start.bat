@echo off
:: ============================================================
::  start.bat - Запуск setup.ps1 от имени администратора
:: ============================================================

setlocal

:: Получаем путь к текущей директории
set "SCRIPT_DIR=%~dp0"

:: Проверяем, запущены ли мы от имени администратора
net session >nul 2>&1
if %errorLevel% == 0 (
    goto :RunScript
)

:: Если не администратор, запрашиваем права через UAC
echo Запрос прав администратора...
powershell -Command "Start-Process '%~f0' -Verb RunAs"
exit /b

:RunScript
:: Запускаем PowerShell скрипт
echo Запуск setup.ps1...
powershell -NoProfile -ExecutionPolicy Bypass -Command "& { Set-Location '%SCRIPT_DIR%'; .\setup.ps1 }"

endlocal
