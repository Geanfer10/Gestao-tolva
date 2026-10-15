@echo off
title Remover envio automatico - Painel da Tolva
chcp 65001 >nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0TolvaSync.ps1" -Desinstalar
echo.
pause
