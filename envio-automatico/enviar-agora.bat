@echo off
title Enviar planilha agora - Painel da Tolva
chcp 65001 >nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0TolvaSync.ps1" -Agora
echo.
pause
