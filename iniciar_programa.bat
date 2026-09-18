@echo off
title DocxFormatter Pro
chcp 65001 >nul
echo ========================================================
echo   Iniciando DocxFormatter Pro...
echo ========================================================
python main.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Ocorreu um erro ao iniciar o programa.
    pause
)
