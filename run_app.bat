@echo off
title Insurance Claim MIS Extractor - Localhost
cd /d "%~dp0"
echo ========================================================
echo   INSURANCE CLAIM MIS EXTRACTOR - LOCALHOST SERVER
echo ========================================================
echo Starting server on http://127.0.0.1:5055 ...
start http://127.0.0.1:5055
.\.venv\Scripts\python.exe app.py
pause
