@echo off
title Stock Tracker
echo.
echo  ================================================
echo   Stock Tracker — Advancer Smart Technology
echo  ================================================
echo.

REM --- Change the PIN here ---
set EDITOR_PIN=ast2026

REM --- Optional: set a separate read-only PIN (leave blank = no PIN needed for viewers) ---
set VIEWER_PIN=

pip install flask --quiet

echo  Starting server...
echo.
python app.py
pause
