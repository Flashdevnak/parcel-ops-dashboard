@echo off
title Parcel Ops Dashboard - Production Preview
cd /d "%~dp0"
echo Starting Parcel Ops Dashboard production preview...
echo.
start "" "http://localhost:5173"
python dev-server.py
pause
