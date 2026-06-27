@echo off
chcp 65001 >nul
title Cap nhat Sales Dashboard
echo ============================================
echo   CAP NHAT SALES DASHBOARD TU EXCEL
echo ============================================
echo.
echo Dang doc Excel va tao lai du lieu, vui long doi...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0update.ps1"
echo.
echo ============================================
echo   HOAN TAT. Nhan phim bat ky de dong cua so.
echo ============================================
pause >nul
