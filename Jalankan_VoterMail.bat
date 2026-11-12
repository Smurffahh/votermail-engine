@echo off
chcp 65001 >nul
title VoterMail Engine — Launcher
echo ========================================================
echo       VoterMail Engine — Otomasi Mail-Merge C6
echo              Created by Danang Satria
echo ========================================================
echo.

where python >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python tidak terdeteksi di sistem PATH!
    echo Silakan unduh dan pasang Python versi 3.10 atau lebih baru dari:
    echo https://www.python.org/downloads/
    echo.
    echo Pastikan untuk mencentang opsi "Add Python to PATH" saat instalasi.
    echo.
    pause
    exit /b 1
)

echo Memulai VoterMail Engine...
python main.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [PERHATIAN] Aplikasi berhenti dengan kode error (%ERRORLEVEL%).
    pause
)
