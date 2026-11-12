@echo off
chcp 65001 >nul
title VoterMail Engine — Build Portable EXE Package
echo ========================================================
echo       VoterMail Engine — Portable EXE Builder
echo              Created by Danang Satria
echo ========================================================
echo.

where python >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python tidak terdeteksi di sistem PATH!
    pause
    exit /b 1
)

python -c "import PyInstaller" >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo [INFO] PyInstaller belum terpasang. Menginstal PyInstaller via pip...
    python -m pip install "pyinstaller>=6.0.0"
)

echo Memulai perakitan paket portabel VoterMail_Engine.exe...
python scripts/build_exe.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Proses perakitan gagal dengan kode (%ERRORLEVEL%).
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo ========================================================
echo [SUKSES] Paket portabel dan arsip .zip telah dibuat di folder dist/
echo ========================================================
pause
