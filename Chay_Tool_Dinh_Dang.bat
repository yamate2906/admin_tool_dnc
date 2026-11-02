@echo off
cd /d "%~dp0"
title DNC AutoReport Formatter
echo Dang khoi dong DNC AutoReport Formatter...

where python >nul 2>nul
if %errorlevel% equ 0 (
    python app_gui.py
    goto end
)

where py >nul 2>nul
if %errorlevel% equ 0 (
    py app_gui.py
    goto end
)

if exist "C:\Users\ASUS\AppData\Local\Programs\Python\Python314\python.exe" (
    "C:\Users\ASUS\AppData\Local\Programs\Python\Python314\python.exe" app_gui.py
    goto end
)

echo Khong tim thay Python tren he thong.
pause

:end
