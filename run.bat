@echo off
chcp 65001 > nul
cd /d "%~dp0"
title Docx Tool - Auto Filler

echo ======================================================================
echo          CONG CU TU DONG DIEN DU LIEU WORD VA BANG TU EXCEL
echo ======================================================================
echo.

:: 1. Kiem tra may da cai Python chua
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [LOI] May tinh cua ban chua cai dat Python!
    echo.
    echo Huong dan cai dat Python nhanh:
    echo 1. Truy cap: https://www.python.org/downloads/
    echo 2. Tai ban cai dat va chay file .exe
    echo 3. [QUAN TRONG] Nho tich chon vao o "Add Python to PATH" o buoc dau tien!
    echo 4. Sau khi cai xong, mo lai file run.bat nay.
    echo.
    pause
    exit /b
)

:: 2. Kiem tra moi truong ao .venv (Chi chay 1 lan dau tien)
if not exist ".venv\Scripts\python.exe" (
    echo [Thong bao] Dang khoi tao moi truong ao va cai dat thu vien cho lan dau...
    echo - Qua trinh nay chi mat khoang 1 - 2 phut, vui long doi...
    echo.
    python -m venv .venv
    call .venv\Scripts\python.exe -m pip install --quiet --upgrade pip
    call .venv\Scripts\python.exe -m pip install --quiet -r requirements.txt
    echo.
    echo [OK] Cai dat hoan tat!
    echo.
)

:: 3. Khoi dong ung dung tren trinh duyet
echo Dang mo ung dung tren trinh duyet web...
call .venv\Scripts\streamlit.exe run app.py
pause
