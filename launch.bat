@echo off
chcp 65001 >nul
cd /d "%~dp0"

if not exist ".venv\Scripts\activate.bat" (
    echo 仮想環境がありません。フォルダで次を実行してください:
    echo   python -m venv .venv
    echo   .venv\Scripts\activate
    echo   pip install -r requirements.txt
    pause
    exit /b 1
)

call .venv\Scripts\activate.bat
streamlit run app.py
if errorlevel 1 pause
