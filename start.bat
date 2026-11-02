@echo off
chcp 65001 >nul
echo =================================================
echo    啟動 Market Physics Lab (股票也有回復力嗎？)   
echo =================================================

python --version >nul 2>&1
if errorlevel 1 (
    echo 錯誤: 系統未安裝 Python，請先至 python.org 安裝 Python 並勾選 Add to PATH。
    pause
    exit /b 1
)

echo [1/2] 正在檢查並安裝相依套件...
pip install -q -r requirements.txt

echo [2/2] 正在啟動互動研究工作台...
python -m streamlit run app.py

pause
