#!/usr/bin/env bash
set -e

DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$DIR"

echo "================================================="
echo "   啟動 Market Physics Lab (股票也有回復力嗎？)   "
echo "================================================="

if ! command -v python3 &> /dev/null; then
    echo "錯誤: 系統未安裝 python3，請先安裝 Python。"
    exit 1
fi

echo "[1/2] 檢查並安裝相依套件..."
python3 -m pip install -q -r requirements.txt

echo "[2/2] 啟動互動工作台網頁..."
python3 -m streamlit run app.py
