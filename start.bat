@echo off
:: 飞书应用凭证（替换为你的实际值）
set FEISHU_APP_ID=cli_aacf00f76a781bb4
set FEISHU_APP_SECRET=uoKG3FkcHtNT6z0wIWMkNcCUahISSAGI

:: 启动服务
cd /d "d:\A工作\question-factory"
venv\xueqiu\Scripts\python.exe main.py
pause
