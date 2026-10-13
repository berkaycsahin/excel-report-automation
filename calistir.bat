@echo off
REM Windows: cift tiklayinca girdi klasorundeki dosyalardan rapor uretir
cd /d "%~dp0"
if not exist .venv (python -m venv .venv && .venv\Scripts\pip install -q -r requirements.txt)
.venv\Scripts\python rapor.py
pause
