@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo First-time setup: creating virtual environment and installing requirements...
    py -m venv .venv
    ".venv\Scripts\python.exe" -m pip install -q -r requirements.txt
)
if not exist "samples\sample-article.docx" (
    echo First-time setup: building the demo article...
    ".venv\Scripts\python.exe" make_sample.py
)
".venv\Scripts\python.exe" app.py
pause
