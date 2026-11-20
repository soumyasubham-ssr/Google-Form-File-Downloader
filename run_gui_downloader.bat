@echo off
title Google Form File Downloader
cd /d "%~dp0"
python -m pip install --disable-pip-version-check -q -r requirements.txt
if errorlevel 1 (
  echo Failed to install dependencies. Make sure Python 3.9+ is installed and on PATH.
  pause
  exit /b 1
)
start "" pythonw google_form_file_downloader.py
