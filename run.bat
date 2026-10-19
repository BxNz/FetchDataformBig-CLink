@echo off
cd /d "%~dp0"
set PATH=C:\Users\KCC_IT-Program2\.local\bin;%PATH%
uv run --python 3.13 --with playwright --with openpyxl --with requests --with pillow python bigc_gui.py
pause
