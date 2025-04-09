@echo off
echo Starting PO File Formatter...
echo.
python po_formatter.py
echo.
if %ERRORLEVEL% NEQ 0 (
  echo Error running PO Formatter. Please make sure Python and required libraries are installed.
  echo Run: pip install -r requirements.txt
  pause
) 