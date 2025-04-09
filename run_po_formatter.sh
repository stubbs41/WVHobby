#!/bin/bash

echo "Starting PO File Formatter..."
echo

python3 po_formatter.py

if [ $? -ne 0 ]; then
  echo "Error running PO Formatter. Please make sure Python and required libraries are installed."
  echo "Run: pip install -r requirements.txt"
  read -p "Press Enter to continue..."
fi 