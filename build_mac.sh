#!/bin/bash

echo "Building LAB14 Excel to SQLite Importer for Mac..."
echo

echo "Checking Python installation..."
python3 --version
if [ $? -ne 0 ]; then
    echo "ERROR: Python 3 is not installed or not in PATH"
    echo "Please install Python 3.8+ from https://python.org"
    exit 1
fi

echo
echo "Installing/updating required packages..."
pip3 install --upgrade pip
pip3 install --upgrade pyinstaller pandas openpyxl xlsxwriter flask

echo
echo "Building executable..."
pyinstaller importxl_mac.spec

if [ $? -ne 0 ]; then
    echo "ERROR: Build failed!"
    exit 1
fi

echo
echo "Build completed successfully!"
echo
echo "The executable is located in: dist/LAB14_Excel_Importer/"
echo
echo "To run the application:"
echo "1. Navigate to dist/LAB14_Excel_Importer/"
echo "2. Run: ./LAB14_Excel_Importer"
echo "3. Open http://localhost:5666 in your web browser"
echo 