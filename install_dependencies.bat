@echo off
:: ============================================================
::  DocumentConverter - Dependency Installer
::  Run this once as Administrator before using the converter.
::  Double-click or run from Command Prompt.
:: ============================================================

echo.
echo ============================================================
echo  DocumentConverter - Installing Python dependencies
echo ============================================================
echo.

:: Upgrade pip first
python -m pip install --upgrade pip

echo.
echo [1/11] Installing docx2pdf (Word to PDF) ...
pip install docx2pdf

echo.
echo [2/11] Installing pywin32 (Word / Excel / PowerPoint COM) ...
pip install pywin32

echo.
echo [3/11] Installing reportlab (MD / TXT / CSV / HTML to PDF) ...
pip install reportlab

echo.
echo [4/11] Installing xlrd (legacy .xls reader) ...
pip install xlrd

echo.
echo [5/11] Installing pymupdf (PDF parsing and page rendering) ...
pip install pymupdf

echo.
echo [6/11] Installing python-docx (write .docx output) ...
pip install python-docx

echo.
echo [7/11] Installing pdfplumber (extract tables from PDFs) ...
pip install pdfplumber

echo.
echo [8/11] Installing openpyxl (write .xlsx output) ...
pip install openpyxl

echo.
echo [9/11] Installing pytesseract (OCR for scanned PDFs) ...
pip install pytesseract

echo.
echo [10/11] Installing Pillow (image handling for OCR) ...
pip install Pillow

echo.
echo [11/11] Installing pypdf + numpy + weasyprint (merge PDFs, HTML export) ...
pip install pypdf numpy weasyprint

echo.
echo ============================================================
echo  All Python packages installed.
echo.
echo  IMPORTANT - Tesseract OCR engine (required for scanned PDFs):
echo  If you have not installed it yet, download and install from:
echo  https://github.com/UB-Mannheim/tesseract/wiki
echo  Default install path: C:\Program Files\Tesseract-OCR\tesseract.exe
echo ============================================================
echo.
pause
