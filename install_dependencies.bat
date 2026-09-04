@echo off
echo ============================================
echo  Installing Document Converter dependencies
echo ============================================
echo.

call "C:\PythonProjects\DocumentConverter\.venv\Scripts\activate.bat"

echo Installing core dependencies...
pip install docx2pdf pywin32 reportlab xlrd pypdf Pillow numpy weasyprint

echo.
echo Installing Phase 4 PDF-conversion dependencies...
pip install pymupdf python-docx pdfplumber openpyxl pytesseract

echo.
echo ============================================
echo  IMPORTANT: Tesseract OCR (for scanned PDFs)
echo  must be installed separately:
echo.
echo  1. Download the installer from:
echo     https://github.com/UB-Mannheim/tesseract/wiki
echo     (choose: tesseract-ocr-w64-setup-5.x.x.exe)
echo.
echo  2. Run the installer — tick "Add to PATH"
echo     during setup.
echo.
echo  3. Restart this window after installing.
echo.
echo  NOTE: Tesseract is only needed for scanned
echo  (image-only) PDFs. Text-based PDFs work
echo  without it.
echo ============================================
echo.
echo  Done. You can close this window.
echo ============================================
pause
