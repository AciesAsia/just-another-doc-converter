# DocumentConverter

A free, local Windows desktop app for converting documents — no cloud, no subscription, no data leaving your machine.

Built with Python and Tkinter. Requires Microsoft Office on the host machine for full conversion quality.

---

## What It Does

**Convert to PDF**
Converts Word, Excel, PowerPoint, HTML, Markdown, CSV, and plain text files to PDF.

**Convert PDF to Office**
Converts PDF files to Word (.docx) or Excel (.xlsx).

---

## System Requirements

- Windows 10 or later (64-bit)
- Microsoft Office installed (Word, Excel, PowerPoint) — required for full conversion quality
- Tesseract OCR (optional) — for scanned PDFs without a text layer

---

## Download and Install

1. Go to the [Releases](../../releases) page
2. Download `DocumentConverter_Setup_v2.0.exe`
3. Run the installer and follow the wizard:
   - Choose an install location (default: `Program Files\DocumentConverter`)
   - Choose your output folder — converted files will be saved here by default
   - Optionally create a Desktop shortcut
4. Click **Finish** — the app launches automatically

---

## How to Use

### Convert Files to PDF
1. Open **DocumentConverter** from the Start Menu or Desktop shortcut
2. Click **Convert Files to PDF**
3. Click **＋ Add Files** to select files, or **＋ Add Folder** to scan a folder
4. The **Output Folder** shows where converted PDFs will be saved — click **Browse…** to change it
5. Click **▶ Convert All**
6. When done, click **📂 Open Output Folder** to view the results

### Convert PDF to Office
1. Click **Convert PDF to Office** from the launcher
2. Add your PDF files
3. Select output format: **Word (.docx)** or **Excel (.xlsx)**
4. Choose your output folder (or use the default)
5. Click **▶ Convert PDF**

---

## Supported Formats

| Input | Output |
|-------|--------|
| .docx, .doc | .pdf |
| .xlsx, .xls | .pdf |
| .pptx, .ppt | .pdf |
| .html, .htm | .pdf |
| .md | .pdf |
| .txt | .pdf |
| .csv | .pdf |
| .pdf | .docx (Word) |
| .pdf | .xlsx (Excel) |

---

## Notes

- **Microsoft Office must be installed** for Word, Excel, and PowerPoint conversions. The app uses Word/Excel/PowerPoint COM automation for the highest quality output.
- **PDF → Word** uses Word COM automation first, then falls back to pymupdf, then OCR.
- **PDF → Excel** uses pdfplumber table extraction — works best on PDFs with clearly defined table borders.
- The output folder can be changed at any time using the **Browse…** button inside either converter window. Changes apply to the current session only; the default is set at install time.

---

## Contributing & Derivative Works

Community contributions, bug fixes, and feature requests are welcome.

This project enforces strict open-source continuity. If you fork, modify, or build a derivative application based on this code, **you are legally required to release your entire project as open-source software under the same GPL-3.0 terms.**

## License

This project is licensed under the **GNU General Public License v3.0 (GPL-3.0)**.

- **Individuals:** Free to use, modify, study, and share.
- **Commercial entities:** GPL-3.0 forbids integrating this code into proprietary closed-source software. Any application that distributes or uses this code must itself be made fully open-source under GPL-3.0.

See the [LICENSE](LICENSE) file for the full licence text.

---

*DocumentConverter v2.0 — Windows only*
