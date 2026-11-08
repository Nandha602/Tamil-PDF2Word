# Third-party software

PDF2Word is licensed under the GNU Affero General Public License v3.0 (see `LICENSE`).
It uses the following components, each under its own license:

| Component | Use | License |
|---|---|---|
| [PyMuPDF](https://github.com/pymupdf/PyMuPDF) | Reading and rendering PDFs | AGPL-3.0 |
| [python-docx](https://github.com/python-openxml/python-docx) | Writing Word files | MIT |
| [Pillow](https://python-pillow.org/) | Image handling | MIT-CMU (HPND) |
| [PyInstaller](https://pyinstaller.org/) | Packaging the .exe | GPL-2.0 with bootloader exception |
| [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) | OCR engine, downloaded on first run (Windows build by [UB Mannheim](https://github.com/UB-Mannheim/tesseract/wiki)) | Apache-2.0 |
| [tessdata_best](https://github.com/tesseract-ocr/tessdata_best) | Tamil and English language data, downloaded on first run | Apache-2.0 |

Tesseract and the language data are downloaded from their official sources and are not
bundled in this repository or in the release `.exe`.
