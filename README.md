# Siwar Image Reader v1.0.1

**Portable OCR tool for Arabic and French images and PDFs with GUI interface.**

---

## Features
- Read images (JPG, PNG) and PDFs
- Extract text to **TXT** and **Word (.docx)**
- Supports **Arabic** and **French**
- Preprocessing and deskewing of images for better OCR results
- Corrects orientation of scanned documents
- Simple and modern GUI with **Tkinter + ttkbootstrap**
- Multi-page PDF support
- Portable EXE (includes Tesseract OCR and Poppler)

---

## Requirements
- **Windows 10/11** ( 64-bit)  
- **No Python installation required** (EXE is portable)  
- Minimum disk space: ~50 MB for EXE + bundled files

> For developers who want to run from source, Python 3.x and the following libraries are required:
```bash
pip install pytesseract Pillow pdf2image python-docx numpy ttkbootstrap


