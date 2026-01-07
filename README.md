# Document Converter (Python · GTK 3)

A simple, local **document conversion application** written in **Python** using **GTK 3**.
The application performs common document format conversions entirely offline.

---

## Features

- Images → PDF
- PDF → Images (PNG)
- PDF → PDF (merge and page range extraction)
- PDF → DOCX
- DOCX → PDF (basic text rendering)

All conversions are performed locally.

---

## Project Structure

```
.
├── document_converter.py
├── main.ui
├── screenshots/
└── README.md
```

---

## Requirements

### System dependencies

GTK 3 and Poppler are required.

**Arch / Manjaro**
```
sudo pacman -S gtk3 poppler
```

**Debian / Ubuntu**
```
sudo apt install python3-gi gir1.2-gtk-3.0 poppler-utils
```

---

### Python dependencies

```
pip install pillow pypdf pdf2image pdf2docx python-docx reportlab
```

---

## Running

```
python3 document_converter.py
```

---

## Notes

- DOCX → PDF conversion is text-only.
- PDF → DOCX quality depends on the source PDF.
- No network access is required.
