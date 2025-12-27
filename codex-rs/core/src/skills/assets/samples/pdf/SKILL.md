---
name: pdf
description: Extract text, tables, and metadata from PDF files. Use when reading PDF documents, extracting data from PDFs, converting PDFs to text, or analyzing PDF structure.
---

# PDF Handling

## Overview

Extract content from PDF files using pdfplumber (best for text/tables) or PyMuPDF (best for speed and images).

## Dependencies

```bash
pip install pdfplumber  # Recommended for text/tables
pip install pymupdf     # Alternative: faster, better for images
```

## Quick Text Extraction

### Using pdfplumber

```python
import pdfplumber

with pdfplumber.open('document.pdf') as pdf:
    for page in pdf.pages:
        text = page.extract_text()
        print(text)
```

### Using the Script

```bash
# Extract all text
python scripts/pdf_tools.py extract document.pdf

# Extract to file
python scripts/pdf_tools.py extract document.pdf -o output.txt

# Get PDF info
python scripts/pdf_tools.py info document.pdf

# Extract specific pages
python scripts/pdf_tools.py extract document.pdf --pages 1-5
```

## Extracting Tables

pdfplumber excels at table extraction:

```python
import pdfplumber

with pdfplumber.open('document.pdf') as pdf:
    page = pdf.pages[0]
    tables = page.extract_tables()

    for table in tables:
        for row in table:
            print(row)
```

### Table to CSV

```bash
python scripts/pdf_tools.py tables document.pdf --output tables.csv
```

## PDF Metadata

```python
import pdfplumber

with pdfplumber.open('document.pdf') as pdf:
    print(f"Pages: {len(pdf.pages)}")
    print(f"Metadata: {pdf.metadata}")
```

## Using PyMuPDF (Alternative)

PyMuPDF (fitz) is faster and better for images:

```python
import fitz  # PyMuPDF

doc = fitz.open('document.pdf')

# Extract text
for page in doc:
    text = page.get_text()
    print(text)

# Extract images
for page_num, page in enumerate(doc):
    images = page.get_images()
    for img_index, img in enumerate(images):
        xref = img[0]
        pix = fitz.Pixmap(doc, xref)
        pix.save(f"page{page_num}_img{img_index}.png")

doc.close()
```

## Page-by-Page Processing

```python
import pdfplumber

with pdfplumber.open('large_document.pdf') as pdf:
    for i, page in enumerate(pdf.pages):
        print(f"--- Page {i + 1} ---")
        text = page.extract_text()
        if text:
            print(text[:500])  # First 500 chars
```

## Searching in PDFs

```python
import pdfplumber

def search_pdf(path: str, query: str):
    results = []
    with pdfplumber.open(path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            if query.lower() in text.lower():
                results.append((i + 1, text))
    return results
```

## Limitations

- Scanned PDFs require OCR (use `pytesseract` + `pdf2image`)
- Complex layouts may not extract perfectly
- PDF creation requires different libraries (reportlab, fpdf2)

## For Scanned PDFs (OCR)

```bash
pip install pytesseract pdf2image
# Also need: brew install tesseract poppler
```

```python
from pdf2image import convert_from_path
import pytesseract

images = convert_from_path('scanned.pdf')
for img in images:
    text = pytesseract.image_to_string(img)
    print(text)
```
