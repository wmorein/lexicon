---
name: docx
description: Create, read, and edit Word documents (.docx files). Use when working with professional documents, reports, creating formatted text documents, or extracting text from Word files.
---

# Word Document Handling

## Overview

Work with Word documents (.docx) using Python's python-docx library. For reading text, pandoc is often simpler.

## Dependencies

```bash
pip install python-docx
# Optional: brew install pandoc (for text extraction)
```

## Quick Text Extraction

For simply reading text content, pandoc is fastest:

```bash
# To plain text
pandoc document.docx -o output.txt

# To markdown (preserves structure)
pandoc document.docx -o output.md
```

## Creating Documents

### Basic Document

```python
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

doc = Document()

# Add heading
doc.add_heading('Document Title', 0)

# Add paragraph
p = doc.add_paragraph('This is a paragraph.')
p.add_run(' Bold text.').bold = True
p.add_run(' Italic text.').italic = True

# Add list
doc.add_paragraph('Item 1', style='List Bullet')
doc.add_paragraph('Item 2', style='List Bullet')

# Add table
table = doc.add_table(rows=2, cols=3)
table.style = 'Table Grid'
table.cell(0, 0).text = 'Header 1'
table.cell(0, 1).text = 'Header 2'
table.cell(0, 2).text = 'Header 3'

doc.save('output.docx')
```

### Using the Script

```bash
# Create from markdown
python scripts/docx_tools.py from-markdown input.md output.docx

# Extract text
python scripts/docx_tools.py to-text document.docx

# Get document info
python scripts/docx_tools.py info document.docx
```

## Reading Documents

```python
from docx import Document

doc = Document('file.docx')

# Read all paragraphs
for para in doc.paragraphs:
    print(para.text)

# Read tables
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            print(cell.text)
```

## Modifying Documents

```python
from docx import Document

doc = Document('existing.docx')

# Modify paragraph
doc.paragraphs[0].text = "New text"

# Add content
doc.add_paragraph('Additional paragraph')

doc.save('modified.docx')
```

## Formatting

```python
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Font formatting
run = doc.add_paragraph().add_run('Formatted text')
run.font.size = Pt(14)
run.font.bold = True
run.font.color.rgb = RGBColor(0x42, 0x24, 0xE9)

# Paragraph alignment
para = doc.add_paragraph('Centered')
para.alignment = WD_ALIGN_PARAGRAPH.CENTER

# Page breaks
doc.add_page_break()
```

## Working with Sections

```python
from docx.shared import Inches
from docx.enum.section import WD_ORIENT

section = doc.sections[0]
section.orientation = WD_ORIENT.LANDSCAPE
section.page_width = Inches(11)
section.page_height = Inches(8.5)
```

## Adding Images

```python
doc.add_picture('image.png', width=Inches(4))
```

## Limitations

- python-docx cannot read/write tracked changes (use pandoc for reading tracked changes)
- Complex formatting may not be preserved when modifying existing docs
- For redlining/track changes, consider using Word directly or LibreOffice CLI
