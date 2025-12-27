---
name: pptx
description: Create, read, and edit PowerPoint presentations (.pptx files). Use when building slide decks, extracting content from presentations, or automating presentation creation.
---

# PowerPoint Presentation Handling

## Overview

Work with PowerPoint files (.pptx) using Python's python-pptx library.

## Dependencies

```bash
pip install python-pptx
```

## Creating Presentations

### Basic Presentation

```python
from pptx import Presentation
from pptx.util import Inches, Pt

prs = Presentation()

# Title slide
title_slide_layout = prs.slide_layouts[0]
slide = prs.slides.add_slide(title_slide_layout)
title = slide.shapes.title
subtitle = slide.placeholders[1]
title.text = "Presentation Title"
subtitle.text = "Subtitle here"

# Content slide
bullet_slide_layout = prs.slide_layouts[1]
slide = prs.slides.add_slide(bullet_slide_layout)
shapes = slide.shapes
title_shape = shapes.title
body_shape = shapes.placeholders[1]

title_shape.text = "Slide Title"
tf = body_shape.text_frame
tf.text = "First bullet point"
p = tf.add_paragraph()
p.text = "Second bullet point"
p.level = 0
p = tf.add_paragraph()
p.text = "Sub-bullet"
p.level = 1

prs.save('presentation.pptx')
```

### Using the Script

```bash
# Create from markdown outline
python scripts/pptx_tools.py from-outline outline.md presentation.pptx

# Extract text from presentation
python scripts/pptx_tools.py extract presentation.pptx

# Get presentation info
python scripts/pptx_tools.py info presentation.pptx
```

## Reading Presentations

```python
from pptx import Presentation

prs = Presentation('existing.pptx')

for slide_num, slide in enumerate(prs.slides, 1):
    print(f"--- Slide {slide_num} ---")
    for shape in slide.shapes:
        if hasattr(shape, "text"):
            print(shape.text)
```

## Slide Layouts

Standard layouts (indices):
- 0: Title Slide
- 1: Title and Content
- 2: Section Header
- 3: Two Content
- 4: Comparison
- 5: Title Only
- 6: Blank

```python
# Blank slide
blank_layout = prs.slide_layouts[6]
slide = prs.slides.add_slide(blank_layout)
```

## Adding Shapes

### Text Box

```python
from pptx.util import Inches, Pt

left = Inches(1)
top = Inches(2)
width = Inches(4)
height = Inches(1)

txBox = slide.shapes.add_textbox(left, top, width, height)
tf = txBox.text_frame
tf.text = "Text in a box"
```

### Images

```python
left = Inches(1)
top = Inches(2)
slide.shapes.add_picture('image.png', left, top, width=Inches(4))
```

### Tables

```python
rows, cols = 3, 4
left = Inches(1)
top = Inches(2)
width = Inches(6)
height = Inches(2)

table = slide.shapes.add_table(rows, cols, left, top, width, height).table
table.cell(0, 0).text = "Header 1"
table.cell(0, 1).text = "Header 2"
table.cell(1, 0).text = "Data 1"
```

## Formatting

### Text Formatting

```python
from pptx.dml.color import RGBColor
from pptx.util import Pt

paragraph = shape.text_frame.paragraphs[0]
run = paragraph.runs[0]
run.font.bold = True
run.font.size = Pt(24)
run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
```

### Shape Fill

```python
from pptx.dml.color import RGBColor

shape.fill.solid()
shape.fill.fore_color.rgb = RGBColor(0x00, 0x80, 0x00)
```

## Modifying Existing Presentations

```python
prs = Presentation('existing.pptx')

# Modify first slide title
slide = prs.slides[0]
title = slide.shapes.title
title.text = "New Title"

# Add a slide
layout = prs.slide_layouts[1]
new_slide = prs.slides.add_slide(layout)

prs.save('modified.pptx')
```

## Notes

```python
# Add speaker notes
slide = prs.slides[0]
notes_slide = slide.notes_slide
notes_slide.notes_text_frame.text = "Speaker notes here"

# Read speaker notes
for slide in prs.slides:
    notes = slide.notes_slide.notes_text_frame.text
    print(notes)
```

## Limitations

- Limited chart editing support
- Cannot edit SmartArt
- Complex animations not supported
- Master slide editing is limited
