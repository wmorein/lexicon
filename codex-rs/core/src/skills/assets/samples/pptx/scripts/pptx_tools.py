#!/usr/bin/env python3
"""
PowerPoint utilities using python-pptx.
"""

import argparse
import sys
from pathlib import Path

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
except ImportError:
    print("Error: python-pptx not installed. Run: pip install python-pptx", file=sys.stderr)
    sys.exit(1)


def extract_text(pptx_path: str) -> str:
    """Extract all text from a PowerPoint presentation."""
    prs = Presentation(pptx_path)
    text_parts = []

    for slide_num, slide in enumerate(prs.slides, 1):
        slide_text = [f"--- Slide {slide_num} ---"]

        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                slide_text.append(shape.text)

        # Include speaker notes
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text
            if notes.strip():
                slide_text.append(f"[Notes: {notes}]")

        text_parts.append('\n'.join(slide_text))

    return '\n\n'.join(text_parts)


def get_info(pptx_path: str) -> dict:
    """Get presentation metadata and structure info."""
    prs = Presentation(pptx_path)
    core_props = prs.core_properties

    slide_info = []
    for i, slide in enumerate(prs.slides, 1):
        title = ""
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text.strip():
                title = shape.text.split('\n')[0][:50]
                break
        slide_info.append({'number': i, 'title': title})

    return {
        'title': core_props.title,
        'author': core_props.author,
        'created': str(core_props.created) if core_props.created else None,
        'modified': str(core_props.modified) if core_props.modified else None,
        'slides': len(prs.slides),
        'slide_details': slide_info,
    }


def from_outline(outline_path: str, pptx_path: str):
    """
    Create a presentation from a markdown outline.

    Format:
    # Presentation Title

    ## Slide Title
    - Bullet point
    - Another bullet
      - Sub-bullet

    ## Another Slide
    Content here
    """
    prs = Presentation()

    with open(outline_path, 'r', encoding='utf-8') as f:
        content = f.read()

    lines = content.strip().split('\n')
    current_slide = None
    current_bullets = []
    presentation_title = None

    def flush_slide():
        nonlocal current_slide, current_bullets
        if current_slide is None:
            return

        if presentation_title and current_slide == presentation_title:
            # Title slide
            layout = prs.slide_layouts[0]
            slide = prs.slides.add_slide(layout)
            slide.shapes.title.text = current_slide
        else:
            # Content slide
            layout = prs.slide_layouts[1]
            slide = prs.slides.add_slide(layout)
            slide.shapes.title.text = current_slide

            if current_bullets:
                body = slide.placeholders[1]
                tf = body.text_frame
                tf.clear()

                for i, (level, text) in enumerate(current_bullets):
                    if i == 0:
                        tf.text = text
                        tf.paragraphs[0].level = level
                    else:
                        p = tf.add_paragraph()
                        p.text = text
                        p.level = level

        current_slide = None
        current_bullets = []

    for line in lines:
        stripped = line.strip()

        if stripped.startswith('# '):
            # Presentation title
            flush_slide()
            presentation_title = stripped[2:].strip()
            current_slide = presentation_title

        elif stripped.startswith('## '):
            # New slide
            flush_slide()
            current_slide = stripped[3:].strip()

        elif stripped.startswith('- ') or stripped.startswith('* '):
            # Bullet point - check indentation for level
            indent = len(line) - len(line.lstrip())
            level = min(indent // 2, 4)  # Max 5 levels (0-4)
            text = stripped[2:].strip()
            current_bullets.append((level, text))

        elif stripped and current_slide:
            # Regular text becomes a bullet
            current_bullets.append((0, stripped))

    flush_slide()
    prs.save(pptx_path)
    print(f"Created {pptx_path} with {len(prs.slides)} slides")


def create_blank(pptx_path: str, title: str = None):
    """Create a blank presentation with optional title slide."""
    prs = Presentation()

    if title:
        layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(layout)
        slide.shapes.title.text = title

    prs.save(pptx_path)
    print(f"Created {pptx_path}")


def main():
    parser = argparse.ArgumentParser(description="PowerPoint utilities")
    subparsers = parser.add_subparsers(dest='command', required=True)

    # extract command
    p_extract = subparsers.add_parser('extract', help='Extract text from presentation')
    p_extract.add_argument('pptx', help='Input PowerPoint file')
    p_extract.add_argument('-o', '--output', help='Output file (default: stdout)')

    # info command
    p_info = subparsers.add_parser('info', help='Show presentation info')
    p_info.add_argument('pptx', help='Input PowerPoint file')

    # from-outline command
    p_outline = subparsers.add_parser('from-outline', help='Create from markdown outline')
    p_outline.add_argument('outline', help='Input markdown outline file')
    p_outline.add_argument('output', help='Output PowerPoint file')

    # create command
    p_create = subparsers.add_parser('create', help='Create blank presentation')
    p_create.add_argument('output', help='Output PowerPoint file')
    p_create.add_argument('--title', help='Title slide text')

    args = parser.parse_args()

    if args.command == 'extract':
        text = extract_text(args.pptx)
        if args.output:
            with open(args.output, 'w', encoding='utf-8') as f:
                f.write(text)
            print(f"Written to {args.output}")
        else:
            print(text)

    elif args.command == 'info':
        info = get_info(args.pptx)
        print(f"Title: {info['title']}")
        print(f"Author: {info['author']}")
        print(f"Created: {info['created']}")
        print(f"Modified: {info['modified']}")
        print(f"Slides: {info['slides']}")
        print("\nSlide overview:")
        for slide in info['slide_details']:
            print(f"  {slide['number']}: {slide['title']}")

    elif args.command == 'from-outline':
        from_outline(args.outline, args.output)

    elif args.command == 'create':
        create_blank(args.output, args.title)


if __name__ == '__main__':
    main()
