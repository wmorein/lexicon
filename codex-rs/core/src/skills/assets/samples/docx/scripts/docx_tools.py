#!/usr/bin/env python3
"""
Word document utilities using python-docx.
"""

import argparse
import sys
from pathlib import Path

try:
    from docx import Document
    from docx.shared import Inches, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    print("Error: python-docx not installed. Run: pip install python-docx", file=sys.stderr)
    sys.exit(1)


def extract_text(docx_path: str) -> str:
    """Extract all text from a Word document."""
    doc = Document(docx_path)
    text_parts = []

    for para in doc.paragraphs:
        text_parts.append(para.text)

    for table in doc.tables:
        for row in table.rows:
            row_text = [cell.text for cell in row.cells]
            text_parts.append('\t'.join(row_text))

    return '\n'.join(text_parts)


def get_info(docx_path: str) -> dict:
    """Get document metadata and structure info."""
    doc = Document(docx_path)
    core_props = doc.core_properties

    return {
        'title': core_props.title,
        'author': core_props.author,
        'created': str(core_props.created) if core_props.created else None,
        'modified': str(core_props.modified) if core_props.modified else None,
        'paragraphs': len(doc.paragraphs),
        'tables': len(doc.tables),
        'sections': len(doc.sections),
    }


def from_markdown(md_path: str, docx_path: str):
    """Create a Word document from a markdown file (basic conversion)."""
    doc = Document()

    with open(md_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    for line in lines:
        line = line.rstrip('\n')

        # Headings
        if line.startswith('# '):
            doc.add_heading(line[2:], level=1)
        elif line.startswith('## '):
            doc.add_heading(line[3:], level=2)
        elif line.startswith('### '):
            doc.add_heading(line[4:], level=3)
        # Bullet lists
        elif line.startswith('- ') or line.startswith('* '):
            doc.add_paragraph(line[2:], style='List Bullet')
        # Numbered lists
        elif len(line) > 2 and line[0].isdigit() and line[1] == '.':
            doc.add_paragraph(line[3:], style='List Number')
        # Empty lines create paragraph breaks
        elif not line.strip():
            continue
        # Regular paragraphs
        else:
            doc.add_paragraph(line)

    doc.save(docx_path)
    print(f"Created {docx_path}")


def create_blank(docx_path: str, title: str = None):
    """Create a blank document with optional title."""
    doc = Document()

    if title:
        doc.add_heading(title, 0)

    doc.save(docx_path)
    print(f"Created {docx_path}")


def main():
    parser = argparse.ArgumentParser(description="Word document utilities")
    subparsers = parser.add_subparsers(dest='command', required=True)

    # to-text command
    p_text = subparsers.add_parser('to-text', help='Extract text from document')
    p_text.add_argument('docx', help='Input Word document')
    p_text.add_argument('-o', '--output', help='Output file (default: stdout)')

    # info command
    p_info = subparsers.add_parser('info', help='Show document info')
    p_info.add_argument('docx', help='Input Word document')

    # from-markdown command
    p_md = subparsers.add_parser('from-markdown', help='Create docx from markdown')
    p_md.add_argument('markdown', help='Input markdown file')
    p_md.add_argument('output', help='Output Word document')

    # create command
    p_create = subparsers.add_parser('create', help='Create blank document')
    p_create.add_argument('output', help='Output Word document')
    p_create.add_argument('--title', help='Document title')

    args = parser.parse_args()

    if args.command == 'to-text':
        text = extract_text(args.docx)
        if args.output:
            with open(args.output, 'w', encoding='utf-8') as f:
                f.write(text)
            print(f"Written to {args.output}")
        else:
            print(text)

    elif args.command == 'info':
        info = get_info(args.docx)
        for key, value in info.items():
            print(f"{key}: {value}")

    elif args.command == 'from-markdown':
        from_markdown(args.markdown, args.output)

    elif args.command == 'create':
        create_blank(args.output, args.title)


if __name__ == '__main__':
    main()
