#!/usr/bin/env python3
"""
PDF extraction utilities using pdfplumber.
"""

import argparse
import csv
import sys
from pathlib import Path

try:
    import pdfplumber
except ImportError:
    print("Error: pdfplumber not installed. Run: pip install pdfplumber", file=sys.stderr)
    sys.exit(1)


def parse_page_range(page_spec: str, total_pages: int) -> list[int]:
    """Parse page specification like '1-5,7,9-11' into list of page indices (0-based)."""
    if not page_spec:
        return list(range(total_pages))

    pages = set()
    for part in page_spec.split(','):
        part = part.strip()
        if '-' in part:
            start, end = part.split('-', 1)
            start = int(start) - 1  # Convert to 0-based
            end = int(end)  # End is inclusive in spec
            pages.update(range(start, min(end, total_pages)))
        else:
            page = int(part) - 1
            if 0 <= page < total_pages:
                pages.add(page)

    return sorted(pages)


def extract_text(pdf_path: str, page_range: str = None) -> str:
    """Extract text from PDF."""
    text_parts = []

    with pdfplumber.open(pdf_path) as pdf:
        pages_to_process = parse_page_range(page_range, len(pdf.pages))

        for page_idx in pages_to_process:
            page = pdf.pages[page_idx]
            text = page.extract_text()
            if text:
                text_parts.append(f"--- Page {page_idx + 1} ---")
                text_parts.append(text)

    return '\n\n'.join(text_parts)


def get_info(pdf_path: str) -> dict:
    """Get PDF metadata and structure info."""
    with pdfplumber.open(pdf_path) as pdf:
        return {
            'path': pdf_path,
            'pages': len(pdf.pages),
            'metadata': pdf.metadata,
        }


def extract_tables(pdf_path: str, page_range: str = None) -> list[list[list]]:
    """Extract all tables from PDF."""
    all_tables = []

    with pdfplumber.open(pdf_path) as pdf:
        pages_to_process = parse_page_range(page_range, len(pdf.pages))

        for page_idx in pages_to_process:
            page = pdf.pages[page_idx]
            tables = page.extract_tables()
            for table in tables:
                all_tables.append({
                    'page': page_idx + 1,
                    'data': table
                })

    return all_tables


def search_pdf(pdf_path: str, query: str) -> list[dict]:
    """Search for text in PDF, return matching pages."""
    results = []
    query_lower = query.lower()

    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            if query_lower in text.lower():
                # Find context around match
                idx = text.lower().find(query_lower)
                start = max(0, idx - 50)
                end = min(len(text), idx + len(query) + 50)
                context = text[start:end]

                results.append({
                    'page': i + 1,
                    'context': f"...{context}..."
                })

    return results


def main():
    parser = argparse.ArgumentParser(description="PDF extraction utilities")
    subparsers = parser.add_subparsers(dest='command', required=True)

    # extract command
    p_extract = subparsers.add_parser('extract', help='Extract text from PDF')
    p_extract.add_argument('pdf', help='Input PDF file')
    p_extract.add_argument('-o', '--output', help='Output file (default: stdout)')
    p_extract.add_argument('--pages', help='Page range (e.g., "1-5,7,9-11")')

    # info command
    p_info = subparsers.add_parser('info', help='Show PDF info')
    p_info.add_argument('pdf', help='Input PDF file')

    # tables command
    p_tables = subparsers.add_parser('tables', help='Extract tables from PDF')
    p_tables.add_argument('pdf', help='Input PDF file')
    p_tables.add_argument('--output', help='Output CSV file')
    p_tables.add_argument('--pages', help='Page range')

    # search command
    p_search = subparsers.add_parser('search', help='Search text in PDF')
    p_search.add_argument('pdf', help='Input PDF file')
    p_search.add_argument('query', help='Search query')

    args = parser.parse_args()

    if args.command == 'extract':
        text = extract_text(args.pdf, args.pages)
        if args.output:
            with open(args.output, 'w', encoding='utf-8') as f:
                f.write(text)
            print(f"Written to {args.output}")
        else:
            print(text)

    elif args.command == 'info':
        info = get_info(args.pdf)
        print(f"File: {info['path']}")
        print(f"Pages: {info['pages']}")
        print("Metadata:")
        for key, value in (info['metadata'] or {}).items():
            print(f"  {key}: {value}")

    elif args.command == 'tables':
        tables = extract_tables(args.pdf, args.pages)
        if not tables:
            print("No tables found")
            return

        if args.output:
            with open(args.output, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                for table_info in tables:
                    writer.writerow([f"=== Page {table_info['page']} ==="])
                    for row in table_info['data']:
                        writer.writerow(row)
                    writer.writerow([])
            print(f"Written to {args.output}")
        else:
            for table_info in tables:
                print(f"\n=== Table from Page {table_info['page']} ===")
                for row in table_info['data']:
                    print('\t'.join(str(cell) if cell else '' for cell in row))

    elif args.command == 'search':
        results = search_pdf(args.pdf, args.query)
        if not results:
            print(f"No matches found for '{args.query}'")
        else:
            print(f"Found {len(results)} page(s) with matches:")
            for result in results:
                print(f"\nPage {result['page']}: {result['context']}")


if __name__ == '__main__':
    main()
