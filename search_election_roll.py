#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "pdf2image",
#     "pytesseract",
#     "pillow",
# ]
# ///
"""
Election Roll PDF Search Tool (FAST VERSION with Pre-extraction)

Usage:
    # Step 1: Extract text from all PDFs (one-time, takes ~30-60 mins)
    uv run search_election_roll.py extract

    # Step 2: Search instantly (takes seconds!)
    uv run search_election_roll.py search
    uv run search_election_roll.py search --name "ஜெகதீசன்"

Prerequisites:
    sudo apt-get install tesseract-ocr tesseract-ocr-tam poppler-utils
"""

import argparse
import json
import os
import sys
from datetime import datetime
from pathlib import Path
from concurrent.futures import ProcessPoolExecutor, as_completed
import multiprocessing

# ============== CONFIGURATION ==============
# SEARCH_NAME = "வினாயக்"
SEARCH_NAME = "விநாயக்"

PDF_FOLDER = "118-eroll"
EXTRACTED_DATA_FILE = "extracted_data.json"
RESULTS_FILE = "results.txt"
DPI = 150
WORKERS = multiprocessing.cpu_count()
SKIP_PAGES = [1, 2]  # Header pages
# ===========================================


def normalize_tamil(text: str) -> str:
    """Remove Tamil vowel signs for fuzzy matching."""
    vowel_signs = [
        '\u0bbe', '\u0bbf', '\u0bc0', '\u0bc1', '\u0bc2',
        '\u0bc6', '\u0bc7', '\u0bc8', '\u0bca', '\u0bcb',
        '\u0bcc', '\u0bcd',
    ]
    for sign in vowel_signs:
        text = text.replace(sign, '')
    return text


# ==================== EXTRACTION ====================

def extract_single_pdf(args) -> dict:
    """Extract text from a single PDF."""
    from pdf2image import convert_from_path
    import pytesseract

    pdf_path, pdf_index, total = args
    pdf_name = Path(pdf_path).name
    result = {'pdf': pdf_name, 'pages': []}

    try:
        pages = convert_from_path(pdf_path, dpi=DPI, thread_count=2)
    except Exception as e:
        print(f"[{pdf_index}/{total}] ERROR {pdf_name}: {e}")
        return result

    for page_num, page_img in enumerate(pages, 1):
        if page_num in SKIP_PAGES:
            continue

        try:
            text = pytesseract.image_to_string(
                page_img,
                lang='tam',
                config='--oem 3 --psm 6'
            )
            if text.strip():
                result['pages'].append({
                    'page': page_num,
                    'text': text
                })
        except Exception:
            continue

    print(f"[{pdf_index}/{total}] ✓ {pdf_name} ({len(result['pages'])} pages)")
    return result


def extract_all_pdfs(folder_path: str, output_file: str, workers: int):
    """Extract text from all PDFs and save to JSON."""
    pdf_files = sorted(Path(folder_path).glob("*.pdf"))
    total = len(pdf_files)

    print(f"\n{'='*60}")
    print(f"EXTRACTING TEXT FROM ALL PDFs")
    print(f"{'='*60}")
    print(f"Folder: {folder_path}")
    print(f"Total PDFs: {total}")
    print(f"Workers: {workers}")
    print(f"Output: {output_file}")
    print(f"{'='*60}\n")

    task_args = [(str(pdf), i+1, total) for i, pdf in enumerate(pdf_files)]
    all_data = []

    with ProcessPoolExecutor(max_workers=workers) as executor:
        futures = {executor.submit(extract_single_pdf, args): args for args in task_args}

        for future in as_completed(futures):
            try:
                result = future.result()
                if result['pages']:
                    all_data.append(result)
            except Exception as e:
                print(f"Error: {e}")

    # Save to JSON
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump({
            'extracted_at': datetime.now().isoformat(),
            'total_pdfs': total,
            'data': all_data
        }, f, ensure_ascii=False, indent=2)

    print(f"\n{'='*60}")
    print(f"EXTRACTION COMPLETE!")
    print(f"{'='*60}")
    print(f"PDFs processed: {total}")
    print(f"PDFs with data: {len(all_data)}")
    print(f"Saved to: {output_file}")
    print(f"\nNow run: uv run search_election_roll.py search")


# ==================== SEARCH ====================

def search_extracted_data(data_file: str, search_name: str, results_file: str):
    """Search pre-extracted data (instant!)."""

    if not os.path.exists(data_file):
        print(f"Error: {data_file} not found!")
        print(f"First run: uv run search_election_roll.py extract")
        sys.exit(1)

    print(f"\nLoading extracted data from {data_file}...")
    with open(data_file, 'r', encoding='utf-8') as f:
        extracted = json.load(f)

    print(f"Searching for: {search_name}")
    print(f"{'='*60}")

    normalized_search = normalize_tamil(search_name)
    results = []

    for pdf_data in extracted['data']:
        pdf_name = pdf_data['pdf']
        for page_data in pdf_data['pages']:
            page_num = page_data['page']
            text = page_data['text']

            # Fuzzy search
            if normalized_search in normalize_tamil(text):
                lines = text.split('\n')
                for line_num, line in enumerate(lines):
                    if normalized_search in normalize_tamil(line):
                        start = max(0, line_num - 3)
                        end = min(len(lines), line_num + 6)
                        context = '\n'.join(lines[start:end])

                        results.append({
                            'pdf': pdf_name,
                            'page': page_num,
                            'line': line.strip(),
                            'context': context
                        })

    # Print results
    print_results(results, search_name)

    # Always save to results.txt
    save_results(results, search_name, results_file)

    return results


def extract_voter_details(context: str) -> dict:
    """Extract voter details from context."""
    import re
    details = {'age': None, 'father_name': None, 'gender': None}

    # Age
    age_match = re.search(r'வயது\s*[:\-]?\s*(\d{2,3})', context)
    if age_match:
        details['age'] = age_match.group(1)

    # Gender
    if 'ஆண்' in context:
        details['gender'] = 'ஆண் (Male)'
    elif 'பெண்' in context:
        details['gender'] = 'பெண் (Female)'

    # Father's name
    father_match = re.search(r'தந்தை(?:யின்)?\s*(?:பெயர்)?\s*[:\-]?\s*([^\n\-]+)', context)
    if father_match:
        details['father_name'] = father_match.group(1).strip()[:40]

    # Husband's name
    husband_match = re.search(r'கணவர்\s*(?:பெயர்)?\s*[:\-]?\s*([^\n\-]+)', context)
    if husband_match:
        details['father_name'] = f"(கணவர்) {husband_match.group(1).strip()[:40]}"

    return details


def print_results(results: list, search_name: str):
    """Print formatted results."""
    print(f"\n{'='*60}")
    print(f"RESULTS FOR: {search_name}")
    print(f"{'='*60}")

    if not results:
        print("\nNo matches found.")
        return

    print(f"\nTotal matches: {len(results)}")

    for i, result in enumerate(results, 1):
        print(f"\n{'─'*40}")
        print(f"Match {i}")
        print(f"{'─'*40}")
        print(f"PDF: {result['pdf']}")
        print(f"Page: {result['page']}")

        details = extract_voter_details(result['context'])
        if details['age']:
            print(f"Age: {details['age']}")
        if details['gender']:
            print(f"Gender: {details['gender']}")
        if details['father_name']:
            print(f"Father/Husband: {details['father_name']}")

        print(f"\nContext:")
        for line in result['context'].split('\n'):
            print(f"    {line}")


def save_results(results: list, search_name: str, output_file: str):
    """Save/update results to file."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(f"{'='*60}\n")
        f.write(f"ELECTION ROLL SEARCH RESULTS\n")
        f.write(f"{'='*60}\n\n")
        f.write(f"Search Name: {search_name}\n")
        f.write(f"Search Time: {timestamp}\n")
        f.write(f"Total Matches: {len(results)}\n")
        f.write(f"\n{'='*60}\n\n")

        if not results:
            f.write("No matches found.\n")
        else:
            for i, result in enumerate(results, 1):
                f.write(f"{'─'*40}\n")
                f.write(f"Match {i}\n")
                f.write(f"{'─'*40}\n")
                f.write(f"PDF: {result['pdf']}\n")
                f.write(f"Page: {result['page']}\n")

                details = extract_voter_details(result['context'])
                if details['age']:
                    f.write(f"Age: {details['age']}\n")
                if details['gender']:
                    f.write(f"Gender: {details['gender']}\n")
                if details['father_name']:
                    f.write(f"Father/Husband: {details['father_name']}\n")

                f.write(f"\nContext:\n")
                for line in result['context'].split('\n'):
                    f.write(f"    {line}\n")
                f.write(f"\n")

    print(f"\n✓ Results saved to: {output_file}")


# ==================== MAIN ====================

def main():
    parser = argparse.ArgumentParser(
        description="Fast Tamil Election Roll PDF Search",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Commands:
  extract    Extract text from all PDFs (one-time, ~30-60 mins)
  search     Search extracted data (instant, seconds!)

Examples:
  uv run search_election_roll.py extract
  uv run search_election_roll.py search
  uv run search_election_roll.py search --name "ஜெகதீசன்"
        """
    )

    subparsers = parser.add_subparsers(dest='command', help='Command to run')

    # Extract command
    extract_parser = subparsers.add_parser('extract', help='Extract text from all PDFs')
    extract_parser.add_argument('--folder', '-f', default=PDF_FOLDER, help=f'PDF folder (default: {PDF_FOLDER})')
    extract_parser.add_argument('--output', '-o', default=EXTRACTED_DATA_FILE, help=f'Output file (default: {EXTRACTED_DATA_FILE})')
    extract_parser.add_argument('--workers', '-w', type=int, default=WORKERS, help=f'Parallel workers (default: {WORKERS})')

    # Search command
    search_parser = subparsers.add_parser('search', help='Search extracted data')
    search_parser.add_argument('--name', '-n', default=SEARCH_NAME, help=f'Name to search (default: {SEARCH_NAME})')
    search_parser.add_argument('--data', '-d', default=EXTRACTED_DATA_FILE, help=f'Extracted data file (default: {EXTRACTED_DATA_FILE})')
    search_parser.add_argument('--results', '-r', default=RESULTS_FILE, help=f'Results file (default: {RESULTS_FILE})')

    args = parser.parse_args()

    if args.command == 'extract':
        if not os.path.exists(args.folder):
            print(f"Error: Folder not found: {args.folder}")
            sys.exit(1)
        extract_all_pdfs(args.folder, args.output, args.workers)

    elif args.command == 'search':
        search_extracted_data(args.data, args.name, args.results)

    else:
        parser.print_help()
        print("\n⚠️  Please specify a command: extract or search")


if __name__ == "__main__":
    main()
