#!/usr/bin/env python3
"""
build_excel.py — Convert JSON scrape data to full analytics Excel file
Usage: python build_excel.py <data.json> [output.xlsx]

This script is called after exporting JSON from the web UI,
or can be run standalone with the scraper engine directly.
"""

import sys
import json
import argparse
from pathlib import Path
from datetime import datetime

sys.path.insert(0, str(Path(__file__).parent))
from scraper_engine import scrape_url
from excel_exporter import create_excel_export
from csv_exporter import export_to_csv


def build_from_json(json_path: str, output_path: str = None) -> str:
    """Build Excel from previously scraped JSON data"""
    with open(json_path, 'r', encoding='utf-8') as f:
        results = json.load(f)

    if not output_path:
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_path = str(Path(json_path).parent / f'webscraper_export_{ts}.xlsx')

    print(f"Building Excel from {len(results)} results...")
    create_excel_export(results, output_path)
    print(f"✅ Excel saved: {output_path}")
    return output_path


def scrape_and_export(urls: list[str], output_dir: str = '.', 
                      format: str = 'both') -> dict:
    """Full pipeline: scrape URLs → export to Excel/CSV"""
    print(f"\n{'='*60}")
    print(f"WebHarvest Pro — Scraping {len(urls)} URL(s)")
    print(f"{'='*60}\n")

    results = []
    for i, url in enumerate(urls, 1):
        print(f"[{i}/{len(urls)}] Fetching: {url}")
        result = scrape_url(url)
        results.append(result)

        if result['success']:
            data = result['data']
            print(f"  ✅ Success — Title: {data.get('title','')[:60]}")
            print(f"     Headings: {len(data.get('headings',[]))} | "
                  f"Paragraphs: {len(data.get('paragraphs',[]))} | "
                  f"Tables: {len(data.get('tables',[]))} | "
                  f"Links: {len(data.get('links',[]))}")
        else:
            print(f"  ❌ {result['error']}")

        # Respect crawl delay
        if i < len(urls):
            import time
            delay = result.get('crawl_delay', 2)
            print(f"  ⏱ Rate limiting: {delay}s delay...")
            time.sleep(delay)

    print(f"\n{'='*60}")
    success = sum(1 for r in results if r['success'])
    print(f"Completed: {success}/{len(results)} successful")

    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_files = []

    if format in ('xlsx', 'both'):
        xlsx_path = str(Path(output_dir) / f'webscraper_{ts}.xlsx')
        create_excel_export(results, xlsx_path)
        output_files.append(xlsx_path)
        print(f"✅ Excel: {xlsx_path}")

    if format in ('csv', 'both'):
        csv_files = export_to_csv(results, output_dir)
        output_files.extend(csv_files)
        for f in csv_files:
            print(f"✅ CSV: {f}")

    # Save raw JSON for record-keeping
    json_path = str(Path(output_dir) / f'webscraper_raw_{ts}.json')
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"✅ Raw JSON: {json_path}")

    print(f"\n{'='*60}")
    print(f"Output directory: {output_dir}")
    return {'results': results, 'files': output_files}


def main():
    parser = argparse.ArgumentParser(
        description='WebHarvest Pro — Ethical web scraper with Excel/CSV export',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Scrape URLs directly:
  python build_excel.py --urls https://news.ycombinator.com https://example.com

  # Build Excel from previously exported JSON:
  python build_excel.py --json webscraper_data.json

  # Scrape with CSV output only:
  python build_excel.py --urls https://example.com --format csv --output ./exports
        """
    )
    parser.add_argument('--json', help='Path to previously exported JSON file')
    parser.add_argument('--urls', nargs='+', help='URLs to scrape')
    parser.add_argument('--format', choices=['xlsx', 'csv', 'both'], default='both',
                        help='Output format (default: both)')
    parser.add_argument('--output', default='.', help='Output directory')
    
    # Also support positional JSON arg for backwards compat
    parser.add_argument('json_pos', nargs='?', help='JSON file (positional)')

    args = parser.parse_args()

    json_file = args.json or args.json_pos

    if json_file:
        build_from_json(json_file, 
                       str(Path(args.output) / 'webscraper_from_json.xlsx') if args.output != '.' else None)
    elif args.urls:
        Path(args.output).mkdir(parents=True, exist_ok=True)
        scrape_and_export(args.urls, args.output, args.format)
    else:
        # Interactive mode
        print("WebHarvest Pro — Interactive Mode")
        print("Enter URLs (one per line, blank line to start scraping):\n")
        urls = []
        while True:
            line = input().strip()
            if not line:
                break
            urls.append(line)
        
        if urls:
            scrape_and_export(urls, '.', 'both')
        else:
            parser.print_help()


if __name__ == '__main__':
    main()
