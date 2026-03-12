"""CSV Export - clean, analysis-ready CSV files"""
import csv
import os
from datetime import datetime


def export_to_csv(all_results: list, output_dir: str) -> list[str]:
    """Export scrape results to multiple analysis-ready CSV files"""
    os.makedirs(output_dir, exist_ok=True)
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    files = []

    # 1. Summary CSV
    summary_path = os.path.join(output_dir, f'summary_{ts}.csv')
    with open(summary_path, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(['url', 'status', 'title', 'description', 'paragraphs',
                    'headings', 'tables', 'links', 'scraped_at', 'content_hash'])
        for r in all_results:
            data = r.get('data') or {}
            w.writerow([
                r['url'],
                'success' if r.get('success') else r.get('error', 'error'),
                data.get('title', ''),
                data.get('description', ''),
                len(data.get('paragraphs', [])),
                len(data.get('headings', [])),
                len(data.get('tables', [])),
                len(data.get('links', [])),
                r.get('timestamp', ''),
                r.get('meta', {}).get('content_hash', ''),
            ])
    files.append(summary_path)

    # 2. Raw content CSV
    raw_path = os.path.join(output_dir, f'raw_content_{ts}.csv')
    with open(raw_path, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(['source_url', 'section_type', 'index', 'content', 'scraped_at'])
        for r in all_results:
            if not r.get('success'):
                continue
            data = r.get('data', {})
            url, ts2 = r['url'], r.get('timestamp', '')
            for h in data.get('headings', []):
                w.writerow([url, 'heading_' + h['level'], '', h['text'], ts2])
            for i, p in enumerate(data.get('paragraphs', [])):
                w.writerow([url, 'paragraph', i+1, p, ts2])
            for i, items in enumerate(data.get('lists', [])):
                for item in items:
                    w.writerow([url, 'list_item', i+1, item, ts2])
            if data.get('description'):
                w.writerow([url, 'meta_description', '', data['description'], ts2])
    files.append(raw_path)

    # 3. Links CSV
    links_path = os.path.join(output_dir, f'links_{ts}.csv')
    with open(links_path, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(['source_url', 'link_text', 'link_url', 'external', 'scraped_at'])
        for r in all_results:
            if not r.get('success'):
                continue
            for link in r.get('data', {}).get('links', []):
                w.writerow([
                    r['url'],
                    link.get('text', ''),
                    link.get('url', ''),
                    'yes' if link.get('external') else 'no',
                    r.get('timestamp', ''),
                ])
    files.append(links_path)

    return files
