"""
WebScraper Engine - Ethical, public-data-only web scraping
Respects robots.txt, rate limits, and copyright guidelines
"""

import re
import time
import json
import hashlib
import urllib.robotparser
from datetime import datetime
from urllib.parse import urlparse, urljoin

import requests
from bs4 import BeautifulSoup


# Sites/patterns that indicate private/paywalled content
BLOCKED_PATTERNS = [
    r'login', r'signin', r'sign-in', r'authenticate', r'auth/',
    r'dashboard', r'account/', r'profile/', r'members/', r'private/',
    r'admin/', r'portal/', r'secure/', r'intranet'
]

BLOCKED_DOMAINS = [
    'mail.google.com', 'gmail.com', 'outlook.com', 'facebook.com',
    'instagram.com', 'linkedin.com/in/', 'twitter.com/messages',
    'patreon.com', 'onlyfans.com'
]

PAYWALL_INDICATORS = [
    'subscribe to read', 'premium content', 'paid subscribers only',
    'sign in to continue', 'create an account', 'members only',
    'subscription required', 'unlock this article'
]

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (compatible; DataResearchBot/1.0; +publicly-available-data-collection)',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.5',
}


def check_robots_txt(url: str) -> tuple[bool, str]:
    """Check if scraping is allowed by robots.txt"""
    parsed = urlparse(url)
    robots_url = f"{parsed.scheme}://{parsed.netloc}/robots.txt"
    rp = urllib.robotparser.RobotFileParser()
    rp.set_url(robots_url)
    try:
        rp.read()
        allowed = rp.can_fetch(HEADERS['User-Agent'], url)
        if not allowed:
            return False, "robots.txt disallows scraping this URL"
        crawl_delay = rp.crawl_delay(HEADERS['User-Agent'])
        return True, str(crawl_delay or 2)
    except Exception:
        return True, "2"  # Default 2s delay if robots.txt unreachable


def is_url_safe(url: str) -> tuple[bool, str]:
    """Validate URL is public, accessible, and ethical to scrape"""
    parsed = urlparse(url)

    # Must have valid scheme
    if parsed.scheme not in ('http', 'https'):
        return False, "Only HTTP/HTTPS URLs are supported"

    # Check blocked domains
    for domain in BLOCKED_DOMAINS:
        if domain in parsed.netloc.lower():
            return False, f"Domain '{parsed.netloc}' is not suitable for public data collection"

    # Check blocked path patterns
    full_url_lower = url.lower()
    for pattern in BLOCKED_PATTERNS:
        if re.search(pattern, full_url_lower):
            return False, f"URL appears to point to private/authenticated content (matched: {pattern})"

    # Must be a real domain
    if not parsed.netloc or '.' not in parsed.netloc:
        return False, "Invalid URL - no valid domain found"

    return True, "OK"


def fetch_page(url: str, timeout: int = 15) -> tuple[BeautifulSoup | None, str, dict]:
    """Fetch a page and return parsed HTML with metadata"""
    meta = {
        'url': url,
        'scraped_at': datetime.now().isoformat(),
        'status_code': None,
        'error': None,
        'is_public': True,
        'content_hash': None,
    }

    try:
        resp = requests.get(url, headers=HEADERS, timeout=timeout, allow_redirects=True)
        meta['status_code'] = resp.status_code
        meta['final_url'] = resp.url

        if resp.status_code != 200:
            meta['error'] = f"HTTP {resp.status_code}"
            return None, meta['error'], meta

        content = resp.text
        meta['content_hash'] = hashlib.sha256(content.encode()).hexdigest()

        # Check for paywall indicators
        content_lower = content.lower()
        for indicator in PAYWALL_INDICATORS:
            if indicator in content_lower:
                meta['is_public'] = False
                meta['error'] = f"Paywall detected: '{indicator}'"
                return None, meta['error'], meta

        soup = BeautifulSoup(content, 'lxml')
        return soup, "OK", meta

    except requests.exceptions.ConnectionError:
        meta['error'] = "Connection failed - site may be down or inaccessible"
        return None, meta['error'], meta
    except requests.exceptions.Timeout:
        meta['error'] = "Request timed out"
        return None, meta['error'], meta
    except Exception as e:
        meta['error'] = str(e)
        return None, meta['error'], meta


def extract_structured_data(soup: BeautifulSoup, url: str) -> dict:
    """Extract all meaningful structured data from a page"""
    data = {
        'title': '',
        'description': '',
        'headings': [],
        'paragraphs': [],
        'tables': [],
        'links': [],
        'images_alt': [],
        'metadata': {},
        'structured_data': [],
        'lists': [],
    }

    # Title
    if soup.title:
        data['title'] = soup.title.get_text(strip=True)

    # Meta tags
    for meta in soup.find_all('meta'):
        name = meta.get('name') or meta.get('property') or ''
        content = meta.get('content') or ''
        if name and content:
            data['metadata'][name] = content
    data['description'] = data['metadata'].get('description', data['metadata'].get('og:description', ''))

    # Headings
    for tag in ['h1', 'h2', 'h3', 'h4']:
        for h in soup.find_all(tag):
            text = h.get_text(strip=True)
            if text:
                data['headings'].append({'level': tag, 'text': text})

    # Paragraphs (skip very short ones)
    for p in soup.find_all('p'):
        text = p.get_text(strip=True)
        if len(text) > 30:
            data['paragraphs'].append(text[:500])  # Cap at 500 chars per para

    # Tables
    for table in soup.find_all('table'):
        rows = []
        headers = []
        thead = table.find('thead')
        if thead:
            headers = [th.get_text(strip=True) for th in thead.find_all(['th', 'td'])]
        for tr in table.find_all('tr'):
            cells = [td.get_text(strip=True) for td in tr.find_all(['td', 'th'])]
            if cells and any(cells):
                rows.append(cells)
        if rows:
            data['tables'].append({'headers': headers, 'rows': rows[:100]})  # Cap rows

    # Links (external only, limit)
    base_domain = urlparse(url).netloc
    seen_links = set()
    for a in soup.find_all('a', href=True):
        href = a['href']
        full_url = urljoin(url, href)
        parsed = urlparse(full_url)
        if parsed.scheme in ('http', 'https') and full_url not in seen_links:
            seen_links.add(full_url)
            data['links'].append({
                'text': a.get_text(strip=True)[:100],
                'url': full_url,
                'external': parsed.netloc != base_domain
            })
            if len(data['links']) >= 50:
                break

    # Lists
    for ul in soup.find_all(['ul', 'ol']):
        items = [li.get_text(strip=True) for li in ul.find_all('li') if li.get_text(strip=True)]
        if len(items) > 1:
            data['lists'].append(items[:20])

    # JSON-LD structured data
    for script in soup.find_all('script', type='application/ld+json'):
        try:
            json_data = json.loads(script.string)
            data['structured_data'].append(json_data)
        except Exception:
            pass

    return data


def scrape_url(url: str) -> dict:
    """Full scrape pipeline for a single URL"""
    result = {
        'url': url,
        'timestamp': datetime.now().isoformat(),
        'success': False,
        'error': None,
        'data': None,
        'meta': {}
    }

    # Safety checks
    safe, reason = is_url_safe(url)
    if not safe:
        result['error'] = f"BLOCKED: {reason}"
        return result

    robots_ok, delay_or_reason = check_robots_txt(url)
    if not robots_ok:
        result['error'] = f"BLOCKED: {delay_or_reason}"
        return result

    crawl_delay = float(delay_or_reason) if delay_or_reason.replace('.', '').isdigit() else 2.0

    # Fetch page
    soup, status, meta = fetch_page(url)
    result['meta'] = meta

    if soup is None:
        result['error'] = status
        return result

    # Extract data
    result['data'] = extract_structured_data(soup, url)
    result['success'] = True
    result['crawl_delay'] = crawl_delay

    return result
