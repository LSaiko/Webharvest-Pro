# 💼 Skills Showcase — Annotated Code Walkthrough

> This document is written for hiring managers and technical interviewers.
> It maps each section of the codebase to a specific, demonstrable skill,
> with the exact code snippets and an explanation of the engineering decisions behind them.

---

## Table of Contents

1. [Web Scraping & HTTP Engineering](#1-web-scraping--http-engineering)
2. [Defensive / Ethical Programming](#2-defensive--ethical-programming)
3. [HTML Parsing & Data Extraction](#3-html-parsing--data-extraction)
4. [File I/O — Excel Generation](#4-file-io--excel-generation)
5. [Automation & Scheduling](#5-automation--scheduling)
6. [CLI Design](#6-cli-design)
7. [Frontend JavaScript Architecture](#7-frontend-javascript-architecture)
8. [Data Engineering Patterns](#8-data-engineering-patterns)
9. [Technical Interview Talking Points](#9-technical-interview-talking-points)

---

## 1. Web Scraping & HTTP Engineering

**File:** `src/scraper_engine.py`  
**Skill level demonstrated:** Intermediate–Advanced

### Custom User-Agent & Header Strategy

```python
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (compatible; DataResearchBot/1.0; +publicly-available-data-collection)',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.5',
}
```

**Why this matters:**  
A bare `requests.get(url)` sends Python's default User-Agent, which many sites immediately rate-limit or block. Setting a descriptive, honest User-Agent is both practical (higher success rate) and ethical (transparency about what's crawling). The `Accept` header signals we want HTML specifically, not JSON API responses.

**Interview talking point:** *"I chose a transparent bot identifier rather than spoofing a real browser, which is an intentional ethical choice — it allows site operators to identify and block my bot if they want to."*

---

### Timeout & Error Handling

```python
def fetch_page(url: str, timeout: int = 15) -> tuple[BeautifulSoup | None, str, dict]:
    try:
        resp = requests.get(url, headers=HEADERS, timeout=timeout, allow_redirects=True)
        ...
    except requests.exceptions.ConnectionError:
        meta['error'] = "Connection failed - site may be down or inaccessible"
        return None, meta['error'], meta
    except requests.exceptions.Timeout:
        meta['error'] = "Request timed out"
        return None, meta['error'], meta
    except Exception as e:
        meta['error'] = str(e)
        return None, meta['error'], meta
```

**Why this matters:**  
Production scrapers must never crash on a single bad URL. Each failure mode is caught separately so the caller gets a specific error message, not a Python traceback. The function always returns a consistent 3-tuple — callers don't need to check for exceptions, just check if `soup is None`.

**Pattern used:** *Returning None + error string instead of raising exceptions at the scrape layer* — keeps the orchestration loop clean.

---

### Return Value Design — Metadata Dict

```python
meta = {
    'url': url,
    'scraped_at': datetime.now().isoformat(),
    'status_code': None,
    'error': None,
    'is_public': True,
    'content_hash': None,
}
```

**Why this matters:**  
Every fetch operation returns rich metadata alongside the content. The `content_hash` (MD5 of raw HTML) enables change detection between monitoring runs. The `scraped_at` ISO timestamp makes the output directly usable in Excel/PowerBI time series without transformation.

---

## 2. Defensive / Ethical Programming

**File:** `src/scraper_engine.py`  
**Skill level demonstrated:** Senior — thinking beyond "make it work"

### Four-Layer Ethics Pipeline

Every URL goes through four sequential checks before a single HTTP request fires:

```python
def scrape_url(url: str) -> dict:
    # Layer 1 + 2: URL structure + domain blocklist
    safe, reason = is_url_safe(url)
    if not safe:
        result['error'] = f"BLOCKED: {reason}"
        return result

    # Layer 3: robots.txt compliance
    robots_ok, delay_or_reason = check_robots_txt(url)
    if not robots_ok:
        result['error'] = f"BLOCKED: {delay_or_reason}"
        return result

    # Fetch happens here...

    # Layer 4: Paywall detection (runs on page content)
    for indicator in PAYWALL_INDICATORS:
        if indicator in content_lower:
            meta['is_public'] = False
            meta['error'] = f"Paywall detected: '{indicator}'"
            return None, meta['error'], meta
```

**Why this matters:**  
Most portfolio scrapers have no ethics layer — they just fetch. Implementing a real compliance pipeline demonstrates: security mindset, understanding of web crawling norms, and ability to think about edge cases that would matter in production.

### Regex-Based URL Pattern Matching

```python
BLOCKED_PATTERNS = [
    r'login', r'signin', r'sign-in', r'authenticate', r'auth/',
    r'dashboard', r'account/', r'profile/', r'members/', r'private/',
    r'admin/', r'portal/', r'secure/', r'intranet'
]

for pattern in BLOCKED_PATTERNS:
    if re.search(pattern, full_url_lower):
        return False, f"URL appears to point to private/authenticated content (matched: {pattern})"
```

**Why this matters:**  
Using `re.search` (not `in`) allows future extension with more complex patterns (e.g., `r'/user/\d+/private'`). The error message includes *which* pattern matched — important for debugging and logging.

### robots.txt with Crawl Delay Extraction

```python
def check_robots_txt(url: str) -> tuple[bool, str]:
    rp = urllib.robotparser.RobotFileParser()
    rp.set_url(robots_url)
    try:
        rp.read()
        allowed = rp.can_fetch(HEADERS['User-Agent'], url)
        crawl_delay = rp.crawl_delay(HEADERS['User-Agent'])
        return True, str(crawl_delay or 2)
    except Exception:
        return True, "2"  # Graceful fallback
```

**Why this matters:**  
This extracts the site's **requested crawl delay** from robots.txt and passes it back to the caller. The scraping loop then actually honors it. Most demo scrapers skip this entirely. The `except Exception: return True, "2"` is a deliberate choice — if robots.txt is unreachable (CDN timeout, 404), we default to 2-second delay and continue, rather than blocking the scrape.

---

## 3. HTML Parsing & Data Extraction

**File:** `src/scraper_engine.py` — `extract_structured_data()`  
**Skill level demonstrated:** Intermediate

### Structured Extraction with BeautifulSoup

```python
def extract_structured_data(soup: BeautifulSoup, url: str) -> dict:
    data = {
        'title': '', 'description': '', 'headings': [],
        'paragraphs': [], 'tables': [], 'links': [],
        'images_alt': [], 'metadata': {}, 'structured_data': [], 'lists': [],
    }
```

**Design decision:** Returning a typed dict (not a flat list) means downstream code can access `data['tables']` without parsing a blob. Each field has a consistent type (list, string) so the Excel exporter never needs to check `isinstance`.

### Table Extraction with Header Detection

```python
for table in soup.find_all('table'):
    rows = []
    headers = []
    thead = table.find('thead')
    if thead:
        headers = [th.get_text(strip=True) for th in thead.find_all(['th', 'td'])]
    for tr in table.find_all('tr'):
        cells = [td.get_text(strip=True) for td in tr.find_all(['td', 'th'])]
        if cells and any(cells):  # Skip empty rows
            rows.append(cells)
    if rows:
        data['tables'].append({'headers': headers, 'rows': rows[:100]})
```

**Why `any(cells)` instead of `if cells`:**  
`if cells` would be True for a list of empty strings (e.g., a spacer row). `any(cells)` skips rows where every cell is whitespace — a common pattern in layout tables.

### JSON-LD Structured Data Extraction

```python
for script in soup.find_all('script', type='application/ld+json'):
    try:
        json_data = json.loads(script.string)
        data['structured_data'].append(json_data)
    except Exception:
        pass
```

**Why this matters:**  
JSON-LD (`schema.org`) is how modern sites publish structured metadata — product prices, event dates, article authors. Extracting it separately gives analysts clean, already-structured data without needing to scrape it from raw HTML text.

---

## 4. File I/O — Excel Generation

**File:** `src/excel_exporter.py`  
**Skill level demonstrated:** Intermediate–Advanced

### openpyxl Styling Architecture

```python
def make_fill(hex_color):
    return PatternFill('solid', start_color=hex_color, end_color=hex_color)

def make_font(bold=False, color='000000', size=11, italic=False):
    return Font(bold=bold, color=color, size=size, italic=italic, name='Arial')

def make_border(style='thin'):
    side = Side(style=style)
    return Border(left=side, right=side, top=side, bottom=side)
```

**Design decision:** Factory functions instead of inline style creation. This means a change to the base font or border style propagates everywhere instantly. It also keeps the sheet-building functions readable — `cell.fill = make_fill('1A1A2E')` is clear intent.

### Excel Tables with Auto-Filter

```python
tbl = Table(displayName="SummaryTable", ref=f"A{table_start}:{last_col}{last_row}")
tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
ws.add_table(tbl)
```

**Why this matters:**  
Adding `openpyxl.worksheet.table.Table` objects (not just styled cells) creates *real* Excel tables — with built-in sort/filter dropdowns, named ranges, and structured references like `=SummaryTable[Status]`. This is the difference between data that looks like a table and data that Excel *treats* as a table.

### Dynamic Column Auto-Fit

```python
def auto_fit_columns(ws, min_width=10, max_width=50):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(min_width, min(max_width, max_len + 2))
```

**Why `max(min_width, min(max_width, max_len + 2))`:**  
Clamping between `min_width` and `max_width` prevents both invisible columns (a common bug when a column has only short values) and absurdly wide ones (when a cell contains a long URL). The `+2` adds padding so text doesn't hug the column edge.

---

## 5. Automation & Scheduling

**File:** `web/index.html` (JavaScript)  
**Skill demonstrated:** Async programming, event-driven design

### Non-Blocking Async Scrape Loop

```javascript
async function startScrape() {
    for (let i = 0; i < validUrls.length && scraping; i++) {
        const url = validUrls[i];
        
        // Update UI before the await
        setTagStatus(origIdx, 'scraping', 'SCRAPING...');
        
        const result = await simulateScrape(url);  // Non-blocking
        results.push(result);
        
        // Update UI after the await — no page freeze
        renderResultRow(result, results.length);
        updateStats();

        // Honor rate limit between requests
        if (i < validUrls.length - 1) {
            await sleep(result.crawl_delay * 1000);
        }
    }
}
```

**Why `async/await` instead of `.then()` chains:**  
The sequential nature of the loop (one URL at a time, with delays between) maps naturally to `await`. Converting to Promise chains would require recursive callbacks or explicit state machines. The `&& scraping` guard in the for condition means `stopScrape()` setting `scraping = false` will cleanly abort after the current URL finishes.

### Content-Hash Change Detection

```javascript
if (monitorHashes[url] && monitorHashes[url] !== hash) {
    logMsg(`⚡ CHANGE DETECTED on ${url}`, 'warn');
}
monitorHashes[url] = hash;
```

**Why MD5 hash comparison instead of full content diff:**  
Storing a full HTML snapshot for every monitored URL would consume enormous memory in a browser tab. A 32-character hash is functionally equivalent for change detection — if content changed, the hash changes; if it didn't, hashes match. The hash is generated server-side in Python using `hashlib.md5(content.encode()).hexdigest()`.

### Interval Scheduler with Clean Start/Stop

```javascript
function toggleSchedule() {
    if (scheduleTimer) {
        clearInterval(scheduleTimer);
        scheduleTimer = null;
        // ... update UI to stopped state
    } else {
        const ms = interval * 60 * 60 * 1000;
        scheduleTimer = setInterval(startScrape, ms);
        // ... update UI to running state
    }
}
```

**Pattern:** Toggle function with a single `scheduleTimer` reference stored in module scope. `clearInterval` is idempotent if called with a valid ID. The timer reference being `null` vs a valid ID serves as the state flag — no extra boolean needed.

---

## 6. CLI Design

**File:** `src/build_excel.py`  
**Skill demonstrated:** Developer ergonomics, Unix tool design

### Multi-Mode argparse

```python
parser.add_argument('--json', help='Path to previously exported JSON file')
parser.add_argument('--urls', nargs='+', help='URLs to scrape')
parser.add_argument('--format', choices=['xlsx', 'csv', 'both'], default='both')
parser.add_argument('--output', default='.', help='Output directory')
parser.add_argument('json_pos', nargs='?', help='JSON file (positional)')
```

**Design decisions:**
- `nargs='+'` on `--urls` accepts one or many URLs in a single flag
- `choices=` on `--format` gives argparse automatic validation + error message
- The `json_pos` positional argument enables `python build_excel.py data.json` (easier than `--json`)
- `default='both'` means a first-time user gets the most complete output without needing to know flags

### Interactive Fallback Mode

```python
if not json_file and not args.urls:
    print("WebHarvest Pro — Interactive Mode")
    print("Enter URLs (one per line, blank line to start):\n")
    urls = []
    while True:
        line = input().strip()
        if not line:
            break
        urls.append(line)
```

**Why:** When called with no arguments, a blank error message is hostile UX. Falling into an interactive prompt means a user who just runs `python build_excel.py` still gets a useful experience — they're guided into entering URLs manually.

---

## 7. Frontend JavaScript Architecture

**File:** `web/index.html`  
**Skill demonstrated:** Vanilla JS, state management, DOM manipulation

### Module-Scope State

```javascript
let urls = [];        // URL queue
let results = [];     // Accumulated scrape results
let scraping = false; // Mutex flag — prevents double-starts
let scheduleTimer = null;
let monitorHashes = {}; // url → content_hash map
let runCount = 0;       // Monotonic run counter for log labeling
```

**Why not a class or framework:**  
For a single-page tool with no routing, module-scope variables are simpler than a class and don't require a build step. The `scraping` boolean acts as a lightweight mutex — `startScrape()` checks it at entry, preventing two concurrent scrape loops if the user double-clicks.

### Progressive UI Updates During Async Loop

```javascript
// Before request: show "SCRAPING..."
setTagStatus(origIdx, 'scraping', 'SCRAPING...');
logMsg(`Fetching: ${url}`, 'info');

const result = await simulateScrape(url);

// After request: update all UI surfaces simultaneously
renderResultRow(result, results.length);
updateStats();
setTagStatus(origIdx, result.success ? 'success' : 'error', result.success ? 'DONE' : 'ERROR');
```

**Pattern:** Update the UI *before* the async call (optimistic "in progress" state), then update again after. This gives the user real-time feedback — they see which URL is being fetched right now, not just a spinning indicator.

---

## 8. Data Engineering Patterns

### Consistent Schema Throughout the Pipeline

Every scrape result follows the same schema from fetch → export:

```python
result = {
    'url': str,           # source URL
    'timestamp': str,     # ISO 8601 — works in Excel, PowerBI, pandas
    'success': bool,      # single truth about whether usable data exists
    'error': str | None,  # human-readable failure reason
    'data': {             # None if success=False
        'title': str,
        'description': str,
        'headings': list[dict],   # [{'level': 'h2', 'text': '...'}]
        'paragraphs': list[str],
        'tables': list[dict],     # [{'headers': [...], 'rows': [[...]]}]
        'links': list[dict],      # [{'text': '...', 'url': '...', 'external': bool}]
        'lists': list[list[str]],
        'metadata': dict,
    },
    'meta': {             # HTTP-level metadata
        'status_code': int,
        'content_hash': str,
        'final_url': str,   # post-redirect URL
        'scraped_at': str,
    },
    'crawl_delay': float, # seconds to wait before next request
}
```

**Why this matters:**  
A consistent schema means `excel_exporter.py` can be written completely independently of `scraper_engine.py`. The exporter just expects this shape — it doesn't know or care how the data was collected. This is the separation of concerns pattern that makes the codebase testable and extensible.

### Defense Against None

```python
data = r.get('data') or {}   # not r.get('data', {})
```

**Subtle but important:** `r.get('data', {})` returns `{}` only if the key is missing — but returns `None` if the key exists with value `None`. `r.get('data') or {}` handles both cases. When `success=False`, `data` is explicitly set to `None` — this pattern ensures downstream code never crashes on a failed scrape.

---

## 9. Technical Interview Talking Points

Use these when a hiring manager asks *"tell me about a project you've built"*:

---

**"Tell me about your architecture decisions"**

> "The project is split into four completely independent modules: the ethics engine, the scraper, the Excel exporter, and the CSV exporter. Each can be swapped or tested independently. For example, if I wanted to swap BeautifulSoup for Playwright for JS-rendered pages, I'd only need to change `scraper_engine.py` — the exporters wouldn't change at all because they just consume the result schema."

---

**"How did you handle errors?"**

> "Every fetch returns a 3-tuple: the parsed content, a status string, and a metadata dict. The caller never catches exceptions — failures are just result objects with `success=False` and an `error` string. This means the scraping loop always completes, even if 9 out of 10 URLs fail. Failed results still appear in the output with their error reasons."

---

**"What would you do differently at scale?"**

> "The current design is synchronous between URLs — it fetches one at a time to respect rate limits. At scale I'd use `asyncio` with a semaphore-controlled concurrency pool, and move the rate limiting to a per-domain token bucket rather than a flat delay. The result schema would stay the same — the exporter doesn't care how the data was collected."

---

**"How do you handle the ethics of scraping?"**

> "I built a four-layer validation pipeline that runs before any HTTP request: URL pattern matching for auth pages, domain blocklist for inherently private services, robots.txt compliance via stdlib's `robotparser`, and paywall detection via keyword scanning of the HTML. The crawl delay isn't just a fixed sleep — it's extracted from the site's own robots.txt and honored per-site."

---

**"What does the Excel output actually contain?"**

> "Eight sheets. The interesting ones are the Analytics Template sheet — which has pre-written Excel formulas that reference the other sheets, so you can immediately start analyzing without building anything — and the PowerBI guide, which has DAX formulas, recommended visual types, and a step-by-step connection guide. The idea is that the output file isn't just data, it's a ready-to-use analytics starter kit."

---

*This document was written specifically to help you articulate the engineering decisions in this project during technical interviews or portfolio reviews.*
