import os, io, re, json, tempfile, urllib.parse
from flask import Flask, request, jsonify, render_template, send_file, session
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XlImage
import requests
from bs4 import BeautifulSoup
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'dlx-offer-tool-dev-key-change-in-prod')
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max

UPLOAD_FOLDER = tempfile.mkdtemp()
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv', 'tsv', 'pdf'}

TEMPLATE_COLUMNS = ['Image', 'Item Name', 'Expiration', 'UPC/Item #', 'Quantity', 'Casepack', 'Cost', 'Retail Link', 'FOB']


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


# ── File Parsing ──────────────────────────────────────────────────────────────
def parse_upload(filepath, filename):
    ext = filename.rsplit('.', 1)[1].lower()
    if ext == 'csv':
        df = pd.read_csv(filepath, dtype=str)
    elif ext == 'tsv':
        df = pd.read_csv(filepath, sep='\t', dtype=str)
    elif ext in ('xlsx', 'xls'):
        df = parse_excel(filepath)
    elif ext == 'pdf':
        df = parse_pdf(filepath)
    else:
        raise ValueError(f'Unsupported file type: {ext}')
    df = df.fillna('')
    df.columns = [str(c).strip() for c in df.columns]
    return df


def parse_excel(filepath):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    data = []
    headers = []
    hyperlinks = {}

    for cell in ws[1]:
        headers.append(str(cell.value or f'Column_{cell.column}').strip())

    # Extract embedded images and save them to temp files, mapped by (row, col)
    image_map = {}
    if hasattr(ws, '_images'):
        for img_idx, img in enumerate(ws._images):
            try:
                anchor = img.anchor
                if hasattr(anchor, '_from'):
                    row = anchor._from.row + 1  # 0-indexed to 1-indexed
                    col = anchor._from.col
                    # Extract image data and save to file
                    img_data = None
                    if hasattr(img, '_data'):
                        img_data = img._data()
                    elif hasattr(img, 'ref'):
                        # Try reading from the image ref
                        try:
                            img_data = img.ref.read()
                            img.ref.seek(0)
                        except Exception:
                            pass
                    if not img_data:
                        # Try the blob directly
                        try:
                            from openpyxl.drawing.image import _import_image
                        except ImportError:
                            pass
                        try:
                            buf = io.BytesIO()
                            img_obj = img._data if callable(getattr(img, '_data', None)) else None
                            if img_obj:
                                img_data = img_obj()
                        except Exception:
                            pass

                    if img_data and len(img_data) > 100:
                        # Determine format from first bytes
                        fmt = 'png'
                        if img_data[:3] == b'\xff\xd8\xff':
                            fmt = 'jpeg'
                        elif img_data[:4] == b'\x89PNG':
                            fmt = 'png'
                        elif img_data[:4] == b'GIF8':
                            fmt = 'gif'

                        img_filename = f'extracted_{img_idx}_{row}_{col}.{fmt}'
                        img_path = os.path.join(UPLOAD_FOLDER, img_filename)
                        with open(img_path, 'wb') as f:
                            f.write(img_data)
                        image_map[(row, col)] = img_filename
                        print(f'Extracted image for row={row}, col={col}: {img_filename} ({len(img_data)} bytes)')
                    else:
                        image_map[(row, col)] = True  # Image exists but couldn't extract data
                        print(f'Image detected at row={row}, col={col} but could not extract data')
            except Exception as e:
                print(f'Image extraction error: {e}')

    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=False), start=2):
        row_data = {}
        for col_idx, cell in enumerate(row):
            if col_idx < len(headers):
                val = cell.value if cell.value is not None else ''
                row_data[headers[col_idx]] = str(val).strip()

                # Check for hyperlinks
                if cell.hyperlink and cell.hyperlink.target:
                    hyperlinks[(row_idx, col_idx)] = cell.hyperlink.target
                    row_data[headers[col_idx] + '__hyperlink'] = cell.hyperlink.target

                # Check if this cell has an extracted embedded image
                img_file = image_map.get((row_idx, col_idx))
                if img_file and isinstance(img_file, str):
                    row_data[headers[col_idx] + '__embedded_image'] = img_file

        data.append(row_data)

    df = pd.DataFrame(data)
    return df


def parse_pdf(filepath):
    try:
        import pdfplumber
        tables = []
        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                page_tables = page.extract_tables()
                for table in page_tables:
                    if table and len(table) > 1:
                        headers = [str(h or '').strip() for h in table[0]]
                        for row in table[1:]:
                            row_dict = {}
                            for i, val in enumerate(row):
                                if i < len(headers):
                                    row_dict[headers[i]] = str(val or '').strip()
                            tables.append(row_dict)
        if tables:
            return pd.DataFrame(tables)
    except ImportError:
        pass
    raise ValueError('Could not parse PDF. Install pdfplumber: pip install pdfplumber')


# ── Product Lookup Services ──────────────────────────────────────────────────
# Priority order for retail sources (1st tier)
RETAIL_PRIORITY = ['amazon', 'walmart', 'target', 'costco', 'kroger', 'walgreens', 'cvs', 'ebay', 'bestbuy', 'homedepot']

# Browser-like headers for web scraping
BROWSER_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.5',
}

# Simple in-memory cache to avoid burning rate limits on repeat lookups
_lookup_cache = {}
# Track which APIs are rate-limited so we stop hitting them
_api_rate_limited = {}


def _resolve_redirect_url(url):
    """Follow redirect URLs (like UPCitemdb's /noredir/) to get the actual destination.
    Returns the final URL after redirects, or the original URL if resolution fails."""
    if not url:
        return url
    # UPCitemdb redirect pattern: contains upcitemdb.com and a redirect param
    # Also handle any generic redirect/tracking URLs
    is_redirect = (
        'upcitemdb.com' in url or
        '/redirect?' in url or
        '/noredir/' in url or
        'go.redirectingat.com' in url
    )
    if not is_redirect:
        return url
    try:
        # Use HEAD request with redirects to find final URL
        resp = requests.head(url, headers=BROWSER_HEADERS, timeout=8, allow_redirects=True)
        final_url = resp.url
        if final_url and final_url != url:
            print(f'Resolved redirect: {url[:80]} → {final_url[:80]}')
            return final_url
    except Exception as e:
        print(f'Redirect resolution failed for {url[:80]}: {e}')
    # Fallback: try to extract destination from URL params
    try:
        parsed = urllib.parse.urlparse(url)
        params = urllib.parse.parse_qs(parsed.query)
        for key in ('to', 'url', 'dest', 'redirect', 'target', 'u'):
            if key in params:
                dest = params[key][0]
                if dest.startswith('http'):
                    return dest
    except Exception:
        pass
    return url


def _is_product_page_url(url):
    """Check if a URL is an actual product page (not a search/listing page).
    Returns (clean_url, source) or (None, None)."""
    if not url:
        return None, None

    # Amazon: must have /dp/ASIN — reject /s? (search), /b/ (browse), /gp/ (generic)
    amazon_match = re.search(r'amazon\.com.*/dp/([A-Z0-9]{10})', url)
    if amazon_match:
        asin = amazon_match.group(1)
        return f'https://www.amazon.com/dp/{asin}', 'Amazon'

    # Walmart: must have /ip/ (item page)
    walmart_match = re.search(r'(https?://[^/]*walmart\.com/ip/[^\s?#]+)', url)
    if walmart_match:
        return walmart_match.group(1), 'Walmart'

    # Target: must have /-/A- (product ID)
    target_match = re.search(r'(https?://[^/]*target\.com/[^\s?#]*/-/A-\d+)', url)
    if target_match:
        return target_match.group(1), 'Target'

    # Home Depot: must have /p/ (product page)
    hd_match = re.search(r'(https?://[^/]*homedepot\.com/p/[^\s?#]+)', url)
    if hd_match:
        return hd_match.group(1), 'Home Depot'

    # Costco: must have /product (product page)
    costco_match = re.search(r'(https?://[^/]*costco\.com[^\s?#]*/product[^\s?#]*)', url)
    if costco_match:
        return costco_match.group(1), 'Costco'

    # Best Buy: must have /site/ with .p?skuId
    bb_match = re.search(r'(https?://[^/]*bestbuy\.com/site/[^\s?#]+\.p)', url)
    if bb_match:
        return bb_match.group(1), 'Best Buy'

    # Kroger: product page
    kroger_match = re.search(r'(https?://[^/]*kroger\.com/p/[^\s?#]+)', url)
    if kroger_match:
        return kroger_match.group(1), 'Kroger'

    # Walgreens: product page
    wg_match = re.search(r'(https?://[^/]*walgreens\.com/store/id=[^\s?#]+)', url)
    if wg_match:
        return wg_match.group(1), 'Walgreens'

    # CVS: product page
    cvs_match = re.search(r'(https?://[^/]*cvs\.com/shop/[^\s?#]+)', url)
    if cvs_match:
        return cvs_match.group(1), 'CVS'

    # eBay: must have /itm/ (item page)
    ebay_match = re.search(r'(https?://[^/]*ebay\.com/itm/[^\s?#]+)', url)
    if ebay_match:
        return ebay_match.group(1), 'eBay'

    return None, None


# SearXNG public instances — open-source meta search engines that provide
# JSON APIs and are designed for programmatic use (unlike Google/Bing/DDG
# which block datacenter IPs like Render's).
SEARXNG_INSTANCES = [
    'https://search.bus-hit.me',
    'https://searx.be',
    'https://search.sapti.me',
    'https://searx.tiekoetter.com',
    'https://search.ononoki.org',
    'https://searx.oxf.app',
    'https://paulgo.io',
]


def _search_searxng(query):
    """Search using SearXNG public instances (JSON API).
    Tries multiple instances for reliability.
    Returns (url, title, source) or (None, None, None)."""
    import random
    # Shuffle instances so we spread load
    instances = SEARXNG_INSTANCES.copy()
    random.shuffle(instances)

    for instance in instances:
        try:
            resp = requests.get(
                f'{instance}/search',
                params={
                    'q': query,
                    'format': 'json',
                    'categories': 'general',
                    'language': 'en',
                },
                headers=BROWSER_HEADERS,
                timeout=12
            )
            if resp.status_code == 200:
                data = resp.json()
                results = data.get('results', [])
                for r in results:
                    url = r.get('url', '')
                    title = r.get('title', '')
                    if not url:
                        continue
                    clean_url, source = _is_product_page_url(url)
                    if clean_url and source:
                        # Clean up title
                        title = re.sub(r'^(Amazon\.com|Walmart\.com|Target|Home Depot|Costco)\s*[:\-–]\s*', '', title).strip()
                        title = re.sub(r'\s*[:\-–|]\s*(Amazon\.com|Walmart|Target|Home Depot|Costco|Best Buy|eBay)\s*$', '', title).strip()
                        print(f'SearXNG ({instance}) found: {clean_url} — "{title}"')
                        return clean_url, title, source
                # Had results but none were product pages
                print(f'SearXNG ({instance}) returned {len(results)} results but no product pages for: {query}')
            else:
                print(f'SearXNG ({instance}) returned status {resp.status_code}')
        except Exception as e:
            print(f'SearXNG ({instance}) error: {e}')
            continue

    return None, None, None


def _fetch_product_title_from_page(url):
    """Fetch a product page and extract the title. Returns title string or None."""
    try:
        resp = requests.get(url, headers=BROWSER_HEADERS, timeout=10, allow_redirects=True)
        if resp.status_code == 200:
            soup = BeautifulSoup(resp.text, 'html.parser')

            # Amazon product title
            if 'amazon.com' in url:
                el = soup.find('span', id='productTitle')
                if el:
                    return el.get_text(strip=True)

            # Walmart product title
            if 'walmart.com' in url:
                el = soup.find('h1', itemprop='name') or soup.find('h1')
                if el:
                    return el.get_text(strip=True)

            # Target product title
            if 'target.com' in url:
                el = soup.find('h1') or soup.find('title')
                if el:
                    t = el.get_text(strip=True)
                    return re.sub(r'\s*[:\-–|]\s*Target$', '', t)

            # Home Depot product title
            if 'homedepot.com' in url:
                el = soup.find('h1', class_='product-details__title') or soup.find('h1')
                if el:
                    return el.get_text(strip=True)

            # Generic: use <title> tag, clean up common suffixes
            title_tag = soup.find('title')
            if title_tag:
                t = title_tag.get_text(strip=True)
                t = re.sub(r'\s*[:\-–|]\s*(Amazon\.com|Walmart|Target|Home Depot|Costco|Best Buy|eBay).*$', '', t)
                return t.strip()
    except Exception as e:
        print(f'Page fetch error for {url}: {e}')
    return None


def _name_similarity(name1, name2):
    """Smart word-overlap similarity between two product names. Returns 0.0 to 1.0.
    Handles cases where original is short ('Tide Pods') and lookup is long
    ('Tide PODS Laundry Detergent Soap Pacs, 81 Count, Spring Meadow').
    The key question: do most words from the SHORT name appear in the LONG name?"""
    if not name1 or not name2:
        return 0.0
    # Normalize: lowercase, remove punctuation, split into words
    words1 = set(re.sub(r'[^a-z0-9\s]', '', name1.lower()).split())
    words2 = set(re.sub(r'[^a-z0-9\s]', '', name2.lower()).split())
    # Remove very common filler words
    stop = {'the', 'a', 'an', 'and', 'or', 'of', 'for', 'in', 'with', 'to', 'by',
            'is', 'at', 'on', 'from', 'pack', 'ct', 'oz', 'lb', 'count', 'ea',
            'each', 'per', 'size', 'new', 'free', 'buy', 'best', 'top', 'item'}
    words1 = words1 - stop
    words2 = words2 - stop
    if not words1 or not words2:
        return 0.0
    overlap = words1 & words2

    # Use the SHORTER name as the reference — "do most words from the original
    # name appear somewhere in the lookup result?"
    shorter = words1 if len(words1) <= len(words2) else words2
    if not shorter:
        return 0.0
    return len(overlap) / len(shorter)


def search_product_on_web(upc=None, name=None):
    """Search for a product on major retail sites using SearXNG.
    ONLY returns results that are actual product pages.
    Returns dict with keys: url, title, source."""
    result = {'url': None, 'title': None, 'source': None}

    upc_clean = re.sub(r'[^0-9]', '', str(upc)) if upc else ''
    has_upc = bool(upc_clean) and upc_clean not in ('0', '')

    # Build search queries — try UPC first, then name-based
    queries = []
    if has_upc:
        queries.append(f'{upc_clean} buy')
        queries.append(f'{upc_clean} amazon walmart target')
    if name:
        clean_name = str(name).strip()
        if clean_name:
            queries.append(f'{clean_name} buy amazon')
            queries.append(f'{clean_name} walmart OR target OR "home depot"')

    for query in queries:
        try:
            url, snippet_title, source = _search_searxng(query)
            if url:
                # ALWAYS fetch the real product title from the actual page.
                # SearXNG snippet titles are often wrong or describe a different
                # product than the URL points to.
                page_title = _fetch_product_title_from_page(url)
                title = page_title if page_title else snippet_title
                print(f'Web search: snippet="{snippet_title}" → page="{page_title}" for {url}')

                # VALIDATION: verify relevance if we have an original name
                if name and title:
                    sim = _name_similarity(name, title)
                    if sim < 0.2:
                        print(f'SearXNG result rejected (sim={sim:.2f}): "{title}" vs "{name}"')
                        continue  # Try next query

                result = {'url': url, 'title': title, 'source': source}
                return result
        except Exception as e:
            print(f'SearXNG search failed for "{query}": {e}')
            continue

    return result


def lookup_product_info(upc=None, name=None):
    """Unified product lookup: returns dict with title, image, retail_link, source, upc_matched.
    upc_matched=True means the result came from a direct UPC barcode lookup (trusted).
    upc_matched=False means it came from a web search (needs validation)."""
    result = {'title': None, 'image': None, 'retail_link': None, 'source': None, 'upc_matched': False}

    upc_clean = re.sub(r'[^0-9]', '', str(upc)) if upc else ''
    has_upc = bool(upc_clean) and upc_clean.lower() not in ('na', 'nan', 'none', '0')

    # Check cache first
    cache_key = f'{upc_clean}|{name or ""}'
    if cache_key in _lookup_cache:
        print(f'Cache hit for: {cache_key}')
        return _lookup_cache[cache_key].copy()

    # ── 1. UPCitemdb (best source — returns title, images, and offers in one call) ──
    if has_upc and not _api_rate_limited.get('upcitemdb'):
        try:
            resp = requests.get(
                f'https://api.upcitemdb.com/prod/trial/lookup?upc={upc_clean}',
                timeout=8,
                headers={'Accept': 'application/json'}
            )
            if resp.status_code == 429:
                print('UPCitemdb rate limited — skipping for remaining items')
                _api_rate_limited['upcitemdb'] = True
            elif resp.status_code == 200:
                data = resp.json()
                items = data.get('items', [])
                if items:
                    item = items[0]
                    # UPC API matched — this is a direct barcode lookup, trusted
                    result['upc_matched'] = True
                    # Get product title
                    title = item.get('title', '').strip()
                    if title:
                        result['title'] = title

                    # Get product image
                    images = item.get('images', [])
                    if images:
                        result['image'] = images[0]

                    # Get retail link — prioritize Amazon, then other 1st-tier retailers
                    # UPCitemdb offers are from a direct UPC match, so they're trustworthy.
                    # Just clean up URLs where possible and skip obvious search pages.
                    offers = item.get('offers', [])
                    best_link = None
                    best_priority = len(RETAIL_PRIORITY) + 1  # worst priority
                    fallback_link = None

                    for offer in offers:
                        link = offer.get('link', '')
                        merchant = offer.get('merchant', '').lower()
                        if not link:
                            continue

                        # Resolve redirect URLs (UPCitemdb uses redirect/tracking URLs)
                        link = _resolve_redirect_url(link)

                        # Try to clean up to a canonical product URL
                        product_url, url_source = _is_product_page_url(link)
                        if product_url:
                            link = product_url

                        # Check against priority list
                        for idx, retailer in enumerate(RETAIL_PRIORITY):
                            if retailer in merchant or (url_source and retailer in url_source.lower()):
                                if idx < best_priority:
                                    best_priority = idx
                                    best_link = link
                                    result['source'] = url_source or retailer.capitalize()
                                break
                        else:
                            # Not in priority list — keep as fallback only if it's a real retailer URL
                            if not fallback_link and not 'upcitemdb.com' in link:
                                fallback_link = link

                    result['retail_link'] = best_link or fallback_link

                    # If we got everything, return early
                    if result['title'] and result['image'] and result['retail_link']:
                        _lookup_cache[cache_key] = result.copy()
                        return result
        except Exception as e:
            print(f'UPCitemdb lookup error: {e}')

    # ── 2. Open Food Facts (good for food/grocery — has title + image) ──
    if has_upc and (not result['title'] or not result['image']):
        try:
            resp = requests.get(
                f'https://world.openfoodfacts.org/api/v0/product/{upc_clean}.json',
                timeout=8
            )
            if resp.status_code == 200:
                data = resp.json()
                product = data.get('product', {})
                if product:
                    result['upc_matched'] = True  # Direct UPC lookup
                    if not result['title']:
                        off_name = product.get('product_name', '').strip()
                        if off_name:
                            result['title'] = off_name
                            if not result['source']:
                                result['source'] = 'Open Food Facts'
                    if not result['image']:
                        img = (product.get('image_url') or
                               product.get('image_front_url') or
                               product.get('image_front_small_url'))
                        if img:
                            result['image'] = img
        except Exception:
            pass

    # ── 3. Go-UPC API ──
    if has_upc and (not result['title'] or not result['image']):
        try:
            resp = requests.get(
                f'https://go-upc.com/api/v1/code/{upc_clean}',
                timeout=8,
                headers={'Accept': 'application/json'}
            )
            if resp.status_code == 200:
                data = resp.json()
                prod = data.get('product', {})
                if prod:
                    if not result['title']:
                        go_name = prod.get('name', '').strip()
                        if go_name:
                            result['title'] = go_name
                            if not result['source']:
                                result['source'] = 'Go-UPC'
                    if not result['image']:
                        img = prod.get('imageUrl')
                        if img:
                            result['image'] = img
        except Exception:
            pass

    # ── 4. Barcode Lookup ──
    if has_upc and (not result['title'] or not result['image']):
        try:
            resp = requests.get(
                f'https://www.barcodelookup.com/restapi?barcode={upc_clean}',
                timeout=8,
                headers={'Accept': 'application/json'}
            )
            if resp.status_code == 200:
                data = resp.json()
                products = data.get('products', [])
                if products:
                    if not result['title']:
                        bl_name = products[0].get('title', '').strip() or products[0].get('product_name', '').strip()
                        if bl_name:
                            result['title'] = bl_name
                            if not result['source']:
                                result['source'] = 'Barcode Lookup'
                    if not result['image']:
                        images = products[0].get('images', [])
                        if images:
                            result['image'] = images[0]
        except Exception:
            pass

    # ── 5. Open Food Facts search by name (fallback when no UPC) ──
    if not result['image'] and name:
        try:
            q = urllib.parse.quote_plus(str(name))
            resp = requests.get(
                f'https://world.openfoodfacts.org/cgi/search.pl?search_terms={q}&search_simple=1&action=process&json=1&page_size=1',
                timeout=8
            )
            if resp.status_code == 200:
                data = resp.json()
                products = data.get('products', [])
                if products:
                    if not result['title']:
                        off_name = products[0].get('product_name', '').strip()
                        if off_name:
                            result['title'] = off_name
                    img = (products[0].get('image_url') or
                           products[0].get('image_front_url') or
                           products[0].get('image_front_small_url'))
                    if img:
                        result['image'] = img
        except Exception:
            pass

    # ── 6. Web search fallback: find actual product page (not just search URL) ──
    needs_link = not result['retail_link']
    needs_title = not result['title']
    if needs_link or needs_title:
        web_result = search_product_on_web(upc=upc, name=name)
        if web_result['url'] and needs_link:
            result['retail_link'] = web_result['url']
        if web_result['title'] and needs_title:
            result['title'] = web_result['title']
        if web_result['source'] and not result['source']:
            result['source'] = web_result['source']

    # ── 7. No fallback to search URLs ──
    # We intentionally do NOT generate amazon.com/s?k= search URLs as fallback.
    # It's better to leave the link blank than to give a search page that may
    # show unrelated products. The enrichment UI will show a dash for missing links.

    # Cache the result (even if empty, to avoid re-hitting rate-limited APIs)
    _lookup_cache[cache_key] = result.copy()

    return result


# ── Image Download Helper ────────────────────────────────────────────────────
def download_image(url, max_size_kb=500):
    """Download an image from URL and return as bytes. Returns (bytes, format) or (None, None)."""
    if not url or not url.startswith('http'):
        return None, None
    # Skip google search URLs
    if 'google.com/search' in url:
        return None, None
    try:
        resp = requests.get(url, timeout=10, stream=True, headers={
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })
        if resp.status_code != 200:
            return None, None
        content_type = resp.headers.get('Content-Type', '').lower()
        if 'image' not in content_type and 'octet-stream' not in content_type:
            return None, None

        img_data = resp.content
        if len(img_data) > max_size_kb * 1024 * 5:  # max ~2.5MB
            return None, None

        # Determine format
        if 'png' in content_type or url.lower().endswith('.png'):
            fmt = 'png'
        elif 'gif' in content_type or url.lower().endswith('.gif'):
            fmt = 'gif'
        elif 'webp' in content_type or url.lower().endswith('.webp'):
            fmt = 'png'  # will need conversion
        else:
            fmt = 'jpeg'

        return img_data, fmt
    except Exception:
        return None, None


def resize_image_bytes(img_bytes, max_width=120, max_height=120):
    """Resize image bytes to fit within max dimensions. Returns (bytes, format)."""
    try:
        from PIL import Image as PILImage
        img = PILImage.open(io.BytesIO(img_bytes))
        # Convert to RGB if needed
        if img.mode in ('RGBA', 'P', 'LA'):
            background = PILImage.new('RGB', img.size, (255, 255, 255))
            if img.mode == 'P':
                img = img.convert('RGBA')
            background.paste(img, mask=img.split()[-1] if 'A' in img.mode else None)
            img = background
        elif img.mode != 'RGB':
            img = img.convert('RGB')

        # Resize
        img.thumbnail((max_width, max_height), PILImage.LANCZOS)

        buf = io.BytesIO()
        img.save(buf, format='PNG', optimize=True)
        buf.seek(0)
        return buf.getvalue(), 'png'
    except ImportError:
        # Pillow not installed, return original
        return img_bytes, 'jpeg'
    except Exception:
        return img_bytes, 'jpeg'


# ── Column Auto-Mapping ──────────────────────────────────────────────────────
COLUMN_ALIASES = {
    'Image': ['image', 'img', 'photo', 'picture', 'product image', 'thumbnail', 'image url', 'image link', 'product photo'],
    'Item Name': ['item name', 'product name', 'name', 'description', 'item description', 'product',
                  'item', 'title', 'product description', 'product title'],
    'Expiration': ['expiration', 'exp', 'exp date', 'expiry', 'expiry date', 'best by', 'use by', 'bb date', 'sell by'],
    'UPC/Item #': ['upc', 'upc code', 'item #', 'item number', 'sku', 'barcode', 'gtin', 'ean',
                   'item#', 'upc/item #', 'product code', 'asin'],
    'Quantity': ['quantity', 'qty', 'units', 'count', 'amount', 'pcs', 'pieces', 'total qty', 'available'],
    'Casepack': ['casepack', 'case pack', 'case qty', 'pack size', 'inner pack', 'units per case', 'per case', 'case'],
    'Cost': ['cost', 'price', 'unit cost', 'unit price', 'wholesale', 'wholesale price', 'our price',
             'your cost', 'net price', 'each'],
    'Retail Link': ['retail link', 'link', 'url', 'product link', 'product url', 'retail url', 'buy link',
                    'store link', 'website'],
    'FOB': ['fob', 'f.o.b.', 'fob location', 'ship from', 'origin', 'warehouse', 'location', 'ship point', 'freight'],
}


def auto_map_columns(source_columns):
    mapping = {}
    source_lower = {c: c.lower().strip() for c in source_columns}
    for template_col, aliases in COLUMN_ALIASES.items():
        for src_col, src_lower in source_lower.items():
            if src_lower in aliases or src_lower == template_col.lower():
                mapping[template_col] = src_col
                break
    return mapping


# ── Routes ────────────────────────────────────────────────────────────────────
@app.route('/')
def index():
    return render_template('index.html')


@app.route('/serve_image/<filename>')
def serve_image(filename):
    """Serve extracted images from uploaded files."""
    safe = secure_filename(filename)
    path = os.path.join(UPLOAD_FOLDER, safe)
    if os.path.exists(path):
        return send_file(path, mimetype='image/png')
    return '', 404


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['file']
    if file.filename == '' or not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file. Upload .xlsx, .csv, .tsv, or .pdf'}), 400

    filename = secure_filename(file.filename)
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    try:
        df = parse_upload(filepath, filename)
        session_id = os.urandom(8).hex()
        cache_path = os.path.join(UPLOAD_FOLDER, f'{session_id}.json')
        df.to_json(cache_path, orient='records')

        source_columns = list(df.columns)
        visible_columns = [c for c in source_columns if not c.endswith('__hyperlink') and not c.endswith('__embedded_image')]
        suggested_mapping = auto_map_columns(visible_columns)
        preview = df.head(5).to_dict(orient='records')

        return jsonify({
            'session_id': session_id,
            'source_columns': visible_columns,
            'suggested_mapping': suggested_mapping,
            'preview': preview,
            'row_count': len(df)
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/process', methods=['POST'])
def process_data():
    data = request.json
    session_id = data.get('session_id')
    mapping = data.get('mapping', {})

    cache_path = os.path.join(UPLOAD_FOLDER, f'{session_id}.json')
    if not os.path.exists(cache_path):
        return jsonify({'error': 'Session expired. Please re-upload.'}), 400

    df = pd.read_json(cache_path, dtype=str).fillna('')
    results = []

    for idx, row in df.iterrows():
        item = {}
        for template_col in TEMPLATE_COLUMNS:
            src_col = mapping.get(template_col, '')
            if src_col and src_col in df.columns:
                item[template_col] = str(row.get(src_col, '')).strip()
            else:
                item[template_col] = ''

        # Handle expiration
        if not item.get('Expiration') or item['Expiration'].lower() in ('', 'nan', 'none', 'null'):
            item['Expiration'] = 'NA'

        # Check for hyperlinks in source (Retail Link)
        src_link_col = mapping.get('Retail Link', '')
        if src_link_col:
            hyperlink_key = src_link_col + '__hyperlink'
            if hyperlink_key in df.columns and row.get(hyperlink_key):
                item['Retail Link'] = row[hyperlink_key]

        # Check for image hyperlinks
        src_img_col = mapping.get('Image', '')
        if src_img_col:
            hyperlink_key = src_img_col + '__hyperlink'
            if hyperlink_key in df.columns and row.get(hyperlink_key):
                item['Image'] = row[hyperlink_key]

            # Check for embedded images extracted from the Excel file
            embedded_key = src_img_col + '__embedded_image'
            if embedded_key in df.columns and row.get(embedded_key):
                # Serve the extracted image via our Flask route
                item['Image'] = '/serve_image/' + row[embedded_key]

        # Also check ALL columns for embedded images if Image is still empty
        if not item.get('Image') or not item['Image'].startswith(('http', '/')):
            for col_name in df.columns:
                if col_name.endswith('__embedded_image') and row.get(col_name):
                    item['Image'] = '/serve_image/' + row[col_name]
                    break

        results.append(item)

    # Save processed results
    result_path = os.path.join(UPLOAD_FOLDER, f'{session_id}_processed.json')
    with open(result_path, 'w') as f:
        json.dump(results, f)

    return jsonify({'results': results, 'session_id': session_id})


@app.route('/debug_lookup')
def debug_lookup():
    """Temporary debug endpoint to diagnose API lookup failures."""
    import traceback
    upc = request.args.get('upc', '805106960230')
    name = request.args.get('name', 'Mini Light')
    results = {'upc': upc, 'name': name, 'tests': {}}

    upc_clean = re.sub(r'[^0-9]', '', str(upc))

    # Test 1: UPCitemdb
    try:
        resp = requests.get(
            f'https://api.upcitemdb.com/prod/trial/lookup?upc={upc_clean}',
            timeout=10, headers={'Accept': 'application/json'}
        )
        results['tests']['upcitemdb'] = {
            'status_code': resp.status_code,
            'body_preview': resp.text[:500],
            'items_count': len(resp.json().get('items', [])) if resp.status_code == 200 else None
        }
    except Exception as e:
        results['tests']['upcitemdb'] = {'error': str(e), 'traceback': traceback.format_exc()}

    # Test 2: Open Food Facts
    try:
        resp = requests.get(
            f'https://world.openfoodfacts.org/api/v0/product/{upc_clean}.json',
            timeout=10
        )
        results['tests']['openfoodfacts'] = {
            'status_code': resp.status_code,
            'status_field': resp.json().get('status') if resp.status_code == 200 else None
        }
    except Exception as e:
        results['tests']['openfoodfacts'] = {'error': str(e)}

    # Test 3: SearXNG (replaces DuckDuckGo/Google/Bing which block datacenter IPs)
    try:
        q = f'{upc_clean} buy'
        url, snippet_title, source = _search_searxng(q)
        page_title = _fetch_product_title_from_page(url) if url else None
        results['tests']['searxng'] = {
            'query': q,
            'found_url': url,
            'snippet_title': snippet_title,
            'page_title': page_title,
            'found_source': source,
        }
    except Exception as e:
        results['tests']['searxng'] = {'error': str(e), 'traceback': traceback.format_exc()}

    # Test 3b: SearXNG with name query
    try:
        q = f'{name} buy amazon'
        url2, snippet_title2, source2 = _search_searxng(q)
        page_title2 = _fetch_product_title_from_page(url2) if url2 else None
        results['tests']['searxng_name'] = {
            'query': q,
            'found_url': url2,
            'snippet_title': snippet_title2,
            'page_title': page_title2,
            'found_source': source2,
        }
    except Exception as e:
        results['tests']['searxng_name'] = {'error': str(e)}

    # Test 4: Full lookup_product_info
    try:
        info = lookup_product_info(upc=upc, name=name)
        results['tests']['lookup_product_info'] = info
    except Exception as e:
        results['tests']['lookup_product_info'] = {'error': str(e), 'traceback': traceback.format_exc()}

    return jsonify(results)


@app.route('/enrich', methods=['POST'])
def enrich_data():
    """Look up images, retail links, and product titles for items that need them."""
    data = request.json
    session_id = data.get('session_id')
    items = data.get('items', [])
    indices = data.get('indices', [])

    enriched = []
    for i in indices:
        if i >= len(items):
            continue
        item = items[i].copy()
        upc = item.get('UPC/Item #', '')
        original_name = item.get('Item Name', '').strip()

        # Check what's missing
        img_val = item.get('Image', '')
        needs_image = not img_val or (not img_val.startswith('http') and not img_val.startswith('/serve_image/'))
        link_val = item.get('Retail Link', '')
        needs_link = not link_val or not link_val.startswith('http')

        # Look up product info
        info = lookup_product_info(upc=upc, name=original_name)

        # ── TRUST LEVEL ──
        # UPC API lookups (UPCitemdb, Open Food Facts by barcode, etc.) are TRUSTED
        # because they match on the exact barcode — the product IS correct.
        # Web search results are UNTRUSTED and need name similarity validation.
        if info['upc_matched']:
            # Direct UPC barcode match — trust the result completely
            lookup_is_relevant = True
            if info['title']:
                print(f'Enrichment UPC MATCH: "{original_name}" → "{info["title"]}" (source: {info["source"]})')
        elif info['title'] and original_name:
            # Web search result — validate against original name
            sim = _name_similarity(original_name, info['title'])
            if sim >= 0.2:
                lookup_is_relevant = True
                print(f'Enrichment NAME MATCH (sim={sim:.2f}): "{original_name}" → "{info["title"]}"')
            else:
                lookup_is_relevant = False
                print(f'Enrichment REJECTED (sim={sim:.2f}): "{original_name}" ≠ "{info["title"]}" — keeping original')
        elif info['title'] and not original_name:
            # No original name to compare — accept whatever we found
            lookup_is_relevant = True
        else:
            lookup_is_relevant = False

        # ── Item Name enrichment ──
        if lookup_is_relevant and info['title']:
            item['Item Name'] = info['title']
            if info['source']:
                item['_name_source'] = info['source']

        # ── Image enrichment ──
        if needs_image and info['image'] and lookup_is_relevant:
            item['Image'] = info['image']

        # ── Retail link enrichment ──
        if needs_link and info['retail_link'] and lookup_is_relevant:
            retail_url = info['retail_link']
            # Reject search page URLs — only accept actual product pages
            is_search_page = (
                '/s?' in retail_url or
                '/s/' in retail_url or
                '/search?' in retail_url or
                '/search/' in retail_url or
                'query=' in retail_url or
                'keywords=' in retail_url
            )
            product_url, _ = _is_product_page_url(retail_url)
            if product_url:
                item['Retail Link'] = product_url
            elif not is_search_page:
                item['Retail Link'] = retail_url

        enriched.append({'index': i, 'item': item})

    return jsonify({'enriched': enriched})


@app.route('/export', methods=['POST'])
def export_excel():
    data = request.json
    items = data.get('items', [])

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Sheet1'

    # Header styling matching DLX template
    header_font_white = Font(name='Aptos Narrow', size=16, bold=True, color='FFFFFF')
    header_fill = PatternFill('solid', fgColor='4472C4')
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell_font = Font(name='Aptos Narrow', size=11)
    cell_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    link_font = Font(name='Aptos Narrow', size=11, color='0563C1', underline='single')
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    alt_fill = PatternFill('solid', fgColor='D9E2F3')

    # Write headers
    for col_idx, col_name in enumerate(TEMPLATE_COLUMNS, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font_white
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border
    ws.row_dimensions[1].height = 46

    # Column widths
    widths = [23, 30, 16, 14, 10, 12, 10, 20, 16]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Write data rows with embedded images
    for row_idx, item in enumerate(items, 2):
        ws.row_dimensions[row_idx].height = 100
        fill = alt_fill if row_idx % 2 == 0 else PatternFill(fill_type=None)

        for col_idx, col_name in enumerate(TEMPLATE_COLUMNS, 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            val = item.get(col_name, '')

            if col_name == 'Image' and val:
                img_embedded = False
                img_path = None

                # Check if it's a local extracted image
                if val.startswith('/serve_image/'):
                    local_filename = val.replace('/serve_image/', '')
                    local_path = os.path.join(UPLOAD_FOLDER, secure_filename(local_filename))
                    if os.path.exists(local_path):
                        img_path = local_path

                # Check if it's a remote URL
                elif val.startswith('http'):
                    try:
                        img_data, fmt = download_image(val)
                        if img_data:
                            img_data, fmt = resize_image_bytes(img_data, max_width=120, max_height=100)
                            dl_path = os.path.join(UPLOAD_FOLDER, f'img_{row_idx}.{fmt}')
                            with open(dl_path, 'wb') as f:
                                f.write(img_data)
                            img_path = dl_path
                    except Exception as e:
                        print(f'Image download failed for row {row_idx}: {e}')

                if img_path and os.path.exists(img_path):
                    try:
                        img = XlImage(img_path)
                        img.width = 100
                        img.height = 90
                        cell_ref = f'A{row_idx}'
                        ws.add_image(img, cell_ref)
                        cell.value = ''
                        img_embedded = True
                    except Exception as e:
                        print(f'Image embed failed for row {row_idx}: {e}')

                if not img_embedded:
                    if val.startswith('http'):
                        cell.value = 'View Image'
                        cell.hyperlink = val
                        cell.font = link_font
                    else:
                        cell.value = ''
                        cell.font = cell_font

            elif col_name == 'Retail Link' and val and val.startswith('http'):
                cell.value = 'View Product'
                cell.hyperlink = val
                cell.font = link_font

            elif col_name == 'Cost':
                try:
                    cost_val = float(re.sub(r'[^\d.]', '', str(val))) if val else 0
                    cell.value = cost_val
                    cell.number_format = '$#,##0.00'
                    cell.font = cell_font
                except (ValueError, TypeError):
                    cell.value = val
                    cell.font = cell_font
            else:
                cell.value = val
                cell.font = cell_font

            cell.alignment = cell_align
            cell.border = thin_border
            if fill.fgColor:
                cell.fill = fill

    # Freeze header row
    ws.freeze_panes = 'A2'

    # Save to temp file
    output_path = os.path.join(UPLOAD_FOLDER, 'DLX_Offer_Export.xlsx')
    wb.save(output_path)

    return send_file(output_path, as_attachment=True,
                     download_name='DLX_Distribution_Offer.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=os.environ.get('FLASK_DEBUG', 'false').lower() == 'true')
