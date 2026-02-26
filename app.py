import os, io, re, json, tempfile, urllib.parse
from flask import Flask, request, jsonify, render_template, send_file, session
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XlImage
import requests
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
def lookup_upc_image(upc):
    """Look up product image by UPC code using multiple free APIs."""
    if not upc or str(upc).lower() in ('na', '', 'nan', 'none'):
        return None
    upc_clean = re.sub(r'[^0-9]', '', str(upc))
    if not upc_clean:
        return None

    # Try UPCitemdb free API
    try:
        resp = requests.get(
            f'https://api.upcitemdb.com/prod/trial/lookup?upc={upc_clean}',
            timeout=8,
            headers={'Accept': 'application/json'}
        )
        if resp.status_code == 200:
            data = resp.json()
            items = data.get('items', [])
            if items:
                images = items[0].get('images', [])
                if images:
                    return images[0]
    except Exception:
        pass

    # Try Open Food Facts
    try:
        resp = requests.get(
            f'https://world.openfoodfacts.org/api/v0/product/{upc_clean}.json',
            timeout=8
        )
        if resp.status_code == 200:
            data = resp.json()
            product = data.get('product', {})
            img = product.get('image_url') or product.get('image_front_url') or product.get('image_front_small_url')
            if img:
                return img
    except Exception:
        pass

    # Try Go-UPC API (free tier)
    try:
        resp = requests.get(
            f'https://go-upc.com/api/v1/code/{upc_clean}',
            timeout=8,
            headers={'Accept': 'application/json'}
        )
        if resp.status_code == 200:
            data = resp.json()
            img = data.get('product', {}).get('imageUrl')
            if img:
                return img
    except Exception:
        pass

    # Try Barcode Lookup
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
                images = products[0].get('images', [])
                if images:
                    return images[0]
    except Exception:
        pass

    return None


def search_product_image(name):
    """Search for a product image by name using free APIs."""
    if not name:
        return None

    # Try Open Food Facts search
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
                img = products[0].get('image_url') or products[0].get('image_front_url') or products[0].get('image_front_small_url')
                if img:
                    return img
    except Exception:
        pass

    return None


def find_retail_link(upc=None, name=None):
    links = {}
    upc_clean = re.sub(r'[^0-9]', '', str(upc)) if upc else ''

    if upc_clean:
        # Try UPCitemdb for offers/links
        try:
            resp = requests.get(
                f'https://api.upcitemdb.com/prod/trial/lookup?upc={upc_clean}',
                timeout=8,
                headers={'Accept': 'application/json'}
            )
            if resp.status_code == 200:
                data = resp.json()
                items = data.get('items', [])
                if items:
                    offers = items[0].get('offers', [])
                    for offer in offers:
                        link = offer.get('link')
                        merchant = offer.get('merchant', '').lower()
                        if link:
                            if 'amazon' in merchant:
                                return link
                            elif 'walmart' in merchant:
                                return link
                            elif not links.get('other'):
                                links['other'] = link
        except Exception:
            pass

    if links.get('other'):
        return links['other']

    # Fallback: generate search URLs
    if upc_clean:
        return f'https://www.amazon.com/s?k={upc_clean}'
    elif name:
        q = urllib.parse.quote_plus(str(name))
        return f'https://www.amazon.com/s?k={q}'
    return ''


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


@app.route('/enrich', methods=['POST'])
def enrich_data():
    """Look up images and retail links for items that need them."""
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
        name = item.get('Item Name', '')

        # Look up image if missing or not a valid URL/path
        img_val = item.get('Image', '')
        needs_image = not img_val or (not img_val.startswith('http') and not img_val.startswith('/serve_image/'))
        if needs_image:
            img_url = lookup_upc_image(upc)
            if not img_url:
                img_url = search_product_image(name)
            if img_url:
                item['Image'] = img_url

        # Look up retail link if missing
        link_val = item.get('Retail Link', '')
        if not link_val or not link_val.startswith('http'):
            link = find_retail_link(upc=upc, name=name)
            item['Retail Link'] = link

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
