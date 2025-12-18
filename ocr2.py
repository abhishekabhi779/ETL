import re
import json
import sys
from typing import List, Dict, Any, Tuple, Optional

# PDF text extraction
from pypdf import PdfReader

# PDFPlumber for table extraction
try:
    import pdfplumber
    _HAS_PDFPLUMBER = True
except Exception:
    pdfplumber = None
    _HAS_PDFPLUMBER = False


def normalize_amount(text: str) -> float:
    if not text:
        return 0.0
    cleaned = re.sub(r'[\s$]', '', text)
    cleaned = cleaned.replace(',', '')
    try:
        return float(cleaned)
    except ValueError:
        cleaned = re.sub(r'[^0-9\.]', '', cleaned)
        return float(cleaned or 0.0)


def normalize_percent(text: str) -> float:
    if not text:
        return 0.0
    cleaned = text.replace('%', '').strip()
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def get_text_from_pdf(pdf_path: str) -> Tuple[List[str], str]:
    reader = PdfReader(pdf_path)
    pages = []
    for p in reader.pages:
        pages.append(p.extract_text() or "")
    full = "\n\n".join(pages)
    return pages, full


def get_tables_from_pdf_pdfplumber(pdf_path: str) -> List:
    """
    Use pdfplumber to extract tables. Returns list of DataFrames.
    """
    if not _HAS_PDFPLUMBER:
        return []
    tables = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_tables = page.extract_tables()
                for table in page_tables:
                    if table and len(table) > 1:  # Skip empty or single-row tables
                        tables.append(table)
    except Exception as e:
        print(f"Error extracting tables with pdfplumber: {e}")
    return tables


def map_table_headers(header_row: List) -> Dict[str, int]:
    """
    Map column headers to their indices. Returns a dict mapping field names to column indices.
    """
    if not header_row:
        return {}
    
    # Clean headers - remove newlines and extra spaces
    cleaned_headers = []
    for h in header_row:
        if h:
            # Replace newlines with spaces and clean
            cleaned = str(h).replace('\n', ' ').strip()
            cleaned = re.sub(r'\s+', ' ', cleaned)
            cleaned_headers.append(cleaned.lower())
        else:
            cleaned_headers.append("")
    
    # Common header variations - ordered by specificity (more specific first)
    header_mappings = {
        'description': ['software description', 'product description', 'item description', 'description', 'desc'],
        'product_code': ['product code', 'product_code', 'item code', 'code', 'sku'],
        'quantity': ['quantity', 'qty', 'qnty'],
        'unit_of_measure': ['unit of measure', 'unit of\nmeasure', 'uom', 'unit', 'units', 'measure'],
        'license_model': ['license modelw', 'license model', 'license\nmodel', 'license type', 'license', 'model'],
        'term_start_date': ['license term start date', 'term start date', 'license\nterm\nstart date', 'start date', 'term start', 'start', 'begin date'],
        'term_end_date': ['license term end date', 'term end date', 'license\nterm end\ndate', 'end date', 'term end', 'end', 'expiration date'],
        'list_unit_price': ['list unit price', 'list price', 'unit price', 'list', 'base price'],
        'discount_percent': ['total discount %', 'discount %', 'discount percent', 'disc %', 'discount rate', 'discount'],
        'net_unit_price': ['net unit price', 'net price', 'discounted price'],
        'net_total_usd': ['net total usd', 'net total\nusd', 'net total', 'extended price', 'net amount', 'total', 'amount']
    }
    
    mapping = {}
    for i, header in enumerate(cleaned_headers):
        if not header:
            continue
        header_lower = header.lower()
        
        # Find the best match for this header
        best_field = None
        best_score = 0
        
        for field, variations in header_mappings.items():
            for variation in variations:
                variation_lower = variation.lower()
                if variation_lower in header_lower:
                    # Score based on match quality (longer matches are better)
                    score = len(variation_lower)
                    if score > best_score:
                        best_score = score
                        best_field = field
        
        if best_field and best_field not in mapping:
            mapping[best_field] = i
    
    return mapping


def clean_whitespace(s: str) -> str:
    return re.sub(r'\s+', ' ', s).strip()


# The main parser using pdfplumber-extracted tables (with a text fallback)
def parse_items_from_pdf(pdf_path: str, verbose: bool = False) -> Dict[str, Any]:
    pages, full_text = get_text_from_pdf(pdf_path)
    block = re.search(r'C\. Software Pricing Detail(.*?)Quote legal terms', full_text, flags=re.S | re.I)
    block_text = block.group(1) if block else full_text

    net_total_match = re.search(r'Net Total Software\s*\$?([0-9\$\.,\s]+)', block_text)
    net_total_software = normalize_amount(net_total_match.group(1)) if net_total_match else None

    items: List[Dict[str, Any]] = []

    parsed_from_tables = False
    tables = get_tables_from_pdf_pdfplumber(pdf_path)
    if tables:
        parsed_from_tables = True
        for table in tables:
            # Process each table row
            try:
                if len(table) < 2:
                    continue
                    
                # First row is header
                header_row = table[0]
                header_mapping = map_table_headers(header_row)
                
                if verbose:
                    print(f"Found table with {len(table)} rows")
                    print(f"Header mapping: {header_mapping}")
                
                if not header_mapping:
                    # If no headers found, skip this table
                    continue
                
                # Process data rows
                for row in table[1:]:  # Start from index 1 to skip header
                    # Skip empty rows
                    if not row or all(not cell or str(cell).strip() == '' for cell in row):
                        continue
                    
                    # Extract data using header mapping
                    description = None
                    if 'description' in header_mapping:
                        col_idx = header_mapping['description']
                        if col_idx < len(row):
                            description = clean_whitespace(str(row[col_idx])) if row[col_idx] else None
                    
                    # Check if this is the "Net Total Software" row
                    if description and 'Net Total Software' in description:
                        # This is the total row, extract the total amount
                        if 'net_total_usd' in header_mapping:
                            col_idx = header_mapping['net_total_usd']
                            if col_idx < len(row) and row[col_idx]:
                                net_total_from_row = normalize_amount(str(row[col_idx]))
                                if net_total_from_row:
                                    net_total_software = net_total_from_row
                        continue  # Skip adding this as a regular item
                    
                    product_code = None
                    if 'product_code' in header_mapping:
                        col_idx = header_mapping['product_code']
                        if col_idx < len(row):
                            product_code = clean_whitespace(str(row[col_idx])) if row[col_idx] else None
                    
                    # Quantity
                    quantity = None
                    if 'quantity' in header_mapping:
                        col_idx = header_mapping['quantity']
                        if col_idx < len(row) and row[col_idx]:
                            try:
                                quantity = int(str(row[col_idx]).strip())
                            except:
                                pass
                    
                    uom = None
                    if 'unit_of_measure' in header_mapping:
                        col_idx = header_mapping['unit_of_measure']
                        if col_idx < len(row):
                            uom = clean_whitespace(str(row[col_idx])) if row[col_idx] else None
                    
                    license_model = None
                    if 'license_model' in header_mapping:
                        col_idx = header_mapping['license_model']
                        if col_idx < len(row):
                            license_model = clean_whitespace(str(row[col_idx])) if row[col_idx] else None
                    
                    term_start = None
                    if 'term_start_date' in header_mapping:
                        col_idx = header_mapping['term_start_date']
                        if col_idx < len(row):
                            term_start = clean_whitespace(str(row[col_idx])) if row[col_idx] else None
                    
                    term_end = None
                    if 'term_end_date' in header_mapping:
                        col_idx = header_mapping['term_end_date']
                        if col_idx < len(row):
                            term_end = clean_whitespace(str(row[col_idx])) if row[col_idx] else None
                    
                    # Prices
                    list_price = None
                    if 'list_unit_price' in header_mapping:
                        col_idx = header_mapping['list_unit_price']
                        if col_idx < len(row):
                            list_price = normalize_amount(str(row[col_idx])) if row[col_idx] else None
                    
                    discount_pct = None
                    if 'discount_percent' in header_mapping:
                        col_idx = header_mapping['discount_percent']
                        if col_idx < len(row):
                            discount_pct = normalize_percent(str(row[col_idx])) if row[col_idx] else None
                    
                    net_unit_price = None
                    if 'net_unit_price' in header_mapping:
                        col_idx = header_mapping['net_unit_price']
                        if col_idx < len(row):
                            net_unit_price = normalize_amount(str(row[col_idx])) if row[col_idx] else None
                    
                    net_total = None
                    if 'net_total_usd' in header_mapping:
                        col_idx = header_mapping['net_total_usd']
                        if col_idx < len(row):
                            net_total = normalize_amount(str(row[col_idx])) if row[col_idx] else None
                    
                    item = {
                        "description": description,
                        "product_code": product_code,
                        "quantity": quantity,
                        "unit_of_measure": uom,
                        "license_model": license_model,
                        "term_start_date": term_start,
                        "term_end_date": term_end,
                        "list_unit_price": list_price,
                        "discount_percent": discount_pct,
                        "net_unit_price": net_unit_price,
                        "net_total_usd": net_total
                    }
                    items.append(item)
                    
            except Exception as e:
                print(f"Error processing table row: {e}")
                continue

    # fallback to text parsing if needed
    if not parsed_from_tables or not items:
        items = []
        raw_items = re.split(r'(?=UiPath\s*-)', block_text)
        for raw in raw_items:
            if not raw.strip().startswith('UiPath'):
                continue
            lines = raw.split('\n')
            item_text = ' '.join(line.strip() for line in lines if line.strip())
            item = _parse_item_text(item_text)
            items.append(item)

    # Clean and deduplicate parsed items: remove rows with no useful data
    def _clean_and_dedupe(raw_items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        cleaned: List[Dict[str, Any]] = []
        seen = set()
        for it in raw_items:
            # normalize quantity: treat 0 as None
            if it.get('quantity') == 0:
                it['quantity'] = None

            # Determine if the row contains meaningful information
            has_meaning = any([
                it.get('product_code'),
                (isinstance(it.get('description'), str) and 'UiPath' in it.get('description')),
                it.get('list_unit_price') is not None,
                it.get('net_unit_price') is not None,
                it.get('net_total_usd') is not None,
                it.get('term_start_date') is not None,
            ])
            if not has_meaning:
                continue

            key = (
                it.get('product_code'),
                it.get('quantity'),
                it.get('list_unit_price'),
                it.get('net_unit_price'),
                it.get('term_start_date'),
                it.get('term_end_date'),
            )
            if key in seen:
                continue
            seen.add(key)
            cleaned.append(it)
        return cleaned

    items = _clean_and_dedupe(items)

    # Calculate sum of net_total_usd from items with descriptions
    calculated_total = 0.0
    for item in items:
        if item.get('description') and item.get('net_total_usd'):
            calculated_total += item.get('net_total_usd')

    # Check if calculated total matches net_total_software
    total_matches = abs(calculated_total - (net_total_software or 0)) < 0.01

    return {
        "items": items, 
        "net_total_software": net_total_software,
        "calculated_total_from_items": calculated_total,
        "totals_match": total_matches
    }


def _parse_item_text(item_text: str) -> Dict[str, Any]:
    item_text = item_text.replace('Concurre nt', 'Concurrent')

    match = re.search(r'(UiPath.*?)\s+([A-Z]+[A-Z0-9\-\s]*)\s+(\d+)\s+(?:Each|N/A)', item_text)
    if match:
        description = clean_whitespace(match.group(1))
        product_code = clean_whitespace(match.group(2))
        try:
            quantity = int(match.group(3))
        except Exception:
            quantity = None
    else:
        description = None
        product_code = None
        # try to pull first integer as quantity
        mqty = re.search(r'\b(\d+)\b', item_text)
        quantity = int(mqty.group(1)) if mqty else None

    if 'Each/User' in item_text:
        uom = 'Each/User per year'
    elif 'Each' in item_text and 'Each/User' not in item_text:
        uom = 'Each'
    elif 'N/A' in item_text:
        uom = 'N/A'
    else:
        uom = None

    if 'Named User' in item_text:
        license_model = 'Named User'
    elif 'Concurrent' in item_text:
        license_model = 'Concurrent Runtime'
    else:
        license_model = None

    dates = re.findall(r'(\d{2}/\d{2}/\d{4})', item_text)
    term_start = dates[0] if len(dates) > 0 else None
    term_end = dates[1] if len(dates) > 1 else None

    price_pattern = r'\$[\s]*([0-9,]+)[\s]*\.[\s]*(\d+)'
    prices = re.findall(price_pattern, item_text)
    prices_numeric = []
    for p in prices:
        amount_str = p[0] + '.' + p[1]
        prices_numeric.append(normalize_amount(amount_str))

    discount = re.search(r'(\d+\.?\d*)\s*%', item_text)
    discount_pct = normalize_percent(discount.group(1)) if discount else None

    list_price = prices_numeric[0] if len(prices_numeric) > 0 else None
    net_unit_price = prices_numeric[1] if len(prices_numeric) > 1 else None
    net_total = prices_numeric[2] if len(prices_numeric) > 2 else None

    return {
        "description": description,
        "product_code": product_code,
        "quantity": quantity,
        "unit_of_measure": uom,
        "license_model": license_model,
        "term_start_date": term_start,
        "term_end_date": term_end,
        "list_unit_price": list_price,
        "discount_percent": discount_pct,
        "net_unit_price": net_unit_price,
        "net_total_usd": net_total
    }


def main(argv=None):
    import argparse
    parser = argparse.ArgumentParser(description="Extract data from PDF quotes using PDFPlumber")
    parser.add_argument('pdf_file', help='Path to the PDF file')
    parser.add_argument('-o', '--output', help='Output JSON file path (default: print to stdout)')
    parser.add_argument('-v', '--verbose', action='store_true', help='Verbose output')
    args = parser.parse_args(argv)

    if args.verbose:
        print('PDFPlumber available:', _HAS_PDFPLUMBER)

    result = parse_items_from_pdf(args.pdf_file, args.verbose)
    result['file'] = args.pdf_file

    output_json = json.dumps(result, indent=2)
    if args.output:
        with open(args.output, 'w', encoding='utf-8') as f:
            f.write(output_json)
        if args.verbose:
            print(f'Wrote output to {args.output}')
    else:
        print(output_json)


if __name__ == '__main__':
    main()
