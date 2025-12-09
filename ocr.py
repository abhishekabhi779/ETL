
import re
import json
from typing import List, Dict, Any, Tuple
from pypdf import PdfReader

# ----------------------------
# Helpers
# ----------------------------
def normalize_amount(text: str) -> float:
    """
    Convert a money-like string to float, fixing thousand splits like '10,00 0.00' -> '10000.00'.
    Keeps digits, commas, dots. Removes spaces around digit groups and '$'.
    """
    if not text:
        return 0.0
    # Remove currency symbol and spaces
    cleaned = re.sub(r'[\s$]', '', text)
    # Fix cases like '10,00 0.00' (remove stray spaces already), any remaining commas are thousand separators
    cleaned = cleaned.replace(',', '')
    try:
        return float(cleaned)
    except ValueError:
        # Last resort: keep only digits and dot
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

def find_between(text: str, start_marker: str, end_marker: str) -> str:
    """
    Return substring between start_marker and end_marker (first occurrences).
    """
    s = text.find(start_marker)
    if s == -1:
        return ""
    e = text.find(end_marker, s + len(start_marker))
    if e == -1:
        return text[s + len(start_marker):]
    return text[s + len(start_marker):e]

def clean_whitespace(s: str) -> str:
    return re.sub(r'\s+', ' ', s).strip()

def extract_key_value(line: str, key: str) -> str:
    """
    Extract value that follows 'key:' in a line. Designed for 'Bill To:' style.
    """
    m = re.search(rf"{re.escape(key)}\s*:\s*(.+)", line, flags=re.IGNORECASE)
    return m.group(1).strip() if m else ""

# ----------------------------
# Parsers for sections
# ----------------------------

def parse_header(full_text: str) -> Dict[str, Any]:
    """
    Extract Quote Number, Quote Date, Quote Expiration Date
    """
    header = {}
    mq = re.search(r'QUOTE NUMBER\s+([A-Z0-9\-]+)', full_text, flags=re.IGNORECASE)
    header['quote_number'] = mq.group(1) if mq else None

    md = re.search(r'Quote Date:\s*([0-9/]{8,10})', full_text, flags=re.IGNORECASE)
    header['quote_date'] = md.group(1) if md else None

    me = re.search(r'Quote Expiration Date:\s*([0-9/]{8,10})', full_text, flags=re.IGNORECASE)
    header['quote_expiration_date'] = me.group(1) if me else None

    return header

def parse_billing_information(full_text: str) -> Dict[str, Any]:
    """
    Extracts Section A: Bill To, Ship To, Partner, End User blocks.
    """
    block = find_between(full_text, "A. Billing Information", "B. Billing terms")
    lines = [clean_whitespace(l) for l in block.splitlines() if l.strip()]
    # We'll collect relevant fields by scanning lines
    out = {
        "bill_to": {"company": None, "address": None},
        "ship_to": {"company": None, "address": None},
        "partner": {"legal_name": None, "tier": None, "address": None},
        "end_user": {"legal_name": None, "address": None}
    }

    # Heuristics based on your PDF structure
    for i, line in enumerate(lines):
        if "Bill To:" in line:
            # Next lines contain company/address
            # Company indicated by '** Ingram Micro Inc.**' in text flow; grab up to '**'
            comp = extract_key_value(line, "Bill To") or None
            # If empty, try neighbor lines (company often appears as '** Ingram Micro Inc.**')
            if not comp and i + 1 < len(lines):
                comp = re.sub(r'\*', '', lines[i+1]).strip()
            out["bill_to"]["company"] = comp

        if "Bill To Address" in line:
            out["bill_to"]["address"] = extract_key_value(line, "Bill To Address")

        if "Ship to:" in line or "Ship To Company Legal Name" in line:
            # Similar extraction
            comp = extract_key_value(line, "Ship to")
            if not comp:
                comp = extract_key_value(line, "Ship To Company Legal Name")
            out["ship_to"]["company"] = comp or out["ship_to"]["company"]

        if "Ship To Address" in line:
            out["ship_to"]["address"] = extract_key_value(line, "Ship To Address")

        if "Partner Legal Name" in line:
            out["partner"]["legal_name"] = extract_key_value(line, "Partner Legal Name")

        if "Partner Tier Level" in line:
            out["partner"]["tier"] = extract_key_value(line, "Partner Tier Level")

        if "Partner Address" in line:
            out["partner"]["address"] = extract_key_value(line, "Partner Address")

        if "End User Legal Name" in line:
            out["end_user"]["legal_name"] = extract_key_value(line, "End User Legal Name")

        if "Address:" in line and ("End User" in lines[i-1] if i > 0 else False):
            out["end_user"]["address"] = extract_key_value(line, "Address")

    # Fallbacks: scrub asterisks
    for k in ["bill_to", "ship_to"]:
        for f in out[k]:
            if isinstance(out[k][f], str):
                out[k][f] = out[k][f].replace("*", "").strip()

    return out

def parse_billing_terms(full_text: str) -> Dict[str, Any]:
    block = find_between(full_text, "B. Billing terms", "C. Software Pricing Detail")
    lines = [clean_whitespace(l) for l in block.splitlines() if l.strip()]
    out = {"payment_term": None, "billing_cycle": None, "currency": None, "quote_total": None, "estimated_partner_rebate": None}
    for line in lines:
        if "Payment term" in line:
            out["payment_term"] = extract_key_value(line, "Payment term")
        elif "Billing cycle" in line:
            out["billing_cycle"] = extract_key_value(line, "Billing cycle")
        elif re.search(r'\bCurrency\b', line, flags=re.IGNORECASE):
            out["currency"] = extract_key_value(line, "Currency")
        elif re.search(r'Quote Total', line, flags=re.IGNORECASE):
            amt = re.search(r'Quote Total\s*\$?([0-9\$\.,\s]+)', line)
            out["quote_total"] = normalize_amount(amt.group(1)) if amt else None
        elif re.search(r'Estimated Partner Rebate', line, flags=re.IGNORECASE):
            amt = re.search(r'Estimated Partner Rebate\s*\$?([0-9\$\.,\s]+)', line)
            out["estimated_partner_rebate"] = normalize_amount(amt.group(1)) if amt else None
    return out

def parse_items(full_text: str) -> Dict[str, Any]:
    """
    Parse Section C items from the pricing table.
    """
    block = find_between(full_text, "C. Software Pricing Detail", "Quote legal terms")
    net_total_match = re.search(r'Net Total Software\s*\$?([0-9\$\.,\s]+)', block)
    net_total_software = normalize_amount(net_total_match.group(1)) if net_total_match else None

    items = []
    
    # Split by UiPath product lines
    raw_items = re.split(r'(?=UiPath\s*-)', block)
    
    for raw in raw_items:
        if not raw.strip().startswith('UiPath'):
            continue
        
        # Rejoin for processing
        lines = raw.split('\n')
        item_text = ' '.join(line.strip() for line in lines if line.strip())
        
        # Clean up spacing issues in text
        item_text = item_text.replace('Concurre nt', 'Concurrent')
        
        # Extract description, product code, and quantity in one pattern
        # Description is everything from UiPath up to the product code
        # Product code is uppercase letters/numbers (may contain spaces like "NU000")
        match = re.search(r'(UiPath.*?)\s+([A-Z]+[A-Z0-9\s]*)\s+(\d+)\s+(?:Each|N/A)', item_text)
        
        if match:
            description = clean_whitespace(match.group(1))
            product_code = clean_whitespace(match.group(2))
            quantity = int(match.group(3))
        else:
            description = None
            product_code = None
            quantity = None
        
        # Unit of Measure
        if 'Each/User' in item_text:
            uom = 'Each/User per year'
        elif 'Each' in item_text and 'Each/User' not in item_text:
            uom = 'Each'
        elif 'N/A' in item_text:
            uom = 'N/A'
        else:
            uom = None
        
        # License Model
        if 'Named User' in item_text:
            license_model = 'Named User'
        elif 'Concurrent' in item_text:
            license_model = 'Concurrent Runtime'
        else:
            license_model = None
        
        # Dates: MM/DD/YYYY format
        dates = re.findall(r'(\d{2}/\d{2}/\d{4})', item_text)
        term_start = dates[0] if len(dates) > 0 else None
        term_end = dates[1] if len(dates) > 1 else None
        
        # Prices: Extract dollar amounts with flexible spacing
        # Pattern accounts for spaces in numbers like "4,400. 00"
        price_pattern = r'\$[\s]*([0-9,]+)[\s]*\.[\s]*(\d+)'
        prices = re.findall(price_pattern, item_text)
        
        # Convert to proper format
        prices_numeric = []
        for p in prices:
            amount_str = p[0] + '.' + p[1]  # e.g., "4400" + "." + "00"
            prices_numeric.append(normalize_amount(amount_str))
        
        # Extract discount percentage
        discount = re.search(r'(\d+)\.(\d+)%', item_text)
        discount_pct = float(f"{discount.group(1)}.{discount.group(2)}") if discount else None
        
        # Assign prices in order: list, regular, net unit, net total
        list_price = prices_numeric[0] if len(prices_numeric) > 0 else None
        regular_price = prices_numeric[1] if len(prices_numeric) > 1 else None
        net_unit_price = prices_numeric[2] if len(prices_numeric) > 2 else None
        net_total = prices_numeric[3] if len(prices_numeric) > 3 else None

        item = {
            "description": description,
            "product_code": product_code,
            "quantity": quantity,
            "unit_of_measure": uom,
            "license_model": license_model,
            "term_start_date": term_start,
            "term_end_date": term_end,
            "list_unit_price": list_price,
            "regular_unit_price": regular_price,
            "discount_percent": discount_pct,
            "net_unit_price": net_unit_price,
            "net_total_usd": net_total
        }
        items.append(item)

    return {"items": items, "net_total_software": net_total_software}


def main():
    import argparse
    
    parser = argparse.ArgumentParser(description="Extract data from PDF quotes")
    parser.add_argument("pdf_file", help="Path to the PDF file")
    parser.add_argument("-o", "--output", help="Output JSON file path (default: print to stdout)")
    parser.add_argument("-v", "--verbose", action="store_true", help="Verbose output")
    
    args = parser.parse_args()
    
    # Extract text from PDF
    pages, full_text = get_text_from_pdf(args.pdf_file)
    
    if args.verbose:
        print(f"Processing {args.pdf_file}...")
        print(f"Total pages: {len(pages)}")
    
    # Parse all sections
    result = {
        "file": args.pdf_file,
        "header": parse_header(full_text),
        "billing_information": parse_billing_information(full_text),
        "billing_terms": parse_billing_terms(full_text),
        "items": parse_items(full_text)
    }
    
    # Output
    output_json = json.dumps(result, indent=2)
    
    if args.output:
        with open(args.output, 'w') as f:
            f.write(output_json)
        if args.verbose:
            print(f"Output written to {args.output}")
    else:
        print(output_json)


if __name__ == "__main__":
    main()
