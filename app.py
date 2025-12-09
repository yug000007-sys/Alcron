import os
import re
import glob
from typing import List, Dict, Optional

import pandas as pd
from PyPDF2 import PdfReader

# --------------------------------------------------------------------
# 1. Alcorn header (same as Headeralcron.xlsx)
# --------------------------------------------------------------------
HEADER_COLUMNS: List[str] = [
    "ReferralManagerCode",
    "ReferralManager",
    "ReferralEmail",
    "Brand",
    "QuoteNumber",
    "QuoteVersion",
    "QuoteDate",
    "QuoteValidDate",
    "Customer Number/ID",
    "Company",
    "Address",
    "County",
    "City",
    "State",
    "ZipCode",
    "Country",
    "FirstName",
    "LastName",
    "ContactEmail",
    "ContactPhone",
    "Webaddress",
    "item_id",
    "item_desc",
    "UOM",
    "Quantity",
    "Unit Price",
    "List Price",
    "TotalSales",
    "Manufacturer_ID",
    "manufacturer_Name",
    "Writer Name",
    "CustomerPONumber",
    "PDF",
    "DemoQuote",
    "Duns",
    "SIC",
    "NAICS",
    "LineOfBusiness",
    "LinkedinProfile",
    "PhoneResearched",
    "PhoneSupplied",
    "ParentName",
]

BRAND_NAME = "Alcorn Industrial Inc"

# --------------------------------------------------------------------
# 2. Regex patterns for header info & money
# --------------------------------------------------------------------
DATE_PATTERN = re.compile(
    r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2},\s+\d{4}"
)
MONEY_RE = re.compile(r"^[0-9,]*\d\.\d{2}$")
QUOTE_NO_RE = re.compile(r"(QT\d{6}|RQ\d{4,}-\d+)")


# --------------------------------------------------------------------
# 3. Helpers to parse header info
# --------------------------------------------------------------------
def extract_quote_number(page_text: str) -> Optional[str]:
    m = QUOTE_NO_RE.search(page_text)
    return m.group(1) if m else None


def extract_quote_date(page_text: str) -> Optional[str]:
    m = DATE_PATTERN.search(page_text)
    return m.group(0) if m else None


def extract_customer_and_salesperson(page_text: str, quote_no: str) -> (Optional[str], Optional[str]):
    """
    Best-effort: get Customer Number/ID and ReferralManagerCode from the top area of the quote.
    Works for both QT and RQ styles you've shown.
    """
    cust = None
    sp = None

    if not quote_no:
        return cust, sp

    # Repair quotes often have RFQ line: "RFQ <...> 2109 SVC ..."
    if quote_no.startswith("RQ"):
        m = re.search(r"RFQ\s+\S+\s+([0-9A-Z\-]+)\s+([A-Z0-9]{1,3})\s+[A-Z]", page_text)
        if m:
            cust, sp = m.group(1), m.group(2)
    else:
        # Normal QT: line like "2109 CR UPS1 NET30 QT00040379"
        pattern = rf"\n([0-9A-Z\-]+)\s+([A-Z0-9]{{1,3}})\s+[A-Z0-9]+\s+[A-Z0-9]+\s+{quote_no}"
        m = re.search(pattern, page_text)
        if m:
            cust, sp = m.group(1), m.group(2)

    return cust, sp


def extract_company_block(page_text: str):
    """
    Pulls Company, Address, City, State, ZipCode, Country from the Ship To: block.
    This matches the quote PDFs you've shared.
    """
    pos = page_text.find("Ship To :")
    if pos == -1:
        return None, None, None, None, None, None

    # Grab next ~9 lines after 'Ship To :'
    block = page_text[pos:].splitlines()[1:10]
    lines = [ln.strip() for ln in block if ln.strip()]
    if not lines:
        return None, None, None, None, None, None

    # Country line
    country_line = None
    country = None
    for ln in reversed(lines):
        if "Canada" in ln:
            country = "Canada"
            country_line = ln
            break
        if "USA" in ln or "United States" in ln:
            country = "USA"
            country_line = ln
            break

    # City / State / Zip line
    city = state = zipcode = None
    cs_line = None
    for ln in lines:
        if "," in ln and any(st in ln for st in [
            " AL ", " AK ", " AZ ", " AR ", " CA ", " CO ", " CT ", " DE ", " FL ", " GA ",
            " HI ", " IA ", " ID ", " IL ", " IN ", " KS ", " KY ", " LA ", " MA ", " MD ",
            " ME ", " MI ", " MN ", " MO ", " MS ", " MT ", " NC ", " ND ", " NE ", " NH ",
            " NJ ", " NM ", " NV ", " NY ", " OH ", " OK ", " OR ", " PA ", " RI ", " SC ",
            " SD ", " TN ", " TX ", " UT ", " VA ", " VT ", " WA ", " WI ", " WV ", " WY ",
            " QC ", " ON "
        ]):
            cs_line = ln
            break

    if cs_line:
        parts = cs_line.split(",")
        city = parts[0].strip()
        rest = ",".join(parts[1:]).strip()
        rest_tokens = rest.split()
        if rest_tokens:
            state = rest_tokens[0]
        if len(rest_tokens) > 1:
            zipcode = " ".join(rest_tokens[1:])

    # Address line (street with digits) before city/state line
    address = None
    if cs_line:
        idx_cs = lines.index(cs_line)
        for ln in reversed(lines[:idx_cs]):
            if any(ch.isdigit() for ch in ln):
                address = ln
                break

    # Company line – first non-empty, non-address, non-city/country
    company = None
    for ln in lines:
        if ln in (cs_line, country_line, address):
            continue
        if ln.upper().startswith("ATTN:") or "@" in ln:
            continue
        company = ln
        break

    # Infer country if missing
    if country is None and state in ("QC", "ON"):
        country = "Canada"
    if country is None:
        country = "USA"

    return company, address, city, state, zipcode, country


# --------------------------------------------------------------------
# 4. Parse line items from text lines
# --------------------------------------------------------------------
def parse_line_item(line: str) -> Optional[Dict]:
    """
    Parse one line that contains a line item.
    Assumes format like:
        "1  QXXD5AT080ES08 80Nm Angle Tool BT/Wireless ETS  8,222.00 EA 8,222.00"
    """
    s = line.strip()
    if not s:
        return None

    tokens = s.split()
    money_idxs = [i for i, t in enumerate(tokens) if MONEY_RE.match(t)]
    if len(money_idxs) < 2:
        return None  # need unit price + total

    i2 = money_idxs[-1]     # total
    i1 = money_idxs[-2]     # unit

    try:
        unit_price = float(tokens[i1].replace(",", ""))
        total_sales = float(tokens[i2].replace(",", ""))
    except ValueError:
        return None

    # leading quantity
    qty_idxs = []
    for i, t in enumerate(tokens):
        if t.isdigit():
            qty_idxs.append(i)
        else:
            break
    if not qty_idxs:
        return None

    qty = int(tokens[qty_idxs[0]])

    # body tokens between qty and unit price
    start_body = qty_idxs[-1] + 1
    end_body = i1 - 1
    if start_body > end_body:
        return None
    body_tokens = tokens[start_body:end_body + 1]

    # UOM tokens between unit and total
    uom_tokens = tokens[i1 + 1:i2]
    uom = " ".join(uom_tokens) if uom_tokens else None

    # Special case: sometimes you get "1 0 1 DYNA 95600 KIT..."
    while body_tokens and body_tokens[0] == "0":
        body_tokens = body_tokens[1:]

    if not body_tokens:
        return None

    item_id = body_tokens[0]
    item_desc = " ".join(body_tokens[1:]) if len(body_tokens) > 1 else None

    return {
        "Quantity": qty,
        "item_id": item_id,
        "item_desc": item_desc,
        "UOM": uom,
        "UnitPrice": unit_price,
        "TotalSales": total_sales,
        "raw": s,
    }


# --------------------------------------------------------------------
# 5. Process a single PDF into rows
# --------------------------------------------------------------------
def process_pdf(pdf_path: str) -> List[Dict]:
    rows: List[Dict] = []
    reader = PdfReader(pdf_path)
    num_pages = len(reader.pages)

    current_quote_no = None
    current_quote_date = None
    current_cust_no = None
    current_sp = None
    current_company = None
    current_address = None
    current_city = None
    current_state = None
    current_zip = None
    current_country = None

    for page_idx in range(num_pages):
        page = reader.pages[page_idx]
        text = page.extract_text() or ""

        # If this page has a new quote number, refresh header info
        q_no = extract_quote_number(text)
        if q_no:
            current_quote_no = q_no
            current_quote_date = extract_quote_date(text)
            current_cust_no, current_sp = extract_customer_and_salesperson(text, current_quote_no)
            (current_company,
             current_address,
             current_city,
             current_state,
             current_zip,
             current_country) = extract_company_block(text)

        # If we still don't have a quote number, skip this page
        if not current_quote_no:
            continue

        lines = text.splitlines()
        in_items = False

        for line in lines:
            if "Please send your order to:" in line:
                in_items = True
                continue
            if in_items and line.strip().startswith("Tax Summary"):
                in_items = False

            if not in_items:
                continue

            parsed = parse_line_item(line)
            if not parsed:
                continue

            row = {h: None for h in HEADER_COLUMNS}
            row["Brand"] = BRAND_NAME
            row["QuoteNumber"] = current_quote_no
            row["QuoteDate"] = current_quote_date
            if current_cust_no:
                row["Customer Number/ID"] = current_cust_no
            if current_sp:
                row["ReferralManagerCode"] = current_sp
            if current_company:
                row["Company"] = current_company
            if current_address:
                row["Address"] = current_address
            if current_city:
                row["City"] = current_city
            if current_state:
                row["State"] = current_state
            if current_zip:
                row["ZipCode"] = current_zip
            if current_country:
                row["Country"] = current_country

            row["item_id"] = parsed["item_id"]
            row["item_desc"] = parsed["item_desc"]
            row["UOM"] = parsed["UOM"]
            row["Quantity"] = parsed["Quantity"]
            row["Unit Price"] = parsed["UnitPrice"]
            row["TotalSales"] = parsed["TotalSales"]
            row["PDF"] = os.path.basename(pdf_path)

            rows.append(row)

    return rows


# --------------------------------------------------------------------
# 6. Main: process all PDFs in a folder and write merged Excel
# --------------------------------------------------------------------
def main():
    input_folder = "input_pdfs"
    output_folder = "output"
    os.makedirs(output_folder, exist_ok=True)

    pdf_files = sorted(glob.glob(os.path.join(input_folder, "*.pdf")))
    if not pdf_files:
        print(f"No PDFs found in {input_folder}/")
        return

    all_rows: List[Dict] = []
    for pdf_path in pdf_files:
        print(f"Processing: {pdf_path}")
        rows = process_pdf(pdf_path)
        all_rows.extend(rows)

    if not all_rows:
        print("No line items found in any PDF.")
        return

    df = pd.DataFrame(all_rows, columns=HEADER_COLUMNS)
    out_path = os.path.join(output_folder, "Alcorn_Quotes_Merged.xlsx")
    df.to_excel(out_path, index=False)
    print(f"Done. Wrote {len(df)} rows to {out_path}")


if __name__ == "__main__":
    main()
