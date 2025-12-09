import os
import re
import glob
from typing import List, Dict, Optional, Tuple

import pandas as pd
from PyPDF2 import PdfReader

# ================================================================
# 1. ALCORN HEADER (must match your Excel header)
# ================================================================
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

# ================================================================
# 2. REGEX PATTERNS
# ================================================================
DATE_PATTERN = re.compile(
    r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2},\s+\d{4}"
)
MONEY_RE = re.compile(r"^[0-9,]*\d\.\d{2}$")
QUOTE_NO_RE = re.compile(r"(QT\d{6}|RQ\d{4,}-\d+)", re.IGNORECASE)


# ================================================================
# 3. HEADER FIELD EXTRACTION
# ================================================================
def extract_quote_number(page_text: str) -> Optional[str]:
    """Find QT000171 / RQ7289-36 etc."""
    m = QUOTE_NO_RE.search(page_text)
    return m.group(1).upper() if m else None


def extract_quote_date(page_text: str) -> Optional[str]:
    """Example: Nov 21, 2025"""
    m = DATE_PATTERN.search(page_text)
    return m.group(0) if m else None


def _fallback_customer(page_text: str) -> Optional[str]:
    """
    Generic: Customer No / Customer # / Customer Number patterns.
    """
    mc = re.search(
        r"Customer\s+No\.?\s*:?[\s#]*([0-9A-Z\-]+)", page_text, re.IGNORECASE
    )
    if mc:
        return mc.group(1)
    return None


def extract_customer_and_salesperson(
    page_text: str, quote_no: Optional[str]
) -> Tuple[Optional[str], Optional[str]]:
    """
    Map to:
      - Customer Number/ID
      - ReferralManagerCode (salesperson code)

    Tuned to match examples like:
      QT000171:  "11007-4 JZ UPSPPA NET30 QT000171"
      QT569MR25: "Brock Beehler 2026-1 MR BRAUN NET30 QT569MR25"
      QT00040347:"3009-1 11 UPSPPA CC QT00040347"
    """

    cust = None
    sp = None

    if not quote_no:
        return _fallback_customer(page_text), None

    if quote_no.startswith("RQ"):
        # Repair quote RFQ line pattern (kept for future RQ usage)
        m = re.search(
            r"RFQ\s+\S+\s+([0-9A-Z\-]+)\s+([A-Z0-9]{1,3})\s+[A-Z]", page_text
        )
        if m:
            cust, sp = m.group(1), m.group(2)
    else:
        # QT header with optional salesperson name at start:
        #   [Name ...] <cust> <sp> <shipvia> <terms> <quote>
        # e.g. "Brock Beehler 2026-1 MR BRAUN NET30 QT569MR25"
        pattern = (
            r"\n"                                   # start of line
            r"(?:(?:[A-Za-z]+\s+){0,3})?"           # optional 0–3 name tokens
            r"([0-9A-Z\-]+)\s+"                     # group 1: customer no
            r"([A-Z0-9]{1,3})\s+"                   # group 2: salesperson code
            r"[A-Z0-9]+\s+"                         # ship via
            r"[A-Z0-9]+\s+"                         # terms
            + re.escape(quote_no)                   # quote number
        )
        m = re.search(pattern, page_text)
        if m:
            cust, sp = m.group(1), m.group(2)

    # Fallback if customer still missing
    if not cust:
        cust = _fallback_customer(page_text)

    return cust, sp


def extract_company_block(page_text: str):
    """
    Map to:
      - Company
      - Address
      - City
      - State
      - ZipCode
      - Country

    Based on 'Ship To' block, tuned for quotes like:
      - Kenworth- Paccar Company / Sainte Therese, QC J7E4K9 / Canada
      - Braun Corp / 631 West 11th St / Winamac, IN 46996
      - Tsugami America, etc.
    """

    lower = page_text.lower()
    idx = lower.find("ship to")
    if idx == -1:
        return None, None, None, None, None, None

    tail = page_text[idx:]
    lines = tail.splitlines()[1:12]  # skip 'Ship To' line
    lines = [ln.strip() for ln in lines if ln.strip()]
    if not lines:
        return None, None, None, None, None, None

    # Country line
    country = None
    country_line = None
    for ln in reversed(lines):
        if "Canada" in ln:
            country = "Canada"
            country_line = ln
            break
        if "USA" in ln or "United States" in ln:
            country = "USA"
            country_line = ln
            break

    # City/State/ZIP line: look for comma + 2-letter state/province
    city = state = zipcode = None
    cs_line = None
    for ln in lines:
        if "," in ln and re.search(r"\b[A-Z]{2}\b", ln):
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

    # Address: street with digits, above city/state line
    address = None
    if cs_line:
        idx_cs = lines.index(cs_line)
        for ln in reversed(lines[:idx_cs]):
            if any(ch.isdigit() for ch in ln):
                address = ln
                break

    # Company: choose last plausible company-style line
    company_candidates = []
    for ln in lines:
        if ln in (cs_line, country_line, address):
            continue
        if ln.upper().startswith("ATTN:"):
            continue
        if "@" in ln:
            continue
        if "Customers ONLY" in ln:
            continue
        if "Counter Sales" in ln:
            continue
        company_candidates.append(ln)

    company = company_candidates[-1] if company_candidates else None

    # Country defaults
    if country is None and state in ("QC", "ON"):
        country = "Canada"
    if country is None:
        country = "USA"

    return company, address, city, state, zipcode, country


# ================================================================
# 4. LINE ITEM PARSING
# ================================================================
def parse_line_item(line: str) -> Optional[Dict]:
    """
    Map to:
      - Quantity
      - item_id
      - item_desc
      - UOM
      - Unit Price
      - TotalSales

    Assumes lines like:
      "2 PARTS & MISC ALCJA-13ST AlcornTCBoltTool ... 21,775.00 43,550.00"
      "1 MISC 273711-0B07 ATE400,60w,Configured-0B07 2,946.00 2,946.00"
      "2 MISC 158012-5013 ASSY ATE400 48INX30IN 2,935.00 5,870.00"
    """

    s = line.strip()
    if not s:
        return None

    tokens = s.split()
    money_idxs = [i for i, t in enumerate(tokens) if MONEY_RE.match(t)]
    if len(money_idxs) < 2:
        return None

    i2 = money_idxs[-1]   # total
    i1 = money_idxs[-2]   # unit

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

    # body tokens between qty-block and unit price
    start_body = qty_idxs[-1] + 1
    end_body = i1 - 1
    if start_body > end_body:
        return None
    body_tokens = tokens[start_body:end_body + 1]

    # UOM between unit price and total
    uom_tokens = tokens[i1 + 1:i2]
    uom = " ".join(uom_tokens) if uom_tokens else None

    # Clean weird leading "0"
    while body_tokens and body_tokens[0] == "0":
        body_tokens = body_tokens[1:]

    if not body_tokens:
        return None

    # Decide item_id:
    # 1) Prefer first token with a digit or '-' (skips PARTS, MISC, etc.)
    # 2) Fallback: first token not in stopwords
    stopwords = {"PARTS", "&", "MISC", "PARTS&MISC"}
    idx_item = None
    for idx, tok in enumerate(body_tokens):
        if re.search(r"[0-9\-]", tok):
            idx_item = idx
            break
    if idx_item is None:
        for idx, tok in enumerate(body_tokens):
            if tok.upper() not in stopwords:
                idx_item = idx
                break

    if idx_item is None:
        return None

    item_id = body_tokens[idx_item]
    desc_tokens = body_tokens[idx_item + 1 :]
    item_desc = " ".join(desc_tokens).strip()

    return {
        "Quantity": qty,
        "item_id": item_id,
        "item_desc": item_desc,
        "UOM": uom,
        "UnitPrice": unit_price,
        "TotalSales": total_sales,
        "raw": s,
    }


# ================================================================
# 5. PROCESS A SINGLE PDF INTO ROWS
# ================================================================
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

        # Detect / refresh quote header on this page
        q_no = extract_quote_number(text)
        if q_no:
            current_quote_no = q_no
            current_quote_date = extract_quote_date(text)
            current_cust_no, current_sp = extract_customer_and_salesperson(
                text, current_quote_no
            )
            (
                current_company,
                current_address,
                current_city,
                current_state,
                current_zip,
                current_country,
            ) = extract_company_block(text)

        if not current_quote_no:
            continue  # skip pages before we know the quote #

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

            # ---- Header fields ----
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

            # ---- Line item fields ----
            row["item_id"] = parsed["item_id"]
            row["item_desc"] = parsed["item_desc"]
            row["UOM"] = parsed["UOM"]
            row["Quantity"] = parsed["Quantity"]
            row["Unit Price"] = parsed["UnitPrice"]
            row["TotalSales"] = parsed["TotalSales"]

            row["PDF"] = os.path.basename(pdf_path)

            rows.append(row)

    return rows


# ================================================================
# 6. MAIN: PROCESS ALL PDFS IN input_pdfs/
# ================================================================
def main():
    input_folder = "input_pdfs"   # put QT000171.pdf, QT569MR25.pdf, QT00040347.pdf etc. here
    output_folder = "output"
    os.makedirs(output_folder, exist_ok=True)

    pdf_files = sorted(glob.glob(os.path.join(input_folder, "*.pdf")))
    if not pdf_files:
        print(f"[!] No PDFs found in {input_folder}/")
        return

    all_rows: List[Dict] = []

    for pdf_path in pdf_files:
        print(f"[+] Processing: {os.path.basename(pdf_path)}")
        try:
            rows = process_pdf(pdf_path)
            all_rows.extend(rows)
        except Exception as e:
            print(f"    ERROR on {pdf_path}: {e}")

    if not all_rows:
        print("[!] No line items found in any PDF.")
        return

    df = pd.DataFrame(all_rows, columns=HEADER_COLUMNS)
    out_path = os.path.join(output_folder, "Alcorn_Quotes_Merged.xlsx")
    df.to_excel(out_path, index=False)
    print(f"[✓] Done. Wrote {len(df)} rows to {out_path}")


if __name__ == "__main__":
    main()
