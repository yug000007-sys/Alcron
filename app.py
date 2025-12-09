import io
import re
from typing import List, Dict, Optional, Tuple

import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader

# ================================================================
# 1. ALCORN HEADER
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

DATE_PATTERN = re.compile(
    r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2},\s+\d{4}"
)
MONEY_RE = re.compile(r"^[0-9,]*\d\.\d{2}$")
QUOTE_NO_RE = re.compile(r"(QT\d{6}|RQ\d{4,}-\d+)", re.IGNORECASE)


# ================================================================
# 2. HELPER FUNCTIONS
# ================================================================
def extract_quote_number(page_text: str) -> Optional[str]:
    m = QUOTE_NO_RE.search(page_text)
    return m.group(1).upper() if m else None


def extract_quote_date(page_text: str) -> Optional[str]:
    m = DATE_PATTERN.search(page_text)
    return m.group(0) if m else None


def _fallback_customer(page_text: str) -> Optional[str]:
    mc = re.search(
        r"Customer\s+No\.?\s*:?[\s#]*([0-9A-Z\-]+)", page_text, re.IGNORECASE
    )
    if mc:
        return mc.group(1)
    return None


def extract_customer_and_salesperson(
    page_text: str, quote_no: Optional[str]
) -> Tuple[Optional[str], Optional[str]]:
    cust = None
    sp = None

    if not quote_no:
        return _fallback_customer(page_text), None

    if quote_no.startswith("RQ"):
        m = re.search(
            r"RFQ\s+\S+\s+([0-9A-Z\-]+)\s+([A-Z0-9]{1,3})\s+[A-Z]", page_text
        )
        if m:
            cust, sp = m.group(1), m.group(2)
    else:
        # handles e.g.:
        #  "11007-4 JZ UPSPPA NET30 QT000171"
        #  "Brock Beehler 2026-1 MR BRAUN NET30 QT569MR25"
        pattern = (
            r"\n"
            r"(?:(?:[A-Za-z]+\s+){0,3})?"
            r"([0-9A-Z\-]+)\s+"
            r"([A-Z0-9]{1,3})\s+"
            r"[A-Z0-9]+\s+"
            r"[A-Z0-9]+\s+"
            + re.escape(quote_no)
        )
        m = re.search(pattern, page_text)
        if m:
            cust, sp = m.group(1), m.group(2)

    if not cust:
        cust = _fallback_customer(page_text)

    return cust, sp


def extract_company_block(page_text: str):
    lower = page_text.lower()
    idx = lower.find("ship to")
    if idx == -1:
        return None, None, None, None, None, None

    tail = page_text[idx:]
    lines = tail.splitlines()[1:12]
    lines = [ln.strip() for ln in lines if ln.strip()]
    if not lines:
        return None, None, None, None, None, None

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

    address = None
    if cs_line:
        idx_cs = lines.index(cs_line)
        for ln in reversed(lines[:idx_cs]):
            if any(ch.isdigit() for ch in ln):
                address = ln
                break

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

    if country is None and state in ("QC", "ON"):
        country = "Canada"
    if country is None:
        country = "USA"

    return company, address, city, state, zipcode, country


def parse_line_item(line: str) -> Optional[Dict]:
    s = line.strip()
    if not s:
        return None

    tokens = s.split()
    money_idxs = [i for i, t in enumerate(tokens) if MONEY_RE.match(t)]
    if len(money_idxs) < 2:
        return None

    i2 = money_idxs[-1]
    i1 = money_idxs[-2]

    try:
        unit_price = float(tokens[i1].replace(",", ""))
        total_sales = float(tokens[i2].replace(",", ""))
    except ValueError:
        return None

    qty_idxs = []
    for i, t in enumerate(tokens):
        if t.isdigit():
            qty_idxs.append(i)
        else:
            break
    if not qty_idxs:
        return None
    qty = int(tokens[qty_idxs[0]])

    start_body = qty_idxs[-1] + 1
    end_body = i1 - 1
    if start_body > end_body:
        return None
    body_tokens = tokens[start_body:end_body + 1]

    uom_tokens = tokens[i1 + 1:i2]
    uom = " ".join(uom_tokens) if uom_tokens else None

    while body_tokens and body_tokens[0] == "0":
        body_tokens = body_tokens[1:]

    if not body_tokens:
        return None

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
    item_desc = " ".join(body_tokens[idx_item + 1 :]).strip()

    return {
        "Quantity": qty,
        "item_id": item_id,
        "item_desc": item_desc,
        "UOM": uom,
        "UnitPrice": unit_price,
        "TotalSales": total_sales,
        "raw": s,
    }


def process_pdf_file(uploaded_file) -> List[Dict]:
    rows: List[Dict] = []
    reader = PdfReader(uploaded_file)
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

            row["PDF"] = uploaded_file.name

            rows.append(row)

    return rows


# ================================================================
# 3. STREAMLIT APP
# ================================================================
def main():
    st.set_page_config(page_title="Alcron Quote Extractor", layout="wide")

    st.title("Alcron Quote PDF → Excel Converter")
    st.write(
        """
        Upload one or more **Alcron quote PDFs** (QT / RQ) and
        download a single Excel file with all line items using your Alcron header.
        """
    )

    uploaded_files = st.file_uploader(
        "Upload PDF quote files",
        type=["pdf"],
        accept_multiple_files=True,
    )

    default_filename = "Alcorn_Quotes_Merged.xlsx"
    output_filename = st.text_input(
        "Output Excel filename",
        value=default_filename,
    )

    if st.button("Process PDFs to Excel"):
        if not uploaded_files:
            st.warning("Please upload at least one PDF file.")
            return

        all_rows: List[Dict] = []

        try:
            with st.spinner("Processing PDFs..."):
                for f in uploaded_files:
                    st.write(f"Processing: **{f.name}**")
                    rows = process_pdf_file(f)
                    all_rows.extend(rows)

            if not all_rows:
                st.error("No line items found in the uploaded PDFs.")
                return

            df = pd.DataFrame(all_rows, columns=HEADER_COLUMNS)

            st.success(f"Extracted {len(df)} rows from {len(uploaded_files)} PDF(s).")
            st.subheader("Preview (first 100 rows)")
            st.dataframe(df.head(100))

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Quotes")
            buffer.seek(0)

            st.download_button(
                label="Download Excel file",
                data=buffer,
                file_name=output_filename or default_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error("An error occurred while processing the PDFs:")
            st.exception(e)


if __name__ == "__main__":
    main()
