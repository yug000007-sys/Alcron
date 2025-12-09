# 🧾 Alcorn Quote PDF → Excel Extractor

This tool converts **Alcorn Industrial quote PDFs (QT / RQ)** into one clean Excel file using the **official Alcron header**.

It supports:

- QT quotes (e.g., QT000171)
- MR quotes (e.g., QT569MR25)
- Multi-page quotes
- Multiple PDFs at once
- Automatic field extraction:
  - Quote Number
  - Quote Date
  - Customer Number/ID
  - Salesperson / Referral Manager Code
  - Company, Address, City, State, ZipCode, Country
  - Line item ID, Description, UOM, Qty, Unit Price, TotalSales

This repository contains a tuned version of the extractor tested against:

- `QT000171.pdf`
- `QT569MR25.pdf`
- `QT00040347.pdf`

These were mapped against a real Alcron Excel entry to ensure accuracy.

---

## 🚀 Usage

### 1. Install dependencies

```bash
pip install -r requirements.txt
