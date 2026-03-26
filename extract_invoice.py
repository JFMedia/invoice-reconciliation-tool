import os
import json
import re
import pdfplumber
from openai import OpenAI

client = OpenAI()

def read_pdf_text(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    return text

def clean_po_number(value):
    if not value:
        return ""
    return re.sub(r"[^A-Za-z0-9]", "", str(value)).upper()

def clean_code(value):
    if not value:
        return ""
    return re.sub(r"[^A-Za-z0-9]", "", str(value)).upper()

def normalize_extracted_data(data):
    items = data.get("items", [])
    cleaned_items = []

    for item in items:
        cleaned_items.append({
            "vendor_id": clean_code(item.get("vendor_id", "")),
            "sku": clean_code(item.get("sku", "")),
            "description": str(item.get("description", "")).strip(),
            "quantity": float(item.get("quantity", 0) or 0),
            "unit_cost": float(item.get("unit_cost", 0) or 0),
            "line_total": float(item.get("line_total", 0) or 0),
        })

    data["vendor"] = str(data.get("vendor", "")).strip()
    data["invoice_number"] = str(data.get("invoice_number", "")).strip()
    data["po_number"] = str(data.get("po_number", "")).strip()
    data["po_number_clean"] = clean_po_number(data.get("po_number_clean") or data.get("po_number", ""))
    data["items"] = cleaned_items

    return data

def extract_invoice_data(pdf_path):
    invoice_text = read_pdf_text(pdf_path)

    prompt = f"""
Extract structured line-item data from this invoice.

Return ONLY valid JSON in exactly this structure:

{{
  "vendor": "",
  "invoice_number": "",
  "po_number": "",
  "po_number_clean": "",
  "items": [
    {{
      "vendor_id": "",
      "sku": "",
      "description": "",
      "quantity": 0,
      "unit_cost": 0,
      "line_total": 0
    }}
  ]
}}

Important rules:
- Extract EVERY invoice line item. Do not skip any line items.
- Do not merge separate line items together.
- Each invoice line should become one item in the items array.
- vendor_id = vendor item number / supplier part number / item number if shown
- sku = manufacturer SKU / MFG SKU / manufacturer part number if shown
- If only one product code is present, put it in vendor_id and leave sku blank unless it is clearly labeled as a manufacturer SKU.
- description = full item description for that line
- quantity must be numeric
- unit_cost must be numeric
- line_total must be numeric
- po_number_clean should be the PO number with spaces, dashes, and special characters removed
- If a value is missing, return empty string or 0
- Do NOT invent rows
- Do NOT combine rows
- Make sure all visible item rows from the invoice table are included

Invoice text:
{invoice_text}
"""

    response = client.chat.completions.create(
        model="gpt-5",
        messages=[
            {"role": "system", "content": "You extract structured invoice data and return only JSON."},
            {"role": "user", "content": prompt}
        ]
    )

    content = response.choices[0].message.content
    data = json.loads(content)
    data = normalize_extracted_data(data)

    if not data["items"]:
        raise ValueError("No invoice items were extracted from the PDF.")

    print("EXTRACTED INVOICE DATA:")
    print(json.dumps(data, indent=2))

    return data

if __name__ == "__main__":
    result = extract_invoice_data("../invoices/sample_invoice.pdf")
    print(json.dumps(result, indent=2))