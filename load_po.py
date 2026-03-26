import pandas as pd
import re

def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()

def clean_sku(value):
    text = clean_text(value).upper()
    text = re.sub(r"[^A-Z0-9]", "", text)
    return text

def clean_po_number(value):
    return re.sub(r"[^A-Za-z0-9]", "", clean_text(value)).upper()

def find_column(df, options):
    lower_map = {str(col).strip().lower(): col for col in df.columns}
    for option in options:
        if option.lower() in lower_map:
            return lower_map[option.lower()]
    return None

def load_po_csv(csv_path, po_number=""):
    df = pd.read_csv(csv_path)

    print("RAW CSV COLUMNS:", df.columns.tolist())

    sku_col = find_column(df, [
        "sku",
        "item number",
        "item_no",
        "supplier code",
        "manufact. sku",
        "custom sku"
    ])

    vendor_id_col = find_column(df, [
        "vendor id",
        "vendor_id",
        "vendor code",
        "supplier id",
        "supplier code",
        "account number"
    ])

    desc_col = find_column(df, [
        "description",
        "product",
        "item description",
        "item name",
        "item"
    ])

    qty_col = find_column(df, [
        "qty",
        "quantity",
        "quantity ordered",
        "order qty",
        "order qty."
    ])

    cost_col = find_column(df, [
        "unit_cost",
        "cost",
        "vendor cost",
        "supply price",
        "unit price",
        "unit cost"
    ])

    total_col = find_column(df, [
        "line_total",
        "total",
        "extended cost",
        "total cost"
    ])

    print("Mapped columns:")
    print("sku_col =", sku_col)
    print("vendor_id_col =", vendor_id_col)
    print("desc_col =", desc_col)
    print("qty_col =", qty_col)
    print("cost_col =", cost_col)
    print("total_col =", total_col)

    out = pd.DataFrame()
    out["po_number"] = po_number
    out["po_number_clean"] = clean_po_number(po_number)

    # Prefer Manufact. SKU first, then SKU column, then Vendor ID only if nothing else exists
    sku_series = None

    if "Manufact. SKU" in df.columns:
        sku_series = df["Manufact. SKU"]

        if sku_series.isna().all() and sku_col:
            sku_series = df[sku_col]
        elif vendor_id_col:
            sku_series = sku_series.fillna(df[vendor_id_col])

    elif sku_col:
        sku_series = df[sku_col]

    elif vendor_id_col:
        sku_series = df[vendor_id_col]

    else:
        sku_series = ""

    out["sku"] = sku_series.apply(clean_sku) if hasattr(sku_series, "apply") else ""
    out["vendor_id"] = df[vendor_id_col].apply(clean_sku) if vendor_id_col else ""
    out["description"] = df[desc_col].apply(clean_text) if desc_col else ""
    out["qty"] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0) if qty_col else 0
    out["unit_cost"] = pd.to_numeric(df[cost_col], errors="coerce").fillna(0) if cost_col else 0
    out["line_total"] = pd.to_numeric(df[total_col], errors="coerce").fillna(0) if total_col else 0

    return out

if __name__ == "__main__":
    df = load_po_csv("../po_exports/sample_po.csv", po_number="12345")
    print(df.head(20).to_string())