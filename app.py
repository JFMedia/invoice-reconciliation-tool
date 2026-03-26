import streamlit as st

password = st.text_input("Enter password", type="password")

if password != st.secrets["APP_PASSWORD"]:
    st.stop()
    
import streamlit.components.v1 as components
import os
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from reconcile import (
    load_all_invoices_for_batch,
    combine_invoice_lines,
    compare_invoice_to_po,
    highlight_excel_report,
)
from load_po import load_po_csv


st.set_page_config(page_title="PO vs Invoice Reconciliation", layout="wide")
st.title("PO vs Invoice Reconciliation")
st.caption("Upload a purchase order and one or more invoice PDFs to automatically detect discrepancies.")


def save_uploaded_files(po_file, invoice_files):
    temp_dir = tempfile.mkdtemp(prefix="recon_")
    temp_path = Path(temp_dir)

    po_path = temp_path / po_file.name
    with open(po_path, "wb") as f:
        f.write(po_file.getbuffer())

    invoice_dir = temp_path / "invoices"
    invoice_dir.mkdir(exist_ok=True)

    for pdf in invoice_files:
        pdf_path = invoice_dir / pdf.name
        with open(pdf_path, "wb") as f:
            f.write(pdf.getbuffer())

    return temp_path, po_path, invoice_dir


def get_clean_po_number(combined_invoice_df):
    po_number = str(combined_invoice_df.iloc[0]["po_number"]).strip()
    po_number = "".join(ch for ch in po_number if ch.isdigit())
    return po_number


def highlight_mismatches(df):
    def to_number(value):
        try:
            if value == "" or pd.isna(value):
                return None
            return float(value)
        except Exception:
            return None

    def style_row(row):
        styles = [""] * len(row)
        cols = list(row.index)

        status = str(row.get("Status", "")).strip()

        po_qty = to_number(row.get("PO Qty"))
        inv_qty = to_number(row.get("Inv Qty"))
        po_cost = to_number(row.get("PO Cost"))
        inv_cost = to_number(row.get("Inv Cost"))

        # Qty mismatch
        if po_qty is not None and inv_qty is not None and po_qty != inv_qty:
            if "PO Qty" in cols:
                styles[cols.index("PO Qty")] = "background-color: #ffcccc"
            if "Inv Qty" in cols:
                styles[cols.index("Inv Qty")] = "background-color: #ccffcc"

        # Cost mismatch
        if po_cost is not None and inv_cost is not None and round(po_cost, 2) != round(inv_cost, 2):
            if "PO Cost" in cols:
                styles[cols.index("PO Cost")] = "background-color: #ffcccc"
            if "Inv Cost" in cols:
                styles[cols.index("Inv Cost")] = "background-color: #ccffcc"

        # Missing from invoice
        if status == "LINE_MISSING_FROM_INVOICE":
            if "PO Qty" in cols:
                styles[cols.index("PO Qty")] = "background-color: #ffcccc"
            if "PO Cost" in cols:
                styles[cols.index("PO Cost")] = "background-color: #ffcccc"

        # Missing from PO
        if status == "LINE_MISSING_FROM_PO":
            if "Inv Qty" in cols:
                styles[cols.index("Inv Qty")] = "background-color: #ffcccc"
            if "Inv Cost" in cols:
                styles[cols.index("Inv Cost")] = "background-color: #ffcccc"

        return styles

    return (
        df.style
        .apply(style_row, axis=1)
        .format(
            {
                "PO Qty": "{:.0f}",
                "Inv Qty": "{:.0f}",
                "Qty Δ": "{:.0f}",
                "PO Cost": "{:.2f}",
                "Inv Cost": "{:.2f}",
                "Cost Δ": "{:.2f}",
            },
            na_rep=""
        )
    )


po_file = st.file_uploader("Upload PO CSV", type=["csv"])
invoice_files = st.file_uploader("Upload Invoice PDFs", type=["pdf"], accept_multiple_files=True)

if st.button("Run", type="primary"):
    if not po_file:
        st.error("Please upload a PO CSV.")
    elif not invoice_files:
        st.error("Please upload at least one invoice PDF.")
    else:
        try:
            with st.spinner("Saving files..."):
                temp_path, po_path, invoice_dir = save_uploaded_files(po_file, invoice_files)

            with st.spinner("Extracting invoice data..."):
                all_invoice_df = load_all_invoices_for_batch(str(invoice_dir))

            if all_invoice_df.empty:
                st.error("No invoice data found in uploaded PDFs.")
                st.stop()

            combined_invoice_df = combine_invoice_lines(all_invoice_df)
            po_number = get_clean_po_number(combined_invoice_df)

            with st.spinner("Loading PO CSV..."):
                po_df = load_po_csv(str(po_path), po_number=po_number)

            with st.spinner("Comparing invoice to PO..."):
                result_df = compare_invoice_to_po(combined_invoice_df, po_df)

            exceptions_df = result_df[result_df["status"] != "MATCH"].copy()

            qty_mismatches = int((result_df["status"] == "QTY_MISMATCH").sum())
            cost_mismatches = int((result_df["status"] == "COST_MISMATCH").sum())
            both_mismatches = int((result_df["status"] == "QTY_AND_COST_MISMATCH").sum())
            missing_from_po = int((result_df["status"] == "LINE_MISSING_FROM_PO").sum())
            missing_from_invoice = int((result_df["status"] == "LINE_MISSING_FROM_INVOICE").sum())

            column_order = [
                "vendor",
                "vendor_id",
                "po_number",
                "invoice_number",
                "sku",
                "description",
                "status",
                "po_qty",
                "invoice_qty",
                "qty_difference",
                "po_unit_cost",
                "invoice_unit_cost",
                "cost_difference",
            ]

            exceptions_df = exceptions_df[[c for c in column_order if c in exceptions_df.columns]]

            for col in ["po_qty", "invoice_qty", "qty_difference"]:
                if col in exceptions_df.columns:
                    exceptions_df[col] = pd.to_numeric(exceptions_df[col], errors="coerce")

            for col in ["po_unit_cost", "invoice_unit_cost", "cost_difference"]:
                if col in exceptions_df.columns:
                    exceptions_df[col] = pd.to_numeric(exceptions_df[col], errors="coerce").round(2)

            st.success("Reconciliation complete")

            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("Qty Mismatch", qty_mismatches)
            m2.metric("Cost Mismatch", cost_mismatches)
            m3.metric("Both", both_mismatches)
            m4.metric("Missing from PO", missing_from_po)
            m5.metric("Missing from Invoice", missing_from_invoice)

            st.subheader("Exceptions")
           
            display_df = exceptions_df.rename(columns={
                "vendor": "Vendor",
                "vendor_id": "Vendor ID",
                "po_number": "PO #",
                "invoice_number": "Invoice #",
                "sku": "SKU",
                "description": "Description",
                "status": "Status",
                "po_qty": "PO Qty",
                "invoice_qty": "Inv Qty",
                "qty_difference": "Qty Δ",
                "po_unit_cost": "PO Cost",
                "invoice_unit_cost": "Inv Cost",
                "cost_difference": "Cost Δ",
            })

            styled_df = highlight_mismatches(display_df)

            html_table = styled_df.to_html(classes="recon-table", index=False)

            table_html = f"""
            <html>
            <head>
            <style>
                body {{
                    font-family: "Inter", "Segoe UI", Arial, sans-serif;
                    margin: 0;
                    padding: 12px;
                    background: #f8fafc;
                    color: #111827;
                }}

                .table-wrap {{
                    border: 1px solid #e5e7eb;
                    border-radius: 14px;
                    overflow: hidden;
                    background: white;
                    box-shadow: 0 1px 3px rgba(0,0,0,0.05);
                }}

                .recon-table {{
                    border-collapse: separate;
                    border-spacing: 0;
                    width: 100%;
                    font-size: 12px;
                    line-height: 1.35;
                }}

                .recon-table thead th {{
                    position: sticky;
                    top: 0;
                    z-index: 2;
                    background: #f9fafb;
                    color: #374151;
                    font-weight: 600;
                    font-size: 11px;
                    letter-spacing: 0.02em;
                    padding: 10px 12px;
                    border-bottom: 1px solid #e5e7eb;
                    text-align: left;
                    white-space: nowrap;
                }}

                .recon-table tbody td {{
                    padding: 10px 12px;
                    border-bottom: 1px solid #f1f5f9;
                    vertical-align: top;
                    background: #ffffff;
                }}

                .recon-table tbody tr:nth-child(even) td {{
                    background: #fcfcfd;
                }}

                .recon-table tbody tr:hover td {{
                    background: #f8fbff;
                }}

                .recon-table tbody tr:last-child td {{
                    border-bottom: none;
                }}

                /* Keep IDs on one line */
                .recon-table td:nth-child(2),
                .recon-table td:nth-child(3),
                .recon-table td:nth-child(4),
                .recon-table td:nth-child(5),
                .recon-table td:nth-child(7) {{
                    white-space: nowrap;
                }}

                /* Description wraps */
                .recon-table td:nth-child(6) {{
                    white-space: normal;
                    min-width: 220px;
                    max-width: 320px;
                }}

                /* Numeric columns */
                .recon-table td:nth-child(8),
                .recon-table td:nth-child(9),
                .recon-table td:nth-child(10),
                .recon-table td:nth-child(11),
                .recon-table td:nth-child(12),
                .recon-table td:nth-child(13) {{
                    text-align: right;
                    white-space: nowrap;
                    font-variant-numeric: tabular-nums;
                }}
            </style>
            </head>
            <body>
                <div class="table-wrap">
                    {html_table}
                </div>
            </body>
            </html>
            """

            components.html(table_html, height=600, scrolling=True)

            csv_data = exceptions_df.to_csv(index=False).encode("utf-8")

            excel_path = os.path.join(temp_path, "discrepancy_report.xlsx")
            exceptions_df.to_excel(excel_path, index=False)
            highlight_excel_report(excel_path)

            with open(excel_path, "rb") as f:
                excel_data = f.read()

            st.download_button(
                "Download CSV",
                data=csv_data,
                file_name="discrepancy_report.csv",
                mime="text/csv",
            )

            st.download_button(
                "Download Excel",
                data=excel_data,
                file_name="discrepancy_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error(f"Error: {e}")