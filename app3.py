import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from datetime import datetime
import io

st.set_page_config(page_title="Houle Invoice Generator", layout="wide")
st.title("🧾 Houle Electric Invoice Generator (HE01)")

uploaded_file = st.file_uploader("📤 Upload Charges Excel", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"❌ Failed to read Excel: {e}")
        st.stop()

    # Validate client is HE01
    if "Client" not in df.columns or df["Client"].iloc[0] != "HE01":
        st.error("❌ Only Client = 'HE01' (Houle Electric Ltd) is allowed.")
        st.stop()

    inv_date = st.date_input("📅 Invoice Date", value=datetime.today())
    due_date = st.date_input("📅 Due Date", value=datetime.today().replace(day=30))

    # Clean data: keep only rows with Charge Amount > 0
    df_clean = df[pd.to_numeric(df["Charge Amount"], errors="coerce").notnull()].copy()
    df_clean["Charge Amount"] = pd.to_numeric(df_clean["Charge Amount"])
    df_clean["Charge Qty"] = pd.to_numeric(df_clean["Charge Qty"])
    df_clean = df_clean[df_clean["Charge Amount"] > 0]

    invoice_no = df_clean["Invoice"].iloc[0] if "Invoice" in df_clean.columns else "N/A"

    # Group by Header ref → Billing Ref for breakdown
    po_groups = {}
    for _, row in df_clean.iterrows():
        po = str(row["Header ref"]) if pd.notna(row["Header ref"]) else "UNSPECIFIED"
        doc = str(row["Billing Ref"]) if pd.notna(row["Billing Ref"]) else "UNSPECIFIED"
        if po not in po_groups:
            po_groups[po] = {}
        if doc not in po_groups[po]:
            po_groups[po][doc] = []
        po_groups[po][doc].append({
            "Service Code": row["Service Code"],
            "Description": row["Description"],
            "Qty": row["Charge Qty"],
            "Unit": row["Charge Unit"],
            "Rate": row["Rate"],
            "Amount": row["Charge Amount"],
            "Date": pd.to_datetime(row["Activity Date"]).strftime("%m/%d/%Y")
        })

    # Global totals (entire invoice)
    grand_subtotal = df_clean["Charge Amount"].sum()
    gst_rate = 0.05
    gst_amount = grand_subtotal * gst_rate
    total_due = grand_subtotal + gst_amount

    def create_pdf():
        buffer = io.BytesIO()
        c = canvas.Canvas(buffer, pagesize=letter)
        width, height = letter

        def txt(x, y, text, size=10, font="Helvetica"):
            c.setFont(font, size)
            c.drawString(x, y, str(text))
            return y - (size + 4)

        y = height - 50
        page_num = 1

        # === BLOCK 1: Company Header ===
        txt(50, y, "18 Wheels Logistics Limited Partnership", 10)
        txt(width - 150, y, "INVOICE", 12)
        y = txt(width - 150, y, f"Invoice No. {invoice_no}", 10)

        # === BLOCK 2: Bill To + Terms (Left + Right) ===
        y -= 20
        y_left = y
        y_right = y

        y_left = txt(50, y_left, "Bill To:", 10)
        y_left = txt(50, y_left, "Houle Electric Ltd", 10)
        y_left = txt(50, y_left, "5050 North Fraser Way Burnaby BC Canada V5J 0H1", 9)

        y_right = txt(350, y_right, "GST Reg No. 749984340", 10)
        y_right = txt(350, y_right, "Net 30 days", 10)
        y_right = txt(350, y_right, f"Due Date: {due_date.strftime('%m/%d/%Y')}", 10)
        txt(50, y_right, "ID: HE01", 10)

        # === BLOCK 3: Warehouse & Payment ===
        y = min(y_left, y_right) - 20
        y = txt(50, y, "Warehouse", 10)
        y = txt(50, y, "18 Wheels Meadow Warehouse", 9)
        y = txt(50, y, "8335 Meadow Ave", 9)
        y = txt(50, y, "Burnaby, BC V3N 2W1 Canada", 9)
        y -= 10
        y = txt(50, y, 'Please make payment to "18 Wheels Supply Chain Ltd."', 9)
        y = txt(50, y, "Mail payment to 7185 11th Ave, Burnaby, BC V3N 2M5.", 9)
        y = txt(50, y, "Email: receivable@18wheels.ca", 9)

        # === BLOCK 4: GLOBAL SUBTOTALS BY SERVICE (entire invoice) ===
        y -= 20
        y = txt(50, y, "Subtotals by Service", 10)
        y -= 10

        # Aggregate global service totals
        global_service = {}
        for _, row in df_clean.iterrows():
            key = (row["Service Code"], row["Description"], row["Charge Unit"])
            if key not in global_service:
                global_service[key] = {"qty": 0, "amt": 0.0}
            global_service[key]["qty"] += row["Charge Qty"]
            global_service[key]["amt"] += float(row["Charge Amount"])

        for (code, desc, unit), v in global_service.items():
            line = f"{code:<12} – {desc:<16} – {int(v['qty']):<2} {unit:<4} – ${v['amt']:,.2f}"
            y = txt(50, y, line, 9)
            y -= 12

        # === BLOCK 5: PO# BREAKDOWN (Header ref → Billing Ref) ===
        y -= 20
        for po, docs in po_groups.items():
            if y < 150:
                c.showPage()
                page_num += 1
                # Repeat header on new page
                txt(50, height - 50, "18 Wheels Logistics Limited Partnership", 10)
                txt(width - 150, height - 50, "INVOICE", 12)
                txt(width - 150, height - 65, f"Invoice No. {invoice_no}", 10)
                y = height - 120

            # Per-PO# service summary
            po_service = {}
            all_lines = [line for lines in docs.values() for line in lines]
            for line in all_lines:
                key = (line["Service Code"], line["Description"], line["Unit"])
                if key not in po_service:
                    po_service[key] = {"qty": 0, "amt": 0}
                po_service[key]["qty"] += line["Qty"]
                po_service[key]["amt"] += line["Amount"]

            y = txt(50, y, f"Totals by Charge Code (PO# {po})", 10)
            y -= 5
            for (code, desc, unit), v in po_service.items():
                line = f"{code:<12} – {desc:<16} – {v['qty']:<2} {unit:<4} – ${v['amt']:,.2f}"
                y = txt(50, y, line, 9)
                y -= 12
            y -= 10

            # Document breakdown
            for doc, lines in docs.items():
                doc_total = sum(l["Amount"] for l in lines)
                y = txt(50, y, f"Document: {doc}", 10)
                y -= 5
                for line in lines:
                    txt(50, y, line["Service Code"], 9)
                    txt(120, y, line["Description"], 9)
                    txt(250, y, str(line["Qty"]), 9)
                    txt(300, y, line["Unit"], 9)
                    txt(350, y, f"{line['Rate']:.2f}", 9)
                    txt(420, y, f"{line['Amount']:.2f}", 9)
                    y -= 14
                    y = txt(50, y, f"Job#: {doc} Date: {line['Date']} PO#: {po}", 8)
                    y -= 16
                y = txt(400, y, f"Subtotal ({doc}): ${doc_total:,.2f}", 9)
                y -= 20

            po_total = sum(sum(l["Amount"] for l in lines) for lines in docs.values())
            y = txt(400, y, f"PO# Total ({po}): ${po_total:,.2f}", 10)
            y -= 30

        # === BLOCK 6: GRAND TOTALS ===
        if y < 200:
            c.showPage()
            page_num += 1
            txt(50, height - 50, "18 Wheels Logistics Limited Partnership", 10)
            txt(width - 150, height - 50, "INVOICE", 12)
            txt(width - 150, height - 65, f"Invoice No. {invoice_no}", 10)
            y = height - 150

        y -= 20
        txt(400, y, f"Subtotal: ${grand_subtotal:,.2f}", 10)
        txt(400, y - 15, f"GST (5%): ${gst_amount:,.2f}", 10)
        c.setFont("Helvetica-Bold", 12)
        c.drawString(400, y - 35, f"TOTAL DUE: ${total_due:,.2f}")

        # === FOOTER ===
        c.setFont("Helvetica", 8)
        c.drawString(50, 50, f"{datetime.now().strftime('%m/%d/%Y %I:%M %p')} Page {page_num} of {page_num}")

        c.save()
        buffer.seek(0)
        return buffer

    # Download button
    pdf = create_pdf()
    st.download_button(
        "📥 Download Final Invoice PDF",
        data=pdf,
        file_name=f"Houle_Invoice_{invoice_no}.pdf",
        mime="application/pdf"
    )