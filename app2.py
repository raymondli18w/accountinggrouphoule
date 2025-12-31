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

    if "Client" not in df.columns or df["Client"].iloc[0] != "HE01":
        st.error("❌ Only Client = 'HE01' (Houle Electric Ltd) is allowed.")
        st.stop()

    inv_date = st.date_input("📅 Invoice Date", value=datetime.today())
    due_date = st.date_input("📅 Due Date", value=datetime.today().replace(day=30))

    # Clean data
    df = df[pd.to_numeric(df["Charge Amount"], errors="coerce").notnull()]
    df["Charge Amount"] = pd.to_numeric(df["Charge Amount"])
    df["Charge Qty"] = pd.to_numeric(df["Charge Qty"])
    df = df[df["Charge Amount"] > 0]

    invoice_no = df["Invoice"].iloc[0] if "Invoice" in df.columns else "N/A"

    # Group by Header ref → Billing Ref
    po_groups = {}
    for _, row in df.iterrows():
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

    # Totals
    grand_subtotal = sum(sum(line["Amount"] for line in lines) for docs in po_groups.values() for lines in docs.values())
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

        # === PAGE 1 ===
        y = height - 50

        # LEFT COLUMN (Bill To + Warehouse)
        y_left = height - 70
        y_left = txt(50, y_left, "Bill To:", 10)
        y_left = txt(50, y_left, "Houle Electric Ltd", 10)
        y_left = txt(50, y_left, "5050 North Fraser Way Burnaby BC Canada V5J 0H1", 9)

        y_left = txt(50, y_left, "", 10)
        y_left = txt(50, y_left, "Warehouse", 10)
        y_left = txt(50, y_left, "18 Wheels Meadow Warehouse", 9)
        y_left = txt(50, y_left, "8335 Meadow Ave", 9)
        y_left = txt(50, y_left, "Burnaby, BC V3N 2W1 Canada", 9)

        # RIGHT COLUMN (Terms, Due Date, ID)
        y_right = height - 70
        y_right = txt(350, y_right, "GST Reg No. 749984340", 10)
        y_right = txt(350, y_right, "Net 30 days", 10)
        y_right = txt(350, y_right, f"Due Date: {due_date.strftime('%m/%d/%Y')}", 10)
        y_right = txt(350, y_right, "ID: HE01", 10)

        # Company Header (top left) + Invoice (top right)
        txt(50, height - 50, "18 Wheels Logistics Limited Partnership", 10)
        txt(width - 150, height - 50, "INVOICE", 12)
        txt(width - 150, height - 65, f"Invoice No. {invoice_no}", 10)

        # Payment instructions (below warehouse)
        y_left = txt(50, y_left, "", 9)
        y_left = txt(50, y_left, 'Please make payment to "18 Wheels Supply Chain Ltd."', 9)
        y_left = txt(50, y_left, "Mail payment to 7185 11th Ave, Burnaby, BC V3N 2M5.", 9)
        y_left = txt(50, y_left, "Email: receivable@18wheels.ca", 9)

        # Start line items below y = 450
        y = 450
        page_num = 1

        # Process each PO#
        for po, docs in po_groups.items():
            if y < 150:
                c.showPage()
                page_num += 1
                # Repeat headers on new page
                txt(50, height - 50, "18 Wheels Logistics Limited Partnership", 10)
                txt(width - 150, height - 50, "INVOICE", 12)
                txt(width - 150, height - 65, f"Invoice No. {invoice_no}", 10)
                y = height - 150

            # Service totals for this PO# (left side)
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
                txt(50, y, f"{code:<12} – {desc:<16} – {v['qty']:<2} {unit:<4} – ${v['amt']:,.2f}", 9)
                y -= 14
            y -= 10

            # Line items (centered block style, like your original)
            for doc, lines in docs.items():
                doc_total = sum(l["Amount"] for l in lines)
                y = txt(50, y, f"Document: {doc}", 10)
                for line in lines:
                    # Service row
                    txt(50, y, f"{line['Service Code']:<8}", 9)
                    txt(120, y, f"{line['Description']:<16}", 9)
                    txt(280, y, f"{line['Qty']:<3}", 9)
                    txt(320, y, f"{line['Unit']:<4}", 9)
                    txt(370, y, f"{line['Rate']:<8}", 9)
                    txt(450, y, f"{line['Amount']:<10}", 9)
                    y -= 14

                    # Job/Date/PO row
                    txt(50, y, f"Job#: {doc} Date: {line['Date']} PO#: {po}", 8)
                    y -= 18
                y = txt(400, y, f"Subtotal ({doc}): ${doc_total:,.2f}", 9)
                y -= 20

            y = txt(400, y, f"PO# Total ({po}): ${sum(sum(l['Amount'] for l in lines) for lines in docs.values()):,.2f}", 10)
            y -= 30

        # Totals (right-aligned, bottom of last page)
        if y < 250:
            c.showPage()
            page_num += 1
            txt(50, height - 50, "18 Wheels Logistics Limited Partnership", 10)
            txt(width - 150, height - 50, "INVOICE", 12)
            txt(width - 150, height - 65, f"Invoice No. {invoice_no}", 10)
            y = height - 150

        # Right-aligned totals (mimic your original invoice)
        c.setFont("Helvetica", 10)
        c.drawString(400, y, f"Subtotal: ${grand_subtotal:,.2f}")
        c.drawString(400, y - 15, f"GST (5%): ${gst_amount:,.2f}")
        c.setFont("Helvetica-Bold", 12)
        c.drawString(400, y - 35, f"TOTAL DUE: ${total_due:,.2f}")

        # Footer
        c.setFont("Helvetica", 8)
        c.drawString(50, 50, f"{datetime.now().strftime('%m/%d/%Y %I:%M %p')} Page {page_num} of {page_num}")

        c.save()
        buffer.seek(0)
        return buffer

    pdf = create_pdf()
    st.download_button(
        "📥 Download Invoice PDF",
        data=pdf,
        file_name=f"Houle_Invoice_{invoice_no}.pdf",
        mime="application/pdf"
    )