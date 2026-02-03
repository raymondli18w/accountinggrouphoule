import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from datetime import datetime
import io
import os

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

    # CRITICAL: Validate new required column exists
    if "Header Reference 2" not in df.columns:
        st.error("❌ Missing required column: 'Header Reference 2'. "
                 "This column is now used for PO grouping per management directive. "
                 "Please verify your Excel file contains this column header exactly.")
        st.stop()
    
    # Optional but recommended: Validate Billing Ref exists
    if "Billing Ref" not in df.columns:
        st.error("❌ Missing required column: 'Billing Ref' (used for document grouping).")
        st.stop()

    # Date inputs side by side
    col1, col2 = st.columns(2)
    with col1:
        inv_date = st.date_input("📅 Invoice Date", value=datetime.today())
    with col2:
        due_date = st.date_input("📅 Due Date", value=datetime.today().replace(day=30))

    # Clean data
    df_clean = df[pd.to_numeric(df["Charge Amount"], errors="coerce").notnull()].copy()
    df_clean["Charge Amount"] = pd.to_numeric(df_clean["Charge Amount"])
    df_clean["Charge Qty"] = pd.to_numeric(df_clean["Charge Qty"])
    df_clean = df_clean[df_clean["Charge Amount"] > 0]

    invoice_no = df_clean["Invoice"].iloc[0] if "Invoice" in df_clean.columns else "N/A"

    # === GROUPING UPDATED: Now uses "Header Reference 2" ===
    po_groups = {}
    for _, row in df_clean.iterrows():
        # PRIMARY CHANGE: Use "Header Reference 2" for outer grouping (PO level)
        po = str(row["Header Reference 2"]) if pd.notna(row["Header Reference 2"]) else "UNSPECIFIED"
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

    # Totals (unchanged)
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

        # === Logo (optional) ===
        logo_path = "logo.png"
        if os.path.exists(logo_path):
            try:
                img = ImageReader(logo_path)
                c.drawImage(img, 50, y - 70, width=120, height=60, preserveAspectRatio=True, anchor='sw')
                y -= 80
            except:
                y -= 20
        else:
            y -= 20

        # === Company Info ===
        y = txt(50, y, "18 Wheels Logistics Limited Partnership", 10)
        y = txt(50, y, "Phone: 604-439-8938", 9)
        y = txt(50, y, "Email: receivable@18wheels.ca", 9)
        y -= 15

        # === Invoice Header ===
        txt(width - 150, y + 25, "INVOICE", 12)
        txt(width - 150, y + 10, f"Invoice No. {invoice_no}", 10)
        txt(width - 150, y - 5, f"Invoice Date: {inv_date.strftime('%m/%d/%Y')}", 9)
        txt(width - 150, y - 20, f"Due Date: {due_date.strftime('%m/%d/%Y')}", 9)
        y -= 30

        # === Bill To ===
        y = txt(50, y, "Bill To:", 10)
        y = txt(50, y, "Houle Electric Ltd", 10)
        y = txt(50, y, "5050 North Fraser Way Burnaby BC Canada V5J 0H1", 9)
        y = txt(50, y, "ID: HE01", 10)
        y -= 20

        # === Warehouse ===
        y = txt(50, y, "Warehouse", 10)
        y = txt(50, y, "18 Wheels Meadow Warehouse", 9)
        y = txt(50, y, "8335 Meadow Ave, Burnaby, BC V3N 2W1 Canada", 9)
        y -= 30

        # === Global Subtotals ===
        y = txt(50, y, "Subtotals by Service", 10)
        y -= 10

        global_service = {}
        for _, row in df_clean.iterrows():
            key = (row["Service Code"], row["Description"], row["Charge Unit"])
            if key not in global_service:
                global_service[key] = {"qty": 0, "amt": 0.0}
            global_service[key]["qty"] += row["Charge Qty"]
            global_service[key]["amt"] += float(row["Charge Amount"])

        for (code, desc, unit), v in global_service.items():
            line = f"{code} – {desc} – {int(v['qty'])} {unit} – ${v['amt']:.2f}"
            y = txt(50, y, line, 9)
            y -= 14
            if y < 150:
                c.showPage()
                page_num += 1
                y = height - 50

        # === PO & Job Breakdown (NOW GROUPED BY HEADER REFERENCE 2) ===
        y -= 20
        for po, docs in po_groups.items():  # 'po' now holds Header Reference 2 values
            if y < 150:
                c.showPage()
                page_num += 1
                y = height - 50
                txt(50, y, "18 Wheels Logistics LP", 10)
                txt(width - 150, y, "INVOICE", 12)
                y -= 80

            # PO Summary (shows Header Reference 2 value)
            po_service = {}
            all_lines = [line for lines in docs.values() for line in lines]
            for line in all_lines:
                key = (line["Service Code"], line["Description"], line["Unit"])
                po_service.setdefault(key, {"qty": 0, "amt": 0})
                po_service[key]["qty"] += line["Qty"]
                po_service[key]["amt"] += line["Amount"]

            # DISPLAYS HEADER REFERENCE 2 VALUE HERE (critical requirement)
            y = txt(50, y, f"PO# {po} Summary", 10)  # Label kept as "PO#" per "rest stay same" requirement
            y -= 5
            for (code, desc, unit), v in po_service.items():
                line = f"  {code} – {desc} – {v['qty']} {unit} – ${v['amt']:.2f}"
                y = txt(50, y, line, 9)
                y -= 12
                if y < 150:
                    c.showPage()
                    page_num += 1
                    y = height - 50
            y -= 10

            # Each Document in a Box
            for doc, lines in docs.items():
                if y < 200:
                    c.showPage()
                    page_num += 1
                    y = height - 50
                    txt(50, y, "18 Wheels Logistics LP", 10)
                    txt(width - 150, y, "INVOICE", 12)
                    y -= 80

                box_top = y
                y = txt(50, y, f"Document: {doc}", 9)
                y -= 8

                # Table header
                c.setFont("Helvetica-Bold", 8)
                c.drawString(50, y, "Code")
                c.drawString(120, y, "Description")
                c.drawString(250, y, "Qty")
                c.drawString(300, y, "Unit")
                c.drawString(350, y, "Rate")
                c.drawString(420, y, "Amount")
                y -= 12
                c.setFont("Helvetica", 9)

                # Table rows
                for line in lines:
                    c.drawString(50, y, str(line["Service Code"]))
                    c.drawString(120, y, str(line["Description"]))
                    c.drawString(250, y, str(line["Qty"]))
                    c.drawString(300, y, str(line["Unit"]))
                    c.drawString(350, y, f"{line['Rate']:.2f}")
                    c.drawString(420, y, f"{line['Amount']:.2f}")
                    y -= 12
                    if y < 100:
                        break

                y -= 5
                c.setFont("Helvetica", 8)
                # CRITICAL: Shows Header Reference 2 value in document footer (as "PO#")
                c.drawString(50, y, f"Job#: {doc} | Date: {lines[0]['Date']} | PO#: {po}")
                y -= 15

                doc_total = sum(l["Amount"] for l in lines)
                c.setFont("Helvetica-Bold", 9)
                c.drawString(400, y, f"Subtotal: ${doc_total:.2f}")
                y -= 25

                # Draw box
                box_height = box_top - y + 20
                c.setStrokeColorRGB(0.9, 0.9, 0.9)
                c.setLineWidth(0.5)
                c.rect(45, y - 10, width - 90, box_height, stroke=1, fill=0)
                c.setStrokeColorRGB(0, 0, 0)

            po_total = sum(sum(l["Amount"] for l in lines) for lines in docs.values())
            c.setFont("Helvetica-Bold", 10)
            c.drawString(400, y, f"PO# Total: ${po_total:.2f}")  # Shows total for this Header Reference 2 group
            y -= 40

        # === Grand Totals ===
        if y < 150:
            c.showPage()
            page_num += 1
            y = height - 100

        c.setFont("Helvetica-Bold", 10)
        c.drawString(400, y, f"Subtotal: ${grand_subtotal:.2f}")
        c.drawString(400, y - 15, f"GST (5%): ${gst_amount:.2f}")
        c.setFont("Helvetica-Bold", 12)
        c.drawString(400, y - 35, f"TOTAL DUE: ${total_due:.2f}")

        # Footer
        c.setFont("Helvetica", 8)
        c.drawString(50, 50, f"Generated on {datetime.now().strftime('%m/%d/%Y %I:%M %p')} | Page {page_num}")

        c.save()
        buffer.seek(0)
        return buffer

    pdf = create_pdf()
    st.download_button(
        "📥 Download Final Invoice PDF",
        data=pdf,
        file_name=f"Houle_Invoice_{invoice_no}.pdf",
        mime="application/pdf"
    )
