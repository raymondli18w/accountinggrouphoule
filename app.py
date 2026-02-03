import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from datetime import datetime, timedelta
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

    # === ROBUST COLUMN DETECTION: Accept "Header Reference 2" OR "Header User 2" ===
    col_map = {col.strip().lower().replace(" ", ""): col for col in df.columns}
    
    po_group_col = None
    po_group_col_display = None
    
    if "headerreference2" in col_map:
        po_group_col = col_map["headerreference2"]
        po_group_col_display = "Header Reference 2"
    elif "headeruser2" in col_map:
        po_group_col = col_map["headeruser2"]
        po_group_col_display = "Header User 2"
    
    if po_group_col is None:
        available = "\n- ".join([f"`{col}` → normalized: `{col.strip().lower().replace(' ', '')}`" for col in df.columns])
        st.error(
            "❌ Missing required grouping column. "
            "Your file must contain **either**:\n"
            "- `Header Reference 2`  **OR**\n"
            "- `Header User 2`\n\n"
            f"Available columns in your file:\n- {available}"
        )
        st.stop()
    
    st.success(f"✓ Grouping invoices by: **{po_group_col_display}** (column name in file: `{po_group_col}`)")

    # Validate Billing Ref exists
    if "Billing Ref" not in df.columns:
        st.error("❌ Missing required column: 'Billing Ref' (used for document grouping).")
        st.stop()

    # Date inputs with safe defaults
    col1, col2 = st.columns(2)
    with col1:
        inv_date = st.date_input("📅 Invoice Date", value=datetime.today())
    with col2:
        due_date = st.date_input("📅 Due Date", value=datetime.today() + timedelta(days=30))

    # Clean data
    df_clean = df[pd.to_numeric(df["Charge Amount"], errors="coerce").notnull()].copy()
    df_clean["Charge Amount"] = pd.to_numeric(df_clean["Charge Amount"])
    df_clean["Charge Qty"] = pd.to_numeric(df_clean["Charge Qty"])
    df_clean = df_clean[df_clean["Charge Amount"] > 0]

    if df_clean.empty:
        st.error("❌ No valid charge records found (all amounts zero or non-numeric).")
        st.stop()

    invoice_no = df_clean["Invoice"].iloc[0] if "Invoice" in df_clean.columns else "N/A"

    # === GROUPING: Uses detected column ===
    po_groups = {}
    for _, row in df_clean.iterrows():
        po = str(row[po_group_col]).strip() if pd.notna(row[po_group_col]) else "UNSPECIFIED"
        doc = str(row["Billing Ref"]).strip() if pd.notna(row["Billing Ref"]) else "UNSPECIFIED"
        
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
            "Date": pd.to_datetime(row["Activity Date"]).strftime("%m/%d/%Y") if pd.notna(row["Activity Date"]) else "N/A"
        })

    # Totals
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

        # === PO & Job Breakdown ===
        y -= 20
        for po, docs in po_groups.items():
            if y < 150:
                c.showPage()
                page_num += 1
                y = height - 50
                txt(50, y, "18 Wheels Logistics LP", 10)
                txt(width - 150, y, "INVOICE", 12)
                y -= 80

            # PO Summary
            po_service = {}
            all_lines = [line for lines in docs.values() for line in lines]
            for line in all_lines:
                key = (line["Service Code"], line["Description"], line["Unit"])
                po_service.setdefault(key, {"qty": 0, "amt": 0})
                po_service[key]["qty"] += line["Qty"]
                po_service[key]["amt"] += line["Amount"]

            y = txt(50, y, f"PO# {po} Summary", 10)
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

                # === FIXED: Description line wrapping logic ===
                # Helper function to split description into 35-char lines
                def wrap_text(text, max_chars=35):
                    if len(text) <= max_chars:
                        return [text]
                    
                    # Split at word boundaries when possible
                    words = text.split()
                    lines = []
                    current_line = ""
                    
                    for word in words:
                        if len(current_line) + len(word) + 1 <= max_chars:
                            if current_line:
                                current_line += " " + word
                            else:
                                current_line = word
                        else:
                            lines.append(current_line)
                            current_line = word
                    if current_line:
                        lines.append(current_line)
                    
                    # If still too long (single long word), force split
                    if len(lines) == 1 and len(lines[0]) > max_chars:
                        lines = [text[i:i+max_chars] for i in range(0, len(text), max_chars)]
                    
                    return lines

                # Draw table rows with wrapped descriptions
                for line in lines:
                    # Get wrapped description lines
                    desc_lines = wrap_text(line["Description"], 35)
                    num_lines = len(desc_lines)
                    
                    # Draw non-description fields (only once)
                    c.drawString(50, y, str(line["Service Code"]))
                    c.drawString(250, y, str(line["Qty"]))
                    c.drawString(300, y, str(line["Unit"]))
                    c.drawString(350, y, f"{line['Rate']:.2f}")
                    c.drawString(420, y, f"{line['Amount']:.2f}")
                    
                    # Draw each line of description
                    for i, desc_line in enumerate(desc_lines):
                        c.drawString(120, y - (i * 12), desc_line)
                    
                    # Adjust y position based on description lines
                    y -= (num_lines * 12)
                    if y < 100:
                        break

                y -= 5
                c.setFont("Helvetica", 8)
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
            c.drawString(400, y, f"PO# Total: ${po_total:.2f}")
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

    try:
        pdf = create_pdf()
        st.download_button(
            "📥 Download Final Invoice PDF",
            data=pdf,
            file_name=f"Houle_Invoice_{invoice_no}.pdf",
            mime="application/pdf"
        )
    except Exception as e:
        st.error(f"❌ PDF generation failed: {str(e)}")
        st.exception(e)
