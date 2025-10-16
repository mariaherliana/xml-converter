# app.py
import streamlit as st
import pdfplumber
import re
import pandas as pd
import io
import os
from datetime import datetime
import xml.etree.ElementTree as ET

# ------------------------
# Config / defaults
# ------------------------
SELLER_ID_TKU_DEFAULT = "0618595813012000000000"  # ID TKU Penjual (22 digits default)
Faktur_columns = [
    "Baris", "Jenis Faktur", "Kode Transaksi", "ID TKU Penjual", "Jenis ID Pembeli",
    "Negara Pembeli", "Email Pembeli", "ID TKU Pembeli", "Nama Pembeli", "Alamat Pembeli",
    "NPWP Pembeli", "Tanggal Faktur", "Invoice No", "Referensi", "Total Amount", "DPP", "PPN"
]
DetailFaktur_columns = [
    "Baris", "Barang/Jasa", "Kode Barang Jasa", "Nama Barang/Jasa", "Nama Satuan Ukur",
    "Harga Satuan", "Jumlah Barang Jasa", "Total Diskon", "DPP", "DPP Nilai Lain",
    "Tarif PPN", "PPN", "Tarif PPnBM", "PPnBM"
]

# ------------------------
# Helpers
# ------------------------
def clean_number(text: str) -> float:
    """
    Convert number string like '762.300' or '830.907' or '762,300.00' into float.
    Handles thousands separator as dot or comma. Assumes decimal separator is either '.' with no thousands or ',' as thousands.
    """
    if text is None:
        return None
    s = str(text).strip()
    # remove currency symbols and spaces
    s = re.sub(r"[^\d,.\-]", "", s)
    if s == "":
        return None
    # If both comma and dot: assume dot is thousands separator when comma present as decimal? we'll normalize:
    # Common Indonesian formatting: "762.300" meaning 762300 (no decimals). If there is a comma as decimals, handle it.
    if s.count(".") > 0 and s.count(",") == 0:
        # remove dots (thousands) -> integer
        s2 = s.replace(".", "")
        try:
            return float(s2)
        except:
            return None
    if s.count(",") > 0 and s.count(".") == 0:
        # maybe "1,234" meaning 1234 -> remove commas
        s2 = s.replace(",", "")
        try:
            return float(s2)
        except:
            return None
    # if both present, treat comma as thousands and dot as decimal OR vice versa — try a couple heuristics
    if s.count(",") > 0 and s.count(".") > 0:
        # heuristic: if last separator is comma, comma is decimal -> replace '.' thousands, ',' decimal
        if s.rfind(",") > s.rfind("."):
            s2 = s.replace(".", "").replace(",", ".")
            try:
                return float(s2)
            except:
                pass
        # else last sep is dot -> remove commas, keep dot as decimal
        s2 = s.replace(",", "")
        try:
            return float(s2)
        except:
            pass
    # fallback
    try:
        return float(s)
    except:
        return None

def pad_npwp_to_22(npwp_raw: str) -> str:
    if not npwp_raw:
        return ""
    digits = re.sub(r"\D", "", npwp_raw)
    # pad left with zeros until length 22
    return digits.zfill(22)

def extract_fields_from_text(text: str) -> dict:
    """
    Improved, line-oriented extractor for your fixed-layout, text-based invoices.
    Returns dict with keys:
      buyer_name, buyer_address, npwp, issue_date, invoice_no, subtotal, vat, total
    """
    out = {
        "buyer_name": "",
        "buyer_address": "",
        "npwp": "",
        "issue_date": "",
        "invoice_no": "",
        "subtotal": None,
        "vat": None,
        "total": None
    }

    # Make sure we preserve lines, remove trailing/leading spaces
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    # unify some known label variants (lowercase copy for searching)
    low_lines = [ln.lower() for ln in lines]

    # --- Helper local functions ---
    def find_line_containing(key_variants):
        """Return first index of a line containing any of the key_variants (case-insensitive)."""
        for idx, l in enumerate(low_lines):
            for k in key_variants:
                if k in l:
                    return idx
        return None

    def extract_after_label(line, label_regex):
        """Try to extract value after label within the same line using label_regex with a capture group."""
        m = re.search(label_regex, line, re.IGNORECASE)
        if m:
            val = m.group(1).strip()
            return val
        return None

    # --- Invoice No ---
    idx = find_line_containing(["invoice no", "invoice no:", "invoice#", "invoice #", "invoice"])
    if idx is not None:
        # try same-line extraction
        val = extract_after_label(lines[idx], r"(?:invoice\s*(?:no|#)\s*[:\-]?\s*)(.+)$")
        if not val:
            # maybe "Invoice" then number on same line separated by spaces
            parts = lines[idx].split()
            if len(parts) > 1:
                val = " ".join(parts[1:])
        if val:
            out["invoice_no"] = val.strip()

    # --- Issue Date / Tanggal Faktur ---
    idx = find_line_containing(["issue date", "tanggal faktur", "issue date:"])
    if idx is not None:
        # try same line first
        val = extract_after_label(lines[idx], r"(?:issue\s*date|tanggal\s*faktur)\s*[:\-]?\s*(.+)$")
        if not val and idx+1 < len(lines):
            # maybe value on next line
            val = lines[idx+1].strip()
        if val:
            # normalize date like '07 Oct 2025' to 'YYYY-MM-DD' if possible
            raw = val
            parsed = None
            for fmt in ("%d %b %Y", "%d %B %Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
                try:
                    parsed = datetime.strptime(raw, fmt)
                    break
                except Exception:
                    pass
            out["issue_date"] = parsed.strftime("%Y-%m-%d") if parsed else raw

    # --- Buyer name and address ---
    # Preferred anchor: a line that starts with "To" or "To :"
    idx = find_line_containing(["to :", "to:", "to "])
    if idx is not None:
        # buyer name may be on same line after "To:" or on next line
        val = extract_after_label(lines[idx], r"to\s*[:\-]?\s*(.+)$")
        if val:
            out["buyer_name"] = val.strip()
            # address follows starting at next line until we hit NPWP, Invoice, or blank
            addr_lines = []
            j = idx + 1
            while j < len(lines):
                l = lines[j]
                if re.search(r"npwp|invoice|issue date|tanggal|total|vat|ppn", l, re.IGNORECASE):
                    break
                addr_lines.append(l)
                j += 1
            out["buyer_address"] = " ".join(addr_lines).strip()
        else:
            # if "To" line only, next line likely buyer name
            if idx+1 < len(lines):
                out["buyer_name"] = lines[idx+1].strip()
                # address from idx+2 until NPWP or label
                addr_lines = []
                j = idx + 2
                while j < len(lines):
                    l = lines[j]
                    if re.search(r"npwp|invoice|issue date|tanggal|total|vat|ppn", l, re.IGNORECASE):
                        break
                    addr_lines.append(l)
                    j += 1
                out["buyer_address"] = " ".join(addr_lines).strip()
    else:
        # fallback: first fully uppercase line that looks like a company name (avoid "INVOICE" etc.)
        for ln in lines[:10]:
            if ln.isupper() and len(ln) > 3 and not re.search(r"invoice|revcomm|to", ln, re.IGNORECASE):
                out["buyer_name"] = ln
                break

    # --- NPWP (must be 16 digits left-padded with zero if shorter) ---
    idx = find_line_containing(["npwp", "npwp/nik", "npwp :"])
    if idx is not None:
        # try same-line capture
        val = extract_after_label(lines[idx], r"npwp(?:\/nik)?\s*[:\-]?\s*([0-9\.\- ]+)")
        if not val and idx+1 < len(lines):
            # maybe next line contains the digits
            val = lines[idx+1].strip()
        if val:
            digits = re.sub(r"\D", "", val)
            if len(digits) < 16:
                digits = digits.zfill(16)
            out["npwp"] = digits
    else:
        # fallback: search for a long digit sequence anywhere (common NPWP patterns)
        for ln in lines:
            m = re.search(r"(\d{8,16})", re.sub(r"\s", "", ln))
            if m:
                digits = m.group(1)
                if len(digits) < 16:
                    digits = digits.zfill(16)
                out["npwp"] = digits
                break

    # --- Monetary amounts: Subtotal/DPP, VAT/PPN, Total ---
    # Search bottom-up for labels because amounts usually placed near bottom
    for i in range(len(lines)-1, -1, -1):
        ln = lines[i]
        low = low_lines[i]
        # subtotal / DPP
        if out["subtotal"] is None and re.search(r"sub\s*total|subtotal|harga satuan, dpp|sub total", low):
            m = re.search(r"([0-9\.,]+)", ln)
            if m:
                out["subtotal"] = clean_number(m.group(1))
                continue
        # VAT / PPN
        if out["vat"] is None and re.search(r"\b(vat|ppn)\b", low):
            m = re.search(r"([0-9\.,]+)", ln)
            if m:
                out["vat"] = clean_number(m.group(1))
                continue
        # total / total amount
        if out["total"] is None and re.search(r"\b(total amount|total)\b", low):
            m = re.search(r"([0-9\.,]+)", ln)
            if m:
                out["total"] = clean_number(m.group(1))
                # try to find subtotal if it's just above
                if out["subtotal"] is None and i-1 >= 0:
                    m2 = re.search(r"([0-9\.,]+)", lines[i-1])
                    if m2:
                        out["subtotal"] = clean_number(m2.group(1))
                continue

    # final fallback heuristics if any are still None
    if out["subtotal"] is None:
        # try to find first large numeric-looking value in the items region
        for ln in lines:
            m = re.search(r"([0-9]{1,3}(?:[.,]\d{3})+)", ln)
            if m:
                out["subtotal"] = clean_number(m.group(1))
                break

    # If issue_date still empty, try to find any dd MMM yyyy anywhere
    if not out["issue_date"]:
        for ln in lines[:20]:
            m = re.search(r"([0-3]?\d\s+[A-Za-z]{3,9}\s+\d{4})", ln)
            if m:
                try:
                    dt = datetime.strptime(m.group(1), "%d %b %Y")
                    out["issue_date"] = dt.strftime("%Y-%m-%d")
                    break
                except Exception:
                    out["issue_date"] = m.group(1)
                    break

    return out

def pdf_to_text(file_bytes: bytes) -> str:
    text = ""
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text += page.extract_text() or ""
            text += "\n"
    return text

def build_excel_and_xml(parsed_rows: list):
    """
    parsed_rows: list of dicts with keys:
      buyer_name, buyer_address, npwp, issue_date, invoice_no, subtotal, vat, total
    Returns bytes for excel and xml
    """
    faktur_rows = []
    detail_rows = []

    for idx, r in enumerate(parsed_rows, start=1):
        # Faktur row mapping
        buyer_tin_padded = pad_npwp_to_22(r.get("npwp", ""))
        faktur_row = {
            "Baris": idx,
            "Jenis Faktur": "Normal",
            "Kode Transaksi": "04",
            "ID TKU Penjual": SELLER_ID_TKU_DEFAULT,
            "Jenis ID Pembeli": "TIN",
            "Negara Pembeli": "IDN",
            "Email Pembeli": "",
            "ID TKU Pembeli": buyer_tin_padded,
            "Nama Pembeli": r.get("buyer_name", ""),
            "Alamat Pembeli": r.get("buyer_address", ""),
            "NPWP Pembeli": r.get("npwp", ""),
            "Tanggal Faktur": r.get("issue_date", ""),
            "Invoice No": r.get("invoice_no", ""),
            "Referensi": "",
            "Total Amount": int(round(r.get("total", 0))) if r.get("total") is not None else "",
            "DPP": int(round(r.get("subtotal", 0))) if r.get("subtotal") is not None else "",
            "PPN": int(round(r.get("vat", 0))) if r.get("vat") is not None else ""
        }
        faktur_rows.append(faktur_row)

        # DetailFaktur fields and calculations:
        dpp = r.get("subtotal") or 0.0
        # DPP Nilai Lain = ROUND((11/12) * DPP)
        dpp_nilai_lain = round((11.0/12.0) * dpp)
        # PPN = ROUND(DPP Nilai Lain * 12%)
        ppn = round(dpp_nilai_lain * 0.12)

        detail_row = {
            "Baris": idx,
            "Barang/Jasa": "B",
            "Kode Barang Jasa": "000000",
            "Nama Barang/Jasa": "MiiTel Subscription",
            "Nama Satuan Ukur": "UM.0033",
            "Harga Satuan": int(round(r.get("subtotal", 0))) if r.get("subtotal") is not None else "",
            "Jumlah Barang Jasa": 1,
            "Total Diskon": 0,
            "DPP": int(round(dpp)),
            "DPP Nilai Lain": int(dpp_nilai_lain),
            "Tarif PPN": 12,
            "PPN": int(ppn),
            "Tarif PPnBM": 0,
            "PPnBM": 0
        }
        detail_rows.append(detail_row)

    # Build Excel in-memory
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df_faktur = pd.DataFrame(faktur_rows, columns=Faktur_columns)
        df_detail = pd.DataFrame(detail_rows, columns=DetailFaktur_columns)
        df_faktur.to_excel(writer, sheet_name="Faktur", index=False)
        df_detail.to_excel(writer, sheet_name="DetailFaktur", index=False)
        writer.save()
    excel_bytes = out.getvalue()

    # Build XML following your TaxInvoiceBulk template structure
    root = ET.Element("TaxInvoiceBulk", {
        "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
        "xsi:noNamespaceSchemaLocation": "TaxInvoice.xsd"
    })
    # Use seller TIN as TIN element (trim/format if necessary)
    tin_elem = ET.SubElement(root, "TIN")
    tin_elem.text = SELLER_ID_TKU_DEFAULT

    list_elem = ET.SubElement(root, "ListOfTaxInvoice")
    for fr, dr in zip(faktur_rows, detail_rows):
        inv = ET.SubElement(list_elem, "TaxInvoice")
        # minimal mapping
        ET.SubElement(inv, "TaxInvoiceDate").text = fr.get("Tanggal Faktur", "")
        ET.SubElement(inv, "TaxInvoiceOpt").text = fr.get("Jenis Faktur", "Normal")
        ET.SubElement(inv, "TrxCode").text = fr.get("Kode Transaksi", "04")
        ET.SubElement(inv, "AddInfo").text = ""
        ET.SubElement(inv, "CustomDoc").text = ""
        ET.SubElement(inv, "RefDesc").text = ""
        ET.SubElement(inv, "FacilityStamp").text = ""
        ET.SubElement(inv, "SellerIDTKU").text = fr.get("ID TKU Penjual")
        ET.SubElement(inv, "BuyerTin").text = fr.get("ID TKU Pembeli")
        ET.SubElement(inv, "BuyerDocument").text = fr.get("Jenis ID Pembeli")
        ET.SubElement(inv, "BuyerCountry").text = fr.get("Negara Pembeli")
        ET.SubElement(inv, "BuyerDocumentNumber").text = fr.get("NPWP Pembeli", "")
        ET.SubElement(inv, "BuyerName").text = fr.get("Nama Pembeli", "")
        ET.SubElement(inv, "BuyerAdress").text = fr.get("Alamat Pembeli", "")
        ET.SubElement(inv, "BuyerEmail").text = fr.get("Email Pembeli", "")
        ET.SubElement(inv, "BuyerIDTKU").text = fr.get("ID TKU Pembeli", "")
        # List of GoodService
        list_gs = ET.SubElement(inv, "ListOfGoodService")
        gs = ET.SubElement(list_gs, "GoodService")
        ET.SubElement(gs, "Opt").text = dr.get("Barang/Jasa", "B")
        ET.SubElement(gs, "Code").text = dr.get("Kode Barang Jasa", "000000")
        ET.SubElement(gs, "Name").text = dr.get("Nama Barang/Jasa", "MiiTel Subscription")
        ET.SubElement(gs, "Unit").text = dr.get("Nama Satuan Ukur", "UM.0033")
        ET.SubElement(gs, "Price").text = str(dr.get("Harga Satuan", ""))
        ET.SubElement(gs, "Qty").text = str(dr.get("Jumlah Barang Jasa", 1))
        ET.SubElement(gs, "TotalDiscount").text = str(dr.get("Total Diskon", 0))
        ET.SubElement(gs, "TaxBase").text = str(dr.get("DPP", ""))
        ET.SubElement(gs, "OtherTaxBase").text = str(dr.get("DPP Nilai Lain", ""))
        ET.SubElement(gs, "VATRate").text = str(dr.get("Tarif PPN", 12))
        ET.SubElement(gs, "VAT").text = str(dr.get("PPN", ""))
        ET.SubElement(gs, "STLGRate").text = str(dr.get("Tarif PPnBM", 0))
        ET.SubElement(gs, "STLG").text = str(dr.get("PPnBM", 0))

    # pretty print xml bytes
    xml_bytes = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    return excel_bytes, xml_bytes

# ------------------------
# Streamlit UI
# ------------------------
st.set_page_config(page_title="Invoice → Excel & XML (bulk)", layout="wide")
st.title("Invoice Bulk Converter — PDF → Excel & XML")
st.info("Upload invoices (PDF) exported using the standard template. The tool will extract fields and produce one Excel and one XML file for download.")

uploaded_files = st.file_uploader("Upload invoice PDFs (multiple allowed)", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    st.write(f"Received {len(uploaded_files)} files — parsing now...")
    parsed = []
    progress = st.progress(0)
    for i, up in enumerate(uploaded_files, start=1):
        try:
            raw = up.read()
            text = pdf_to_text(raw)
            fields = extract_fields_from_text(text)
            # for traceability, attach filename as invoice_no fallback
            if not fields.get("invoice_no"):
                fields["invoice_no"] = os.path.splitext(up.name)[0]
            parsed.append(fields)
            progress.progress(int(i/len(uploaded_files)*100))
        except Exception as e:
            st.warning(f"Failed to parse {up.name}: {e}")
            parsed.append({
                "buyer_name": "",
                "buyer_address": "",
                "npwp": "",
                "issue_date": "",
                "invoice_no": os.path.splitext(up.name)[0],
                "subtotal": None,
                "vat": None,
                "total": None
            })

    st.success("Parsing complete. Preview extracted fields:")
    preview_df = pd.DataFrame(parsed)
    st.dataframe(preview_df)

    if st.button("Generate Excel & XML"):
        with st.spinner("Building Excel and XML..."):
            excel_bytes, xml_bytes = build_excel_and_xml(parsed)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            excel_name = f"faktur_bulk_{ts}.xlsx"
            xml_name = f"taxinvoice_bulk_{ts}.xml"

            st.download_button("⬇️ Download Excel (Faktur + DetailFaktur)", data=excel_bytes, file_name=excel_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.download_button("⬇️ Download XML (TaxInvoiceBulk)", data=xml_bytes, file_name=xml_name, mime="application/xml")
            st.success("Files ready for download. Review the Excel for any missing / mis-parsed fields before submitting elsewhere.")
else:
    st.info("Upload one or more invoice PDFs to begin.")

st.markdown("<small>Note: this extractor assumes identical invoice layout. If some fields are wrongly parsed, open the generated Excel and correct manually — the XML is generated from the Excel-style rows.</small>", unsafe_allow_html=True)
