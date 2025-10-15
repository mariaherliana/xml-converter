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
    Given extracted PDF text (single string), attempt to parse required fields.
    This is tuned for the invoice layout you provided. If not found, leave blank.
    Fields to extract: buyer name, buyer address (multi-line capture), NPWP, issue date, invoice no, subtotal (DPP), VAT, total amount
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

    # Normalize spaces
    t = re.sub(r"\r", "\n", text)
    t = re.sub(r"\n\s+\n", "\n\n", t)
    # try to find "To :" block or buyer name uppercase line near top
    # Buyer name shown in screenshot (bold uppercase) - attempt to find line starting with uppercase words > 3 letters
    lines = [l.strip() for l in t.splitlines() if l.strip()]
    # find "To" or "To :" and then next lines
    for i, ln in enumerate(lines):
        if re.search(r"^To\s*:?", ln, re.IGNORECASE):
            # Next non-empty lines might include buyer name and address
            if i+1 < len(lines):
                out["buyer_name"] = lines[i+1]
                # capture next few lines as address until we hit a numeric line containing NPWP or phone or "NPWP/NIK"
                addr_lines = []
                for j in range(i+2, min(i+8, len(lines))):
                    l2 = lines[j]
                    if re.search(r"NPWP|NIK|NPWP\/NIK|NPWP\s*:|Invoice|Issue Date|Tanggal", l2, re.IGNORECASE):
                        break
                    addr_lines.append(l2)
                out["buyer_address"] = " ".join(addr_lines).strip()
            break

    # If still empty buyer name, fallback: first very uppercase line with > 2 words
    if not out["buyer_name"]:
        for ln in lines[:8]:
            if ln.isupper() and len(ln.split()) >= 2:
                out["buyer_name"] = ln
                break

    # NPWP: search for long digit sequences (10-22 digits)
    npwp_match = None
    for ln in lines:
        m = re.search(r"(\d{9,22})", ln.replace(" ", ""))
        if m:
            maybe = m.group(1)
            # filter out phone numbers (10-12 digits) vs NPWP often >= 15? we'll accept >=9 and then pad later
            npwp_match = maybe
            break
    if npwp_match:
        out["npwp"] = npwp_match

    # Issue Date: look for patterns like "Issue Date", "Tanggal Faktur" etc.
    date_patterns = [
        r"Issue Date[:\s]*([0-3]?\d[\/\-][0-1]?\d[\/\-]\d{2,4})",
        r"Tanggal Faktur[:\s]*([0-3]?\d[\/\-][0-1]?\d[\/\-]\d{2,4})",
        r"Issue Date[:\s]*([A-Za-z]{3,9}\s+\d{1,2}\s+\d{4})",
        r"(\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4})",
        r"(\d{4}-\d{2}-\d{2})"
    ]
    for pat in date_patterns:
        m = re.search(pat, t, re.IGNORECASE)
        if m:
            out["issue_date"] = m.group(1).strip()
            break
    # invoice no
    m = re.search(r"Invoice No[:\s]*([A-Za-z0-9\-\_\/]+)", t, re.IGNORECASE)
    if m:
        out["invoice_no"] = m.group(1).strip()
    else:
        # alternative look
        m2 = re.search(r"Invoice\s*#[:\s]*([A-Za-z0-9\-\_\/]+)", t, re.IGNORECASE)
        if m2:
            out["invoice_no"] = m2.group(1).strip()

    # Monetary values: Sub Total / Subtotal / TOTAL AMOUNT / TOTAL
    # Find lines containing 'Sub Total' or 'SubTotal' or 'Sub Total' or the highlighted 'Sub Total' in screenshot
    for ln in lines[::-1]:  # search from bottom up
        if re.search(r"Sub\s*Total|Subtotal|Harga Satuan, DPP", ln, re.IGNORECASE):
            # attempt to extract number from that line
            num = re.search(r"([0-9\.,]+)", ln)
            if num:
                out["subtotal"] = clean_number(num.group(1))
                break
    # fallback: try VAT and Total lines
    for ln in lines[::-1]:
        if re.search(r"\bVAT\b", ln, re.IGNORECASE) and out["vat"] is None:
            m = re.search(r"([0-9\.,]+)", ln)
            if m:
                out["vat"] = clean_number(m.group(1))
        if re.search(r"\bTotal\b", ln, re.IGNORECASE):
            m = re.search(r"([0-9\.,]+)", ln)
            if m:
                out["total"] = clean_number(m.group(1))
                # maybe subtotal still missing
                if out["subtotal"] is None:
                    # attempt to get number from nearby line above
                    # find index and check previous line
                    try:
                        idx = lines.index(ln)
                        if idx > 0:
                            prev = lines[idx-1]
                            nm = re.search(r"([0-9\.,]+)", prev)
                            if nm:
                                out["subtotal"] = clean_number(nm.group(1))
                    except ValueError:
                        pass

    # another heuristic: if subtotal still None, look for first numeric value in the items table region (like unit price qty total)
    if out["subtotal"] is None:
        for ln in lines[::-1]:
            # often a big number like 762.300
            m = re.search(r"([0-9]{1,3}(?:[.,]\d{3})+)", ln)
            if m:
                out["subtotal"] = clean_number(m.group(1))
                break

    # Normalize issue date to yyyy-mm-dd if possible
    if out["issue_date"]:
        try:
            # try parsing common formats
            possible = out["issue_date"].replace(".", "/").replace("-", "/")
            dt = None
            for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y/%m/%d", "%d %b %Y", "%d %B %Y", "%Y-%m-%d"):
                try:
                    dt = datetime.strptime(possible, fmt)
                    break
                except:
                    pass
            if dt:
                out["issue_date"] = dt.strftime("%Y-%m-%d")
        except Exception:
            pass

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
