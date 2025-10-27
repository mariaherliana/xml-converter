# app.py
import re
import io
import datetime
import base64
from typing import List, Dict, Any

import streamlit as st
import pandas as pd
import pdfplumber
from supabase import create_client
from dotenv import load_dotenv
import os
import xml.etree.ElementTree as ET

load_dotenv()

# -----------------------
# Config / Supabase
# -----------------------
SUPABASE_URL = st.secrets["SUPABASE"]["URL"] if "SUPABASE" in st.secrets else os.getenv("SUPABASE_URL")
SUPABASE_KEY = st.secrets["SUPABASE"]["KEY"] if "SUPABASE" in st.secrets else os.getenv("SUPABASE_KEY")
if SUPABASE_URL and SUPABASE_KEY:
    supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
else:
    supabase = None

# -----------------------
# Palette / Theme colors
# -----------------------
PALETTE = {
    "primary": "#E2A16F",  # warm tan
    "bg_light": "#FFF0DD",
    "muted": "#D1D3D4",
    "accent": "#86B0BD"    # blueish
}

st.set_page_config(page_title="Coretax PDF Extractor", layout="wide",
                   initial_sidebar_state="expanded")

# Minimal theming via CSS using palette
st.markdown(
    f"""
    <style>
    .stApp {{ background-color: {PALETTE['bg_light']} }}
    .header {{ background-color: {PALETTE['primary']}; padding: 10px; border-radius: 8px; color: white }}
    .stButton>button {{ background-color: {PALETTE['accent']}; color: white; border: none }}
    .reset-btn .stButton>button {{ background-color: #999; color: white }}
    table.dataframe tbody tr:hover {{ background-color: {PALETTE['muted']} }}
    </style>
    """, unsafe_allow_html=True)

# -----------------------
# Helper extraction functions
# -----------------------
MONTHS_ID = {
    "Januari": "01", "Februari": "02", "Maret": "03", "April": "04",
    "Mei": "05", "Juni": "06", "Juli": "07", "Agustus": "08",
    "September": "09", "Oktober": "10", "November": "11", "Desember": "12"
}

def parse_date_from_text(text: str) -> str:
    """
    Looks for date formats like:
    'KOTA ADM. JAKARTA SELATAN, 30 September 2025'
    or 'Jakarta 30 September 2025'
    Returns DD/MM/YYYY or empty string.
    """
    m = re.search(r"(?:,|\b)\s*(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", text)
    if not m:
        return ""
    day, mon_name, year = m.group(1), m.group(2), m.group(3)
    month_num = MONTHS_ID.get(mon_name.capitalize())
    if not month_num:
        return ""
    return f"{int(day):02d}/{month_num}/{year}"

def parse_kode_seri_type(text: str) -> Dict[str, str]:
    """
    Find 'Kode dan Nomor Seri Faktur Pajak: 0400250031...' -> get first 3 digits
    If starts with 040 => Normal else Pembetulan
    """
    m = re.search(r"Kode dan Nomor Seri Faktur Pajak\s*:\s*([0-9A-Za-z\-]+)", text)
    if m:
        code = m.group(1).strip()
        prefix = code[:3]
        ftype = "Normal" if prefix == "040" else "Pembetulan"
        return {"raw_code": code, "type": ftype}
    return {"raw_code": "", "type": ""}

def parse_reference(text: str) -> str:
    """
    Captures full reference text after 'Referensi:' until newline and trims trailing ).
    """
    m = re.search(r"Referensi\s*:\s*(.+)", text)
    if not m:
        return ""
    ref_line = m.group(1).strip()
    ref_line = ref_line.splitlines()[0].strip()
    ref_line = ref_line.rstrip(")")
    return ref_line

def extract_buyer_block(text: str) -> str:
    """Return the text block for 'Pembeli Barang Kena Pajak / Penerima Jasa Kena Pajak' section."""
    m = re.search(
        r"Pembeli Barang Kena Pajak\/Penerima Jasa Kena Pajak(.*?)(?:Nama Barang Kena Pajak|Dasar Pengenaan Pajak)",
        text, re.S)
    return m.group(1) if m else ""

def parse_buyer_fields(text: str) -> Dict[str, str]:
    b = extract_buyer_block(text)
    result = {"buyer_npwp": "", "buyer_name": "", "buyer_address": "", "buyer_email": "", "buyer_id_tku": ""}

    m = re.search(r"NPWP\s*:\s*([0-9\.]+)", b)
    if m:
        result["buyer_npwp"] = re.sub(r"\D", "", m.group(1))

    m = re.search(r"Nama\s*:\s*(.+)", b)
    if m:
        result["buyer_name"] = m.group(1).strip()

    m = re.search(r"Alamat\s*:\s*(.*?)\s*(?:NPWP|Email|$)", b, re.S)
    if m:
        address = " ".join(line.strip() for line in m.group(1).splitlines() if line.strip())
        result["buyer_address"] = address.strip()

    m = re.search(r"Email\s*:\s*([\w\.-]+@[\w\.-]+)", b)
    if m:
        result["buyer_email"] = m.group(1).strip()

    m = re.search(r"#\s*(\d{8,30})", b)
    if m:
        result["buyer_id_tku"] = m.group(1).strip()

    return result

def extract_text_from_pdf_bytes(file_bytes: bytes) -> str:
    text = ""
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            text += page_text + "\n"
    return text

# -----------------------
# UI
# -----------------------
st.markdown('<div class="header"><h2>Coretax PDF (Faktur) Extractor</h2></div>', unsafe_allow_html=True)
st.write("Upload one or multiple Coretax-format Faktur Pajak PDFs. Note: app expects Coretax layout.")

# top small controls
col1, col2, col3 = st.columns([2,1,1])
with col1:
    st.markdown("**Actions**")
    st.markdown("- Upload Coretax PDFs\n- Click *Extract* to parse and log to Supabase\n- Download results as CSV/XLSX")
with col2:
    st.markdown("**Download template**")
    st.markdown("[Download Faktur Pajak Keluaran Excel template](https://pajak.go.id/sites/default/files/2025-03/Sample%20Faktur%20PK%20Template%20v.1.4.xml.zip)")
with col3:
    st.markdown("**Pages**")
    page = st.selectbox("", ["Extractor", "XML Converter"])

if page != "XML Converter":
    uploaded = st.file_uploader(
        "Upload Coretax PDF(s)",
        type=["pdf"],
        accept_multiple_files=True,
        key=f"uploader_{st.session_state.get('uploader_key', 0)}"
    )
else:
    st.markdown("### XML Converter")
    st.write("Upload the Coretax Faktur Excel (format like *Sample Faktur PK Template v.1.6.xlsx*) to generate a Faktur Pajak XML file.")

    xlsx_file = st.file_uploader("Upload Faktur Excel", type=["xlsx"], key="xml_converter_uploader")

    if xlsx_file is not None:
        df_faktur = pd.read_excel(xlsx_file, sheet_name="Faktur")
        df_detail = pd.read_excel(xlsx_file, sheet_name="DetailFaktur")

        st.write(f"📄 Loaded {len(df_faktur)} Faktur rows and {len(df_detail)} DetailFaktur rows")

        if st.button("Convert to XML"):
            with st.spinner("Converting XLSX to XML..."):
                # Build XML
                root = ET.Element("TaxInvoiceBulk")
                tin_value = str(df_faktur.get("NPWP Penjual", ["999999999999999"])[0]).strip() if "NPWP Penjual" in df_faktur.columns else "999999999999999"
                ET.SubElement(root, "TIN").text = tin_value
                list_invoices = ET.SubElement(root, "ListOfTaxInvoice")

                def set_elem(parent, tag, val):
                    if pd.notna(val) and str(val).strip() != "":
                        ET.SubElement(parent, tag).text = str(val).strip()
                    else:
                        ET.SubElement(parent, tag).text = ""

                for _, row in df_faktur.iterrows():
                    faktur_elem = ET.SubElement(list_invoices, "TaxInvoice")
                    set_elem(faktur_elem, "TaxInvoiceDate", row.get("Tanggal Faktur", ""))
                    set_elem(faktur_elem, "TaxInvoiceOpt", row.get("Jenis Faktur", "Normal"))
                    set_elem(faktur_elem, "TrxCode", row.get("Kode Transaksi", "07"))
                    set_elem(faktur_elem, "AddInfo", row.get("Referensi", ""))
                    set_elem(faktur_elem, "CustomDoc", row.get("Nomor Seri Faktur", ""))
                    set_elem(faktur_elem, "CustomDocMonthYear", row.get("Masa Pajak", ""))
                    set_elem(faktur_elem, "RefDesc", "")
                    set_elem(faktur_elem, "FacilityStamp", row.get("Kode Dokumen", ""))
                    set_elem(faktur_elem, "SellerIDTKU", row.get("IDTKU Penjual", ""))
                    set_elem(faktur_elem, "BuyerTin", row.get("NPWP Pembeli", ""))
                    set_elem(faktur_elem, "BuyerDocument", "TIN")
                    set_elem(faktur_elem, "BuyerCountry", "IND")
                    set_elem(faktur_elem, "BuyerDocumentNumber", "")
                    set_elem(faktur_elem, "BuyerName", row.get("Nama Pembeli", ""))
                    set_elem(faktur_elem, "BuyerAdress", row.get("Alamat Pembeli", ""))
                    set_elem(faktur_elem, "BuyerEmail", row.get("Email Pembeli", ""))
                    set_elem(faktur_elem, "BuyerIDTKU", row.get("IDTKU Pembeli", ""))

                    list_goods = ET.SubElement(faktur_elem, "ListOfGoodService")
                    baris = row.get("Baris")
                    if pd.notna(baris):
                        detail_rows = df_detail[df_detail["Baris"] == baris]
                        for _, d in detail_rows.iterrows():
                            goods_elem = ET.SubElement(list_goods, "GoodService")
                            set_elem(goods_elem, "Opt", d.get("Opt", "A"))
                            set_elem(goods_elem, "Code", d.get("Kode Barang", "000000"))
                            set_elem(goods_elem, "Name", d.get("Nama Barang", "Barang"))
                            set_elem(goods_elem, "Unit", d.get("Unit", "UM.0001"))
                            set_elem(goods_elem, "Price", d.get("Harga Satuan", "0"))
                            set_elem(goods_elem, "Qty", d.get("Jumlah", "1"))
                            set_elem(goods_elem, "TotalDiscount", d.get("Diskon", "0"))
                            set_elem(goods_elem, "TaxBase", d.get("DPP", "0"))
                            set_elem(goods_elem, "OtherTaxBase", d.get("DPP", "0"))
                            set_elem(goods_elem, "VATRate", d.get("Tarif PPN", "11"))
                            set_elem(goods_elem, "VAT", d.get("PPN", "0"))
                            set_elem(goods_elem, "STLGRate", d.get("Tarif PPnBM", "0"))
                            set_elem(goods_elem, "STLG", d.get("PPnBM", "0"))

                # Build XML bytes
                xml_bytes = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                st.success("✅ Conversion complete!")

                # ---- Supabase log ----
                if supabase:
                    log_payload = {
                        "processed_count": len(df_faktur),
                        "status": "Converted",
                        "details": {"filename": xlsx_file.name}
                    }
                    try:
                        supabase.table("extraction_logs").insert(log_payload).execute()
                        st.caption("Log saved to Supabase.")
                    except Exception as e:
                        st.warning(f"Supabase logging failed: {e}")
                else:
                    st.info("Supabase not configured; skipping logging.")

                st.download_button(
                    "Download XML",
                    data=xml_bytes,
                    file_name="FakturPajak.xml",
                    mime="application/xml"
                )

    else:
        st.info("Please upload a Faktur Excel file to begin conversion.")

st.checkbox("I confirm these are Coretax-format Faktur Pajak PDFs", value=False, key="confirm_coretax")

# Session storage for results
if "results_df" not in st.session_state:
    st.session_state["results_df"] = None
if "last_log_id" not in st.session_state:
    st.session_state["last_log_id"] = None

extract_col, reset_col = st.columns([1,1])
with extract_col:
    extract_btn = st.button("Extract", key="extract_btn")
with reset_col:
    reset_btn = st.button("Reset (clear uploads & results)", key="reset_btn")

if reset_btn:
    # Clear stored results and uploaded files
    st.session_state.clear()
    # Change uploader key to reset the widget
    if "uploader_key" not in st.session_state:
        st.session_state["uploader_key"] = 0
    else:
        st.session_state["uploader_key"] += 1
    st.rerun()

if extract_btn:
    if not uploaded:
        st.warning("Please upload at least one Coretax PDF before extracting.")
    elif not st.session_state.get("confirm_coretax", False):
        st.warning("Please confirm the files are Coretax-format PDFs.")
    else:
        st.session_state["is_extracting"] = True
        with st.spinner("⏳ Extracting data, please wait..."):
            rows = []
            for f in uploaded:
                raw = f.read()
                text = extract_text_from_pdf_bytes(raw)
                date = parse_date_from_text(text)
                kode_info = parse_kode_seri_type(text)
                reference = parse_reference(text)
                buyer = parse_buyer_fields(text)

                row = {
                    "source_filename": f.name,
                    "date": date,
                    "facture_type": kode_info.get("type"),
                    "kode_seri_raw": kode_info.get("raw_code"),
                    "reference": reference,
                    "buyer_npwp": buyer.get("buyer_npwp"),
                    "buyer_name": buyer.get("buyer_name"),
                    "buyer_address": buyer.get("buyer_address"),
                    "buyer_email": buyer.get("buyer_email"),
                    "buyer_id_tku": buyer.get("buyer_id_tku")
                }
                rows.append(row)

            df = pd.DataFrame(rows)
            st.session_state["results_df"] = df

            if supabase:
                log_payload = {
                    "processed_count": len(uploaded),
                    "status": "Processed",
                    "details": {"files": [f.name for f in uploaded]}
                }
                try:
                    resp = supabase.table("extraction_logs").insert(log_payload).execute()
                    try:
                        inserted = resp.data[0]
                        st.session_state["last_log_id"] = inserted.get("id")
                    except Exception:
                        st.session_state["last_log_id"] = None
                except Exception as e:
                    st.error(f"Supabase logging failed: {e}")
            else:
                st.info("Supabase not configured; skipping logging.")
        st.session_state["is_extracting"] = False

# Show results
if st.session_state["results_df"] is not None:
    df = st.session_state["results_df"]
    st.markdown("### Extraction results")
    st.dataframe(df)

    # Downloads
    csv = df.to_csv(index=False).encode("utf-8")
    xlsx_buffer = io.BytesIO()
    with pd.ExcelWriter(xlsx_buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="extraction")
    xlsx_data = xlsx_buffer.getvalue()

    col_a, col_b, col_c = st.columns([1,1,1])
    with col_a:
        st.download_button("Download CSV", data=csv, file_name="extraction.csv", mime="text/csv")
    with col_b:
        st.download_button("Download XLSX", data=xlsx_data, file_name="extraction.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with col_c:
        # Copy as plain text to clipboard is browser-level; provide a textarea for manual copy
        st.text_area("Copy results (select all to copy):", value=df.to_csv(sep="\t", index=False), height=120)

    # If user downloads — we try to update Supabase status to Downloaded.
    # Note: Streamlit's download button cannot give a callback event when click completes.
    # We provide a manual "Mark Downloaded" button to update the log.
    if st.session_state.get("last_log_id"):
        if st.button("Mark as Downloaded (update log)"):
            if supabase:
                try:
                    supabase.table("extraction_logs").update({"status":"Downloaded"}).eq("id", st.session_state["last_log_id"]).execute()
                    st.success("Log updated to Downloaded.")
                except Exception as e:
                    st.error(f"Failed to update log: {e}")
            else:
                st.info("Supabase not configured; cannot update log.")

# Footer / notes
st.markdown("---")
st.markdown("**Notes & assumptions**")
st.markdown("""
- This tool expects Coretax layout PDFs (the invoice you uploaded was used to tune heuristics). Example sample used: the Akasa file you uploaded. :contentReference[oaicite:1]{index=1}  
- Date parsing is heuristic: it looks for patterns like `, 30 September 2025` and normalizes to `DD/MM/YYYY`.  
- Facture type is decided by the first 3 digits of `Kode dan Nomor Seri Faktur Pajak` — `040` => `Normal`, other values => `Pembetulan`.  
- Buyer ID TKU is the numeric string that appears after `#` in the buyer address line.  
- Goods & DPP extraction is table-heuristic-based and may need adjustments for edge cases.  
- If you want automatic update of logs on real download events, we can implement a signed-download route (server) to capture the click; for now the app provides a manual "Mark as Downloaded" button to update Supabase.
""")

