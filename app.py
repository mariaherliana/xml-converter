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
def parse_date_from_text(text: str) -> str:
    """
    Looks for lines like:
    'KOTA ADM. JAKARTA SELATAN, 30 September 2025'
    Returns DD/MM/YYYY or empty string.
    """
    # common pattern: , <day> <MonthName> <year>
    m = re.search(r",\s*(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", text)
    if m:
        day, mon_name, year = m.group(1), m.group(2), m.group(3)
        try:
            dt = datetime.datetime.strptime(f"{day} {mon_name} {year}", "%d %B %Y")
        except ValueError:
            # try English month
            try:
                dt = datetime.datetime.strptime(f"{day} {mon_name} {year}", "%d %b %Y")
            except Exception:
                return ""
        return dt.strftime("%d/%m/%Y")
    return ""

def parse_kode_seri_type(text: str) -> Dict[str,str]:
    """
    Find 'Kode dan Nomor Seri Faktur Pajak: 0400250031...' -> get first 3 digits
    If starts with 040 => Normal else Pembetulan
    """
    m = re.search(r"Kode dan Nomor Seri Faktur Pajak[:\s]+([0-9A-Za-z\-]+)", text)
    if m:
        code = m.group(1).strip()
        prefix = code[:3]
        if prefix == "040":
            ftype = "Normal"
        else:
            ftype = "Pembetulan"
        return {"raw_code": code, "type": ftype}
    return {"raw_code": "", "type": ""}

def parse_reference(text: str) -> str:
    # look for (Referensi: something) or Referensi: something
    m = re.search(r"Referensi[:\s]*([A-Za-z0-9\-\_\/]+)", text)
    if m:
        return m.group(1).strip()
    m2 = re.search(r"\(Referensi:\s*([^\)]+)\)", text)
    if m2:
        return m2.group(1).strip()
    return ""

def extract_buyer_block(text: str) -> str:
    """Return the textual block for 'Pembeli Barang Kena Pajak' section (approx)."""
    # crude approach: find 'Pembeli Barang Kena Pajak' then take next 400-800 chars
    m = re.search(r"(Pembeli Barang Kena Pajak\/Penerima Jasa Kena Pajak:.*?)(?:\n\n|\Z)", text, re.S)
    if m:
        return m.group(1)
    # fallback: search for 'Pembeli Barang' and take chunk
    m2 = re.search(r"Pembeli Barang.*", text, re.S)
    if m2:
        start = m2.start()
        return text[start:start+800]
    return ""

def parse_buyer_fields(text: str) -> Dict[str,str]:
    b = extract_buyer_block(text)
    result = {"buyer_npwp":"", "buyer_name":"", "buyer_address":"", "buyer_email":"", "buyer_id_tku":""}
    # NPWP : digits
    m = re.search(r"NPWP\s*[:]\s*([0-9]{15,20})", b)
    if m:
        result["buyer_npwp"] = m.group(1).strip()
    # Name
    m = re.search(r"Nama\s*[:]\s*([^\n\r]+)", b)
    if m:
        result["buyer_name"] = m.group(1).strip()
    # Alamat
    m = re.search(r"Alamat\s*[:]\s*([^\n\r\(]+(?:\n[^\n\r]+)?)", b)
    if m:
        addr = m.group(1).strip()
        # remove trailing NPWP if present at same line or following
        result["buyer_address"] = re.sub(r"\s*NPWP.*", "", addr).strip()
    else:
        # try to capture full address by taking up to NPWP line
        m2 = re.search(r"Alamat\s*[:]\s*(.*?)\nNPWP", b, re.S)
        if m2:
            result["buyer_address"] = " ".join([ln.strip() for ln in m2.group(1).splitlines()]).strip()
    # Email
    m = re.search(r"Email[:\s]*([A-Za-z0-9\.\-_]+@[A-Za-z0-9\.\-]+)", b)
    if m:
        result["buyer_email"] = m.group(1).strip()
    # Buyer ID TKU: find '#' followed by digits in address line
    m = re.search(r"#\s*([0-9]{8,30})", b)
    if m:
        result["buyer_id_tku"] = m.group(1).strip()
    return result

def extract_goods_and_dpp(text: str) -> List[Dict[str,Any]]:
    """
    Extract lines under the "Nama Barang Kena Pajak / Jasa Kena Pajak" and the price column.
    This is heuristic: find the table area and parse each item block.
    """
    # We'll search for the table header then grab following lines until summary area like "Dasar Pengenaan Pajak"
    m = re.search(r"Nama Barang Kena Pajak\s*\/\s*Jasa Kena Pajak(.*?)(?:Dasar Pengenaan Pajak|Jumlah PPN)", text, re.S|re.I)
    items = []
    if m:
        block = m.group(1)
        # Items are often numbered: "1 000000\n\nCall fee\nRp 762.300,00 x 1,00"
        # Split by lines that start with a digit number and possibly space
        parts = re.split(r"\n(?=\s*\d+\s+)", block)
        for p in parts:
            # name: first non-empty line
            lines = [ln.strip() for ln in p.splitlines() if ln.strip()]
            if not lines:
                continue
            # find a line that looks like a currency: 'Rp 762.300,00' or ends with numbers and comma
            price = ""
            name = lines[0]
            # try to locate price inside lines
            for ln in lines:
                mprice = re.search(r"Rp\s*([\d\.\,]+)", ln)
                if mprice:
                    price = mprice.group(1).strip()
                    break
            # convert price to plain number
            price_num = None
            if price:
                price_num = price.replace(".", "").replace(",", ".")
                try:
                    price_num = float(price_num)
                except:
                    price_num = None
            items.append({"goods_service": name, "dpp": price_num, "dpp_raw": price})
    return items

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
    page = st.selectbox("", ["Extractor", "XML Converter (placeholder)"])

if page != "XML Converter (placeholder)":
    uploaded = st.file_uploader("Upload Coretax PDF(s)", type=["pdf"], accept_multiple_files=True)
else:
    st.info("XML Converter page — placeholder. We'll implement this later.")
    uploaded = []

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
    st.session_state["results_df"] = None
    st.session_state["last_log_id"] = None
    st.experimental_rerun()

if extract_btn:
    if not uploaded:
        st.warning("Please upload at least one Coretax PDF before extracting.")
    else:
        if not st.session_state.get("confirm_coretax", False):
            st.warning("Please confirm the files are Coretax-format PDFs.")
        else:
            rows = []
            for f in uploaded:
                raw = f.read()
                text = extract_text_from_pdf_bytes(raw)
                date = parse_date_from_text(text)
                kode_info = parse_kode_seri_type(text)
                reference = parse_reference(text)
                buyer = parse_buyer_fields(text)
                goods = extract_goods_and_dpp(text)
                # If goods empty, create fallback single row with blank goods
                if not goods:
                    goods = [{"goods_service":"", "dpp": None, "dpp_raw": ""}]

                # For each goods/service produce a row
                for g in goods:
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
                        "buyer_id_tku": buyer.get("buyer_id_tku"),
                        "goods_service": g.get("goods_service"),
                        "dpp": g.get("dpp"),
                        "dpp_raw": g.get("dpp_raw")
                    }
                    rows.append(row)

            df = pd.DataFrame(rows)
            st.session_state["results_df"] = df

            # Log to Supabase: Processed
            if supabase:
                log_payload = {
                    "processed_count": len(uploaded),
                    "status": "Processed",
                    "details": {"files": [f.name for f in uploaded]}
                }
                try:
                    resp = supabase.table("extraction_logs").insert(log_payload).execute()
                    # store returned id if available
                    try:
                        inserted = resp.data[0]
                        st.session_state["last_log_id"] = inserted.get("id")
                    except Exception:
                        st.session_state["last_log_id"] = None
                except Exception as e:
                    st.error(f"Supabase logging failed: {e}")
            else:
                st.info("Supabase not configured; skipping logging.")

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

