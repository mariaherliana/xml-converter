# app.py (revised extractor logic)
import re
import io
import datetime
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
supabase = create_client(SUPABASE_URL, SUPABASE_KEY) if SUPABASE_URL and SUPABASE_KEY else None

# -----------------------
# UI Config
# -----------------------
st.set_page_config(page_title="Coretax PDF Extractor", layout="wide")

# -----------------------
# Indonesian months map
# -----------------------
MONTHS_ID = {
    "Januari": "01", "Februari": "02", "Maret": "03", "April": "04",
    "Mei": "05", "Juni": "06", "Juli": "07", "Agustus": "08",
    "September": "09", "Oktober": "10", "November": "11", "Desember": "12"
}

# -----------------------
# Extraction functions
# -----------------------
def extract_text_from_pdf_bytes(file_bytes: bytes) -> str:
    text = ""
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            text += page_text + "\n"
    return text


def extract_date(text: str) -> str:
    # Match both with and without comma before date
    m = re.search(r"(?:,|\b)\s*(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", text)
    if not m:
        return ""
    day, month, year = m.groups()
    month_num = MONTHS_ID.get(month.capitalize())
    if not month_num:
        return ""
    return f"{int(day):02d}/{month_num}/{year}"


def extract_facture_type(text: str):
    m = re.search(r"Kode dan Nomor Seri Faktur Pajak\s*:\s*([0-9A-Za-z\-]+)", text)
    if not m:
        return "", ""
    code = m.group(1).strip()
    ftype = "Normal" if code.startswith("040") else "Pembetulan"
    return code, ftype


def extract_reference(text: str):
    # Match Referensi: followed by any text until newline
    m = re.search(r"Referensi\s*:\s*(.+)", text)
    if not m:
        return ""
    ref = m.group(1).strip()
    # clean up possible trailing words like “Tanggal”
    ref = ref.splitlines()[0].strip()
    return ref


def extract_buyer_fields(text: str):
    buyer_block = re.search(
        r"Pembeli Barang Kena Pajak\/Penerima Jasa Kena Pajak(.*?)(?:Nama Barang Kena Pajak|Dasar Pengenaan Pajak)",
        text, re.S)
    if not buyer_block:
        return {}, ""
    block = buyer_block.group(1)
    buyer = {
        "npwp": "",
        "name": "",
        "address": "",
        "email": "",
        "id_tku": ""
    }

    m = re.search(r"NPWP\s*:\s*([0-9\.]+)", block)
    if m:
        buyer["npwp"] = re.sub(r"\D", "", m.group(1))

    m = re.search(r"Nama\s*:\s*(.+)", block)
    if m:
        buyer["name"] = m.group(1).strip()

    # Address can span multiple lines before next label
    m = re.search(r"Alamat\s*:\s*(.*?)\s*(?:NPWP|Email|$)", block, re.S)
    if m:
        address = " ".join(line.strip() for line in m.group(1).splitlines() if line.strip())
        buyer["address"] = address.strip()

    m = re.search(r"Email\s*:\s*([\w\.-]+@[\w\.-]+)", block)
    if m:
        buyer["email"] = m.group(1).strip()

    # Buyer ID TKU after "#"
    m = re.search(r"#\s*(\d{8,30})", block)
    if m:
        buyer["id_tku"] = m.group(1).strip()

    return buyer, block


def extract_goods_and_dpp(text: str):
    """
    Extracts the first occurrence of 'Nama Barang Kena Pajak / Jasa Kena Pajak'
    and its corresponding 'Harga Jual / Penggantian / Uang Muka / Termin'.
    Returns one row only.
    """
    block = re.search(
        r"Nama Barang Kena Pajak\s*/\s*Jasa Kena Pajak(.*?)Jumlah PPN",
        text, re.S)
    if not block:
        return {"goods_service": "", "dpp": None, "dpp_raw": ""}
    part = block.group(1)

    # Find “Uang Muka / Termin” and its numeric value
    m = re.search(r"(Uang Muka\s*/\s*Termin).*?Rp\s*([\d\.\,]+)", part)
    if not m:
        return {"goods_service": "", "dpp": None, "dpp_raw": ""}
    goods = m.group(1).strip()
    dpp_raw = m.group(2).strip()
    dpp_num = float(dpp_raw.replace(".", "").replace(",", ".")) if dpp_raw else None
    return {"goods_service": goods, "dpp": dpp_num, "dpp_raw": dpp_raw}


# -----------------------
# Streamlit UI
# -----------------------
st.title("Coretax PDF Extractor")

uploaded = st.file_uploader("Upload Coretax PDF(s)", type=["pdf"], accept_multiple_files=True)

if st.button("Extract") and uploaded:
    results = []
    for f in uploaded:
        pdf_bytes = f.read()
        text = extract_text_from_pdf_bytes(pdf_bytes)

        date = extract_date(text)
        kode_seri_raw, facture_type = extract_facture_type(text)
        reference = extract_reference(text)
        buyer, _ = extract_buyer_fields(text)
        goods = extract_goods_and_dpp(text)

        row = {
            "source_filename": f.name,
            "date": date,
            "facture_type": facture_type,
            "kode_seri_raw": kode_seri_raw,
            "reference": reference,
            "buyer_npwp": buyer.get("npwp", ""),
            "buyer_name": buyer.get("name", ""),
            "buyer_address": buyer.get("address", ""),
            "buyer_email": buyer.get("email", ""),
            "buyer_id_tku": buyer.get("id_tku", ""),
            "goods_service": goods["goods_service"],
            "dpp": goods["dpp"],
            "dpp_raw": goods["dpp_raw"]
        }
        results.append(row)

        # Log to Supabase
        if supabase:
            try:
                supabase.table("extraction_logs").insert({
                    "processed_count": len(uploaded),
                    "status": "Processed",
                    "details": {"files": [f.name for f in uploaded]}
                }).execute()
            except Exception as e:
                st.warning(f"Supabase log failed: {e}")

    df = pd.DataFrame(results)
    st.dataframe(df)

    csv = df.to_csv(index=False).encode("utf-8")
    st.download_button("Download CSV", csv, "extracted.csv", "text/csv")

else:
    st.info("Upload Coretax-format Faktur Pajak PDFs to begin.")
