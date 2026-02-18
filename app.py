import re
from io import BytesIO
from datetime import datetime

import pandas as pd
import streamlit as st
from pdfminer.high_level import extract_text


# -------------------------
# UI
# -------------------------
st.set_page_config(page_title="Revisión PDFs", layout="wide")
st.title("📄 Revisión de PDFs (Regex → Excel)")

with st.sidebar:
    st.header("⚙️ Opciones")
    to_upper = st.checkbox("Convertir texto a MAYÚSCULAS", value=True)
    prefix = st.text_input("Prefijo del Excel", value="RentaMAX")


# -------------------------
# Patrones
# -------------------------
PATRONES = {
    "NUM_POL": r"PÓLIZA\s+N[°º]\s*([A-Z0-9\/\.\-]+)",
    "NUM_DOC": r"N[°º][\s\n]*([0-9 ]{8,})",
    "FEC_NAC": r"FECHA\s+DE\s+NACIMIENTO[\s\n]*([0-9 ]{6,})",
    "TASA_VENTA": r"TASA\s+DE\s+VENTA[\s\n]*([0-9]+(?:\.[0-9]+)?)\s*%?"
}


def pdf_to_text(uploaded_file):
    try:
        return extract_text(BytesIO(uploaded_file.getvalue())) or ""
    except Exception:
        return ""


def extract_field(text, pattern):
    match = re.search(pattern, text, flags=re.MULTILINE)
    if not match:
        return "0"
    return re.sub(r"\s+", "", match.group(1))


# -------------------------
# Upload
# -------------------------
files = st.file_uploader("📤 Sube PDFs", type=["pdf"], accept_multiple_files=True)

if not files:
    st.info("Sube al menos un PDF.")
    st.stop()


# -------------------------
# Procesar
# -------------------------
rows = []

for f in files:
    text = pdf_to_text(f)
    if not text.strip():
        rows.append({"ARCHIVO": f.name, "ERROR": "Texto vacío/no extraíble"})
        continue

    if to_upper:
        text = text.upper()

    row = {"ARCHIVO": f.name}
    for field, pattern in PATRONES.items():
        row[field] = extract_field(text, pattern)

    rows.append(row)

df = pd.DataFrame(rows)
st.dataframe(df, use_container_width=True)


# -------------------------
# Descargar Excel (xlsxwriter)
# -------------------------
output = BytesIO()

with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
    df.to_excel(writer, index=False, sheet_name="DATA")

output.seek(0)

filename = f"{prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

st.download_button(
    "⬇️ Descargar Excel",
    data=output,
    file_name=filename,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
