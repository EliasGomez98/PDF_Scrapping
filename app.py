import re
from io import BytesIO
from datetime import datetime

import pandas as pd
import streamlit as st

# pdfminer.six
from pdfminer.high_level import extract_text


# =========================
# STREAMLIT CONFIG (debe ir antes de cualquier otro st.*)
# =========================
st.set_page_config(page_title="Automatización revisión PDFs", layout="wide")
st.title("📄 Automatización de revisión de Expedientes")
st.caption("Sube uno o varios PDFs, aplica Expresiones Regulares y descarga un Excel consolidado.")


CAMPOS = [
    "NUM_POL", "MON", "NUM_DOC", "FEC_NAC", "INI_VIG_POL", "FIN_VIG_POL",
    "PER_DIF", "PER_GAR", "REM_BASE", "PER_PAGO_RENTA",
    "K_SEPELIO", "P_UNICA", "PORC_DEV_PRIMA", "TASA_VENTA"
]

PATRONES = {
    "NUM_POL": r"PÓLIZA\s+N[°º]\s*([A-Z0-9\/\.\-]+)",
    "MON": r"MONTO\s+PRIMA\s+ÚNICA[\s\n]*([A-Z$\/\.]+)",
    "NUM_DOC": r"N[°º][\s\n]*([0-9 ]{8,})",
    "FEC_NAC": r"FECHA\s+DE\s+NACIMIENTO[\s\n]*([0-9 ]{6,})",
    "INI_VIG_POL": r"FECHA(?:\s+DE)?\s+INICIO\s+VIGENCIA\s+(?:DE\s+LA\s+PÓLIZA|DEL\s+PG)[\s\n]*([0-9 ]{6,})",
    "FIN_VIG_POL": r"FECHA(?:\s+DE)?\s+FIN\s+VIGENCIA\s+(?:DE\s+LA\s+PÓLIZA|DEL\s+PG)[\s\n]*([0-9 ]{6,})",
    "PER_DIF": r"DIFERIMIENTO\s+DEL\s+PAGO\s*\(N[°º]\s*DE\s+AÑOS\)[\s\n]*([0-9]{1,3})",
    "PER_GAR": r"N[°º]\s*MESES\s+PERIODO\s+GARANTIZADO\s*\(PG\)[\s\n]*([0-9]{1,3})",
    "REM_BASE": r"MONTO\s+RENTA\s+BASE[\s\S]*?([A-Z$\/\.]+\s*\d[\d,\.]*)",
    "PER_PAGO_RENTA": r"PERIODICIDAD\s+DEL\s+PAGO[\s\n]*([A-ZÁÉÍÓÚ]+)",
    "K_SEPELIO": r"SUMA\s+ASEGURADA\s+COB\.?\s+DE\s+SEPELIO[\s\n]*([A-Z$\/\.]+\s*\d[\d,\.]*)",
    "P_UNICA": r"MONTO\s+PRIMA\s+ÚNICA[\s\n]*([A-Z$\/\.]+\s*\d[\d,\.]*)",
    "PORC_DEV_PRIMA": r"MONTO\s+DE\s+DEVOLUCIÓN\s+DE\s+PRIMA[\s\n]*([0-9]+%?)",
    "TASA_VENTA": r"(?:TASA\s+DE\s+VENTA\s+DE\s+LA\s+PÓLIZA(?:\s*\(TV\))?|TASA\s+DE\s+VENTA\s*\(TV\)\s*DE\s+LA\s+PÓLIZA)[\s\n]*([0-9]+(?:\.[0-9]+)?)\s*%?"
}


def extraer_texto_pdf(uploaded_file) -> str:
    """
    Streamlit UploadedFile -> BytesIO -> pdfminer.extract_text
    (más estable que pasar el UploadedFile directo).
    """
    try:
        data = uploaded_file.getvalue()
        if not data:
            return ""
        bio = BytesIO(data)
        bio.seek(0)
        txt = extract_text(bio)  # pdfminer.six
        return txt or ""
    except Exception:
        return ""


def extraer_campo(texto: str, patron: str) -> str:
    m = re.search(patron, texto, flags=re.MULTILINE)
    if not m:
        return ""
    # Compacta espacios/saltos para que salgan fechas/números limpitos
    return re.sub(r"\s+", "", m.group(1)).strip()


def elegir_engine_excel() -> str:
    """
    Evita que la app no cargue si openpyxl no está instalado.
    """
    try:
        import openpyxl  # noqa: F401
        return "openpyxl"
    except Exception:
        try:
            import xlsxwriter  # noqa: F401
            return "xlsxwriter"
        except Exception:
            # Último intento: que pandas elija (puede fallar si no hay engine)
            return None


with st.sidebar:
    st.header("⚙️ Parámetros")
    to_upper = st.toggle("Convertir texto a MAYÚSCULAS", value=True)
    show_debug = st.toggle("Mostrar texto extraído (debug)", value=False)
    excel_prefix = st.text_input("Prefijo del Excel", value="RentaMAX")


uploaded_files = st.file_uploader(
    "📤 Sube uno o varios archivos PDF",
    type=["pdf"],
    accept_multiple_files=True
)

if not uploaded_files:
    st.info("Sube al menos un PDF para comenzar.")
    st.stop()


if st.button("▶️ Procesar PDFs", type="primary"):
    registros, errores = [], []
    progress = st.progress(0)

    total = len(uploaded_files)

    for idx, file in enumerate(uploaded_files, start=1):
        texto = extraer_texto_pdf(file)

        if not texto.strip():
            errores.append({"ARCHIVO": file.name, "ERROR": "Texto vacío o no extraíble"})
            progress.progress(idx / total)
            continue

        texto_proc = texto.upper() if to_upper else texto

        if show_debug:
            with st.expander(f"Texto extraído: {file.name}"):
                st.text(texto_proc[:20000])

        fila = {"ARCHIVO": file.name}
        for campo in CAMPOS:
            try:
                valor = extraer_campo(texto_proc, PATRONES[campo])
                fila[campo] = valor if valor else "0"
            except Exception as e:
                fila[campo] = "0"
                errores.append({"ARCHIVO": file.name, "ERROR": f"{campo}: {e}"})

        registros.append(fila)
        progress.progress(idx / total)

    df = pd.DataFrame(registros)

    st.success("✅ Procesamiento terminado")
    st.dataframe(df, use_container_width=True)

    if errores:
        st.warning(f"Se registraron {len(errores)} observaciones")
        with st.expander("Ver detalles"):
            st.dataframe(pd.DataFrame(errores), use_container_width=True)

    bio = BytesIO()
    engine = elegir_engine_excel()

    try:
        if engine:
            with pd.ExcelWriter(bio, engine=engine) as writer:
                df.to_excel(writer, index=False, sheet_name="DATA")
        else:
            # sin engine explícito (puede fallar si no hay ninguno instalado)
            with pd.ExcelWriter(bio) as writer:
                df.to_excel(writer, index=False, sheet_name="DATA")
    except Exception as e:
        st.error(
            "No pude generar el Excel (falta un engine). "
            "Instala openpyxl o xlsxwriter en requirements.txt."
        )
        st.exception(e)
        st.stop()

    bio.seek(0)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{excel_prefix}_{timestamp}.xlsx"

    st.download_button(
        "⬇️ Descargar Excel",
        data=bio,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
