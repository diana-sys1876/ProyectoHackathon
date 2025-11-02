import re
import pandas as pd
import pdfplumber
from etl.utils import convertir_fecha, limpiar_cliente
from etl.transform import extraer_descripcion

def _texto_filas_desde_excel(file):
    df_raw = pd.read_excel(file, header=None)
    texto_filas = []
    for _, fila in df_raw.iterrows():
        celdas = [str(v).strip() for v in fila if pd.notna(v) and str(v).strip().lower() != "nan"]
        if celdas:
            texto_filas.append(" ".join(celdas))
    return texto_filas

def _texto_filas_desde_pdf(file):
    with pdfplumber.open(file) as pdf:
        texto_paginas = [p.extract_text() or "" for p in pdf.pages]
    texto_filas = []
    for pagina in texto_paginas:
        for line in pagina.split("\n"):
            line = line.strip()
            if line and line.lower() != "nan":
                texto_filas.append(line)
    return texto_filas

def _campos_comunes(text):
    fecha_ingreso_m = re.search(r"FECHA\s*INGRESO\s*[:\-]?\s*([\d/\-]+)", text, re.IGNORECASE)
    fecha_salida_m  = re.search(r"FECHA\s*SALIDA\s*[:\-]?\s*([\d/\-]+)", text, re.IGNORECASE)
    fecha_informe_m = re.search(r"FECHA\s*INFORME\s*[:\-]?\s*([\d/\-]+)", text, re.IGNORECASE)

    idequipo = re.search(r"(?:ID\s*EQUIPO|EQUIPO\s*ID|EQP)\s*[:\-]?\s*([A-Z]+\s*-?\s*\d+)", text, re.IGNORECASE)
    cotizacion = re.search(r"COTIZACI[ÓO]N\s*[:\-]?\s*([A-Z0-9\-]+)", text, re.IGNORECASE)
    cliente_raw = re.search(r"CLIENTE\s*[:\-]?\s*([A-Z0-9 \.\-]+?)(?=\s+(?:ACTIVIDAD|INGENIERO|$))", text, re.IGNORECASE)
    actividad_raw = re.search(
        r"ACTIVIDAD\s*[:\-]?\s*(.*?)(?=\b(?:INGENIERO|AUTOBOMBA|PREVENTIVO|EMERGENCIA|ACT\.?PROGRAMA|OT\b|O\.T\.|ORDEN\s+DE\s+TRABAJO|COTIZACI[ÓO]N|CLIENTE|$))",
        text, re.IGNORECASE | re.DOTALL
    )
    ot_raw = re.search(r"(?:\bOT\b|O\.T\.|ORDEN\s+DE\s+TRABAJO)\s*[:\-]?\s*([A-Z0-9\-]+)", text, re.IGNORECASE)

    return {
        "IDEquipo": idequipo.group(1).strip() if idequipo else None,
        "Cliente": limpiar_cliente(cliente_raw.group(1) if cliente_raw else None),
        "FechaIngreso": convertir_fecha(fecha_ingreso_m.group(1)) if fecha_ingreso_m else None,
        "FechaSalida": convertir_fecha(fecha_salida_m.group(1)) if fecha_salida_m else None,
        "FechaInforme": convertir_fecha(fecha_informe_m.group(1)) if fecha_informe_m else None,
        "Cotizacion": cotizacion.group(1) if cotizacion else None,
        "Actividad": actividad_raw.group(1).strip() if actividad_raw else None,
        "OT": ot_raw.group(1).strip() if ot_raw else None,
    }

def extraer_excel(file):
    try:
        filas = _texto_filas_desde_excel(file)
        text = " ".join(filas)
        campos = _campos_comunes(text)
        descripcion = extraer_descripcion(filas)
        registros = [{ **campos, "Descripcion": descripcion }]
        return pd.DataFrame(registros)
    except Exception as e:
        from tkinter import messagebox
        messagebox.showerror("Error", f"No se pudo leer el Excel: {e}")
        return pd.DataFrame()

def extraer_pdf(file):
    try:
        filas = _texto_filas_desde_pdf(file)
        text = " ".join(filas)
        campos = _campos_comunes(text)
        descripcion = extraer_descripcion(filas)
        registros = [{ **campos, "Descripcion": descripcion }]
        return pd.DataFrame(registros)
    except Exception as e:
        from tkinter import messagebox
        messagebox.showerror("Error", f"No se pudo leer el PDF: {e}")
        return pd.DataFrame()
