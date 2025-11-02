import re
import pandas as pd

PATRONES_EXCLUIR = [
    r"FECHA\s*INGRESO", r"FECHA\s*SALIDA", r"FECHA\s*INFORME",
    r"\bOT\b", r"O\.T\.", r"ORDEN\s+DE\s+TRABAJO", r"COTIZACI[ÓO]N",
    r"CLIENTE", r"ACTIVIDAD", r"INGENIERO", r"AUTOBOMBA",
    r"PREVENTIVO", r"EMERGENCIA", r"ACT\.?PROGRAMA",
    r"REALIZADO POR", r"REVISADO POR",
    r"INFORME\s+T[ÉE]CNICO", r"MANTENIMIENTO\s+EQUIPOS",
    r"G\.C\.A", r"GCA-DOC",
    r"\b\d{10,}\b",
    r"^[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+$",    r"(REALIZADO|REVISADO)\s+POR.*"
]

def extraer_descripcion(texto_filas):
    descripcion = []
    for fila in texto_filas:
        if not fila.strip():
            continue
        if any(re.search(p, fila, re.IGNORECASE) for p in PATRONES_EXCLUIR):
            continue
        descripcion.append(fila.strip())
    return " ".join(descripcion).strip() if descripcion else None

def transformar_datos(df):
    if df is None or df.empty:
        return pd.DataFrame(columns=[
            "IDEquipo","Cliente","FechaIngreso","FechaSalida",
            "FechaInforme","Cotizacion","Actividad","OT","Descripcion"
        ])
    df = df.drop_duplicates()
    columnas = ["IDEquipo","Cliente","FechaIngreso","FechaSalida",
                "FechaInforme","Cotizacion","Actividad","OT","Descripcion"]
    for c in columnas:
        if c not in df.columns:
            df[c] = None
    # Orden sugerido
    return df[columnas]
