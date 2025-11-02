import os
import pandas as pd
from tkinter import messagebox
from sqlalchemy import create_engine, text
from config.settings import DB_CONFIG

SALIDA_XLSX = os.path.join(
    os.path.dirname(os.path.dirname(__file__)),
    'data',
    'Datos_GCA_Limpios.xlsx'
)


def _tipos(df):
    """Normaliza tipos de columnas específicas."""
    for col_texto in ["OT", "IDEquipo", "Cliente"]:
        if col_texto in df.columns:
            df[col_texto] = df[col_texto].astype(str).str.strip() 
    if "Cotizacion" in df.columns:
        df["Cotizacion"] = pd.to_numeric(df["Cotizacion"], errors="coerce")
    return df


def guardar_datos(df, output_file=SALIDA_XLSX):
    """Guarda los datos en un archivo Excel, evitando duplicados."""
    try:
        if df is None or df.empty:
            messagebox.showwarning("Aviso", "No hay datos para exportar")
            return
        df = _tipos(df.copy())

        if os.path.exists(output_file):
            try:
                df_existente = pd.read_excel(output_file, dtype=str)
            except Exception:
                df_existente = pd.DataFrame()
            df_comb = pd.concat([df_existente, df], ignore_index=True).drop_duplicates(ignore_index=True)
            if "Cotizacion" in df_comb.columns:
                df_comb["Cotizacion"] = pd.to_numeric(df_comb["Cotizacion"], errors="coerce")
            df_comb.to_excel(output_file, index=False)
            messagebox.showinfo("Éxito", f"Datos actualizados en {output_file}")
        else:
            df.to_excel(output_file, index=False)
            messagebox.showinfo("Éxito", f"Datos guardados en {output_file}")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo guardar el archivo: {e}")


def guardar_en_postgres(df):
    """
    Guarda los datos en PostgreSQL sin borrar la tabla.
    Evita duplicados usando la clave única compuesta en la DB.
    Retorna (insertados, repetidos, total)
    """
    try:
        if df is None or df.empty:
            messagebox.showwarning("Aviso", "No hay datos para guardar en la base de datos.")
            return 0, 0, 0

        conn = f"postgresql+psycopg2://{DB_CONFIG['USER']}:{DB_CONFIG['PASSWORD']}@" \
               f"{DB_CONFIG['HOST']}:{DB_CONFIG['PORT']}/{DB_CONFIG['DATABASE']}"
        engine = create_engine(conn)

        df.columns = df.columns.str.strip().str.lower()
        if "ot" in df.columns:
            df["ot"] = df["ot"].astype(str).str.strip()

        df_db = df.rename(columns={
            "idequipo": "idequipo",
            "cliente": "cliente",
            "fechaingreso": "fecha_ingreso",
            "fechasalida": "fecha_salida",
            "fechainforme": "fecha_informe",
            "cotizacion": "cotizacion",
            "actividad": "actividad",
            "ot": "ot",
            "descripcion": "descripcion"
        }).copy()

        total = len(df_db)
        insertados = 0
        repetidos = 0

        with engine.begin() as conn_trx:
            for _, row in df_db.iterrows():
                try:
                    result = conn_trx.execute(text(f"""
                        INSERT INTO {DB_CONFIG['TABLE']} 
                        (idequipo, cliente, fecha_ingreso, fecha_salida, fecha_informe,
                         cotizacion, actividad, ot, descripcion)
                        VALUES (:idequipo, :cliente, :fecha_ingreso, :fecha_salida, :fecha_informe,
                                :cotizacion, :actividad, :ot, :descripcion)
                        ON CONFLICT ON CONSTRAINT intervenciones_unica DO NOTHING
                    """), row.to_dict())

                    if result.rowcount and result.rowcount > 0:
                        insertados += 1
                    else:
                        repetidos += 1
                except Exception as e:
                    print(f"⚠️ Error insertando fila OT={row.get('ot')}: {e}")

        messagebox.showinfo(
            "Resultado",
            f"✅ {insertados} registros nuevos cargados.\n⚠️ {repetidos} ya existían."
        )
        return insertados, repetidos, total

    except Exception as e:
        messagebox.showerror("Error", f"No se pudo guardar en PostgreSQL: {e}")
        return 0, 0, 0
