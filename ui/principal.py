import os
import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

from etl.extract import extraer_excel, extraer_pdf
from etl.transform import transformar_datos
from etl.load import guardar_en_postgres   # debe devolver (insertados, repetidos, total)
from etl.utils import centrar_ventana

data_final = pd.DataFrame()
archivos_procesados = set()

def mostrar_tabla(tree, df):
    """Muestra los datos en la tabla."""
    for row in tree.get_children():
        tree.delete(row)

    tree["columns"] = list(df.columns)
    tree["show"] = "headings"

    for col in df.columns:
        tree.heading(col, text=col)
        width = 300 if col == "Descripcion" else 120
        tree.column(col, width=width, anchor="center")

    for _, row in df.iterrows():
        valores = [("" if (v is None or (isinstance(v, float) and pd.isna(v))) else str(v)) for v in row]
        tree.insert("", "end", values=valores)

def seleccionar_archivos(tree):
    files = filedialog.askopenfilenames(
        title="Seleccionar archivos",
        filetypes=[("Excel y PDF", ".xlsx .xls .pdf"), ("Todos los archivos", ".*")]
    )
    if files:
        procesar_archivos(files, tree)

def procesar_archivos(files, tree):
    global data_final, archivos_procesados
    dfs = []

    for file in files:
        if file in archivos_procesados:
            messagebox.showinfo("Aviso", f"El archivo ya fue procesado y será ignorado:\n{os.path.basename(file)}")
            continue

        ext = os.path.splitext(file)[1].lower()
        if ext in (".xlsx", ".xls"):
            df = extraer_excel(file)
        elif ext == ".pdf":
            df = extraer_pdf(file)
        else:
            continue

        df = transformar_datos(df)
        if not df.empty:
            dfs.append(df)
            archivos_procesados.add(file)

    if dfs:
        df_final = pd.concat(dfs, ignore_index=True).drop_duplicates(ignore_index=True)
        if "IDEquipo" in df_final.columns:
            df_final = df_final.sort_values(by="IDEquipo", kind="stable", ignore_index=True)

        data_final = df_final
        mostrar_tabla(tree, data_final)
    else:
        messagebox.showinfo("Aviso", "No se extrajo información nueva de los archivos seleccionados.")

def exportar_postgres():
    global data_final
    if not data_final.empty:
        try:
            insertados, repetidos, total = guardar_en_postgres(data_final)

            if repetidos > 0:
                messagebox.showinfo(
                    "Resultado",
                    f"✅ Se subieron {insertados} registros nuevos correctamente.\n"
                    f"⚠️ {repetidos} registros ya estaban en la base y no se subieron."
                )
            else:
                messagebox.showinfo("Resultado", f"✅ Se subieron {insertados} registros nuevos correctamente.")

        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un problema al exportar:\n{e}")
    else:
        messagebox.showwarning("Aviso", "No hay datos para exportar")

def abrir_ventana_principal(posicion=None, padre=None):
    """Abre la ventana principal ETL como subventana encima del menú."""
    root = ctk.CTkToplevel()
    root.title("G.C.A - Cargar Archivos")
    root.geometry("1100x650")

    ancho, alto = 1100, 650

    if posicion:
        x = posicion[0] - ancho // 2
        y = posicion[1] - alto // 2
        root.geometry(f"{ancho}x{alto}+{x}+{y}")
    else:
        centrar_ventana(root, ancho, alto)

    if padre:
        root.transient(padre)
    root.lift()
    root.focus_force()

    # ==============================
    # CONTENIDO
    # ==============================
    ctk.CTkLabel(root, text="Reto GCA - ETL (Excel + PDF)", font=("Arial", 16)).pack(pady=10)

    # Tabla
    frame_tabla = ctk.CTkFrame(root)
    frame_tabla.pack(fill="both", expand=True, padx=10, pady=10)

    y_scroll = ttk.Scrollbar(frame_tabla, orient="vertical")
    y_scroll.pack(side="right", fill="y")
    x_scroll = ttk.Scrollbar(frame_tabla, orient="horizontal")
    x_scroll.pack(side="bottom", fill="x")

    tree = ttk.Treeview(
        frame_tabla, show="headings", height=20,
        yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set
    )
    tree.pack(side="left", fill="both", expand=True)
    y_scroll.config(command=tree.yview)
    x_scroll.config(command=tree.xview)

    btn_frame = ctk.CTkFrame(root)
    btn_frame.pack(pady=10)

    ctk.CTkButton(btn_frame, text="Seleccionar archivos",
                  command=lambda: seleccionar_archivos(tree),
                  width=200, height=40).pack(side="left", padx=10)

    ctk.CTkButton(btn_frame, text="Exportar a PostgreSQL",
                  command=exportar_postgres,
                  width=200, height=40).pack(side="left", padx=10)

    ctk.CTkLabel(root, text="No hay datos cargados").pack(pady=5)

    # ==============================
    # CONFIRMACIÓN DE CIERRE
    # ==============================
    def confirmar_cierre():
        if messagebox.askyesno("Confirmar salida", "¿Está seguro de que desea cerrar esta ventana?"):
            root.destroy()

    root.protocol("WM_DELETE_WINDOW", confirmar_cierre)

    return root
