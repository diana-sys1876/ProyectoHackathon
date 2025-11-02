import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import psycopg2
import pandas as pd
import os
import subprocess


def abrir_ventana_informes(posicion=None, padre=None):
    ventana = ctk.CTkToplevel()
    ventana.title("📊 Todos los Informes")

    ancho_ventana = 1100
    alto_ventana = 650

    if posicion:
        x = posicion[0] - (ancho_ventana // 2)
        y = posicion[1] - (alto_ventana // 2)
        ventana.geometry(f"{ancho_ventana}x{alto_ventana}+{x}+{y}")
    else:
        ancho_pantalla = ventana.winfo_screenwidth()
        alto_pantalla = ventana.winfo_screenheight()
        x = (ancho_pantalla // 2) - (ancho_ventana // 2)
        y = (alto_pantalla // 2) - (alto_ventana // 2)
        ventana.geometry(f"{ancho_ventana}x{alto_ventana}+{x}+{y}")

    ventana.resizable(False, False)

    if padre:
        ventana.transient(padre)
    ventana.grab_set()
    ventana.focus_force()

    # ==============================
    # 🔹 CONFIRMACIÓN AL CERRAR
    # ==============================
    def confirmar_cierre():
        if messagebox.askyesno("Confirmar salida", "¿Está seguro de que desea cerrar la ventana de informes?"):
            ventana.destroy()

    ventana.protocol("WM_DELETE_WINDOW", confirmar_cierre)

    # Frame principal
    frame = ctk.CTkFrame(ventana, fg_color="white")
    frame.pack(fill="both", expand=True, padx=10, pady=10)

    # ==============================
    # 🔹 FRAME SUPERIOR
    # ==============================
    frame_top = ctk.CTkFrame(frame, fg_color="white")
    frame_top.pack(fill="x", pady=5)

    entrada_buscar = ctk.CTkEntry(frame_top, placeholder_text="🔎 Buscar por Cliente o Equipo...", width=250)
    entrada_buscar.pack(side="left", padx=5)

    btn_buscar = ctk.CTkButton(frame_top, text="Buscar", width=100, command=lambda: buscar())
    btn_buscar.pack(side="left", padx=5)

    btn_exportar = ctk.CTkButton(frame_top, text="📤 Exportar a Excel", width=150, command=lambda: exportar_excel())
    btn_exportar.pack(side="left", padx=5)

    btn_refrescar = ctk.CTkButton(frame_top, text="🔄 Refrescar", width=120, command=lambda: cargar_datos())
    btn_refrescar.pack(side="left", padx=5)

    btn_dashboard = ctk.CTkButton(frame_top, text="📊 Dashboard", width=150, command=lambda: abrir_dashboard())
    btn_dashboard.pack(side="left", padx=5)

    btn_cerrar = ctk.CTkButton(frame_top, text="❌ Cerrar", fg_color="red", width=100, command=confirmar_cierre)
    btn_cerrar.pack(side="left", padx=5)

    # ==============================
    # 🔹 TABLA
    # ==============================
    columnas = (
        "ID", "Equipo", "Cliente", "Fecha Ingreso", "Fecha Salida",
        "Fecha Informe", "Cotización", "Actividad", "OT", "Descripción"
    )

    tabla = ttk.Treeview(frame, columns=columnas, show="headings", height=15)
    for col in columnas:
        tabla.heading(col, text=col)
        tabla.column(col, width=120, anchor="center")
    tabla.pack(fill="both", expand=True)

    scrollbar_y = ttk.Scrollbar(frame, orient="vertical", command=tabla.yview)
    scrollbar_x = ttk.Scrollbar(frame, orient="horizontal", command=tabla.xview)
    tabla.configure(yscroll=scrollbar_y.set, xscroll=scrollbar_x.set)
    scrollbar_y.pack(side="right", fill="y")
    scrollbar_x.pack(side="bottom", fill="x")

    # ==============================
    # 🔹 FRAME DESCRIPCIÓN DETALLADA
    # ==============================
    frame_desc = ctk.CTkFrame(frame, fg_color="#f5f5f5")
    frame_desc.pack(fill="x", pady=10)

    lbl_desc = ctk.CTkLabel(frame_desc, text="📝 Descripción Detallada:", anchor="w", font=("Arial", 14, "bold"))
    lbl_desc.pack(anchor="w", padx=5, pady=2)

    txt_descripcion = tk.Text(frame_desc, height=5, wrap="word", font=("Arial", 12))
    txt_descripcion.pack(fill="x", padx=5, pady=5)

    # ==============================
    # 🔹 FUNCIONES
    # ==============================
    def cargar_datos(filtro=None):
        for i in tabla.get_children():
            tabla.delete(i)
        try:
            conn = psycopg2.connect(
                host="localhost",
                dbname="gca_Inventario",
                user="postgres",
                password="123456",
                port="5432"
            )
            cursor = conn.cursor()

            if filtro:
                cursor.execute("""
                    SELECT id, idequipo, cliente, fecha_ingreso, fecha_salida,
                           fecha_informe, cotizacion, actividad, ot, descripcion
                    FROM intervenciones
                    WHERE LOWER(cliente) LIKE %s OR LOWER(idequipo::text) LIKE %s
                    ORDER BY id ASC;
                """, (f"%{filtro}%", f"%{filtro}%"))
            else:
                cursor.execute("""
                    SELECT id, idequipo, cliente, fecha_ingreso, fecha_salida,
                           fecha_informe, cotizacion, actividad, ot, descripcion
                    FROM intervenciones
                    ORDER BY id ASC;
                """)

            registros = cursor.fetchall()
            for fila in registros:
                fila = ["" if v is None else v for v in fila]
                tabla.insert("", tk.END, values=fila)

            conn.close()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron cargar los informes:\n{e}")

    def mostrar_detalle(event):
        item = tabla.selection()
        if item:
            valores = tabla.item(item, "values")
            descripcion = valores[-1] if valores else ""
            txt_descripcion.delete("1.0", tk.END)
            txt_descripcion.insert(tk.END, descripcion)

    def exportar_excel():
        try:
            items = tabla.get_children()
            if not items:
                messagebox.showwarning("Aviso", "No hay datos en la tabla para exportar.")
                return

            datos = [tabla.item(i, "values") for i in items]
            df = pd.DataFrame(datos, columns=columnas)

            # 📂 Elegir ruta de guardado
            ruta_salida = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Archivos Excel", "*.xlsx")],
                title="Guardar informe como"
            )
            if not ruta_salida:
                return  # Canceló

            # Guardar con pandas
            df.to_excel(ruta_salida, index=False)

            # ✅ Ajustar columnas al contenido
            from openpyxl import load_workbook
            wb = load_workbook(ruta_salida)
            ws = wb.active

            for col in ws.columns:
                max_length = 0
                col_letter = col[0].column_letter
                for cell in col:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                ws.column_dimensions[col_letter].width = max_length + 2  # margen extra

            wb.save(ruta_salida)

            messagebox.showinfo("✅ Exportación", f"Los informes se exportaron correctamente a:\n{ruta_salida}")

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo exportar:\n{e}")

    def buscar():
        filtro = entrada_buscar.get().lower()
        cargar_datos(filtro=filtro)

    def abrir_dashboard():
        try:
            ruta_dashboard = os.path.join(os.path.dirname(__file__), "..", "dashboard", "dashboardEmpresa.pbix")
            ruta_dashboard = os.path.abspath(ruta_dashboard)
            if os.path.exists(ruta_dashboard):
                subprocess.Popen([ruta_dashboard], shell=True)
            else:
                messagebox.showwarning("Archivo no encontrado", f"No se encontró el dashboard en:\n{ruta_dashboard}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el dashboard:\n{e}")

    tabla.bind("<Double-1>", mostrar_detalle)
    tabla.bind("<<TreeviewSelect>>", mostrar_detalle)

    cargar_datos()

    return ventana
