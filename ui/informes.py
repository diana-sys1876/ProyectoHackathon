import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import psycopg2
import pandas as pd
import os
import subprocess


# =====================================================
# 🔧 EDIT WINDOW — COMPACT DESIGN WITH 2 COLUMNS
# =====================================================
def abrir_ventana_editar(ventana_padre, datos):
    # Create edit window
    ventana_editar = ctk.CTkToplevel(ventana_padre)
    ventana_editar.title("✏️ Editar Registro")
    ventana_editar.geometry("620x520")
    ventana_editar.resizable(False, False)

    # List of field labels to generate entries dynamically
    labels = [
        "ID Equipo", "Cliente",
        "Fecha Ingreso", "Fecha Salida",
        "Fecha Informe", "Cotización",
        "Actividad", "OT"
    ]

    entradas = []

    # Main frame that contains all form inputs
    frame_form = ctk.CTkFrame(ventana_editar)
    frame_form.pack(fill="both", expand=True, padx=20, pady=10)

    # ------------------------------------------------------------
    # Create input fields using grid layout (real 2-column layout)
    # ------------------------------------------------------------
    for i, campo in enumerate(labels):
        fila = i // 2
        columna = i % 2

        lbl = ctk.CTkLabel(frame_form, text=campo, font=("Arial", 13))
        lbl.grid(row=fila * 2, column=columna, sticky="w", padx=10, pady=(5, 0))

        entry = ctk.CTkEntry(frame_form, width=250)
        entry.grid(row=fila * 2 + 1, column=columna, sticky="we", padx=10)
        entradas.append(entry)

    # Make both grid columns expand proportionally
    frame_form.grid_columnconfigure(0, weight=1)
    frame_form.grid_columnconfigure(1, weight=1)

    # ----------------- DESCRIPTION FIELD -----------------
    lbl_desc = ctk.CTkLabel(frame_form, text="Descripción", font=("Arial", 13))
    lbl_desc.grid(row=10, column=0, columnspan=2, sticky="w", pady=(15, 5), padx=10)

    # Text widget for multiline description
    txt_descripcion = tk.Text(frame_form, height=6, font=("Arial", 12), wrap="word")
    txt_descripcion.grid(row=11, column=0, columnspan=2, sticky="we", padx=10)
    entradas.append(txt_descripcion)

    # Fill fields with initial data from selected row
    valores_iniciales = list(datos[1:])  # omit ID

    for i, valor in enumerate(valores_iniciales):
        if i == 8:  # description index
            txt_descripcion.insert("1.0", valor)
        else:
            entradas[i].insert(0, valor)

    # ---------------- SAVE AND CANCEL BUTTONS ----------------
    frame_botones = ctk.CTkFrame(ventana_editar)
    frame_botones.pack(fill="x", pady=15)


    # Function to update the record in the database
    def guardar_cambios():
        try:
            # Connect to PostgreSQL database
            conn = psycopg2.connect(
                host="localhost",
                dbname="gca_Inventario",
                user="postgres",
                password="123456",
                port="5432"
            )
            cursor = conn.cursor()

            # SQL update query
            sql = """
                UPDATE intervenciones SET
                    idequipo = %s,
                    cliente = %s,
                    fecha_ingreso = %s,
                    fecha_salida = %s,
                    fecha_informe = %s,
                    cotizacion = %s,
                    actividad = %s,
                    ot = %s,
                    descripcion = %s
                WHERE id = %s
            """

            # Gather updated values
            valores = [
                entradas[0].get(), entradas[1].get(),
                entradas[2].get(), entradas[3].get(),
                entradas[4].get(), entradas[5].get(),
                entradas[6].get(), entradas[7].get(),
                txt_descripcion.get("1.0", tk.END).strip(),
                datos[0]
            ]

            cursor.execute(sql, valores)
            conn.commit()
            conn.close()

            messagebox.showinfo("✔️ Éxito", "Registro actualizado correctamente.")
            ventana_editar.destroy()

        except Exception as e:
            # Show error message if update fails
            messagebox.showerror("❌ Error", f"No se pudo actualizar:\n{e}")

    # Save button
    btn_guardar = ctk.CTkButton(frame_botones, text="💾 Guardar", fg_color="green", width=150, command=guardar_cambios)
    btn_guardar.pack(side="left", padx=20)

    # Cancel button
    btn_cancelar = ctk.CTkButton(frame_botones, text="Cancelar", fg_color="red", width=150, command=ventana_editar.destroy)
    btn_cancelar.pack(side="right", padx=20)



# =====================================================
# 🔧 MAIN WINDOW — SHOW ALL REPORTS
# =====================================================
def abrir_ventana_informes(posicion=None, padre=None):

    # Main window that displays all records
    ventana = ctk.CTkToplevel()
    ventana.title("📊 Todos los Informes")

    ancho_ventana = 1100
    alto_ventana = 650

    # Center window on screen
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

    # Confirmation when closing window
    def confirmar_cierre():
        if messagebox.askyesno("Confirmar salida", "¿Está seguro de que desea cerrar la ventana de informes?"):
            ventana.destroy()

    ventana.protocol("WM_DELETE_WINDOW", confirmar_cierre)

    # ---------------- MAIN FRAME ----------------
    frame = ctk.CTkFrame(ventana, fg_color="white")
    frame.pack(fill="both", expand=True, padx=10, pady=10)

    # Top frame with action buttons
    frame_top = ctk.CTkFrame(frame, fg_color="white")
    frame_top.pack(fill="x", pady=5)

    # Search bar
    entrada_buscar = ctk.CTkEntry(frame_top, placeholder_text="🔎 Buscar por Cliente o Equipo...", width=250)
    entrada_buscar.pack(side="left", padx=5)

    # Search button
    btn_buscar = ctk.CTkButton(frame_top, text="Buscar", width=100, command=lambda: buscar())
    btn_buscar.pack(side="left", padx=5)

    # Export to Excel button
    btn_exportar = ctk.CTkButton(frame_top, text="📤 Exportar a Excel", width=150, command=lambda: exportar_excel())
    btn_exportar.pack(side="left", padx=5)

    # Refresh data button
    btn_refrescar = ctk.CTkButton(frame_top, text="🔄 Refrescar", width=120, command=lambda: cargar_datos())
    btn_refrescar.pack(side="left", padx=5)

    # Open dashboard button
    btn_dashboard = ctk.CTkButton(frame_top, text="📊 Dashboard", width=150, command=lambda: abrir_dashboard())
    btn_dashboard.pack(side="left", padx=5)

    # Edit selected record button
    btn_editar = ctk.CTkButton(frame_top, text="✏️ Editar", width=120, command=lambda: editar_registro())
    btn_editar.pack(side="left", padx=5)

    # Close window button
    btn_cerrar = ctk.CTkButton(frame_top, text="❌ Cerrar", fg_color="red", width=100, command=confirmar_cierre)
    btn_cerrar.pack(side="left", padx=5)

    # --------------- TABLE VIEW ----------------
    columnas = (
        "ID", "Equipo", "Cliente", "Fecha Ingreso", "Fecha Salida",
        "Fecha Informe", "Cotización", "Actividad", "OT", "Descripción"
    )

    # Treeview widget for showing all database records
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

    # ------------ DESCRIPTION INSPECTOR ------------
    frame_desc = ctk.CTkFrame(frame, fg_color="#f5f5f5")
    frame_desc.pack(fill="x", pady=10)

    lbl_desc = ctk.CTkLabel(frame_desc, text="📝 Descripción Detallada:", anchor="w", font=("Arial", 14, "bold"))
    lbl_desc.pack(anchor="w", padx=5, pady=2)

    txt_descripcion = tk.Text(frame_desc, height=5, wrap="word", font=("Arial", 12))
    txt_descripcion.pack(fill="x", padx=5, pady=5)



    # =====================================================
    # INTERNAL FUNCTIONS
    # =====================================================

    # Load all records from database
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

            # Query with optional filter
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

            # Insert rows into Treeview
            for fila in registros:
                fila = ["" if v is None else v for v in fila]
                tabla.insert("", tk.END, values=fila)

            conn.close()

        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron cargar los informes:\n{e}")


    # Display detailed description of a selected row
    def mostrar_detalle(event):
        item = tabla.selection()
        if item:
            valores = tabla.item(item, "values")
            descripcion = valores[-1]
            txt_descripcion.delete("1.0", tk.END)
            txt_descripcion.insert(tk.END, descripcion)


    # Export data to an Excel file
    def exportar_excel():
        try:
            items = tabla.get_children()
            if not items:
                messagebox.showwarning("Aviso", "No hay datos en la tabla para exportar.")
                return

            datos = [tabla.item(i, "values") for i in items]
            df = pd.DataFrame(datos, columns=columnas)

            ruta_salida = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Archivos Excel", "*.xlsx")],
                title="Guardar informe como"
            )
            if not ruta_salida:
                return

            df.to_excel(ruta_salida, index=False)

            # Auto-adjust Excel column widths
            from openpyxl import load_workbook
            wb = load_workbook(ruta_salida)
            ws = wb.active

            for col in ws.columns:
                max_length = 0
                col_letter = col[0].column_letter
                for cell in col:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                ws.column_dimensions[col_letter].width = max_length + 2

            wb.save(ruta_salida)

            messagebox.showinfo("Éxito", f"Exportado correctamente en:\n{ruta_salida}")

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo exportar:\n{e}")


    # Search function
    def buscar():
        filtro = entrada_buscar.get().lower()
        cargar_datos(filtro=filtro)


    # Open Power BI dashboard
    def abrir_dashboard():
        try:
            ruta_dashboard = os.path.join(os.path.dirname(__file__), "..", "dashboard", "dashboardEmpresa.pbix")
            ruta_dashboard = os.path.abspath(ruta_dashboard)

            if os.path.exists(ruta_dashboard):
                subprocess.Popen([ruta_dashboard], shell=True)
            else:
                messagebox.showwarning("No encontrado", f"No se encontró:\n{ruta_dashboard}")

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir:\n{e}")


    # Open edit window with selected row data
    def editar_registro():
        item = tabla.selection()
        if not item:
            messagebox.showwarning("Aviso", "Seleccione un registro para editar.")
            return

        valores = tabla.item(item, "values")
        abrir_ventana_editar(ventana, valores)

    # Bind events
    tabla.bind("<Double-1>", mostrar_detalle)
    tabla.bind("<<TreeviewSelect>>", mostrar_detalle)

    # Load data when window opens
    cargar_datos()

    return ventana
