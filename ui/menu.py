import customtkinter as ctk
from PIL import Image
import os
import sys
import subprocess
import shutil
from datetime import datetime
import tkinter.messagebox as msg
from ui.principal import abrir_ventana_principal  


def resource_path(relative_path: str) -> str:
    """Obtiene la ruta absoluta, compatible con PyInstaller."""
    try:
        base_path = sys._MEIPASS  
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def mostrar_menu(pos=(100, 100)):
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    ventana = ctk.CTk()
    ventana.title("G.C.A - Menú Principal")
    ventana.geometry(f"1100x650+{pos[0]}+{pos[1]}")
    ventana.resizable(False, False)

    # ==============================
    # 🔹 CONFIRMACIÓN AL CERRAR
    # ==============================
    def confirmar_cierre():
        if msg.askyesno("Confirmar salida", "¿Está seguro de que desea salir de la aplicación?"):
            ventana.destroy()

    ventana.protocol("WM_DELETE_WINDOW", confirmar_cierre)

    ruta_img = resource_path("image/Menu.png")
    if os.path.exists(ruta_img):
        try:
            imagen_pil = Image.open(ruta_img)
            imagen_fondo = ctk.CTkImage(light_image=imagen_pil, size=(1100, 650))
            fondo = ctk.CTkLabel(ventana, image=imagen_fondo, text="")
            fondo.place(x=0, y=0, relwidth=1, relheight=1)
        except Exception as e:
            msg.showwarning("Error cargando imagen", f"No se pudo abrir Menu.png\n{e}")
    else:
        ventana.configure(fg_color="#f0f0f0") 

    def abrir_opcion(modo):
        ventana.update_idletasks()
        x = ventana.winfo_x()
        y = ventana.winfo_y()
        w = ventana.winfo_width()
        h = ventana.winfo_height()
        centro_menu = (x + w // 2, y + h // 2)

        if modo == "cargar":
            nueva = abrir_ventana_principal(posicion=centro_menu, padre=ventana)
            nueva.grab_set()  

        elif modo == "nuevo":
            plantilla = resource_path(os.path.join("archivo", "PLANTILLA_INFORME TÉCNICO MANTENIMIENTO EQUIPOS.xlsx"))
            carpeta_salida = resource_path(os.path.join("archivo", "informes_generados"))
            os.makedirs(carpeta_salida, exist_ok=True)

            if os.path.exists(plantilla):
                timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                copia = os.path.join(carpeta_salida, f"INFORME_{timestamp}.xlsx")
                shutil.copy(plantilla, copia)

                try:
                    if sys.platform.startswith("win"):
                        os.startfile(copia)
                    elif sys.platform == "darwin":
                        subprocess.call(["open", copia])
                    else:
                        subprocess.call(["xdg-open", copia])
                except Exception as e:
                    msg.showerror("Error", f"No se pudo abrir el archivo:\n{e}")
            else:
                msg.showwarning("Archivo no encontrado", f"No existe el archivo:\n{plantilla}")

        elif modo == "todos":
            try:
                from ui.informes import abrir_ventana_informes
                nueva = abrir_ventana_informes(posicion=centro_menu, padre=ventana)
                nueva.grab_set()
            except ImportError:
                msg.showerror("Error", "No se pudo cargar la interfaz de informes.")

    botones = [
        ("Iniciar", "cargar", 110, 430),
        ("Crear", "nuevo", 450, 430),
        ("Ver", "todos", 800, 430)
    ]

    for texto, modo, x, y in botones:
        btn = ctk.CTkButton(
            ventana,
            text=texto,
            width=200,
            height=60,
            corner_radius=15,
            fg_color="#4a90e2",
            hover_color="#357ABD",
            text_color="white",
            font=("Arial", 16, "bold"),
            command=lambda m=modo: abrir_opcion(m)
        )
        btn.place(x=x, y=y)

    ventana.mainloop()
