import customtkinter as ctk
from PIL import Image
import tkinter as tk
from etl.utils import centrar_ventana, fade_out, fade_in
from ui.menu import mostrar_menu  

def mostrar_inicio():
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    splash = ctk.CTk()
    splash.title("GCA-DATA")

    alice_font = ctk.CTkFont(family="Alice", size=18)

    imagen_pil = Image.open("image/inicial.png")
    ancho, alto = imagen_pil.size
    imagen_fondo = ctk.CTkImage(light_image=imagen_pil, size=(ancho, alto))

    centrar_ventana(splash, ancho, alto)
    splash.attributes("-alpha", 0.0)

    label_fondo = ctk.CTkLabel(splash, image=imagen_fondo, text="")
    label_fondo.place(x=0, y=0, relwidth=1, relheight=1)

    fade_in(splash)

    # -------------------
    # Acción al pulsar "Entrar"
    # -------------------
    def ir_a_menu():
        splash.update_idletasks()
        geo = splash.geometry()

        import re
        match = re.search(r"\+(\d+)\+(\d+)", geo)
        x, y = (match.group(1), match.group(2)) if match else ("100", "100")

        try:
            afters = splash.tk.call("after", "info")
            if isinstance(afters, (tuple, list)):
                for task in afters:
                    try:
                        splash.after_cancel(task)
                    except tk.TclError:
                        pass
        except tk.TclError:
            pass

        def despues():
            if splash.winfo_exists():
                splash.destroy()
            mostrar_menu((x, y))

        fade_out(splash, despues)

    btn = ctk.CTkButton(
        splash,
        text="Entrar",
        corner_radius=2,
        fg_color="#4a90e2",
        hover_color="#357ABD",
        text_color="white",
        font=alice_font,
        command=ir_a_menu,
        width=140,
        height=45
    )
    btn.place(x=ancho - 330, y=alto - 200)  

    splash.mainloop()
