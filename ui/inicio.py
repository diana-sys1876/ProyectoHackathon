import customtkinter as ctk
from PIL import Image
import tkinter as tk
from etl.utils import centrar_ventana, fade_out, fade_in
from ui.menu import mostrar_menu  

def mostrar_inicio():
    # Set UI appearance
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    splash = ctk.CTk()
    splash.title("GCA-DATA")

    # Custom font for UI
    alice_font = ctk.CTkFont(family="Alice", size=18)

    # Load splash background image
    imagen_pil = Image.open("image/inicial.png")
    ancho, alto = imagen_pil.size
    imagen_fondo = ctk.CTkImage(light_image=imagen_pil, size=(ancho, alto))

    # Center window on screen
    centrar_ventana(splash, ancho, alto)

    # Start fully transparent for fade-in effect
    splash.attributes("-alpha", 0.0)

    # Set background image as full window
    label_fondo = ctk.CTkLabel(splash, image=imagen_fondo, text="")
    label_fondo.place(x=0, y=0, relwidth=1, relheight=1)

    # Smooth fade-in animation
    fade_in(splash)

    # ------------------------------------------------
    # Action executed when the user clicks "Entrar"
    # ------------------------------------------------
    def ir_a_menu():

        # Get current window position to open menu in same location
        splash.update_idletasks()
        geo = splash.geometry()

        import re
        match = re.search(r"\+(\d+)\+(\d+)", geo)
        x, y = (match.group(1), match.group(2)) if match else ("100", "100")

        # Cancel any pending animations (if running)
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

        # Destroy splash after fade-out and open menu
        def despues():
            if splash.winfo_exists():
                splash.destroy()
            mostrar_menu((x, y))

        # Smooth fade-out animation
        fade_out(splash, despues)

    # ------------------------------------------------
    # Main button to enter the system
    # ------------------------------------------------
    btn = ctk.CTkButton(
        splash,
        text="Entrar",
        corner_radius=2,
        fg_color="#4a90e2",  # Button background color
        hover_color="#357ABD",  # Hover color
        text_color="white",
        font=alice_font,
        command=ir_a_menu,  # Action to execute
        width=140,
        height=45
    )

    # Position button over the background image
    btn.place(x=ancho - 330, y=alto - 200)  

    # Start splash screen loop
    splash.mainloop()
