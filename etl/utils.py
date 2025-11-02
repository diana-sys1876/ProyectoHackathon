import re
import pandas as pd
import tkinter as tk

# --------------------
# UI helpers
# --------------------
def centrar_ventana(ventana, ancho, alto):
    """Centra la ventana en la pantalla con las dimensiones indicadas."""
    ventana.update_idletasks()
    x = (ventana.winfo_screenwidth() // 2) - (ancho // 2)
    y = (ventana.winfo_screenheight() // 2) - (alto // 2)
    ventana.geometry(f"{ancho}x{alto}+{x}+{y}")

# ==========================
# Funciones para animaciones fade
# ==========================

def fade_in(window, alpha=0.0, step=0.05, delay=50):
    """Efecto de aparición gradual (fade in) seguro."""
    try:
        if not window.winfo_exists():
            return
        alpha = min(1.0, alpha + step)
        window.attributes("-alpha", alpha)
        if alpha < 1.0 and window.winfo_exists():
            window.after(delay, lambda: fade_in(window, alpha, step, delay))
    except tk.TclError:
        return


def fade_out(window, callback=None, alpha=None, step=0.05, delay=50):
    """Efecto de desaparición gradual (fade out) seguro."""
    try:
        if not window.winfo_exists():
            if callback:
                callback()
            return
        if alpha is None:
            try:
                alpha = float(window.attributes("-alpha"))
            except Exception:
                alpha = 1.0
        alpha = max(0.0, alpha - step)
        window.attributes("-alpha", alpha)
        if alpha > 0.0 and window.winfo_exists():
            window.after(delay, lambda: fade_out(window, callback, alpha, step, delay))
        else:
            if callback:
                callback()
    except tk.TclError:
        if callback:
            callback()
        return

# --------------------
# Data helpers
# --------------------
def limpiar_cliente(texto):
    """Limpia el nombre de cliente eliminando prefijos como 'Cliente:'."""
    if not texto:
        return None
    return re.sub(r"(?i)^cliente\s*[:\-]?\s*", "", str(texto)).strip()

def convertir_fecha(valor):
    """
    Convierte:
      - Serial de Excel (int/str de dígitos) con origen 1899-12-30
      - Cadenas día/mes/año o variantes comunes
    Devuelve en formato DD/MM/YYYY o None si no se puede.
    """
    try:
        if valor is None:
            return None
        s = str(valor).strip()
        if s == "" or s.lower() in ["nan", "none", "null"]:
            return None
        if s.isdigit():
            fecha = pd.to_datetime(int(s), origin="1899-12-30", unit="D", errors="coerce")
            if pd.notnull(fecha):
                return fecha.strftime("%d/%m/%Y")
            return None
        fecha = pd.to_datetime(s, dayfirst=True, errors="coerce")
        return fecha.strftime("%d/%m/%Y") if pd.notnull(fecha) else None
    except Exception:
        return None
