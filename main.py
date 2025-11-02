import warnings

# 🔹 Ocultar advertencias de openpyxl relacionadas con imágenes/dibujos
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

from ui.inicio import mostrar_inicio

if __name__ == "__main__":
    mostrar_inicio()
