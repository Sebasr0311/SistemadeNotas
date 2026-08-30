"""
Sistema de Digitalización de Planillas de Notas — punto de entrada.

Crea la ventana principal de CustomTkinter y arranca el flujo.
Para empaquetar a .exe se usa este archivo como script principal.
"""

import customtkinter as ctk

from gui.screens import App
from gui import styles


def main():
    styles.configurar_tema()
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
