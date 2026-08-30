"""Estilos y paleta de colores para la interfaz (CustomTkinter).

Colores suaves y amigables (nada de gris plano), tipografía grande y legible
pensada para una usuaria no técnica.

Contraste WCAG AA (S13): los colores de texto sobre fondos claros cumplen
>= 4.5:1 (verificado numéricamente con la fórmula de luminancia relativa).
"""

# Nombre único de la app (S14): todo texto visible a la usuaria usa esta
# constante, para no divergir (antes convivían "Notas Digital" y el nombre
# completo). El nombre técnico interno (carpeta %APPDATA%\\SistemaNotas, el
# .exe) NO se toca: es identidad de instalación, no texto de interfaz.
NOMBRE_APP = "Sistema de Digitalización de Planillas de Notas"

# --- Paleta suave ---
COLOR_FONDO = "#F4F6FB"          # fondo general, azul grisáceo muy suave
COLOR_FONDO_SECUNDARIO = "#FFFFFF"
COLOR_PRINCIPAL = "#5B8DEF"      # azul suave (botones principales)
COLOR_PRINCIPAL_HOVER = "#4A7BE0"
COLOR_ACCENTO = "#7C6BF0"        # violeta suave (detalles)
COLOR_TEXTO = "#2B3A55"
COLOR_TEXTO_SECUNDARIO = "#4E5E7C"  # AA: 6.04:1 sobre #F4F6FB y 6.53:1 sobre blanco
COLOR_REVISAR = "#FFE08A"        # amarillo suave para celdas dudosas
COLOR_REVISAR_BORDE = "#E0B64E"
COLOR_VERDE = "#6FBF8B"          # éxito
COLOR_ROJO = "#A94444"           # AA: 5.40:1 sobre #F4F6FB y 5.84:1 sobre blanco
COLOR_BLANCO = "#FFFFFF"

# --- Tipografía ---
FUENTE = "Segoe UI"
TAM_TITULO = 28
TAM_SUBTITULO = 20
TAM_TEXTO = 15
TAM_TEXTO_CHICO = 13
TAM_BOTON_GRANDE = 18
TAM_BOTON = 15
TAM_CELDA = 14


def configurar_tema():
    """Aplica el tema claro y la fuente por defecto a CustomTkinter."""
    import customtkinter as ctk

    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")
