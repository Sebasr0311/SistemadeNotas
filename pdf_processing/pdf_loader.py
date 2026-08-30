"""
Carga de PDF y conversión de páginas a imágenes.

Usa PyMuPDF (fitz) para abrir el archivo PDF y convertir cada página en una
imagen PNG de buena calidad (entre 200 y 300 DPI), lista para que el motor de
visión la lea.

Excepciones:
- Errores legibles para la usuaria ante archivos que no abren o no son PDF.
"""

import os

import pymupdf as fitz  # PyMuPDF (el paquete moderno; expone la API `fitz`).

DPI_POR_DEFECTO = 250
DPI_MINIMO = 150
DPI_MAXIMO = 400


class PdfError(Exception):
    """Error de lectura del PDF con mensaje amigable para la usuaria."""


def validar_pdf(ruta_pdf: str) -> None:
    """
    Valida que la ruta exista, sea un archivo PDF y pueda abrirse con PyMuPDF.
    Lanza PdfError con un mensaje claro si algo no anda.
    """
    if not ruta_pdf or not os.path.exists(ruta_pdf):
        raise PdfError("No se encontró el archivo. Revisá que hayas elegido el PDF correcto.")
    if not os.path.isfile(ruta_pdf):
        raise PdfError("La ruta elegida no es un archivo válido.")
    nombre = os.path.basename(ruta_pdf).lower()
    if not nombre.endswith(".pdf"):
        raise PdfError("El archivo no parece ser un PDF (debe terminar en .pdf).")
    try:
        doc = fitz.open(ruta_pdf)
    except Exception as e:
        raise PdfError("No se pudo abrir el PDF. Asegurate de que no esté dañado o protegido.") from e
    if doc.page_count < 1:
        doc.close()
        raise PdfError("El PDF está vacío: no tiene ninguna página.")
    doc.close()


def _dpi_sano(dpi: int) -> int:
    """Devuelve un DPI razonable (200-300 por defecto) sin valores extremos."""
    if not dpi:
        return DPI_POR_DEFECTO
    return max(DPI_MINIMO, min(DPI_MAXIMO, int(dpi)))


def cargar_paginas(ruta_pdf: str, dpi: int = DPI_POR_DEFECTO):
    """
    Abre el PDF y devuelve una lista de páginas.

    Cada página es un dict:
        {
            "indice": int (0-based y 1-based),
            "imagen": PIL.Image (imagen de la página a buena resolución),
            "alto_px": int,
            "ancho_px": int,
        }

    Lanza PdfError si el archivo no se puede leer.
    """
    validar_pdf(ruta_pdf)
    dpi = _dpi_sano(dpi)
    zoom = dpi / 72.0  # PyMuPDF trabaja en puntos; 72 puntos por pulgada.

    paginas = []
    try:
        doc = fitz.open(ruta_pdf)
        for i, page in enumerate(doc):
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
            datos = pix.tobytes("png")
            # Convertir los bytes PNG a una imagen PIL para el cliente de Gemini.
            from PIL import Image
            import io
            imagen = Image.open(io.BytesIO(datos)).convert("RGB")
            paginas.append(
                {
                    "indice": i + 1,
                    "imagen": imagen,
                    "ancho_px": pix.width,
                    "alto_px": pix.height,
                }
            )
        doc.close()
    except PdfError:
        raise
    except fitz.FileDataError as e:
        raise PdfError("No se pudo abrir el PDF. Asegurate de que no esté dañado.") from e
    except Exception as e:
        raise PdfError("Hubo un problema al convertir el PDF a imágenes.") from e

    if not paginas:
        raise PdfError("El PDF no tiene páginas para leer.")

    return paginas


def contar_paginas(ruta_pdf: str) -> int:
    """Devuelve el número de páginas del PDF (sin convertirlas todas)."""
    validar_pdf(ruta_pdf)
    doc = fitz.open(ruta_pdf)
    n = doc.page_count
    doc.close()
    return n
