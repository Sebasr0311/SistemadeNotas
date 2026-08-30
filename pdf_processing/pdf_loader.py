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
    Abre el PDF y devuelve una lista de páginas SIN materializar las imágenes
    (carga perezosa, S7): rasterizar las decenas de páginas de una planilla
    anual a 250 DPI puede ocupar cientos de MB de RAM; acá cada página expone
    `_render`, un callable que rasteriza ESA página recién cuando se pide.

    Cada página es un dict:
        {
            "indice": int (1-based),
            "ancho_px": int,   # dimensiones de la página (page.rect, barato)
            "alto_px": int,
            "_render": callable() -> PIL.Image RGB de ESA página,
            "_doc": fitz.Document interno (NO tocar salvo cerrar_paginas),
        }

    IMPORTANTE: el documento queda ABIERTO mientras se renderizan páginas con
    `_render`. Cuando el procesamiento termine (éxito, error o cancelación),
    llamá a `cerrar_paginas(paginas)` para liberar el archivo.

    Lanza PdfError si el archivo no se puede leer.
    """
    validar_pdf(ruta_pdf)
    dpi = _dpi_sano(dpi)
    zoom = dpi / 72.0  # PyMuPDF trabaja en puntos; 72 puntos por pulgada.

    paginas = []
    try:
        doc = fitz.open(ruta_pdf)
        for i, page in enumerate(doc):
            ancho = int(round(page.rect.width))
            alto = int(round(page.rect.height))

            def _render(pagina=page, factor=zoom):
                # Rasteriza UNA página a la resolución pedida y devuelve la
                # imagen PIL RGB lista para el cliente de Gemini.
                from PIL import Image
                import io
                pix = pagina.get_pixmap(matrix=fitz.Matrix(factor, factor), alpha=False)
                datos = pix.tobytes("png")
                return Image.open(io.BytesIO(datos)).convert("RGB")

            paginas.append(
                {
                    "indice": i + 1,
                    "ancho_px": ancho,
                    "alto_px": alto,
                    "_render": _render,
                    "_doc": doc,
                }
            )
    except PdfError:
        raise
    except fitz.FileDataError as e:
        raise PdfError("No se pudo abrir el PDF. Asegurate de que no esté dañado.") from e
    except Exception as e:
        raise PdfError("Hubo un problema al convertir el PDF a imágenes.") from e

    if not paginas:
        raise PdfError("El PDF no tiene páginas para leer.")

    return paginas


def cerrar_paginas(paginas) -> None:
    """
    Cierra el documento PDF interno de las páginas cargadas con cargar_paginas.

    El doc queda abierto mientras se renderizan páginas con `_render` (PyMuPDF
    necesita el documento para rasterizar); una vez terminado el procesamiento
    (éxito, error o cancelación), se cierra acá para liberar los handles.

    Idempotente: tolera None, listas vacías y dicts sin "_doc" (por ejemplo,
    fixtures legacy de tests). Llamarla dos veces es seguro.
    """
    try:
        for pag in paginas or []:
            if not isinstance(pag, dict):
                continue
            doc = pag.get("_doc")
            if doc is None:
                continue
            try:
                doc.close()
            except Exception:
                pass
            finally:
                # Marcar como cerrado para que una segunda llamada no reintente.
                pag["_doc"] = None
    except Exception:
        # Cerrar es un best-effort: nunca debe romper el flujo principal.
        pass


def contar_paginas(ruta_pdf: str) -> int:
    """Devuelve el número de páginas del PDF (sin convertirlas todas)."""
    validar_pdf(ruta_pdf)
    doc = fitz.open(ruta_pdf)
    n = doc.page_count
    doc.close()
    return n
