"""
Carga de PDF e imágenes, y conversión de páginas a imágenes.

Usa PyMuPDF (fitz) para abrir PDFs y convertir cada página en una imagen PNG de
buena calidad (entre 200 y 300 DPI), y PIL para leer imágenes (JPG/PNG) y
tratarlas como una planilla cada una. Todo queda listo para que el motor de
visión la lea.

La app acepta una mezcla de PDFs y de imágenes en el MISMO lote: cada imagen es
UNA planilla, y cada página de PDF es UNA planilla. `cargar_archivo` y
`contar_planillas` despachan según la extensión de cada archivo.

Excepciones:
- Errores legibles para la usuaria ante archivos que no abren o no son válidos.
"""

import os

import pymupdf as fitz  # PyMuPDF (el paquete moderno; expone la API `fitz`).

DPI_POR_DEFECTO = 250
DPI_MINIMO = 150
DPI_MAXIMO = 400

# Extensiones de archivos de imagen aceptadas como planillas.
EXTENSIONES_IMAGEN = (".jpg", ".jpeg", ".png")

# Límite de planillas por lote (PDFs + imágenes) para no saturar la cuota
# gratuita de la API.
MAX_PLANILLAS_POR_LOTE = 30
# Alias histórico: conserva el valor anterior para no romper los imports y los
# tests existentes que comparan contra el límite por PDF.
MAX_PLANILLAS_POR_PDF = MAX_PLANILLAS_POR_LOTE


class PdfError(Exception):
    """Error de lectura del archivo (PDF o imagen) con mensaje amigable."""


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
    """Devuelve el número de páginas del PDF (sin rasterizar nada).

    Barato: no convierte las páginas a imágenes, solo abre el documento y lee
    `page_count`. Si la ruta no existe, no es un PDF o no se puede abrir,
    lanza PdfError (igual que validar_pdf). El documento siempre se cierra.
    """
    validar_pdf(ruta_pdf)
    doc = fitz.open(ruta_pdf)
    try:
        return doc.page_count
    finally:
        doc.close()


def es_imagen(ruta: str) -> bool:
    """Devuelve True si la ruta termina en una extensión de imagen (JPG/PNG)."""
    return bool(ruta) and os.path.basename(str(ruta)).lower().endswith(EXTENSIONES_IMAGEN)


def es_pdf(ruta: str) -> bool:
    """Devuelve True si la ruta termina en .pdf."""
    return bool(ruta) and os.path.basename(str(ruta)).lower().endswith(".pdf")


def validar_imagen(ruta_imagen: str) -> None:
    """Valida que la ruta exista, sea una imagen y pueda abrirse con PIL.

    Lanza PdfError con un mensaje claro si algo no anda.
    """
    if not ruta_imagen or not os.path.exists(ruta_imagen):
        raise PdfError("No se encontró el archivo. Revisá que hayas elegido la imagen correcta.")
    if not os.path.isfile(ruta_imagen):
        raise PdfError("La ruta elegida no es un archivo válido.")
    if not es_imagen(ruta_imagen):
        raise PdfError(
            "El archivo no es una imagen compatible. Usá imágenes JPG o PNG "
            "(o un PDF) para cargar planillas."
        )
    try:
        from PIL import Image
        with Image.open(ruta_imagen) as img:
            img.verify()  # Fuerza la lectura de la cabecera; detecta archivos rotos.
    except Exception as e:
        raise PdfError("No se pudo abrir la imagen. Asegurate de que no esté dañada o cortada.") from e


def cargar_imagen(ruta_imagen: str, dpi: int = DPI_POR_DEFECTO):
    """Carga UNA imagen (JPG/PNG) y la devuelve como una lista de UN dict.

    Cada imagen se trata como UNA planilla, con la MISMA forma que una página
    de `cargar_paginas` (S7: carga perezosa con `_render`):

        {
            "indice": 1,
            "ancho_px": int,
            "alto_px": int,
            "_render": callable() -> PIL.Image RGB (materializa la imagen recién
                       cuando se pide),
            "_doc": None (no hay PDF interno que cerrar),
            "es_imagen": True,
        }

    La imagen se lee de disco recién en `_render` (perezoso), igual que las
    páginas de PDF, para no guardar la imagen completa en memoria al cargar.
    Lanza PdfError si el archivo no se puede leer.
    """
    validar_imagen(ruta_imagen)
    dpi = _dpi_sano(dpi)

    def _render():
        from PIL import Image
        # La imagen se materializa y se devuelve RGB (como hacen las páginas).
        with Image.open(ruta_imagen) as img:
            img.load()
            return img.convert("RGB")

    try:
        from PIL import Image
        with Image.open(ruta_imagen) as img:
            ancho = int(img.size[0])
            alto = int(img.size[1])
    except Exception as e:
        raise PdfError("No se pudo abrir la imagen. Asegurate de que no esté dañada.") from e

    return [
        {
            "indice": 1,
            "ancho_px": ancho,
            "alto_px": alto,
            "_render": _render,
            "_doc": None,
            "es_imagen": True,
        }
    ]


def cargar_archivo(ruta: str, dpi: int = DPI_POR_DEFECTO):
    """Carga un archivo (PDF o imagen) y devuelve la lista de páginas/planillas.

    Despacha según la extensión:
    - PDF      -> cargar_paginas(ruta, dpi)
    - Imagen   -> cargar_imagen(ruta, dpi) (lista de UN dict)
    - Otro     -> PdfError amigable.
    """
    if es_pdf(ruta):
        return cargar_paginas(ruta, dpi)
    if es_imagen(ruta):
        return cargar_imagen(ruta, dpi)
    raise PdfError(
        "El archivo no es un PDF ni una imagen compatible. "
        "Usá archivos .pdf, .jpg, .jpeg o .png."
    )


def contar_planillas(ruta: str) -> int:
    """Devuelve la cantidad de planillas que aporta un archivo.

    - PDF    -> contar_paginas(ruta) (una páginas = una planilla).
    - Imagen -> 1 (una imagen = una planilla).
    - Otro   -> PdfError amigable.
    """
    if es_pdf(ruta):
        return contar_paginas(ruta)
    if es_imagen(ruta):
        validar_imagen(ruta)
        return 1
    raise PdfError(
        "El archivo no es un PDF ni una imagen compatible. "
        "Usá archivos .pdf, .jpg, .jpeg o .png."
    )


def render_imagen_planilla(planilla: dict, dpi: int = DPI_POR_DEFECTO):
    """Rasteriza y devuelve la imagen RGB de la planilla indicada.

    Dado un dict planilla (el de `extraer_planilla_pdf` con el campo adiciona
    `"fuente"`), renderiza la imagen de ESA planilla usando su origen:

    - fuente tipo "pdf"   : abre el PDF con fitz y rasteriza la página
                            `fuente["indice"]` con zoom dpi/72. El doc se
                            cierra SIEMPRE (try/finally).
    - fuente tipo "imagen": abre la imagen con PIL y la convierte a RGB.
    - Sin campo "fuente"  : devuelve None (no hay forma de saber el origen).

    NO guarda nada en disco.
    """
    if not isinstance(planilla, dict):
        return None
    fuente = planilla.get("fuente")
    if not isinstance(fuente, dict):
        return None
    tipo = fuente.get("tipo")
    ruta = fuente.get("ruta")

    dpi = _dpi_sano(dpi)
    if tipo == "pdf" and ruta:
        doc = None
        try:
            doc = fitz.open(ruta)
            indice = fuente.get("indice", 1)
            if indice < 1 or indice > doc.page_count:
                return None
            zoom = dpi / 72.0
            pix = doc[indice - 1].get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
            from PIL import Image
            import io
            return Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
        except Exception:
            return None
        finally:
            if doc is not None:
                try:
                    doc.close()
                except Exception:
                    pass
    if tipo == "imagen" and ruta:
        try:
            from PIL import Image
            with Image.open(ruta) as img:
                img.load()
                return img.convert("RGB")
        except Exception:
            return None
    return None
