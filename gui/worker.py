"""
Procesamiento en segundo plano (hilo) para no congelar la interfaz.

El flujo completo (leer el PDF, convertir a imágenes, leer cada planilla con
Gemini) puede tardar. Se corre en un hilo aparte y se notifica el avance a la
interfaz mediante una cola de mensajes que la GUI revisa periódicamente.
"""

import queue
import threading

from config import app_config
from pdf_processing import pdf_loader
from vision import gemini_extractor


# Tipos de mensaje que se envían a la cola.
MSG_PROGRESO = "progreso"      # {"mensaje": str, "valor": float 0..1 (opcional)}
MSG_RESULTADO = "resultado"    # {"planillas": [...], "paginas_total": int, "fallidas": [...]}
MSG_ERROR = "error"            # {"mensaje": str}
MSG_CANCELADO = "cancelado"    # {"mensaje": str} — cancelación, no un error


class ProcesadorEnSegundoPlano:
    """
    Ejecuta la extracción en un hilo de fondo y envía mensajes a una cola.

    Uso:
        proc = ProcesadorEnSegundoPlano(ruta_pdf, cola)
        proc.iniciar()
        # La GUI consume cola.get() hasta ver MSG_RESULTADO o MSG_ERROR.
    """

    def __init__(self, ruta_pdf: str, cola: queue.Queue):
        self.ruta_pdf = ruta_pdf
        self.cola = cola
        self._hilo = None
        self.cancelado = False

    def iniciar(self):
        self._hilo = threading.Thread(target=self._ejecutar, daemon=True)
        self._hilo.start()

    def cancelar(self):
        self.cancelado = True

    def _ejecutar(self):
        try:
            dpi = app_config.load_config().get("preferencias", {}).get("dpi_pdf") or 250
            api_key = app_config.get_api_key()
            modelo = app_config.get_modelo_vision()

            # 1) Validar y convertir el PDF a páginas.
            self.cola.put({"tipo": MSG_PROGRESO, "mensaje": "Abriendo el PDF...", "valor": 0.02})
            paginas = pdf_loader.cargar_paginas(self.ruta_pdf, dpi=dpi)
            total = len(paginas)

            # 2) Leer cada planilla con Gemini.
            def _progreso(mensaje, valor=None):
                if self.cancelado:
                    raise _Cancelado()
                # None es sólo un chequeo de cancelación (lo llama el extractor
                # dentro del loop de reintentos); no se encola un mensaje.
                if mensaje is None:
                    return
                msg = {"tipo": MSG_PROGRESO, "mensaje": mensaje}
                if valor is not None:
                    msg["valor"] = valor
                self.cola.put(msg)

            planillas, fallidas = gemini_extractor.extraer_planilla_pdf(
                paginas,
                api_key=api_key,
                modelo=modelo,
                progreso_cb=_progreso,
            )

            if self.cancelado:
                raise _Cancelado()

            self.cola.put(
                {
                    "tipo": MSG_RESULTADO,
                    "planillas": planillas,
                    "paginas_total": total,
                    "fallidas": fallidas,
                }
            )
        except _Cancelado:
            # La cancelación NO es un error: se informa como un estado propio
            # (W4) para que la GUI lo muestre de forma amigable.
            self.cola.put({"tipo": MSG_CANCELADO, "mensaje": "El proceso fue cancelado."})
        except pdf_loader.PdfError as e:
            self.cola.put({"tipo": MSG_ERROR, "mensaje": str(e)})
        except gemini_extractor.VisionError as e:
            self.cola.put({"tipo": MSG_ERROR, "mensaje": str(e)})
        except Exception as e:
            app_config.escribir_log(f"Error inesperado en el procesamiento: {e!r}")
            self.cola.put(
                {
                    "tipo": MSG_ERROR,
                    "mensaje": "Ocurrió un problema inesperado al procesar el PDF. "
                    "Se guardó el detalle en el registro de errores.",
                }
            )


class _Cancelado(Exception):
    """Excepción interna para abortar el procesamiento si la usuaria cancela."""
