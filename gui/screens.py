"""
Pantallas de la interfaz (CustomTkinter).

Flujo lineal de 5 pantallas pensado para una usuaria no técnica:
  1. Configuración inicial (sólo la primera vez, si no hay clave).
  2. Pantalla principal: cargar el PDF de planillas.
  3. Progreso: barra + mensajes mientras se lee.
  4. Revisión: resumen por curso, celdas dudosas en amarillo, corregibles.
  5. Final: éxito, botones para abrir el archivo o la carpeta.

Todo el texto está en español, simple y cercano. Los errores siempre se
muestran con mensajes claros (no tracebacks) y se guardan en el log.
"""

import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

from config import app_config
from excel import generar_excel_notas
from excel.generar_excel_notas import ColumnConfig, nombres_columnas_areas
from excel.agrupacion import agrupar_por_curso, agrupar_por_forma, combinar_estudiantes, _asignatura_limpia
from pdf_processing import pdf_loader
from . import styles
from .worker import ProcesadorEnSegundoPlano, MSG_PROGRESO, MSG_RESULTADO, MSG_ERROR, MSG_CANCELADO

ABRIR_ARCHIVO = "open_file"
ABRIR_CARPETA = "open_folder"


class App(ctk.CTk):
    """Ventana principal y controlador del flujo entre pantallas."""

    def __init__(self):
        super().__init__()
        styles.configurar_tema()
        self.title(styles.NOMBRE_APP)
        self.geometry("860x640")
        self.minsize(720, 560)
        self._configurar_apariencia()

        self.planillas = []           # planillas extraídas (una por página/imagen)
        self.paginas_total = 0
        self.paginas_fallidas = []    # páginas que no se pudieron leer (S2)
        self.planilla_actual_idx = 0  # índice usado en la pantalla de revisión
        self._column_configs = {}     # {clave_forma: ColumnConfig} por forma
        self._worker_cola = None
        self._worker = None

        # Contenedor único donde se montan las pantallas.
        self._contenedor = ctk.CTkFrame(self, fg_color=styles.COLOR_FONDO)
        self._contenedor.pack(fill="both", expand=True)
        self._pantalla_actual = None

        self._mostrar_inicio()

        # S9: la X de la ventana pasa por acá. Antes, cerrar con el worker
        # activo mataba el hilo daemon a mitad de proceso y la GUI podía
        # programar afters sobre una ventana ya destruida.
        self.protocol("WM_DELETE_WINDOW", self._al_cerrar)

    @property
    def _column_config(self):
        """Compatibilidad: devuelve la config de la primera forma si existe.

        Antes había UNA sola ColumnConfig global (self._column_config). Ahora
        hay una por forma (self._column_configs). Esta property se conserva
        para no romper usos viejos dentro de este mismo archivo que esperaban
        el atributo plano.
        """
        if not self._column_configs:
            return None
        # Devuelve la config de la primera forma (orden de inserción).
        primera = next(iter(self._column_configs.values()))
        return primera

    @_column_config.setter
    def _column_config(self, valor):
        """Setter de compatibilidad: si se asigna una config sin forma conocida,
        se guarda bajo una forma genérica si aún no hay una."""
        if valor is None:
            return
        if not self._column_configs and valor.columnas:
            clave = ("n_areas", len(valor.columnas))
            self._column_configs[clave] = valor
        elif not self._column_configs:
            # Sin columnas: guardar bajo una forma genérica de 0 para no perderlo.
            self._column_configs[("n_areas", 0)] = valor

    def _al_cerrar(self):
        """Cierra la ventana de forma segura (S9).

        Si hay un procesamiento en curso, se pregunta ANTES de cancelar: la
        usuaria decide si salir igual (se cancela el proceso) o quedarse.
        """
        worker_activo = (
            self._worker is not None
            and getattr(self._worker, "_hilo", None) is not None
            and self._worker._hilo.is_alive()
        )
        if worker_activo:
            confirma = messagebox.askyesno(
                "¿Salir?",
                "Hay un procesamiento en curso. ¿Querés salir igual? "
                "El procesamiento se cancelará.",
            )
            if not confirma:
                return
            self._worker.cancelar()
        self.destroy()

    def _configurar_apariencia(self):
        self.configure(fg_color=styles.COLOR_FONDO)

    # ------------------------------------------------------------------ #
    # Navegación entre pantallas
    # ------------------------------------------------------------------ #
    def _cambiar_pantalla(self, widget):
        if self._pantalla_actual is not None:
            self._pantalla_actual.destroy()
        self._pantalla_actual = widget
        widget.pack(fill="both", expand=True, padx=24, pady=24)

    # ------------------------------------------------------------------ #
    # 1) Pantalla de configuración inicial
    # ------------------------------------------------------------------ #
    def _mostrar_inicio(self):
        if app_config.has_api_key():
            self.mostrar_principal()
        else:
            self.mostrar_configuracion()

    def mostrar_configuracion(self):
        pantalla = ctk.CTkFrame(self._contenedor, fg_color=styles.COLOR_FONDO)

        titulo = ctk.CTkLabel(
            pantalla, text="¡Bienvenida!",
            font=(styles.FUENTE, styles.TAM_TITULO, "bold"), text_color=styles.COLOR_TEXTO,
        )
        titulo.pack(pady=(20, 4))

        ctk.CTkLabel(
            pantalla, text=styles.NOMBRE_APP,
            font=(styles.FUENTE, styles.TAM_SUBTITULO, "bold"),
            text_color=styles.COLOR_TEXTO_SECUNDARIO,
        ).pack(pady=(0, 8))

        intro = (
            "Para poder leer las notas escritas a mano de tus planillas, la app\n"
            "necesita una clave de Google AI (es gratis y se obtiene en un minuto)."
        )
        ctk.CTkLabel(
            pantalla, text=intro, font=(styles.FUENTE, styles.TAM_TEXTO),
            text_color=styles.COLOR_TEXTO_SECUNDARIO, justify="center",
        ).pack(pady=(0, 4))

        ctk.CTkLabel(
            pantalla,
            text="Se obtiene gratis en Google AI Studio. Sólo se guarda en tu computador "
            "y se usa para leer las planillas.",
            font=(styles.FUENTE, styles.TAM_TEXTO_CHICO),
            text_color=styles.COLOR_TEXTO_SECUNDARIO, justify="center", wraplength=560,
        ).pack(pady=(0, 24))

        caja = ctk.CTkFrame(pantalla, fg_color=styles.COLOR_BLANCO, corner_radius=14)
        caja.pack(padx=20, pady=8)

        ctk.CTkLabel(
            caja, text="Pegá acá tu clave de Google AI:",
            font=(styles.FUENTE, styles.TAM_SUBTITULO, "bold"), text_color=styles.COLOR_TEXTO,
        ).pack(pady=(18, 8))

        self._campo_clave = ctk.CTkEntry(
            caja, width=460, height=44, show="•", font=(styles.FUENTE, styles.TAM_TEXTO),
            fg_color=styles.COLOR_FONDO_SECUNDARIO, border_color=styles.COLOR_PRINCIPAL,
        )
        self._campo_clave.pack(pady=(0, 8), padx=20)

        ctk.CTkLabel(
            caja, text="También podés obtener una clave gratuita en Google AI Studio.",
            font=(styles.FUENTE, styles.TAM_TEXTO_CHICO), text_color=styles.COLOR_TEXTO_SECUNDARIO,
        ).pack(pady=(0, 16))

        boton = ctk.CTkButton(
            caja, text="Guardar y continuar", height=48, font=(styles.FUENTE, styles.TAM_BOTON_GRANDE, "bold"),
            fg_color=styles.COLOR_PRINCIPAL, hover_color=styles.COLOR_PRINCIPAL_HOVER,
            command=self._guardar_clave,
        )
        boton.pack(pady=(0, 22), padx=20)

        # Validar y mostrar error amigable si está vacía.
        self._error_clave = ctk.CTkLabel(
            caja, text="", font=(styles.FUENTE, styles.TAM_TEXTO_CHICO), text_color=styles.COLOR_ROJO
        )
        self._error_clave.pack(pady=(0, 10))

        self._cambiar_pantalla(pantalla)
        self._campo_clave.focus_set()

    def _guardar_clave(self):
        clave = (self._campo_clave.get() or "").strip()
        if not clave:
            self._error_clave.configure(text="Necesitás pegar la clave para continuar.")
            return
        try:
            app_config.set_api_key(clave)
        except Exception:
            messagebox.showerror(
                "No se pudo guardar",
                "No se pudo guardar la clave. Revisá que tengas permiso para escribir "
                "en la carpeta de configuración del usuario.",
            )
            return
        self.mostrar_principal()

    # ------------------------------------------------------------------ #
    # 2) Pantalla principal
    # ------------------------------------------------------------------ #
    def mostrar_principal(self):
        pantalla = ctk.CTkFrame(self._contenedor, fg_color=styles.COLOR_FONDO)

        ctk.CTkLabel(
            pantalla, text="Digitalizar planillas de notas",
            font=(styles.FUENTE, styles.TAM_TITULO, "bold"), text_color=styles.COLOR_TEXTO,
        ).pack(pady=(30, 10))

        ctk.CTkLabel(
            pantalla,
            text="Elegí los PDFs o las imágenes (JPG/PNG) de una misma asignatura\n"
            "(puede haber varios cursos: la app los ordena solos). Cada imagen es una planilla.",
            font=(styles.FUENTE, styles.TAM_TEXTO), text_color=styles.COLOR_TEXTO_SECUNDARIO,
            justify="center",
        ).pack(pady=(0, 30))

        boton_cargar = ctk.CTkButton(
            pantalla,
            text="📄  Elegir planillas (PDF o imágenes)",
            height=72, width=360, font=(styles.FUENTE, styles.TAM_BOTON_GRANDE, "bold"),
            fg_color=styles.COLOR_PRINCIPAL, hover_color=styles.COLOR_PRINCIPAL_HOVER,
            corner_radius=16, command=self._elegir_planillas,
        )
        boton_cargar.pack(pady=16)

        ctk.CTkLabel(
            pantalla, text="¿Cómo funciona?",
            font=(styles.FUENTE, styles.TAM_SUBTITULO, "bold"), text_color=styles.COLOR_TEXTO,
        ).pack(pady=(30, 6))

        pasos = (
            "1. Escaneá o sacale foto a las planillas de una asignatura (PDF o imágenes).\n"
            "2. Elegilas en la app.\n"
            "3. Revisá las notas que queden marcadas en amarillo.\n"
            "4. La app genera el Excel con las definitivas ya calculadas."
        )
        ctk.CTkLabel(
            pantalla, text=pasos, font=(styles.FUENTE, styles.TAM_TEXTO),
            text_color=styles.COLOR_TEXTO_SECUNDARIO, justify="center",
        ).pack(pady=(0, 30))

        ctk.CTkButton(
            pantalla, text="Cambiar la clave de Google AI",
            font=(styles.FUENTE, styles.TAM_TEXTO_CHICO), fg_color="transparent",
            text_color=styles.COLOR_TEXTO_SECUNDARIO, hover_color=styles.COLOR_FONDO_SECUNDARIO,
            command=self.mostrar_configuracion,
        ).pack(side="bottom", pady=(0, 12))

        self._cambiar_pantalla(pantalla)

    def _elegir_pdf(self):
        """Alias de compatibilidad: selecciona un solo PDF y arranca."""
        ruta = filedialog.askopenfilename(
            title="Elegí el PDF de planillas",
            filetypes=[("Archivos PDF", "*.pdf"), ("Todos los archivos", "*.*")],
        )
        if not ruta:
            return
        self._elegir_planillas(archivos_previos=[ruta])

    def _elegir_planillas(self, archivos_previos=None):
        """Selecciona una o varias planillas (PDFs y/o imágenes JPG/PNG).

        Cada imagen es una planilla; cada página de PDF es una planilla. Se
        valida cada archivo al elegir y se aborta si el total supera el límite
        del lote, antes de arrancar el procesamiento (que es caro).
        """
        if archivos_previos:
            rutas = archivos_previos
        else:
            rutas = filedialog.askopenfilenames(
                title="Elegí planillas (PDF o imágenes)",
                filetypes=[
                    ("Planillas", "*.pdf *.jpg *.jpeg *.png"),
                    ("Archivos PDF", "*.pdf"),
                    ("Imágenes", "*.jpg *.jpeg *.png"),
                ],
            )
            if not rutas:
                return
            rutas = list(rutas)

        # Validación temprana: contar las planillas (páginas + imágenes) es
        # barato y evita arrancar un lote ilegible o demasiado grande para la
        # cuota gratuita de la API.
        total = 0
        for ruta in rutas:
            try:
                if pdf_loader.es_pdf(ruta):
                    pdf_loader.validar_pdf(ruta)
                elif pdf_loader.es_imagen(ruta):
                    pdf_loader.validar_imagen(ruta)
                else:
                    raise pdf_loader.PdfError(
                        "El archivo no es un PDF ni una imagen compatible. "
                        "Usá archivos .pdf, .jpg, .jpeg o .png."
                    )
                total += pdf_loader.contar_planillas(ruta)
            except pdf_loader.PdfError as e:
                messagebox.showerror("No se pudo leer el archivo", str(e))
                return
        if total > pdf_loader.MAX_PLANILLAS_POR_LOTE:
            messagebox.showinfo(
                "Demasiadas planillas",
                f"Elegiste {total} planillas en total. Por ahora la app procesa "
                f"hasta {pdf_loader.MAX_PLANILLAS_POR_LOTE} planillas por lote para "
                "no saturar el servicio de lectura. Elegí menos planillas y volvé "
                "a intentar.",
            )
            return
        self.mostrar_progreso(rutas)

    # ------------------------------------------------------------------ #
    # 3) Pantalla de progreso
    # ------------------------------------------------------------------ #
    def mostrar_progreso(self, rutas):
        # Acepta una ruta suelta (compatibilidad) o una lista de rutas.
        if isinstance(rutas, str):
            rutas = [rutas]
        rutas = list(rutas or [])

        pantalla = ctk.CTkFrame(self._contenedor, fg_color=styles.COLOR_FONDO)

        ctk.CTkLabel(
            pantalla, text="Leyendo las planillas...",
            font=(styles.FUENTE, styles.TAM_TITULO, "bold"), text_color=styles.COLOR_TEXTO,
        ).pack(pady=(40, 12))

        self._barra = ctk.CTkProgressBar(
            pantalla, width=520, height=22, fg_color=styles.COLOR_FONDO_SECUNDARIO,
            progress_color=styles.COLOR_PRINCIPAL, corner_radius=10,
        )
        self._barra.pack(pady=16)
        self._barra.set(0)

        self._estado = ctk.CTkLabel(
            pantalla, text="Preparando...", font=(styles.FUENTE, styles.TAM_TEXTO),
            text_color=styles.COLOR_TEXTO_SECUNDARIO,
        )
        self._estado.pack(pady=8)

        self._boton_cancelar = ctk.CTkButton(
            pantalla, text="Cancelar", font=(styles.FUENTE, styles.TAM_TEXTO),
            fg_color="transparent", text_color=styles.COLOR_TEXTO_SECUNDARIO,
            hover_color=styles.COLOR_FONDO_SECUNDARIO, command=self._cancelar_proceso,
        )
        self._boton_cancelar.pack(pady=(20, 10))

        self._cambiar_pantalla(pantalla)

        # Arrancar el hilo de procesamiento con el lote de archivos.
        self._worker_cola = queue.Queue()
        self._worker = ProcesadorEnSegundoPlano(rutas, self._worker_cola)
        self._worker.iniciar()
        self.after(80, self._revisar_cola_progreso)

    def _revisar_cola_progreso(self):
        try:
            while True:
                msg = self._worker_cola.get_nowait()
                if msg["tipo"] == MSG_PROGRESO:
                    self._estado.configure(text=msg["mensaje"])
                    if "valor" in msg and msg["valor"] is not None:
                        self._barra.set(msg["valor"])
                elif msg["tipo"] == MSG_RESULTADO:
                    self.planillas = msg["planillas"]
                    self.paginas_total = msg.get("paginas_total", len(msg["planillas"]))
                    self.paginas_fallidas = msg.get("fallidas", [])
                    self.planilla_actual_idx = 0
                    if not self.planillas:
                        # S2: ninguna planilla se pudo leer -> error claro, y
                        # NO se pasa a una pantalla de revisión vacía.
                        messagebox.showerror(
                            "No se pudo leer",
                            "No se pudo leer ninguna planilla. Revisá que las "
                            "planillas estén bien escaneadas y volvé a intentar.",
                        )
                        self.mostrar_principal()
                        return
                    if self.paginas_fallidas:
                        # S2: algunas fallaron pero el resto sirve: se avisa y
                        # se continúa igual a la revisión.
                        lista = ", ".join(str(f["pagina"]) for f in self.paginas_fallidas)
                        messagebox.showwarning(
                            "Algunas planillas no se leyeron",
                            f"Se generaron {len(self.planillas)} de {self.paginas_total} "
                            f"planillas. No se pudieron leer: {lista}. "
                            "Podés revisar las que sí se leyeron, o escanear de nuevo "
                            "las que fallaron.",
                        )
                    self.mostrar_revision()
                    return
                elif msg["tipo"] == MSG_ERROR:
                    messagebox.showerror("No se pudo completar", msg["mensaje"])
                    self.mostrar_principal()
                    return
                elif msg["tipo"] == MSG_CANCELADO:
                    # La cancelación es un estado esperado (W4), no una falla:
                    # se informa con un mensaje amigable, no con un error.
                    messagebox.showinfo(
                        "Proceso cancelado",
                        "Proceso cancelado. Podés volver a cargar las planillas cuando quieras.",
                    )
                    self.mostrar_principal()
                    return
        except queue.Empty:
            pass
        # S9: si la ventana ya fue destruida (la usuaria cerró con la X mientras
        # el worker seguía), no programar más afters sobre widgets muertos.
        try:
            if not self.winfo_exists():
                return
        except tk.TclError:
            return
        # Seguir revisando mientras el hilo siga vivo.
        if self._worker is not None and getattr(self._worker, "_hilo", None) is not None:
            self.after(80, self._revisar_cola_progreso)

    def _cancelar_proceso(self):
        if self._worker:
            self._worker.cancelar()
        self._estado.configure(text="Cancelando...")

    # ------------------------------------------------------------------ #
    # 4) Pantalla de revisión
    # ------------------------------------------------------------------ #
    def mostrar_revision(self):
        pantalla = ctk.CTkFrame(self._contenedor, fg_color=styles.COLOR_FONDO)

        ctk.CTkLabel(
            pantalla, text="Revisar lo que se leyó",
            font=(styles.FUENTE, styles.TAM_TITULO, "bold"), text_color=styles.COLOR_TEXTO,
        ).pack(pady=(20, 4))

        ctk.CTkLabel(
            pantalla,
            text="Las celdas en amarillo son notas dudadas. Corregilas si hace falta "
            "y después generá el Excel.",
            font=(styles.FUENTE, styles.TAM_TEXTO_CHICO), text_color=styles.COLOR_TEXTO_SECUNDARIO,
            wraplength=640,
        ).pack(pady=(0, 12))

        # Contenedor que recorre los cursos.
        self._frame_cursos = ctk.CTkScrollableFrame(
            pantalla, fg_color="transparent", width=780, height=420,
        )
        self._frame_cursos.pack(fill="both", expand=True, pady=(0, 12))

        self._render_cursos()

        boton_guardar = ctk.CTkButton(
            pantalla, text="Continuar →", height=52,
            font=(styles.FUENTE, styles.TAM_BOTON_GRANDE, "bold"),
            fg_color=styles.COLOR_PRINCIPAL, hover_color=styles.COLOR_PRINCIPAL_HOVER,
            command=self.mostrar_configuracion_calculo,
        )
        boton_guardar.pack(pady=(0, 6))

        ctk.CTkLabel(
            pantalla, text="Amarillo = nota dudosa, verificarla en la planilla física.\n"
            "Después de revisar, elegís qué columnas entran al promedio y cómo se calcula.",
            font=(styles.FUENTE, styles.TAM_TEXTO_CHICO), text_color=styles.COLOR_REVISAR_BORDE,
            wraplength=640,
        ).pack(pady=(0, 10))

        self._cambiar_pantalla(pantalla)

    def _render_cursos(self):
        # Limpiar el frame por si se vuelve a entrar.
        for w in self._frame_cursos.winfo_children():
            w.destroy()

        # S8: la agrupación por curso y la combinación de estudiantes viven en
        # excel/agrupacion.py, la MISMA fuente de verdad que usa el generador
        # de Excel: lo que se muestra acá es exactamente lo que se escribe.
        # S11: combinar_estudiantes descarta estudiantes duplicados (páginas
        # repetidas o solapadas) antes de mostrarlos.
        por_curso, orden = agrupar_por_curso(self.planillas)
        self._editores = {}  # (curso, idx_fila, celda_idx) -> variable StringVar
        self._periodo_combos = {}  # curso -> CTkComboBox de corrección de periodo

        for clave in orden:
            # S8: cada clave es una tupla (grupo, asignatura): una tarjeta por
            # asignatura+curso, para que asignaturas distintas de un mismo curso
            # no se mezclen y no se pierdan notas.
            grupo, _asignatura = clave
            paginas = por_curso[clave]
            # Combinar estudiantes de páginas del mismo curso (con dedupe S11).
            estudiantes = combinar_estudiantes(paginas)
            enc = paginas[0]["encabezado"]

            # Ancho REAL de la zona de área de trabajo (spec v2: 1 a 16 notas):
            # se calcula con la MISMA fuente de verdad que usa el generador de
            # Excel (S8) — lo que se muestra acá es exactamente lo que se escribe.
            n_areas = generar_excel_notas.calcular_n_areas(enc, estudiantes)
            # Achicar el ancho de las celdas para que entren hasta 16 columnas:
            # 640//n_areas reparte el espacio horizontal disponible.
            ancho_celda = max(48, min(100, 640 // n_areas))

            tarjeta = ctk.CTkFrame(
                self._frame_cursos, fg_color=styles.COLOR_BLANCO, corner_radius=12,
                border_width=1, border_color="#E3E9F5",
            )
            tarjeta.pack(fill="x", pady=8, padx=4)

            titulo = f"Curso {grupo} — Periodo {enc.get('periodo')}"
            if enc.get("grupo_erroneo"):
                titulo += "  ⚠ (curso no reconocido, verificá el número)"
            if enc.get("periodo_erroneo"):
                titulo += "  ⚠ (periodo no reconocido, verificá el número)"
            ctk.CTkLabel(
                tarjeta, text=titulo, font=(styles.FUENTE, styles.TAM_SUBTITULO, "bold"),
                text_color=styles.COLOR_TEXTO,
            ).pack(anchor="w", padx=14, pady=(10, 2))

            # Advertencia visible si el periodo o el grupo parecen incorrectos
            # (W1): texto naranja/rojo para que la usuaria lo verifique.
            if enc.get("periodo_erroneo") or enc.get("grupo_erroneo"):
                ctk.CTkLabel(
                    tarjeta,
                    text="⚠ El periodo o el grupo de esta planilla parece incorrecto, "
                         "verificá contra el papel.",
                    font=(styles.FUENTE, styles.TAM_TEXTO_CHICO, "bold"),
                    text_color="#C0392B", wraplength=620, anchor="w",
                ).pack(anchor="w", padx=14, pady=(0, 6))

            # W-A Part 2: si el periodo llegó inválido, ofrecer corregirlo acá
            # mismo, en la pantalla de revisión (antes no había forma de hacerlo).
            if enc.get("periodo_erroneo"):
                fila_periodo = ctk.CTkFrame(tarjeta, fg_color="transparent")
                fila_periodo.pack(anchor="w", padx=14, pady=(0, 6))
                ctk.CTkLabel(
                    fila_periodo, text="Elegí el periodo correcto:",
                    font=(styles.FUENTE, styles.TAM_TEXTO_CHICO),
                    text_color=styles.COLOR_TEXTO, anchor="w",
                ).pack(side="left", padx=(0, 8))
                combo = ctk.CTkComboBox(
                    fila_periodo, values=["1", "2", "3", "4"], width=90,
                    state="normal",
                    font=(styles.FUENTE, styles.TAM_TEXTO_CHICO),
                )
                combo.pack(side="left")
                self._periodo_combos[clave] = combo

            sub = f"{enc.get('asignatura','')}  •  {enc.get('docente','')}"
            ctk.CTkLabel(
                tarjeta, text=sub, font=(styles.FUENTE, styles.TAM_TEXTO_CHICO),
                text_color=styles.COLOR_TEXTO_SECUNDARIO,
            ).pack(anchor="w", padx=14, pady=(0, 6))

            # S3: aviso si alguna nota de Ev. Anteriores del curso quedó marcada
            # (fuera de rango o dígitos ambiguos). Las celdas de Ev. Anteriores no
            # son editables acá, así que la usuaria debe verificarlas contra el
            # papel antes de generar el Excel.
            if any(est.get("revisar_ev") for est in estudiantes):
                ctk.CTkLabel(
                    tarjeta,
                    text="Hay notas de Ev. Anteriores que parecen fuera de rango en "
                         "este curso — verificá los valores antes de generar.",
                    font=(styles.FUENTE, styles.TAM_TEXTO_CHICO, "bold"),
                    text_color="#B8860B", wraplength=620, anchor="w",
                ).pack(anchor="w", padx=14, pady=(0, 6))

            # Spec v2: aviso si alguna planilla del curso quedó marcada para
            # revisión manual (cantidad de notas o de alumnos no coincide con lo
            # esperado). No bloquea el flujo: la usuaria verifica contra el papel
            # y puede generar el Excel igual.
            if any(p.get("revisar_planilla") for p in paginas):
                ctk.CTkLabel(
                    tarjeta,
                    text=f"Curso {grupo} requiere revisión manual: la cantidad de notas "
                         "o de alumnos no coincide. Verificá contra la planilla física.",
                    font=(styles.FUENTE, styles.TAM_TEXTO_CHICO, "bold"),
                    text_color="#C0392B", wraplength=620, anchor="w",
                ).pack(anchor="w", padx=14, pady=(0, 6))

            # Encabezados de columnas.
            cabecera = ctk.CTkFrame(tarjeta, fg_color="#EDF2FB", corner_radius=8)
            cabecera.pack(fill="x", padx=10)
            ctk.CTkLabel(cabecera, text="No.", font=(styles.FUENTE, styles.TAM_TEXTO_CHICO, "bold"),
                         text_color=styles.COLOR_TEXTO, width=44).pack(side="left", padx=(10, 4), pady=6)
            ctk.CTkLabel(cabecera, text="Nombre del Alumno", font=(styles.FUENTE, styles.TAM_TEXTO_CHICO, "bold"),
                         text_color=styles.COLOR_TEXTO, width=280).pack(side="left", pady=6)
            for k in range(1, n_areas + 1):
                ctk.CTkLabel(cabecera, text=f"Área Trabajo {k}", font=(styles.FUENTE, styles.TAM_TEXTO_CHICO, "bold"),
                             text_color=styles.COLOR_TEXTO, width=110).pack(side="left", pady=6)

            for idx, est in enumerate(estudiantes):
                fila = ctk.CTkFrame(tarjeta, fg_color="transparent")
                fila.pack(fill="x", padx=10, pady=1)

                if est.get("retirado"):
                    ctk.CTkLabel(
                        fila, text=str(est.get("no", "")), font=(styles.FUENTE, styles.TAM_CELDA),
                        text_color=styles.COLOR_TEXTO, width=44,
                    ).pack(side="left", padx=(10, 4))
                    ctk.CTkLabel(
                        fila, text="👤 Retirado — " + (est.get("nombre") or ""),
                        font=(styles.FUENTE, styles.TAM_CELDA), text_color=styles.COLOR_TEXTO_SECUNDARIO,
                        width=280, anchor="w",
                    ).pack(side="left")
                    continue

                at = est.get("area_trabajo") or []
                rev = est.get("revisar") or []

                ctk.CTkLabel(
                    fila, text=str(est.get("no", "")), font=(styles.FUENTE, styles.TAM_CELDA),
                    text_color=styles.COLOR_TEXTO, width=44,
                ).pack(side="left", padx=(10, 4))

                ctk.CTkLabel(
                    fila, text=est.get("nombre", ""), font=(styles.FUENTE, styles.TAM_CELDA),
                    text_color=styles.COLOR_TEXTO, width=280, anchor="w",
                ).pack(side="left")

                for k in range(n_areas):
                    var = tk.StringVar(value=_fmt_celda(at[k] if at and k < len(at) else None))
                    self._editores[(clave, idx, k)] = var
                    flag_rev = rev[k] if rev and k < len(rev) else False
                    color_fondo = styles.COLOR_REVISAR if flag_rev else styles.COLOR_FONDO_SECUNDARIO
                    entrada = ctk.CTkEntry(
                        fila, textvariable=var, width=ancho_celda, height=30,
                        font=(styles.FUENTE, styles.TAM_CELDA),
                        fg_color=color_fondo,
                        border_color=styles.COLOR_REVISAR_BORDE if flag_rev else "#D5DEEF",
                    )
                    _clear_tooltip(entrada, flag_rev)
                    entrada.pack(side="left", padx=6, pady=3)

    # ------------------------------------------------------------------ #
    # 4b) Pantalla de configuración de cálculo
    # ------------------------------------------------------------------ #
    def mostrar_configuracion_calculo(self):
        """
        Pantalla donde la usuaria define CÓMO se calcula la nota definitiva:
        qué columnas de Área de Trabajo entran al promedio y si es un promedio
        simple o con pesos (porcentajes) distintos por columna.

        La configuración es POR FORMA (cantidad de columnas de notas): cada
        planilla se agrupa por su cantidad de columnas, y la usuaria elige UNA
        vez la config para todas las planillas de esa forma. Junto al formulario
        se muestra la imagen (escalada) de la primera planilla de la forma,
        clickeable para ver la planilla completa.

        El formulario de cada forma se reconstruye dinámicamente según el modo:
        - simple:  checkboxes para marcar qué columnas entran.
        - pesos:   checkboxes + campo de porcentaje por columna.
        """
        # Aplicar correcciones manuales antes de configurar (lo que se muestra
        # en el resumen es lo que se escribirá en el Excel).
        self._aplicar_periodo_seleccionado()
        self._aplicar_ediciones()

        por_curso, orden = agrupar_por_curso(self.planillas)
        if not orden:
            # Sin planillas no hay nada que configurar: volver a inicio.
            self.mostrar_principal()
            return

        # Agrupar las planillas POR FORMA (cantidad de columnas de notas).
        por_forma, orden_formas = agrupar_por_forma(self.planillas)
        if not orden_formas:
            self.mostrar_principal()
            return

        dpi = app_config.load_config().get("preferencias", {}).get("dpi_pdf") or 250

        pantalla = ctk.CTkFrame(self._contenedor, fg_color=styles.COLOR_FONDO)

        ctk.CTkLabel(
            pantalla, text="¿Cómo se calcula la nota?",
            font=(styles.FUENTE, styles.TAM_TITULO, "bold"), text_color=styles.COLOR_TEXTO,
        ).pack(pady=(20, 4))

        ctk.CTkLabel(
            pantalla,
            text="Elegí qué notas entran a la definitiva del periodo. Cada tipo de "
            "planilla (según la cantidad de columnas) se configura por separado.\n"
            "Hacé clic en la imagen para ver la planilla completa.",
            font=(styles.FUENTE, styles.TAM_TEXTO_CHICO), text_color=styles.COLOR_TEXTO_SECUNDARIO,
            justify="center", wraplength=720,
        ).pack(pady=(0, 10))

        # ── Contenedor scrolleable con un frame interior que empaqueta cada
        #    forma. Si hay muchas formas (planillas de distinta cantidad de
        #    columnas), la pantalla scrollea en vez de desbordar.
        self._frame_formas = ctk.CTkScrollableFrame(
            pantalla, fg_color="transparent", width=880, height=400,
        )
        self._frame_formas.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        interior = ctk.CTkFrame(self._frame_formas, fg_color="transparent")
        interior.pack(fill="both", expand=True, padx=4, pady=4)
        self._interior_formas = interior

        # Resetear el estado de widgets por forma.
        self._var_modo = {}        # forma -> StringVar
        self._chk_col = {}         # forma -> {índice: BooleanVar}
        self._chk_col_frames = {}  # forma -> frame del formulario
        self._peso_col = {}        # forma -> {índice: StringVar}
        self._col_rows = {}        # forma -> {índice: frame de fila}
        self._label_pesos = {}     # forma -> CTkLabel de resumen de pesos

        # Asegurar una config por defecto para cada forma.
        for clave_forma in orden_formas:
            n = clave_forma[1]
            if clave_forma not in self._column_configs:
                self._column_configs[clave_forma] = ColumnConfig.crear_desde_planilla(n)

        # Para cada forma: un CTkFrame con dos columnas (formulario + imagen).
        for i, clave_forma in enumerate(orden_formas, start=1):
            n = clave_forma[1]
            planillas_forma = por_forma[clave_forma]

            frame_forma = ctk.CTkFrame(
                interior, fg_color=styles.COLOR_BLANCO, corner_radius=12,
                border_width=1, border_color="#E3E9F5",
            )
            frame_forma.pack(fill="x", pady=8, padx=2)

            ctk.CTkLabel(
                frame_forma,
                text=self._forma_titulo(i, n, len(planillas_forma)),
                font=(styles.FUENTE, styles.TAM_SUBTITULO, "bold"),
                text_color=styles.COLOR_TEXTO,
            ).pack(anchor="w", padx=14, pady=(10, 2))

            # Dos columnas: formulario a la izquierda, imagen a la derecha.
            dos_col = ctk.CTkFrame(frame_forma, fg_color="transparent")
            dos_col.pack(fill="x", padx=10, pady=(0, 8))
            dos_col.grid_columnconfigure(0, weight=1)

            col_izq = ctk.CTkFrame(dos_col, fg_color="transparent")
            col_izq.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

            col_der = ctk.CTkFrame(dos_col, fg_color="transparent")
            col_der.grid(row=0, column=1, sticky="n", padx=(8, 0))

            # ── Modo de cálculo (por forma) ──
            self._var_modo[clave_forma] = tk.StringVar(
                value=self._column_configs[clave_forma].modo
            )
            fila_modo = ctk.CTkFrame(col_izq, fg_color="transparent")
            fila_modo.pack(anchor="w", pady=(0, 4))
            ctk.CTkRadioButton(
                fila_modo, text="Promedio simple (todas las notas valen lo mismo)",
                variable=self._var_modo[clave_forma], value="simple",
                command=lambda f=clave_forma: self._on_modo_changed(f),
                font=(styles.FUENTE, styles.TAM_TEXTO_CHICO), text_color=styles.COLOR_TEXTO,
            ).pack(anchor="w", pady=2)
            ctk.CTkRadioButton(
                fila_modo, text="Cada nota tiene su porcentaje (pesos)",
                variable=self._var_modo[clave_forma], value="pesos",
                command=lambda f=clave_forma: self._on_modo_changed(f),
                font=(styles.FUENTE, styles.TAM_TEXTO_CHICO), text_color=styles.COLOR_TEXTO,
            ).pack(anchor="w", pady=2)

            # ── Formulario de columnas (por forma) ──
            frame_col = ctk.CTkFrame(col_izq, fg_color="transparent")
            frame_col.pack(fill="both", expand=True, pady=(2, 0))
            self._chk_col_frames[clave_forma] = frame_col

            # ── Aviso de pesos (por forma) ──
            label_pesos = ctk.CTkLabel(
                frame_forma, text="", font=(styles.FUENTE, styles.TAM_TEXTO_CHICO),
                text_color=styles.COLOR_TEXTO_SECUNDARIO, wraplength=660,
            )
            label_pesos.pack(anchor="w", padx=14, pady=(0, 4))
            self._label_pesos[clave_forma] = label_pesos

            # ── Imagen de la planilla (primera de la forma), clickeable ──
            primera = planillas_forma[0]
            self._render_imagen_forma(col_der, primera, i, dpi)

            self._reconstruir_formulario_columnas(clave_forma, n)

        # ── Botones ──
        fila_botones = ctk.CTkFrame(pantalla, fg_color="transparent")
        fila_botones.pack(pady=(0, 10))

        ctk.CTkButton(
            fila_botones, text="Volver a la revisión", height=44, width=180,
            font=(styles.FUENTE, styles.TAM_TEXTO, "bold"),
            fg_color="transparent", text_color=styles.COLOR_TEXTO_SECUNDARIO,
            border_width=1, border_color="#C6D2E8",
            hover_color=styles.COLOR_FONDO_SECUNDARIO,
            command=self.mostrar_revision,
        ).pack(side="left", padx=8)

        ctk.CTkButton(
            fila_botones, text="Generar Excel", height=52, width=260,
            font=(styles.FUENTE, styles.TAM_BOTON_GRANDE, "bold"),
            fg_color=styles.COLOR_VERDE, hover_color="#5AA87A",
            command=self._guardar_config_y_generar,
        ).pack(side="left", padx=8)

        self._cambiar_pantalla(pantalla)

    @staticmethod
    def _forma_titulo(indice, n, x):
        """Título de una forma: 'Plantilla tipo i · n columnas de notas · x planillas'."""
        return f"Plantilla tipo {indice} · {n} columnas de notas · {x} planillas"

    def _render_imagen_forma(self, contenedor, planilla, indice, dpi):
        """Renderiza la imagen de la planilla (escalada a ~460px) y la hace
        clickeable para abrir la ventana con la planilla completa."""
        imagen = pdf_loader.render_imagen_planilla(planilla, dpi)
        if imagen is None:
            ctk.CTkLabel(
                contenedor, text="Sin imagen",
                font=(styles.FUENTE, styles.TAM_TEXTO_CHICO),
                text_color=styles.COLOR_TEXTO_SECUNDARIO,
                width=200, height=120,
            ).pack(pady=4)
            return

        try:
            ancho_objetivo = 460
            ancho, alto = imagen.size
            if ancho > 0:
                escala = ancho_objetivo / ancho
            else:
                escala = 1.0
            alto_esc = max(1, int(alto * escala))
            ctk_img = ctk.CTkImage(light_image=imagen, dark_image=imagen,
                                   size=(ancho_objetivo, alto_esc))
        except Exception:
            # Imagen ilegible: mostrar el placeholder en vez de romper la pantalla.
            imagen.close()
            ctk.CTkLabel(
                contenedor, text="Sin imagen",
                font=(styles.FUENTE, styles.TAM_TEXTO_CHICO),
                text_color=styles.COLOR_TEXTO_SECUNDARIO,
                width=200, height=120,
            ).pack(pady=4)
            return

        # OJO: NO cerrar `imagen` acá: CTkImage la usa (resize) recién cuando se
        # renderiza el botón. La referencia queda en manos del CTkImage/CtkButton,
        # que se libera al destruir la pantalla.

        # Botón con la imagen como contenido (fg transparent) -> clickeable.
        boton = ctk.CTkButton(
            contenedor, text="", image=ctk_img, width=ancho_objetivo,
            height=alto_esc, fg_color="transparent", hover_color=styles.COLOR_FONDO_SECUNDARIO,
            command=lambda p=planilla: self._ver_planilla_grande(p),
        )
        boton.pack(pady=(0, 4))
        ctk.CTkLabel(
            contenedor, text=f"Clic para ver la planilla {indice} completa",
            font=(styles.FUENTE, styles.TAM_TEXTO_CHICO),
            text_color=styles.COLOR_TEXTO_SECUNDARIO,
        ).pack(pady=(0, 4))

    def _ver_planilla_grande(self, planilla):
        """Abre una ventana con la planilla completa (imagen a tamaño natural)
        dentro de un Canvas con scrollbars si es más grande que la pantalla."""
        imagen = pdf_loader.render_imagen_planilla(
            planilla,
            int(app_config.load_config().get("preferencias", {}).get("dpi_pdf") or 250),
        )
        if imagen is None:
            messagebox.showinfo("Planilla completa", "No se pudo obtener la imagen de esta planilla.")
            return

        ventana = ctk.CTkToplevel(self)
        ventana.title("Planilla completa")
        ventana.geometry("880x680")
        ventana.minsize(500, 400)

        ancho, alto = imagen.size
        # Para el Canvas se usa un PhotoImage de PIL (CTkImage sirve para los
        # widgets de customtkinter, no para tk.Canvas).
        from PIL import ImageTk
        foto = ImageTk.PhotoImage(imagen)
        imagen.close()

        # Canvas + scrollbars para cuando la imagen es más grande que la ventana.
        lienzo = tk.Canvas(
            ventana, highlightthickness=0, bg="#FFFFFF",
            scrollregion=(0, 0, ancho, alto),
        )
        barra_v = ttk.Scrollbar(ventana, orient="vertical", command=lienzo.yview)
        barra_h = ttk.Scrollbar(ventana, orient="horizontal", command=lienzo.xview)
        lienzo.configure(yscrollcommand=barra_v.set, xscrollcommand=barra_h.set)

        lienzo.grid(row=0, column=0, sticky="nsew")
        barra_v.grid(row=0, column=1, sticky="ns")
        barra_h.grid(row=1, column=0, sticky="ew")
        ventana.grid_rowconfigure(0, weight=1)
        ventana.grid_columnconfigure(0, weight=1)

        lienzo.create_image(0, 0, anchor="nw", image=foto)
        # Referencia para que no se recolecte la imagen.
        lienzo._img_ref = foto

        fila_btn = ctk.CTkFrame(ventana, fg_color="transparent")
        fila_btn.grid(row=2, column=0, columnspan=2, pady=(8, 10))
        ctk.CTkButton(
            fila_btn, text="Cerrar", height=38, width=120,
            font=(styles.FUENTE, styles.TAM_TEXTO), fg_color=styles.COLOR_PRINCIPAL,
            hover_color=styles.COLOR_PRINCIPAL_HOVER, command=ventana.destroy,
        ).pack()

    def _reconstruir_formulario_columnas(self, clave_forma, n_areas):
        """Reconstruye el formulario de columnas de UNA forma según el modo
        elegido para esa forma. Los estados de widgets son POR FORMA."""
        frame_col = self._chk_col_frames[clave_forma]
        for w in frame_col.winfo_children():
            w.destroy()

        config = self._column_configs[clave_forma]
        modo = self._var_modo[clave_forma].get()
        self._chk_col[clave_forma] = {}   # índice -> BooleanVar
        self._peso_col[clave_forma] = {}  # índice -> StringVar
        self._col_rows[clave_forma] = {}  # índice -> frame de fila

        if not config.columnas:
            ctk.CTkLabel(
                frame_col,
                text="No se detectaron columnas de notas en esta planilla.",
                font=(styles.FUENTE, styles.TAM_TEXTO_CHICO), text_color=styles.COLOR_TEXTO_SECUNDARIO,
            ).pack(pady=12)
            return

        # Cabecera
        cabecera = ctk.CTkFrame(frame_col, fg_color="#EDF2FB", corner_radius=8)
        cabecera.pack(fill="x", pady=(2, 4))
        ctk.CTkLabel(cabecera, text="Incluir", font=(styles.FUENTE, styles.TAM_TEXTO_CHICO, "bold"),
                     text_color=styles.COLOR_TEXTO, width=60).pack(side="left", padx=(10, 4), pady=5)
        ctk.CTkLabel(cabecera, text="Columna de notas", font=(styles.FUENTE, styles.TAM_TEXTO_CHICO, "bold"),
                     text_color=styles.COLOR_TEXTO, width=190, anchor="w").pack(side="left", pady=5)
        if modo == "pesos":
            ctk.CTkLabel(cabecera, text="Porcentaje (%)", font=(styles.FUENTE, styles.TAM_TEXTO_CHICO, "bold"),
                         text_color=styles.COLOR_TEXTO, width=110).pack(side="left", pady=5)

        for i, col in enumerate(config.columnas):
            fila = ctk.CTkFrame(frame_col, fg_color="transparent")
            fila.pack(fill="x", pady=2)
            self._col_rows[clave_forma][i] = fila

            var_chk = tk.BooleanVar(value=col.get("incluida", True))
            self._chk_col[clave_forma][i] = var_chk
            ctk.CTkCheckBox(
                fila, text="", variable=var_chk, width=50,
                command=lambda f=clave_forma, idx=i: self._on_toggle_col(f, idx),
                checkbox_width=22, checkbox_height=22,
            ).pack(side="left", padx=(10, 4))

            nombre = col.get("nombre") or f"Área Trabajo {i+1}"
            ctk.CTkLabel(
                fila, text=nombre, font=(styles.FUENTE, styles.TAM_TEXTO_CHICO),
                text_color=styles.COLOR_TEXTO, width=185, anchor="w",
            ).pack(side="left")

            if modo == "pesos":
                var_peso = tk.StringVar(value=_fmt_peso(col.get("peso")))
                self._peso_col[clave_forma][i] = var_peso
                entrada = ctk.CTkEntry(
                    fila, textvariable=var_peso, width=100, height=28,
                    font=(styles.FUENTE, styles.TAM_TEXTO_CHICO),
                    fg_color=styles.COLOR_FONDO_SECUNDARIO,
                    border_color="#D5DEEF",
                )
                entrada.pack(side="left", padx=(10, 0))
                estado = "normal" if col.get("incluida", True) else "disabled"
                entrada.configure(state=estado)
            else:
                # En simple no se muestran pesos
                ctk.CTkLabel(fila, text="", width=110).pack(side="left")

        self._actualizar_resumen_pesos(clave_forma)

    def _on_modo_changed(self, clave_forma):
        """Reconstruye el formulario de UNA forma cuando cambia simple ↔ pesos."""
        n_areas = clave_forma[1]
        self._reconstruir_formulario_columnas(clave_forma, n_areas)

    def _on_toggle_col(self, clave_forma, idx):
        """Habilita/deshabilita el campo de peso cuando se marca una columna."""
        if idx in self._peso_col[clave_forma] and idx in self._col_rows[clave_forma]:
            # Buscar la entrada dentro de la fila (es el único CTkEntry)
            fila = self._col_rows[clave_forma][idx]
            for w in fila.winfo_children():
                if isinstance(w, ctk.CTkEntry):
                    w.configure(
                        state="normal" if self._chk_col[clave_forma][idx].get() else "disabled"
                    )
        self._actualizar_resumen_pesos(clave_forma)

    def _actualizar_resumen_pesos(self, clave_forma):
        """Muestra la suma actual de porcentajes de UNA forma (y si no llega a 100)."""
        modo = self._var_modo[clave_forma].get()
        label = self._label_pesos[clave_forma]
        if modo != "pesos":
            label.configure(text="")
            return
        incluidas = [
            i for i, var in self._chk_col[clave_forma].items() if var.get()
        ]
        if not incluidas:
            label.configure(
                text="Marcá al menos una columna para calcular la nota.",
                text_color=styles.COLOR_ROJO,
            )
            return
        total = 0.0
        for i in incluidas:
            try:
                total += float((self._peso_col[clave_forma][i].get() or "").replace(",", "."))
            except ValueError:
                pass
        if abs(total - 100.0) < 0.01:
            texto = f"Suma de porcentajes: {total:.1f}% ✓"
            color = styles.COLOR_TEXTO_SECUNDARIO
        else:
            texto = f"Suma de porcentajes: {total:.1f}% (debe sumar 100%)"
            color = styles.COLOR_ROJO if incluidas else styles.COLOR_TEXTO_SECUNDARIO
        label.configure(text=texto, text_color=color)

    def _guardar_config_y_generar(self):
        """Valida la configuración y, si es correcta, pasa a generar el Excel."""
        if not self._validar_configuracion():
            return
        self._generar_excel()

    def _validar_configuracion(self) -> bool:
        """
        Lee el formulario de CADA forma, actualiza `self._column_configs` y valida:

        - Al menos una columna marcada por forma.
        - En modo pesos: cada columna marcada tiene un porcentaje numérico
          >= 0, y la suma da ≈ 100 (con tolerancia 0.01).
        Devuelve True si todo está OK; si no, muestra un error amigable y
        devuelve False.
        """
        por_forma, orden_formas = agrupar_por_forma(self.planillas)
        if not orden_formas:
            return False

        for clave_forma in orden_formas:
            n_areas = clave_forma[1]
            config_original = self._column_configs.get(
                clave_forma, ColumnConfig.crear_desde_planilla(n_areas)
            )

            modo = self._var_modo[clave_forma].get()
            chk = self._chk_col[clave_forma]
            peso = self._peso_col[clave_forma]

            columnas = []
            for i in range(n_areas):
                col_original = (
                    config_original.columnas[i]
                    if i < len(config_original.columnas)
                    else {}
                )
                incl = chk[i].get() if i in chk else col_original.get("incluida", True)
                valor_peso = col_original.get("peso", 0)
                if i in peso:
                    texto = (peso[i].get() or "").strip()
                    if texto:
                        try:
                            valor_peso = float(texto.replace(",", "."))
                        except ValueError:
                            messagebox.showerror(
                                "Porcentaje inválido",
                                f"El porcentaje de la columna \"{col_original.get('nombre', f'Área Trabajo {i+1}')}\" "
                                "no es un número válido. Usá punto o coma para decimales (ej. 33.3).",
                            )
                            return False
                columnas.append({
                    "nombre": col_original.get("nombre") or f"Área Trabajo {i+1}",
                    "incluida": bool(incl),
                    "peso": valor_peso,
                })

            config_nueva = ColumnConfig(modo=modo, columnas=columnas)

            # ── Validaciones por forma ──
            seleccionadas = [c for c in columnas if c["incluida"]]
            if not seleccionadas:
                messagebox.showerror(
                    "Falta elegir columnas",
                    "En la forma de " + _n_descripcion(n_areas) +
                    " tenés que marcar al menos una columna de notas para calcular la definitiva.",
                )
                return False

            if modo == "pesos":
                for c in seleccionadas:
                    if c["peso"] < 0:
                        messagebox.showerror(
                            "Porcentaje inválido",
                            f"El porcentaje de \"{c['nombre']}\" no puede ser negativo.",
                        )
                        return False
                if not config_nueva.pesos_suman_cien():
                    messagebox.showerror(
                        "Los porcentajes no suman 100",
                        f"En la forma de {_n_descripcion(n_areas)} los porcentajes suman "
                        f"{sum(c['peso'] for c in seleccionadas):.1f}% y deben sumar 100%. "
                        "Revisá los valores e intentá de nuevo.",
                    )
                    return False

            self._column_configs[clave_forma] = config_nueva

        return True

    def _generar_excel(self):
        # Aplicar las correcciones manuales de la pantalla de revisión.
        self._aplicar_periodo_seleccionado()
        self._aplicar_ediciones()

        # Elegir dónde guardar (el diálogo pregunta antes de sobrescribir).
        ruta = filedialog.asksaveasfilename(
            title="Guardar el Excel de notas",
            defaultextension=".xlsx",
            filetypes=[("Libro de Excel", "*.xlsx")],
            initialdir=app_config.default_output_dir(),
            initialfile=self._nombre_archivo_sugerido(),
        )
        if not ruta:
            return  # la usuaria canceló

        # W-B: guardar la carpeta de preferencia es accesorio, NUNCA debe
        # abortar la generación del Excel. Si falla la config, seguimos igual
        # con la ruta elegida.
        try:
            app_config.set_output_dir(os.path.dirname(ruta))
        except Exception as e:
            app_config.escribir_log(f"No se pudo guardar la carpeta de salida: {e!r}")

        try:
            generar_excel_notas.generar_excel_asignatura(
                self.planillas, ruta, column_configs=self._column_configs
            )
        except ValueError as e:
            # W-A Part 1: periodo inválido detectado por el generador (guard).
            # Mensaje amigable nombrando el curso, sin traceback.
            app_config.escribir_log(f"Error generando el Excel: {e!r}")
            messagebox.showerror(
                "Periodo inválido",
                str(e) + "\n\nElegí el periodo correcto en la pantalla de "
                "revisión y volvé a generar.",
            )
            return
        except Exception as e:
            app_config.escribir_log(f"Error generando el Excel: {e!r}")
            messagebox.showerror(
                "No se pudo generar",
                "No se pudo crear el archivo Excel. Revisá que la carpeta elegida "
                "esté disponible y volvé a intentarlo.",
            )
            return

        self.archivo_final = ruta
        self.mostrar_final()

    def _aplicar_periodo_seleccionado(self):
        """W-A Part 2: vuelca el periodo elegido en los combobox de revisión a
        las planillas. Si la usuaria eligió un valor válido, se corrige el
        periodo y se quita el flag de error; si no eligió nada, se deja igual y
        el guard del generador (ValueError) atrapa el periodo inválido."""
        if not getattr(self, "_periodo_combos", None):
            return
        por_curso, _ = agrupar_por_curso(self.planillas)
        for curso, combo in self._periodo_combos.items():
            valor = (combo.get() or "").strip()
            if valor not in ("1", "2", "3", "4"):
                continue  # sin elección válida: lo atrapa el guard del generador
            for p in por_curso.get(curso, []):
                enc = p["encabezado"]
                enc["periodo"] = int(valor)
                enc.pop("periodo_erroneo", None)

    def _aplicar_ediciones(self):
        """Vuelca los valores editados de la pantalla de revisión a las planillas."""
        por_curso, orden = agrupar_por_curso(self.planillas)
        for curso in orden:
            paginas = por_curso[curso]
            # El mismo combinar_estudiantes (con dedupe S11) que usó la
            # pantalla: los índices de fila coinciden uno a uno.
            estudiantes = combinar_estudiantes(paginas)
            enc = paginas[0]["encabezado"]
            # Mismo cálculo que la pantalla de revisión y el generador (S8):
            # el ancho de edición coincide con el ancho que se muestra/escribe.
            n_areas = generar_excel_notas.calcular_n_areas(enc, estudiantes)
            for idx, est in enumerate(estudiantes):
                if est.get("retirado"):
                    continue
                area = list(est.get("area_trabajo") or [])
                revisar = list(est.get("revisar") or [])
                # Padding hasta el ancho calculado (celdas vacías al final). Si
                # la lista fuera MÁS larga, no se trunca: se conserva todo y el
                # generador decide su propio ancho (defensivo).
                while len(area) < n_areas:
                    area.append(None)
                while len(revisar) < n_areas:
                    revisar.append(False)
                for k in range(n_areas):
                    var = self._editores.get((curso, idx, k))
                    if var is None:
                        continue
                    valor_texto = (var.get() or "").strip()
                    area[k] = _parse_celda(valor_texto)
                    # Si la usuaria borró o corrigió un valor dudoso, se quita el
                    # resaltado sólo cuando ya no es dudoso por rango.
                    if valor_texto:
                        try:
                            num = float(valor_texto.replace(",", "."))
                            if 0 <= num <= 100:
                                revisar[k] = False
                        except ValueError:
                            pass
                    else:
                        revisar[k] = False
                est["area_trabajo"] = area
                est["revisar"] = revisar
                # S3: el flag de Ev. Anteriores se preserva tal cual al editarla;
                # no se recalcula ni se borra (no hay celdas editables para ev).
                est["revisar_ev"] = bool(est.get("revisar_ev"))

    def _nombre_archivo_sugerido(self):
        # S8: si el PDF trae más de una asignatura distinta (normalizada), el
        # nombre genérico evita atribuirle un solo nombre a varias planillas.
        asignaturas = {
            _asignatura_limpia(p.get("encabezado", {}).get("asignatura"))
            for p in self.planillas
        }
        if len(asignaturas) > 1:
            return "notas_planillas.xlsx"
        enc = self.planillas[0]["encabezado"] if self.planillas else {}
        base = (enc.get("asignatura") or "notas").split()
        asignatura = " ".join(base[:3]) if base else "notas"
        return f"notas_{asignatura.replace(' ', '_')}.xlsx"

    # ------------------------------------------------------------------ #
    # 5) Pantalla final
    # ------------------------------------------------------------------ #
    def mostrar_final(self):
        pantalla = ctk.CTkFrame(self._contenedor, fg_color=styles.COLOR_FONDO)

        ctk.CTkLabel(
            pantalla, text="✅ ¡Listo!",
            font=(styles.FUENTE, styles.TAM_TITULO, "bold"), text_color=styles.COLOR_VERDE,
        ).pack(pady=(40, 10))

        ctk.CTkLabel(
            pantalla, text="El Excel con tus notas ya está guardado.\nPodés abrirlo "
            "directamente o ver la carpeta donde quedó.",
            font=(styles.FUENTE, styles.TAM_TEXTO), text_color=styles.COLOR_TEXTO,
            justify="center",
        ).pack(pady=(0, 8))

        archivo = getattr(self, "archivo_final", "")
        ctk.CTkLabel(
            pantalla, text=archivo, font=(styles.FUENTE, styles.TAM_TEXTO_CHICO),
            text_color=styles.COLOR_TEXTO_SECUNDARIO, wraplength=560,
        ).pack(pady=(0, 30))

        fila_btn = ctk.CTkFrame(pantalla, fg_color="transparent")
        fila_btn.pack(pady=12)

        ctk.CTkButton(
            fila_btn, text="Abrir archivo", height=48, width=190,
            font=(styles.FUENTE, styles.TAM_BOTON, "bold"),
            fg_color=styles.COLOR_PRINCIPAL, hover_color=styles.COLOR_PRINCIPAL_HOVER,
            command=lambda: self._abrir(ABRIR_ARCHIVO, archivo),
        ).pack(side="left", padx=8)

        ctk.CTkButton(
            fila_btn, text="Abrir carpeta", height=48, width=190,
            font=(styles.FUENTE, styles.TAM_BOTON, "bold"),
            fg_color=styles.COLOR_ACCENTO, hover_color="#6A59E0",
            command=lambda: self._abrir(ABRIR_CARPETA, archivo),
        ).pack(side="left", padx=8)

        ctk.CTkButton(
            pantalla, text="Cargar otras planillas", font=(styles.FUENTE, styles.TAM_TEXTO_CHICO),
            fg_color="transparent", text_color=styles.COLOR_TEXTO_SECUNDARIO,
            hover_color=styles.COLOR_FONDO_SECUNDARIO, command=self.mostrar_principal,
        ).pack(side="bottom", pady=(0, 14))

        self._cambiar_pantalla(pantalla)

    def _abrir(self, accion, ruta):
        try:
            if accion == ABRIR_ARCHIVO:
                os.startfile(ruta)
            else:
                os.startfile(os.path.dirname(ruta))
        except Exception as e:
            app_config.escribir_log(f"No se pudo abrir el archivo: {e!r}")
            messagebox.showerror(
                "No se pudo abrir",
                "No se pudo abrir el archivo. Podés buscarlo manualmente en la carpeta "
                "donde lo guardaste.",
            )


# ---------------------------------------------------------------------- #
# Helpers
# ---------------------------------------------------------------------- #
def _n_descripcion(n):
    """'4 columnas de notas' (para mensajes de error legibles por forma)."""
    return f"{n} columnas de notas"


def _fmt_celda(valor):
    if valor is None or valor == "":
        return ""
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return str(valor)


def _parse_celda(texto):
    """Convierte el texto editado a número o None si está vacío."""
    texto = (texto or "").strip()
    if not texto:
        return None
    try:
        return float(texto.replace(",", "."))
    except ValueError:
        return None


def _clear_tooltip(entrada, es_dudoso):
    """Sin tooltips por ahora: la celda amarilla ya comunica la duda."""
    _ = es_dudoso


def _fmt_peso(valor):
    """Formatea un peso (porcentaje) para mostrarlo en el formulario.

    Sin decimales si es entero (40 -> "40"), con un decimal si no (33.3 -> "33.3").
    """
    try:
        num = float(valor)
    except (TypeError, ValueError):
        return ""
    if num.is_integer():
        return str(int(num))
    return f"{num:.1f}"
