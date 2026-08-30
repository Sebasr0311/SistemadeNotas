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
from tkinter import filedialog, messagebox

import customtkinter as ctk

from config import app_config
from excel import generar_excel_notas
from . import styles
from .worker import ProcesadorEnSegundoPlano, MSG_PROGRESO, MSG_RESULTADO, MSG_ERROR, MSG_CANCELADO

ABRIR_ARCHIVO = "open_file"
ABRIR_CARPETA = "open_folder"


class App(ctk.CTk):
    """Ventana principal y controlador del flujo entre pantallas."""

    def __init__(self):
        super().__init__()
        styles.configurar_tema()
        self.title("Sistema de Digitalización de Planillas de Notas")
        self.geometry("860x640")
        self.minsize(720, 560)
        self._configurar_apariencia()

        self.planillas = []           # planillas extraídas (una por página)
        self.paginas_total = 0
        self.paginas_fallidas = []    # páginas que no se pudieron leer (S2)
        self.planilla_actual_idx = 0  # índice usado en la pantalla de revisión
        self._worker_cola = None
        self._worker = None

        # Contenedor único donde se montan las pantallas.
        self._contenedor = ctk.CTkFrame(self, fg_color=styles.COLOR_FONDO)
        self._contenedor.pack(fill="both", expand=True)
        self._pantalla_actual = None

        self._mostrar_inicio()

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
            pantalla, text="¡Bienvenida a Notas Digital!",
            font=(styles.FUENTE, styles.TAM_TITULO, "bold"), text_color=styles.COLOR_TEXTO,
        )
        titulo.pack(pady=(20, 8))

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
            text="Elegí el PDF con todas las planillas de una misma asignatura\n"
            "(puede tener varios cursos: la app los ordena solos).",
            font=(styles.FUENTE, styles.TAM_TEXTO), text_color=styles.COLOR_TEXTO_SECUNDARIO,
            justify="center",
        ).pack(pady=(0, 30))

        boton_cargar = ctk.CTkButton(
            pantalla,
            text="📄  Cargar PDF de planillas",
            height=72, width=360, font=(styles.FUENTE, styles.TAM_BOTON_GRANDE, "bold"),
            fg_color=styles.COLOR_PRINCIPAL, hover_color=styles.COLOR_PRINCIPAL_HOVER,
            corner_radius=16, command=self._elegir_pdf,
        )
        boton_cargar.pack(pady=16)

        ctk.CTkLabel(
            pantalla, text="¿Cómo funciona?",
            font=(styles.FUENTE, styles.TAM_SUBTITULO, "bold"), text_color=styles.COLOR_TEXTO,
        ).pack(pady=(30, 6))

        pasos = (
            "1. Escaneá las planillas de una asignatura y armá un PDF.\n"
            "2. Cargalo en la app.\n"
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
        ruta = filedialog.askopenfilename(
            title="Elegí el PDF de planillas",
            filetypes=[("Archivos PDF", "*.pdf"), ("Todos los archivos", "*.*")],
        )
        if not ruta:
            return
        self.mostrar_progreso(ruta)

    # ------------------------------------------------------------------ #
    # 3) Pantalla de progreso
    # ------------------------------------------------------------------ #
    def mostrar_progreso(self, ruta_pdf):
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

        # Arrancar el hilo de procesamiento.
        self._worker_cola = queue.Queue()
        self._worker = ProcesadorEnSegundoPlano(ruta_pdf, self._worker_cola)
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
                        # S2: ninguna página se pudo leer -> error claro, y
                        # NO se pasa a una pantalla de revisión vacía.
                        messagebox.showerror(
                            "No se pudo leer",
                            "No se pudo leer ninguna página del PDF. Revisá que las "
                            "planillas estén bien escaneadas y volvé a intentar.",
                        )
                        self.mostrar_principal()
                        return
                    if self.paginas_fallidas:
                        # S2: algunas páginas fallaron pero el resto sirve: se
                        # avisa y se continúa igual a la revisión.
                        lista = ", ".join(str(f["pagina"]) for f in self.paginas_fallidas)
                        messagebox.showwarning(
                            "Algunas páginas no se leyeron",
                            f"Se generaron {len(self.planillas)} de {self.paginas_total} "
                            f"planillas. No se pudieron leer las páginas: {lista}. "
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
                        "Proceso cancelado. Podés volver a cargar el PDF cuando quieras.",
                    )
                    self.mostrar_principal()
                    return
        except queue.Empty:
            pass
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
            pantalla, text="Generar Excel", height=52,
            font=(styles.FUENTE, styles.TAM_BOTON_GRANDE, "bold"),
            fg_color=styles.COLOR_VERDE, hover_color="#5AA87A", command=self._generar_excel,
        )
        boton_guardar.pack(pady=(0, 6))

        ctk.CTkLabel(
            pantalla, text="Amarillo = nota dudosa, verificarla en la planilla física.",
            font=(styles.FUENTE, styles.TAM_TEXTO_CHICO), text_color=styles.COLOR_REVISAR_BORDE,
        ).pack(pady=(0, 10))

        self._cambiar_pantalla(pantalla)

    def _agrupar_por_curso(self):
        """Devuelve dict curso -> lista de planillas (páginas)."""
        por_curso = {}
        orden = []
        for p in self.planillas:
            g = p["encabezado"].get("grupo") or "SIN CURSO"
            if g not in por_curso:
                por_curso[g] = []
                orden.append(g)
            por_curso[g].append(p)
        return por_curso, orden

    def _render_cursos(self):
        # Limpiar el frame por si se vuelve a entrar.
        for w in self._frame_cursos.winfo_children():
            w.destroy()

        por_curso, orden = self._agrupar_por_curso()
        self._editores = {}  # (curso, idx_fila, celda_idx) -> variable StringVar
        self._periodo_combos = {}  # curso -> CTkComboBox de corrección de periodo

        for curso in orden:
            paginas = por_curso[curso]
            # Combinar estudiantes de páginas del mismo curso.
            estudiantes = [est for pg in paginas for est in pg["estudiantes"]]
            enc = paginas[0]["encabezado"]

            tarjeta = ctk.CTkFrame(
                self._frame_cursos, fg_color=styles.COLOR_BLANCO, corner_radius=12,
                border_width=1, border_color="#E3E9F5",
            )
            tarjeta.pack(fill="x", pady=8, padx=4)

            titulo = f"Curso {curso} — Periodo {enc.get('periodo')}"
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
                self._periodo_combos[curso] = combo

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

            # Encabezados de columnas.
            cabecera = ctk.CTkFrame(tarjeta, fg_color="#EDF2FB", corner_radius=8)
            cabecera.pack(fill="x", padx=10)
            ctk.CTkLabel(cabecera, text="No.", font=(styles.FUENTE, styles.TAM_TEXTO_CHICO, "bold"),
                         text_color=styles.COLOR_TEXTO, width=44).pack(side="left", padx=(10, 4), pady=6)
            ctk.CTkLabel(cabecera, text="Nombre del Alumno", font=(styles.FUENTE, styles.TAM_TEXTO_CHICO, "bold"),
                         text_color=styles.COLOR_TEXTO, width=280).pack(side="left", pady=6)
            ctk.CTkLabel(cabecera, text="Área Trabajo 1", font=(styles.FUENTE, styles.TAM_TEXTO_CHICO, "bold"),
                         text_color=styles.COLOR_TEXTO, width=110).pack(side="left", pady=6)
            ctk.CTkLabel(cabecera, text="Área Trabajo 2", font=(styles.FUENTE, styles.TAM_TEXTO_CHICO, "bold"),
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

                at = est.get("area_trabajo") or [None, None]
                rev = est.get("revisar") or [False, False]

                ctk.CTkLabel(
                    fila, text=str(est.get("no", "")), font=(styles.FUENTE, styles.TAM_CELDA),
                    text_color=styles.COLOR_TEXTO, width=44,
                ).pack(side="left", padx=(10, 4))

                ctk.CTkLabel(
                    fila, text=est.get("nombre", ""), font=(styles.FUENTE, styles.TAM_CELDA),
                    text_color=styles.COLOR_TEXTO, width=280, anchor="w",
                ).pack(side="left")

                for celda_idx in (0, 1):
                    var = tk.StringVar(value=_fmt_celda(at[celda_idx]))
                    self._editores[(curso, idx, celda_idx)] = var
                    color_fondo = styles.COLOR_REVISAR if rev[celda_idx] else styles.COLOR_FONDO_SECUNDARIO
                    entrada = ctk.CTkEntry(
                        fila, textvariable=var, width=100, height=30,
                        font=(styles.FUENTE, styles.TAM_CELDA),
                        fg_color=color_fondo,
                        border_color=styles.COLOR_REVISAR_BORDE if rev[celda_idx] else "#D5DEEF",
                    )
                    _clear_tooltip(entrada, rev[celda_idx])
                    entrada.pack(side="left", padx=6, pady=3)

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
            generar_excel_notas.generar_excel_asignatura(self.planillas, ruta)
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
        por_curso, _ = self._agrupar_por_curso()
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
        por_curso, orden = self._agrupar_por_curso()
        for curso in orden:
            paginas = por_curso[curso]
            estudiantes = [est for pg in paginas for est in pg["estudiantes"]]
            for idx, est in enumerate(estudiantes):
                if est.get("retirado"):
                    continue
                area = est.get("area_trabajo") or [None, None]
                revisar = est.get("revisar") or [False, False]
                for celda_idx in (0, 1):
                    var = self._editores.get((curso, idx, celda_idx))
                    if var is None:
                        continue
                    valor_texto = (var.get() or "").strip()
                    area[celda_idx] = _parse_celda(valor_texto)
                    # Si la usuaria borró o corrigió un valor dudoso, se quita el
                    # resaltado sólo cuando ya no es dudoso por rango.
                    if valor_texto:
                        try:
                            num = float(valor_texto.replace(",", "."))
                            if 0 <= num <= 100:
                                revisar[celda_idx] = False
                        except ValueError:
                            pass
                    else:
                        revisar[celda_idx] = False
                est["area_trabajo"] = area
                est["revisar"] = revisar
                # S3: el flag de Ev. Anteriores se preserva tal cual al editarla;
                # no se recalcula ni se borra (no hay celdas editables para ev).
                est["revisar_ev"] = bool(est.get("revisar_ev"))

    def _nombre_archivo_sugerido(self):
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
            pantalla, text="Cargar otro PDF", font=(styles.FUENTE, styles.TAM_TEXTO_CHICO),
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
