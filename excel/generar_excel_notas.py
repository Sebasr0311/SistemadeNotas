"""
Generador de Excel de notas a partir de datos extraídos de una planilla.

Esta es la segunda mitad del pipeline: recibe los datos YA EXTRAÍDOS
(en producción, extraídos por el modelo de visión celda por celda desde
la imagen de la planilla) y arma el Excel con fórmulas reales de Excel
(no valores fijos), para que la definitiva se recalcule sola si se
corrige una nota a mano en el Excel.

Uso: se llama una vez por cada planilla (cada hoja = un grupo+asignatura+periodo).
"""

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


# ────────────────────────────────────────────────────────────────────
# Configuración de cálculo de notas (columnas seleccionables + pesos)
# ────────────────────────────────────────────────────────────────────

class ColumnConfig:
    """
    Representa la configuración de cálculo de notas para un curso.

    Attributes:
        modo: "simple" (promedio aritmético) o "pesos" (ponderado).
        columnas: lista de dicts, uno por cada columna de notas candidata
            ("Def. Periodo 1..N" y/o "Área Trabajo 1..M"):
            {
                "nombre": str,       # Nombre legible (ej. "Área Trabajo 1")
                "incluida": bool,    # True si entra al cálculo de la definitiva
                "peso": float,       # Porcentaje (0-100). Solo relevante si modo == "pesos".
                "tipo": str,         # "ev" (Def. Periodo) o "area" (Área Trabajo).
                                     # Retrocompat: ausente => "area".
                "pos": int,          # índice 0-based dentro de su tipo.
                                     # Retrocompat: ausente => índice en la lista.
            }
    """

    def __init__(self, modo="simple", columnas=None):
        self.modo = modo if modo in ("simple", "pesos") else "simple"
        self.columnas = columnas or []

    @property
    def columnas_seleccionadas(self):
        """Índices (0-based) de las columnas que entran al cálculo."""
        return [i for i, c in enumerate(self.columnas) if c.get("incluida", False)]

    @property
    def pesos_seleccionados(self):
        """Dict {índice: peso} de las columnas seleccionadas."""
        return {
            i: c.get("peso", 0)
            for i, c in enumerate(self.columnas)
            if c.get("incluida", False)
        }

    @property
    def n_columnas_seleccionadas(self):
        return len(self.columnas_seleccionadas)

    def es_pesado(self):
        return self.modo == "pesos" and self.n_columnas_seleccionadas > 1

    def pesos_suman_cien(self, tolerancia=0.01):
        """True si la suma de pesos de columnas seleccionadas ≈ 100."""
        if not self.es_pesado():
            return True
        total = sum(self.pesos_seleccionados.values())
        return abs(total - 100.0) <= tolerancia

    def to_dict(self):
        return {"modo": self.modo, "columnas": self.columnas}

    @classmethod
    def from_dict(cls, data):
        if not data:
            return cls()
        return cls(modo=data.get("modo", "simple"), columnas=data.get("columnas", []))

    @classmethod
    def crear_desde_planilla(cls, n_areas, n_ev=0, nombres_otras=None):
        """Crea una config por defecto: simple, todas las columnas incluidas,
        peso por defecto = 100 / total.

        Si `n_ev > 0`, la lista empieza con las "Def. Periodo 1..n_ev" (campo
        `ev_anteriores`, tipo "ev"), sigue con las "Área Trabajo 1..n_areas"
        (tipo "area") y termina con las columnas "otras" (tipo "otra", un dict
        por título con `{"nombre": <título>, "tipo": "otra", "pos": k, ...}`).
        Con los defaults `n_ev=0` y `nombres_otras=None` se conserva el
        comportamiento histórico (solo columnas de área).
        """
        total = n_areas + n_ev + (len(nombres_otras) if nombres_otras else 0)
        peso = round(100.0 / total, 1) if total > 0 else 0
        columnas = []
        for k in range(n_ev):
            columnas.append({
                "nombre": f"Def. Periodo {k+1}",
                "tipo": "ev",
                "pos": k,
                "incluida": True,
                "peso": peso,
            })
        for k in range(n_areas):
            columnas.append({
                "nombre": f"Área Trabajo {k+1}",
                "tipo": "area",
                "pos": k,
                "incluida": True,
                "peso": peso,
            })
        for k, titulo in enumerate(nombres_otras or []):
            columnas.append({
                "nombre": str(titulo or "").strip() or f"Otras notas {k+1}",
                "tipo": "otra",
                "pos": k,
                "incluida": True,
                "peso": peso,
            })
        return cls(modo="simple", columnas=columnas)


def nombres_columnas_areas(enc, n_areas):
    """Devuelve la lista de nombres legibles de las columnas de Área de Trabajo."""
    return [f"Área Trabajo {k+1}" for k in range(n_areas)]


def _ref_columna(col_ev_start, col_at1, r, i, col, col_otras_start=None):
    """Resuelve la referencia de celda de una columna candidata según su tipo.

    Devuelve la referencia completa con fila (ej. "C12").

    - tipo "ev"   -> letra de `col_ev_start + pos` + fila `r`.
    - tipo "area" (o sin tipo, retrocompat) -> letra de `col_at1 + pos` + fila `r`.
      Si no hay `pos`, se usa el índice `i` en la lista (comportamiento histórico,
      donde todas las columnas eran de área).
    - tipo "otra" -> letra de `col_otras_start + pos` + fila `r`. Requiere
      `col_otras_start` (columna donde empiezan las "otras" ya corridas en la
      hoja); si es None -> ValueError (candidata "otra" sin zona de otras).
    """
    from openpyxl.utils import get_column_letter as gcl

    tipo = col.get("tipo")
    if tipo == "ev":
        return f"{gcl(col_ev_start + col.get('pos', 0))}{r}"
    if tipo == "otra":
        if col_otras_start is None:
            raise ValueError(
                "Columna candidata tipo 'otra' sin col_otras_start: "
                "no hay zona de columnas 'otras' en la hoja."
            )
        pos = col.get("pos")
        if pos is None:
            pos = 0
        return f"{gcl(col_otras_start + int(pos))}{r}"
    pos = col.get("pos")
    if pos is None:
        pos = i
    return f"{gcl(col_at1 + pos)}{r}"


def _formula_definitiva(col_ev_start, col_at1, r, n_areas, column_config=None,
                        col_otras_start=None):
    """
    Genera la fórmula de Excel para la definitiva de un curso.

    Parámetros:
        col_ev_start: columna de la primera "Def. Periodo" (1-based)
        col_at1: columna de la primera Área de Trabajo (1-based)
        r: fila del estudiante
        n_areas: total de columnas de área de trabajo (para el modo simple sin config)
        column_config: ColumnConfig (None = modo simple con todas las columnas).
        col_otras_start: columna de la primera columna "otra" (1-based, ya
            corrida en la hoja). Solo necesario si la config trae tipo "otra".

    Cada columna de la config se resuelve según su tipo: "ev" -> col_ev_start + pos,
    "area" (o sin tipo) -> col_at1 + pos, "otra" -> col_otras_start + pos.

    Fórmulas generadas:
        Sin config / simple:    =IFERROR(SUM(col_at1:col_def-1)/n_areas,"")
        Simple con seleccionadas: =IFERROR(AVERAGE(refs),"") (1 → IFERROR(ref,""))
        Pesos:                  =IFERROR((ref*peso+...+ref*peso)/divisor,"")
    """
    from openpyxl.utils import get_column_letter as gcl

    if column_config is None:
        # Modo legacy: simple con todas las columnas de área
        at1_ref = f"{gcl(col_at1)}{r}"
        atn_ref = f"{gcl(col_at1 + n_areas - 1)}{r}"
        return f"=IFERROR(SUM({at1_ref}:{atn_ref})/{n_areas},\"\")"

    columnas = column_config.columnas
    seleccionadas = column_config.columnas_seleccionadas
    if not seleccionadas:
        return ""

    if not column_config.es_pesado():
        # Promedio simple de columnas seleccionadas
        refs = [_ref_columna(col_ev_start, col_at1, r, i, columnas[i], col_otras_start)
                for i in seleccionadas]
        if len(refs) == 1:
            return f"=IFERROR({refs[0]},\"\")"
        return f"=IFERROR(AVERAGE({','.join(refs)}),\"\")"

    # ── Modo pesos ──
    pesos = column_config.pesos_seleccionados
    n = len(seleccionadas)
    if n == 0:
        return ""
    if n == 1:
        # Un solo factor: se usa directamente (peso irrelevante)
        ref = _ref_columna(col_ev_start, col_at1, r, seleccionadas[0],
                           columnas[seleccionadas[0]], col_otras_start)
        return f"=IFERROR({ref},\"\")"

    # Construir la fórmula con una referencia de columna por cada peso
    # para que el recálculo automático funcione.
    partes = []
    for i in seleccionadas:
        letra = _ref_columna(col_ev_start, col_at1, r, i, columnas[i], col_otras_start)

        # el "letra" ya incluye la fila (ej. "E11"); el peso va multiplicando.
        partes.append(f"{letra}*{pesos[i]}")

    numerador = "+".join(partes)

    if column_config.pesos_suman_cien():
        divisor = "100"
    else:
        pesos_suma = "+".join(str(pesos[i]) for i in seleccionadas)
        divisor = f"({pesos_suma})"

    return f"=IFERROR(({numerador})/{divisor},\"\")"


def calcular_n_areas(enc, estudiantes) -> int:
    """Cantidad de columnas de área de trabajo (spec v2: de 1 a 16 notas).

    Misma lógica que usa el generador para dimensionar la hoja, extraída como
    función pura para que la pantalla de revisión de la GUI calcule EXACTAMENTE
    el mismo ancho (S8: lo que se muestra es lo que se escribe):
    - Si el encabezado declara n_area_trabajo como int (no bool) en 1..16,
      prevalece el declarado (celdas vacías al final si hay menos observadas).
    - Si no, se usa lo observado: la mayor cantidad de notas leídas por alumno.
    - Sin nada que medir, se cae al comportamiento histórico: planilla
      tradicional de 2 notas.
    """
    declarado_raw = enc.get("n_area_trabajo")
    declarado = (
        declarado_raw
        if isinstance(declarado_raw, int) and not isinstance(declarado_raw, bool)
        and 1 <= declarado_raw <= 16
        else 0
    )
    longitudes = [
        len(est.get("area_trabajo") or [])
        for est in estudiantes
    ]
    observado = max(longitudes) if longitudes else 0
    n_areas = max(declarado, observado)
    if n_areas == 0:
        # Sin declaración y sin ninguna nota: planillas tradicionales de 2 notas.
        n_areas = 2
    return n_areas


def calcular_n_otras(enc, estudiantes) -> int:
    """Cantidad de columnas "otras" de una planilla (trabajo práctico, parcial,
    recuperatorio, definitivas intermedias, etc.).

    Misma lógica que usa el generador para dimensionar la hoja, extraída como
    función pura para que la GUI calcule EXACTAMENTE el mismo ancho:
    - Si el encabezado declara `otras_columnas` como lista, prevalece su largo
      (celdas vacías al final si hay menos observadas).
    - Si no, se usa lo observado: la mayor cantidad de notas "otras" por alumno.
    - Sin nada que medir -> 0 (sin zona de columnas otras).
    """
    otras_enc = enc.get("otras_columnas")
    if isinstance(otras_enc, list):
        return len(otras_enc)
    longitudes = [
        len(est.get("otras_notas") or [])
        for est in estudiantes
    ]
    return max(longitudes) if longitudes else 0


def _escribir_hoja(ws, planilla: dict, column_config=None):
    """
    Escribe una hoja de Excel con los datos de una planilla.

    Args:
        ws: hoja de openpyxl.
        planilla: dict con encabezado + estudiantes.
        column_config: ColumnConfig (None = promedio simple de todas las columnas).
    """
    enc = planilla["encabezado"]
    estudiantes = planilla["estudiantes"]
    periodo = enc["periodo"]
    # Guard W1/W-A: nunca generar un Excel corrupto con un periodo inválido.
    if not (isinstance(periodo, int) and 1 <= periodo <= 4):
        raise ValueError(
            f"Periodo inválido ({periodo!r}) en el curso {enc.get('grupo', '?')}: "
            "debe ser 1, 2, 3 o 4."
        )
    n_ev_anteriores = periodo - 1

    # --- Ancho de la zona de área de trabajo (spec v2: 1 a 16 notas) ---
    n_areas = calcular_n_areas(enc, estudiantes)

    # --- Zona de columnas "otras" (trabajo práctico, parcial, etc.) ---
    n_otras = calcular_n_otras(enc, estudiantes)
    otras_titulos = list((enc.get("otras_columnas") or [])[:n_otras])
    while len(otras_titulos) < n_otras:
        otras_titulos.append(f"Otras notas {len(otras_titulos)+1}")

    bold = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="D9E1F2")
    revisar_fill = PatternFill("solid", fgColor="FFF2CC")
    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")

    # --- Encabezado informativo ---
    info_rows = [
        ("Institución", enc.get("institucion", "")),
        ("Sede", enc.get("sede", "")),
        ("Año lectivo", enc.get("año_lectivo", "")),
        ("Jornada", enc.get("jornada", "")),
        ("Grupo", enc.get("grupo", "")),
        ("Asignatura", enc.get("asignatura", "")),
        ("Docente", enc.get("docente", "")),
        ("Periodo", periodo),
    ]

    # Agregar info de modo de cálculo si hay config
    if column_config:
        modo_label = "Pesos" if column_config.modo == "pesos" else "Promedio simple"
        info_rows.append(("Modo de cálculo", modo_label))
        if column_config.modo == "pesos":
            seleccionadas = column_config.columnas_seleccionadas
            pesos = column_config.pesos_seleccionados
            partes = []
            for idx in seleccionadas:
                peso = pesos[idx]
                nombre = column_config.columnas[idx]["nombre"]
                partes.append(f"{nombre}: {peso:.1f}%")
            info_rows.append(("Pesos", " | ".join(partes) if partes else "Ninguna columna"))

    for i, (label, value) in enumerate(info_rows, start=1):
        ws.cell(row=i, column=1, value=label).font = bold
        ws.cell(row=i, column=2, value=value)

    header_row = len(info_rows) + 2

    # --- Columnas dinámicas ---
    cols = ["No.", "Nombre del Alumno"]
    for p in range(1, n_ev_anteriores + 1):
        cols.append(f"Def. Periodo {p}")
    cols += [f"Área Trabajo {k}" for k in range(1, n_areas + 1)]
    cols += otras_titulos
    cols.append(f"Definitiva Periodo {periodo}")
    incluir_anual = periodo == 4 and n_ev_anteriores == 3
    if incluir_anual:
        cols.append("Definitiva Anual")

    for j, title in enumerate(cols, start=1):
        c = ws.cell(row=header_row, column=j, value=title)
        c.font = bold
        c.fill = header_fill
        c.alignment = center
        c.border = border

    col_ev_start = 3
    col_at1 = col_ev_start + n_ev_anteriores
    col_otras_start = col_at1 + n_areas
    col_def = col_otras_start + n_otras
    col_anual = col_def + 1 if incluir_anual else None

    r = header_row + 1
    for est in estudiantes:
        ws.cell(row=r, column=1, value=est["no"]).border = border
        ws.cell(row=r, column=2, value=est["nombre"]).border = border

        if est.get("retirado"):
            ws.cell(row=r, column=2).font = Font(italic=True, color="999999")
            c = ws.cell(row=r, column=3, value="Retirado / sin datos")
            c.font = Font(italic=True, color="999999")
            r += 1
            continue

        ev = est.get("ev_anteriores") or []
        for k in range(n_ev_anteriores):
            v = ev[k] if k < len(ev) else None
            ws.cell(row=r, column=col_ev_start + k, value=v).border = border
            ws.cell(row=r, column=col_ev_start + k).alignment = center

        at = est.get("area_trabajo")
        revisar = est.get("revisar") or []
        for k in range(n_areas):
            v = at[k] if at and k < len(at) else None
            c = ws.cell(row=r, column=col_at1 + k, value=v)
            c.border = border
            c.alignment = center
            if k < len(revisar) and revisar[k]:
                c.fill = revisar_fill

        otr = est.get("otras_notas")
        revisar_otras = est.get("revisar_otras") or []
        for k in range(n_otras):
            v = otr[k] if otr and k < len(otr) else None
            c = ws.cell(row=r, column=col_otras_start + k, value=v)
            c.border = border
            c.alignment = center
            if k < len(revisar_otras) and revisar_otras[k]:
                c.fill = revisar_fill

        # ── Definitiva del periodo (con column_config si la hay) ──
        cdef = ws.cell(row=r, column=col_def)
        if at and len(at) >= 1:
            cdef.value = _formula_definitiva(
                col_ev_start, col_at1, r, n_areas, column_config, col_otras_start
            )
        cdef.border = border; cdef.alignment = center
        cdef.font = bold

        if incluir_anual:
            refs = [f"{get_column_letter(col_ev_start + k)}{r}" for k in range(n_ev_anteriores)]
            refs.append(f"{get_column_letter(col_def)}{r}")
            canual = ws.cell(row=r, column=col_anual)
            canual.value = f"=IFERROR(AVERAGE({','.join(refs)}),\"\")"
            canual.border = border; canual.alignment = center
            canual.font = bold

        r += 1

    # --- Leyenda ---
    r += 1
    leyenda = ws.cell(row=r, column=2, value="Amarillo = nota dudosa (tachón/letra ambigua), verificar contra la planilla física")
    leyenda.font = Font(italic=True, size=9, color="7F6000")
    ws.cell(row=r, column=2).fill = revisar_fill

    # --- Anchos de columna ---
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 34
    for j in range(3, len(cols) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 14


def _clave_forma(planilla: dict):
    """Clave de forma de una planilla: tupla ("n_areas", n) con su cantidad de
    columnas de notas calculada POR PLANILLA (misma clave que agrupar_por_forma
    y que usa la GUI para guardar las configs por forma)."""
    enc = planilla.get("encabezado") or {}
    estudiantes = planilla.get("estudiantes") or []
    return ("n_areas", calcular_n_areas(enc, estudiantes))


def _resolver_column_config(planilla: dict, column_config, column_configs):
    """
    Resuelve la ColumnConfig para una planilla según la precedencia:

    - Si `column_configs` no es None -> se busca por clave de forma
      `_clave_forma(planilla)`; si no hay clave para esa planilla -> None
      (legacy para ese grupo).
    - Si `column_configs` es None y `column_config` no -> la única config.
    - Ambos None -> None (legacy histórico).
    """
    if column_configs is not None:
        return column_configs.get(_clave_forma(planilla))
    return column_config


def generar_excel_planilla(planilla: dict, ruta_salida: str, column_config=None,
                           column_configs=None):
    """
    Genera un Excel de una sola planilla (un curso).

    Args:
        planilla: dict con encabezado + estudiantes.
        ruta_salida: ruta del archivo .xlsx de salida.
        column_config: ColumnConfig (None = promedio simple de todas las columnas).
        column_configs: dict {clave_forma: ColumnConfig} opcional. Si se pasa,
                        se resuelve la config de esta planilla por su forma
                        (precedencia sobre column_config).
    """
    config = _resolver_column_config(planilla, column_config, column_configs)
    wb = openpyxl.Workbook()
    ws = wb.active
    enc = planilla["encabezado"]
    ws.title = f"P{enc['periodo']} - {enc['grupo']}"
    _escribir_hoja(ws, planilla, config)
    wb.save(ruta_salida)
    return ruta_salida


def generar_excel_asignatura(planillas: list, ruta_salida: str, column_config=None,
                             column_configs=None):
    """
    Caso real de uso: un lote sube TODAS las planillas de una misma asignatura
    (varios cursos/grupos). Se agrupan automáticamente por curso y se genera
    UN SOLO Excel con una hoja por cada curso.

    Args:
        planillas: lista de dicts, cada uno con encabezado + estudiantes.
        ruta_salida: ruta del archivo .xlsx de salida.
        column_config: ColumnConfig (None = promedio simple de todas las columnas).
                       Se aplica la MISMA config a todas las hojas.
        column_configs: dict {clave_forma: ColumnConfig} opcional. Si se pasa,
                        cada hoja resuelve su config POR FORMA (la cantidad de
                        columnas de notas de ese grupo), con precedencia sobre
                        `column_config`. Si una forma no tiene clave -> None
                        (legacy para ese grupo).
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # se reemplaza por una hoja por curso

    from excel.agrupacion import agrupar_por_curso, combinar_estudiantes

    por_curso, orden_grupos = agrupar_por_curso(planillas)

    nombres_usados = set()
    for clave in orden_grupos:
        grupo = clave[0]
        asignatura = clave[1]
        paginas = por_curso[clave]
        base = dict(paginas[0])
        base["estudiantes"] = combinar_estudiantes(paginas)
        planilla_final = base

        # Resolver la config de ESTA forma para este grupo (None si no aplica).
        config_hoja = _resolver_column_config(planilla_final, column_config, column_configs)

        asignatura_corta = asignatura[:22].strip()
        nombre_hoja = f"Curso {grupo} - {asignatura_corta}"[:31]
        original = nombre_hoja
        i = 2
        while nombre_hoja in nombres_usados:
            nombre_hoja = f"{original[:28]} ({i})"
            i += 1
        nombres_usados.add(nombre_hoja)

        ws = wb.create_sheet(title=nombre_hoja)
        _escribir_hoja(ws, planilla_final, config_hoja)

    wb.save(ruta_salida)
    return ruta_salida


if __name__ == "__main__":
    # Datos transcritos de la planilla de ejemplo (Pedro Castro Monsalvo,
    # grupo 0302, periodo 3, Educación Física) para validar el formato del Excel.
    # En producción esta lista la llena la extracción por visión, celda a celda.
    estudiantes = [
        (1, "ALFARO OCHOA LYHAM ANDRES", [45, 45], [40, 50]),
        (2, "ARIAS ECHAVEZ CELESTE SOFIA", [45, 45], [40, 45]),
        (3, "ATENCIO SIERRA KEYLLER DAVID", [45, 50], [40, 45]),
        (4, "BOTELLO ORTIZ ALVARO JOSUE", [45, 50], [45, 40]),
        (5, "CASTRO PINTO SALOMON", [45, 45], [45, 40]),
        (6, "CONTRERAS CARPIO ELIO DE JESUS", [45, 45], [45, 45]),
        (7, "DIAZ QUINTERO DYLAN DAVID", [45, 45], [45, 50]),
        (8, "ESCAMILLA GUTIERREZ ANA JULIA", [45, 45], [40, 45]),
        (9, "ESCORCIA PEDROZA MARIA SALOME", [45, 45], [40, 45]),
        (10, "GARCIA MARIN LUISA FERNANDA", [45, 44], [40, 45]),
        (11, "GOMEZ AMAYA MARIA VICTORIA", [45, 44], [45, 50]),
        (12, "GONZALEZ CASTRO LIAM DAVID", [45, 45], [40, 40]),
        (13, "GUTIERREZ ARAGON MARJALYS", [40, 40], [40, 35]),
        (14, "MARTELO ARIAS JUAN MIGUEL", [45, 50], [45, 45]),
        (15, "MARTINEZ CASTRO ELIAS DAVID", [45, 45], [40, 50]),
        (16, "MERIÑO THOMPSON DAIRO JUNIOR", [45, 45], [40, 40]),
        (17, "MOLINA OCHOA VALENTINO JOSE", [45, 45], [40, 50]),
        (18, "MURGAS MARTINEZ ERIYETH ANTONELA", [45, 50], [45, 50]),
        (19, "NAVARRO AMARIS JEREMY", [44, 40], [40, 45]),
        (20, "NAVARRO BUSTO AMANDA SOFIA", [45, 46], [40, 50]),
        (21, "PEÑA MUÑOZ ALEXANDRA", [45, 45], [40, 45]),
        (22, "PEÑA MUÑOZ MARIA FERNANDA", [45, 45], [45, 50]),
        (23, "PEREIRA JAIMES KALET DAVID", [45, 44], [40, 40]),
        (24, "PEREIRA JAIMES SEBASTIAN DAVID", [45, 45], [40, 40]),
        (25, "PEREIRA PEÑA KATHELYN VICTORIA", [45, 40], None),
        (26, "PEREIRA VILLEGAS ABRAHAM DAVID", [45, 40], [40, 40]),
        (27, "RASGO RAMIREZ JHAZIEL", [45, 50], [45, 45]),
        (28, "REBOLLEDO DE LA HOZ JORGE DE JESUS", [45, 42], [40, 40]),
        (29, "RIOS MUÑOZ MARIA FERNANDA", [45, 40], [40, 40]),
        (30, "ROMERO MACHUCA DANIELA PATRICIA", [45, 46], [40, 50]),
        (31, "SALAZAR MALDONADO IAN LUCA", [45, 45], [40, 40]),
        (32, "TERAN DE LA HOZ MARIA ELENA", [45, 45], [40, 45]),
        (33, "VENERA ARAMENDIZ ANDRES DAVID", [45, 45], [35, 40]),
        (34, "YEPES ALVAREZ VALERY", [45, 40], [40, 45]),
        (35, "ZAGARRA MEJIA JUAN DIEGO", [45, 46], [40, 50]),
    ]

    def encabezado_base(grupo):
        return {
            "institucion": "INST. ED. TEC. INDUSTRIAL PEDRO CASTRO MONSALVO",
            "sede": "SEDE CINCO DE ENERO",
            "año_lectivo": "2026",
            "jornada": "MAÑANA",
            "grupo": grupo,
            "asignatura": "0300501 EDUCACION FISICA REC. Y DEPORTES",
            "docente": "Narlis Ester Farelo Calderon",
            "periodo": 3,
        }

    def planilla_para(grupo):
        return {
            "encabezado": encabezado_base(grupo),
            "estudiantes": [
                {
                    "no": no, "nombre": nom, "ev_anteriores": ev, "area_trabajo": at, "retirado": False,
                    "revisar": [False, True] if (no == 11) else [False, False],
                }
                for no, nom, ev, at in estudiantes
            ] + [
                {"no": 36, "nombre": "LASTRE SANCHEZ ANTONELLA", "ev_anteriores": [45], "area_trabajo": None, "retirado": True}
            ],
        }

    # Simulación del caso real: un PDF con las planillas de 3 cursos distintos
    # de la misma asignatura -> un solo Excel con una hoja por curso.
    planillas_pdf = [planilla_para("0302"), planilla_para("0401"), planilla_para("0501")]

    import os
    import tempfile

    # Ruta demo portable: usa la carpeta temporal del sistema (S5), nunca una
    # ruta fija a C:\Users\JUAN.
    ruta_salida = os.path.join(
        tempfile.gettempdir(),
        "notas_educacion_fisica_periodo3.xlsx",
    )
    os.makedirs(os.path.dirname(ruta_salida), exist_ok=True)
    ruta = generar_excel_asignatura(planillas_pdf, ruta_salida)
    print("Generado:", ruta)
