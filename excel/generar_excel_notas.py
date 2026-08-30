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


def _escribir_hoja(ws, planilla: dict):
    enc = planilla["encabezado"]
    estudiantes = planilla["estudiantes"]
    periodo = enc["periodo"]
    n_ev_anteriores = periodo - 1

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
    for i, (label, value) in enumerate(info_rows, start=1):
        ws.cell(row=i, column=1, value=label).font = bold
        ws.cell(row=i, column=2, value=value)

    header_row = len(info_rows) + 2

    # --- Columnas dinámicas ---
    cols = ["No.", "Nombre del Alumno"]
    for p in range(1, n_ev_anteriores + 1):
        cols.append(f"Def. Periodo {p}")
    cols += ["Área Trabajo 1", "Área Trabajo 2", f"Definitiva Periodo {periodo}"]
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
    col_at2 = col_at1 + 1
    col_def = col_at2 + 1
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
        revisar = est.get("revisar") or [False, False]
        c1 = ws.cell(row=r, column=col_at1, value=(at[0] if at and len(at) > 0 else None))
        c2 = ws.cell(row=r, column=col_at2, value=(at[1] if at and len(at) > 1 else None))
        c1.border = border; c1.alignment = center
        c2.border = border; c2.alignment = center
        if revisar[0]:
            c1.fill = revisar_fill
        if len(revisar) > 1 and revisar[1]:
            c2.fill = revisar_fill

        at1_ref = f"{get_column_letter(col_at1)}{r}"
        at2_ref = f"{get_column_letter(col_at2)}{r}"
        cdef = ws.cell(row=r, column=col_def)
        if at and len(at) >= 1:
            cdef.value = f"=IFERROR(AVERAGE({at1_ref}:{at2_ref}),\"\")"
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


def generar_excel_planilla(planilla: dict, ruta_salida: str):
    """Genera un Excel de una sola planilla (un curso). Ver generar_excel_asignatura
    para el caso real de uso: un PDF con varias planillas de la misma asignatura."""
    wb = openpyxl.Workbook()
    ws = wb.active
    enc = planilla["encabezado"]
    ws.title = f"P{enc['periodo']} - {enc['grupo']}"
    _escribir_hoja(ws, planilla)
    wb.save(ruta_salida)
    return ruta_salida


def generar_excel_asignatura(planillas: list, ruta_salida: str):
    """
    Caso real de uso: un PDF sube TODAS las planillas de una misma asignatura
    (varios cursos/grupos). Se agrupan automáticamente por curso y se genera
    UN SOLO Excel con una hoja por cada curso.

    planillas: lista de dicts, cada uno con la misma forma que en
    generar_excel_planilla (un dict por planilla/página reconocida en el PDF).
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # se reemplaza por una hoja por curso

    # Agrupar por grupo (curso) — normalmente 1 planilla por curso, pero si llegan
    # varias páginas del mismo curso (ej. planilla partida en 2 hojas) se combinan.
    por_grupo = {}
    orden_grupos = []
    for p in planillas:
        grupo = p["encabezado"]["grupo"]
        if grupo not in por_grupo:
            por_grupo[grupo] = []
            orden_grupos.append(grupo)
        por_grupo[grupo].append(p)

    nombres_usados = set()
    for grupo in orden_grupos:
        paginas = por_grupo[grupo]
        if len(paginas) > 1:
            base = dict(paginas[0])
            base["estudiantes"] = [est for pg in paginas for est in pg["estudiantes"]]
            planilla_final = base
        else:
            planilla_final = paginas[0]

        nombre_hoja = f"Curso {grupo}"[:31]
        original = nombre_hoja
        i = 2
        while nombre_hoja in nombres_usados:
            nombre_hoja = f"{original[:28]} ({i})"
            i += 1
        nombres_usados.add(nombre_hoja)

        ws = wb.create_sheet(title=nombre_hoja)
        _escribir_hoja(ws, planilla_final)

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
    ruta_salida = r"C:\Users\JUAN\AppData\Local\Temp\opencode\demo_excel\notas_educacion_fisica_periodo3.xlsx"
    os.makedirs(os.path.dirname(ruta_salida), exist_ok=True)
    ruta = generar_excel_asignatura(planillas_pdf, ruta_salida)
    print("Generado:", ruta)
