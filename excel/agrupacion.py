"""
Agrupación de planillas por curso y combinación de estudiantes.

Fuente Única de Verdad (S8): tanto la pantalla de revisión de la GUI como el
generador de Excel usan estas funciones para agrupar por curso y combinar los
estudiantes de las páginas del mismo curso, con dedupe (S11). Si cambia la
regla, se cambia acá y los dos lados quedan idénticos.

Módulo puro: sin tkinter, sin red, sin efectos secundarios.
"""


def agrupar_por_forma(planillas):
    """
    Agrupa las planillas POR FORMA (cantidad de columnas de notas).

    Un PDF con varias planillas puede tener planillas con distinta cantidad de
    columnas de área de trabajo (ej. un curso con 4 columnas y otro con 6). La
    configuración de columnas se elige UNA VEZ por "forma", y se aplica a todas
    las planillas iguales.

    Clave = tupla `("n_areas", n)` donde `n` = `calcular_n_areas(...)` calculado
    POR PLANILLA (no por curso). Orden estable por primera aparición.

    Devuelve (por_forma: dict[tuple, list], orden_formas: list[tuple]).
    """
    # Import local para evitar ciclos: este módulo (excel.agrupacion) puede ser
    # importado por excel.generar_excel_notas, que a su vez importa acá; el
    # cálculo de n_areas vive en generar_excel_notas.
    from excel.generar_excel_notas import calcular_n_areas

    por_forma = {}
    orden_formas = []
    for p in planillas:
        enc = p.get("encabezado") or {}
        estudiantes = p.get("estudiantes") or []
        n = calcular_n_areas(enc, estudiantes)
        clave = ("n_areas", n)
        if clave not in por_forma:
            por_forma[clave] = []
            orden_formas.append(clave)
        por_forma[clave].append(p)
    return por_forma, orden_formas


def _asignatura_limpia(nombre) -> str:
    """Normaliza el nombre de una asignatura para usarla como clave de agrupación.

    - Mayúsculas y espacios colapsados.
    - Quita un prefijo de código numérico (5 a 8 dígitos) seguido de espacio
      (ej. "0500101 CIENCIAS NATURALES..." -> "CIENCIAS NATURALES...").
    - Si tras el recorte el nombre queda vacío, devuelve "" (el caller aplica
      el fallback, ej. "SIN ASIGNATURA").
    """
    limpio = " ".join(str(nombre or "").strip().upper().split())
    # Quitar el prefijo de código numérico (5 a 8 dígitos) + espacio.
    import re
    limpio = re.sub(r"^\d{5,8}\s+", "", limpio)
    return limpio


def agrupar_por_curso(planillas):
    """
    Agrupa las planillas por curso Y asignatura.

    Clave = tupla `(grupo, asignatura)`:
    - `grupo`      = `encabezado.get("grupo")` o "SIN CURSO" si falta/viene vacío.
    - `asignatura` = `_asignatura_limpia(encabezado.get("asignatura"))` o
                     "SIN ASIGNATURA" si queda vacía/falta.
    Orden = orden de primera aparición en la lista.

    DOS planillas pertenecen al MISMO grupo sólo si coinciden curso Y
    asignatura: un mismo curso (ej. 0501) puede tener planillas de asignaturas
    DISTINTAS en el mismo PDF (caso diagnosticado: Naturales e Informática en
    páginas distintas del mismo grupo). Antes se agrupaba sólo por `grupo` y la
    segunda página de cada curso se descartaba en silencio al combinar los
    estudiantes por número (los mismos 1..N).

    Devuelve (por_curso: dict[tuple[str, str], list], orden: list[tuple[str, str]]).
    """
    por_curso = {}
    orden = []
    for p in planillas:
        enc = p.get("encabezado") or {}
        g = enc.get("grupo") or "SIN CURSO"
        asignatura = _asignatura_limpia(enc.get("asignatura")) or "SIN ASIGNATURA"
        clave = (g, asignatura)
        if clave not in por_curso:
            por_curso[clave] = []
            orden.append(clave)
        por_curso[clave].append(p)
    return por_curso, orden


def _nombre_normalizado(nombre) -> str:
    """Mayúsculas, espacios colapsados: base para comparar nombres iguales."""
    return " ".join(str(nombre or "").strip().upper().split())


def combinar_estudiantes(paginas):
    """
    Concatena los estudiantes de las páginas de un curso y elimina duplicados.

    Regla de dedupe (S11: escaneos duplicados o solapados no deben entrar 2 veces):
    - Se conserva la PRIMERA aparición de cada `no` cuando `no` es entero > 0.
    - Si `no` es 0 o está ausente EN AMBOS, se deduplica por nombre normalizado
      (`nombre.strip().upper()` colapsando espacios).
    El estudiante duplicado se descarta: el primero gana, con sus valores.
    """
    resultado = []
    vistos_no = set()
    vistos_nombre = set()
    for pg in paginas:
        for est in pg.get("estudiantes") or []:
            try:
                no_int = int(est.get("no"))
            except (TypeError, ValueError):
                no_int = 0
            if no_int > 0:
                if no_int in vistos_no:
                    continue
                vistos_no.add(no_int)
            else:
                clave_nombre = _nombre_normalizado(est.get("nombre"))
                if clave_nombre:
                    if clave_nombre in vistos_nombre:
                        continue
                    vistos_nombre.add(clave_nombre)
            resultado.append(est)
    return resultado