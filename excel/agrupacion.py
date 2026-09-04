"""
Agrupación de planillas por curso y combinación de estudiantes.

Fuente Única de Verdad (S8): tanto la pantalla de revisión de la GUI como el
generador de Excel usan estas funciones para agrupar por curso y combinar los
estudiantes de las páginas del mismo curso, con dedupe (S11). Si cambia la
regla, se cambia acá y los dos lados quedan idénticos.

Módulo puro: sin tkinter, sin red, sin efectos secundarios.
"""


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