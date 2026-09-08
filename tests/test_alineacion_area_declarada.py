"""
Tests de la alineación posicional de "area_trabajo" contra n_area_trabajo
(anti-colapso de celdas vacías).

El problema reportado: cuando una planilla tiene una celda vacía EN MEDIO de las
notas (ej. "45 40 [vacía] 50"), el reporte mostraba "45 40 50 [vacía]" — el
modelo colapsaba la lista compactando los nulls posicionales al final y se
"perdía el orden". Como con una lista más corta NO se puede reconstruir si la
celda vacía era interna, cuando n_area_trabajo está declarado:

- Lista MÁS CORTA que el declarado -> se rellena con None hasta el declarado y
  TODA la fila se marca revisar=True (corrimiento posible).
- Lista MÁS LARGA que el declarado -> se trunca al declarado y TODA la fila se
  marca revisar=True (sobran valores -> sospechoso).
- Lista con la longitud exacta -> sin marca extra (los flags del modelo se
  preservan tal cual).
- Sin n_area_trabajo declarado -> comportamiento previo intacto.

Además, el prompt (_PROMPT_PLANILLA) incluye la regla anti-colapso con el
ejemplo concreto y la regla de celdas ilegibles ("+", mancha, tachón -> null
en su posición + revisar, nunca inventar un número).

Uso (desde la raíz del proyecto):
    python -m unittest discover -s tests -v
"""

import os
import sys
import unittest

_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

import vision.gemini_extractor as gem  # noqa: E402


def _planilla(areas, n_area_trabajo, revisar_s=None):
    """Planilla cruda con 12 estudiantes (dentro del rango sano 10-60).

    areas: lista de listas (una por estudiante) o lista única que se repite
    para todos.
    revisar_s: igual que areas pero para los flags "revisar" (opcional).
    """
    if areas and isinstance(areas[0], list):
        areas_por_est = list(areas)
    else:
        areas_por_est = [list(areas)] * 12
    encabezado = {
        "institucion": "I", "sede": "S", "año_lectivo": "2026",
        "jornada": "M", "grupo": "0302", "asignatura": "A",
        "docente": "D", "periodo": 3,
    }
    if n_area_trabajo is not None:
        encabezado["n_area_trabajo"] = n_area_trabajo
    estudiantes = []
    for i, area in enumerate(areas_por_est, start=1):
        est = {
            "no": i, "nombre": f"ALUMNO {i}",
            "ev_anteriores": [45, 45],
            "area_trabajo": list(area), "retirado": False,
        }
        if revisar_s is not None:
            est["revisar"] = (
                list(revisar_s[i - 1]) if isinstance(revisar_s[0], list)
                else list(revisar_s)
            )
        estudiantes.append(est)
    return {"encabezado": encabezado, "estudiantes": estudiantes}


# ---------------------------------------------------------------------- #
# Alineación a nivel planilla (_normalizar_planilla)
# ---------------------------------------------------------------------- #

class TestAlineacionEnPlanilla(unittest.TestCase):

    def test_lista_corta_se_rellena_y_marca_toda_la_fila(self):
        # El modelo colapsó: devolvió [45, 40, 50] para 4 columnas declaradas
        # (el null posicional de la celda vacía se perdió). Se rellena con None
        # al ancho declarado, TODA la fila queda en revisión y la planilla se
        # marca revisar_planilla (no se puede reconstruir el orden).
        areas = [[45, 40, 50]] + [[40, 50, 60, 70]] * 11
        p = gem._normalizar_planilla(_planilla(areas, n_area_trabajo=4))
        est = p["estudiantes"][0]
        self.assertEqual(est["area_trabajo"], [45, 40, 50, None])
        self.assertEqual(est["revisar"], [True, True, True, True])
        self.assertTrue(p["revisar_planilla"])
        # Los demás alumnos, bien alineados, no llevan marca extra.
        bien = p["estudiantes"][1]
        self.assertEqual(bien["area_trabajo"], [40, 50, 60, 70])
        self.assertEqual(bien["revisar"], [False, False, False, False])

    def test_lista_bien_alineada_no_marca_extra(self):
        # [45, 40, null, 50] con 4 columnas: el null va en su posición, sin
        # marca extra y sin revisar_planilla.
        areas = [[45, 40, None, 50]] + [[40, 50, 60, 70]] * 11
        p = gem._normalizar_planilla(_planilla(areas, n_area_trabajo=4))
        est = p["estudiantes"][0]
        self.assertEqual(est["area_trabajo"], [45, 40, None, 50])
        self.assertEqual(est["revisar"], [False, False, False, False])
        self.assertFalse(p["revisar_planilla"])

    def test_flags_del_modelo_se_preservan_en_lista_alineada(self):
        # Longitud exacta + el modelo marca revisar en la celda vacía (2): la
        # marca se conserva tal cual (nunca se sobreescribe a False).
        areas = [[45, 40, None, 50]] * 12
        revs = [[False, False, True, False]] * 12
        p = gem._normalizar_planilla(
            _planilla(areas, n_area_trabajo=4, revisar_s=revs)
        )
        self.assertEqual(p["estudiantes"][0]["revisar"],
                         [False, False, True, False])

    def test_celda_con_mas_queda_null_y_revisar(self):
        # Celda ilegible (símbolo "+"): null en su posición + revisar True,
        # con longitud correcta no hay marca extra de fila.
        areas = [[45, "+", 50, 60]] + [[40, 50, 60, 70]] * 11
        p = gem._normalizar_planilla(_planilla(areas, n_area_trabajo=4))
        est = p["estudiantes"][0]
        self.assertEqual(est["area_trabajo"], [45, None, 50, 60])
        self.assertTrue(est["revisar"][1])
        self.assertFalse(p["revisar_planilla"])

    def test_lista_larga_se_trunca_y_marca_toda_la_fila(self):
        # El modelo devolvió 5 valores para 4 columnas: sobran -> se trunca al
        # declarado y toda la fila queda en revisión.
        areas = [[45, 40, 50, 60, 70]] + [[40, 50, 60, 70]] * 11
        p = gem._normalizar_planilla(_planilla(areas, n_area_trabajo=4))
        est = p["estudiantes"][0]
        self.assertEqual(est["area_trabajo"], [45, 40, 50, 60])
        self.assertEqual(est["revisar"], [True, True, True, True])
        self.assertTrue(p["revisar_planilla"])

    def test_sin_declarado_comportamiento_previo_intacto(self):
        # Sin n_area_trabajo: longitud observada, sin marcas extra y sin
        # revisar_planilla por cantidad de notas.
        areas = [[45, 40, 50]] + [[40, 50, 60]] * 11
        p = gem._normalizar_planilla(_planilla(areas, n_area_trabajo=None))
        est = p["estudiantes"][0]
        self.assertEqual(est["area_trabajo"], [45, 40, 50])
        self.assertEqual(est["revisar"], [False, False, False])
        self.assertFalse(p["revisar_planilla"])


# ---------------------------------------------------------------------- #
# _normalizar_area directa con la firma nueva (n_area_declarado)
# ---------------------------------------------------------------------- #

class TestNormalizarAreaConDeclarado(unittest.TestCase):

    def test_lista_corta_rellena_con_none_y_marca_toda_la_fila(self):
        area, revisar = gem._normalizar_area([45, 40], [False, False], 4)
        self.assertEqual(area, [45, 40, None, None])
        self.assertEqual(revisar, [True, True, True, True])

    def test_lista_larga_trunca_y_marca_toda_la_fila(self):
        area, revisar = gem._normalizar_area(
            [45, 40, 50, 60, 70], [False] * 5, 4
        )
        self.assertEqual(area, [45, 40, 50, 60])
        self.assertEqual(revisar, [True, True, True, True])

    def test_sin_declarado_igual_que_antes(self):
        # Sin n_area_declarado (argumento omitido -> None): comportamiento
        # previo intacto (longitud observada, sin marcas extra).
        area, revisar = gem._normalizar_area([45, 40, 50], [False, False, False])
        self.assertEqual(area, [45, 40, 50])
        self.assertEqual(revisar, [False, False, False])

    def test_vacio_con_declarado_rellena_y_marca(self):
        # El modelo no trajo NADA para una planilla que declara 4 columnas.
        area, revisar = gem._normalizar_area(None, [], 4)
        self.assertEqual(area, [None, None, None, None])
        self.assertEqual(revisar, [True, True, True, True])


# ---------------------------------------------------------------------- #
# FORMATO POSICIONAL: area_trabajo como [{"col": N, "valor": X}]
# (contract v3). La posición de cada celda es explícita, así un colapso de la
# lista (celdas vacías omitidas) NO corre las notas de columna.
# ---------------------------------------------------------------------- #

def _planilla_posicional(areas_pos, n_area_trabajo, revisar_s=None):
    """Planilla cruda con formato posicional.

    areas_pos: lista de listas de dicts {"col": N, "valor": X}.
    revisar_s: flags "revisar" legacy (lista plana) si se quiere probar el
    camino de compatibilidad (cuando el modelo no trae "revisar" por objeto).
    """
    if areas_pos and isinstance(areas_pos[0], list):
        por_est = list(areas_pos)
    else:
        por_est = [list(areas_pos)] * 12
    encabezado = {
        "institucion": "I", "sede": "S", "año_lectivo": "2026",
        "jornada": "M", "grupo": "0302", "asignatura": "A",
        "docente": "D", "periodo": 3,
    }
    if n_area_trabajo is not None:
        encabezado["n_area_trabajo"] = n_area_trabajo
    estudiantes = []
    for i, area in enumerate(por_est, start=1):
        est = {
            "no": i, "nombre": f"ALUMNO {i}",
            "ev_anteriores": [45, 45],
            # El área puede venir como dicts posicionales o como lista plana
            # (formato mixto permitido): copiar tal cual para no mutar.
            "area_trabajo": (
                [dict(e) for e in area]
                if area and isinstance(area[0], dict)
                else list(area)
            ),
            "retirado": False,
        }
        if revisar_s is not None:
            est["revisar"] = (
                list(revisar_s[i - 1]) if isinstance(revisar_s[0], list)
                else list(revisar_s)
            )
        estudiantes.append(est)
    return {"encabezado": encabezado, "estudiantes": estudiantes}


class TestFormatoPosicional(unittest.TestCase):

    def test_celdas_vacias_omitidas_no_corren_posiciones(self):
        # Planilla de 4 columnas: "45 40 [vacía] 50". El modelo omite la vacía
        # (col 3) y las posiciones se reconstruyen exactas: el 50 queda en la
        # columna 4, NUNCA corrido a la columna de la vacía.
        areas = [[
            {"col": 1, "valor": 45},
            {"col": 2, "valor": 40},
            {"col": 4, "valor": 50},
        ]] + [[
            {"col": 1, "valor": 40},
            {"col": 2, "valor": 50},
            {"col": 3, "valor": 60},
            {"col": 4, "valor": 70},
        ]] * 11
        p = gem._normalizar_planilla(
            _planilla_posicional(areas, n_area_trabajo=4)
        )
        est = p["estudiantes"][0]
        self.assertEqual(est["area_trabajo"], [45, 40, None, 50])
        self.assertFalse(p["revisar_planilla"])
        # Sin marca extra de fila: la omisión de una celda vacía es el
        # comportamiento NORMAL del formato posicional, no un colapso.
        self.assertEqual(est["revisar"], [False, False, False, False])

    def test_primera_celda_vacia_omitida(self):
        # "[vacía] 50 50 50 50": col 1 omitida, posiciones exactas.
        areas = [[
            {"col": 2, "valor": 50},
            {"col": 3, "valor": 50},
            {"col": 4, "valor": 50},
            {"col": 5, "valor": 50},
        ]] + [[
            {"col": 1, "valor": 45},
            {"col": 2, "valor": 50},
            {"col": 3, "valor": 50},
            {"col": 4, "valor": 50},
            {"col": 5, "valor": 50},
        ]] * 11
        p = gem._normalizar_planilla(
            _planilla_posicional(areas, n_area_trabajo=5)
        )
        est = p["estudiantes"][0]
        self.assertEqual(est["area_trabajo"], [None, 50, 50, 50, 50])
        self.assertFalse(p["revisar_planilla"])

    def test_varias_vacias_internas_omitidas(self):
        areas = [[
            {"col": 1, "valor": 2},
            {"col": 2, "valor": 46},
            {"col": 4, "valor": 50},
            {"col": 5, "valor": 40},
        ]] + [[
            {"col": 1, "valor": 45},
            {"col": 2, "valor": 50},
            {"col": 3, "valor": 50},
            {"col": 4, "valor": 50},
            {"col": 5, "valor": 50},
        ]] * 11
        p = gem._normalizar_planilla(
            _planilla_posicional(areas, n_area_trabajo=5)
        )
        est = p["estudiantes"][0]
        self.assertEqual(est["area_trabajo"], [2, 46, None, 50, 40])
        self.assertFalse(p["revisar_planilla"])

    def test_null_explicito_con_col_conserva_posicion(self):
        # El modelo incluye la celda con null explícito: la posición se
        # conserva igual que si la omitiera.
        areas = [[
            {"col": 1, "valor": 48},
            {"col": 2, "valor": 45},
            {"col": 3, "valor": None},
            {"col": 4, "valor": 44},
        ]] * 12
        p = gem._normalizar_planilla(
            _planilla_posicional(areas, n_area_trabajo=4)
        )
        est = p["estudiantes"][0]
        self.assertEqual(est["area_trabajo"], [48, 45, None, 44])
        self.assertFalse(p["revisar_planilla"])

    def test_celda_ilegible_valor_null_y_revisar_true(self):
        # Celda ilegible ("+"): null con su col y revisar en esa posición.
        areas = [[
            {"col": 1, "valor": 45},
            {"col": 2, "valor": None, "revisar": True},
            {"col": 3, "valor": 50},
            {"col": 4, "valor": 60},
        ]] * 12
        p = gem._normalizar_planilla(
            _planilla_posicional(areas, n_area_trabajo=4)
        )
        est = p["estudiantes"][0]
        self.assertEqual(est["area_trabajo"], [45, None, 50, 60])
        self.assertTrue(est["revisar"][1])
        self.assertFalse(p["revisar_planilla"])

    def test_col_fuera_de_rango_marca_planilla_y_toda_la_fila(self):
        # El modelo devuelve una col 5 con n_area_trabajo=4 declarado: sobran
        # valores -> revisar_planilla y la fila completa en revisión (igual que
        # el colapso plano: el corrimiento es sospechoso).
        areas = [[
            {"col": 1, "valor": 45},
            {"col": 2, "valor": 40},
            {"col": 4, "valor": 50},
            {"col": 5, "valor": 60},
        ]] + [[
            {"col": 1, "valor": 40},
            {"col": 2, "valor": 50},
            {"col": 3, "valor": 60},
            {"col": 4, "valor": 70},
        ]] * 11
        p = gem._normalizar_planilla(
            _planilla_posicional(areas, n_area_trabajo=4)
        )
        est = p["estudiantes"][0]
        # Los valores que caen dentro del declarado se conservan; el de la col
        # fuera de rango no tiene posición válida.
        self.assertTrue(p["revisar_planilla"])
        self.assertTrue(all(est["revisar"]))
        self.assertEqual(len(est["area_trabajo"]), 4)

    def test_col_duplicada_marca_toda_la_fila(self):
        areas = [[
            {"col": 2, "valor": 45},
            {"col": 2, "valor": 40},
            {"col": 4, "valor": 50},
        ]] + [[
            {"col": 1, "valor": 40},
            {"col": 2, "valor": 50},
            {"col": 3, "valor": 60},
            {"col": 4, "valor": 70},
        ]] * 11
        p = gem._normalizar_planilla(
            _planilla_posicional(areas, n_area_trabajo=4)
        )
        est = p["estudiantes"][0]
        self.assertTrue(all(est["revisar"]))
        self.assertTrue(p["revisar_planilla"])

    def test_sin_declarado_usar_max_col(self):
        # Sin n_area_trabajo: el ancho se infiere del col más alto observado.
        areas = [[
            {"col": 1, "valor": 2},
            {"col": 2, "valor": 46},
            {"col": 4, "valor": 50},
        ]] * 12
        p = gem._normalizar_planilla(
            _planilla_posicional(areas, n_area_trabajo=None)
        )
        est = p["estudiantes"][0]
        self.assertEqual(est["area_trabajo"], [2, 46, None, 50])

    def test_revisar_legacy_plano_se_alinea_por_posicion(self):
        # Compatibilidad: formato posicional pero "revisar" viene como lista
        # plana legacy (longitud observada de area_trabajo, sin dict). El flag
        # de la posición 2 marca la col 2 aunque el área venga posicional.
        areas = [[
            {"col": 1, "valor": 45},
            {"col": 2, "valor": None},
            {"col": 3, "valor": 50},
            {"col": 4, "valor": 60},
        ]] * 12
        revs = [False, True, False, False]
        p = gem._normalizar_planilla(
            _planilla_posicional(areas, n_area_trabajo=4, revisar_s=revs)
        )
        est = p["estudiantes"][0]
        self.assertEqual(est["area_trabajo"], [45, None, 50, 60])
        self.assertTrue(est["revisar"][1])
        self.assertFalse(p["revisar_planilla"])

    def test_prompt_pide_formato_posicional(self):
        texto = gem._PROMPT_PLANILLA
        # El contrato v3: cada celda con su columna explícita.
        self.assertIn('"col"', texto)
        self.assertIn('"valor"', texto)
        # La posición NUNCA se infiere por el orden de una lista compactada.
        self.assertIn("compact", texto.lower())

    def test_formato_mixto_por_estudiante(self):
        # Cada estudiante resuelve independientemente: posicional y plano
        # pueden convivir en la misma planilla (el modelo no siempre cambia de
        # formato a la vez).
        areas = [
            [{"col": 1, "valor": 45}, {"col": 3, "valor": 50}],
            [45, 40, 50, 60],
        ] * 6
        p = gem._normalizar_planilla(
            _planilla_posicional(areas, n_area_trabajo=4)
        )
        self.assertEqual(p["estudiantes"][0]["area_trabajo"],
                         [45, None, 50, None])
        self.assertEqual(p["estudiantes"][1]["area_trabajo"],
                         [45, 40, 50, 60])
        self.assertFalse(p["revisar_planilla"])


# ---------------------------------------------------------------------- #
# Guardas de regresión del prompt (regla anti-colapso + celdas ilegibles)
# ---------------------------------------------------------------------- #

class TestPromptAntiColapso(unittest.TestCase):

    def test_prompt_tiene_regla_anticolapso_con_ejemplo(self):
        texto = gem._PROMPT_PLANILLA
        # La regla anti-colapso (ahora en formato posicional) está presente.
        self.assertIn("NUNCA compactes", texto)
        # El ejemplo concreto con la celda vacía en el medio y las cols reales.
        self.assertIn('{"col": 1, "valor": 45}', texto)
        self.assertIn('{"col": 4, "valor": 50}', texto)
        self.assertIn('el siguiente objeto lleva "col": 4', texto)
        # Celdas ilegibles: null en su posición + revisar, nunca inventar.
        self.assertIn("celda es ilegible", texto)
        self.assertIn("nunca inventes un número", texto)


if __name__ == "__main__":
    unittest.main()