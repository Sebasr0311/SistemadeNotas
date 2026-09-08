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
# Guardas de regresión del prompt (regla anti-colapso + celdas ilegibles)
# ---------------------------------------------------------------------- #

class TestPromptAntiColapso(unittest.TestCase):

    def test_prompt_tiene_regla_anticolapso_con_ejemplo(self):
        texto = gem._PROMPT_PLANILLA
        self.assertIn("REGLA ANTI-COMPACTACIÓN", texto)
        # El ejemplo concreto: [45, 40, null, 50] y los dos colapsos prohibidos.
        self.assertIn("[45, 40, null, 50]", texto)
        self.assertIn("[45, 40, 50]", texto)
        self.assertIn("n_area_trabajo para TODOS los alumnos", texto)
        # Celdas ilegibles: null en su posición + revisar, nunca inventar.
        self.assertIn("celda es ilegible", texto)
        self.assertIn("nunca inventes un número", texto)


if __name__ == "__main__":
    unittest.main()