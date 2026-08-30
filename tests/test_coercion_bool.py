"""
Test de la coerción explícita de booleanos (W6).

Cubre el bug donde `bool("false")` daba True: un estudiante "retirado" con el
valor "false" se perdía del cálculo. La nueva función `_coercion_bool` de
vision.gemini_extractor debe mapear correctamente los valores leídos del modelo.

Uso (desde la raíz del proyecto):
    python -m tests.test_coercion_bool
"""

import os
import sys
import unittest

_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

from vision.gemini_extractor import _coercion_bool, _normalizar_planilla  # noqa: E402


class TestCoercionBool(unittest.TestCase):

    def test_string_false_no_es_true(self):
        self.assertFalse(_coercion_bool("false")[0])

    def test_string_true_es_true(self):
        self.assertTrue(_coercion_bool("true")[0])

    def test_variantes_afirmativas(self):
        for v in ("true", "TRUE", "1", "sí", "si", "Si", "yes", 1, 1.0):
            self.assertTrue(_coercion_bool(v)[0], f"debería ser True: {v!r}")

    def test_variantes_negativas(self):
        for v in ("false", "FALSE", "0", "no", "No", 0, 0.0, "", None):
            self.assertFalse(_coercion_bool(v)[0], f"debería ser False: {v!r}")

    def test_valor_no_reconocido_se_marca_para_revision(self):
        # No se adivina: valor raro -> confiable=False (marcar revisar).
        result, confiable = _coercion_bool("quizás")
        self.assertFalse(confiable)

    def test_retirado_false_no_desaparece_del_calculo(self):
        # Bug W6: bool("false") era True y sacaba al alumno del cálculo.
        datos = {
            "encabezado": {
                "institucion": "I", "sede": "S", "año_lectivo": "2026",
                "jornada": "M", "grupo": "0302", "asignatura": "A",
                "docente": "D", "periodo": 3,
            },
            "estudiantes": [
                {
                    "no": 1, "nombre": "ALUMNO", "ev_anteriores": [45, 45],
                    "area_trabajo": [40, 50], "retirado": "false",
                    "revisar": ["false", "false"],
                }
            ],
        }
        planilla = _normalizar_planilla(datos)
        est = planilla["estudiantes"][0]
        self.assertFalse(est["retirado"], "retirado='false' no debe marcar al alumno como retirado")
        self.assertEqual(est["revisar"], [False, False])

    def test_retirado_true_se_marca(self):
        datos = {
            "encabezado": {
                "institucion": "I", "sede": "S", "año_lectivo": "2026",
                "jornada": "M", "grupo": "0302", "asignatura": "A",
                "docente": "D", "periodo": 3,
            },
            "estudiantes": [
                {
                    "no": 1, "nombre": "ALUMNO", "ev_anteriores": [],
                    "area_trabajo": None, "retirado": "true",
                    "revisar": [False, False],
                }
            ],
        }
        planilla = _normalizar_planilla(datos)
        self.assertTrue(planilla["estudiantes"][0]["retirado"])


if __name__ == "__main__":
    unittest.main()
