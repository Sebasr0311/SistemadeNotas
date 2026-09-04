"""
Tests de integridad de datos (S2, S3, S6) de la auditoría.

S6: los dígitos separados ("4 5") NUNCA se fusionan en un número inventado.
S3: las Ev. Anteriores fuera de rango o ambiguas se marcan para revisión.
S2: una página fallida no tira el lote completo del PDF.

Uso (desde la raíz del proyecto):
    python -m unittest discover -s tests -v
"""

import os
import sys
import unittest
from unittest import mock

_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

from vision import gemini_extractor  # noqa: E402
from vision.gemini_extractor import (  # noqa: E402
    _numero_plausible,
    _normalizar_ev,
    _normalizar_planilla,
    extraer_planilla_pdf,
    VisionError,
    _PlanillaParseError,
)


class TestNumeroPlausible(unittest.TestCase):
    """S6: no fusionar dígitos separados; decimales españoles tolerados."""

    def test_entero(self):
        self.assertEqual(_numero_plausible("45"), (45.0, True))

    def test_decimal_con_coma(self):
        self.assertEqual(_numero_plausible("4,5"), (4.5, True))

    def test_decimal_con_espacio_alrededor_del_punto(self):
        self.assertEqual(_numero_plausible("4 ,5"), (4.5, True))

    def test_digitos_separados_no_se_fusionan(self):
        self.assertEqual(_numero_plausible("4 5"), (None, False))

    def test_digitos_separados_largos_no_se_fusionan(self):
        self.assertEqual(_numero_plausible("12 3"), (None, False))

    def test_espacios_alrededor_ok(self):
        self.assertEqual(_numero_plausible(" 45 "), (45.0, True))

    def test_none_es_confiable(self):
        self.assertEqual(_numero_plausible(None), (None, True))


class TestNormalizarEv(unittest.TestCase):
    """S3: ev_anteriores fuera de rango o ambigua marca revisar_ev."""

    def test_fuera_de_rango_superior_marca(self):
        ev, revisar = _normalizar_ev([150], 1)
        self.assertEqual(ev, [150.0])
        self.assertTrue(revisar)

    def test_negativo_marca(self):
        _, revisar = _normalizar_ev([-3], 1)
        self.assertTrue(revisar)

    def test_en_rango_es_confiable(self):
        ev, revisar = _normalizar_ev([4.5], 1)
        self.assertEqual(ev, [4.5])
        self.assertFalse(revisar)

    def test_planilla_con_ev_150_marca_revisar_ev(self):
        datos = {
            "encabezado": {"periodo": 2, "grupo": "0301", "asignatura": "LENGUA"},
            "estudiantes": [
                {"no": 1, "nombre": "PEREZ JUAN", "retirado": False,
                 "ev_anteriores": [150], "area_trabajo": [40, 50]}
            ],
        }
        planilla = _normalizar_planilla(datos)
        est = planilla["estudiantes"][0]
        self.assertTrue(est["revisar_ev"])
        self.assertEqual(est["ev_anteriores"], [150.0])

    def test_planilla_con_ev_normal_no_marca(self):
        datos = {
            "encabezado": {"periodo": 2, "grupo": "0301", "asignatura": "LENGUA"},
            "estudiantes": [
                {"no": 1, "nombre": "PEREZ JUAN", "retirado": False,
                 "ev_anteriores": [45], "area_trabajo": [40, 50]}
            ],
        }
        planilla = _normalizar_planilla(datos)
        self.assertFalse(planilla["estudiantes"][0]["revisar_ev"])

    def test_planilla_retirado_no_marca_revisar_ev(self):
        datos = {
            "encabezado": {"periodo": 2, "grupo": "0301", "asignatura": "LENGUA"},
            "estudiantes": [
                {"no": 2, "nombre": "**** GOMEZ ANA", "retirado": False,
                 "ev_anteriores": [150], "area_trabajo": None}
            ],
        }
        planilla = _normalizar_planilla(datos)
        est = planilla["estudiantes"][0]
        self.assertTrue(est["retirado"])
        self.assertFalse(est["revisar_ev"])


class TestExtraerPlanillaPdf(unittest.TestCase):
    """S2: páginas fallidas no tiran el lote; excepciones raras propagan."""

    @classmethod
    def setUpClass(cls):
        # Fuga de log: estos tests fuerzan páginas fallidas a propósito y
        # extraer_planilla_pdf registra cada una en el log REAL del usuario
        # (%APPDATA%\\SistemaNotas\\log.txt). Se neutraliza escribiendo en el
        # vacío para TODA la clase, así la suite no contamina el diagnóstico
        # de corridas reales.
        cls._parche_log = mock.patch.object(
            gemini_extractor.app_config, "escribir_log", return_value=None
        )
        cls._parche_log.start()
        cls.addClassCleanup(cls._parche_log.stop)

    def _paginas(self, n):
        return [{"imagen": object()} for _ in range(n)]

    def _planilla_ok(self):
        return {"encabezado": {"periodo": 3, "grupo": "0301"}, "estudiantes": []}

    @mock.patch.object(gemini_extractor, "extraer_planilla_pagina")
    def test_una_fallida_no_tira_el_lote(self, mock_pagina):
        ok = self._planilla_ok()
        mock_pagina.side_effect = [ok, VisionError("falla"), ok]
        planillas, fallidas = extraer_planilla_pdf(self._paginas(3), api_key="x")
        self.assertEqual(len(planillas), 2)
        self.assertEqual(len(fallidas), 1)
        self.assertEqual(fallidas[0]["pagina"], 2)
        self.assertEqual(fallidas[0]["total"], 3)
        self.assertEqual(fallidas[0]["tipo"], "vision")

    @mock.patch.object(gemini_extractor, "extraer_planilla_pagina")
    def test_parse_error_se_registra_y_sigue(self, mock_pagina):
        ok = self._planilla_ok()
        mock_pagina.side_effect = [_PlanillaParseError("no interpretable"), ok]
        planillas, fallidas = extraer_planilla_pdf(self._paginas(2), api_key="x")
        self.assertEqual(len(planillas), 1)
        self.assertEqual(len(fallidas), 1)
        self.assertEqual(fallidas[0]["tipo"], "parseo")

    @mock.patch.object(gemini_extractor, "extraer_planilla_pagina")
    def test_todas_fallan_devuelve_vacio_sin_excepcion(self, mock_pagina):
        mock_pagina.side_effect = VisionError("falla")
        planillas, fallidas = extraer_planilla_pdf(self._paginas(3), api_key="x")
        self.assertEqual(planillas, [])
        self.assertEqual([f["pagina"] for f in fallidas], [1, 2, 3])

    @mock.patch.object(gemini_extractor, "extraer_planilla_pagina")
    def test_excepcion_no_esperada_propaga(self, mock_pagina):
        ok = self._planilla_ok()
        mock_pagina.side_effect = [ok, RuntimeError("boom")]
        with self.assertRaises(RuntimeError):
            extraer_planilla_pdf(self._paginas(2), api_key="x")


if __name__ == "__main__":
    unittest.main()