"""
Tests de _normalizar_area (especificación v2, 1 a 16 notas por alumno).

El bug confirmado: _normalizar_area estaba hardcodeada a 2 celdas, así que TODA
planilla salía con exactamente 2 notas aunque el modelo leyera 5, 10 o 16. Esta
suite verifica que ahora se devuelven TANTAS celdas como valores traiga el
modelo (hasta MAX_N_AREAS = 16), con los flags de "revisar" alineados al mismo
largo y las reglas por celda intactas (blanco -> None sin revisar, ambiguo ->
None + revisar, fuera de rango 0-100 -> dejar valor + revisar).

También cubre el pipeline/alto nivel: una planilla con n_area_trabajo=5 debe
normalizarse a 5 notas por alumno y el circuito completo (Gemini mockeado ->
normalización -> Excel) debe correr sin romperse.

Uso (desde la raíz del proyecto):
    python -m unittest discover -s tests -v
"""

import json
import os
import sys
import tempfile
import unittest
from unittest import mock

_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

import vision.gemini_extractor as gem  # noqa: E402


# ---------------------------------------------------------------------- #
# (a-e) _normalizar_area: cantidad de celdas dinámica y alineación de flags
# ---------------------------------------------------------------------- #

class TestNormalizarArea(unittest.TestCase):

    def test_cinco_valores_devuelve_cinco_celdas(self):
        # 5 notas -> 5 celdas (no 2, como el bug original).
        valores = [40, 50, 60, 70, 80]
        area, revisar = gem._normalizar_area(valores, [])
        self.assertEqual(len(area), 5)
        self.assertEqual(len(revisar), 5)
        self.assertEqual(area, [40, 50, 60, 70, 80])

    def test_dieciséis_valores_devuelve_dieciséis_celdas(self):
        valores = list(range(1, 17))
        area, revisar = gem._normalizar_area(valores, [])
        self.assertEqual(len(area), gem.MAX_N_AREAS)
        self.assertEqual(len(revisar), gem.MAX_N_AREAS)
        self.assertEqual(area, valores)

    def test_diecisiete_valores_se_trunca_a_dieciséis(self):
        # Tope de seguridad: nunca más de MAX_N_AREAS (16).
        valores = list(range(1, 18))
        area, revisar = gem._normalizar_area(valores, [])
        self.assertEqual(len(area), gem.MAX_N_AREAS)
        self.assertEqual(len(revisar), gem.MAX_N_AREAS)
        self.assertEqual(area, valores[:gem.MAX_N_AREAS])

    def test_valores_vacios_devuelve_listas_vacias(self):
        self.assertEqual(gem._normalizar_area([], []), ([], []))
        self.assertEqual(gem._normalizar_area(None, []), ([], []))

    def test_blanco_intercalado_queda_none_sin_revisar(self):
        # Celda en blanco (None) -> None, sin marcar revisión.
        valores = [40, None, 50]
        area, revisar = gem._normalizar_area(valores, [])
        self.assertEqual(area, [40, None, 50])
        self.assertEqual(revisar, [False, False, False])

    def test_flags_mas_cortos_que_valores_se_rellenan_con_false(self):
        # Solamente se marcan las celdas cubiertas por los flags; el resto False.
        valores = [40, 50, 60, 70]
        area, revisar = gem._normalizar_area(valores, [True, False])
        self.assertEqual(len(revisar), 4)
        self.assertEqual(revisar, [True, False, False, False])

    def test_flags_mas_largos_que_valores_se_truncan(self):
        # Si el modelo trae más flags que valores, sobran: se ignoran.
        valores = [40, 50]
        area, revisar = gem._normalizar_area(valores, [True, True, True])
        self.assertEqual(len(revisar), 2)
        self.assertEqual(revisar, [True, True])

    def test_valor_ambiguo_digitos_separados_marca_revision(self):
        # "4 5" (dígitos muy separados, S6) -> None + revisar True en esa celda.
        valores = ["4 5", 50]
        area, revisar = gem._normalizar_area(valores, [False, False])
        self.assertIsNone(area[0])
        self.assertEqual(revisar[0], True)
        self.assertEqual(area[1], 50)
        self.assertEqual(revisar[1], False)

    def test_fuera_de_rango_deja_valor_y_marca_revision(self):
        # Nota fuera de 0-100 (ej. 150) -> se conserva y se marca revisar.
        valores = [150, 50]
        area, revisar = gem._normalizar_area(valores, [False, False])
        self.assertEqual(area[0], 150)
        self.assertEqual(revisar[0], True)
        self.assertEqual(area[1], 50)


# ---------------------------------------------------------------------- #
# (a-f) _normalizar_area vía planilla: revisar del modelo alineado dinámico
# ---------------------------------------------------------------------- #

class TestNormalizarAreaEnPlanilla(unittest.TestCase):

    def _planilla_cruda(self, area_trabajo, n_area_trabajo, revisar=None):
        """Una planilla cruda (1 alumno) con la cantidad de áreas declarada."""
        encabezado = {
            "institucion": "I", "sede": "S", "año_lectivo": "2026",
            "jornada": "M", "grupo": "0302", "asignatura": "A",
            "docente": "D", "periodo": 3, "n_area_trabajo": n_area_trabajo,
        }
        est = {
            "no": 1, "nombre": "ALUMNO 1",
            "ev_anteriores": [45, 45],
            "area_trabajo": area_trabajo, "retirado": False,
        }
        if revisar is not None:
            est["revisar"] = revisar
        return {"encabezado": encabezado, "estudiantes": [est]}

    def test_cinco_notas_por_alumno_normalizan_cinco_celdas(self):
        # n_area_trabajo=5 coincide con las 5 notas leídas -> sin flag.
        planilla = gem._normalizar_planilla(
            self._planilla_cruda([40, 50, 60, 70, 80], n_area_trabajo=5)
        )
        est = planilla["estudiantes"][0]
        self.assertEqual(est["area_trabajo"], [40, 50, 60, 70, 80])
        self.assertEqual(len(est["revisar"]), 5)

    def test_revisar_del_modelo_alineado_a_cinco_celdas(self):
        # El modelo marca la celda 2 de 5 -> flags de largo dinámico.
        planilla = gem._normalizar_planilla(
            self._planilla_cruda(
                [40, 50, 60, 70, 80], n_area_trabajo=5,
                revisar=[False, False, "true", False, False],
            )
        )
        est = planilla["estudiantes"][0]
        self.assertEqual(est["revisar"], [False, False, True, False, False])
        self.assertEqual(len(est["revisar"]), 5)

    def test_retirado_no_lleva_revisar_fijo_de_dos(self):
        # Retirado -> "revisar" vacío (no [False, False] fijo de 2).
        ret = gem._normalizar_planilla({
            "encabezado": {
                "institucion": "I", "sede": "S", "año_lectivo": "2026",
                "jornada": "M", "grupo": "0302", "asignatura": "A",
                "docente": "D", "periodo": 3,
                "n_area_trabajo": 5,
            },
            "estudiantes": [{
                "no": 1, "nombre": "ALUMNO 1", "ev_anteriores": [],
                "area_trabajo": None, "retirado": True,
            }],
        })
        self.assertTrue(ret["estudiantes"][0]["retirado"])
        self.assertEqual(ret["estudiantes"][0]["revisar"], [])


# ---------------------------------------------------------------------- #
# Pipeline/alto nivel: 5 notas por alumno terminan en el Excel real
# ---------------------------------------------------------------------- #

class _FakeResponse:
    def __init__(self, texto):
        self.text = texto


class _FakeModels:
    def __init__(self, respuestas):
        self._respuestas = respuestas
        self._contador = 0

    def generate_content(self, **kwargs):
        texto = self._respuestas[self._contador]
        self._contador += 1
        return _FakeResponse(texto)


class _FakeClient:
    def __init__(self, respuestas):
        self.models = _FakeModels(respuestas)


class TestPipelineCincoNotas(unittest.TestCase):

    def test_planilla_con_cinco_notas_corre_el_circuito_completo(self):
        # Planilla cruda con n_area_trabajo=5 y 5 notas por alumno (2 alumnos:
        # dentro del rango sano de alumnos para no marcar revisar_planilla).
        planilla_cruda = {
            "encabezado": {
                "institucion": "INST", "sede": "SEDE", "año_lectivo": "2026",
                "jornada": "M", "grupo": "0302", "asignatura": "MATES",
                "docente": "DOC", "periodo": 3, "n_area_trabajo": 5,
            },
            "estudiantes": [
                {"no": 1, "nombre": "GARCIA LUZ", "ev_anteriores": [45, 45],
                 "area_trabajo": [40, 50, 60, 70, 80], "retirado": False,
                 "revisar": [False, False, False, False, False]},
                {"no": 2, "nombre": "PEREZ JUAN", "ev_anteriores": [45, 45],
                 "area_trabajo": [35, 45, 55, 65, 75], "retirado": False,
                 "revisar": [False, False, False, False, False]},
                {"no": 3, "nombre": "LOPEZ ANA", "ev_anteriores": [45, 45],
                 "area_trabajo": [30, 40, 50, 60, 70], "retirado": False,
                 "revisar": [False, False, False, False, False]},
                {"no": 4, "nombre": "DIAZ MAR", "ev_anteriores": [45, 45],
                 "area_trabajo": [25, 35, 45, 55, 65], "retirado": False,
                 "revisar": [False, False, False, False, False]},
                {"no": 5, "nombre": "MORA RITA", "ev_anteriores": [45, 45],
                 "area_trabajo": [20, 30, 40, 50, 60], "retirado": False,
                 "revisar": [False, False, False, False, False]},
                {"no": 6, "nombre": "RUIZ SOL", "ev_anteriores": [45, 45],
                 "area_trabajo": [10, 20, 30, 40, 50], "retirado": False,
                 "revisar": [False, False, False, False, False]},
                {"no": 7, "nombre": "SOTO NINA", "ev_anteriores": [45, 45],
                 "area_trabajo": [15, 25, 35, 45, 55], "retirado": False,
                 "revisar": [False, False, False, False, False]},
                {"no": 8, "nombre": "VEGA LEO", "ev_anteriores": [45, 45],
                 "area_trabajo": [5, 15, 25, 35, 45], "retirado": False,
                 "revisar": [False, False, False, False, False]},
                {"no": 9, "nombre": "CRUZ IDA", "ev_anteriores": [45, 45],
                 "area_trabajo": [8, 18, 28, 38, 48], "retirado": False,
                 "revisar": [False, False, False, False, False]},
                {"no": 10, "nombre": "RIVAS KAI", "ev_anteriores": [45, 45],
                 "area_trabajo": [12, 22, 32, 42, 52], "retirado": False,
                 "revisar": [False, False, False, False, False]},
            ],
        }
        respuestas = [json.dumps(planilla_cruda, ensure_ascii=False)]

        dir_tmp = tempfile.mkdtemp(prefix="notas_5_")
        try:
            # Lo único mockeado: la conexión HTTP a Gemini.
            with mock.patch.object(gem, "_cliente", return_value=_FakeClient(respuestas)):
                planillas, fallidas = gem.extraer_planilla_pdf(
                    [{"imagen": None, "indice": 1}], api_key="clave-de-prueba"
                )
            self.assertEqual(fallidas, [])
            self.assertEqual(len(planillas), 1)

            # La normalización conserva las 5 notas por alumno.
            est0 = planillas[0]["estudiantes"][0]
            self.assertEqual(est0["area_trabajo"], [40, 50, 60, 70, 80])
            self.assertEqual(len(est0["revisar"]), 5)
            # La validación de planilla cierra: 5 == n_area_trabajo y la
            # cantidad de alumnos (10) está dentro del rango sano.
            self.assertFalse(planillas[0]["revisar_planilla"])

            # El circuito termina en el Excel real: debe generar sin romperse
            # y con la hoja del curso.
            ruta_xlsx = os.path.join(dir_tmp, "notas_5.xlsx")
            from excel.generar_excel_notas import generar_excel_asignatura
            generar_excel_asignatura(planillas, ruta_xlsx)

            import openpyxl
            wb = openpyxl.load_workbook(ruta_xlsx)
            self.assertEqual(wb.sheetnames, ["Curso 0302 - MATES"])
        finally:
            import shutil
            shutil.rmtree(dir_tmp, ignore_errors=True)


class TestNAreaTrabajoEncabezado(unittest.TestCase):
    """S13: el n_area_trabajo declarado se conserva en el encabezado normalizado
    y el prompt lo pide explícitamente (el modelo contaba celdas llenas, no
    columnas del encabezado: planillas con celdas vacías salían con menos de
    las notas reales)."""

    def test_encabezado_conserva_declarado(self):
        p = gem._normalizar_planilla(
            self._planilla([40, 50, 60], n_area_trabajo=4)
        )
        self.assertEqual(p["encabezado"]["n_area_trabajo"], 4)

    def test_encabezado_sin_declarado_queda_none(self):
        p = gem._normalizar_planilla(
            self._planilla([40, 50, 60], n_area_trabajo=None)
        )
        self.assertIsNone(p["encabezado"]["n_area_trabajo"])

    def test_declarado_bool_no_es_conteo(self):
        # Un bool NO es un conteo: se conserva como None (nunca 1 ni 4).
        p = gem._normalizar_planilla(
            self._planilla([40], n_area_trabajo=True)
        )
        self.assertIsNone(p["encabezado"]["n_area_trabajo"])

    def test_prompt_pide_n_area_trabajo_contando_encabezado(self):
        # Guarda de regresión: el prompt debe pedir el campo y la regla de
        # contar columnas del ENCABEZADO aunque las celdas estén vacías.
        self.assertIn("n_area_trabajo", gem._PROMPT_PLANILLA)
        self.assertIn("EXACTAMENTE n_area_trabajo", gem._PROMPT_PLANILLA)
        self.assertIn("columnas del encabezado", gem._PROMPT_PLANILLA)

    def test_prompt_incluye_refuerzos_nuevos(self):
        # Guardas de regresión (S16): membrete, rótulos cortados, apuntes
        # debajo de la tabla, fila retirado con asteriscos en todas las celdas,
        # Min/Fls. ignoradas aunque traigan datos, y cantidades de columnas
        # independientes entre planillas del mismo curso. Substrings sin
        # acentos para evitar problemas de encoding.
        self.assertIn("LISTA AUXILIAR DE CLASE", gem._PROMPT_PLANILLA)
        self.assertIn("Firma Docente", gem._PROMPT_PLANILLA)
        self.assertIn("Nunca asumas", gem._PROMPT_PLANILLA)
        self.assertIn("Min", gem._PROMPT_PLANILLA)
        self.assertIn("Fls", gem._PROMPT_PLANILLA)
        self.assertIn("retirado", gem._PROMPT_PLANILLA)
        self.assertIn("membrete", gem._PROMPT_PLANILLA)
        self.assertIn("CamScanner", gem._PROMPT_PLANILLA)

    def _planilla(self, area_trabajo, n_area_trabajo):
        return {
            "encabezado": {
                "institucion": "I", "sede": "S", "año_lectivo": "2026",
                "jornada": "M", "grupo": "0302", "asignatura": "A",
                "docente": "D", "periodo": 3, "n_area_trabajo": n_area_trabajo,
            },
            "estudiantes": [
                {"no": 1, "nombre": "ALUMNO 1", "ev_anteriores": [45, 45],
                 "area_trabajo": area_trabajo, "retirado": False,
                 "revisar": [False] * 10},
            ],
        }


if __name__ == "__main__":
    unittest.main()
