"""
Tests de validación del spec v2.

DELTA 1 — confianza explícita del modelo: el campo "revisar" que devuelve el
modelo se adopta por celda (con coerciones booleanas explícitas, _coercion_bool)
y se UNE con la heurística existente de _normalizar_area: una celda queda en
revisión si el modelo la marcó O la heurística la marcó. Si el modelo no trae
"revisar", se usa sólo la heurística.

DELTA 2 — validación fuerte de planilla: si el encabezado declara
n_area_trabajo y algún estudiante no retirado trae otra cantidad de notas, o si
la cantidad de alumnos no retirados cae fuera de [MIN_ESTUDIANTES_SANOS,
MAX_ESTUDIANTES_SANOS], la planilla se marca con revisar_planilla=True pero NO
se descarta: sigue en `planillas` (nunca en `fallidas`).

Uso (desde la raíz del proyecto):
    python -m unittest discover -s tests -v
"""

import json
import os
import sys
import unittest
from unittest import mock

_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

import vision.gemini_extractor as gem  # noqa: E402


# ---------------------------------------------------------------------- #
# Fixtures
# ---------------------------------------------------------------------- #

def _estudiante(no, area_trabajo=None, revisar=None, retirado=False):
    """Estudiante crudo (formato que devuelve el modelo)."""
    est = {
        "no": no,
        "nombre": f"ALUMNO {no}",
        "ev_anteriores": [45, 45],
        "area_trabajo": area_trabajo if area_trabajo is not None else [40, 50],
        "retirado": retirado,
    }
    if revisar is not None:
        est["revisar"] = revisar
    return est


def _planilla_cruda(estudiantes, n_area_trabajo=None, periodo=3):
    """Planilla cruda (encabezado + estudiantes), opcional n_area_trabajo."""
    encabezado = {
        "institucion": "I", "sede": "S", "año_lectivo": "2026",
        "jornada": "M", "grupo": "0302", "asignatura": "A",
        "docente": "D", "periodo": periodo,
    }
    if n_area_trabajo is not None:
        encabezado["n_area_trabajo"] = n_area_trabajo
    return {"encabezado": encabezado, "estudiantes": estudiantes}


def _doce_alumnos():
    """12 alumnos con 2 notas cada uno (cantidad dentro del rango sano)."""
    return [_estudiante(n) for n in range(1, 13)]


class _FakeResponse:
    def __init__(self, texto):
        self.text = texto


class _FakeModels:
    """Devuelve una planilla distinta por llamada (una por página del PDF)."""

    def __init__(self, respuestas: list):
        self._respuestas = respuestas
        self._contador = 0

    def generate_content(self, **kwargs):
        if self._contador >= len(self._respuestas):
            raise AssertionError("generate_content se llamó más veces que páginas")
        texto = self._respuestas[self._contador]
        self._contador += 1
        return _FakeResponse(texto)


class _FakeClient:
    def __init__(self, respuestas: list):
        self.models = _FakeModels(respuestas)


# ---------------------------------------------------------------------- #
# (a) DELTA 1: "revisar" del modelo, adoptado por celda y unido con heurística
# ---------------------------------------------------------------------- #

class TestRevisarDelModeloYHeuristica(unittest.TestCase):

    def test_revisar_del_modelo_se_adopta_por_celda(self):
        # El modelo marca la celda 0 y no la 1: el resultado respeta esa marca
        # ("true"/"false" como strings, que _coercion_bool interpreta bien).
        estudiantes = [_estudiante(n, revisar=["true", "false"]) for n in range(1, 13)]
        planilla = gem._normalizar_planilla(_planilla_cruda(estudiantes))
        est = planilla["estudiantes"][0]
        self.assertEqual(est["revisar"], [True, False])

    def test_revisar_con_enteros_y_variantes_afirmativas(self):
        # 1/sí -> True; 0/no -> False (coerción explícita por celda).
        estudiantes = [
            _estudiante(1, revisar=[1, 0]),
            _estudiante(2, revisar=["sí", "no"]),
        ]
        planilla = gem._normalizar_planilla(_planilla_cruda(estudiantes))
        self.assertEqual(planilla["estudiantes"][0]["revisar"], [True, False])
        self.assertEqual(planilla["estudiantes"][1]["revisar"], [True, False])

    def test_union_con_heuristica_fuera_de_rango(self):
        # El modelo dice "revisar": false, pero 150 queda fuera de rango
        # (0-100): la heurística lo marca igual -> UNIÓN (O lógico).
        estudiantes = [_estudiante(1, area_trabajo=[150, 50], revisar=["false", "false"])]
        planilla = gem._normalizar_planilla(_planilla_cruda(estudiantes))
        self.assertEqual(planilla["estudiantes"][0]["revisar"], [True, False])

    def test_union_con_heuristica_digitos_ambiguos(self):
        # "4 5" (dígitos muy separados, S6) tampoco se lee: la heurística marca
        # revisión aunque el modelo haya dicho false, y el valor queda en null.
        estudiantes = [_estudiante(1, area_trabajo=["4 5", 50], revisar=["false", "false"])]
        planilla = gem._normalizar_planilla(_planilla_cruda(estudiantes))
        est = planilla["estudiantes"][0]
        self.assertEqual(est["revisar"], [True, False])
        self.assertIsNone(est["area_trabajo"][0])

    def test_sin_revisar_del_modelo_queda_solo_la_heuristica(self):
        # Sin el campo "revisar" en el JSON crudo, no se rompe: la heurística
        # sigue marcando lo ambiguo y dejando sano lo sano.
        ambiguos = [_estudiante(1, area_trabajo=["4 5", 50])]
        planilla = gem._normalizar_planilla(_planilla_cruda(ambiguos))
        self.assertEqual(planilla["estudiantes"][0]["revisar"], [True, False])

        sanos = _doce_alumnos()
        planilla = gem._normalizar_planilla(_planilla_cruda(sanos))
        self.assertEqual(planilla["estudiantes"][0]["revisar"], [False, False])


# ---------------------------------------------------------------------- #
# (b) DELTA 2: mismatch len(area_trabajo) vs n_area_trabajo -> revisar_planilla
# ---------------------------------------------------------------------- #

class TestRevisarPlanillaMismatch(unittest.TestCase):

    def test_mismatch_marca_revisar_planilla(self):
        # Planilla con 3 áreas declaradas y alumnos con 2 notas: mismatch.
        estudiantes = _doce_alumnos()
        planilla = gem._normalizar_planilla(_planilla_cruda(estudiantes, n_area_trabajo=3))
        self.assertTrue(planilla["revisar_planilla"])
        # No se descarta: los 12 estudiantes siguen en la planilla.
        self.assertEqual(len(planilla["estudiantes"]), 12)

    def test_cantidad_que_coincide_no_marca(self):
        # n_area_trabajo=2 coincide con las 2 notas leídas: sin flag.
        estudiantes = _doce_alumnos()
        planilla = gem._normalizar_planilla(_planilla_cruda(estudiantes, n_area_trabajo=2))
        self.assertFalse(planilla["revisar_planilla"])

    def test_retirado_no_participa_del_mismatch(self):
        # Un retirado no tiene notas (area_trabajo None): no debe contar en la
        # comparación; sólo miran los estudiantes NO retirados.
        estudiantes = _doce_alumnos()
        estudiantes.append(_estudiante(13, retirado=True))
        planilla = gem._normalizar_planilla(_planilla_cruda(estudiantes, n_area_trabajo=2))
        self.assertFalse(planilla["revisar_planilla"])
        self.assertTrue(planilla["estudiantes"][-1]["retirado"])

    def test_planilla_con_mismatch_no_va_a_fallidas(self):
        # Pipeline real con la llamada a Gemini mockeada: la planilla marcada
        # sigue en `planillas` (no queda en `fallidas`) y el lote no se rompe.
        datos = _planilla_cruda(_doce_alumnos(), n_area_trabajo=3)
        respuestas = [json.dumps(datos, ensure_ascii=False)]
        paginas = [{"imagen": None, "indice": 1}]
        with mock.patch.object(gem, "_cliente", return_value=_FakeClient(respuestas)):
            planillas, fallidas = gem.extraer_planilla_pdf(
                paginas, api_key="clave-de-prueba"
            )
        self.assertEqual(fallidas, [])
        self.assertEqual(len(planillas), 1)
        self.assertTrue(planillas[0]["revisar_planilla"])


# ---------------------------------------------------------------------- #
# (c)/(d) DELTA 2: sanidad de la cantidad de alumnos
# ---------------------------------------------------------------------- #

class TestSanidadCantidadAlumnos(unittest.TestCase):

    def test_menos_del_minimo_marca_revision(self):
        # 4 alumnos < MIN_ESTUDIANTES_SANOS (10) -> revisar_planilla True.
        estudiantes = [_estudiante(n) for n in range(1, 5)]
        planilla = gem._normalizar_planilla(_planilla_cruda(estudiantes))
        self.assertTrue(planilla["revisar_planilla"])
        self.assertEqual(len(planilla["estudiantes"]), 4)  # no se descarta

    def test_mas_del_maximo_marca_revision(self):
        # 61 alumnos > MAX_ESTUDIANTES_SANOS (60) -> revisar_planilla True.
        estudiantes = [_estudiante(n) for n in range(1, 62)]
        planilla = gem._normalizar_planilla(_planilla_cruda(estudiantes))
        self.assertTrue(planilla["revisar_planilla"])

    def test_cantidad_normal_no_marca(self):
        # 20 alumnos, dentro de 10-60 -> sin flag.
        estudiantes = [_estudiante(n) for n in range(1, 21)]
        planilla = gem._normalizar_planilla(_planilla_cruda(estudiantes))
        self.assertFalse(planilla["revisar_planilla"])

    def test_retirados_no_cuentan_como_alumnos(self):
        # 9 activos + 3 retirados: los retirados no cuentan, 9 < 10 -> flag.
        estudiantes = [_estudiante(n) for n in range(1, 10)]
        for n in range(10, 13):
            estudiantes.append(_estudiante(n, retirado=True))
        planilla = gem._normalizar_planilla(_planilla_cruda(estudiantes))
        self.assertTrue(planilla["revisar_planilla"])
        self.assertEqual(len(planilla["estudiantes"]), 12)


# ---------------------------------------------------------------------- #
# (e) DELTA 2: sin n_area_trabajo declarado, comportamiento actual intacto
# ---------------------------------------------------------------------- #

class TestSinNAreaDeclarado(unittest.TestCase):

    def test_planilla_sin_n_area_no_se_rompe(self):
        estudiantes = _doce_alumnos()
        planilla = gem._normalizar_planilla(_planilla_cruda(estudiantes))
        # Sin n_area_trabajo no hay mismatch que revisar; con 12 alumnos el
        # conteo está sano -> sin flag, estructura intacta.
        self.assertFalse(planilla["revisar_planilla"])
        self.assertIn("encabezado", planilla)
        self.assertIn("estudiantes", planilla)
        self.assertEqual(len(planilla["estudiantes"]), 12)

    def test_n_area_en_string_se_interpreta_y_lo_no_numerico_se_ignora(self):
        estudiantes = _doce_alumnos()
        # "2" como string coincide -> sin flag.
        planilla = gem._normalizar_planilla(
            _planilla_cruda(estudiantes, n_area_trabajo="2")
        )
        self.assertFalse(planilla["revisar_planilla"])
        # "abc" no se puede interpretar: se ignora sin romper y sin flag.
        planilla = gem._normalizar_planilla(
            _planilla_cruda(estudiantes, n_area_trabajo="abc")
        )
        self.assertFalse(planilla["revisar_planilla"])

    def test_n_area_bool_se_ignora(self):
        # True es un bool (no un conteo): se descarta sin romper y sin flag.
        estudiantes = _doce_alumnos()
        planilla = gem._normalizar_planilla(
            _planilla_cruda(estudiantes, n_area_trabajo=True)
        )
        self.assertFalse(planilla["revisar_planilla"])


if __name__ == "__main__":
    unittest.main()