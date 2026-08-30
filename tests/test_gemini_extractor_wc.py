"""
Test W-C: el error de parseo de una planilla debe quedar registrado en el log
(app_config.escribir_log) antes de elevarse como _PlanillaParseError.

Se mockea el cliente de Gemini para no tocar la red: la respuesta del modelo
es texto que no es JSON, por lo que _extraer_json falla y se dispara el path
de parseo que debe escribir el log.

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

import vision.gemini_extractor as gem  # noqa: E402


class _FakeResponse:
    text = "esto no es un JSON valido {{"


class _FakeModels:
    def generate_content(self, **kwargs):
        return _FakeResponse()


class _FakeClient:
    models = _FakeModels()


class TestParseErrorLogueado(unittest.TestCase):

    def test_error_de_parseo_escribe_el_log(self):
        log_lineas = []

        def _fake_log(msg):
            log_lineas.append(msg)

        with mock.patch.object(gem.app_config, "escribir_log", side_effect=_fake_log), \
             mock.patch.object(gem, "_cliente", return_value=_FakeClient()):
            with self.assertRaises(gem._PlanillaParseError):
                gem.extraer_planilla_pagina(
                    imagen="IGNORADA", api_key="clave-de-prueba",
                    pagina_label="planilla 1 de 2",
                )

        self.assertTrue(
            any("Error de parseo en planilla 1 de 2" in ln for ln in log_lineas),
            f"El log no quedó registrado. Líneas de log: {log_lineas}",
        )


if __name__ == "__main__":
    unittest.main()
