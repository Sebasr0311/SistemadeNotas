"""
Tests de la entrada mixta PDF + imágenes.

Cubre:
- cargar_imagen: UN dict con la forma de página (indice=1, _render perezoso
  que devuelve PIL RGB, _doc=None, es_imagen=True) y cerrar_paginas tolerante.
- contar_planillas: imagen -> 1, PDF -> páginas.
- cargar_archivo: despacha por extensión y rechaza otros archivos.
- extraer_planilla_pdf con `origen`: adjunta `planilla["fuente"]` con el
  `indice` correcto a cada planilla (1:1 página->planilla); sin `origen` no
  agrega la clave (compat total).
- agrupar_por_forma: agrupa por cantidad de columnas de notas con orden
  estable de primera aparición.
- render_imagen_planilla: devuelve PIL RGB según la fuente (PDF o imagen) y
  None si la planilla no tiene fuente.

Uso (desde la raíz del proyecto):
    python -m unittest discover -s tests -v
"""

import json
import os
import shutil
import sys
import tempfile
import unittest
from unittest import mock

_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

import pymupdf as fitz  # noqa: E402
from PIL import Image  # noqa: E402

from excel.agrupacion import agrupar_por_forma  # noqa: E402
from pdf_processing import pdf_loader  # noqa: E402
from vision import gemini_extractor as gem  # noqa: E402


# ---------------------------------------------------------------------- #
# Fixtures
# ---------------------------------------------------------------------- #

def _crear_png(ruta, ancho=400, alto=300, color=(200, 200, 240)):
    """Crea un PNG real chico con PIL."""
    Image.new("RGB", (ancho, alto), color).save(ruta, format="PNG")


def _crear_pdf(ruta, paginas=2):
    doc = fitz.open()
    for _ in range(paginas):
        doc.new_page()
    doc.save(ruta)
    doc.close()


def _json_planilla(grupo: str, estudiantes: list) -> str:
    planilla = {
        "encabezado": {
            "institucion": "INSTITUCION PRUEBA",
            "sede": "SEDE PRINCIPAL",
            "año_lectivo": "2026",
            "jornada": "MAÑANA",
            "grupo": grupo,
            "asignatura": "MATEMATICAS",
            "docente": "DOCENTE PRUEBA",
            "periodo": 3,
        },
        "estudiantes": estudiantes,
    }
    return json.dumps(planilla, ensure_ascii=False)


class _FakeResponse:
    def __init__(self, texto):
        self.text = texto


class _FakeModels:
    def __init__(self, respuestas):
        self._respuestas = respuestas
        self._contador = 0

    def generate_content(self, **kwargs):
        if self._contador >= len(self._respuestas):
            raise AssertionError("generate_content se llamó más veces que páginas")
        texto = self._respuestas[self._contador]
        self._contador += 1
        return _FakeResponse(texto)


class _FakeClient:
    def __init__(self, respuestas):
        self.models = _FakeModels(respuestas)


def _estudiante(no, nombre, n_notas=2):
    return {
        "no": no, "nombre": nombre, "ev_anteriores": [45, 45],
        "area_trabajo": [40 + i for i in range(n_notas)],
        "retirado": False, "revisar": [False] * n_notas,
    }


# ---------------------------------------------------------------------- #
# cargar_imagen / contar_planillas / cargar_archivo
# ---------------------------------------------------------------------- #

class TestCargarImagen(unittest.TestCase):

    def setUp(self):
        self.dir_tmp = tempfile.mkdtemp(prefix="notas_imagenes_")

    def tearDown(self):
        shutil.rmtree(self.dir_tmp, ignore_errors=True)

    def _ruta_png(self):
        ruta = os.path.join(self.dir_tmp, "planilla.png")
        _crear_png(ruta)
        return ruta

    def test_cargar_imagen_forma_de_pagina(self):
        paginas = pdf_loader.cargar_imagen(self._ruta_png())
        self.assertEqual(len(paginas), 1)
        pag = paginas[0]
        self.assertEqual(pag["indice"], 1)
        self.assertGreater(pag["ancho_px"], 0)
        self.assertGreater(pag["alto_px"], 0)
        self.assertIsNone(pag["_doc"], "una imagen no tiene PDF interno que cerrar")
        self.assertTrue(pag["es_imagen"])
        # _render perezoso: devuelve una imagen PIL RGB recién al llamarlo.
        img = pag["_render"]()
        self.assertIsInstance(img, Image.Image)
        self.assertEqual(img.mode, "RGB")
        self.assertEqual(img.size, (400, 300))
        img.close()

    def test_cerrar_paginas_tolerante_con_imagenes(self):
        paginas = pdf_loader.cargar_imagen(self._ruta_png())
        # No debe explotar (no hay _doc) y queda idem para la segunda llamada.
        pdf_loader.cerrar_paginas(paginas)
        pdf_loader.cerrar_paginas(paginas)
        self.assertIsNone(paginas[0]["_doc"])

    def test_contar_planillas_imagen_es_una(self):
        self.assertEqual(pdf_loader.contar_planillas(self._ruta_png()), 1)

    def test_contar_planillas_pdf_es_sus_paginas(self):
        ruta = os.path.join(self.dir_tmp, "planillas.pdf")
        _crear_pdf(ruta, paginas=3)
        self.assertEqual(pdf_loader.contar_planillas(ruta), 3)

    def test_cargar_archivo_despacha_por_extension(self):
        ruta_png = self._ruta_png()
        ruta_pdf = os.path.join(self.dir_tmp, "planillas.pdf")
        _crear_pdf(ruta_pdf, paginas=2)

        paginas_img = pdf_loader.cargar_archivo(ruta_png)
        self.assertEqual(len(paginas_img), 1)
        self.assertTrue(paginas_img[0]["es_imagen"])
        pdf_loader.cerrar_paginas(paginas_img)

        paginas_pdf = pdf_loader.cargar_archivo(ruta_pdf)
        self.assertEqual(len(paginas_pdf), 2)
        self.assertNotIn("es_imagen", paginas_pdf[0])
        try:
            self.assertTrue(callable(paginas_pdf[0]["_render"]))
        finally:
            pdf_loader.cerrar_paginas(paginas_pdf)

    def test_cargar_archivo_rechaza_otros(self):
        ruta = os.path.join(self.dir_tmp, "texto.txt")
        with open(ruta, "w", encoding="utf-8") as f:
            f.write("no soy una planilla")
        with self.assertRaises(pdf_loader.PdfError):
            pdf_loader.cargar_archivo(ruta)
        with self.assertRaises(pdf_loader.PdfError):
            pdf_loader.contar_planillas(ruta)


# ---------------------------------------------------------------------- #
# extraer_planilla_pdf con origen (fuente por planilla)
# ---------------------------------------------------------------------- #

class TestExtraerConOrigen(unittest.TestCase):

    def setUp(self):
        self.dir_tmp = tempfile.mkdtemp(prefix="notas_origen_")
        self.ruta_png = os.path.join(self.dir_tmp, "foto_planilla.png")
        _crear_png(self.ruta_png)

    def tearDown(self):
        shutil.rmtree(self.dir_tmp, ignore_errors=True)

    def _paginas_fixture(self):
        """Páginas estilo legacy (imagen materializada) con indices 1 y 3 para
        verificar que el indice se hereda tal cual (1:1 página->planilla)."""
        with Image.open(self.ruta_png) as img:
            imagen = img.convert("RGB")
        return [
            {"indice": 1, "imagen": imagen, "ancho_px": 400, "alto_px": 300},
            {"indice": 3, "imagen": imagen, "ancho_px": 400, "alto_px": 300},
        ]

    def test_origen_adjunta_fuente_con_indice_correcto(self):
        respuestas = [
            _json_planilla("0201", [_estudiante(1, "GARCIA LUZ")]),
            _json_planilla("0302", [_estudiante(2, "PEREZ JUAN")]),
        ]
        paginas = self._paginas_fixture()
        origen = {"tipo": "imagen", "ruta": self.ruta_png}
        with mock.patch.object(gem, "_cliente", return_value=_FakeClient(respuestas)):
            planillas, fallidas = gem.extraer_planilla_pdf(
                paginas, api_key="clave-de-prueba", origen=origen
            )

        self.assertEqual(fallidas, [])
        self.assertEqual(len(planillas), 2)
        # El origen se fusiona y el indice de la página se hereda 1:1.
        self.assertEqual(planillas[0]["fuente"]["tipo"], "imagen")
        self.assertEqual(planillas[0]["fuente"]["ruta"], self.ruta_png)
        self.assertEqual(planillas[0]["fuente"]["indice"], 1)
        self.assertEqual(planillas[1]["fuente"]["indice"], 3)

    def test_sin_origen_no_agrega_la_clave_fuente(self):
        respuestas = [
            _json_planilla("0201", [_estudiante(1, "GARCIA LUZ")]),
        ]
        paginas = self._paginas_fixture()[:1]
        with mock.patch.object(gem, "_cliente", return_value=_FakeClient(respuestas)):
            planillas, fallidas = gem.extraer_planilla_pdf(
                paginas, api_key="clave-de-prueba"
            )

        self.assertEqual(fallidas, [])
        self.assertEqual(len(planillas), 1)
        self.assertNotIn("fuente", planillas[0])


# ---------------------------------------------------------------------- #
# agrupar_por_forma
# ---------------------------------------------------------------------- #

class TestAgruparPorForma(unittest.TestCase):

    def _planilla(self, grupo, n_notas):
        estudiantes = [_estudiante(1, "ALUMNO UNO", n_notas=n_notas)]
        return {"encabezado": {"grupo": grupo}, "estudiantes": estudiantes}

    def test_agrupa_por_cantidad_de_columnas_con_orden_estable(self):
        # 2 planillas de 5 columnas + 1 de 7 -> 2 formas, orden de aparición.
        p1 = self._planilla("0201", 5)
        p2 = self._planilla("0302", 7)
        p3 = self._planilla("0202", 5)
        por_forma, orden = agrupar_por_forma([p1, p2, p3])

        self.assertEqual(orden, [("n_areas", 5), ("n_areas", 7)])
        self.assertEqual(len(por_forma[("n_areas", 5)]), 2)
        self.assertEqual(len(por_forma[("n_areas", 7)]), 1)
        self.assertEqual(por_forma[("n_areas", 5)][0], p1)
        self.assertEqual(por_forma[("n_areas", 5)][1], p3)

    def test_planillas_vacias_no_rompen(self):
        por_forma, orden = agrupar_por_forma([])
        self.assertEqual(orden, [])
        self.assertEqual(por_forma, {})

    def test_no_muta_las_planillas(self):
        p = self._planilla("0201", 4)
        agrupar_por_forma([p])
        self.assertEqual(p, self._planilla("0201", 4))


# ---------------------------------------------------------------------- #
# render_imagen_planilla
# ---------------------------------------------------------------------- #

class TestRenderImagenPlanilla(unittest.TestCase):

    def setUp(self):
        self.dir_tmp = tempfile.mkdtemp(prefix="notas_render_")

    def tearDown(self):
        shutil.rmtree(self.dir_tmp, ignore_errors=True)

    def test_fuente_imagen_devuelve_pil_rgb(self):
        ruta = os.path.join(self.dir_tmp, "planilla.png")
        _crear_png(ruta)
        planilla = {"encabezado": {}, "fuente": {"tipo": "imagen", "ruta": ruta}}
        img = pdf_loader.render_imagen_planilla(planilla)
        self.assertIsInstance(img, Image.Image)
        self.assertEqual(img.mode, "RGB")
        img.close()

    def test_fuente_pdf_devuelve_pil_rgb_de_esa_pagina(self):
        ruta = os.path.join(self.dir_tmp, "planillas.pdf")
        _crear_pdf(ruta, paginas=2)
        planilla = {"encabezado": {}, "fuente": {"tipo": "pdf", "ruta": ruta, "indice": 2}}
        img = pdf_loader.render_imagen_planilla(planilla, dpi=150)
        self.assertIsInstance(img, Image.Image)
        self.assertEqual(img.mode, "RGB")
        img.close()

    def test_sin_fuente_devuelve_none(self):
        self.assertIsNone(pdf_loader.render_imagen_planilla({"encabezado": {}}))
        self.assertIsNone(pdf_loader.render_imagen_planilla(None))

    def test_fuente_invalida_devuelve_none(self):
        ruta = os.path.join(self.dir_tmp, "no_existe.png")
        planilla = {"encabezado": {}, "fuente": {"tipo": "imagen", "ruta": ruta}}
        self.assertIsNone(pdf_loader.render_imagen_planilla(planilla))


if __name__ == "__main__":
    unittest.main()