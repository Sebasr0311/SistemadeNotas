"""
Tests de contar_paginas y del límite de planillas por PDF (S16).

PDFs reales creados con PyMuPDF en un directorio temporal: se verifica que
contar_paginas devuelva el conteo correcto sin rasterizar, que el límite de
MAX_PLANILLAS_POR_PDF (10) se aplique al conteo, y que una ruta inexistente
lance PdfError con mensaje amigable.

Uso (desde la raíz del proyecto):
    python -m unittest discover -s tests -v
"""

import os
import shutil
import sys
import tempfile
import unittest

_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

import pymupdf as fitz  # noqa: E402

from pdf_processing import pdf_loader  # noqa: E402


def _crear_pdf_prueba(ruta, paginas=2):
    """Crea un PDF real chico con PyMuPDF (una página en blanco por planilla)."""
    doc = fitz.open()
    for _ in range(paginas):
        doc.new_page()
    doc.save(ruta)
    doc.close()


class TestContarPaginas(unittest.TestCase):

    def setUp(self):
        self.dir_tmp = tempfile.mkdtemp(prefix="notas_contar_")

    def tearDown(self):
        shutil.rmtree(self.dir_tmp, ignore_errors=True)

    def _ruta_pdf(self, paginas):
        ruta = os.path.join(self.dir_tmp, f"planillas_{paginas}.pdf")
        _crear_pdf_prueba(ruta, paginas=paginas)
        return ruta

    def test_conteo_correcto(self):
        # 3 páginas -> 3 (sin rasterizar nada).
        ruta = self._ruta_pdf(3)
        self.assertEqual(pdf_loader.contar_paginas(ruta), 3)

    def test_diez_paginas_no_supera_el_limite(self):
        ruta = self._ruta_pdf(10)
        n = pdf_loader.contar_paginas(ruta)
        self.assertEqual(n, pdf_loader.MAX_PLANILLAS_POR_PDF)
        self.assertFalse(n > pdf_loader.MAX_PLANILLAS_POR_PDF)

    def test_once_paginas_supera_el_limite(self):
        ruta = self._ruta_pdf(11)
        n = pdf_loader.contar_paginas(ruta)
        self.assertTrue(n > pdf_loader.MAX_PLANILLAS_POR_PDF)

    def test_ruta_inexistente_lanza_pdf_error(self):
        ruta = os.path.join(self.dir_tmp, "no_existe.pdf")
        with self.assertRaises(pdf_loader.PdfError):
            pdf_loader.contar_paginas(ruta)

    def test_no_es_un_pdf_lanza_pdf_error(self):
        ruta = os.path.join(self.dir_tmp, "texto.txt")
        with open(ruta, "w", encoding="utf-8") as f:
            f.write("esto no es un pdf")
        with self.assertRaises(pdf_loader.PdfError):
            pdf_loader.contar_paginas(ruta)


if __name__ == "__main__":
    unittest.main()