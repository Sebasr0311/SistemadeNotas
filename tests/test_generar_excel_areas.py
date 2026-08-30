"""
Test del ancho dinámico de columnas de área de trabajo en el generador de Excel.

Cubre el camino de salida del pipeline (spec v2): la extracción puede devolver
de 1 a 16 notas por alumno y el generador debe reflejar ESE ancho real en el
Excel, no un fijo de 2 columnas. También cubre la inferencia cuando no se
declara n_area_trabajo pero sí hay notas observadas, y la regla de fallback a
la planilla tradicional de 2 notas cuando no hay nada que medir.

Uso (desde la raíz del proyecto):
    python -m tests.test_generar_excel_areas
"""

import os
import sys
import tempfile
import unittest

# La raíz del proyecto se agrega al path para poder importar el paquete excel/.
_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

import openpyxl  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402

import excel.generar_excel_notas as generador  # noqa: E402


def _planilla(area_por_alumno, n_area_trabajo=None, periodo=3):
    """Construye una planilla con `area_por_alumno` (lista de listas o listas
    de notas por alumno). Quien dé None en `area_por_alumno` es un retirado."""
    encabezado = {
        "institucion": "INSTITUCION DEMO",
        "sede": "SEDE",
        "año_lectivo": "2026",
        "jornada": "MAÑANA",
        "grupo": "0302",
        "asignatura": "MATEMATICAS",
        "docente": "DOCENTE DEMO",
        "periodo": periodo,
    }
    if n_area_trabajo is not None:
        encabezado["n_area_trabajo"] = n_area_trabajo

    estudiantes = []
    for i, notas in enumerate(area_por_alumno, start=1):
        retirado = notas is None
        estudiantes.append({
            "no": i,
            "nombre": f"ALUMNO {i}",
            "ev_anteriores": [45, 45] if periodo > 1 else [],
            "area_trabajo": notas if not retirado else None,
            "retirado": retirado,
            "revisar": [False] * (len(notas) if not retirado else 0),
        })
    return {"encabezado": encabezado, "estudiantes": estudiantes}


def _hoja_cargada(planilla):
    """Genera el Excel de la planilla y devuelve la hoja activa ya cargada."""
    directorio = tempfile.mkdtemp(prefix="notas_areas_")
    ruta = os.path.join(directorio, "salida.xlsx")
    generador.generar_excel_planilla(planilla, ruta)
    wb = openpyxl.load_workbook(ruta)
    return wb[wb.sheetnames[0]]


class TestAnchoDinamicoAreas(unittest.TestCase):
    """El ancho de las columnas de área sigue la cantidad real de notas."""

    def test_planilla_de_5_notas_escribe_5_columnas_y_formula_EI(self):
        # Periodo 3 => 2 columnas de definitiva previa (C y D), área arranca en E.
        planilla = _planilla(
            [
                [40, 50, 60, 70, 80],
                [41, 51, 61, 71, 81],
                [42, 52, 62, 72, 82],
            ],
            n_area_trabajo=5,
        )
        ws = _hoja_cargada(planilla)

        # Fila 10 = header (8 filas informativas + 2). Verifico los títulos.
        headers = [ws.cell(row=10, column=j).value for j in range(1, 12)]
        for k in range(1, 6):
            self.assertIn(f"Área Trabajo {k}", headers)
        self.assertIn("Definitiva Periodo 3", headers)

        # Alumno 1 (fila 11): nota 5 en la columna I (9), nota 1 en E (5),
        # y la definitiva en J (10) con AVERAGE(E:I).
        self.assertEqual(ws.cell(row=11, column=5).value, 40)
        self.assertEqual(ws.cell(row=11, column=9).value, 80)
        # La fórmula usa el rango completo de las 5 celdas de área (E:I).
        self.assertEqual(
            ws.cell(row=11, column=10).value,
            '=IFERROR(AVERAGE(E11:I11),"")',
        )

    def test_alumno_con_menos_notas_deja_celda_vacia(self):
        # Un alumno con 4 notas en una planilla de 5: la 5ª celda (columna I)
        # queda None (vacía), sin excepción por índice fuera de rango.
        planilla = _planilla(
            [
                [40, 50, 60, 70, 80],
                [41, 51, 61, 71],  # solo 4 notas
            ],
            n_area_trabajo=5,
        )
        ws = _hoja_cargada(planilla)
        self.assertEqual(ws.cell(row=12, column=5).value, 41)
        self.assertEqual(ws.cell(row=12, column=8).value, 71)  # 4ª nota (H)
        self.assertIsNone(ws.cell(row=12, column=9).value)  # 5ª celda vacía (I)

    def test_una_sola_nota_genera_una_columna_y_formula_EE(self):
        # Periodo 3 => área arranca en E; una sola nota => col_def = F, rango E:E.
        planilla = _planilla(
            [[55]],
            n_area_trabajo=1,
        )
        ws = _hoja_cargada(planilla)

        headers = [ws.cell(row=10, column=j).value for j in range(1, 8)]
        self.assertIn("Área Trabajo 1", headers)
        self.assertNotIn("Área Trabajo 2", headers)

        self.assertEqual(ws.cell(row=11, column=5).value, 55)
        self.assertEqual(
            ws.cell(row=11, column=6).value,
            '=IFERROR(AVERAGE(E11:E11),"")',
        )

    def test_inferencia_sin_declaracion_usa_lo_observado(self):
        # Sin n_area_trabajo en el encabezado pero con 3 notas por alumno:
        # el ancho se infiere de lo observado (3 columnas).
        planilla = _planilla(
            [
                [40, 50, 60],
                [41, 51, 61],
            ],
            n_area_trabajo=None,
        )
        ws = _hoja_cargada(planilla)

        headers = [ws.cell(row=10, column=j).value for j in range(1, 9)]
        self.assertIn("Área Trabajo 1", headers)
        self.assertIn("Área Trabajo 3", headers)
        self.assertNotIn("Área Trabajo 4", headers)

        # Área arranca en E (3), con 3 notas => col_def = 5 + 3 = 8 (H);
        # rango E:G en la celda de definitiva.
        self.assertEqual(ws.cell(row=11, column=8).value,
                         '=IFERROR(AVERAGE(E11:G11),"")')


if __name__ == "__main__":
    unittest.main()
