"""
Test del generador de Excel.

Corre el demo del módulo excel.generar_excel_notas (ejecutando su bloque
__main__) y valida que el .xlsx generado existe y tiene las tres hojas de
curso esperadas (Curso 0302, 0401 y 0501), además de fórmulas de definitiva.

Uso (desde la raíz del proyecto):
    python -m tests.test_generar_excel
"""

import os
import runpy
import sys
import unittest

# La raíz del proyecto se agrega al path para poder importar el paquete excel/.
_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

# Se importa el módulo del generador (requisito del test).
import excel.generar_excel_notas  # noqa: F401

# Ruta de salida del demo (definida en el bloque __main__ del generador).
_DEMO_XLSX = os.path.normpath(
    r"C:\Users\JUAN\AppData\Local\Temp\opencode\demo_excel\notas_educacion_fisica_periodo3.xlsx"
)
_HOJAS_ESPERADAS = ["Curso 0302", "Curso 0401", "Curso 0501"]


class TestGeneradorExcel(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        # Ejecutar el bloque __main__ del generador (el demo), que genera el Excel.
        runpy.run_module("excel.generar_excel_notas", run_name="__main__")

    def test_archivo_generado_existe(self):
        self.assertTrue(
            os.path.exists(_DEMO_XLSX),
            f"El demo no generó el archivo en {_DEMO_XLSX}",
        )

    def test_hojas_de_curso_esperadas(self):
        self.assertTrue(os.path.exists(_DEMO_XLSX), "Falta el archivo generado")
        import openpyxl

        wb = openpyxl.load_workbook(_DEMO_XLSX)
        for hoja in _HOJAS_ESPERADAS:
            self.assertIn(
                hoja, wb.sheetnames,
                f"El Excel no tiene la hoja {hoja} (hojas: {wb.sheetnames})",
            )

    def test_existencia_formula_definitiva(self):
        self.assertTrue(os.path.exists(_DEMO_XLSX), "Falta el archivo generado")
        import openpyxl

        wb = openpyxl.load_workbook(_DEMO_XLSX)
        ws = wb["Curso 0302"]
        hay_formula = any(
            isinstance(cell.value, str) and cell.value.startswith("=IFERROR(AVERAGE")
            for row in ws.iter_rows()
            for cell in row
        )
        self.assertTrue(hay_formula, "No se encontró ninguna fórmula de definitiva en Curso 0302")


if __name__ == "__main__":
    unittest.main()
