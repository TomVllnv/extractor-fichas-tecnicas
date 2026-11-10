"""Verificaciones reproducibles de compatibilidad de strings ricos de Hematología."""

import io
import re
import sys
import zipfile
import unittest
from pathlib import Path
from xml.etree import ElementTree as ET

from openpyxl import Workbook, load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import compatibilidad_excel as ce


def bytes_libro(numero=11):
    wb = Workbook()
    wb.active.title = 'EXAMEN 1'
    for i in range(2, numero + 1):
        wb.create_sheet(f'EXAMEN {i}')
    ws = wb.worksheets[10]
    ws.title = 'FIBRINOGENO'
    ws['A7'] = 'Preparación del paciente'
    ws['B7'] = CellRichText(
        TextBlock(InlineFont(b=True), 'Antes de la toma'),
        '\n', '  • Primera instrucción', '\n',
        '  • Segunda instrucción', '\n',
        TextBlock(InlineFont(b=True), 'Documentación requerida'),
        '\n', '  • Cédula de identidad',
    )
    salida = io.BytesIO()
    wb.save(salida)
    return salida.getvalue()


class TestCompatibilidadExcelHematologia(unittest.TestCase):
    def test_corrige_cadenas_de_sheet11_conserva_texto_sin_formato_enriquecido(self):
        original = bytes_libro()
        limpio = ce.normalizar_texto_enriquecido_para_excel(original)
        self.assertEqual(limpio, ce.normalizar_texto_enriquecido_para_excel(limpio))
        with zipfile.ZipFile(io.BytesIO(limpio)) as z:
            self.assertIsNone(z.testzip())
            contenido_11 = z.read('xl/worksheets/sheet11.xml')
            self.assertIn(b't="inlineStr"', contenido_11)
            self.assertNotIn('xl/sharedStrings.xml', z.namelist())
            self.assertNotIn(b'<rPr>', contenido_11)
            self.assertNotIn(b'<t/>', contenido_11)
            self.assertNotIn(b'<t />', contenido_11)
            for ruta in z.namelist():
                if ruta.endswith(('.xml', '.rels')):
                    ET.fromstring(z.read(ruta))
        wb = load_workbook(io.BytesIO(limpio), rich_text=True)
        self.assertEqual(len(wb.sheetnames), 11)
        valor = wb['FIBRINOGENO']['B7'].value
        self.assertEqual(str(valor).count('•'), 3)
        self.assertIn('\n  • Primera instrucción\n  • Segunda instrucción\n', str(valor))
        self.assertIsInstance(valor, str)

    def test_fragmento_con_namespace_heredado_de_hoja_no_provoca_parseerror(self):
        """Regresión: <si> aislado falla si una propiedad usa un prefijo heredado."""
        original = bytes_libro()
        salida = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(original)) as origen, zipfile.ZipFile(salida, 'w') as destino:
            for info in origen.infolist():
                contenido = origen.read(info.filename)
                if info.filename == 'xl/worksheets/sheet11.xml':
                    contenido = contenido.replace(
                        b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
                        b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:pro="urn:propiedades:inherited"',
                        1,
                    )
                    contenido = contenido.replace(b'<rPr>', b'<rPr pro:tipo="local">', 1)
                    ET.fromstring(contenido)  # La hoja completa es XML válido.
                    self.assertIn(b'pro:tipo=', contenido)
                destino.writestr(info, contenido)
        limpio = ce.normalizar_texto_enriquecido_para_excel(salida.getvalue())
        with zipfile.ZipFile(io.BytesIO(limpio)) as contenido:
            ET.fromstring(contenido.read('xl/worksheets/sheet11.xml'))
            self.assertNotIn('xl/sharedStrings.xml', contenido.namelist())
        excel = load_workbook(io.BytesIO(limpio), rich_text=True)
        self.assertIn('• Primera instrucción', str(excel['FIBRINOGENO']['B7'].value))

    def test_xml_invalido_presenta_error_controlado_en_lugar_de_parseerror(self):
        original = bytes_libro()
        salida = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(original)) as origen, zipfile.ZipFile(salida, 'w') as destino:
            for info in origen.infolist():
                contenido = origen.read(info.filename)
                if info.filename == 'xl/worksheets/sheet11.xml':
                    contenido = contenido.replace(b'<rPr>', b'<rPr><malformado', 1)
                destino.writestr(info, contenido)
        with self.assertRaisesRegex(ValueError, 'XML inválido'):
            ce.normalizar_texto_enriquecido_para_excel(salida.getvalue())

    def test_descarta_fragmentos_vacios_sin_quitar_vinetas(self):
        libro = Workbook()
        libro.active.title = 'FIBRINOGENO'
        libro.active['A1'] = 'Preparación del paciente'
        libro.active['B1'] = CellRichText(
            TextBlock(InlineFont(b=True), 'Antes de la toma'),
            '\n', '', '\n', '  • Tomar muestra según indicación',
            '\n', TextBlock(InlineFont(b=True), 'Documentación requerida'),
            '\n', '  • Cédula de identidad',
        )
        datos = io.BytesIO(); libro.save(datos)
        limpio = ce.normalizar_texto_enriquecido_para_excel(datos.getvalue())
        with zipfile.ZipFile(io.BytesIO(limpio)) as archivo:
            self.assertNotIn('xl/sharedStrings.xml', archivo.namelist())
            raiz = ET.fromstring(archivo.read('xl/worksheets/sheet1.xml'))
            ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
            for run in raiz.iter(ns + 'r'):
                nodo = run.find(ns + 't')
                self.assertTrue(nodo is None or (nodo.text or '') != '')
        cargado = load_workbook(io.BytesIO(limpio), rich_text=True)
        valor = cargado.active['B1'].value
        self.assertIn('\n\n  • Tomar muestra', str(valor))
        self.assertEqual(str(valor).count('•'), 2)
        self.assertIsInstance(valor, str)

    def test_libro_sin_rich_text_no_cambia(self):
        wb = Workbook(); wb.active['B7'] = 'Datos normales'
        bf = io.BytesIO(); wb.save(bf)
        original = bf.getvalue()
        self.assertEqual(original, ce.normalizar_texto_enriquecido_para_excel(original))

    def test_macro_opcional_tiene_instruccion_solicitada(self):
        macro = (ROOT / 'AutoajustarColumnas.bas').read_text(encoding='utf-8')
        self.assertIn('Cells.EntireColumn.AutoFit', macro)


if __name__ == '__main__':
    unittest.main()
