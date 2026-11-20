"""Evita decenas de fragmentos vacíos o de salto de línea en celdas ricas."""
import io
import sys
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from openpyxl import load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import compatibilidad_excel as ce


class TestCompactacion(unittest.TestCase):
    def test_no_aplana_y_reduce_runs_sin_perder_estilos(self):
        partes = CellRichText(
            TextBlock(InlineFont(b=True), 'Antes de la toma'),
            '\n', '\n', '  • Ayuno', '\n', '  • Agua', '\n',
            TextBlock(InlineFont(b=True), 'Documentación requerida'),
            '\n', '  • Orden',
        )
        texto = str(partes)
        compacto = ce.compactar_fragmentos_enriquecidos(partes)
        self.assertEqual(str(compacto), texto)
        self.assertLess(len(compacto), len(partes))
        self.assertEqual(sum(isinstance(p, TextBlock) and p.font.b for p in compacto), 2)
        self.assertEqual([x for x in compacto if isinstance(x, str)],
                         ['\n\n  • Ayuno\n  • Agua\n', '\n  • Orden'])

    def test_compactacion_recupera_excel_anterior(self):
        anterior = Path('/mnt/data/Ficha_Hematologia_reparada_para_comprobar.xlsx')
        if not anterior.exists():
            self.skipTest('No hay copia real del libro de diagnóstico')
        original = anterior.read_bytes()
        limpio = ce.normalizar_texto_enriquecido_para_excel(original)
        with zipfile.ZipFile(io.BytesIO(original)) as src, zipfile.ZipFile(io.BytesIO(limpio)) as out:
            a = ET.fromstring(src.read('xl/worksheets/sheet11.xml'))
            b = ET.fromstring(out.read('xl/worksheets/sheet11.xml'))
            ns = {'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            for coordenada in ('B7','B21'):
                r_a = a.find(f'.//s:c[@r="{coordenada}"]', ns)
                r_b = b.find(f'.//s:c[@r="{coordenada}"]', ns)
                self.assertLessEqual(len(r_b.findall('.//s:r', ns)), len(r_a.findall('.//s:r', ns)))
            self.assertNotIn('xl/sharedStrings.xml', out.namelist())
            self.assertEqual(out.testzip(), None)
        w_original = load_workbook(io.BytesIO(original), rich_text=True)
        w_limpio = load_workbook(io.BytesIO(limpio), rich_text=True)
        self.assertEqual(w_original.sheetnames, w_limpio.sheetnames)
        self.assertEqual(str(w_original.worksheets[10]['B7'].value),
                         str(w_limpio.worksheets[10]['B7'].value))
        self.assertEqual(str(w_original.worksheets[10]['B21'].value),
                         str(w_limpio.worksheets[10]['B21'].value))


if __name__ == '__main__':
    unittest.main()
