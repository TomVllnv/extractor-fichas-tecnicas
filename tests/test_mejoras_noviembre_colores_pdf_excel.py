"""Nueva selección cromática, visor con flechas y control XLSX de hojas extensas."""
import ast
import io
import sys
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from openpyxl import Workbook, load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import coincidencias_ui as cui
import compatibilidad_excel as ce

APP = (ROOT / 'app.py').read_text(encoding='utf-8')
BRIDGE = (ROOT / 'interfaz_fluida.py').read_text(encoding='utf-8')


def libro_de_11_hojas(valor):
    wb = Workbook()
    for i in range(2, 12):
        wb.create_sheet(f'EXAMEN {i}')
    wb.worksheets[10].title = 'FIBRINOGENO'
    wb.worksheets[10]['B7'] = valor
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


class TestExcelSinPropiedadesIncompatibles(unittest.TestCase):
    def test_no_descarga_cadena_rich_demasiado_larga_en_hoja11(self):
        valor = CellRichText(TextBlock(InlineFont(b=True), 'Aviso:'), 'A' * 32767)
        with self.assertRaisesRegex(ValueError, 'sheet11.xml.*32,767'):
            ce.normalizar_texto_enriquecido_para_excel(libro_de_11_hojas(valor))

    def test_longitud_cuenta_caracteres_utf16(self):
        valor = CellRichText(TextBlock(InlineFont(i=True), 'A'), '🙂' * 16384)
        with self.assertRaisesRegex(ValueError, 'sheet11.xml'):
            ce.normalizar_texto_enriquecido_para_excel(libro_de_11_hojas(valor))

    def test_rpr_vacio_se_corrige_y_no_pierde_texto(self):
        valor = CellRichText(TextBlock(InlineFont(), 'Primero'), '\n',
                             TextBlock(InlineFont(b=True), 'Segundo'))
        libro = libro_de_11_hojas(valor)
        limpio = ce.normalizar_texto_enriquecido_para_excel(libro)
        with zipfile.ZipFile(io.BytesIO(limpio)) as z:
            raiz = ET.fromstring(z.read('xl/worksheets/sheet11.xml'))
            ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
            for run in raiz.iter(ns + 'r'):
                propiedades = run.find(ns + 'rPr')
                self.assertFalse(propiedades is not None and len(propiedades) == 0)
        celda = load_workbook(io.BytesIO(limpio), rich_text=True).worksheets[10]['B7']
        self.assertEqual(str(celda.value), 'Primero\nSegundo')
        self.assertIsInstance(celda.value, str)

    def test_reconoce_celda_que_cabe_exactamente(self):
        valor = CellRichText(TextBlock(InlineFont(i=True), 'A'), 'B' * 32766)
        ce.normalizar_texto_enriquecido_para_excel(libro_de_11_hojas(valor))


class TestSeleccionVisualYVisor(unittest.TestCase):
    def test_gradiente_de_colores_desde_rojo_a_verde(self):
        valores = [0.55, 0.65, 0.75, 0.85, 0.95, 1.0]
        matices = [cui.matiz_similitud(x) for x in valores]
        self.assertEqual(matices, sorted(matices))
        self.assertEqual(matices[0], 0)
        self.assertEqual(matices[-1], 132)
        self.assertEqual(cui.porcentaje_similitud(1.0), 100)
        self.assertEqual(cui.porcentaje_similitud(0.789), 79)

    def test_nombre_y_seleccionar_tienen_mismo_color_y_hover_blanco(self):
        nombre = cui.claves_coincidencia('crs', 'EXAMEN 1')
        reglas = cui.estilos_coincidencia(nombre, 0.75)
        self.assertIn('st-key-' + nombre, reglas)
        self.assertNotIn('coincidencia_elegir_', reglas)
        self.assertIn('background: #ffffff', reglas)
        self.assertIn('button:hover', reglas)
        self.assertIn('on_click=accion', APP)
        self.assertNotIn('st.button("Seleccionar", key=clave_boton, on_click=accion,', APP)
        self.assertIn('"crs", nombre_sugerido', APP)
        self.assertIn('"acreditacion", identificador', APP)

    def test_sugerencia_unica_tambien_se_puede_pulsar(self):
        self.assertNotIn('if len(sugeridas_crs) > 1 or seleccion_manual_cartera:', APP)
        self.assertNotIn('if len(sugeridas) > 1:', APP)
        self.assertIn('for nombre_sugerido, laboratorio_sugerido, puntaje in sugeridas_crs:', APP)

    def test_flechas_sincronizadas_con_slider_y_documento(self):
        self.assertIn('pag_seleccionada = st.slider(', APP)
        self.assertIn('key=clave_pagina', APP)
        self.assertIn('on_click=cambiar_pagina_preview', APP)
        self.assertIn('args=(clave_pagina, -1, num_paginas)', APP)
        self.assertIn('args=(clave_pagina, 1, num_paginas)', APP)
        self.assertIn('disabled=pag_seleccionada <= 1', APP)
        self.assertIn('disabled=pag_seleccionada >= num_paginas', APP)
        fn = next(n for n in ast.parse(APP).body
                  if isinstance(n, ast.FunctionDef) and n.name == 'cambiar_pagina_preview')
        class StreamlitFake:
            session_state = {'pagina': 2}
        scope = {'st': StreamlitFake}
        exec(compile(ast.Module(body=[fn], type_ignores=[]), 'app.py', 'exec'), scope)
        cb = scope['cambiar_pagina_preview']
        cb('pagina', 1, 3)
        self.assertEqual(StreamlitFake.session_state['pagina'], 3)
        cb('pagina', 1, 3)
        self.assertEqual(StreamlitFake.session_state['pagina'], 3)
        cb('pagina', -5, 3)
        self.assertEqual(StreamlitFake.session_state['pagina'], 1)

    def test_referencias_no_se_cierran_ni_se_interceptan(self):
        self.assertIn('key="contenedor_referencias"', APP)
        self.assertNotIn('detalles.open = false', BRIDGE)
        self.assertNotIn('stopImmediatePropagation', BRIDGE)
        self.assertIn("registrar('referencias')", BRIDGE)


if __name__ == '__main__':
    unittest.main()
