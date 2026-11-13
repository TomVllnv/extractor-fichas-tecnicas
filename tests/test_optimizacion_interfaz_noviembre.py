"""Regresiones de rendimiento visual sin depender de Streamlit o SQLite."""
from pathlib import Path
import ast
import hashlib
import types
import unittest

ROOT = Path(__file__).resolve().parents[1]
APP = (ROOT / 'app.py').read_text(encoding='utf-8')
BRIDGE = (ROOT / 'interfaz_fluida.py').read_text(encoding='utf-8')


class TestOptimizacionInterfaz(unittest.TestCase):
    def test_elimina_solo_expander_secundario_de_avisos(self):
        self.assertIn('with st.expander("Referencias bibliográficas", expanded=True)', APP)
        self.assertNotIn('with st.expander("Datos bibliográficos que conviene verificar")', APP)
        self.assertIn('apa.advertencias_fuentes(seleccionadas)', APP)
        self.assertIn('for aviso in avisos:', APP)
        self.assertIn('st.caption("• " + aviso)', APP)

    def test_referencias_no_cierra_panel_tras_confirmar(self):
        self.assertNotIn('::details-content {', APP)
        self.assertNotIn('setInterval', BRIDGE)
        self.assertIn('UI.observador.disconnect()', BRIDGE)
        self.assertIn("registrar('referencias')", BRIDGE)
        self.assertNotIn('detalles.open = false', BRIDGE)
        self.assertNotIn('animandoReferencias', BRIDGE)
        self.assertNotIn('stopImmediatePropagation', BRIDGE)
        self.assertIn('key="contenedor_referencias"', APP)
        self.assertNotIn('st.session_state["generacion_panel_referencias"] = (', APP)

    def test_bienvenida_no_se_muestra_entre_vistas(self):
        self.assertIn('ocultarBienvenidaDuranteCambio()', BRIDGE)
        self.assertIn('st-key-abrir_area_', BRIDGE)
        self.assertIn("principal.style.visibility = 'hidden'", BRIDGE)
        self.assertIn("principal.style.removeProperty('visibility')", BRIDGE)
        self.assertIn("!principal.textContent.includes('¡Bienvenido/a", BRIDGE)
        self.assertIn('seleccionar_area', APP)

    def test_descargas_sidebar_no_rehacen_la_pagina(self):
        sidebar = APP.split('# --- SIDEBAR ---', 1)[1].split('# --- VISTA PRINCIPAL ---', 1)[0]
        for key in ('dl_area_btn_', 'descargar_acreditacion_original', 'dl_cartera_btn',
                    'descargar_normas_original'):
            self.assertIn(key, sidebar)
        self.assertGreaterEqual(sidebar.count('on_click="ignore"'), 4)

    def test_bibliografia_pdf_en_cache_sin_alterar_pdf(self):
        target = next(n for n in ast.parse(APP).body
                      if isinstance(n, ast.FunctionDef) and n.name == '_fuente_pdf_apa_cacheada')
        code = compile(ast.Module(body=[target], type_ignores=[]), str(ROOT/'app.py'), 'exec')
        llamadas = []
        apa = types.SimpleNamespace(fuente_pdf=lambda nombre, contenido, examen:
                                    llamadas.append((nombre, contenido, examen)) or len(llamadas))
        st = types.SimpleNamespace(session_state={})
        scope = {'st': st, 'apa': apa, 'hashlib': hashlib}
        exec(code, scope)
        lectura = scope['_fuente_pdf_apa_cacheada']

        class PdfFalso:
            def __init__(self, contenido):
                self.name = 'documento.pdf'
                self.contenido = contenido
            def getvalue(self):
                return self.contenido

        archivo = PdfFalso(b'%PDF-1.4 contenido original')
        self.assertEqual(lectura(archivo, 'Examen 1'), 1)
        self.assertEqual(lectura(archivo, 'Examen 1'), 1)
        self.assertEqual(len(llamadas), 1)
        archivo.contenido = b'%PDF-1.4 contenido actualizado'
        self.assertEqual(lectura(archivo, 'Examen 1'), 2)
        self.assertEqual(lectura(archivo, 'Examen 2'), 3)
        self.assertEqual(len(llamadas), 3)
        self.assertEqual(llamadas[0][1], b'%PDF-1.4 contenido original')
        self.assertEqual(llamadas[1][1], b'%PDF-1.4 contenido actualizado')
        self.assertLessEqual(len(st.session_state['cache_apa_pdf']), 12)

    def test_boton_reinicio_sigue_funcionando_con_menos_actualizaciones(self):
        self.assertIn('@st.fragment\ndef _reinicio_flotante_rapido():', APP)
        self.assertNotIn('run_every=1.0', APP)
        self.assertIn('SEGUNDOS_CONFIRMACION_REINICIO = 3.0', APP)
        self.assertIn('SEGUNDOS_FLASH_CONFIRMACION = 1.5', APP)
        self.assertIn('mostrar_reinicio_flotante()', APP)


if __name__ == '__main__':
    unittest.main()
