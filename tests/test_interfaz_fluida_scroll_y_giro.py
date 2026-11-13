"""Regresiones de la navegación suave sin requerir un navegador en Windows."""
from pathlib import Path
import unittest

ROOT = Path(__file__).resolve().parents[1]
APP = (ROOT / 'app.py').read_text(encoding='utf-8')
BRIDGE = (ROOT / 'interfaz_fluida.py').read_text(encoding='utf-8')


class TestNavegacionSuave(unittest.TestCase):
    def test_descarga_excel_sin_rerun(self):
        bloque = APP.split('key="download_updated_xlsx",', 1)[1][:260]
        self.assertIn('on_click="ignore"', bloque)
        bloque2 = APP.split('key="download_imported_excel",', 1)[1][:150]
        self.assertIn('on_click="ignore"', bloque2)

    def test_aplicar_referencias_conserva_su_funcion(self):
        self.assertIn('key="contenedor_referencias"', APP)
        self.assertIn('btn_aplicar_referencias_apa', APP)
        self.assertIn('actualizar_campo_excel(', APP)
        self.assertNotIn('duration: 560', BRIDGE)
        self.assertNotIn('detalles.open = false', BRIDGE)
        self.assertIn('registrar(\'referencias\')', BRIDGE)
        self.assertIn('UI.observador.observe(doc.body', BRIDGE)
        self.assertIn('UI.observador = null', BRIDGE)
        self.assertNotIn('setInterval', BRIDGE)
        self.assertNotIn('::details-content {', APP)

    def test_actualizacion_no_dispara_segundo_rerun(self):
        parte = APP.split('key="btn_trigger_save_template"', 1)[1]
        parte = parte.split('with col_next:', 1)[0]
        self.assertIn('modal_confirmar_planilla_abierto', parte)
        self.assertNotIn('st.rerun()', parte)

    def test_respuesta_a_scroll_del_usuario(self):
        self.assertIn("doc.addEventListener('wheel'", BRIDGE)
        self.assertIn("doc.addEventListener('touchmove'", BRIDGE)

    def test_flecha_pausada_sin_reorientar(self):
        css = APP.split('/* Giro limitado a tres segundos en CSS', 1)[0][-600:]
        self.assertIn('animation-play-state: paused', css)
        self.assertIn('animation-play-state: running !important', APP)
        self.assertIn('key="boton_reinicio_flotante"', APP)

    def test_no_toca_sqlite(self):
        self.assertNotIn('sqlite3', BRIDGE)
        self.assertNotIn('base_de_datos.db', BRIDGE)


if __name__ == '__main__':
    unittest.main()
