"""Prioridad legible y referencias que conservan su estado al confirmar."""
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import fuentes_documentales as fd

APP = (ROOT / 'app.py').read_text(encoding='utf-8')


class TestPrioridadVisualYReferencias(unittest.TestCase):
    def test_orden_de_preparacion(self):
        self.assertEqual(fd.PRIORIDAD_FUENTES_POR_CAMPO['Preparación del paciente'], (
            'Cartera de Exámenes del CRS', 'Cartera de acreditación',
            'Normas Generales de Toma de Muestras',
        ))
        self.assertEqual(fd.identificar_fuente_procedencia(
            'Cartera de Acreditación · Heredado'), 'Cartera de acreditación')
        self.assertEqual(fd.identificar_fuente_procedencia(
            'Cartera de Exámenes del CRS · Indicaciones'), 'Cartera de Exámenes del CRS')
        self.assertEqual(fd.identificar_fuente_procedencia(
            'Normas Generales de Toma de Muestras'), 'Normas Generales de Toma de Muestras')

    def test_todos_los_campos_controlados_tienen_prioridad(self):
        self.assertEqual(set(fd.PRIORIDAD_FUENTES_POR_CAMPO), set(fd.CAMPOS_CONTROLADOS))
        self.assertEqual(fd.PRIORIDAD_FUENTES_POR_CAMPO['Recepción de muestras'],
                         ('Manual de toma de muestras',))
        self.assertEqual(fd.PRIORIDAD_FUENTES_POR_CAMPO['Rotulación'],
                         ('Cartera de acreditación', 'Normas Generales de Toma de Muestras'))

    def test_referencias_no_reinicia_contenedor_tras_aplicar(self):
        self.assertIn('key="contenedor_referencias"', APP)
        pos = APP.index('st.session_state["modo_referencia_aplicado"] = modo_referencia')
        fragmento = APP[pos:pos+180]
        self.assertNotIn('st.session_state["generacion_panel_referencias"]', fragmento)
        self.assertIn('st.rerun()', fragmento)
        self.assertIn('with st.expander("Referencias bibliográficas", expanded=False)', APP)


if __name__ == '__main__':
    unittest.main()
