"""Pruebas del trazado de líneas del panel de procedencia.

Las líneas deben depender del alto real de cada campo, no de un patrón
repetido independiente del contenido (que puede atravesar el texto).
"""
import re
import unittest
from pathlib import Path

APP = (Path(__file__).resolve().parents[1] / 'app.py').read_text(encoding='utf-8')


class TestCuadernoProcedencia(unittest.TestCase):
    def test_lineas_vinculadas_a_cada_fila_y_no_al_fondo(self):
        inicio = APP.index('  .cuaderno-procedencia {')
        fin = APP.index('  .cuaderno-campo {', inicio)
        estilos = APP[inicio:fin]
        self.assertNotIn('repeating-linear-gradient(', estilos)
        fila = re.search(r'\.cuaderno-fila\s*\{([^}]+)\}', estilos)
        self.assertIsNotNone(fila)
        self.assertRegex(fila.group(1), r'border-bottom\s*:\s*1px\s+solid')
        self.assertIn('box-sizing: border-box', fila.group(1))
        self.assertIn('min-height: 53px', fila.group(1))

    def test_conserva_fuentes_y_azul_marino(self):
        self.assertIn('color: #17365D', APP)
        self.assertIn('color: #556479', APP)
        self.assertIn('for campo in aliases_area:', APP)
        self.assertIn('html.escape(str(procedencia or fuente_aplicada or', APP)
        self.assertIn('cuaderno-sin-fuente', APP)
        self.assertIn('SIN FUENTE APLICADA', APP)
        self.assertIn('pc.procedencia_campo(', APP)
        self.assertIn('cuaderno-fuente-seleccionada', APP)
        self.assertIn('Procedencia de campos documentales', APP)


if __name__ == '__main__':
    unittest.main()
