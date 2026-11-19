"""Evita que el top-N del CRS descarte alternativas por nombre parecido."""
from pathlib import Path
import sqlite3
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import cartera_crs as crs


class TestSugerenciasCRSAmpliadas(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.db = str(Path(tmp.name) / 'cartera.db')
        self.area = 'Hematología y coagulación'
        with sqlite3.connect(self.db) as conn:
            conn.execute('CREATE TABLE cartera_examenes (nombre_examen TEXT, laboratorio TEXT)')
            conn.executemany('INSERT INTO cartera_examenes VALUES (?, ?)', [
                ('PROTROMBINA (TP)', 'HEMATOLOGÍA Y COAGULACIÓN'),
                ('PROTROMBINA TACO NUEVO', 'HEMATOLOGÍA Y COAGULACIÓN'),
                # Estas sugerencias no incluyen el nombre clínico principal.
                ('PROGESTERONA', 'DERIVADO EN CONVENIO'),
                ('PROLACTINA', 'HORMONAS E INMUNOLOGÍA'),
                ('Procalcitonina', 'DERIVADO EN CONVENIO'),
                # Muchos registros de la MISMA área pero sin parecido:
                # antes ocupaban el top-8 y luego eran filtrados por 55 %.
                *[(f'HEMOGRAMA ESPECIAL {i}', 'HEMATOLOGÍA Y COAGULACIÓN')
                  for i in range(14)],
            ])

    def test_muestra_cinco_y_no_solo_las_dos_variantes(self):
        sugerencias = crs.candidatos_cartera(
            self.db, 'PROTROMBINA', self.area, limite=5, similitud_minima=0.55)
        nombres = [nombre for nombre, _, _ in sugerencias]
        self.assertEqual(len(sugerencias), 5)
        self.assertEqual(nombres[:2], ['PROTROMBINA (TP)', 'PROTROMBINA TACO NUEVO'])
        self.assertEqual(set(nombres[2:]), {'PROGESTERONA', 'PROLACTINA', 'Procalcitonina'})
        self.assertTrue(all(puntaje >= 0.55 for _, _, puntaje in sugerencias))

    def test_la_variante_completa_tiene_mayor_porcentaje_que_un_nombre_parecido(self):
        sugerencias = crs.candidatos_cartera(
            self.db, 'PROTROMBINA', self.area, limite=5, similitud_minima=0.55)
        porcentajes = {n: round(p * 100) for n, _, p in sugerencias}
        self.assertEqual(porcentajes['PROTROMBINA (TP)'], 91)
        self.assertEqual(porcentajes['PROTROMBINA TACO NUEVO'], 75)
        self.assertEqual(porcentajes['PROGESTERONA'], 70)
        self.assertGreater(porcentajes['PROTROMBINA TACO NUEVO'], porcentajes['PROGESTERONA'])
        self.assertLess(
            [n for n, _, _ in sugerencias].index('PROTROMBINA TACO NUEVO'),
            [n for n, _, _ in sugerencias].index('PROGESTERONA'))

    def test_variante_del_nombre_no_es_una_coincidencia_clinica_automatica(self):
        self.assertEqual(crs.similitud_nombre_cartera('PROTROMBINA', 'PROTROMBINA'), 1.0)
        self.assertGreater(crs.similitud_nombre_cartera('PROTROMBINA', 'PROTROMBINA TACO NUEVO'), 0.70)
        self.assertLess(crs.similitud_nombre_cartera('FIBRINOGENO', 'FIBRINOGENOS'), 1.0)
        self.assertFalse(crs.es_variante_con_nombre_incluido('FIBRINOGENO', 'FIBRINOGENOS'))
        self.assertIsNone(crs.coincidencia_exacta_cartera(self.db, 'PROTROMBINA', self.area))

    def test_el_area_no_influye_en_el_porcentaje(self):
        misma = crs.candidatos_cartera(self.db, 'PROTROMBINA', self.area, 5, 0.55)
        otra = crs.candidatos_cartera(self.db, 'PROTROMBINA', 'MICROBIOLOGÍA', 5, 0.55)
        self.assertEqual({n: p for n, _, p in misma}, {n: p for n, _, p in otra})

    def test_sin_cambios_en_enlace_automatico(self):
        self.assertIsNone(crs.coincidencia_exacta_cartera(self.db, 'PROTROMBINA', self.area))

    def test_la_interfaz_filtra_antes_de_limitar_a_cinco(self):
        app = (ROOT / 'app.py').read_text(encoding='utf-8')
        self.assertIn('top_n=5, similitud_minima=0.55', app)
        self.assertNotIn('if c[2] >= 0.55][:5]', app)


if __name__ == '__main__':
    unittest.main()
