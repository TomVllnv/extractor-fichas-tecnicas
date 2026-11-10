"""Regresiones del vínculo FIBRINOGENO vs FIBRINOGENO – HOSMET."""
from pathlib import Path
import sqlite3
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import cartera_crs as crs


class TestCoincidenciaExactaConVariante(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.db = str(Path(self.tmp.name) / 'cartera.db')
        with sqlite3.connect(self.db) as con:
            con.execute('''CREATE TABLE cartera_examenes (
                nombre_examen TEXT UNIQUE, laboratorio TEXT)''')
            con.executemany('INSERT INTO cartera_examenes VALUES (?,?)', [
                ('FIBRINOGENO', 'HEMATOLOGÍA Y COAGULACIÓN'),
                ('FIBRINOGENO – HOSMET', 'DERIVADO EN CONVENIO'),
                ('FIBRINOGENOS', 'HEMATOLOGÍA Y COAGULACIÓN'),
                ('HEMOGLOBINA', 'HEMATOLOGÍA Y COAGULACIÓN'),
            ])

    def test_vincula_exacto_y_deja_variante_alternativa(self):
        exacto = crs.coincidencia_exacta_cartera(self.db, 'Fibrinógeno', 'Hematología y coagulación')
        self.assertEqual(exacto, 'FIBRINOGENO')
        candidatos = crs.candidatos_cartera(self.db, 'Fibrinógeno', 'Hematología y coagulación')
        self.assertEqual(candidatos[0], ('FIBRINOGENO', 'HEMATOLOGÍA Y COAGULACIÓN', 1.0))
        derivado = next(c for c in candidatos if c[0] == 'FIBRINOGENO – HOSMET')
        self.assertTrue(0.55 < derivado[2] < 1.0)
        self.assertTrue(crs.es_variante_con_nombre_incluido('Fibrinógeno', derivado[0]))
        self.assertEqual(crs.candidatos_variantes_cartera(
            self.db, 'Fibrinógeno','Hematología y coagulación')[0][0], 'FIBRINOGENO – HOSMET')
        self.assertFalse(crs.es_variante_con_nombre_incluido('Fibrinógeno','FIBRINOGENOS'))

    def test_nombre_exacto_unico_y_variante_si_falta_area(self):
        with sqlite3.connect(self.db) as con:
            con.execute("UPDATE cartera_examenes SET laboratorio='' WHERE nombre_examen='FIBRINOGENO'")
        self.assertEqual(crs.coincidencia_exacta_cartera(self.db, 'FIBRINOGENO','Hematología y coagulación'),
                         'FIBRINOGENO')

    def test_regla_excepcional_solo_si_hay_variante_con_nombre_completo(self):
        self.assertIsNone(crs.coincidencia_exacta_cartera(self.db, 'HEMOGLOBINA', 'Microbiología'))
        self.assertFalse(crs.es_variante_con_nombre_incluido('FIBRINOGENO','FIBRI - HOSMET'))

    def test_nombres_exactos_duplicados_sin_area_compatible_son_ambiguos(self):
        with sqlite3.connect(self.db) as con:
            con.execute("INSERT INTO cartera_examenes VALUES ('FIBRINÓGENO','BIOQUÍMICA')")
        self.assertIsNone(crs.coincidencia_exacta_cartera(self.db, 'FIBRINOGENO', 'Microbiología'))
        self.assertEqual(crs.coincidencia_exacta_cartera(self.db, 'FIBRINOGENO','Bioquímica'),
                         'FIBRINÓGENO')

    def test_porcentaje_100_no_puede_ser_bonus_de_area(self):
        with sqlite3.connect(self.db) as con:
            con.execute("INSERT INTO cartera_examenes VALUES ('FIBRINOGENX','HEMATOLOGÍA Y COAGULACIÓN')")
        candidatos = crs.candidatos_cartera(self.db, 'FIBRINOGENO', 'Hematología y coagulación', 10)
        self.assertEqual(candidatos[0][0], 'FIBRINOGENO')
        self.assertEqual(candidatos[0][2], 1.0)
        self.assertLess(next(x[2] for x in candidatos if x[0] == 'FIBRINOGENX'), 1.0)


class TestAparienciaSolicitada(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.codigo = (ROOT / 'app.py').read_text(encoding='utf-8')

    def test_normas_no_tiene_subtitulo_repetido(self):
        self.assertNotIn('st.caption("Normas generales de toma de muestras")', self.codigo)
        self.assertIn('with tab_normas:', self.codigo)

    def test_etiqueta_vinculada_y_alternativas_persistentes(self):
        self.assertIn('Vinculada: **{seleccion_manual_cartera}**', self.codigo)
        self.assertIn('descripcion_vinculo = "exacta"', self.codigo)
        self.assertIn('crs.candidatos_variantes_cartera',self.codigo)
        self.assertIn('_vincular_candidato_cartera, nombre_sugerido',self.codigo)
        self.assertIn('mostrar_coincidencia_seleccionable(', self.codigo)

    def test_solo_titulo_posibles_coincidencias_sin_avisos(self):
        self.assertIn('<u><strong>Posibles coincidencias:</strong></u>', self.codigo)
        self.assertNotIn('separador-emparejamientos', self.codigo)
        self.assertIn('key="lista_coincidencias_crs"', self.codigo)
        self.assertNotIn('Posibles coincidencias. Elige una solo si corresponde al examen activo:',self.codigo)

    def test_division_entre_las_dos_carteras_sin_lineas_entre_sugerencias(self):
        bloque_inicio = self.codigo.index('st.markdown("##### Emparejamiento con **Cartera de Exámenes del CRS**")')
        bloque_final = self.codigo.index('st.markdown("##### Emparejamiento con **Cartera de acreditación**")', bloque_inicio)
        entre_bloques = self.codigo[bloque_inicio:bloque_final]
        self.assertIn('st.divider()\n            # Vinculación independiente', entre_bloques)
        self.assertNotIn('<hr class="separador-emparejamientos">', entre_bloques)
        acreditacion_hasta_siguiente = self.codigo[bloque_final:self.codigo.index('st.markdown("##### Parámetros operacionales de la ficha")', bloque_final)]
        self.assertNotIn('<hr class="separador-emparejamientos">', acreditacion_hasta_siguiente)

    def test_selector_actualiza_estado_antes_de_instanciarse(self):
        self.assertIn('crs.sincronizar_vinculo_en_sesion(', self.codigo)
        self.assertIn('on_change=_cambiar_selector_cartera', self.codigo)
        self.assertIn('cartera_seleccion_confirmada_manualmente', self.codigo)
        self.assertIn('revision_cartera', self.codigo)
        self.assertIn('if crs.coincidencia_exacta_cartera(DB_PATH, examen_activo, area_activa) is None:', self.codigo)


if __name__ == '__main__':
    unittest.main()
