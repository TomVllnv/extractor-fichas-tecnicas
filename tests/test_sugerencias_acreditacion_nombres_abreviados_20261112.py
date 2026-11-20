"""Regresión: la Cartera de Acreditación no debe perder nombres clínicos relevantes."""
from pathlib import Path
import json
import sqlite3
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import fuentes_documentales as fd


class TestSugerenciasAcreditacion(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.db = str(Path(tmp.name) / 'acreditacion.db')
        fd.inicializar_fuentes(self.db)
        self.area = 'Hematología y coagulación'
        registros = [
            ('ANTICOAGULANTE LÚPICO', 'HEMATOLOGÍA Y COAGULACIÓN'),
            ('ANTICOAGULANTES CIRCULANTES', 'HEMATOLOGÍA Y COAGULACIÓN'),
            ('AC. CIRCULANTES O AC. LUPICO NUEVO', 'HEMATOLOGÍA Y COAGULACIÓN'),
            ('ÁCIDO ÚRICO', 'BIOQUÍMICA'),
            # Distractores deliberadamente parecidos por letras, pero sin
            # palabras distintivas relacionadas con lúpico/circulantes.
            *[(f'AC TRATAMIENTO NUEVO {i}', 'HEMATOLOGÍA Y COAGULACIÓN')
              for i in range(18)],
        ]
        with sqlite3.connect(self.db) as con:
            con.executemany('''INSERT INTO prestaciones_acreditacion (
                fila_excel, identificador, codigo, prestacion, prestacion_norm,
                area, area_norm, datos_json, estados_json
            ) VALUES (?,?,?,?,?,?,?,?,?)''', [
                (i, str(i), str(i), nombre, fd.normalizar(nombre), area,
                 fd.area_equivalente(area), json.dumps({}), json.dumps({}))
                for i, (nombre, area) in enumerate(registros, 1)
            ])

    def test_sugerencias_priorizan_lupico_y_circulantes(self):
        result = fd.candidatos_acreditacion(
            self.db, 'AC. CIRCULANTES O AC. LUPICO NUEVO', self.area,
            limite=5, similitud_minima=0.55)
        nombres = [item[1] for item in result]
        self.assertEqual(nombres[0], 'AC. CIRCULANTES O AC. LUPICO NUEVO')
        self.assertLess(nombres.index('ANTICOAGULANTE LÚPICO'), 4)
        self.assertLess(nombres.index('ANTICOAGULANTES CIRCULANTES'), 4)
        self.assertNotIn('ÁCIDO ÚRICO', nombres)
        self.assertGreaterEqual(next(p for _, n, _, p in result if n == 'ANTICOAGULANTE LÚPICO'), .55)

    def test_busca_tambien_con_el_examen_crs_confirmado(self):
        # En pantalla el nombre activo puede ser abreviado o diferente del
        # examen que el usuario acaba de vincular a la Cartera CRS.
        q = 'EXAMEN CODIFICADO LOCALMENTE'
        ref = 'AC. CIRCULANTES O AC. LUPICO NUEVO'
        sin_referencia = fd.candidatos_acreditacion(
            self.db, q, self.area, limite=5, similitud_minima=0.55)
        self.assertNotIn(ref, [n for _, n, _, _ in sin_referencia])
        con_referencia = fd.candidatos_acreditacion(
            self.db, q, self.area, limite=5, similitud_minima=0.55,
            nombre_referencia=ref)
        self.assertEqual(con_referencia[0][1], ref)
        self.assertEqual(con_referencia[0][3], 1.0)
        # La selección manual del CRS jamás enlaza automáticamente los datos
        # de Acreditación: un examen activo diferente sigue sin match exacto.
        self.assertEqual(fd.coincidencia_exacta_acreditacion(self.db, q, self.area), (None, None))
        _, _, elegida, _ = fd.obtener_valores_fuentes(self.db, q, self.area)
        self.assertIsNone(elegida)

    def test_lupico_reconocido_con_abreviatura_y_sin_inventar_sinonimos(self):
        prioridad, similitud = fd._similitud_sugerencia_acreditacion(
            'AC. LÚPICO', 'ANTICOAGULANTE LUPICO')
        self.assertEqual(prioridad, 1)
        self.assertGreaterEqual(similitud, .55)
        prioridad_no, similitud_no = fd._similitud_sugerencia_acreditacion(
            'AC. LUPICO', 'ACIDO URICO')
        self.assertEqual(prioridad_no, 0)
        self.assertLess(similitud_no, similitud)

    def test_area_no_modifica_porcentaje(self):
        a = fd.candidatos_acreditacion(self.db, 'AC. LUPICO', self.area, 5, .55)
        b = fd.candidatos_acreditacion(self.db, 'AC. LUPICO', 'BIOQUÍMICA', 5, .55)
        self.assertEqual({n: p for _, n, _, p in a}, {n: p for _, n, _, p in b})

    def test_filtra_antes_del_limite_de_sugerencias(self):
        coincidentes = fd.candidatos_acreditacion(
            self.db, 'AC. CIRCULANTES O AC. LUPICO NUEVO', self.area,
            limite=3, similitud_minima=.55)
        self.assertEqual(len(coincidentes), 3)
        self.assertEqual({n for _, n, _, _ in coincidentes}, {
            'ANTICOAGULANTE LÚPICO', 'ANTICOAGULANTES CIRCULANTES',
            'AC. CIRCULANTES O AC. LUPICO NUEVO'
        })

    def test_interfaz_pasa_alternativa_crs_y_umbral_antes_de_limitar(self):
        app = (ROOT / 'app.py').read_text(encoding='utf-8')
        self.assertIn('nombre_referencia=seleccion_manual_cartera', app)
        self.assertIn('limite=5, similitud_minima=0.55', app)
        self.assertNotIn('if c[3] >= 0.55][:5]', app)


if __name__ == '__main__':
    unittest.main()
