"""Comprueba prioridad real de Normas para tres campos y no afecta Preparación."""
import json
import sqlite3
import tempfile
import unittest
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import fuentes_documentales as fd


class TestNormasComoPrimeraFuente(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.db = str(Path(self.tmp.name) / 'datos.db')
        fd.inicializar_fuentes(self.db)
        self.normas = {
            'Preparación del paciente': 'Preparación de Normas',
            'Procedimiento de toma de muestras': 'Procedimiento de Normas',
            'Rotulación': 'Rotulación de Normas',
            'Criterios de rechazo': 'Criterios de Normas',
        }
        self.acreditacion = {
            'Preparación del paciente': 'Preparación de Acreditación',
            'Procedimiento de toma de muestras': 'Procedimiento de Acreditación',
            'Rotulación': 'Rotulación de Acreditación',
            'Criterios de rechazo': 'Criterios de Acreditación',
            'Muestra requerida y recipiente': 'Tubo de Acreditación',
            'Condiciones de envío al Laboratorio': 'Envío de Acreditación',
        }
        with sqlite3.connect(self.db) as cx:
            for campo, valor in self.normas.items():
                cx.execute('INSERT INTO normas_generales_campos(campo,contenido) VALUES (?,?)',
                           (campo, valor))
            cx.execute('''INSERT INTO prestaciones_acreditacion(
                fila_excel,prestacion,prestacion_norm,area,area_norm,datos_json,estados_json)
                VALUES (1, 'GLUCOSA', 'glucosa', 'Bioquímica', 'bioquimica', ?, ?)''',
                (json.dumps(self.acreditacion),
                 json.dumps({campo: 'Específico' for campo in self.acreditacion})))

    def test_los_tres_campos_priorizan_normas_con_ambas_fuentes_validas(self):
        valores, origen, seleccionado, mensaje = fd.obtener_valores_fuentes(self.db, 'GLUCOSA', 'Bioquímica')
        self.assertEqual(mensaje, 'exacta')
        self.assertIsNotNone(seleccionado)
        for campo in fd.CAMPOS_NORMAS_PRIORITARIAS:
            self.assertEqual(valores[campo], self.normas[campo], campo)
            self.assertEqual(origen[campo], 'Normas Generales de Toma de Muestras', campo)
            self.assertEqual(fd.PRIORIDAD_FUENTES_POR_CAMPO[campo],
                             ('Normas Generales de Toma de Muestras', 'Cartera de acreditación'))
        for campo in ('Muestra requerida y recipiente', 'Condiciones de envío al Laboratorio'):
            self.assertEqual(valores[campo], self.acreditacion[campo])

    def test_respaldo_acreditacion_cuando_faltan_normas(self):
        with sqlite3.connect(self.db) as cx:
            cx.execute("DELETE FROM normas_generales_campos WHERE campo='Rotulación'")
            cx.execute("UPDATE normas_generales_campos SET contenido='Pendiente' WHERE campo='Criterios de rechazo'")
        valores, origen, _, _ = fd.obtener_valores_fuentes(self.db, 'GLUCOSA', 'Bioquímica')
        self.assertEqual(valores['Procedimiento de toma de muestras'], 'Procedimiento de Normas')
        for campo in ('Rotulación', 'Criterios de rechazo'):
            self.assertEqual(valores[campo], self.acreditacion[campo])
            self.assertTrue(origen[campo].startswith('Cartera de Acreditación'), campo)

    def test_preparacion_preserva_orden_actual(self):
        valores, origen, _, _ = fd.obtener_valores_fuentes(self.db, 'GLUCOSA', 'Bioquímica')
        self.assertEqual(valores['Preparación del paciente'], 'Preparación de Acreditación')
        valores, origen, _, _ = fd.obtener_valores_fuentes(
            self.db, 'GLUCOSA', 'Bioquímica', indicaciones_cartera='Ayuno desde CRS')
        self.assertEqual(valores['Preparación del paciente'], 'Ayuno desde CRS')
        self.assertTrue(origen['Preparación del paciente'].startswith('Cartera de Exámenes'))

    def test_imagen_para_pestana_existe_con_su_nombre_correcto(self):
        self.assertTrue((ROOT / 'Pestaña.png').is_file())
        self.assertFalse((ROOT / 'Pesta├▒a.png').exists())
        self.assertIn('"Pestaña.png"', (ROOT / 'app.py').read_text(encoding='utf-8'))


if __name__ == '__main__':
    unittest.main()
