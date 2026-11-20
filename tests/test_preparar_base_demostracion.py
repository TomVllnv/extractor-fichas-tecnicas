"""La limpieza de trazabilidad solamente afecta una copia de SQLite."""
import sqlite3
import tempfile
import unittest
from pathlib import Path
from preparar_base_demostracion import preparar_copia


class TestCopiaTrazabilidad(unittest.TestCase):
    def test_copia_sin_historial_conserva_original_y_fichas(self):
        with tempfile.TemporaryDirectory() as carpeta:
            raiz = Path(carpeta)
            origen, destino = raiz / 'original.db', raiz / 'demo.db'
            with sqlite3.connect(origen) as con:
                con.executescript('''
                    CREATE TABLE auditoria_cambios (id INTEGER PRIMARY KEY, accion TEXT);
                    CREATE TABLE eventos_acceso (id INTEGER PRIMARY KEY, evento TEXT);
                    CREATE TABLE examenes_area (id INTEGER PRIMARY KEY, nombre TEXT);
                    INSERT INTO auditoria_cambios VALUES (1, 'Modificar');
                    INSERT INTO eventos_acceso VALUES (1, 'Inicio sesión');
                    INSERT INTO examenes_area VALUES (1, 'FIBRINOGENO');
                ''')
            resumen = preparar_copia(origen, destino)
            self.assertEqual(resumen, {'auditoria_cambios': 1, 'eventos_acceso': 1})
            with sqlite3.connect(origen) as con:
                self.assertEqual(con.execute('SELECT COUNT(*) FROM auditoria_cambios').fetchone()[0], 1)
                self.assertEqual(con.execute('SELECT COUNT(*) FROM eventos_acceso').fetchone()[0], 1)
            with sqlite3.connect(destino) as con:
                self.assertEqual(con.execute('SELECT COUNT(*) FROM auditoria_cambios').fetchone()[0], 0)
                self.assertEqual(con.execute('SELECT COUNT(*) FROM eventos_acceso').fetchone()[0], 0)
                self.assertEqual(con.execute('SELECT nombre FROM examenes_area').fetchone()[0], 'FIBRINOGENO')
            with self.assertRaises(FileExistsError):
                preparar_copia(origen, destino)
            with self.assertRaises(ValueError):
                preparar_copia(origen, origen)

    def test_rechaza_esquema_no_valido_sin_tocar_original(self):
        with tempfile.TemporaryDirectory() as carpeta:
            raiz = Path(carpeta)
            origen, destino = raiz / 'original.db', raiz / 'demo.db'
            with sqlite3.connect(origen) as con:
                con.execute('CREATE TABLE examenes_area (nombre TEXT)')
                con.execute('INSERT INTO examenes_area VALUES (?)', ('FIBRINOGENO',))
            with self.assertRaisesRegex(ValueError, 'tablas de trazabilidad'):
                preparar_copia(origen, destino)
            self.assertFalse(destino.exists())
            with sqlite3.connect(origen) as con:
                self.assertEqual(con.execute('SELECT COUNT(*) FROM examenes_area').fetchone()[0], 1)


if __name__ == '__main__':
    unittest.main()
