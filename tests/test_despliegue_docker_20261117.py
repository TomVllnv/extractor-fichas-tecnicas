"""Protecciones del paquete Docker sin necesidad del demonio Docker."""
import importlib.util
import sqlite3
import tempfile
import unittest
from pathlib import Path


RAIZ = Path(__file__).resolve().parents[1]


class TestDespliegueDocker(unittest.TestCase):
    def test_paths_y_separacion_datos(self):
        compose = (RAIZ / 'compose.yaml').read_text(encoding='utf8')
        dockerfile = (RAIZ / 'Dockerfile').read_text(encoding='utf8')
        self.assertIn('laboratorio_datos:/datos', compose)
        self.assertIn('HERRAMIENTA_DB_PATH: /datos/base_de_datos.db', compose)
        self.assertIn('127.0.0.1', compose)
        self.assertIn('USER 10001:10001', dockerfile)
        self.assertIn('/_stcore/health', dockerfile)
        self.assertIn('tesseract-ocr-spa', dockerfile)
        self.assertIn('TZ: America/Santiago', compose)

    def test_secretos_y_sqlite_excluidos_del_build(self):
        ignorados = (RAIZ / '.dockerignore').read_text(encoding='utf8')
        for pattern in ('*.db', '*.sqlite', '*.sqlite3', '.env', 'secrets.toml'):
            self.assertIn(pattern, ignorados)
        self.assertFalse((RAIZ / 'base_de_datos.db').exists())

    def test_importacion_no_sobrescribe_ni_altera_original(self):
        spec = importlib.util.spec_from_file_location(
            'importador_de_prueba', RAIZ / 'despliegue/importar_base_existente.py')
        modulo = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(modulo)
        with tempfile.TemporaryDirectory() as d:
            carpeta = Path(d)
            original = carpeta / 'original.db'
            destino = carpeta / 'datos' / 'base_de_datos.db'
            with sqlite3.connect(original) as conn:
                conn.execute('CREATE TABLE plantilla (texto TEXT)')
                conn.execute("INSERT INTO plantilla VALUES ('conservado')")
            modulo.importar(original, destino)
            with sqlite3.connect(destino) as conn:
                self.assertEqual(conn.execute('SELECT texto FROM plantilla').fetchone(), ('conservado',))
            with self.assertRaises(RuntimeError):
                modulo.importar(original, destino)
            with sqlite3.connect(original) as conn:
                self.assertEqual(conn.execute('SELECT texto FROM plantilla').fetchone(), ('conservado',))


if __name__ == '__main__':
    unittest.main()
