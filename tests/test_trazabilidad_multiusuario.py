"""Casos reales de concurrencia, cuentas y respaldo SQLite, sin navegador."""
import io
import os
import sys
import shutil
import tempfile
import sqlite3
from pathlib import Path
import unittest
from concurrent.futures import ThreadPoolExecutor

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import gestion_compartida as gp
import fuentes_documentales as fd


class TestPersistenciaMultiusuario(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.db = str(Path(self.temp.name) / "datos" / "base_de_datos.db")
        Path(self.db).parent.mkdir()
        shutil.copyfile(ROOT / "base_de_datos.db", self.db)
        gp.inicializar(self.db)
        fd.inicializar_fuentes(self.db)
        gp.sincronizar_legado(self.db)
        for nombre in ("operador1", "operador2", "operador3"):
            gp.crear_usuario(self.db, nombre, "Encargado " + nombre, nombre + "_ClaveSegura2026!")

    def test_3_cuentas_son_independientes_y_las_claves_no_estan_en_claro(self):
        for nombre in ("operador1", "operador2", "operador3"):
            valido = gp.autenticar(self.db, nombre, nombre + "_ClaveSegura2026!")
            self.assertEqual(valido["usuario"], nombre)
            self.assertIsNone(gp.autenticar(self.db, nombre, "clave-erronea"))
        with gp.conexion_cerrada(self.db) as con:
            self.assertEqual(con.execute("SELECT COUNT(*) FROM usuarios_aplicacion").fetchone()[0], 3)
            self.assertGreater(len(con.execute("SELECT sal FROM usuarios_aplicacion WHERE usuario='operador1'").fetchone()[0]), 16)

    def test_metadatos_preexistentes_no_inventan_usuario(self):
        estado = gp.estado_recurso(self.db, "fuente:manual_toma_muestras")
        self.assertEqual(estado["revision"], 0)
        self.assertIn("No registrado", estado["usuario"])
        self.assertIsNotNone(estado["fecha"])

    def test_carrera_entre_dos_operadores_impide_sobrescritura(self):
        recurso = "catalogo:laboratorios"
        version = gp.estado_recurso(self.db, recurso)["revision"]
        def operacion(con):
            con.execute("INSERT OR IGNORE INTO laboratorios(nombre,activo) VALUES('Laboratorio prueba',1)")
            return "Guardado"
        def intentar(usuario):
            try:
                gp.ejecutar_cambio(self.db, recurso, usuario, version, "Modificar lista", operacion)
                return "ok"
            except gp.ConflictoConcurrente:
                return "conflicto"
        with ThreadPoolExecutor(max_workers=2) as executor:
            uno = executor.submit(intentar, "operador1")
            dos = executor.submit(intentar, "operador2")
            self.assertEqual(sorted((uno.result(), dos.result())), ["conflicto", "ok"])
        self.assertEqual(gp.estado_recurso(self.db,recurso)["revision"], version + 1)
        cambios = [x for x in gp.historial(self.db) if x[2] == recurso]
        self.assertEqual(len(cambios), 1)

    def test_error_antes_de_confirmar_no_deja_cambios_ni_auditoria(self):
        recurso = "catalogo:laboratorios"
        def fallo(con):
            con.execute("INSERT OR IGNORE INTO laboratorios(nombre,activo) VALUES('Temporal',1)")
            raise ValueError("Archivo incorrecto")
        with self.assertRaisesRegex(ValueError, "Archivo incorrecto"):
            gp.ejecutar_cambio(self.db,recurso,"operador1",0,"Prueba",fallo)
        with gp.conexion_cerrada(self.db) as con:
            self.assertIsNone(con.execute("SELECT nombre FROM laboratorios WHERE nombre='Temporal'").fetchone())
        self.assertEqual(gp.estado_recurso(self.db,recurso)["revision"], 0)
        self.assertEqual(len(gp.historial(self.db)),0)

    def test_backup_consistente_con_wal(self):
        gp.ejecutar_cambio(self.db,"catalogo:encargados","operador3",0,"Registro demo",
            lambda con: con.execute("INSERT OR IGNORE INTO encargados(nombre,activo) VALUES('Auxiliar demo',1)"))
        ruta = gp.respaldar(self.db)
        with sqlite3.connect(ruta) as con:
            self.assertEqual(con.execute("PRAGMA integrity_check").fetchone()[0], "ok")
            self.assertEqual(con.execute("SELECT COUNT(*) FROM auditoria_cambios").fetchone()[0], 1)
            self.assertEqual(con.execute("SELECT COUNT(*) FROM cartera_examenes").fetchone()[0], 393)

    def test_fuentes_importadas_en_transaccion_comun(self):
        anterior = fd.documento_fuente(self.db, fd.NORMAS_ID, True)
        # Reimportar el PDF existente y auditar todo de manera atómica.
        gp.ejecutar_cambio(self.db,"fuente:normas_generales","operador2",0,
            "Actualizar normas",lambda con: fd.guardar_normas(self.db, anterior["archivo"],anterior["nombre_archivo"],conexion=con),
            archivo=anterior["nombre_archivo"])
        self.assertEqual(gp.estado_recurso(self.db,"fuente:normas_generales")["revision"],1)
        with gp.conexion_cerrada(self.db) as con:
            self.assertEqual(con.execute("SELECT COUNT(*) FROM normas_generales_campos").fetchone()[0],4)
            self.assertEqual(con.execute("SELECT COUNT(*) FROM auditoria_cambios").fetchone()[0],1)

    def test_usuarios_pueden_trabajar_con_lecturas_simultaneas(self):
        with ThreadPoolExecutor(max_workers=3) as executor:
            consultas = [executor.submit(fd.obtener_valores_fuentes, self.db, "VITAMINA D", "Hormonas e inmunidad") for _ in range(3)]
            resultados = [f.result()[0] for f in consultas]
        self.assertTrue(all("Recepción de muestras" in x for x in resultados))
        self.assertEqual(resultados[0], resultados[1])


if __name__ == '__main__':
    unittest.main()
