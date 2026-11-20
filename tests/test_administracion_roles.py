"""Migración segura de las cuentas existentes y permisos de administración."""
import concurrent.futures
import hashlib
import os
import sqlite3
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import gestion_compartida as gp


class AdminRolesTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.db = str(Path(self.temp.name) / 'aplicacion.db')
        # Emular la BD de la versión anterior con tres cuentas ya creadas.
        with sqlite3.connect(self.db) as con:
            con.execute('''CREATE TABLE usuarios_aplicacion (
              usuario TEXT PRIMARY KEY COLLATE NOCASE, nombre_visible TEXT NOT NULL,
              sal BLOB NOT NULL, clave_derivada BLOB NOT NULL,
              iteraciones INTEGER NOT NULL, activo INTEGER NOT NULL DEFAULT 1,
              creado_en TEXT NOT NULL)''')
            for username in ('ADMIN', 'encargado2', 'encargado3'):
                salt = os.urandom(32)
                clave = 'ClaveSeguraLaboratorio2026!'
                digest = hashlib.pbkdf2_hmac('sha256', clave.encode(), salt, gp.ITERACIONES)
                con.execute('''INSERT INTO usuarios_aplicacion
                  (usuario,nombre_visible,sal,clave_derivada,iteraciones,activo,creado_en)
                  VALUES (?,?,?,?,?,1,'2026-11-06T12:00:00-03:00')''',
                  (username, 'Nombre ' + username,salt,digest,gp.ITERACIONES))
        gp.inicializar(self.db)

    def test_migracion_preserva_contrasenas_e_incorpora_solo_admin(self):
        self.assertTrue(gp.es_administrador(self.db, 'admin'))
        self.assertFalse(gp.es_administrador(self.db, 'encargado2'))
        for u in ('ADMIN','encargado2','encargado3'):
            self.assertIsNotNone(gp.autenticar(self.db, u,'ClaveSeguraLaboratorio2026!'))
        self.assertEqual(gp.estado_recurso(self.db, 'usuarios:administracion')['revision'],1)
        gp.inicializar(self.db)
        self.assertEqual(gp.estado_recurso(self.db, 'usuarios:administracion')['revision'],1)
        with sqlite3.connect(self.db) as con:
            self.assertEqual(con.execute("SELECT count(*) FROM auditoria_cambios WHERE accion LIKE 'Asignar rol%' ").fetchone()[0], 1)

    def test_usuarios_normales_no_pueden_listar_ni_cambiar_usuarios(self):
        with self.assertRaises(PermissionError):
            gp.listar_usuarios(self.db,'encargado2')
        with self.assertRaises(PermissionError):
            gp.administrar_cuenta(self.db,'encargado2',1,'crear','noautorizado',
                                  nombre_visible='Nuevo', rol='encargado',clave='ClaveSeguraOtro2026')
        with sqlite3.connect(self.db) as con:
            self.assertEqual(con.execute('SELECT count(*) FROM usuarios_aplicacion').fetchone()[0],3)

    def test_admin_crea_usuario_y_registra_auditoria_sin_contrasena(self):
        gp.administrar_cuenta(self.db,'ADMIN',1,'crear','nuevo_usuario',
                             nombre_visible='Encargado nuevo',rol='encargado',clave='ClaveNuevaExtensa2026!')
        cuenta = gp.autenticar(self.db,'nuevo_usuario','ClaveNuevaExtensa2026!')
        self.assertEqual(cuenta['rol'],'encargado')
        with sqlite3.connect(self.db) as con:
            self.assertEqual(con.execute("SELECT count(*) FROM usuarios_aplicacion").fetchone()[0],4)
            registro = con.execute("SELECT detalle FROM auditoria_cambios ORDER BY id DESC LIMIT 1").fetchone()[0]
            self.assertNotIn('ClaveNuevaExtensa',registro)
        usuarios = gp.listar_usuarios(self.db,'ADMIN')
        self.assertNotIn('sal',usuarios[0])
        self.assertNotIn('clave_derivada',usuarios[0])

    def test_edicion_clave_y_estado_sin_borrar_historial(self):
        gp.administrar_cuenta(self.db,'ADMIN',1,'editar','encargado2',
                             nombre_visible='María',rol='encargado')
        gp.administrar_cuenta(self.db,'ADMIN',2,'clave','encargado2', clave='NuevaContrasenaMuySegura2026')
        self.assertIsNone(gp.autenticar(self.db,'encargado2','ClaveSeguraLaboratorio2026!'))
        self.assertIsNotNone(gp.autenticar(self.db,'encargado2','NuevaContrasenaMuySegura2026'))
        gp.administrar_cuenta(self.db,'ADMIN',3,'estado','encargado2',activo=False)
        self.assertIsNone(gp.autenticar(self.db,'encargado2','NuevaContrasenaMuySegura2026'))
        gp.administrar_cuenta(self.db,'ADMIN',4,'estado','encargado2',activo=True)
        self.assertIsNotNone(gp.autenticar(self.db,'encargado2','NuevaContrasenaMuySegura2026'))
        self.assertEqual(gp.estado_recurso(self.db,'usuarios:administracion')['revision'],5)

    def test_ultimo_administrador_no_puede_perder_el_rol_ni_desactivarse(self):
        with self.assertRaisesRegex(ValueError,'último administrador'):
            gp.administrar_cuenta(self.db,'ADMIN',1,'editar','ADMIN',
                                 nombre_visible='Nombre ADMIN', rol='encargado')
        with self.assertRaises(ValueError):
            gp.administrar_cuenta(self.db,'ADMIN',1,'estado','ADMIN',activo=False)
        self.assertTrue(gp.es_administrador(self.db,'ADMIN'))
        self.assertEqual(gp.estado_recurso(self.db,'usuarios:administracion')['revision'],1)

    def test_conflicto_de_dos_administradores_no_sobrescribe(self):
        gp.administrar_cuenta(self.db,'ADMIN',1,'editar','encargado2',
                             nombre_visible='Segundo admin',rol='administrador')
        def editar(actor,nombre):
            try:
                gp.administrar_cuenta(self.db,actor,2,'editar','encargado3',
                                     nombre_visible=nombre,rol='encargado')
                return 'ok'
            except gp.ConflictoConcurrente:
                return 'conflicto'
        with concurrent.futures.ThreadPoolExecutor(max_workers=2) as ex:
            a=ex.submit(editar,'ADMIN','primero')
            b=ex.submit(editar,'encargado2','segundo')
            self.assertEqual(sorted([a.result(), b.result()]),['conflicto','ok'])
        self.assertEqual(gp.estado_recurso(self.db,'usuarios:administracion')['revision'],3)

    def test_contrasenas_tienen_salt_diferente_incluso_si_son_iguales(self):
        gp.administrar_cuenta(self.db,'ADMIN',1,'crear','nuevousuario',
                             nombre_visible='Nuevo',rol='encargado',clave='ClaveSeguraLaboratorio2026!')
        with sqlite3.connect(self.db) as con:
            items = con.execute("SELECT sal,clave_derivada FROM usuarios_aplicacion WHERE usuario IN ('ADMIN','nuevousuario') ORDER BY usuario").fetchall()
        self.assertEqual(len(items),2)
        self.assertNotEqual(items[0][0],items[1][0])
        self.assertNotEqual(items[0][1],items[1][1])

    def test_restablecer_contrasena_invalida_sesiones_anteriores(self):
        vieja = gp.autenticar(self.db, 'encargado3', 'ClaveSeguraLaboratorio2026!')
        self.assertTrue(gp.usuario_habilitado(self.db, 'encargado3', vieja['revision_sesion']))
        gp.administrar_cuenta(self.db, 'ADMIN',1,'clave','encargado3',
                             clave='ContrasenaRestablecida2026!')
        self.assertFalse(gp.usuario_habilitado(self.db,'encargado3',vieja['revision_sesion']))
        nueva = gp.autenticar(self.db,'encargado3','ContrasenaRestablecida2026!')
        self.assertTrue(gp.usuario_habilitado(self.db,'encargado3',nueva['revision_sesion']))

    def test_cambio_de_rol_invalida_sesion_previa(self):
        vieja = gp.autenticar(self.db,'encargado2','ClaveSeguraLaboratorio2026!')
        gp.administrar_cuenta(self.db,'ADMIN',1,'editar','encargado2',
                             nombre_visible='Encargado segundo',rol='administrador')
        self.assertFalse(gp.usuario_habilitado(self.db,'encargado2',vieja['revision_sesion']))
        nuevo = gp.autenticar(self.db,'encargado2','ClaveSeguraLaboratorio2026!')
        self.assertTrue(gp.es_administrador(self.db,'encargado2'))
        self.assertEqual(nuevo['rol'],'administrador')

    def test_el_admin_ya_creado_esta_disponible_en_la_sesion(self):
        login=gp.autenticar(self.db,'ADMIN','ClaveSeguraLaboratorio2026!')
        self.assertEqual(login['rol'],'administrador')
        self.assertIsNone(gp.autenticar(self.db,'admin','ClaveSeguraLaboratorio2026!'))
        self.assertEqual(login['usuario'],'ADMIN')


if __name__ == '__main__':
    unittest.main()
