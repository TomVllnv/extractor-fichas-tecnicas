"""Altas desde el login: autorización real de administrador, sin claves en claro."""
import sqlite3
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import gestion_compartida as gp
import registro_usuarios_login as rul


class AltaDesdeLoginTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.db = str(Path(self.temp.name) / 'bd_prueba.db')
        gp.inicializar(self.db)
        gp.crear_usuario(self.db, 'ADMIN', 'Administrador', 'ClaveSeguraInicial2026!')
        gp.crear_usuario(self.db, 'CRSCO1', 'Encargado 1', 'OtraClaveSegura2026!')
        self.revision = gp.estado_recurso(self.db, 'usuarios:administracion')['revision']

    def autorizar(self):
        ticket = rul.verificar_administrador(self.db, 'ADMIN', 'ClaveSeguraInicial2026!')
        self.assertIsNone(rul.verificar_administrador(self.db, 'admin', 'ClaveSeguraInicial2026!'))
        self.assertIsNotNone(ticket)
        return ticket

    def test_rechaza_clave_erronea_y_encargados(self):
        self.assertIsNone(rul.verificar_administrador(self.db,'ADMIN','Incorrecta2026!'))
        self.assertIsNone(rul.verificar_administrador(self.db,'CRSCO1','OtraClaveSegura2026!'))
        self.assertIsNone(rul.verificar_administrador(self.db,'NoExiste','ClaveSeguraInicial2026!'))

    def test_crea_encargado_y_registra_auditoria_sin_secretos(self):
        ticket = self.autorizar()
        resultado = rul.crear_encargado_autorizado(
            self.db, ticket, self.revision,
            usuario='CRSCO3', nombre_visible='Nuevo encargado',
            contrasena='ClaveNuevaSegura2026!', repetir_contrasena='ClaveNuevaSegura2026!',
        )
        self.assertIn('CRSCO3',resultado)
        usuario = gp.autenticar(self.db,'CRSCO3','ClaveNuevaSegura2026!')
        self.assertEqual(usuario['rol'],'encargado')
        with sqlite3.connect(self.db) as con:
            detalle = con.execute('SELECT detalle FROM auditoria_cambios ORDER BY id DESC LIMIT 1').fetchone()[0]
            self.assertNotIn('ClaveNuevaSegura2026!', detalle)
            self.assertEqual(con.execute("SELECT COUNT(*) FROM usuarios_aplicacion WHERE usuario='ADMIN'").fetchone()[0],1)
        self.assertEqual(gp.estado_recurso(self.db,'usuarios:administracion')['revision'], self.revision+1)

    def test_no_acepta_clave_distinta_ni_corta(self):
        ticket = self.autorizar()
        with self.assertRaisesRegex(ValueError,'coinciden'):
            rul.crear_encargado_autorizado(self.db,ticket,self.revision,
                usuario='nuevo',nombre_visible='Nuevo',contrasena='ClaveSeguraLarga',repetir_contrasena='Distinta!')
        with self.assertRaisesRegex(ValueError,'12'):
            rul.crear_encargado_autorizado(self.db,ticket,self.revision,
                usuario='nuevo',nombre_visible='Nuevo',contrasena='corta',repetir_contrasena='corta')
        self.assertIsNone(gp.autenticar(self.db,'nuevo','corta'))

    def test_no_se_reutiliza_autorizacion_caducada(self):
        ticket = self.autorizar()
        with patch.object(rul.time,'monotonic',return_value=ticket['vence']+0.01):
            self.assertFalse(rul.autorizacion_vigente(self.db,ticket))
            with self.assertRaises(PermissionError):
                rul.crear_encargado_autorizado(self.db,ticket,self.revision,
                    usuario='nuevo',nombre_visible='Nuevo',
                    contrasena='ClaveSeguraNueva2026!',repetir_contrasena='ClaveSeguraNueva2026!')

    def test_no_se_reutiliza_autorizacion_si_cambia_clave_admin(self):
        ticket = self.autorizar()
        gp.administrar_cuenta(self.db,'ADMIN',self.revision,'clave','ADMIN',clave='OtraNuevaClaveAdmin2026!')
        self.assertFalse(rul.autorizacion_vigente(self.db,ticket))

    def test_cuenta_existente_se_conserva_y_conflictos_no_sobrescriben(self):
        ticket = self.autorizar()
        with self.assertRaisesRegex(ValueError,'Ya existe'):
            rul.crear_encargado_autorizado(self.db,ticket,self.revision,
                usuario='CRSCO1',nombre_visible='Alterado',contrasena='NuevaClaveSegura2026!',
                repetir_contrasena='NuevaClaveSegura2026!')
        self.assertIsNotNone(gp.autenticar(self.db,'CRSCO1','OtraClaveSegura2026!'))
        gp.administrar_cuenta(self.db,'ADMIN',self.revision,'crear','OTRO',
            nombre_visible='Otro',rol='encargado',clave='ClaveOtro2026Segura!')
        with self.assertRaises(gp.ConflictoConcurrente):
            rul.crear_encargado_autorizado(self.db,ticket,self.revision,
                usuario='nuevo',nombre_visible='Nuevo',contrasena='NuevaClaveSegura2026!',
                repetir_contrasena='NuevaClaveSegura2026!')

    def test_no_eleva_rol_y_no_autoriza_encargado_recien_creado(self):
        ticket = self.autorizar()
        rul.crear_encargado_autorizado(self.db,ticket,self.revision,
            usuario='nuevo',nombre_visible='Nuevo',contrasena='NuevaClaveSegura2026!',
            repetir_contrasena='NuevaClaveSegura2026!')
        self.assertIsNone(rul.verificar_administrador(self.db,'nuevo','NuevaClaveSegura2026!'))


if __name__ == '__main__':
    unittest.main()
