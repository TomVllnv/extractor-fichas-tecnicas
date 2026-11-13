"""Pruebas sin BD institucional para nuevas cuentas, páginas y etiquetas históricas."""
import ast
import io
import re
import sqlite3
import sys
import tempfile
import unittest
from pathlib import Path
from openpyxl import Workbook, load_workbook
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
import gestion_compartida as gp
APP=(ROOT/'app.py').read_text(encoding='utf8')
ARBOL=ast.parse(APP)
funciones=('norm','nombre_campo_visible','normalizar_etiquetas_ficha',
           'buscar_hoja_por_examen','comparar_plantillas_excel')
namespace={'re':re,'io':io,'load_workbook':load_workbook}
exec(compile(ast.Module(body=[n for n in ARBOL.body if isinstance(n,ast.FunctionDef) and n.name in funciones],
                       type_ignores=[]), str(ROOT/'app.py'),'exec'),namespace)


def libro(items):
    wb=Workbook(); ws=wb.active; ws.title='FIBRINOGENO'
    for i,(campo,valor) in enumerate(items,1):
        ws.cell(i,1,campo);ws.cell(i,2,valor)
    bio=io.BytesIO();wb.save(bio);return bio.getvalue()


class CuentasYCampos(unittest.TestCase):
    def setUp(self):
        t=tempfile.TemporaryDirectory();self.addCleanup(t.cleanup)
        self.db=str(Path(t.name)/'prueba.db')
        gp.inicializar(self.db)
        gp.crear_usuario(self.db,'ADMIN','Administración','ClaveMaestraLaboratorio2026!')
        gp.crear_usuario(self.db,'operador1','Operador uno','ClaveOperadorLaboratorio2026!')
        gp.crear_usuario(self.db,'operador2','Operador dos','ClaveOperadorLaboratorio2026!')

    def test_cambiar_usuario_preserva_clave_y_cierra_sesion_vieja(self):
        anterior=gp.autenticar(self.db,'operador1','ClaveOperadorLaboratorio2026!')
        gp.administrar_cuenta(self.db,'ADMIN',0,'editar','operador1',
              nuevo_usuario='operador_modificado',nombre_visible='Operador uno',rol='encargado')
        self.assertFalse(gp.usuario_habilitado(self.db,'operador1',anterior['revision_sesion']))
        self.assertIsNotNone(gp.autenticar(self.db,'operador_modificado','ClaveOperadorLaboratorio2026!'))
        with sqlite3.connect(self.db) as con:
            detalle=con.execute('SELECT detalle FROM auditoria_cambios ORDER BY id DESC LIMIT 1').fetchone()[0]
        self.assertIn('operador_modificado',detalle)
        self.assertNotIn('ClaveOperador',detalle)

    def test_usuario_duplicado_no_realiza_operacion(self):
        with self.assertRaisesRegex(ValueError,'existe'):
            gp.administrar_cuenta(self.db,'ADMIN',0,'editar','operador1',
                 nuevo_usuario='OPERADOR2',nombre_visible='Operador uno',rol='encargado')
        self.assertEqual(gp.estado_recurso(self.db,'usuarios:administracion')['revision'],0)

    def test_cambiar_propio_admin_conserva_permisos_e_invalida_sesion(self):
        anterior=gp.autenticar(self.db,'ADMIN','ClaveMaestraLaboratorio2026!')
        gp.administrar_cuenta(self.db,'ADMIN',0,'editar','ADMIN',
              nuevo_usuario='ADMIN_NUEVO',nombre_visible='Administración',rol='administrador')
        self.assertFalse(gp.usuario_habilitado(self.db,'ADMIN',anterior['revision_sesion']))
        self.assertTrue(gp.es_administrador(self.db,'ADMIN_NUEVO'))

    def test_paginacion_sql_mas_recientes_primero(self):
        with sqlite3.connect(self.db) as con:
            for idx in range(26):
                con.execute("""INSERT INTO eventos_acceso(fecha,usuario,evento) VALUES ('2026-11-11','ADMIN',?)""",(f'evento {idx}',))
                con.execute("""INSERT INTO auditoria_cambios(fecha,usuario,recurso,accion,revision_anterior,revision_nueva)
                     VALUES ('2026-11-11','ADMIN','catalogo:encargados',?,0,1)""",(f'acción {idx}',))
        self.assertEqual(gp.total_historial(self.db),26)
        self.assertEqual(gp.total_historial_actividad(self.db),26)
        self.assertEqual(len(gp.historial(self.db,5)),5)
        self.assertEqual(gp.historial(self.db,5)[0][3],'acción 25')
        self.assertEqual(len(gp.historial(self.db,10,10)),10)
        self.assertEqual(len(gp.historial(self.db,10,20)),6)
        self.assertEqual(gp.historial_actividad(self.db,10,10)[0][2],'evento 15')

    def test_normalizacion_no_cambia_los_valores(self):
        wb=Workbook(); ws=wb.active
        ws['A1']='Autorizado por:';ws['A2']='Fecha actualización: ';ws['B1']='Entrega: 2 horas'
        namespace['normalizar_etiquetas_ficha'](ws)
        self.assertEqual(ws['A1'].value,'Autorizado por')
        self.assertEqual(ws['A2'].value,'Fecha actualización')
        self.assertEqual(ws['B1'].value,'Entrega: 2 horas')

    def test_comparar_plantilla_antigua_no_crea_cambios_falsos(self):
        antes=libro([('Autorizado por:','Marta'),('Fecha actualización:','11/11/2026')])
        ahora=libro([('Autorizado por','Marta'),('Fecha actualización','11/11/2026')])
        comp=namespace['comparar_plantillas_excel'](antes,ahora,'FIBRINOGENO')
        self.assertEqual(comp['diferencias'],[])
        self.assertEqual(comp['campos_completados_antes'],2)
        otro=libro([('Autorizado por','Gabriela'),('Fecha actualización','11/11/2026')])
        diferencia=namespace['comparar_plantillas_excel'](antes,otro,'FIBRINOGENO')['diferencias']
        self.assertEqual([f['Campo'] for f in diferencia],['Autorizado por'])

if __name__=='__main__': unittest.main()
