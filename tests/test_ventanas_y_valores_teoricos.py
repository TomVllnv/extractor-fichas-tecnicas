"""Pruebas aisladas de diálogo/callback y extracción, sin arrancar Streamlit."""
import ast
from pathlib import Path
import re
import unittest

RUTA_APP = Path(__file__).resolve().parents[1] / 'app.py'
ORIGEN = RUTA_APP.read_text(encoding='utf8')
ARBOL = ast.parse(ORIGEN)


def cargar_funciones(nombres, constantes=()):
    nodos = []
    for nodo in ARBOL.body:
        if isinstance(nodo, (ast.FunctionDef, ast.AsyncFunctionDef)) and nodo.name in nombres:
            nodo_sin_decoradores = ast.FunctionDef(
                name=nodo.name,
                args=nodo.args,
                body=nodo.body,
                decorator_list=[],
                returns=nodo.returns,
                type_comment=nodo.type_comment,
                type_params=getattr(nodo, 'type_params', []),
            )
            nodos.append(nodo_sin_decoradores)
        if isinstance(nodo, ast.Assign) and any(isinstance(x, ast.Name) and x.id in constantes for x in nodo.targets):
            nodos.append(nodo)
    namespace = {'re': re}
    exec(compile(ast.fix_missing_locations(ast.Module(body=nodos, type_ignores=[])), str(RUTA_APP), 'exec'), namespace)
    return namespace


FUNCIONES = cargar_funciones([
    'extraer_valores_teoricos_hematologia', '_limpiar_fallback_valores_teoricos',
    'quitar_citas_superindices', 'limpiar_valor_excel', 'convertir_subindices',
    '_superindices_visuales_en_pagina',
    'es_linea_basura_o_encabezado', '_limpiar_seleccion_area',
    '_cerrar_confirmacion_por_x', '_cerrar_alerta_por_x',
    '_cerrar_candidatos_por_x', '_cerrar_discrepancias_por_x',
    '_cerrar_planilla_por_x', '_cerrar_examenes_por_x',
    '_cerrar_administracion_por_x', '_cerrar_alta_usuario_login',
], constantes=('SECCIONES_DOCUMENTO_MAESTRAS',))


class StreamlitFalso:
    def __init__(self, estado):
        self.session_state = estado


class TestCierreDeDialogos(unittest.TestCase):
    def test_los_seis_dialogos_restantes_tienen_callback_de_cierre(self):
        decoradores = re.findall(r'@st\.dialog\((.*)\)', ORIGEN)
        self.assertEqual(len(decoradores), 7)
        self.assertTrue(all('on_dismiss=_cerrar_' in x for x in decoradores))

    def test_cierre_del_dialogo_no_reabre_el_modal(self):
        estados = [
            ('_cerrar_confirmacion_por_x', 'cambio_pendiente', {'tipo': 'normas'}),
            ('_cerrar_alerta_por_x', 'mostrar_dialogo_expiracion', True),
            ('_cerrar_candidatos_por_x', 'candidatos_cartera_pendientes', [('uno', 90)]),
            ('_cerrar_discrepancias_por_x', 'modal_discrepancias_unificado', True),
            ('_cerrar_planilla_por_x', 'modal_confirmar_planilla_abierto', True),
            ('_cerrar_examenes_por_x', 'mostrar_modal_examenes', True),
            ('_cerrar_administracion_por_x', 'modal_admin_usuarios_abierto', True),
            ('_cerrar_alta_usuario_login', 'alta_usuario_login_abierta', True),
        ]
        for funcion, clave, valor in estados:
            with self.subTest(funcion=funcion):
                estado = {clave: valor, 'examen_activo': 'Glucosa'}
                FUNCIONES[funcion].__globals__['st'] = StreamlitFalso(estado)
                FUNCIONES[funcion]()
                self.assertFalse(estado.get(clave))
                self.assertEqual(estado['examen_activo'], 'Glucosa')

    def test_limpiar_seleccion_usa_callback_previo_al_widget(self):
        self.assertIn('on_click=_limpiar_seleccion_area', ORIGEN)
        estado = {'selector_examen_Hematología y coagulación': 'Hemograma',
                  'examen_activo': 'Hemograma', 'mostrar_modal_examenes': True,
                  'modal_area_activa': 'Hematología y coagulación'}
        FUNCIONES['_limpiar_seleccion_area'].__globals__['st'] = StreamlitFalso(estado)
        FUNCIONES['_limpiar_seleccion_area']('Hematología y coagulación',
                                           'selector_examen_Hematología y coagulación')
        self.assertIsNone(estado['selector_examen_Hematología y coagulación'])
        self.assertIsNone(estado['examen_activo'])
        self.assertFalse(estado['mostrar_modal_examenes'])
        self.assertNotIn('modal_area_activa', estado)


class TestValoresDeReferencia(unittest.TestCase):
    def test_prioridad_para_mg_dl_sin_parrafos(self):
        ejemplo = ('Valores teóricos\n193‑412 mg/dL\nCada laboratorio debería comprobar...\n'
                   '220 - 490 mg/dL\nDatos específicos de funcionamiento del test\n4 mg/dL')
        self.assertEqual(FUNCIONES['extraer_valores_teoricos_hematologia'](ejemplo),
                         '193‑412 mg/dL | 220 - 490 mg/dL')

    def test_fallback_transcribe_ejemplo_de_fibrinogeno(self):
        ejemplo = ('Valores teóricos\n'
                   '< 0.5\u00a0µg de unidades de equivalente de fibrinógeno/mL (µg de UEF/mL).29\n'
                   'El equivalente de fibrinógeno indicado se basa en la cantidad de\n'
                   'fibrinógeno utilizada en la preparación del estándar original de\n'
                   'Asserachrom.\n'
                   'Cada laboratorio debería comprobar si los intervalos de referencia pueden\n'
                   'aplicarse a su grupo de pacientes y, en caso necesario, establecer sus\n'
                   'propios valores.\n'
                   '56\nCódigo :LC-GC-MA02\nCod. Caract.: APL 1.2\n'
                   'CLÍNICO MANUAL DEL DE CENTRO TOMA DE DE MUESTRAS REFERENCIA LBORATORIO DE SALUD DE _\n'
                   'PEÑALOLÉN CORDILLERA ORIENTE\nVigencia: 27/03/2028\nPágina : 57 de 75\n'
                   'Datos específicos de funcionamiento del test\n'
                   'No incluir un valor posterior: 100 mg/dL')
        resultado = FUNCIONES['extraer_valores_teoricos_hematologia'](ejemplo)
        self.assertTrue(resultado.startswith('< 0.5 µg de unidades'))
        self.assertIn('Cada laboratorio debería comprobar', resultado)
        self.assertIn('propios valores.', resultado)
        for basura in ('Código', 'Vigencia:', 'Página :', 'MANUAL DEL', '100 mg/dL', ').29'):
            self.assertNotIn(basura, resultado)

    def test_citas_unicode_y_aplanadas_dentro_de_parrafos(self):
        quitar = FUNCIONES['quitar_citas_superindices']
        self.assertEqual(quitar('La muestra²⁹ corresponde al paciente³, sin cambios.'),
                         'La muestra corresponde al paciente, sin cambios.')
        self.assertEqual(quitar('La muestra29 corresponde al paciente23, sin cambios.'),
                         'La muestra corresponde al paciente, sin cambios.')
        self.assertEqual(quitar('equivalente de fibrinógeno/mL (µg de UEF/mL).²⁹'),
                         'equivalente de fibrinógeno/mL (µg de UEF/mL).')
        self.assertEqual(quitar('equivalente de fibrinógeno/mL (µg de UEF/mL).29'),
                         'equivalente de fibrinógeno/mL (µg de UEF/mL).')

    def test_pdf_detecta_superindice_visual_aunque_sea_numero_normal(self):
        from types import SimpleNamespace
        class PaginaFalsa:
            def get_text(self, modo):
                self.modo = modo
                return {'blocks': [
                    {'number': 4, 'lines': [
                        {'spans': [{'text': 'equivalente (µg/mL).', 'flags': 0},
                                   {'text': '29', 'flags': 1},
                                   {'text': ' El resultado es 0.5 µg/mL.', 'flags': 0}]}
                    ]},
                    {'number': 5, 'lines': [
                        {'spans': [{'text': 'Se conserva la cifra 29 mg/dL', 'flags': 0}]}
                    ]},
                ]}
        func = FUNCIONES['_superindices_visuales_en_pagina']
        func.__globals__['fitz'] = SimpleNamespace(TEXT_FONT_SUPERSCRIPT=1)
        nuevos = func(PaginaFalsa())
        self.assertEqual(nuevos, {4: 'equivalente (µg/mL).²⁹ El resultado es 0.5 µg/mL.'})
        self.assertEqual(FUNCIONES['quitar_citas_superindices'](nuevos[4]),
                         'equivalente (µg/mL). El resultado es 0.5 µg/mL.')

    def test_valores_cientificos_y_subindices_no_se_descartan(self):
        quitar = FUNCIONES['quitar_citas_superindices']
        original = 'CO2, T3, CD4, HbA1c, 193-412 mg/dL, < 0.5 µg/mL, 10⁶, x², m², cm³.'
        self.assertEqual(quitar(original), original)

    def test_fallback_con_citas_voladas_entre_palabras(self):
        ejemplo = ('Valores teóricos\n'
                   '< 0.5 µg de unidades de equivalente de fibrinógeno/mL (µg de UEF/mL).²⁹\n'
                   'El equivalente de fibrinógeno indicado³⁰ se basa en la cantidad de\n'
                   'fibrinógeno utilizada en la preparación del estándar original.\n'
                   'Cada laboratorio debería comprobar los intervalos de referencia.\n'
                   'Datos específicos de funcionamiento del test\n')
        res = FUNCIONES['extraer_valores_teoricos_hematologia'](ejemplo)
        self.assertIn('El equivalente de fibrinógeno indicado se basa', res)
        self.assertIn('(µg de UEF/mL). El equivalente', res)
        self.assertNotIn('²⁹', res)
        self.assertNotIn('³⁰', res)

    def test_respetar_proximo_encabezado_no_solo_datos_especificos(self):
        ejemplo = ('Valores teóricos\n< 0.5 µg/mL\n'
                   'Límites e intervalos\nEste párrafo ya es de otra sección.\n'
                   'Datos específicos de funcionamiento del test\n')
        self.assertEqual(FUNCIONES['extraer_valores_teoricos_hematologia'](ejemplo),
                         '< 0.5 µg/mL')

    def test_sin_delimitacion_no_copiar_resto_del_pdf(self):
        ejemplo = 'Valores teóricos\n< 0.5 µg/mL\nPárrafo sin cierre delimitado.'
        self.assertEqual(FUNCIONES['extraer_valores_teoricos_hematologia'](ejemplo), '')

    def test_conservar_numero_clinico_aislado_si_no_es_pagina(self):
        ejemplo = 'Valores teóricos\nValor esperado:\n5\nDatos específicos de funcionamiento del test'
        self.assertIn('5', FUNCIONES['extraer_valores_teoricos_hematologia'](ejemplo))


if __name__ == '__main__':
    unittest.main()
