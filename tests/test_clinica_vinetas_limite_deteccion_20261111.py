"""Regresiones clínicas y visuales sin necesidad de la base institucional."""
import ast
import io
from pathlib import Path
import re
import sqlite3
import sys
import tempfile
import unittest
from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import gestion_compartida as gp

FUENTE = (ROOT / 'app.py').read_text(encoding='utf-8')
ARBOL = ast.parse(FUENTE)
HEMATO = {
    'Información clínica': ['Características', 'características'],
    'Límite de detección': ['límite de detección', 'rango de linealidad'],
    'Valores de referencia': ['valores teóricos'],
}
FUNCIONES = {
    'parse_labeled_text', 'extraer_limite_deteccion_hematologia',
    'limpiar_valor_excel', 'quitar_citas_superindices', 'convertir_subindices',
    'norm', 'es_alias_hematologia', 'clean_value',
}
ENV = {
    're': re,
    'FIELD_ALIASES_BASE': HEMATO,
    'FIELD_ALIASES_HEMATOLOGIA': HEMATO,
    'SECCIONES_DOCUMENTO_MAESTRAS': [],
    'es_linea_basura_o_encabezado': lambda linea, examen_activo=None: False,
    'extraer_solo_valor_referencia': lambda txt: txt,
    'extraer_valores_teoricos_hematologia': lambda txt: '',
}
exec(compile(ast.Module(body=[n for n in ARBOL.body if isinstance(n, ast.FunctionDef) and n.name in FUNCIONES], type_ignores=[]), str(ROOT/'app.py'), 'exec'), ENV)


class ViñetasYLimites(unittest.TestCase):
    def test_vinetas_no_quedan_en_renglones_aislados(self):
        texto = ('Características\n'
                 'La monitorización de los productos de degradación de la fibrina puede contribuir a\n'
                 '▪\nConfirmar o refutar un diagnóstico de sospecha\n'
                 '▪\n\nEstimar el riesgo potencial para los pacientes con CID\n'
                 '▪\nControlar un tratamiento ya iniciado\n'
                 'Más allá de su importancia para el diagnóstico')
        parsed = ENV['parse_labeled_text'](texto, HEMATO)
        clin = parsed['Información clínica']
        self.assertIn('• Confirmar o refutar', clin)
        self.assertIn('• Estimar el riesgo', clin)
        self.assertIn('• Controlar un tratamiento', clin)
        self.assertNotRegex(clin, r'(?m)^\s*[▪•]\s*$')
        self.assertLess(clin.index('• Confirmar'), clin.index('• Estimar'))
        self.assertLess(clin.index('• Estimar'), clin.index('• Controlar'))
        # Escritura íntegra de los saltos de línea en una celda, como la exportación.
        wb = Workbook()
        ws = wb.active
        ws['B1'] = clin
        buffer = io.BytesIO()
        wb.save(buffer)
        ws = load_workbook(io.BytesIO(buffer.getvalue())).active
        self.assertEqual(ws['B1'].value, clin)

    def test_vinetas_normales_no_se_duplican(self):
        texto = 'Texto introductorio\n  • Primera indicación\n  • Segunda indicación'
        salida = ENV['limpiar_valor_excel'](texto)
        self.assertEqual(salida.count('•'), 2)
        self.assertIn('• Primera indicación', salida)

    def test_limite_numerico_hematologia_con_unidad(self):
        texto = ('Características\nTexto clínico\n'
                 'Límites e intervalos\nLímite de detección= 0,12 UEF/mL\n'
                 'Rango de linealidad: 10 a 300 UEF/mL\n')
        parseado = ENV['parse_labeled_text'](texto, HEMATO)
        self.assertEqual(parseado['Límite de detección'], '0,12 UEF/mL')

    def test_limite_entre_lineas_y_desigualdad(self):
        txt = 'Límites e intervalos\nLímite de detección =\n≤ 0,05 UEF/mL\nOtras notas.'
        self.assertEqual(ENV['extraer_limite_deteccion_hematologia'](txt), '≤ 0,05 UEF/mL')

    def test_ausencia_de_limite_no_copia_otro_parametro(self):
        txt = 'Rango de linealidad: 0,05 a 300 UEF/mL\nCaracterísticas\nTexto clínico'
        parsed = ENV['parse_labeled_text'](txt, HEMATO)
        self.assertNotIn('Límite de detección', parsed)

    def test_limite_sin_unidad_no_se_inventa(self):
        self.assertEqual(ENV['extraer_limite_deteccion_hematologia']('Límites e intervalos\nLímite de detección= 2'), '')

    def test_limites_e_intervalos_reales_ug_de_uef_ml(self):
        texto = (
            'Características\nEl test mide concentraciones del analito.\n'
            'Límites e intervalos\n'
            'Intervalo de medición\n0.20-9 µg de UEF/mL\n'
            'Límites inferiores de medición\n'
            'Límite de Blanco = 0,08 µg de UEF/mL\n'
            'Límite de Detección= 0,15 µg de UEF/mL\n'
            'Límite de Cuantificación = 0,20 µg de UEF/mL\n'
            'Valores teóricos\nDatos adicionales\n'
        )
        self.assertEqual(ENV['extraer_limite_deteccion_hematologia'](texto), '0,15 µg de UEF/mL')
        self.assertEqual(ENV['parse_labeled_text'](texto, HEMATO)['Límite de detección'],
                         '0,15 µg de UEF/mL')

    def test_limite_fuera_de_limites_e_intervalos_se_descarta(self):
        texto = ('Características\nLímite de Detección= 0,13 µg de UEF/mL\n'
                 'Límites e intervalos\nLímite de Blanco = 0,08 µg de UEF/mL\n'
                 'Precisión\nLímite de Detección= 0,99 µg de UEF/mL')
        self.assertEqual(ENV['extraer_limite_deteccion_hematologia'](texto), '')
        self.assertNotIn('Límite de detección', ENV['parse_labeled_text'](texto, HEMATO))

    def test_limite_compuesto_partido_en_lineas_y_decimales(self):
        texto = ('Limites e intervalos\nLímite de Detección =\n'
                 '≤ 0.15 μg de UEF/mL\nLímite de Cuantificación = 0.20 UEF/mL')
        self.assertEqual(ENV['extraer_limite_deteccion_hematologia'](texto),
                         '≤ 0.15 μg de UEF/mL')

    def test_mayusculas_preservadas_y_verificadas_en_inicio_sesion(self):
        with tempfile.TemporaryDirectory() as temp:
            db = str(Path(temp)/'prueba.db')
            gp.inicializar(db)
            gp.crear_usuario(db, 'ADMIN', 'Admin', 'ClaveDeAdministradorLarga2026')
            gp.administrar_cuenta(db, 'ADMIN', 0, 'crear', 'MariaLab',
                nombre_visible='María', rol='encargado', clave='ClaveDeLaboratorioMuyLarga2026')
            self.assertIsNone(gp.autenticar(db, 'marialab', 'ClaveDeLaboratorioMuyLarga2026'))
            login = gp.autenticar(db, 'MariaLab', 'ClaveDeLaboratorioMuyLarga2026')
            self.assertEqual(login['usuario'], 'MariaLab')
            gp.administrar_cuenta(db, 'ADMIN', 1, 'editar', 'MariaLab',
                nuevo_usuario='MariaLAB', nombre_visible='María', rol='encargado')
            self.assertEqual(gp.autenticar(db, 'MariaLAB', 'ClaveDeLaboratorioMuyLarga2026')['usuario'], 'MariaLAB')
            self.assertIsNone(gp.autenticar(db, 'MariaLab', 'ClaveDeLaboratorioMuyLarga2026'))
            self.assertEqual(gp.listar_usuarios(db, 'ADMIN')[1]['usuario'], 'MariaLAB')

    def test_modal_usuarios_ancho_con_pestanas_independientes(self):
        self.assertIn('@st.dialog("Administración de usuarios", width="large"', FUENTE)
        self.assertIn('st.dataframe(', FUENTE)
        self.assertIn('with tab_crear:', FUENTE)
        self.assertIn('with tab_editar:', FUENTE)
        self.assertIn('with tab_clave:', FUENTE)
        self.assertIn('with tab_estado:', FUENTE)
        for contenido in ('crear', 'editar', 'clave', 'estado'):
            self.assertIn(f'with contenido_{contenido}:', FUENTE)
        self.assertNotIn('with contenido_centro:', FUENTE)
        self.assertNotIn('st.caption("Eventos recientes.")', FUENTE)


if __name__ == '__main__':
    unittest.main()
