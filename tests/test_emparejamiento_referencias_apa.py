"""Regresión: emparejamiento en la página y doble vista previa de APA."""
from pathlib import Path
import io
import unittest

import openpyxl

import referencias_apa as apa

SOURCE = (Path(__file__).resolve().parents[1] / 'app.py').read_text(encoding='utf-8')


class TestEmparejamientoEnLinea(unittest.TestCase):
    def test_nombre_exactamente_solicitado(self):
        self.assertIn('Emparejamiento con **Cartera de Exámenes del CRS**', SOURCE)
        self.assertIn('Emparejamiento con **Cartera de acreditación**', SOURCE)
        self.assertNotIn('Emparejamiento institucional (Cartera de Exámenes en SQLite)', SOURCE)

    def test_flujo_directo_y_vinculacion_opcional(self):
        self.assertNotIn('@st.dialog("Seleccionar coincidencia en Cartera de Exámenes"', SOURCE)
        self.assertNotIn('with st.expander("Opciones cuando no se encuentra el examen"', SOURCE)
        self.assertNotIn('¿Cómo deseas continuar?', SOURCE)
        self.assertNotIn('Continuar sin vinculación', SOURCE)
        self.assertNotIn('Preparar incorporación a la Cartera', SOURCE)
        self.assertIn('_vincular_candidato_cartera, nombre_sugerido', SOURCE)
        self.assertIn('Selecciona la prestación adecuada para evitar datos de otro examen.', SOURCE)
        self.assertIn('placeholder="Selecciona un examen (opcional)"', SOURCE)
        self.assertIn('cui.porcentaje_similitud(puntaje)', SOURCE)

    def test_acreditacion_examen_negrita_cursiva_area_normal(self):
        self.assertIn('st.button(etiqueta, key=clave_nombre, on_click=accion,', SOURCE)
        self.assertIn('f"Similitud: {cui.porcentaje_similitud(puntaje)} %"', SOURCE)
        self.assertNotIn('Similitud estimada:', SOURCE)

    def test_no_emparejar_nombres_parecidos_automaticamente(self):
        self.assertIn('crs.sincronizar_vinculo_en_sesion(', SOURCE)
        self.assertIn('for nombre_sugerido, laboratorio_sugerido, puntaje in sugeridas_crs:', SOURCE)
        self.assertIn('descripcion_vinculo = "exacta"', SOURCE)

    def test_previsualizacion_doble_antes_de_aplicar(self):
        self.assertIn('**Opción 1 · Solo PDF técnico**', SOURCE)
        self.assertIn('**Opción 2 · Todas las fuentes utilizadas**', SOURCE)
        self.assertIn('Aplicar opción seleccionada a Referencias', SOURCE)
        self.assertIn('Confirmo reemplazar mis referencias editadas', SOURCE)
        self.assertIn('cartera_examenes', SOURCE)


class TestAPA(unittest.TestCase):
    def test_nunca_usa_fecha_de_carga_como_publicacion(self):
        f = apa.fuente_institucional('normas_generales', {
            'nombre_archivo': 'documentos institucionales.docx.pdf',
            'fecha_actualizacion': '2026-11-06T17:00:00',
        })
        self.assertEqual(f.anio, '')
        self.assertIn('(s. f.)', apa.referencia_apa(f))
        self.assertNotIn('2026', apa.referencia_apa(f))
        self.assertTrue(f.titulo_provisional)
        self.assertNotIn('.docx', f.titulo)

    def test_version_sin_repeticion_del_numero(self):
        f = apa.fuente_institucional('cartera_acreditacion', {
            'nombre_archivo': 'Cartera_Acreditacion_CRS_v0.2.3.xlsx'
        })
        t = apa.referencia_apa(f)
        self.assertEqual(t.count('0.2.3'), 1)
        self.assertIn('versión 0.2.3', t)

    def test_aplica_orden_y_elimina_duplicados(self):
        f1 = apa.FuenteBibliografica('doc1', 'Zeta', anio='2023')
        f2 = apa.FuenteBibliografica('doc2', 'Alfa', anio='2024')
        texto = apa.ordenar_y_unir([f1, f2, f1])
        self.assertEqual(len(texto.split("\n\n")), 2)
        self.assertTrue(texto.startswith('Alfa. (2024).'))

    def test_previsualizacion_apa_html_segura_y_cursiva(self):
        f = apa.FuenteBibliografica('x', 'Manual <CRS> de muestras', autor='Centro & CRS', anio='2023')
        html = apa.ordenar_y_unir_html([f])
        self.assertIn('<em>Manual &lt;CRS&gt; de muestras</em>', html)
        self.assertIn('Centro &amp; CRS', html)
        self.assertNotIn('<CRS>', html)

    def test_referencias_separadas_en_excel_y_texto(self):
        fuentes = [apa.FuenteBibliografica('a', 'Alfa', anio='2023'),
                   apa.FuenteBibliografica('b', 'Beta', anio='2024')]
        self.assertIn('\n\n', apa.ordenar_y_unir(fuentes))
        self.assertIn('\n\n', str(apa.referencias_excel_con_cursiva(fuentes)))
        self.assertIn('<br><br>', apa.ordenar_y_unir_html(fuentes))

    def test_referencias_excel_plano_sin_cambiar_texto(self):
        f = apa.FuenteBibliografica('manual', 'Manual de toma de muestras',
                                    autor='CRSCO', anio='2023', version='04', tipo='Manual institucional')
        rich = apa.referencias_excel_con_cursiva([f])
        self.assertEqual(str(rich), apa.ordenar_y_unir([f]))
        wb = openpyxl.Workbook()
        wb.active['A1'] = rich
        buf = io.BytesIO()
        wb.save(buf)
        after = openpyxl.load_workbook(io.BytesIO(buf.getvalue()), rich_text=True)
        self.assertEqual(str(after.active['A1'].value), str(rich))
        self.assertIsInstance(after.active['A1'].value, str)


if __name__ == '__main__':
    unittest.main()
