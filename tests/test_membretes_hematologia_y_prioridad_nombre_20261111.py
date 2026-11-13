"""Regresión de membretes en saltos de página y ranking por nombre clínico."""
from __future__ import annotations

import ast
import io
import re
import sqlite3
import tempfile
import unittest
from pathlib import Path

import fitz

ROOT = Path(__file__).resolve().parents[1]
SOURCE = (ROOT/'app.py').read_text(encoding='utf-8')
TREE = ast.parse(SOURCE)
NAMES = {
    'extract_pdf_text', '_membrete_clinico_hematologia',
    '_limpiar_membrete_ocr_hematologia', '_superindices_visuales_en_pagina',
    'norm', 'es_linea_basura_o_encabezado', 'parse_labeled_text',
    'clean_value', 'limpiar_valor_excel', 'quitar_citas_superindices',
    'convertir_subindices', 'es_alias_hematologia',
}
CONST = {'SECCIONES_DOCUMENTO_MAESTRAS'}
ENV = {'fitz': fitz, 're': re, 'io': io, 'FIELD_ALIASES_BASE': {},
       'FIELD_ALIASES_HEMATOLOGIA': {},
       'extraer_limite_deteccion_hematologia': lambda text: '',
       'extraer_valores_teoricos_hematologia': lambda text: '',
       'extraer_solo_valor_referencia': lambda text: text}
nodes = [node for node in TREE.body if (
    isinstance(node, ast.FunctionDef) and node.name in NAMES
    or isinstance(node, ast.Assign) and any(isinstance(t, ast.Name) and t.id in CONST for t in node.targets)
)]
exec(compile(ast.Module(body=nodes, type_ignores=[]), str(ROOT/'app.py'), 'exec'), ENV)


def pdf_demo():
    doc = fitz.open()
    p1 = doc.new_page()
    p1.insert_text((48, 80), 'PT Rec', fontsize=24)
    p1.insert_text((48, 104), 'Prothrombin Time', fontsize=11)
    p1.insert_text((48, 420), 'Estabilidad:', fontsize=12)
    p1.insert_text((48, 440), 'A 15-25 C: 12 horas', fontsize=11)
    p1.insert_text((48, 470), 'La muestra debe conservarse dentro del tubo.', fontsize=11)
    p1.insert_text((48, 755), '2024-02 V5.0 Espanol 1/5', fontsize=8)

    p2 = doc.new_page()
    p2.insert_text((48, 40), 'ms_07153414190V5.0', fontsize=8)
    p2.insert_text((48, 83), 'PT Rec', fontsize=24)
    p2.insert_text((48, 105), 'Prothrombin Time', fontsize=11)
    p2.insert_text((48, 240), 'a -20 C: 7 semanas', fontsize=11)
    p2.insert_text((48, 260), 'a -80 C: 11 meses', fontsize=11)
    p2.insert_text((48, 285), 'Las alicuotas deben descongelarse en condiciones adecuadas.', fontsize=10)
    p2.insert_text((48, 333), 'Material suministrado', fontsize=12)
    p2.insert_text((48, 355), 'Otro contenido que no pertenece a estabilidad.', fontsize=10)
    p2.insert_text((48, 730), '2024 Roche Diagnostics www.roche.com', fontsize=8)
    data = doc.tobytes()
    doc.close()
    return data


class TestMembreteHematologia(unittest.TestCase):
    def test_remueve_membrete_solo_de_los_campos_clinicos(self):
        datos = pdf_demo()
        texto, metodo = ENV['extract_pdf_text'](datos, 'ptrec.pdf', 'Protrombina', hematologia=True)
        self.assertIn('A 15-25 C: 12 horas', texto)
        self.assertIn('a -20 C: 7 semanas', texto)
        self.assertIn('a -80 C: 11 meses', texto)
        self.assertNotIn('ms_07153414190V5.0', texto)
        self.assertEqual(texto.count('PT Rec'), 1)  # En la portada original, no en página 2
        self.assertEqual(texto.count('Prothrombin Time'), 1)
        self.assertIn('Material suministrado', texto)
        self.assertIn('PyMuPDF', metodo)

    def test_no_modifica_otros_laboratorios_ni_pdf_original(self):
        datos = pdf_demo()
        texto, _ = ENV['extract_pdf_text'](datos, 'ptrec.pdf', 'Protrombina', hematologia=False)
        self.assertIn('ms_07153414190V5.0', texto)
        self.assertEqual(texto.count('PT Rec'), 2)
        with fitz.open(stream=datos, filetype='pdf') as original:
            self.assertIn('ms_07153414190V5.0', original[1].get_text())
            self.assertIn('www.roche.com', original[1].get_text())

    def test_ocr_sin_codigos_no_borra_lineas_clinicas(self):
        linea = 'Estabilidad\na -20 C: 7 semanas\nInstrucciones complementarias'
        self.assertEqual(ENV['_limpiar_membrete_ocr_hematologia'](linea, 2), linea)
        con_codigo = 'ms_07153414190V5.0\nPT Rec\nProthrombin Time\nEstabilidad\n12 horas'
        resultado = ENV['_limpiar_membrete_ocr_hematologia'](con_codigo, 2)
        self.assertIn('Estabilidad', resultado)
        self.assertIn('12 horas', resultado)
        self.assertNotIn('PT Rec', resultado)

    def test_referencias_usan_pdf_completo_original(self):
        import sys
        sys.path.insert(0, str(ROOT))
        import referencias_apa as apa
        fuente = apa.fuente_pdf('ptrec.pdf', pdf_demo(), 'Protrombina')
        self.assertEqual(fuente.url, 'roche.com')
        self.assertEqual(fuente.autor, 'Roche Diagnostics')


class TestOrdenCartera(unittest.TestCase):
    def test_por_nombre_clinico_antes_que_parecido_literal(self):
        import sys
        sys.path.insert(0, str(ROOT))
        import cartera_crs as crs
        with tempfile.TemporaryDirectory() as carpeta:
            db = str(Path(carpeta)/'prueba.db')
            with sqlite3.connect(db) as con:
                con.execute('CREATE TABLE cartera_examenes (nombre_examen TEXT, laboratorio TEXT)')
                con.executemany('INSERT INTO cartera_examenes VALUES (?,?)', [
                    ('PROTROMBINA (TP)', 'HEMATOLOGÍA Y COAGULACIÓN'),
                    ('PROGESTERONA', 'DERIVADO EN CONVENIO'),
                    ('PROTROMBINA TACO NUEVO', 'HEMATOLOGÍA Y COAGULACIÓN'),
                    ('PROLACTINA', 'HORMONAS E INMUNOLOGÍA'),
                ])
            candidatos = crs.candidatos_cartera(db, 'PROTROMBINA', 'Hematología y coagulación')
            nombres = [x[0] for x in candidatos]
            self.assertEqual(set(nombres[:2]), {'PROTROMBINA (TP)', 'PROTROMBINA TACO NUEVO'})
            self.assertGreater(nombres.index('PROGESTERONA'), nombres.index('PROTROMBINA TACO NUEVO'))
            self.assertLess(next(x[2] for x in candidatos if x[0] == 'PROTROMBINA TACO NUEVO'), 1.0)


if __name__ == '__main__':
    unittest.main()
