"""Cinco archivos separados en la barra lateral y sin cruces entre áreas."""
import io
import sqlite3
from pathlib import Path

import openpyxl
import pytest

import fuentes_documentales as fd
import homologacion_tests as ht


def excel(registros):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Registros'
    ws.append(('EXAMEN', 'AREA', '¿QUÉ EXAMEN O RESULTADOS INCLUYE?'))
    for nombre, area, resultados in registros:
        ws.append((nombre, area, resultados))
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def database(tmp_path):
    db = tmp_path / 'db.sqlite'
    fd.inicializar_fuentes(db)
    ht.inicializar(db)
    return db


def test_cinco_areas_realmente_independientes(tmp_path):
    db = database(tmp_path)
    assert len(ht.AREAS) == 5
    for area in ht.AREAS:
        ht.guardar(db, excel([('EXAMEN TEST', '', f'Tests de {area}')]),
                   f'{area}.xlsx', area_fuente=area)
    for area in ht.AREAS:
        valor, meta = ht.obtener_tests(db, examen='EXAMEN TEST', area=area)
        assert valor == f'Tests de {area}'
        assert meta['estado'] == 'encontrado'
        assert fd.documento_fuente(db, ht.id_fuente_area(area), True)['nombre_archivo'] == f'{area}.xlsx'
    assert len({ht.id_fuente_area(area) for area in ht.AREAS}) == 5
    assert len({ht.recurso_area(area) for area in ht.AREAS}) == 5


def test_reemplazar_un_area_no_borra_las_demais(tmp_path):
    db = database(tmp_path)
    hema = excel([('TEST A', 'Hematología', 'A1')])
    bio = excel([('TEST A', 'Bioquímica', 'B1')])
    ht.guardar(db, hema, 'hemato.xlsx', area_fuente='Hematología y coagulación')
    ht.guardar(db, bio, 'bio.xlsx', area_fuente='Bioquímica')
    ht.guardar(db, excel([('TEST A', 'Hematología', 'A2')]),
              'hemato_nuevo.xlsx', area_fuente='Hematología y coagulación')
    assert ht.obtener_tests(db, examen='TEST A', area='Hematología y coagulación')[0] == 'A2'
    assert ht.obtener_tests(db, examen='TEST A', area='Bioquímica')[0] == 'B1'
    assert fd.documento_fuente(db, ht.id_fuente_area('Bioquímica'), True)['archivo'] == bio
    assert ht.obtener_tests(db, examen='TEST A', area='Derivados')[0] == ''


def test_documento_area_incorrecta_no_reemplaza_vigente(tmp_path):
    db = database(tmp_path)
    anterior = excel([('TEST A', 'Hematología', 'Valor correcto')])
    ht.guardar(db, anterior, 'antes.xlsx', area_fuente='Hematología y coagulación')
    with pytest.raises(ValueError, match='otra área'):
        ht.guardar(db, excel([('TEST A', 'Bioquímica', 'Valor incorrecto')]),
                  'despues.xlsx', area_fuente='Hematología y coagulación')
    assert fd.documento_fuente(db, ht.id_fuente_area('Hematología y coagulación'), True)['archivo'] == anterior
    assert ht.obtener_tests(db, examen='TEST A', area='Hematología y coagulación')[0] == 'Valor correcto'


def test_no_rellenar_desde_otras_areas_y_sin_area_no_buscar_en_nuevas(tmp_path):
    db = database(tmp_path)
    ht.guardar(db, excel([('PRUEBA A', '', 'Sólo Hematología')]),
              'hema.xlsx', area_fuente='Hematología y coagulación')
    assert ht.obtener_tests(db, examen='PRUEBA A', area='Microbiología')[0] == ''
    assert ht.obtener_tests(db, examen='PRUEBA A', area='Derivados')[0] == ''
    assert ht.obtener_tests(db, examen='PRUEBA A')[0] == ''


def test_migracion_previa_exclusivamente_hematologica(tmp_path):
    db = database(tmp_path)
    original = excel([('PRUEBA A', 'Hematología', 'Test hematológico')])
    ht.guardar(db, original, 'anterior.xlsx')
    ht.inicializar(db)
    assert ht.obtener_tests(db, examen='PRUEBA A', area='Hematología y coagulación')[0] == 'Test hematológico'
    assert ht.obtener_tests(db, examen='PRUEBA A', area='Bioquímica')[0] == ''
    assert fd.documento_fuente(db, ht.id_fuente_area('Hematología y coagulación'), True)['archivo'] == original


def test_migracion_no_presume_que_un_documento_mixto_es_hematologia(tmp_path):
    db = database(tmp_path)
    original = excel([('PRUEBA A', 'Hematología', 'A'), ('PRUEBA B', 'Bioquímica', 'B')])
    ht.guardar(db, original, 'mixto.xlsx')
    ht.inicializar(db)
    assert ht.obtener_tests(db, examen='PRUEBA A', area='Hematología y coagulación')[0] == ''
    assert fd.documento_fuente(db, ht.id_fuente_area('Hematología y coagulación'), True) is None


def test_pestanas_y_confirmacion_por_area_en_interfaz():
    contenido = (Path(__file__).resolve().parents[1] / 'app.py').read_text()
    assert 'st.tabs(list(htest.AREAS))' in contenido
    assert 'area_fuente=area_ht' in contenido
    assert 'htest.recurso_area(area_ht)' in contenido


def test_migracion_desde_esquema_antiguo_sin_columna_area_fuente(tmp_path):
    db = tmp_path / 'esquema_anterior.sqlite'
    fd.inicializar_fuentes(db)
    original = excel([('AC. LUPICO', 'Hematología', 'Confirmación')])
    import hashlib
    with sqlite3.connect(db) as con:
        con.execute('''CREATE TABLE homologacion_tests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            hoja TEXT NOT NULL, fila_excel INTEGER NOT NULL,
            examen TEXT, examen_norm TEXT, codigo TEXT, codigo_norm TEXT,
            area TEXT, area_norm TEXT, resultados TEXT NOT NULL)''')
        con.execute('''INSERT INTO homologacion_tests
            (hoja,fila_excel,examen,examen_norm,codigo,codigo_norm,area,area_norm,resultados)
            VALUES (?,?,?,?,?,?,?,?,?)''',
            ('Registros', 2, 'AC. LUPICO', ht._norm('AC. LUPICO'), '', '',
             'Hematología', ht._norm('Hematología'), 'Confirmación'))
        con.execute('''INSERT INTO documentos_fuente
            (fuente_id,nombre_archivo,tipo_mime,archivo,sha256,fecha_actualizacion,registros)
            VALUES (?,?,?,?,?,?,?)''',
            (ht.FUENTE_ID, 'original.xlsx', ht.MIME_XLSX, original,
             hashlib.sha256(original).hexdigest(), '2026-11-19T01:00:00', 1))
    ht.inicializar(db)
    ht.inicializar(db)  # Idempotente al reiniciar Streamlit.
    assert ht.obtener_tests(db, examen='AC. LUPICO', area='Hematología y coagulación')[0] == 'Confirmación'
    assert ht.obtener_tests(db, examen='AC. LUPICO', area='Microbiología')[0] == ''
    assert fd.documento_fuente(db, ht.id_fuente_area('Hematología y coagulación'), True)['archivo'] == original
