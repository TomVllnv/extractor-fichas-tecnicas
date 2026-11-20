"""Casos sin depender de la base institucional de producción."""
import io
import sqlite3
from pathlib import Path

import openpyxl

import fuentes_documentales as fd
import homologacion_tests as ht
import procedencia_campos as pc


def _excel(filas, encabezado='¿QUÉ EXAMEN O RESULTADOS INCLUYE?'):
    wb = openpyxl.Workbook()
    sh = wb.active
    sh.title = 'Homologacion'
    sh.append(['EXAMEN', 'CODIGO', 'AREA', encabezado])
    for fila in filas:
        sh.append(fila)
    b = io.BytesIO()
    wb.save(b)
    return b.getvalue()


def _db(tmp_path):
    db = tmp_path / 'datos.sqlite'
    fd.inicializar_fuentes(db)
    ht.inicializar(db)
    return db


def test_importa_excel_y_recupera_por_examen_y_codigo(tmp_path):
    db = _db(tmp_path)
    hematologia = _excel([('AC. LUPICO','03100','Hematología','Anticoagulante lúpico; confirmación')])
    bioquimica = _excel([('PERFIL HEPATICO','02402','Bioquímica','AST, ALT, FA, GGT')])
    meta = ht.guardar(db, hematologia, 'Hemato.xlsx', area_fuente='Hematología y coagulación')
    ht.guardar(db, bioquimica, 'Bioquimica.xlsx', area_fuente='Bioquímica')
    assert meta['registros'] == 1
    valor, detalle = ht.obtener_tests(db, examen='Perfil Hepático', area='Bioquímica', codigo='02402')
    assert valor == 'AST, ALT, FA, GGT'
    assert detalle['estado'] == 'encontrado'
    assert fd.documento_fuente(db, ht.id_fuente_area('Hematología y coagulación'), incluir_archivo=True)['archivo'] == hematologia
    assert fd.documento_fuente(db, ht.id_fuente_area('Bioquímica'), incluir_archivo=True)['archivo'] == bioquimica


def test_sin_codigo_no_inventa_relacion_y_con_ambiguedad_no_auto_completa(tmp_path):
    db = _db(tmp_path)
    ht.guardar(db, _excel([('AC. LUPICO', '', '', 'Prueba A'),
                          ('AC. LUPICO', '', '', 'Prueba B')]), 'homologacion.xlsx')
    valor, meta = ht.obtener_tests(db, examen='AC. LUPICO')
    assert valor == '' and meta['estado'] == 'requiere_revision'
    assert ht.obtener_tests(db, examen='AC. LUPIO')[0] == ''


def test_importacion_erronea_no_reemplaza_datos_anteriores(tmp_path):
    db = _db(tmp_path)
    ht.guardar(db, _excel([('PERFIL LIPIDICO','34','Bioquímica','HDL, LDL, colesterol')]), 'original.xlsx')
    import pytest
    with pytest.raises(ValueError):
        ht.guardar(db, _excel([('PERFIL LIPIDICO','34','Bioquímica','HDL')], encabezado='COLUMNA INCORRECTA'), 'nuevo.xlsx')
    assert ht.obtener_tests(db, examen='PERFIL LIPIDICO')[0] == 'HDL, LDL, colesterol'


def test_campos_reservados_exclusivamente_al_encargado():
    assert set(fd.CAMPOS_MANUAL) == {'Recepción de muestras','Procesamiento de muestras','Valor crítico'}
    for campo in fd.CAMPOS_MANUAL:
        fuentes, aplicada = pc.procedencia_campo(campo, '', es_escaneado=True,
                                               campos_pdf={campo: 'resultado falso'})
        assert fuentes == ['Completados manualmente por encargado']
        assert aplicada is None
        _, aplicada = pc.procedencia_campo(campo,'Texto del encargado',
                                             campos_editados={campo})
        assert aplicada == 'Completados manualmente por encargado'
        assert campo not in fd.MAPA_ACREDITACION


def test_columna_de_examen_no_conocida_puede_seleccionarse_expresamente(tmp_path):
    db = _db(tmp_path)
    blob = _excel([('AC. LUPICO','X1','Hematología','Pruebas confirmatorias')])
    # Sin la columna normalizada de examen no se debe adivinar una identidad.
    wb = openpyxl.load_workbook(io.BytesIO(blob))
    wb.active['A1'] = 'IDENTIFICACIÓN SEGÚN SISTEMA LOCAL'
    wb.active['B1'] = 'CODIGO FONASA'
    nuevo = io.BytesIO(); wb.save(nuevo)
    from pytest import raises
    with raises(ValueError):
        ht.leer_excel(nuevo.getvalue())
    ht.guardar(db, nuevo.getvalue(), 'Homologació de test.xlsx', columna_examen=0)
    valor, meta = ht.obtener_tests(db, examen='AC. LUPICO')
    assert valor == 'Pruebas confirmatorias' and meta['estado'] == 'encontrado'
    # El código FONASA no se equipara al Código LIS por defecto.
    assert ht.obtener_tests(db, examen='EXAMEN DESCONOCIDO', codigo='X1')[0] == ''


def test_area_incompatible_no_se_acepta_automaticamente(tmp_path):
    db = _db(tmp_path)
    blob = _excel([('Examen X','ABC','Microbiología','Panel A')])
    ht.guardar(db, blob, 'Homologació de test.xlsx', columna_codigo=1, area_fuente='Microbiología')
    valor, meta = ht.obtener_tests(db, examen='Examen X', area='Bioquímica', codigo='ABC')
    assert not valor and meta['estado'] == 'sin_documento_area'
