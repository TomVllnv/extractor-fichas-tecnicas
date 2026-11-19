"""Pruebas con el archivo original aportado por el usuario, sin base institucional."""
from pathlib import Path
import hashlib

import pytest

import fuentes_documentales as fd
import homologacion_tests as ht


@pytest.fixture
def ruta_excel():
    ruta = Path(ht.__file__).resolve().parent / 'recursos_iniciales' / 'Homologacion_de_test.xlsx'
    assert ruta.is_file(), 'Falta el Excel original incluido'
    return ruta


def test_estructura_real_identifica_24_examenes_y_cabeceras(ruta_excel):
    datos, meta = ht.leer_excel(ruta_excel.read_bytes())
    assert meta['hoja'] == 'MAPA SIMPLE 24'
    assert meta['registros'] == len(datos) == 24
    assert meta['columna_codigo'] == 0
    assert meta['columna_examen'] == 1
    assert meta['encabezados'][5] == '¿QUÉ EXAMEN O RESULTADOS INCLUYE?'
    grupos = {fila[6] for fila in datos}
    assert {'Hematología', 'Coagulación', 'Microscopía', 'Perfil compuesto'} <= grupos
    assert all(ht._area_celda_compatible(g, 'Hematología y coagulación') for g in grupos)
    assert not ht._area_celda_compatible('Microscopía', 'Microbiología')


def test_importar_incluido_recupera_descripcion_exacta_sin_cruces(ruta_excel, tmp_path):
    db = tmp_path / 'prueba_institucional.sqlite'
    fd.inicializar_fuentes(db)
    ht.inicializar(db)
    original = ruta_excel.read_bytes()
    meta = ht.guardar(db, original, ruta_excel.name, area_fuente='Hematología y coagulación')
    assert meta['registros'] == 24
    valor, detalle = ht.obtener_tests(db, examen='HEMOGRAMA AUTOMATIZADO',
                                    codigo='0301045-001', area='Hematología y coagulación')
    assert detalle['estado'] == 'encontrado'
    assert detalle['coincidencia'] == 'codigo_lis'
    assert 'WBC, RBC, HGB, HCT, PLT' in valor
    assert 'frotis solo por criterio' in valor
    fetal, _ = ht.obtener_tests(db, examen='HEMOGRAMA AUT. EN SANGRE FETAL',
                               codigo='0301045-002', area='Hematología y coagulación')
    assert 'NO incluye frotis' in fetal
    misc, _ = ht.obtener_tests(db, examen='MIELOGRAMA CRS (FROTIS)',
                              area='Hematología y coagulación')
    assert 'Citomorfología' in misc
    no_cruce, info = ht.obtener_tests(db, examen='HEMOGRAMA AUTOMATIZADO',
                                     codigo='0301045-001', area='Microbiología')
    assert not no_cruce and info['estado'] == 'sin_documento_area'
    fuente = fd.documento_fuente(db, ht.id_fuente_area('Hematología y coagulación'), incluir_archivo=True)
    assert hashlib.sha256(fuente['archivo']).digest() == hashlib.sha256(original).digest()


def test_identidad_contradictoria_no_asigna_otro_test(ruta_excel, tmp_path):
    db = tmp_path / 'test.sqlite'
    fd.inicializar_fuentes(db)
    ht.guardar(db, ruta_excel.read_bytes(), ruta_excel.name, area_fuente='Hematología y coagulación')
    valor, detalle = ht.obtener_tests(db, examen='HEMOGLOBINA', codigo='0301045-001',
                                    area='Hematología y coagulación')
    assert valor == ''
    assert detalle['estado'] == 'requiere_revision'
    valor, detalle = ht.obtener_tests(db, examen='Desconocido', area='Hematología y coagulación')
    assert valor == '' and detalle['estado'] == 'sin_coincidencia'


def test_archivo_real_no_modificado_y_no_autoinstala(ruta_excel, tmp_path):
    assert ht.archivo_inicial_hematologia() == ruta_excel.read_bytes()
    db = tmp_path / 'vacia.sqlite'
    fd.inicializar_fuentes(db)
    ht.inicializar(db)
    assert fd.documento_fuente(db, ht.id_fuente_area('Hematología y coagulación')) is None


def test_interfaz_permite_confirmar_documento_real():
    app = (Path(ht.__file__).resolve().parent / 'app.py').read_text(encoding='utf-8')
    assert 'Incorporar Excel de Hematología incluido' in app
    assert 'solicitar_cambio(' in app
    assert 'ARCHIVO_INICIAL_HEMATOLOGIA.name' in app


def test_plantilla_incluida_en_imagen_docker():
    raiz = Path(ht.__file__).resolve().parent
    ignorados = (raiz / '.dockerignore').read_text(encoding='utf-8')
    assert '*.xlsx' in ignorados
    assert '!recursos_iniciales/Homologacion_de_test.xlsx' in ignorados
