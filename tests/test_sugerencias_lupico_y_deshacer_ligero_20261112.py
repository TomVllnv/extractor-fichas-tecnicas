"""Reproduce el caso «AC. LUPICO» y protege el deshacer sin rerun doble."""
from pathlib import Path
import json
import sqlite3
import sys
import tempfile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import cartera_crs as crs
import fuentes_documentales as fd
import historial_tabla as ht


NOMBRE = 'AC. LUPICO'
VARIANTE = 'AC. CIRCULANTES O AC. LUPICO NUEVO'
OTROS = ('ACIDO URICO', 'ACIDO LACTICO')
AREA = 'Hematología y coagulación'


def test_porcentaje_conserva_frase_clinica_completa_y_relega_falsos_positivos():
    esperado = fd.similitud_sugerencia_clinica(NOMBRE, VARIANTE)
    assert esperado[0] == 2
    assert 0.80 < esperado[1] < 1.0
    assert round(esperado[1] * 100) == 86
    for otro in OTROS:
        prioridad, similitud = fd.similitud_sugerencia_clinica(NOMBRE, otro)
        assert prioridad == 0
        assert similitud < .55
        assert similitud < esperado[1]
        assert abs(crs.similitud_nombre_cartera(NOMBRE, otro) - similitud) < 1e-9
    assert abs(crs.similitud_nombre_cartera(NOMBRE, VARIANTE) - esperado[1]) < 1e-9
    assert crs._prioridad_nombre_examen(NOMBRE, VARIANTE) == 2


def test_las_dos_carteras_ordenan_por_similitud_coherente_y_sin_enlazar():
    with tempfile.TemporaryDirectory() as tmp:
        db = str(Path(tmp) / 'fichas.db')
        fd.inicializar_fuentes(db)
        with sqlite3.connect(db) as con:
            con.execute('CREATE TABLE cartera_examenes (nombre_examen TEXT, laboratorio TEXT)')
            con.executemany('INSERT INTO cartera_examenes VALUES (?,?)', [
                (nombre, AREA if nombre == VARIANTE else 'BIOQUÍMICA')
                for nombre in (VARIANTE, *OTROS)
            ])
            con.executemany('''INSERT INTO prestaciones_acreditacion (
                fila_excel, identificador, codigo, prestacion, prestacion_norm,
                area, area_norm, datos_json, estados_json
            ) VALUES (?,?,?,?,?,?,?,?,?)''', [
                (i, str(i), str(i), nombre, fd.normalizar(nombre), AREA,
                 fd.area_equivalente(AREA), json.dumps({}), json.dumps({}))
                for i, nombre in enumerate((VARIANTE, *OTROS), 1)
            ])
        crs_result = crs.candidatos_cartera(db, NOMBRE, AREA, 5, .55)
        ac_result = fd.candidatos_acreditacion(db, NOMBRE, AREA, 5, .55)
        assert [n for n, _, _ in crs_result] == [VARIANTE]
        assert [n for _, n, _, _ in ac_result] == [VARIANTE]
        assert round(crs_result[0][2] * 100) == 86
        assert round(ac_result[0][3] * 100) == 86
        assert crs.coincidencia_exacta_cartera(db, NOMBRE, AREA) is None
        assert fd.coincidencia_exacta_acreditacion(db, NOMBRE, AREA) == (None, None)


def test_no_cambia_prioridad_anterior_de_protrombina_ni_nombre_exactos():
    assert round(crs.similitud_nombre_cartera('PROTROMBINA', 'PROTROMBINA (TP)')*100) == 91
    assert round(crs.similitud_nombre_cartera('PROTROMBINA', 'PROTROMBINA TACO NUEVO')*100) == 75
    assert round(crs.similitud_nombre_cartera('PROTROMBINA', 'PROGESTERONA')*100) == 70
    assert fd.similitud_sugerencia_clinica('AC. LUPICO', 'AC. LUPICO') == (3, 1.0)
    assert fd.similitud_sugerencia_clinica('LUPICO', 'LUPICOS')[1] < 1.0


def test_ctrl_z_una_sola_sincronizacion_sin_rehacer_en_interfaz():
    codigo = (ROOT / 'app.py').read_text(encoding='utf-8')
    assert "grid_options['context'] = {" in codigo
    assert "primerNodo.setDataValue('__historial_accion'" in codigo
    assert "node.data['Información extraída'] = valor;" in codigo
    assert "evento.shiftKey" in codigo
    assert "tecla === 'y'" not in codigo
    assert 'btn_rehacer_tabla' not in codigo
    assert 'actualizar_grid=False' in codigo
    bloque = codigo[codigo.index('            comando_historial = ht.comando_desde_filas('):
                    codigo.index('            seleccion_actual = st.session_state.setdefault(',
                                  codigo.index('            comando_historial = ht.comando_desde_filas('))]
    assert 'st.rerun()' not in '\n'.join(
        linea for linea in bloque.splitlines() if not linea.lstrip().startswith('#'))
    assert 'ht.registrar_cambios' in bloque
    assert 'campos_anteriores' in bloque
    assert 'if comando_historial == "deshacer"' in bloque
    assert 'Documentos (arrastrar para ordenar)' in codigo


def test_se_puede_recuperar_borrado_y_sigue_en_el_historial_del_servidor():
    historia = ht.contexto_nuevo({}, ('Hematología', 'AC. LUPICO', 0))
    datos = {'Estabilidad': '2 horas', 'Información clínica': 'Texto de prueba'}
    filas = [
        {'Campo': 'Estabilidad', 'Información extraída': '', 'Estado': 'No encontrado'},
        {'Campo': 'Información clínica', 'Información extraída': 'Texto de prueba', 'Estado': 'Encontrado'}
    ]
    assert ht.registrar_cambios(historia, datos, set(), filas)
    datos['Estabilidad'] = ''
    assert ht.aplicar_accion(historia, datos, set(), 'deshacer') == ['Estabilidad']
    assert datos['Estabilidad'] == '2 horas'
    assert historia['deshacer'] == []


def test_marca_oculta_no_vuelve_a_desencadenar_otro_deshacer():
    filas = [{'__historial_accion': 'deshacer:1234:abc', 'Campo': 'Estabilidad'}]
    marca = ht.marca_desde_filas(filas)
    consumida = None
    assert marca != consumida
    consumida = marca
    assert ht.marca_desde_filas(filas) == consumida
    assert ht.marca_desde_filas([{'__historial_accion': ''}]) is None
    assert ht.marca_desde_filas([{'__historial_accion': 'deshacer:1235:def'}]) != consumida
    codigo = (ROOT / 'app.py').read_text(encoding='utf-8')
    assert 'marca_historial != st.session_state.get("marca_historial_consumida")' in codigo
