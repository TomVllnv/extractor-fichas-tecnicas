"""Las tablas referenciales no son estabilidad por sí mismas."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import derivados_estructurado as de


TABLA_HORIZONTAL = '''Tabla de Temperaturas
Ambiente          Refrigerada         Congelada
18 – 28 °C         2 – 8 °C             −18 °C
*Tabla referencial. La estabilidad de la muestra corresponde a la indicada anteriormente (según cada examen).'''


def test_tabla_tres_columnas_complementa_estabilidad_sin_asumir_tiempo_a_temperatura():
    texto = ('''Estabilidad de la muestra: 48 horas
'''
             '''Condiciones de transporte: Temperatura ambiente
'''
             + TABLA_HORIZONTAL)
    datos = de.extraer_ficha_derivados(texto)
    valor = datos['Estabilidad de la muestra']
    assert valor.startswith('48 horas\n')
    assert 'ambiente: 18–28 °C' in valor
    assert 'refrigerada' not in valor and '−18 °C' not in valor
    assert '48 horas a temperatura ambiente' not in valor
    assert datos['Condiciones de envío al Laboratorio'] == 'Temperatura ambiente'
    assert 'Tabla de Temperaturas' not in str(datos['Condiciones de envío al Laboratorio'])


def test_tabla_vertical_se_reconoce_con_nombres_y_rangos_intercalados():
    texto = '''Estabilidad de la muestra: 5 días refrigerada
Tabla de Temperaturas
Ambiente: 18-28 °C
Refrigerada: 2-8 °C
Congelada: -18 °C'''
    datos = de.extraer_ficha_derivados(texto)
    assert 'refrigerada: 2–8 °C' in datos['Estabilidad de la muestra']
    assert 'ambiente' not in datos['Estabilidad de la muestra']
    assert 'congelada' not in datos['Estabilidad de la muestra']


def test_tabla_aislada_no_se_interpreta_como_condicion_del_examen():
    texto = 'Estabilidad de la muestra: 48 horas\n' + TABLA_HORIZONTAL
    datos = de.extraer_ficha_derivados(texto)
    assert datos['Estabilidad de la muestra'] == '48 horas'


def test_sin_tabla_o_sin_estabilidad_no_se_anaden_temperaturas():
    texto = '''Estabilidad de la muestra: 48 horas
Condiciones de transporte: Temperatura ambiente
Ambiente 18 – 28 °C'''
    assert de.extraer_ficha_derivados(texto)['Estabilidad de la muestra'] == '48 horas'
    texto = '''Condiciones de transporte: Temperatura ambiente\n''' + TABLA_HORIZONTAL
    datos = de.extraer_ficha_derivados(texto)
    assert 'Estabilidad de la muestra' not in datos


def test_no_modifica_una_estabilidad_con_temperatura_ya_explicita():
    texto = '''Estabilidad de la muestra: 48 horas entre 2 y 8 °C
Condiciones de transporte: Ambiente
''' + TABLA_HORIZONTAL
    datos = de.extraer_ficha_derivados(texto)
    assert datos['Estabilidad de la muestra'] == '48 horas entre 2 y 8 °C'


def test_tabla_pdf_digital_espacial_y_multilinea():
    import fitz
    doc = fitz.open()
    pag = doc.new_page(width=700, height=850)
    for y, etiqueta, valor in (
        (100, 'Estabilidad de la muestra', '48 horas'),
        (145, 'Condiciones de transporte', 'Temperatura ambiente'),
    ):
        pag.insert_text((28,y), etiqueta, fontsize=11)
        pag.insert_text((270,y), valor, fontsize=11)
    pag.insert_text((28,400), 'Tabla de Temperaturas', fontsize=15)
    for x, tipo, temp in ((100, 'Ambiente', '18 - 28 °C'),
                          (325, 'Refrigerada', '2 - 8 °C'),
                          (540, 'Congelada', '-18 °C')):
        pag.insert_text((x,435), tipo, fontsize=11)
        pag.insert_text((x,465), temp, fontsize=10)
    pag.insert_text((28,505), 'Tabla referencial', fontsize=9)
    b = doc.tobytes()
    doc.close()
    tabla = de._tabla_temperaturas_en_pdf(b)
    assert tabla == {'ambiente': '18–28 °C', 'refrigerada': '2–8 °C', 'congelada': '−18 °C'}
    datos = de.extraer_ficha_derivados('Estabilidad de la muestra: 48 horas\nCondiciones de transporte: Temperatura ambiente', b)
    assert 'ambiente: 18–28 °C' in datos['Estabilidad de la muestra']
    assert 'Congelada' not in datos['Estabilidad de la muestra']


def test_complemento_sigue_siendo_valor_documental_y_no_cambia_prioridad():
    import procedencia_campos as pc
    texto = 'Estabilidad de la muestra: 48 horas\nCondiciones de transporte: Temperatura ambiente\n' + TABLA_HORIZONTAL
    datos = de.extraer_ficha_derivados(texto)
    fuentes, aplicada = pc.procedencia_campo(
        'Estabilidad de la muestra', datos['Estabilidad de la muestra'],
        campos_pdf=datos, es_derivados=True,
    )
    assert pc.PDF_TECNICO in fuentes
    assert aplicada == pc.PDF_TECNICO
    # La fuente institucional continúa priorizada en los campos para los que
    # realmente existe; no inventar Cartera para Estabilidad de la muestra.
    fuentes_transporte, _ = pc.procedencia_campo(
        'Condiciones de envío al Laboratorio', 'Temperatura ambiente',
        campos_pdf=datos, es_derivados=True,
    )
    assert fuentes_transporte.index('Cartera de acreditación') < fuentes_transporte.index(pc.PDF_TECNICO)
