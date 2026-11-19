"""Escenarios documentales de Derivados (sin nombres de laboratorio obligatorios)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import derivados_estructurado as de
import manuales_procedimiento as mp
import procedencia_campos as pc


SAF = """1. INTRODUCCIÓN
El síndrome antifosfolipídico (SAF) es una enfermedad autoinmune que se diagnostica mediante criterios clínicos y de laboratorio. Se caracteriza por una tendencia aumentada a desarrollar trombosis, pudiendo afectar vasos sanguíneos venosos, arteriales o de pequeño calibre. El SAF se vincula con complicaciones en el embarazo, como muerte fetal, abortos espontáneos recurrentes y partos prematuros asociados a problemas como preeclampsia. Para diagnosticar este síndrome, es necesario determinar la presencia de anticuerpos antifosfolípidos (aPL) en la sangre, como anticardiolipinas y anticoagulante lúpico. El SAF puede ser primario o estar relacionado con otras enfermedades autoinmunes como lupus eritematoso sistémico.
2. PREPARACIÓN DE REACTIVOS
Los reactivos deben conservarse en refrigeración."""


def test_informacion_clinica_conserva_cuatro_ideas_clinicas_del_saf():
    valor = mp.extraer_campos_manual(SAF)['Información clínica']
    for frase in ('enfermedad autoinmune', 'desarrollar trombosis',
                  'complicaciones en el embarazo', 'anticuerpos antifosfolípidos'):
        assert frase in valor
    assert 'El SAF puede ser primario' not in valor
    assert 'reactivos' not in valor
    assert '\n\n' not in valor  # En el origen son oraciones del mismo párrafo.


def test_clinica_no_incorpora_frases_tecnicas_intercaladas():
    documento = '''INTRODUCCIÓN
La ferritina es un marcador que se relaciona con las reservas de hierro.
El analizador se calibra mediante fotometría y reactivos.
La muestra debe centrifugarse en un tubo de ensayo.
MATERIALES Y REACTIVOS
Se utilizan calibradores.'''
    val = mp.extraer_campos_manual(documento)['Información clínica']
    assert 'reservas de hierro' in val
    assert 'analizador' not in val
    assert 'tubo' not in val


def test_derivados_ocr_lineas_con_campos_adyacentes_y_desconocidos():
    texto = '''Información general
Código BK: CARO
Tiempo de respuesta: Mismo día a las 20:00
Principio del método: Espectrofotometría
Propósito del examen: Confirmar el diagnóstico de carotenoderma y alteraciones de absorción de grasas.
Componentes/Alérgeno
Muestra y Transporte
Tipo de muestra primaria: Sangre
Tipo de muestra requerida: Suero
Volumen mínimo: 3 mL
Indicaciones toma de muestra: Ninguna
Preparación del paciente: Ninguna
Medio de transporte: Tubo de polipropileno
Estabilidad de la muestra: 48 horas
Condiciones de transporte: Temperatura ambiente
Observaciones adicionales: Instrucciones específicas para otras variantes de la prueba.'''
    datos = de.extraer_ficha_derivados(texto)
    assert datos['Información clínica'] == 'Confirmar el diagnóstico de carotenoderma y alteraciones de absorción de grasas.'
    assert datos['Método de procesamiento'] == 'Espectrofotometría'
    assert datos['Plazo entrega resultados'] == 'Mismo día a las 20:00'
    assert datos['Estabilidad de la muestra'] == '48 horas'
    assert datos['Condiciones de envío al Laboratorio'] == 'Temperatura ambiente'
    assert 'Suero' in datos['Muestra requerida y recipiente']
    assert 'Volumen mínimo: 3 mL' in datos['Muestra requerida y recipiente']
    assert 'polipropileno' not in str(datos)
    assert 'Código' not in datos
    assert 'Observaciones' not in str(datos)


def test_derivados_pdf_digital_tabla_de_dos_columnas_y_multilinea():
    import fitz
    pdf = fitz.open()
    pag = pdf.new_page(width=680, height=842)
    y = 100
    for titulo, valor in [
        ('Código BK', 'CARO'), ('Código FONASA', '302017'),
        ('Tiempo de respuesta', 'Mismo día a las 20:00'),
        ('Principio del método', 'Espectrofotometría'),
        ('Propósito del examen', 'Confirmar el diagnóstico de carotenoderma.'),
        ('Componentes/Alérgeno', ''), ('Muestra y Transporte',''),
        ('Tipo de muestra requerida', 'Suero'), ('Volumen mínimo', '3 mL'),
        ('Estabilidad de la muestra', '48 horas'),
        ('Condiciones de transporte', 'Temperatura ambiente'),
        ('Observaciones adicionales', 'No consumir alcohol antes de la toma.')]:
        pag.insert_text((28,y), titulo, fontsize=10)
        if valor:
            pag.insert_text((262,y), valor, fontsize=10)
        y+=27
    datos = de.extraer_ficha_derivados(pag.get_text('text'), pdf.tobytes())
    assert datos['Información clínica'] == 'Confirmar el diagnóstico de carotenoderma.'
    assert datos['Estabilidad de la muestra'] == '48 horas'
    assert datos['Condiciones de envío al Laboratorio'] == 'Temperatura ambiente'
    assert 'No consumir' not in str(datos)


def test_comparacion_detecta_discrepancia_sin_inventarla():
    assert de.son_valores_distintos('Estabilidad de la muestra', '48 horas', '72 horas')
    assert not de.son_valores_distintos('Método de procesamiento', 'Espectrofotometría', 'espectrofotometria')
    assert not de.son_valores_distintos('Preparación del paciente', 'Ninguna', '')
    assert de.son_valores_distintos('Equipamiento utilizado', 'Cobas T511', 'Cobas T500')
    assert not de.son_valores_distintos('Equipamiento utilizado', 'COBAS T511', 'Cobas T 511')


def test_procedencia_derivados_digital_conserva_primero_base():
    fuentes, aplicada = pc.procedencia_campo(
        'Condiciones de envío al Laboratorio', 'Temperatura ambiente',
        campos_pdf={'Condiciones de envío al Laboratorio': 'Temperatura ambiente'},
        es_derivados=True)
    assert fuentes.index('Cartera de acreditación') < fuentes.index(pc.PDF_TECNICO)
    assert aplicada == pc.PDF_TECNICO


def test_ocr_de_derivados_reconoce_etiquetas_sin_separador():
    texto = '''Tipo de muestra requerida Suero
Estabilidad de la muestra 48 horas
Condiciones de transporte Temperatura ambiente'''
    d = de.extraer_ficha_derivados(texto)
    assert d['Muestra requerida y recipiente'] == 'Suero'
    assert d['Estabilidad de la muestra'] == '48 horas'
    assert d['Condiciones de envío al Laboratorio'] == 'Temperatura ambiente'


def test_discrepancias_derivados_con_valores_institucionales_independientes():
    """Cotejo posterior a seleccionar entre varios PDF, sin tocar SQLite."""
    import ast
    from types import SimpleNamespace
    from pathlib import Path
    import equipamientos as eq
    app = ast.parse((Path(__file__).resolve().parents[1] / 'app.py').read_text())
    func = next(n for n in app.body if isinstance(n, ast.FunctionDef) and
                n.name == '_conflictos_cartera_actualizados')
    datos = {
        'Estabilidad de la muestra': '48 horas',
        'Condiciones de envío al Laboratorio': 'Temperatura ambiente',
        'Método de procesamiento': 'Espectrofotometría',
    }
    sesion = {
        'area_resultado': 'Derivados',
        'cartera_examen_emparejado': 'CAROTENO',
        'procedencia_valores_documentales': {
            'Estabilidad de la muestra': '72 horas',
            'Condiciones de envío al Laboratorio': 'Refrigerado',
        },
        'origen_campos_documentales': {
            'Estabilidad de la muestra': 'Cartera de Acreditación (Específico)',
            'Condiciones de envío al Laboratorio': 'Cartera de Acreditación (Específico)',
        },
        'pdf_para_cartera': dict(datos),
    }
    ns = {'st': SimpleNamespace(session_state=sesion), 'de': de, 'eq': eq,
          'obtener_info_cartera_por_nombre': lambda _: {}}
    exec(compile(ast.Module(body=[func], type_ignores=[]), '<app aislada>', 'exec'), ns)
    resultado = ns['_conflictos_cartera_actualizados'](datos)
    assert set(resultado) == {'Estabilidad de la muestra', 'Condiciones de envío al Laboratorio'}
    assert resultado['Estabilidad de la muestra'] == {
        'pdf': '48 horas', 'cartera': '72 horas',
        'fuente': 'Cartera de Acreditación (Específico)'}


def test_interfaz_muestra_boton_para_discrepancias_entre_pdfs():
    app = (Path(__file__).resolve().parents[1] / 'app.py').read_text()
    assert 'diferencias_pdf = st.session_state.get("discrepancias_pdf", {})' in app
    assert 'if diferencias_pdf or diferencias_cartera:' in app
    assert 'de.son_valores_distintos(campo, propuesto, valor_inst)' in app
