"""Pruebas de etiquetas y procedencia documental sin fuentes supuestas."""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import procedencia_campos as pc
import fuentes_documentales as fd

ROOT = Path(__file__).resolve().parents[1]
APP = (ROOT / 'app.py').read_text(encoding='utf-8')
README = (ROOT / 'README.md').read_text(encoding='utf-8')


def test_interfaz_renombrada_sin_leyendas_ni_diagnostico():
    assert 'Administrar parámetros operacionales' in APP
    assert 'Administrar opciones (Añadir / Eliminar de las listas arriba)' not in APP
    assert 'OCR de compatibilidad con RapidOCR:' not in APP
    assert 'Solo se propondrán datos explícitos del documento escaneado para campos vacíos.' not in APP
    assert 'Diagnóstico de lectura del documento escaneado' not in APP
    assert '"Fichas técnicas", "Documentos escaneados"' in APP
    assert 'help="Fichas técnicas: análisis habitual. Documentos escaneados: lectura de PDF escaneados y manuales."' in APP


def test_fuentes_no_documentales_eliminadas_de_listado():
    for campo, valor in [('Código', 'A1'), ('Muestra requerida y recipiente', 'Plasma'),
                         ('Recepción de muestras', 'Tubo'), ('Referencias', 'Referencia')]:
        fuentes, _ = pc.procedencia_campo(campo, valor, es_escaneado=True)
        assert 'Edición manual' not in fuentes
        assert 'Ficha consolidada existente' not in fuentes
    fuentes, aplicada = pc.procedencia_campo('Referencias', 'Documento técnico')
    assert fuentes == [pc.REFERENCIAS]
    assert not any('PDF importado' in f for f in fuentes)
    assert aplicada is None  # no se demuestra procedencia por texto preexistente
    _, aprobada = pc.procedencia_campo('Referencias', 'Documento técnico',
                                      referencias_aplicadas='Documento técnico')
    assert aprobada == pc.REFERENCIAS
    assert pc.PDF_ESCANEADO == 'Escáner'


def test_sin_fuente_cuando_no_hay_prueba_o_campo_editado():
    fuentes, fuente = pc.procedencia_campo('Código', 'Valor anterior')
    assert fuentes == [pc.CRS, pc.PDF_TECNICO]
    assert fuente is None
    # No atribuir un dato de Cartera a un campo que solo admite Acreditación.
    fuentes_extra, falsa = pc.procedencia_campo('Muestra requerida y recipiente', 'Plasma',
                                                campos_cartera={'Muestra requerida y recipiente': 'Plasma'})
    assert fuentes_extra == ['Cartera de acreditación']
    assert falsa is None
    _, editada = pc.procedencia_campo('Código', 'Valor editado', '',
                                     {'Código': 'Valor editado'}, {}, {'Código'})
    assert editada is None
    assert 'cuaderno-sin-fuente' in APP
    assert 'SIN FUENTE APLICADA</span>' in APP
    assert ' (sin fuente aplicada)' not in APP


def test_prioridades_clinicas_no_cambiaron():
    fuentes, aplicada = pc.procedencia_campo('Preparación del paciente', 'No aplica',
        'Cartera de Acreditación · Específico', {}, {}, set(), False,
        {'Preparación del paciente':'No aplica'})
    assert fuentes == list(fd.PRIORIDAD_FUENTES_POR_CAMPO['Preparación del paciente'])
    assert aplicada == 'Cartera de acreditación'
    fuentes, aplicada = pc.procedencia_campo('Recepción de muestras', 'Texto', '',
        {'Recepción de muestras': 'Texto'}, {}, set(), True)
    assert fuentes == [pc.PDF_ESCANEADO]
    assert aplicada == pc.PDF_ESCANEADO


def test_se_mantiene_orden_tabla_y_no_inventan_faltantes():
    seccion = APP.split('with st.expander("Procedencia de campos documentales"', 1)[1].split(
        'st.markdown(', 1)[0]
    assert 'for campo in aliases_area:' in seccion
    assert 'cuaderno-fuente-seleccionada' in seccion
    assert 'cuaderno-sin-fuente' in seccion
    assert 'if not fuente_aplicada:' in seccion


def test_readme_nuevos_nombres_sin_efectos_en_ruta_sqlite():
    assert 'OCR.zip' in README and 'OCR/' in README
    assert 'No borres ni reemplaces' in README
    assert 'HERRAMIENTA_DB_PATH' in README
