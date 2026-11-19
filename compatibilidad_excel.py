"""Compatibilidad de cadenas enriquecidas en libros Excel.

La aplicación utiliza openpyxl para crear los XLSX. Las versiones anteriores
convertían manualmente sus cadenas a ``xl/sharedStrings.xml``. Esa conversión
introducía fragmentos de texto enriquecido vacíos, que Microsoft Excel podía
reportar como «Registros reparados: Propiedades de cadena».

Aquí NO se edita el XML OOXML a mano. Solo cuando hace falta se vuelve a
serializar el libro con el escritor de openpyxl, eliminando fragmentos vacíos
y manteniendo sus estilos, viñetas y saltos de línea.
"""
from __future__ import annotations

import io
import zipfile
from xml.etree import ElementTree as ET

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
N = '{' + NS + '}'


LIMITE_CARACTERES_EXCEL = 32767


def _unidades_utf16(texto: str) -> int:
    """Excel contabiliza unidades UTF-16, no puntos Unicode de Python."""
    return len(texto.encode('utf-16-le', 'surrogatepass')) // 2


def validar_cadenas_ooxml(paquete: zipfile.ZipFile) -> None:
    """Detecta texto que Excel repararía aunque el XML y el ZIP sean válidos.

    Las cadenas enriquecidas no pasan por el límite de ``check_string`` de
    openpyxl. No recortarlas en silencio: hay información clínica y normativa.
    """
    for ruta in paquete.namelist():
        if ruta == 'xl/sharedStrings.xml':
            raiz = ET.fromstring(paquete.read(ruta))
            unidades = ((f'cadena compartida {i + 1}', nodo)
                        for i, nodo in enumerate(raiz.findall(N + 'si')))
        elif ruta.startswith('xl/worksheets/') and ruta.endswith('.xml'):
            raiz = ET.fromstring(paquete.read(ruta))
            unidades = ((f'celda {celda.get("r", "desconocida")}', celda.find(N + 'is'))
                        for celda in raiz.iter(N + 'c')
                        if celda.get('t') == 'inlineStr')
        else:
            continue
        for direccion, nodo in unidades:
            if nodo is None:
                continue
            texto = ''.join((t.text or '') for t in nodo.iter(N + 't'))
            longitud = _unidades_utf16(texto)
            if longitud > LIMITE_CARACTERES_EXCEL:
                raise ValueError(
                    f'La {direccion} en {ruta} contiene {longitud:,} caracteres '
                    f'UTF-16 (máximo de Excel: {LIMITE_CARACTERES_EXCEL:,}). '
                    'Es necesario acortar ese campo antes de descargar; '
                    'no se eliminó información automáticamente.'
                )


def _contiene_propiedades_rich_incompletas(paquete: zipfile.ZipFile) -> bool:
    """Detecta runs sin texto o propiedades de fuente vacías en XML heredado."""
    for ruta in paquete.namelist():
        if ruta == 'xl/sharedStrings.xml' or (ruta.startswith('xl/worksheets/')
                                               and ruta.endswith('.xml')):
            raiz = ET.fromstring(paquete.read(ruta))
            for run in raiz.iter(N + 'r'):
                fuente = run.find(N + 'rPr')
                if (run.find(N + 't') is None
                        or fuente is not None and not len(fuente) and not fuente.attrib):
                    return True
    return False


def _contiene_fragmentos_vacios(paquete: zipfile.ZipFile) -> bool:
    """Identifica los `<r><t/></r>` inválidos/arriesgados del archivo real."""
    for ruta in paquete.namelist():
        if ruta == 'xl/sharedStrings.xml' or (ruta.startswith('xl/worksheets/')
                                               and ruta.endswith('.xml')):
            try:
                raiz = ET.fromstring(paquete.read(ruta))
            except ET.ParseError as error:
                raise ValueError(f'XML inválido en {ruta}: {error}') from error
            for fragmento in raiz.iter(N + 'r'):
                texto = fragmento.find(N + 't')
                if texto is not None and not (texto.text or ''):
                    return True
    return False


def compactar_fragmentos_enriquecidos(partes):
    """Une saltos, texto sin formato y runs del mismo estilo contiguos.

    Se evita emitir decenas de <r> con sólo un salto de línea en Excel.
    No se altera el texto completo, ni la negrita/cursiva de cada carácter.
    """
    from openpyxl.cell.rich_text import CellRichText, TextBlock

    compactos = []
    for parte in partes:
        es_estilo = isinstance(parte, TextBlock)
        texto = parte.text if es_estilo else str(parte)
        if not texto:
            continue
        if es_estilo and not any(
                valor not in (None, False) for valor in (
                    parte.font.rFont, parte.font.charset, parte.font.family,
                    parte.font.b, parte.font.i, parte.font.strike,
                    parte.font.outline, parte.font.shadow, parte.font.condense,
                    parte.font.extend, parte.font.color, parte.font.sz,
                    parte.font.u, parte.font.vertAlign, parte.font.scheme)):
            # Un <rPr/> vacío puede gatillar reparación de propiedades de
            # cadena en Excel; mantener su texto como fragmento sin formato.
            es_estilo = False
        if es_estilo:
            if (compactos and isinstance(compactos[-1], TextBlock)
                    and compactos[-1].font == parte.font):
                compactos[-1] = TextBlock(parte.font, compactos[-1].text + texto)
            else:
                compactos.append(TextBlock(parte.font, texto))
        elif compactos and isinstance(compactos[-1], str):
            compactos[-1] += texto
        else:
            compactos.append(texto)
    return CellRichText(*compactos)


def _contiene_runs_de_saltos_separados(paquete: zipfile.ZipFile) -> bool:
    """Busca formatos previos que serializaban cada salto como un run OOXML."""
    for ruta in paquete.namelist():
        if ruta == 'xl/sharedStrings.xml' or (ruta.startswith('xl/worksheets/')
                                               and ruta.endswith('.xml')):
            raiz = ET.fromstring(paquete.read(ruta))
            for contenedor in list(raiz.iter(N + 'is')) + list(raiz.iter(N + 'si')):
                runs = contenedor.findall(N + 'r')
                if len(runs) < 2:
                    continue
                anterior_sin_estilo = False
                for run in runs:
                    texto = run.find(N + 't')
                    contenido = (texto.text or '') if texto is not None else ''
                    tiene_estilo = run.find(N + 'rPr') is not None
                    if contenido in ('\n', '\n\n'):
                        return True
                    if anterior_sin_estilo and not tiene_estilo:
                        return True
                    anterior_sin_estilo = not tiene_estilo
    return False




def _contiene_cadenas_enriquecidas(paquete: zipfile.ZipFile) -> bool:
    """Detecta cualquier run enriquecido presente en shared strings u hojas."""
    for ruta in paquete.namelist():
        if ruta == 'xl/sharedStrings.xml' or (ruta.startswith('xl/worksheets/') and ruta.endswith('.xml')):
            raiz = ET.fromstring(paquete.read(ruta))
            if any(True for _ in raiz.iter(N + 'r')):
                return True
    return False


def normalizar_texto_enriquecido_para_excel(datos: bytes) -> bytes:
    """Corrige cadenas enriquecidas mediante openpyxl, sin edición manual XML.

    * No transforma viñetas ni saltos de línea del contenido.
    * Prioriza compatibilidad: cualquier texto enriquecido se convierte
      a texto plano antes de la descarga.
    * Reescribe si detecta cadenas compartidas heredadas, fragmentos vacíos
      o cualquier run enriquecido que Microsoft Excel pudiera reparar.
    """
    try:
        with zipfile.ZipFile(io.BytesIO(datos)) as paquete:
            if paquete.testzip() is not None:
                raise ValueError('El archivo Excel contiene un miembro ZIP dañado.')
            requiere_reescritura = ('xl/sharedStrings.xml' in paquete.namelist()
                                    or _contiene_fragmentos_vacios(paquete)
                                    or _contiene_runs_de_saltos_separados(paquete)
                                    or _contiene_propiedades_rich_incompletas(paquete)
                                    or _contiene_cadenas_enriquecidas(paquete))
            validar_cadenas_ooxml(paquete)
    except (zipfile.BadZipFile, OSError) as error:
        raise ValueError('El archivo generado no es un XLSX válido.') from error
    if not requiere_reescritura:
        return datos

    from openpyxl import load_workbook
    from openpyxl.cell.rich_text import CellRichText

    try:
        libro = load_workbook(io.BytesIO(datos), rich_text=True)
    except Exception as error:
        raise ValueError(f'No se pudo leer el texto enriquecido del Excel: {error}') from error

    try:
        for hoja in libro.worksheets:
            for fila in hoja.iter_rows():
                for celda in fila:
                    valor = celda.value
                    texto = str(valor) if isinstance(valor, CellRichText) else valor
                    if isinstance(texto, str) and _unidades_utf16(texto) > LIMITE_CARACTERES_EXCEL:
                        raise ValueError(
                            f'El texto de {hoja.title}!{celda.coordinate} supera '
                            'los 32.767 caracteres admitidos por Excel.'
                        )
                    if isinstance(valor, CellRichText):
                        celda.value = texto
                        celda.data_type = 's'
        resultado = io.BytesIO()
        libro.save(resultado)
    except ValueError:
        raise
    except Exception as error:
        raise ValueError(f'No se pudo preparar el Excel para descarga: {error}') from error
    finally:
        libro.close()
    limpio = resultado.getvalue()
    with zipfile.ZipFile(io.BytesIO(limpio)) as paquete:
        validar_cadenas_ooxml(paquete)
        if (_contiene_fragmentos_vacios(paquete)
                or _contiene_propiedades_rich_incompletas(paquete)
                or _contiene_cadenas_enriquecidas(paquete)):
            raise ValueError('Quedaron propiedades de cadena incompatibles en el Excel.')
    return limpio
