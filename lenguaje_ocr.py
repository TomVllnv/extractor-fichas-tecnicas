"""Correcciones lingüísticas limitadas a errores inequívocos de lectura óptica.

No es una normalización de datos clínicos: códigos, magnitudes, modelos,
fechas y palabras desconocidas se mantienen intactos. Sin servicios externos.
"""
from __future__ import annotations

import re
import unicodedata

# Léxico de términos habituales que se espera encontrar en texto de laboratorio.
# Solo se corrigen variantes cuyo reemplazo es inequívoco dentro de este léxico.
_TERMINOS = """
Bioquímica Hematología Inmunología Microbiología Laboratorio Laboratorios
Recepción Procesamiento Centrifugación Preparación Rotulación Identificación
Muestra Muestras Paciente Pacientes Análisis Analítica Analítico Analíticos
Clínica Clínico Clínicos Clínicas Procedimiento Procedimientos
Examen Exámenes Resultado Resultados Valor Valores Referencia Referencias
Equipo Equipos Equipamiento Método Métodos Sistema Sistemas
Extracción Conservación Estabilidad Transporte Envío Rechazo
Detección Límite Límites Crítico Críticos Indicaciones
Temperatura Temperaturas Suero Plasma Sangre Orina Hemólisis
Hemoglobina Hemograma Coagulación Inmunoensayo Quimioluminiscencia
Turbidimetría Fotometría Espectrofotometría Enzimático Enzimática
Calibración Concentración Sensibilidad Especificidad
Médico Médica Diagnóstico Diagnóstica Diagnósticos
Información Registro Registros Revisión Versión Emisión
Publicación Elaboración Aprobación Vigencia Septiembre
Enero Febrero Marzo Abril Mayo Junio Julio Agosto Octubre Noviembre Diciembre
""".split()


def _base(palabra):
    descompuesta = unicodedata.normalize("NFKD", palabra.casefold())
    return "".join(c for c in descompuesta if not unicodedata.combining(c))


_LEXICO = {_base(termino): termino.casefold() for termino in _TERMINOS}
# Ejemplos con dos interpretaciones no se autocorrigen: la lista queda
# restringida a términos inequívocos y el usuario conserva el original.
_PALABRA = re.compile(r"(?<![\w])([A-Za-zÁÉÍÓÚÜÑáéíóúüñ0-9]+)(?![\w])")


def _forma_habitual(original: str, canonica: str) -> str:
    if original.isupper():
        return canonica.upper()
    if original[:1].isupper():
        return canonica.capitalize()
    return canonica


def _normalizar_vinetas_opticas(texto: str) -> str:
    """Recupera viñetas confundidas con «e-» o «+» sin alterar datos.

    Se exige que el signo esté separado del texto y encabece una instrucción.
    Un número después de «+» podría indicar una magnitud o un resultado, de
    modo que no se transforma automáticamente. Los términos cortos ambiguos
    se conservan salvo que pertenezcan a una lista ya reconocible.
    """
    lineas = str(texto or '').splitlines(keepends=True)
    salida = []
    patron_marcador = re.compile(
        r'^([ \t]*)([eE][ \t]*-|\+)([ \t]*)(.*?)(\r?\n)?$'
    )
    patron_lista = re.compile(r'^[ \t]*(?:[•●▪◦∙]|[+][ \t]+|[eE][ \t]*-[ \t]+)')

    def otra_vineta_valida(linea: str) -> bool:
        m = patron_lista.match(linea)
        if not m:
            return False
        texto_item = linea[m.end():].strip()
        if re.match(r'^[+\-−]?[\d.,<>=≤≥]', texto_item):
            return False
        return len(texto_item.split()) >= 2 and len(texto_item) >= 12

    for indice, linea in enumerate(lineas):
        m = patron_marcador.match(linea)
        if not m:
            salida.append(linea)
            continue
        prefijo, marcador, espacio, resto, salto = m.groups()
        frase = resto.strip()
        if frase and not espacio:
            # No modificar identificadores o cadenas unidas al signo.
            salida.append(linea)
            continue
        proxima = next((x.strip() for x in lineas[indice + 1:] if x.strip()), '')
        candidata = frase or proxima
        if not candidata or re.match(r'^[+\-−]?[\d.,<>=≤≥]', candidata):
            salida.append(linea)
            continue
        if re.match(r'(?i)^(?:mail|coli|commerce|learning)\b', candidata):
            salida.append(linea)
            continue

        coherente = len(candidata.split()) >= 2 and len(candidata) >= 12
        adyacentes = (
            (indice > 0 and otra_vineta_valida(lineas[indice - 1]))
            or (indice + 1 < len(lineas) and otra_vineta_valida(lineas[indice + 1]))
        )
        if coherente or (adyacentes and len(candidata) >= 5):
            salida.append(prefijo + '• ' + frase + (salto or ''))
        else:
            salida.append(linea)
    return ''.join(salida)


def corregir_texto_lectura_optica(texto: str) -> str:
    """Corrige confusiones 0/o y acentos de términos reconocidos.

    No utiliza similitud difusa: evitar cambiar una palabra correcta por
    otra clínica de significado distinto es más importante que completar
    todas las correcciones posibles.
    """
    def corregir(match):
        palabra = match.group(1)
        # Cualquier cifra no nula indica probable identificador, resultado,
        # unidad o modelo. Tampoco tocar secuencias totalmente numéricas.
        if re.search(r"[1-9]", palabra) or not re.search(r"[a-záéíóúüñ]", palabra, re.I):
            return palabra
        # 0 por O: solo se permite si el resultado está en el léxico.
        candidato = palabra.replace("0", "o")
        canonica = _LEXICO.get(_base(candidato))
        if not canonica:
            return palabra
        # Evitar alterar palabras correctas únicamente por capitalización.
        if palabra.casefold() == canonica:
            return palabra
        return _forma_habitual(palabra, canonica)

    return _normalizar_vinetas_opticas(_PALABRA.sub(corregir, str(texto or "")))
