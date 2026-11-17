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
    """Repara «e-» confundido con una viñeta al inicio de renglón.

    No se cambia «e-» dentro de palabras, códigos, correos o frases. Una
    secuencia aislada debe preceder a otro renglón con contenido suficiente.
    """
    lineas = str(texto or '').splitlines(keepends=True)
    resultado = []
    for i, linea in enumerate(lineas):
        coincidencia = re.match(r'^([ \t]*)[eE][ \t]*-([ \t]*)(.*?)(\r?\n)?$', linea)
        if not coincidencia:
            resultado.append(linea)
            continue
        prefijo, espacio_tras_guion, resto, salto = coincidencia.groups()
        frase = resto.strip()
        if frase and not espacio_tras_guion:
            # e-123 o e-mail sin separación es una cadena, no una lista.
            resultado.append(linea)
            continue
        if not frase:
            siguiente = next((v.strip() for v in lineas[i + 1:] if v.strip()), '')
            es_vineta = len(siguiente.split()) >= 3 and len(siguiente) >= 15
        else:
            # «e-mail», «e-coli», «e-commerce» no son viñetas.
            es_vineta = (len(frase.split()) >= 3 and len(frase) >= 15
                         and not re.match(r'(?i)^(?:mail|coli|commerce|learning)\b', frase))
        if es_vineta:
            resultado.append(prefijo + '• ' + frase + (salto or ''))
        else:
            resultado.append(linea)
    return ''.join(resultado)


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
