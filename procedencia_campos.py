"""Procedencia documental por campo, sin atribuir datos previos a fuentes no verificadas.

Las fuentes propuestas se muestran de mayor a menor prioridad; solo se marca
una como utilizada si hay evidencia para el valor actual de la celda.
"""
from __future__ import annotations

import fuentes_documentales as fd

PDF_TECNICO = "PDF importado (ficha técnica)"
PDF_ESCANEADO = "Escáner"
CRS = "Cartera de Exámenes del CRS"
EXAMEN = "Administrador de exámenes"
SISTEMA = "Sistema"
PARAMETROS = "Parámetros operacionales"
REFERENCIAS = "Referencias bibliográficas"
MANUAL = "Completados manualmente por encargado"
HOMOLOGACION = "Homologació de test"


def procedencia_campo(campo, valor, origen_documental="", campos_pdf=None,
                      campos_cartera=None, campos_editados=None, es_escaneado=False,
                      valores_documentales=None, referencias_aplicadas=None, fuentes_elegidas_usuario=None,
                      manual_digital=False, es_derivados=False):
    """Retorna (fuentes posibles en orden de prioridad, fuente comprobada o None).

    No se muestran «Edición manual» ni un Excel preexistente como fuentes
    documentales. Si el valor viene de uno de ellos y no se ha registrado su
    procedencia, queda SIN FUENTE APLICADA, sin atribuciones inventadas.
    """
    campos_pdf = campos_pdf or {}
    campos_cartera = campos_cartera or {}
    campos_editados = campos_editados or set()
    valores_documentales = valores_documentales or {}
    fuente_pdf = PDF_ESCANEADO if es_escaneado else PDF_TECNICO
    fuentes_elegidas_usuario = fuentes_elegidas_usuario or set()
    tiene_valor = valor is not None and str(valor).strip().lower() not in ('', 'nan', 'none', 'null')

    if campo in fd.CAMPOS_MANUAL:
        # Ningún documento ni lectura óptica constituye procedencia de estos campos.
        return [MANUAL], (MANUAL if tiene_valor and campo in campos_editados else None)

    if campo == "Nombre LIS":
        fuentes = [EXAMEN]
    elif campo in ("Laboratorio de procesamiento", "Actualizado por", "Autorizado por"):
        fuentes = [PARAMETROS]
    elif campo == "Fecha actualización":
        fuentes = [SISTEMA]
    elif campo == "Referencias":
        # El listado de referencias se construye desde su propio panel APA;
        # PDF importado no debe figurar como otra fuente de este campo.
        fuentes = [REFERENCIAS]
    elif campo == "Test incluidos":
        fuentes = ([fuente_pdf, HOMOLOGACION] if es_escaneado else
                   [HOMOLOGACION, fuente_pdf])
    elif campo in fd.PRIORIDAD_FUENTES_POR_CAMPO:
        fuentes = list(fd.PRIORIDAD_FUENTES_POR_CAMPO[campo])
        # El documento escaneado puede proponer datos de cualquiera de los
        # campos clínicos, pero solo se marca como fuente utilizada si su
        # contenido coincide exactamente con la propuesta vigente.
        if es_escaneado:
            fuentes.insert(0, fuente_pdf)
        elif (es_derivados or (manual_digital and campo in (
                'Muestra requerida y recipiente'))):
            # Derivados permite fichas de formato heterogéneo con pares
            # etiqueta/valor. En digital la base conserva su prioridad.
            fuentes.append(fuente_pdf)
    else:
        fuentes = []
        if campo in ("Código", "Equipamiento utilizado", "Plazo entrega resultados"):
            fuentes.append(CRS)
        if es_escaneado and campo != "Código":
            fuentes.insert(0, fuente_pdf)
        else:
            fuentes.append(fuente_pdf)

    aplicada = None
    if tiene_valor and campo not in campos_editados:
        if campo == "Nombre LIS":
            aplicada = EXAMEN
        elif campo in ("Laboratorio de procesamiento", "Actualizado por", "Autorizado por"):
            aplicada = PARAMETROS
        elif campo == "Fecha actualización":
            aplicada = SISTEMA
        elif campo == "Referencias" and str(referencias_aplicadas or "").strip() == str(valor).strip():
            aplicada = REFERENCIAS
        elif (es_escaneado and campo != "Código"
              and campo not in fuentes_elegidas_usuario
              and fuente_pdf in fuentes and campo in campos_pdf
              and str(campos_pdf[campo]).strip() == str(valor).strip()):
            aplicada = fuente_pdf
        elif (campo == "Test incluidos" and HOMOLOGACION in fuentes
              and origen_documental == HOMOLOGACION
              and campo in valores_documentales
              and str(valores_documentales[campo]).strip() == str(valor).strip()):
            aplicada = HOMOLOGACION
        elif (fd.identificar_fuente_procedencia(origen_documental) in fuentes
              and campo in valores_documentales
              and str(valores_documentales[campo]).strip() == str(valor).strip()):
            aplicada = fd.identificar_fuente_procedencia(origen_documental)
        elif (CRS in fuentes and campo in campos_cartera
              and str(campos_cartera[campo]).strip() == str(valor).strip()):
            aplicada = CRS
        elif (campo in campos_pdf and fuente_pdf in fuentes
              and str(campos_pdf[campo]).strip() == str(valor).strip()):
            aplicada = fuente_pdf

    # Jamás añadir orígenes no previstos por el campo, ni cambiar la prioridad.
    return list(dict.fromkeys(fuentes)), aplicada
