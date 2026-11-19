"""Operaciones independientes de la interfaz para una cola de fichas PDF.

Omitir un archivo no crea hojas, no transforma el libro consolidado y no
escribe en SQLite. La confirmación de guardado continúa a cargo de app.py.
"""

import io
from openpyxl import load_workbook


def omitir_documento(estado, indice, total, nombre):
    """Avanza al siguiente sin analizar; el último termina la cola.

    Solo admite colas de al menos dos documentos individuales.
    No modifica result_bytes, bytes_plantilla_inicial ni la base de datos.
    """
    if total <= 1 or not (0 <= indice < total):
        raise ValueError("Solo se pueden omitir documentos de una cola múltiple válida.")
    omitidos = estado.setdefault("cola_pdf_omitidos", [])
    if nombre not in omitidos:
        omitidos.append(nombre)
    estado["cola_pdf_indice"] = indice + 1
    estado["cola_pdf_finalizada_omision"] = indice == total - 1
    return estado["cola_pdf_finalizada_omision"]


def resumen_cola(estado, total):
    procesados = list(estado.get("cola_pdf_procesados") or [])
    omitidos = list(estado.get("cola_pdf_omitidos") or [])
    puede_guardar = bool(procesados and estado.get("result_bytes"))
    return {"total": total, "procesados": procesados,
            "omitidos": omitidos, "puede_guardar": puede_guardar}


def comparar_libros_completos(bytes_anterior, bytes_nuevo):
    """Compara todos los exámenes para la confirmación tras omitir el último.

    La comparación normal solo inspecciona la hoja del examen activo; al
    finalizar una cola sin examen activo es fundamental revisar el libro entero.
    """
    def leer(contenido):
        if not contenido:
            return {}, []
        wb = load_workbook(io.BytesIO(contenido), data_only=True)
        try:
            hojas = list(wb.sheetnames)
            valores = {}
            for ws in wb.worksheets:
                for fila in range(1, ws.max_row + 1):
                    campo = str(ws.cell(fila, 1).value or "").strip()
                    if campo:
                        valor_celda = ws.cell(fila, 2).value
                        valor = str(valor_celda).strip() if valor_celda is not None else ""
                        if valor.casefold() in ("nan", "none", "null"):
                            valor = ""
                        valores[(ws.title, campo)] = valor
            return valores, hojas
        finally:
            wb.close()

    anterior, hojas_ant = leer(bytes_anterior)
    nuevo, hojas_nue = leer(bytes_nuevo)
    diferencias = []
    for hoja, campo in dict.fromkeys((*anterior, *nuevo)):
        antes = anterior.get((hoja, campo), "")
        despues = nuevo.get((hoja, campo), "")
        if antes == despues:
            continue
        tipo = ("Campo nuevo incorporado" if not antes else
                "Campo eliminado / borrado" if not despues else
                "Información modificada")
        diferencias.append({
            "Campo": f"{hoja} · {campo}", "Tipo de cambio": tipo,
            "Valor anterior": antes or "(Vacío)",
            "Valor nuevo": despues or "(Vacío)",
        })
    return {
        "hojas_antes": len(hojas_ant),
        "hojas_despues": len(hojas_nue),
        "campos_completados_antes": sum(bool(v) for v in anterior.values()),
        "campos_completados_despues": sum(bool(v) for v in nuevo.values()),
        "es_nueva_hoja": any(hoja not in hojas_ant for hoja in hojas_nue),
        "diferencias": diferencias,
    }
