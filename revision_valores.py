"""Separación de resultados actuales, valores antiguos y decisión «Aplicar»."""
from __future__ import annotations
import io
from openpyxl import load_workbook


def valores_anteriores(excel_bytes, nombre_examen, normalizar):
    """Consulta una plantilla sin modificarla ni atribuirle fuentes documentales."""
    if not excel_bytes:
        return {}
    try:
        wb = load_workbook(io.BytesIO(excel_bytes), read_only=True, data_only=True)
        try:
            hoja = next((ws for ws in wb.worksheets
                         if normalizar(ws.title) == normalizar(nombre_examen)), None)
            if hoja is None:
                return {}
            return {normalizar(str(campo)): str(valor) if valor is not None else ""
                    for campo, valor in hoja.iter_rows(min_col=1, max_col=2,
                                                        values_only=True) if campo}
        finally:
            wb.close()
    except (ValueError, OSError, TypeError):
        return {}


def decidir_propuesta(pdf, institucional):
    """El origen seleccionado prevalece como propuesta, sin ocultar una discrepancia."""
    pdf = str(pdf or '').strip()
    institucional = str(institucional or '').strip()
    conflicto = bool(pdf and institucional and ' '.join(pdf.casefold().split()) !=
                      ' '.join(institucional.casefold().split()))
    return institucional or pdf, conflicto
