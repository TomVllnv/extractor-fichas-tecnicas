"""Historial de edición de celdas para la revisión de fichas técnicas.

El historial vive exclusivamente en st.session_state; nunca toca SQLite ni
almacena contenido clínico en almacenamiento persistente del navegador.
Ctrl+Z restaura temporalmente las celdas en memoria de AG Grid y envía una
única marca de acción al servidor para sincronizar la revisión y el Excel.
"""

LIMITE = 20
COLUMNA_COMANDO = "__historial_accion"


def texto_celda(valor):
    """Normaliza valores vacíos del grid igual que la tabla de revisión."""
    if valor is None:
        return ""
    try:
        import pandas as pd
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass
    texto = str(valor).strip()
    return "" if texto.casefold() in ("nan", "none", "null") else texto


def contar_campos_encontrados(datos):
    """Cuenta solo campos con contenido real después de sincronizar AG Grid.

    No cuenta espacios ni marcadores vacíos como 'nan'; conserva el mismo
    conjunto de campos que el indicador previo.
    """
    return sum(bool(texto_celda(valor)) for valor in datos.values())


def contexto_nuevo(historial, identificador):
    """Aísla los cambios por documento/examen; nunca mezcla dos fichas."""
    if historial.get("contexto") != identificador:
        historial.clear()
        historial.update(contexto=identificador, deshacer=[], rehacer=[])
    return historial


def registrar_cambios(historial, valores_previos, editados_previos, filas):
    """Registra una operación de AG Grid, incluso un borrado con Suprimir.

    Varias celdas cambiadas durante el mismo evento se revierten juntas.
    No registra reruns sin cambios, colores o casillas de la columna Aplicar.
    """
    cambios = []
    for fila in filas:
        campo = str(fila["Campo"])
        anterior = texto_celda(valores_previos.get(campo, ""))
        nuevo = texto_celda(fila.get("Información extraída", ""))
        if anterior == nuevo:
            continue
        cambios.append({
            "campo": campo,
            "antes": anterior,
            "despues": nuevo,
            "antes_editado": campo in editados_previos,
            "despues_editado": bool(nuevo) and fila.get("Estado") == "Modificado por usuario",
        })
    if cambios:
        historial.setdefault("deshacer", []).append(cambios)
        del historial["deshacer"][:-LIMITE]
        historial["rehacer"] = []
    return bool(cambios)


def aplicar_accion(historial, datos, editados, accion):
    """Deshace/rehace y devuelve campos a forzar en el Excel consolidado."""
    origen, destino = (("deshacer", "rehacer") if accion == "deshacer"
                       else ("rehacer", "deshacer") if accion == "rehacer"
                       else (None, None))
    if not origen or not historial.get(origen):
        return []
    cambios = historial[origen].pop()
    historial.setdefault(destino, []).append(cambios)
    del historial[destino][:-LIMITE]
    for cambio in cambios:
        campo = cambio["campo"]
        datos[campo] = cambio["antes"] if accion == "deshacer" else cambio["despues"]
        es_editado = cambio["antes_editado"] if accion == "deshacer" else cambio["despues_editado"]
        if es_editado:
            editados.add(campo)
        else:
            editados.discard(campo)
    return [c["campo"] for c in cambios]


def comando_desde_filas(filas):
    """Lee solo órdenes válidas recibidas desde el grid (no datos de usuario)."""
    for fila in filas:
        orden = str(fila.get(COLUMNA_COMANDO, "") or "").split(":", 1)[0]
        if orden in ("deshacer", "rehacer"):
            return orden
    return None


def marca_desde_filas(filas):
    """Identificador de la última orden del grid para consumirla una vez.

    AG Grid puede mantener la columna oculta después de un rerun sin
    reconstruir las celdas; eso no debe deshacer varias acciones por error.
    """
    for fila in filas:
        marca = str(fila.get(COLUMNA_COMANDO, "") or "")
        if marca.split(":", 1)[0] in ("deshacer", "rehacer"):
            return marca
    return None
