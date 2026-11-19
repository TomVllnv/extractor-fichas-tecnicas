"""Regresión: el primer Suprimir habilita inmediatamente Deshacer en revisión."""
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import historial_tabla as ht

CODIGO = (ROOT / "app.py").read_text(encoding="utf-8")


def test_primer_borrado_habilita_el_boton_sin_esperar_una_segunda_edicion():
    """Proteger el orden de render: primero evento de grid, luego botón."""
    indice_placeholder = CODIGO.index('espacio_boton_deshacer = st.empty()')
    indice_grid = CODIGO.index('grid_response = AgGrid(', indice_placeholder)
    indice_registro = CODIGO.index('ht.registrar_cambios(', indice_grid)
    indice_boton = CODIGO.index('with espacio_boton_deshacer.container():', indice_registro)
    indice_excel = CODIGO.index('updated_excel_bytes = aplicar_ediciones_a_excel(', indice_boton)
    assert indice_placeholder < indice_grid < indice_registro < indice_boton < indice_excel
    assert 'disabled=not bool(historial_ediciones.get("deshacer"))' in CODIGO[indice_boton:indice_excel]
    assert 'on_click=_aplicar_historial_tabla, args=("deshacer",)' in CODIGO[indice_boton:indice_excel]


def test_el_primer_suprimir_se_registra_y_se_puede_recuperar_inmediatamente():
    historial = ht.contexto_nuevo({}, ("Hematología y coagulación", "AC. LUPICO", 0))
    datos = {"Estabilidad": "2 horas a temperatura ambiente"}
    assert not historial["deshacer"]  # botón inicialmente deshabilitado
    filas = [{"Campo": "Estabilidad", "Información extraída": "", "Estado": "No encontrado"}]
    assert ht.registrar_cambios(historial, datos, set(), filas)
    assert len(historial["deshacer"]) == 1  # botón habilitado en ESTE rerun
    datos["Estabilidad"] = ""
    assert ht.aplicar_accion(historial, datos, set(), "deshacer") == ["Estabilidad"]
    assert datos["Estabilidad"] == "2 horas a temperatura ambiente"
    assert not historial["deshacer"]


def test_ctrl_z_conoce_el_primer_valor_antes_de_la_respuesta_del_servidor():
    """La pila local funciona sin esperar un rerun de Streamlit."""
    captura = CODIGO[CODIGO.index("grid_options['onCellValueChanged'] = JsCode("):]
    assert "const antes = params.oldValue" in captura
    assert "estadoUndo.pila.push([cambio])" in captura
    assert "params.api.__fichasUndoState" in captura
    teclado = CODIGO[CODIGO.index("grid_options['onGridReady'] = JsCode("):
                     CODIGO.index("grid_options['onCellValueChanged'] = JsCode(")]
    assert "estado.pila.pop()" in teclado
    assert "primerNodo.setDataValue('__historial_accion'" in teclado
    assert "evento.shiftKey" in teclado  # rehacer sigue desactivado
    assert "api.__fichasUndoTecla = deshacer" in teclado
