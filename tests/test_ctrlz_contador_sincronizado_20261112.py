"""Pruebas de Ctrl+Z cliente y contador luego de editar en AG Grid."""
import json
import shutil
import subprocess
from pathlib import Path

import historial_tabla as ht

ROOT = Path(__file__).resolve().parents[1]
CODIGO = (ROOT / "app.py").read_text(encoding="utf-8")


def _js_evento(nombre):
    inicio = CODIGO.index(f"grid_options['{nombre}'] = JsCode(\"\"\"")
    inicio = CODIGO.index('"""', inicio) + 3
    fin = CODIGO.index('"""', inicio)
    return CODIGO[inicio:fin].strip()


def test_contador_se_renderiza_despues_de_recibir_los_cambios():
    placeholder = CODIGO.index('espacio_campos_encontrados = st.empty()')
    grid = CODIGO.index('grid_response = AgGrid(', placeholder)
    asignacion = CODIGO.index('st.session_state["data_extracted"][field] = val_editado', grid)
    contador = CODIGO.index('espacio_campos_encontrados.metric(', asignacion)
    assert placeholder < grid < asignacion < contador
    # Solo las propuestas actuales marcadas cuentan en el indicador.
    assert 'ht.contar_campos_encontrados({' in CODIGO
    assert 'if seleccion_actual.get(campo, True)' in CODIGO
    assert ht.contar_campos_encontrados({"Estabilidad": "8 horas", "Equipo": "X"}) == 2
    assert ht.contar_campos_encontrados({"Estabilidad": "", "Equipo": "X"}) == 1
    assert ht.contar_campos_encontrados({"Estabilidad": "  ", "Equipo": "nan"}) == 0


def test_ctrlz_js_primera_edicion_y_dos_ediciones_consecutivas():
    """Simula el navegador: pulsaciones, eventos y datos AG Grid sin Streamlit."""
    if not shutil.which("node"):
        import pytest
        pytest.skip("Node no está instalado para ejecutar JavaScript")
    scripts = [_js_evento(nombre) for nombre in
               ('onGridReady', 'onCellKeyDown', 'onCellValueChanged')]
    codigo_node = r'''
const assert = require('assert');
const vm = require('vm');
const fuente = JSON.parse(process.argv[1]);
const listeners = {};
const window = {};
const document = {
  addEventListener: (name, handler) => { listeners[name] = handler; },
  removeEventListener: (name, handler) => {
    if (listeners[name] === handler) delete listeners[name];
  },
};
const sandbox = {window, document, Map, Array, Date, Math, String};
const handlers = fuente.map(js => vm.runInNewContext('(' + js + ')', sandbox));
const filas = ['Estabilidad', 'Tipo de muestra'].map((campo, i) => ({
  data: {'Campo': campo, 'Información extraída': i ? 'Plasma' : '8 horas',
         'Estado': 'Encontrado', '__historial_accion': ''},
  setDataValue(k, v) { this.data[k] = v; if (k === '__historial_accion') ordenes.push(v); }
}));
let ordenes = [];
let refreshes = 0;
const api = {
  forEachNode(fn) { filas.forEach(fn); },
  refreshCells() { refreshes++; },
  getEditingCells() { return []; },
};
const context = {pilaDeshacer: [], ultimaEdicionReversible: []};
handlers[0]({api, context});
function editar(node, nuevo) {
  const anterior = node.data['Información extraída'];
  node.data['Información extraída'] = nuevo;
  handlers[2]({api, context, column: {colDef: {field: 'Información extraída'}},
    data: node.data, node, oldValue: anterior, newValue: nuevo});
}
function undo() {
  const evento = {ctrlKey: true, key: 'z',
    preventDefault() {this.prevented = true;},
    stopPropagation() {}, stopImmediatePropagation() {}};
  listeners.keydown(evento);
  handlers[1]({api, event: evento, node: filas[0]});
  return evento;
}
// El primer Suprimir se puede deshacer en la misma sesión, sin respuesta backend.
editar(filas[0], '');
assert.equal(api.__fichasUndoState.pila.length, 1);
assert.equal(undo().prevented, true);
assert.equal(filas[0].data['Información extraída'], '8 horas');
assert.equal(ordenes.length, 1, 'la misma pulsación no emite dos órdenes');
// Edición sucesiva de dos campos y dos deshacer en orden inverso.
editar(filas[0], 'Valor diferente');
editar(filas[1], '');
assert.equal(api.__fichasUndoState.pila.length, 2);
assert.equal(undo().prevented, true);
assert.equal(filas[1].data['Información extraída'], 'Plasma');
assert.equal(ordenes.length, 2);
assert.equal(undo().prevented, true);
assert.equal(filas[0].data['Información extraída'], '8 horas');
assert.equal(ordenes.length, 3);
// Sin historial no se intercepta el navegador ni se emite evento.
assert.equal(undo().prevented, undefined);
assert.equal(ordenes.length, 3);
// Si está el editor activo, Ctrl+Z queda reservado para edición de texto.
editar(filas[0], 'Temporal');
api.getEditingCells = () => [{rowIndex: 0, colKey: 'Información extraída'}];
assert.equal(undo().prevented, undefined);
assert.equal(filas[0].data['Información extraída'], 'Temporal');
assert.equal(ordenes.length, 3);
console.log('OK: primer borrado, doble Ctrl+Z, sin comandos duplicados y editor nativo');
'''
    resultado = subprocess.run(["node", "-e", codigo_node, json.dumps(scripts)],
                               capture_output=True, text=True, timeout=15)
    assert resultado.returncode == 0, resultado.stderr
    assert 'OK:' in resultado.stdout
