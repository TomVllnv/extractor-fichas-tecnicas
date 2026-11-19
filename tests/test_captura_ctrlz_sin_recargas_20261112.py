"""Regresión: teclado fiable sin emitir eventos de modelo redundantes."""

from pathlib import Path
import shutil
import subprocess
import pytest

ROOT = Path(__file__).resolve().parents[1]
SOURCE = (ROOT / 'app.py').read_text(encoding='utf-8')


def test_ctrlz_se_captura_en_documento_y_no_requiere_evento_de_celda():
    assert "grid_options['onGridReady'] = JsCode(" in SOURCE
    assert "document.addEventListener('keydown', deshacer, true);" in SOURCE
    assert "document.removeEventListener('keydown', window.__fichasCtrlZHandler, true);" in SOURCE
    assert "evento.stopImmediatePropagation()" in SOURCE
    assert "api.getEditingCells && api.getEditingCells().length" in SOURCE
    assert "evento.target.closest('input, textarea, [contenteditable=\"true\"]')" in SOURCE
    assert "update_mode=GridUpdateMode.VALUE_CHANGED" in SOURCE
    assert "update_mode=GridUpdateMode.MODEL_CHANGED" not in SOURCE


def test_estado_de_fila_no_emite_un_segundo_evento_de_valor():
    manejador = SOURCE.split("grid_options['onCellValueChanged'] = JsCode(\"\"\"", 1)[1].split('""")', 1)[0]
    assert "params.data['Estado'] = 'No encontrado'" in manejador
    assert "params.data['Estado'] = 'Modificado por usuario'" in manejador
    assert "params.node.setDataValue('Estado'" not in manejador


@pytest.mark.skipif(not shutil.which('node'), reason='Node no está instalado en este entorno')
def test_en_javascript_el_primer_suprimir_se_recupera_y_no_genera_reruns_duplicados():
    script = r'''
const fs = require('fs');
const codigo = fs.readFileSync(process.argv[1], 'utf8');
function extraer(nombre) {
    const inicio = codigo.indexOf(`grid_options['${nombre}'] = JsCode("""`);
    if (inicio < 0) throw new Error('No existe ' + nombre);
    const js = codigo.slice(inicio).split('JsCode("""', 2)[1].split('""")', 1)[0].trim();
    return eval('(' + js + ')');
}
const inicio = extraer('onGridReady');
const edicion = extraer('onCellValueChanged');
let escuchas = [], eventosServidor = 0, refrescos = 0, interceptados = 0;
global.window = {};
global.document = {
    addEventListener: (nombre, fn, captura) => { if(nombre==='keydown' && captura) escuchas.push(fn); },
    removeEventListener: (nombre, fn) => { escuchas = escuchas.filter(x => x !== fn); }
};
const contexto = {ultimaEdicionReversible:[]};
const fila = {
    data: {Campo:'Estabilidad', 'Información extraída':'', Estado:'No encontrado'},
    setDataValue(nombre, valor) { this.data[nombre]=valor; eventosServidor++; }
};
const api = {
    getEditingCells: () => [], getGridOption: nombre => nombre==='context' ? contexto : null,
    forEachNode: fn => fn(fila), refreshCells: () => refrescos++
};
const p = {api,context:contexto};
inicio(p);
if (escuchas.length !== 1) throw Error('No escucha el atajo');
edicion({column:{colDef:{field:'Información extraída'}}, oldValue:'8 horas', newValue:'', data:fila.data, node:fila,context:contexto,api});
if (contexto.ultimaEdicionReversible.length !== 1 || eventosServidor !== 0) {
    throw Error('El primer borrado no quedó capturado o el estado disparó otra actualización');
}
const evento = (elemento, shiftKey=false) => ({ctrlKey:true,metaKey:false,altKey:false,shiftKey,
    key:'z',target:{closest:()=>elemento},preventDefault(){interceptados++},
    stopPropagation(){},stopImmediatePropagation(){}});
escuchas[0](evento(null));
if (fila.data['Información extraída'] !== '8 horas' || eventosServidor !== 1 || interceptados !== 1) {
    throw Error('No restauró correctamente la primera eliminación');
}
escuchas[0](evento(null));
if (eventosServidor !== 1) throw Error('Procesó dos veces la misma operación');
contexto.ultimaEdicionReversible=[{campo:'Estabilidad',antes:'8 horas', antes_editado:false}];
escuchas[0](evento({tagName:'TEXTAREA'}));
if (eventosServidor !== 1) throw Error('Interceptó el deshacer del editor de texto');
escuchas[0](evento(null,true));
if (eventosServidor !== 1) throw Error('Rehacer no debe estar activo');
inicio(p);
if (escuchas.length !== 1) throw Error('Duplicó listeners al reconstruir el grid');
console.log('OK');
'''
    ejecucion = subprocess.run(
        ['node', '-e', script, str(ROOT / 'app.py')], capture_output=True,
        text=True, check=False, timeout=10,
    )
    assert ejecucion.returncode == 0, ejecucion.stderr
    assert 'OK' in ejecucion.stdout
