"""Regresiones: contexto compatible st_aggrid y posición al analizar un PDF."""
import json
import re
import shutil
import subprocess
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
APP = (ROOT / 'app.py').read_text(encoding='utf-8')
BRIDGE = (ROOT / 'interfaz_fluida.py').read_text(encoding='utf-8')


def _recorrer_grid_options_como_st_aggrid(obj):
    """Reproduce el recorrido recursivo que fallaba en shared.walk_gridOptions.

    Un bloque de historial en la forma [[{...}]] acaba entrando al walker
    como lista, cuyos elementos son diccionarios; go[dict] lanza TypeError.
    """
    from collections.abc import Mapping
    for key in obj:
        if isinstance(obj[key], Mapping):
            _recorrer_grid_options_como_st_aggrid(obj[key])
        elif isinstance(obj[key], list):
            for value in obj[key]:
                _recorrer_grid_options_como_st_aggrid(value)


def test_error_original_con_listas_anidadas_y_json_seguro():
    historial = [[{'campo': 'Estabilidad', 'antes': '8 horas', 'despues': ''}]]
    with pytest.raises(TypeError, match='list indices must be integers'):
        _recorrer_grid_options_como_st_aggrid({'context': {'pilaDeshacer': historial}})

    contexto_seguro = {'context': {'pilaDeshacerJson': json.dumps(historial, ensure_ascii=False),
                                  'ultimaEdicionReversible': historial[-1]}}
    _recorrer_grid_options_como_st_aggrid(contexto_seguro)
    assert json.loads(contexto_seguro['context']['pilaDeshacerJson']) == historial
    assert "'pilaDeshacerJson': json.dumps(" in APP
    assert 'JSON.parse(opciones.pilaDeshacerJson' in APP
    assert "'pilaDeshacer': historial_ediciones" not in APP


def test_analizar_sin_segundo_rerun_y_con_texto_reducido():
    assert 'else "Analizar documento", type="primary", key="analyze_pdf"' in APP
    assert 'Analizar documento activo' not in APP
    parte = APP.split('if iniciar_analisis:', 1)[1].split('elif not pdfs:', 1)[0]
    assert 'st.success(f"Se analizaron' in parte
    # El clic dispara ya la ejecución que renderiza abajo la revisión.
    assert not re.search(r'^\s+st\.rerun\(\)\s*$', parte, flags=re.MULTILINE)


def test_el_puente_guarda_la_posicion_y_respeta_los_otros_flujos():
    assert "boton.closest('[class*=\"st-key-analyze_pdf\"]')" in BRIDGE
    assert "registrar('analizar');" in BRIDGE
    assert "dato.tipo === 'analizar'" in BRIDGE
    assert 'boton.getBoundingClientRect().top - dato.referenciaY' in BRIDGE
    assert "registrar('referencias')" in BRIDGE
    assert "registrar('actualizar')" in BRIDGE
    assert 'UI.observador.disconnect()' in BRIDGE


@pytest.mark.skipif(not shutil.which('node'), reason='Node no está instalado')
def test_js_deserializa_historial_sin_perder_ctrlz():
    inicio = APP.index("grid_options['onGridReady'] = JsCode(\"\"\"")
    codigo = APP[APP.index('"""', inicio)+3:APP.index('"""', APP.index('"""', inicio)+3)]
    script = r'''
const assert = require('assert');
const vm = require('vm');
const iniciar = vm.runInNewContext('(' + process.argv[1] + ')', {
 window: {}, document: {addEventListener(){},removeEventListener(){}}
});
const previos = [[{campo:'Estabilidad',antes:'8 horas',despues:'',antes_editado:false}],
                 [{campo:'Tipo de muestra',antes:'Plasma',despues:'',antes_editado:false}]];
const contexto = {pilaDeshacerJson:JSON.stringify(previos),ultimaEdicionReversible:previos[1]};
const api = {};
iniciar({api,context:contexto});
assert.equal(api.__fichasUndoState.pila.length, 2);
assert.equal(api.__fichasUndoState.pila[0][0].antes,'8 horas');
assert.equal(api.__fichasUndoState.pila[1][0].campo,'Tipo de muestra');
const api2 = {};
iniciar({api:api2,context:{pilaDeshacerJson:'{contenido-roto'}});
assert.equal(api2.__fichasUndoState.pila.length, 0);
console.log('Historial JSON seguro y Ctrl+Z preparado');
'''
    p = subprocess.run(['node', '-e', script, codigo], capture_output=True, text=True, timeout=15)
    assert p.returncode == 0, p.stderr
    assert 'Historial JSON seguro' in p.stdout

@pytest.mark.skipif(not shutil.which('node'), reason='Node no está instalado')
def test_js_analisis_restaura_scroll_sin_interferir_en_el_boton():
    fuente = BRIDGE.split('<script>', 1)[1].split('</script>', 1)[0]
    script = r'''
const assert = require('assert');
const vm = require('vm');
const codigo = process.argv[1];
const eventos = {};
let mutacion;
let alerta = null;
let observadorActivo = false;
const area = {scrollTop:250,scrollHeight:4000,clientHeight:650};
const boton = {disabled:false,getBoundingClientRect(){return {top:550-area.scrollTop}},
    closest(sel){return sel.includes('analyze_pdf') ? {} : null;}};
const doc = {
  body:{},
  scrollingElement:area,documentElement:area,
  querySelector(sel){
    if (sel.includes('[data-testid="stMain"]')) return area;
    if (sel.includes('st-key-analyze_pdf')) return boton;
    return null;
  },
  querySelectorAll(sel){return sel.includes('stAlert') && alerta ? [alerta] : [];},
  addEventListener(name,fn){eventos[name]=fn;}
};
const host = {document:doc,requestAnimationFrame(fn){fn();},setTimeout(fn){return 1;},clearTimeout(){}};
const sandbox = {window:{parent:host},MutationObserver:class {
 constructor(cb){mutacion=cb;}
 observe(){observadorActivo=true;}
 disconnect(){observadorActivo=false;}
 },Array};
vm.runInNewContext(codigo,sandbox);
assert.equal(typeof eventos.click,'function');
eventos.click({target:{closest(sel){return sel==='button'?boton:null;}}});
assert.equal(observadorActivo,true);
area.scrollTop=1420;
alerta={textContent:'¡Se procesó correctamente el archivo «lupico.pdf»!'};
mutacion();
assert.equal(area.scrollTop,250,'El botón debe quedar en la misma ubicación visible');
assert.equal(observadorActivo,false,'El observador debe desconectarse después de restaurar');
console.log('Navegación analítica restaurada');
'''
    p = subprocess.run(['node', '-e', script, fuente], capture_output=True, text=True, timeout=15)
    assert p.returncode == 0, p.stderr
    assert 'Navegación analítica restaurada' in p.stdout
