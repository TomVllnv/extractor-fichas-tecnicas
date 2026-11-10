"""Interacciones ligeras para la ficha técnica de Streamlit.

El puente conserva el desplazamiento al aplicar Referencias o actualizar una
plantilla. No lee ni envía PDF, datos de pacientes, credenciales ni SQLite.
Solo mantiene observadores/temporizadores mientras espera una acción concreta.
"""
from __future__ import annotations

SOPORTE_NAVEGACION = r"""<!doctype html><html><body style='margin:0'>
<script>
(() => {
  'use strict';
  let host;
  try { host = window.parent; void host.document.body; } catch (_) { return; }
  const doc = host.document;
  const ID = '__testeoInterfazFluidaV3';
  if (host[ID]) {
    // El iframe puede recrearse durante un rerun, pero el oyente principal
    // permanece registrado una sola vez en la ventana de Streamlit.
    host[ID].recuperar();
    return;
  }

  const UI = { pendiente: null, observador: null, reloj: null,
               fotograma: false, cambioArea: null };
  host[ID] = UI;

  const encontrarScroll = () => {
    const candidatos = [doc.querySelector('[data-testid="stMain"]'),
      doc.querySelector('section.main'),
      doc.querySelector('[data-testid="stAppViewContainer"]')];
    for (const elemento of candidatos) {
      if (elemento && elemento.scrollHeight > elemento.clientHeight + 20) return elemento;
    }
    return doc.scrollingElement || doc.documentElement;
  };
  const marcadorReferencias = () => doc.querySelector(
    '[class*="st-key-contenedor_referencias"] details > summary'
  );
  const cerrarEspera = () => {
    if (UI.observador) UI.observador.disconnect();
    if (UI.reloj) host.clearTimeout(UI.reloj);
    UI.observador = null;
    UI.reloj = null;
    UI.pendiente = null;
  };
  const restaurar = (respaldo = false) => {
    const dato = UI.pendiente;
    if (!dato) return;
    const scroll = encontrarScroll();
    if (dato.tipo === 'referencias') {
      const nuevo = marcadorReferencias();
      // Restaurar una vez cuando Streamlit haya montado el NUEVO panel.
      // No corregir la posición continuamente mientras el usuario navega.
      if (!respaldo && (!nuevo || nuevo === dato.anterior)) return;
      if (nuevo && dato.referenciaY != null) {
        const diferencia = nuevo.getBoundingClientRect().top - dato.referenciaY;
        if (Math.abs(diferencia) > 2) scroll.scrollTop += diferencia;
      } else if (Math.abs(scroll.scrollTop - dato.posicion) > 2) {
        scroll.scrollTop = dato.posicion;
      }
    } else {
      // El flujo de actualización abre/cierra un diálogo al terminar el rerun.
      // Esperar al cambio del diálogo, con un único respaldo temporal.
      const dialogo = !!doc.querySelector('[role="dialog"]');
      if (!respaldo && dialogo === dato.dialogoAnterior) return;
      if (Math.abs(scroll.scrollTop - dato.posicion) > 2) scroll.scrollTop = dato.posicion;
    }
    cerrarEspera();
  };
  UI.recuperar = () => restaurar(false);

  const registrar = tipo => {
    cerrarEspera();
    const scroll = encontrarScroll();
    const marcador = tipo === 'referencias' ? marcadorReferencias() : null;
    UI.pendiente = {
      tipo, posicion: scroll.scrollTop || 0,
      anterior: marcador,
      referenciaY: marcador ? marcador.getBoundingClientRect().top : null,
      dialogoAnterior: !!doc.querySelector('[role="dialog"]')
    };
    UI.observador = new MutationObserver(() => {
      if (UI.fotograma) return;
      UI.fotograma = true;
      host.requestAnimationFrame(() => {
        UI.fotograma = false;
        restaurar(false);
      });
    });
    UI.observador.observe(doc.body, {childList: true, subtree: true});
    UI.reloj = host.setTimeout(() => restaurar(true), 2500);
  };


  const ocultarBienvenidaDuranteCambio = () => {
    // Streamlit puede conservar unos instantes la vista anterior mientras
    // vuelve a procesar todo el sidebar. Solo ocultar la bienvenida antigua,
    // no las fichas ni otros cambios de área ya iniciados.
    const principal = doc.querySelector('[data-testid="stMain"]') ||
      doc.querySelector('section.main');
    if (!principal || !principal.textContent.includes(
        '¡Bienvenido/a a la Herramienta de Extracción de Datos!')) return;
    if (UI.cambioArea) UI.cambioArea();
    principal.style.visibility = 'hidden';
    let observador = null;
    let respaldo = null;
    const mostrar = () => {
      if (observador) observador.disconnect();
      if (respaldo) host.clearTimeout(respaldo);
      principal.style.removeProperty('visibility');
      UI.cambioArea = null;
    };
    UI.cambioArea = mostrar;
    observador = new MutationObserver(() => {
      // Mostrar solo cuando la nueva vista por área ya tiene sus pestañas.
      if (principal.querySelector('[role="tablist"]') &&
          !principal.textContent.includes('¡Bienvenido/a a la Herramienta de Extracción de Datos!')) mostrar();
    });
    observador.observe(principal, {childList: true, subtree: true});
    respaldo = host.setTimeout(mostrar, 6000);
  };

  const alPulsar = evento => {
    const boton = evento.target && evento.target.closest && evento.target.closest('button');
    if (!boton) return;
    if (boton.closest('[class*="st-key-abrir_area_"]')) {
      ocultarBienvenidaDuranteCambio();
      return;
    }
    const aplicar = boton.closest('[class*="st-key-btn_aplicar_referencias_apa"]');
    if (aplicar && !boton.disabled) {
      // Registrar la posición, pero dejar que Streamlit procese el clic.
      // No cerrar el expander, ni interceptar o repetir la pulsación.
      registrar('referencias');
      return;
    }
    const actualizar = boton.closest('[class*="st-key-btn_trigger_save_template"]');
    const aceptar = boton.closest('[class*="st-key-btn_confirmar_guardado_ficha"]');
    const cancelar = boton.closest('[class*="st-key-btn_cancelar_guardado_ficha"]');
    if (actualizar || aceptar || cancelar) registrar('actualizar');
  };

  doc.addEventListener('click', alPulsar, true);
  // Al desplazarse manualmente se respeta de inmediato al usuario.
  doc.addEventListener('wheel', cerrarEspera, {passive: true, capture: true});
  doc.addEventListener('touchmove', cerrarEspera, {passive: true, capture: true});
})();
</script></body></html>"""


def activar():
    """Habilita navegación sin realizar operaciones en los documentos o SQLite."""
    import streamlit as st
    st.components.v1.html(SOPORTE_NAVEGACION, height=0, scrolling=False)
