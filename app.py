import io, os, re, html, zipfile, tempfile, subprocess, shutil, csv, json, hashlib, copy, base64, sqlite3, unicodedata, time, xml.etree.ElementTree as ET
from difflib import SequenceMatcher
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import openpyxl
import fitz  # PyMuPDF
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, JsCode
from streamlit_sortables import sort_items
import fuentes_documentales as fd
import homologacion_tests as htest
import gestion_compartida as gp
import referencias_apa as apa
import compatibilidad_excel as ce
import registro_usuarios_login as rul
import interfaz_fluida as iflu
import coincidencias_ui as cui
import historial_tabla as ht
import cola_fichas as cf
import manuales_procedimiento as mp
import procedencia_campos as pc
import ocr_documentos_escaneados as oe
import equipamientos as eq
import revision_valores as rv
import deteccion_pdf as dp
import derivados_estructurado as de
from uuid import uuid4

APP_TITLE = 'Herramienta de extracción de datos para laboratorio'
MAX_PDFS_PERMITIDOS = 5
DB_PATH = gp.ruta_base()
gp.inicializar(DB_PATH)

AREA_CONFIG = {
    "Microbiología": {"template": "plantilla_microbiologia.xlsx"},
    "Hematología y coagulación": {"template": "plantilla_hematologia.xlsx"},
    "Bioquímica": {"template": "plantilla_bioquimica.xlsx"},
    "Hormonas e inmunidad": {"template": None},
    "Derivados": {"template": None},
}


@st.cache_data(show_spinner=False, max_entries=24)
def previsualizar_tipo_pdf(contenido: bytes):
    """Identifica PDF escaneado/textual/mixto sin ejecutar el motor OCR."""
    return dp.inspeccionar_pdf(contenido)


def formatear_fecha_fuente(fecha_iso):
    """Presenta la fecha ISO guardada en SQLite como DD/MM/AAAA HH:MM.

    No cambia el valor almacenado (ISO se mantiene para ordenar/consultar).
    """
    if not fecha_iso:
        return "Sin registro"
    try:
        fecha = datetime.fromisoformat(fecha_iso)
        if fecha.tzinfo is not None:
            fecha = fecha.astimezone(ZoneInfo("America/Santiago"))
        return fecha.strftime("%d/%m/%Y %H:%M")
    except (TypeError, ValueError):
        return "Fecha no disponible"


def fecha_hoy_nombre_archivo():
    """DD-MM-AAAA para nombres válidos en Windows; UI usa DD/MM/AAAA."""
    return datetime.now(ZoneInfo("America/Santiago")).strftime("%d-%m-%Y")


def nombre_ficha_tecnica(area, fecha_extraccion=None):
    """Nombre uniforme para las fichas: área + fecha de extracción DD-MM-AAAA."""
    area_segura = re.sub(r'[<>:"/\\|?*]+', '_', str(area or "General")).strip()
    area_segura = re.sub(r"\s+", "_", area_segura).strip("_ .") or "General"
    fecha = fecha_extraccion or fecha_hoy_nombre_archivo()
    return f"Ficha_Técnica_{area_segura}_Actualizada_{fecha}.xlsx"


def nombre_descarga_excel_generado(nombre, area=None, fecha_registro=None):
    """Normaliza descargas históricas sin renombrar los archivos guardados.

    Para fichas generadas se conserva la fecha original de extracción, incluso
    cuando se descargan días después desde SQLite. En documentos importados,
    la mejor fecha disponible es la registrada al guardar la plantilla.
    """
    original = str(nombre or "")
    patron_fecha = r"(20\d{2}-\d{2}-\d{2}|\d{2}-\d{2}-20\d{2})"
    auto = re.fullmatch(
        rf"Ficha_T[eé]cnica_Actualizada_(.+)_({patron_fecha})\.xlsx",
        original, re.IGNORECASE)
    moderno = re.fullmatch(
        rf"Ficha_T[eé]cnica_(.+)_Actualizada_({patron_fecha})\.xlsx",
        original, re.IGNORECASE)
    hallado = auto or moderno
    if hallado:
        area_archivo = hallado.group(1)
        fecha_archivo = hallado.group(2)
        if re.fullmatch(r"20\d{2}-\d{2}-\d{2}", fecha_archivo):
            fecha_archivo = datetime.strptime(fecha_archivo, "%Y-%m-%d").strftime("%d-%m-%Y")
        return nombre_ficha_tecnica(area or area_archivo, fecha_archivo)

    if area is None:
        # Archivo de usuario sin formato generado: respetar el nombre original.
        return nombre

    fecha = None
    if fecha_registro and str(fecha_registro) != "Sin registro":
        marca = re.search(r"(\d{2}/\d{2}/\d{4})", str(fecha_registro))
        if marca:
            fecha = marca.group(1).replace("/", "-")
        else:
            marca_iso = re.search(r"(20\d{2}-\d{2}-\d{2})", str(fecha_registro))
            if marca_iso:
                fecha = datetime.strptime(marca_iso.group(1), "%Y-%m-%d").strftime("%d-%m-%Y")
    return nombre_ficha_tecnica(area, fecha)


def clave_orden_alfabetico(nombre):
    """Ordena exámenes en español sin distinguir tildes ni mayúsculas.

    Solo transforma la clave de comparación: los nombres originales de SQLite,
    Nombre LIS y las fichas técnicas permanecen intactos.
    """
    base = unicodedata.normalize("NFD", str(nombre).strip())
    sin_tildes = "".join(caracter for caracter in base if not unicodedata.combining(caracter))
    return (sin_tildes.casefold(), str(nombre).casefold())


def mostrar_coincidencia_seleccionable(tipo, identificador, nombre, origen, puntaje,
                                     accion, valor_accion):
    """Muestra una sola tarjeta clicable con color según similitud."""
    clave_nombre = cui.claves_coincidencia(tipo, identificador)
    st.html(cui.estilos_coincidencia(clave_nombre, puntaje))
    etiqueta = (
        f"{str(nombre).strip()} — {str(origen).strip()} · "
        f"Similitud: {cui.porcentaje_similitud(puntaje)} %"
    )
    st.button(etiqueta, key=clave_nombre, on_click=accion,
              args=(valor_accion,), use_container_width=True)


def _total_paginas_pdf_cacheadas(pdf_bytes: bytes) -> int:
    """Cuenta las páginas por PDF y guarda el resultado sólo en esta sesión."""
    clave = hashlib.sha256(pdf_bytes).hexdigest()
    cache = st.session_state.setdefault("pdf_preview_cuenta_cache", {})
    if clave not in cache:
        with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
            cache[clave] = len(doc)
        if len(cache) > 12:
            cache.pop(next(iter(cache)))
    return cache[clave]


def _imagen_pagina_pdf_cacheada(pdf_bytes: bytes, pagina: int, dpi: int):
    """Rasteriza sólo cuando hace falta; las imágenes permanecen en la sesión."""
    clave = (hashlib.sha256(pdf_bytes).hexdigest(), pagina, dpi)
    cache = st.session_state.setdefault("pdf_preview_imagen_cache", {})
    if clave not in cache:
        with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
            pix = doc[pagina].get_pixmap(dpi=dpi, alpha=False)
            cache[clave] = (base64.b64encode(pix.tobytes("png")).decode("ascii"),
                            pix.width, pix.height)
        if len(cache) > 16:
            cache.pop(next(iter(cache)))
    return cache[clave]


@st.fragment
def _visor_pdf_rapido(pdf_activo, idx_actual):
    """Visor aislado por documento de la cola, sin selector de PDF conjunto."""
    try:
        pdf_vista = pdf_activo

        pdf_bytes_vista = pdf_vista.getvalue()
        num_paginas = _total_paginas_pdf_cacheadas(pdf_bytes_vista)
        st.caption(f"Visualizando: **{pdf_vista.name}** | Páginas: **{num_paginas}**")

        zoom_level = 120

        col_prev1, col_prev2 = st.columns([2, 1])

        # Cada documento conserva su propia página seleccionada.
        # El índice está ligado a la posición en la lista de archivos.
        indice_pagina = idx_actual
        # La página se conserva por documento; las flechas y el
        # slider usan exactamente la misma clave de estado.
        clave_pagina = (f"pdf_page_slider_cola_"
                        f"{indice_pagina}_{hashlib.sha256(pdf_vista.name.encode('utf-8')).hexdigest()[:12]}")
        with col_prev1:
            pag_seleccionada = st.slider(
                "Página a visualizar", 1, num_paginas, 1,
                key=clave_pagina) if num_paginas > 1 else 1

        img_b64, ancho_pixeles, alto_pixeles = _imagen_pagina_pdf_cacheada(
            pdf_bytes_vista, pag_seleccionada - 1, 130)

        ancho_visual = int(750 * (zoom_level / 100))
        alto_real_imagen = int(ancho_visual * (alto_pixeles / ancho_pixeles))
        alto_componente = alto_real_imagen + 105

        html_lupa = f"""
        <!doctype html>
        <html lang="es"><head><meta charset="utf-8">
        <style>
            * {{ box-sizing: border-box; }}
            body {{ margin:0; padding:0; background:transparent; color:#66788a;
                   font-family:system-ui,-apple-system,Segoe UI,sans-serif; }}
            .pdf-visor {{ position:relative; overflow:hidden; display:block;
                margin:0 auto; width:min(100%,{ancho_visual}px);
                border:1px solid #a8b7c8; border-radius:9px;
                user-select:none; touch-action:none; cursor:zoom-in;
                background:#f6f8fb; }}
            .pdf-visor.activo {{ cursor:grab; border-color:#4b91c7;
                outline:2px solid rgba(74,144,192,.27); outline-offset:-2px; }}
            .pdf-visor.arrastrando {{ cursor:grabbing; }}
            .pdf-visor img {{ width:100%; height:auto; display:block;
                transform-origin:50% 50%; will-change:transform; pointer-events:none;
                user-select:none; }}
            .pdf-ayuda {{ text-align:center; font-size:12px; margin:7px 0 3px;
                color:#718398; }}
            .pdf-acciones {{ display:flex; justify-content:center; align-items:center;
                gap:8px; margin-top:6px; }}
            .pdf-acciones button {{ border:1px solid #a3b7c9; color:#234d73;
                background:#edf4fa; border-radius:6px; min-width:32px; min-height:28px;
                cursor:pointer; font:600 13px system-ui; }}
            .pdf-acciones button:hover {{ background:white; }}
            .pdf-acciones span {{ font-size:12px; color:#586e81; min-width:38px; text-align:center; }}
        </style></head><body>
            <div id="pdf-visor" class="pdf-visor" tabindex="0"
                 role="region" aria-label="Vista previa del PDF ampliable">
                <img id="pdf-imagen" src="data:image/png;base64,{img_b64}"
                     alt="Página del documento PDF" draggable="false">
            </div>
            <div class="pdf-acciones">
                <button id="pdf-menos" type="button" aria-label="Alejar">−</button>
                <span id="pdf-nivel">100 %</span>
                <button id="pdf-mas" type="button" aria-label="Acercar">+</button>
                <button id="pdf-reiniciar" type="button">Restablecer</button>
            </div>
            <p class="pdf-ayuda" id="pdf-ayuda">
              Clic para enfocar · rueda para ampliar · arrastrar para mover · doble clic para restablecer
            </p>
        <script>
        (() => {{
            'use strict';
            const marco = document.getElementById('pdf-visor');
            const imagen = document.getElementById('pdf-imagen');
            const nivel = document.getElementById('pdf-nivel');
            let escala=1, x=0, y=0, enfocado=false, arrastrando=false;
            let inicioX=0, inicioY=0;
            const maxZoom=4;
            const limitar = (valor,min,max) => Math.max(min,Math.min(max,valor));
            function ajustar() {{
                const excesoX = marco.clientWidth * (escala - 1) / 2;
                const excesoY = marco.clientHeight * (escala - 1) / 2;
                x=limitar(x,-excesoX,excesoX);
                y=limitar(y,-excesoY,excesoY);
                imagen.style.transform=`translate(${{x}}px,${{y}}px) scale(${{escala}})`;
                nivel.textContent=Math.round(escala*100)+' %';
            }}
            function foco() {{
                enfocado=true;
                marco.classList.add('activo');
                marco.focus({{preventScroll:true}});
            }}
            function cambiarZoom(destino,centroX=0,centroY=0) {{
                const previo=escala;
                escala=limitar(destino,1,maxZoom);
                // Mantener la zona apuntada próxima al mismo lugar al ampliar.
                if(escala!==previo) {{
                    x -= centroX*(escala/previo-1);
                    y -= centroY*(escala/previo-1);
                }}
                ajustar();
            }}
            marco.addEventListener('pointerdown', e => {{
                if(e.button!==0) return;
                foco();
                arrastrando=true;
                marco.classList.add('arrastrando');
                inicioX=e.clientX-x; inicioY=e.clientY-y;
                marco.setPointerCapture(e.pointerId);
            }});
            marco.addEventListener('pointermove', e => {{
                if(!arrastrando) return;
                x=e.clientX-inicioX; y=e.clientY-inicioY; ajustar();
            }});
            function soltar(e) {{
                arrastrando=false;
                marco.classList.remove('arrastrando');
                if(e && marco.hasPointerCapture(e.pointerId)) marco.releasePointerCapture(e.pointerId);
            }}
            marco.addEventListener('pointerup',soltar);
            marco.addEventListener('pointercancel',soltar);
            marco.addEventListener('wheel',e => {{
                if(!enfocado) return;
                e.preventDefault();
                const r=marco.getBoundingClientRect();
                cambiarZoom(escala*(e.deltaY<0?1.13:1/1.13),
                    e.clientX-r.left-r.width/2,e.clientY-r.top-r.height/2);
            }},{{passive:false}});
            function restablecer() {{
                escala=1; x=0; y=0; ajustar();
            }}
            marco.addEventListener('dblclick',e=>{{ e.preventDefault(); restablecer(); }});
            marco.addEventListener('keydown',e=>{{
                if(e.key==='Escape') {{ enfocado=false; marco.classList.remove('activo'); }}
                else if(e.key==='+' || e.key==='=') {{ cambiarZoom(escala*1.2); e.preventDefault(); }}
                else if(e.key==='-') {{ cambiarZoom(escala/1.2); e.preventDefault(); }}
                else if(e.key==='0') {{ restablecer(); e.preventDefault(); }}
            }});
            document.getElementById('pdf-mas').onclick=()=>{{ foco(); cambiarZoom(escala*1.2); }};
            document.getElementById('pdf-menos').onclick=()=>{{ foco(); cambiarZoom(escala/1.2); }};
            document.getElementById('pdf-reiniciar').onclick=()=>{{ restablecer(); }};
            ajustar();
        }})();
        </script></body></html>
        """

        if num_paginas > 1:
            col_izq, col_documento, col_der = st.columns(
                [0.8, 10, 0.8], vertical_alignment="center", gap="small")
            with col_izq:
                st.button("❮", key=f"pdf_ir_izquierda_{clave_pagina}",
                          on_click=cambiar_pagina_preview,
                          args=(clave_pagina, -1, num_paginas),
                          disabled=pag_seleccionada <= 1,
                          help="Página anterior", use_container_width=True)
            with col_documento:
                st.components.v1.html(html_lupa, height=alto_componente,
                                      scrolling=False)
            with col_der:
                st.button("❯", key=f"pdf_ir_derecha_{clave_pagina}",
                          on_click=cambiar_pagina_preview,
                          args=(clave_pagina, 1, num_paginas),
                          disabled=pag_seleccionada >= num_paginas,
                          help="Página siguiente", use_container_width=True)
        else:
            st.components.v1.html(html_lupa, height=alto_componente,
                                  scrolling=False)

    except Exception as err_preview:
        st.error(f"No se pudo generar la vista previa del PDF: {err_preview}")



def cambiar_pagina_preview(clave_slider, diferencia, total_paginas):
    """Una flecha actualiza el mismo slider, sin alterar la cola ni el análisis."""
    pagina = st.session_state.get(clave_slider, 1)
    st.session_state[clave_slider] = min(total_paginas, max(1, pagina + diferencia))


# --- GESTIÓN DE BASE DE DATOS SQLITE (CARTERA Y ÁREAS) ---
def init_db():
    """Inicializa la base de datos SQLite con tablas para Cartera, Áreas y Metadata."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Cartera de Exámenes
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS cartera_examenes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre_examen TEXT UNIQUE,
            codigo_examen TEXT,
            indicaciones TEXT,
            tecnologia TEXT,
            entrega_normal TEXT,
            entrega_urgente TEXT,
            laboratorio TEXT
        )
    """)
    # Compatibilidad con instalaciones anteriores: guardar una nueva Cartera
    # debe funcionar aunque la base todavía no tenga la columna LABORATORIO.
    # La operación es aditiva y no modifica usuarios, documentos ni fichas.
    columnas_cartera = {r[1] for r in cursor.execute("PRAGMA table_info(cartera_examenes)")}
    if "laboratorio" not in columnas_cartera:
        cursor.execute("ALTER TABLE cartera_examenes ADD COLUMN laboratorio TEXT")

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS cartera_metadata (
            clave TEXT PRIMARY KEY,
            valor TEXT
        )
    """)
    
    # Plantillas/Fichas por Área del Laboratorio
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS plantillas_area (
            area TEXT PRIMARY KEY,
            excel_bytes BLOB,
            fecha_actualizacion TEXT
        )
    """)
    
    # Exámenes pertenecientes a cada área
    # Compatibilidad con bases antiguas: la columna se añade sin perder las plantillas.
    columnas_plantillas = {r[1] for r in cursor.execute("PRAGMA table_info(plantillas_area)").fetchall()}
    if "nombre_archivo" not in columnas_plantillas:
        cursor.execute("ALTER TABLE plantillas_area ADD COLUMN nombre_archivo TEXT")

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS examenes_area (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            area TEXT,
            nombre_examen TEXT,
            UNIQUE(area, nombre_examen)
        )
    """)
    
    # Catálogos de selección persistentes (sin dependencia de archivos JSON).
    cursor.execute("""CREATE TABLE IF NOT EXISTS laboratorios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT NOT NULL COLLATE NOCASE UNIQUE,
        activo INTEGER NOT NULL DEFAULT 1 CHECK (activo IN (0, 1))
    )""")
    cursor.execute("""CREATE TABLE IF NOT EXISTS encargados (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT NOT NULL COLLATE NOCASE UNIQUE,
        activo INTEGER NOT NULL DEFAULT 1 CHECK (activo IN (0, 1))
    )""")
    conn.commit()

    # Cargar plantillas predeterminadas desde disco a la BD si aún no existen
    ZONA_HORARIA = ZoneInfo("America/Santiago")
    fecha_init = datetime.now(ZONA_HORARIA).isoformat()

    for area, config in AREA_CONFIG.items():
        cursor.execute("SELECT area FROM plantillas_area WHERE area = ?", (area,))
        if not cursor.fetchone():
            plantilla_nombre = config.get("template")
            if plantilla_nombre:
                ruta = os.path.join(os.path.dirname(__file__), plantilla_nombre)
                if os.path.exists(ruta):
                    try:
                        with open(ruta, "rb") as f:
                            b_data = f.read()
                        cursor.execute("""
                            INSERT OR REPLACE INTO plantillas_area
                            (area, excel_bytes, fecha_actualizacion, nombre_archivo)
                            VALUES (?, ?, ?, ?)
                        """, (area, b_data, fecha_init, plantilla_nombre))
                    except Exception:
                        pass
    # Identificar los archivos originales en instalaciones anteriores solo cuando
    # el contenido SQLite coincide exactamente con el XLSX disponible en disco.
    for area, config in AREA_CONFIG.items():
        nombre = config.get("template")
        if not nombre:
            continue
        ruta = os.path.join(os.path.dirname(__file__), nombre)
        if not os.path.isfile(ruta):
            continue
        cursor.execute("SELECT excel_bytes, nombre_archivo FROM plantillas_area WHERE area=?", (area,))
        fila = cursor.fetchone()
        if fila and fila[0] and not fila[1]:
            with open(ruta, "rb") as f:
                if hashlib.sha256(fila[0]).digest() == hashlib.sha256(f.read()).digest():
                    cursor.execute("UPDATE plantillas_area SET nombre_archivo=? WHERE area=?", (nombre, area))
    conn.commit()
    conn.close()

init_db()
# Al retirar la fuente antigua, respaldar la SQLite institucional ANTES de
# eliminar su PDF y dos textos. La limpieza no afecta Normas ni auditorías.
if fd.hay_manual_heredado(DB_PATH):
    gp.respaldar(DB_PATH)
fd.inicializar_fuentes(DB_PATH)
htest.inicializar(DB_PATH)
# Migración aditiva: recupera «LABORATORIO» desde el Excel original ya guardado.
# No borra fichas, cuentas, documentos ni el historial de SQLite.
import cartera_crs as crs
crs.migrar_laboratorio_cartera(DB_PATH)
gp.sincronizar_legado(DB_PATH)


def obtener_plantilla_area_db(area):
    """Obtiene los bytes del archivo Excel y la última fecha de actualización para un área."""
    if not area:
        return None, "Sin registro"
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT excel_bytes, fecha_actualizacion FROM plantillas_area WHERE area = ?", (area,))
    row = cursor.fetchone()
    conn.close()

    if row:
        b_data, f_act = row
        if f_act:
            try:
                f_dt = datetime.fromisoformat(f_act)
                msg_f = f_dt.strftime("%d/%m/%Y %H:%M")
            except Exception:
                msg_f = f_act
        else:
            msg_f = "Sin registro"
        return b_data, msg_f
    return None, "Sin registro"


def obtener_nombre_plantilla_area_db(area):
    """Devuelve el nombre del Excel almacenado o None si la versión antigua no lo registró."""
    with sqlite3.connect(DB_PATH) as conn:
        fila = conn.execute("SELECT nombre_archivo FROM plantillas_area WHERE area=?", (area,)).fetchone()
    return fila[0] if fila and fila[0] else None


def nombres_hojas_plantilla(area, excel_bytes):
    """Valida el Excel ANTES de escribir y devuelve los nombres de examen."""
    wb = load_workbook(io.BytesIO(excel_bytes), read_only=True, data_only=True)
    try:
        if area in ("Derivados", "Hormonas e inmunidad"):
            hojas = [str(ws.cell(2, 2).value or ws.title).strip() for ws in wb.worksheets]
        else:
            hojas = list(wb.sheetnames)
        hojas = sorted(set(h for h in hojas if h), key=clave_orden_alfabetico)
        if not hojas:
            raise ValueError("La planilla no contiene hojas de examen válidas.")
        return hojas
    finally:
        wb.close()


def actualizar_lista_area_conexion(con, area, nombres):
    """Modifica UNA SOLA área, sin sobrescribir exámenes de otras sesiones."""
    con.execute("DELETE FROM examenes_area WHERE area=?", (area,))
    for nombre in dict.fromkeys(str(n).strip() for n in nombres if str(n).strip()):
        con.execute("INSERT OR IGNORE INTO examenes_area(area,nombre_examen) VALUES(?,?)", (area,nombre))


def conservar_fichas_al_reconstruir_hematologia(plantilla_nueva, plantilla_anterior):
    """Importa la nueva estructura sin borrar datos ya ingresados en SQLite.

    Se ejecuta SOLO al cargar una plantilla desde el panel lateral, no al
    guardar la ficha generada y editada en «Actualizar en base de datos».
    """
    if not plantilla_anterior:
        return plantilla_nueva
    from copy import copy as copiar
    nueva = load_workbook(io.BytesIO(plantilla_nueva), rich_text=True)
    anterior = load_workbook(io.BytesIO(plantilla_anterior), rich_text=True)
    try:
        hojas_nuevas = {norm(hoja.title): hoja for hoja in nueva.worksheets}
        for hoja_antigua in anterior.worksheets:
            hoja_nueva = hojas_nuevas.get(norm(hoja_antigua.title))
            if hoja_nueva is None:
                # Las prestaciones creadas en el programa no se deben perder.
                hoja_nueva = nueva.create_sheet(hoja_antigua.title)
                for fila in hoja_antigua:
                    for celda in fila:
                        destino = hoja_nueva[celda.coordinate]
                        destino.value = copiar(celda.value)
                        if celda.has_style:
                            destino._style = copiar(celda._style)
                for nombre, columna in hoja_antigua.column_dimensions.items():
                    hoja_nueva.column_dimensions[nombre].width = columna.width
                continue
            mapa = {norm(hoja_antigua.cell(fila, 1).value): fila
                    for fila in range(1, hoja_antigua.max_row + 1)
                    if hoja_antigua.cell(fila, 1).value}
            for fila in range(1, hoja_nueva.max_row + 1):
                anterior_fila = mapa.get(norm(hoja_nueva.cell(fila, 1).value))
                if anterior_fila is None:
                    continue
                valor = hoja_antigua.cell(anterior_fila, 2).value
                if valor is not None and str(valor).strip():
                    destino = hoja_nueva.cell(fila, 2)
                    destino.value = copiar(valor)
                    destino.alignment = copiar(hoja_antigua.cell(anterior_fila, 2).alignment)
                    ajustar_alto_fila_con_listas(hoja_nueva, fila)
        buf = io.BytesIO()
        nueva.save(buf)
        return buf.getvalue()
    finally:
        nueva.close()
        anterior.close()


def guardar_plantilla_area_db(area, excel_bytes, nombre_archivo=None, conexion=None,
                              preservar_fichas_importadas=False):
    """Guarda plantilla y exámenes de esa área en una transacción compartida."""
    if area not in AREA_CONFIG or not excel_bytes:
        return False, "Datos no válidos."
    propia = conexion is None
    con = conexion or sqlite3.connect(DB_PATH, timeout=30)
    try:
        if preservar_fichas_importadas and area == "Hematología y coagulación":
            anterior = con.execute("SELECT excel_bytes FROM plantillas_area WHERE area=?",
                                   (area,)).fetchone()
            excel_bytes = conservar_fichas_al_reconstruir_hematologia(
                excel_bytes, anterior[0] if anterior else None)
        hojas = nombres_hojas_plantilla(area, excel_bytes)
        fecha = datetime.now(ZoneInfo("America/Santiago")).isoformat()
        con.execute("""INSERT OR REPLACE INTO plantillas_area
            (area,excel_bytes,fecha_actualizacion,nombre_archivo) VALUES(?,?,?,?)""",
            (area, sqlite3.Binary(excel_bytes),fecha, os.path.basename(nombre_archivo) if nombre_archivo else None))
        actualizar_lista_area_conexion(con, area, hojas)
        if propia:
            con.commit()
        return True, f"Se actualizó la base de datos de «{area}» exitosamente."
    except Exception:
        if propia:
            con.rollback()
        raise
    finally:
        if propia:
            con.close()


def cargar_examenes_por_area():
    """Carga los exámenes de cada área desde la BD; de estar vacía, consulta las plantillas cargadas."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    resultado = {area: [] for area in AREA_CONFIG}
    
    cursor.execute("SELECT area, nombre_examen FROM examenes_area ORDER BY nombre_examen ASC")
    rows = cursor.fetchall()
    conn.close()

    for area, nombre in rows:
        if area in resultado:
            resultado[area].append(nombre)

    # Si alguna área no tiene exámenes registrados en la tabla, poblar desde la plantilla SQLite
    for area in AREA_CONFIG:
        if not resultado[area]:
            b_bytes, _ = obtener_plantilla_area_db(area)
            if b_bytes:
                try:
                    wb = load_workbook(io.BytesIO(b_bytes), read_only=True, data_only=True)
                    if area in ("Derivados", "Hormonas e inmunidad"):
                        resultado[area] = [str(ws.cell(2, 2).value or ws.title).strip()
                                           for ws in wb.worksheets]
                    else:
                        resultado[area] = list(wb.sheetnames)
                    wb.close()
                except Exception:
                    pass

    return {a: sorted(list(dict.fromkeys(str(x).strip() for x in xs if str(x).strip())), key=clave_orden_alfabetico) for a, xs in resultado.items()}


def guardar_examenes_por_area(examenes_dict):
    """Guarda los exámenes asociados a cada área en SQLite."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM examenes_area")

    for area, lista in examenes_dict.items():
        vistos = set()
        for nombre in lista:
            nombre = str(nombre).strip()
            if nombre and nombre.casefold() not in vistos:
                cursor.execute("INSERT OR IGNORE INTO examenes_area (area, nombre_examen) VALUES (?, ?)", (area, nombre))
                vistos.add(nombre.casefold())

    conn.commit()
    conn.close()
    return True


def guardar_cartera_db(df, nombre_archivo=None, archivo_original=None, conexion=None):
    """Guarda la Cartera y, si se proporciona, conserva el XLSX original y su nombre."""
    propia = conexion is None
    conn = conexion or sqlite3.connect(DB_PATH, timeout=30)
    cursor = conn.cursor()

    col_nombre_series = None
    
    for col in df.columns:
        if str(col).strip().upper() in ["DESCRIPCION", "DESCRIPCIÓN"]:
            col_nombre_series = df[col]
            break

    if col_nombre_series is None and len(df.columns) >= 6:
        col_nombre_series = df.iloc[:, 5]
    elif col_nombre_series is None and len(df.columns) > 0:
        col_nombre_series = df.iloc[:, 0]

    cols_exactas = {norm(str(c)): c for c in df.columns}

    def buscar_col(alias_list):
        for a in alias_list:
            if a in cols_exactas:
                return cols_exactas[a]
        for a in alias_list:
            for k_norm, k_real in cols_exactas.items():
                if a in k_norm:
                    return k_real
        return None

    col_codigo = buscar_col(["codigo examen", "codigo host", "codigo", "código"])
    col_indicaciones = buscar_col(["indicaciones", "preparacion paciente", "preparación paciente", "preparacion"])
    col_laboratorio = buscar_col(["laboratorio"])
    col_tecnologia = buscar_col(["tecnologia", "tecnología", "equipamiento", "equipo", "instrumento"])
    col_normal = buscar_col(["entrega normal", "plazo normal", "normal", "tat normal"])
    col_urgente = buscar_col(["entrega urgente", "plazo urgente", "urgente", "tat urgente"])

    cursor.execute("DELETE FROM cartera_examenes")

    filas_guardadas = 0
    for idx, nombre_raw in col_nombre_series.items():
        nombre = str(nombre_raw).strip()

        if not nombre or nombre.casefold() in ["nan", "none", "null", "descripcion", "descripción"] or nombre.isdigit():
            continue

        row = df.loc[idx]

        codigo = str(row.get(col_codigo, "")).strip() if col_codigo else ""
        indicaciones = str(row.get(col_indicaciones, "")).strip() if col_indicaciones else ""
        laboratorio = str(row.get(col_laboratorio, "")).strip() if col_laboratorio else ""
        tecnologia = str(row.get(col_tecnologia, "")).strip() if col_tecnologia else ""
        normal = str(row.get(col_normal, "")).strip() if col_normal else ""
        urgente = str(row.get(col_urgente, "")).strip() if col_urgente else ""

        codigo = "" if codigo.casefold() in ["nan", "none", "null"] else codigo
        indicaciones = "" if indicaciones.casefold() in ["nan", "none", "null"] else indicaciones
        laboratorio = "" if laboratorio.casefold() in ["nan", "none", "null"] else laboratorio
        tecnologia = "" if tecnologia.casefold() in ["nan", "none", "null"] else tecnologia
        normal = "" if normal.casefold() in ["nan", "none", "null"] else normal
        urgente = "" if urgente.casefold() in ["nan", "none", "null"] else urgente

        cursor.execute("""
            INSERT OR REPLACE INTO cartera_examenes (nombre_examen, codigo_examen, indicaciones, tecnologia, entrega_normal, entrega_urgente, laboratorio)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (nombre, codigo, indicaciones, tecnologia, normal, urgente, laboratorio))
        filas_guardadas += 1

    if filas_guardadas == 0:
        if propia:
            conn.rollback()
            conn.close()
        return False, "No se encontraron prestaciones válidas en la Cartera. No se modificó la base de datos."

    ZONA_HORARIA = ZoneInfo("America/Santiago")
    fecha_actual = datetime.now(ZONA_HORARIA).isoformat()
    cursor.execute("INSERT OR REPLACE INTO cartera_metadata (clave, valor) VALUES ('ultima_actualizacion', ?)", (fecha_actual,))
    if nombre_archivo and archivo_original:
        nombre_seguro = os.path.basename(nombre_archivo)
        cursor.execute("INSERT OR REPLACE INTO cartera_metadata (clave, valor) VALUES ('nombre_archivo', ?)", (nombre_seguro,))
        # El original se almacena igual que los otros documentos de referencia.
        cursor.execute("""INSERT OR REPLACE INTO documentos_fuente
            (fuente_id,nombre_archivo,tipo_mime,archivo,sha256,fecha_actualizacion,registros)
            VALUES (?,?,?,?,?,?,?)""",
            ("cartera_examenes", nombre_seguro,
             "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
             sqlite3.Binary(archivo_original),hashlib.sha256(archivo_original).hexdigest(),
             fecha_actual, cursor.execute("SELECT COUNT(*) FROM cartera_examenes").fetchone()[0]))

    if propia:
        conn.commit()
        conn.close()
    return True, f"Se importaron exitosamente {filas_guardadas} exámenes desde la Columna 6 (DESCRIPCION) a la Cartera."


def obtener_lista_cartera_db():
    """Retorna la lista ordenada de nombres/DESCRIPCION de exámenes almacenados en la Cartera."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    # No filtrar nombres que empiecen por un dígito: son exámenes legítimos,
    # como 17-HIDROXIPROGESTERONA y 17-BETA-ESTRADIOL.
    cursor.execute("SELECT nombre_examen FROM cartera_examenes ORDER BY nombre_examen COLLATE NOCASE ASC")
    rows = cursor.fetchall()
    conn.close()
    return sorted((r[0] for r in rows if r[0]), key=clave_orden_alfabetico)


def buscar_top_coincidencias_cartera(nombre_buscar, area=None, top_n=8, similitud_minima=0.0):
    """Sugerencias CRS: primero el nombre clínico y luego parecidos válidos.

    Las opciones de otra área siguen disponibles; ninguna coincidencia difusa
    se vincula automáticamente sin confirmación del encargado.
    """
    return crs.candidatos_cartera(
        DB_PATH, nombre_buscar, area, top_n, similitud_minima=similitud_minima)


def obtener_info_cartera_por_nombre(nombre_exacto_cartera):
    """Consulta SQLite obteniendo los datos completos de un examen de la Cartera."""
    if not nombre_exacto_cartera:
        return None
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT codigo_examen, indicaciones, tecnologia, entrega_normal, entrega_urgente FROM cartera_examenes WHERE nombre_examen = ?", (nombre_exacto_cartera.strip(),))
    row = cursor.fetchone()
    conn.close()

    if row:
        codigo, indicaciones, tecnologia, normal, urgente = row
        plazo_partes = []
        if normal:
            plazo_partes.append(f"Entrega Normal: {normal}")
        if urgente:
            plazo_partes.append(f"Entrega Urgente: {urgente}")
        plazo_total = " | ".join(plazo_partes)

        return {
            "Código": codigo or "",
            "Preparación del paciente": indicaciones or "",
            "Equipamiento utilizado": tecnologia or "",
            "Plazo entrega resultados": plazo_total or ""
        }
    return None


def fecha_actualizacion_cartera():
    """Consulta la última fecha de actualización y valida los 6 meses (180 días)."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT valor FROM cartera_metadata WHERE clave = 'ultima_actualizacion'")
    row = cursor.fetchone()
    conn.close()

    if not row or not row[0]:
        return None, True, "No se ha subido ninguna Cartera de Exámenes a la base de datos."

    ZONA_HORARIA = ZoneInfo("America/Santiago")
    fecha_dt = datetime.fromisoformat(row[0])
    dias_transcurridos = (datetime.now(ZONA_HORARIA) - fecha_dt).days
    es_antiguo = dias_transcurridos >= 180

    msg_fecha = fecha_dt.strftime("%d/%m/%Y %H:%M")
    return fecha_dt, es_antiguo, msg_fecha

def obtener_cartera_excel_bytes():
    """Genera y retorna el archivo Excel (.xlsx) con todos los exámenes registrados en la Cartera SQLite."""
    conn = sqlite3.connect(DB_PATH)
    query = """
        SELECT 
            nombre_examen AS "DESCRIPCION",
            codigo_examen AS "Código Examen",
            indicaciones AS "Indicaciones / Preparación del paciente",
            tecnologia AS "Equipamiento / Tecnología",
            entrega_normal AS "Entrega Normal",
            entrega_urgente AS "Entrega Urgente",
            laboratorio AS "Laboratorio"
        FROM cartera_examenes
        ORDER BY nombre_examen ASC
    """
    df = pd.read_sql_query(query, conn)
    conn.close()

    if df.empty:
        return None

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Cartera de Exámenes")
    return output.getvalue()

# --- CATÁLOGOS PERSISTENTES EN SQLITE ---
# Los dos selectbox de responsables comparten la misma tabla de encargados.
def _normalizar_nombres(nombres):
    resultado = {}
    for nombre in nombres:
        nombre = " ".join(str(nombre).split())
        if nombre:
            resultado.setdefault(nombre.casefold(), nombre)
    return resultado


def _listar_catalogo(tabla):
    if tabla not in ("laboratorios", "encargados"):
        raise ValueError("Catálogo no permitido")
    with sqlite3.connect(DB_PATH, timeout=15) as conn:
        filas = conn.execute(f"SELECT nombre FROM {tabla} WHERE activo = 1").fetchall()
    return sorted((fila[0] for fila in filas), key=clave_orden_alfabetico)


def _guardar_catalogo(tabla, nombres):
    """Sincroniza opciones activas; desactiva sin borrar los registros históricos."""
    if tabla not in ("laboratorios", "encargados"):
        raise ValueError("Catálogo no permitido")
    nombres = _normalizar_nombres(nombres)
    try:
        with sqlite3.connect(DB_PATH, timeout=15) as conn:
            existentes = {fila[0].casefold(): fila[0] for fila in
                          conn.execute(f"SELECT nombre FROM {tabla}")}
            for clave, nombre in nombres.items():
                if clave in existentes:
                    conn.execute(f"UPDATE {tabla} SET activo = 1 WHERE nombre = ?", (existentes[clave],))
                else:
                    conn.execute(f"INSERT INTO {tabla} (nombre, activo) VALUES (?, 1)", (nombre,))
            for clave, nombre in existentes.items():
                if clave not in nombres:
                    conn.execute(f"UPDATE {tabla} SET activo = 0 WHERE nombre = ?", (nombre,))
        return True
    except sqlite3.Error as exc:
        st.error(f"No se pudieron guardar los cambios en la base de datos: {exc}")
        return False


def cargar_laboratorios():
    return _listar_catalogo("laboratorios")


def guardar_lista_laboratorios(laboratorios):
    return _guardar_catalogo("laboratorios", laboratorios)


def guardar_laboratorio(nombre):
    nombre = " ".join(str(nombre).split())
    actuales = cargar_laboratorios()
    if not nombre or nombre.casefold() in {n.casefold() for n in actuales}:
        return False
    return guardar_lista_laboratorios(actuales + [nombre])


def cargar_encargados(tipo="actualizado"):
    """Ambos tipos comparten las mismas opciones almacenadas en SQLite."""
    return _listar_catalogo("encargados")


def guardar_encargados(tipo, lista):
    """Una modificación desde cualquier pestaña afecta a ambos selectbox."""
    return _guardar_catalogo("encargados", lista)


def nombre_examen_valido(nombre):
    """Conserva íntegra DESCRIPCION aunque Excel requiera un título de hoja corto."""
    nombre = str(nombre or "").strip()
    if not nombre:
        return False, "Escribe un nombre para el examen."
    if len(nombre) > 250:
        return False, "El nombre del examen supera los 250 caracteres permitidos."
    return True, ""


def titulo_hoja_para_examen(wb, nombre):
    """Título único y compatible con Excel (31 caracteres y sin signos prohibidos)."""
    base = re.sub(r'[\\/*?:\[\]]', '-', str(nombre or '').strip()).strip(" '")[:31] or "Examen"
    usados = {titulo.casefold() for titulo in wb.sheetnames}
    candidato = base
    numero = 2
    while candidato.casefold() in usados:
        sufijo = f" ({numero})"
        candidato = base[:31-len(sufijo)].rstrip() + sufijo
        numero += 1
    return candidato


def buscar_hoja_por_examen(wb, nombre):
    """Localiza una ficha por su nombre completo LIS, aunque su hoja esté abreviada."""
    if nombre in wb.sheetnames:
        return wb[nombre]
    buscado = norm(nombre)
    for hoja in wb.worksheets:
        if norm(hoja.cell(2, 2).value) == buscado:
            return hoja
    return None


def ajustar_ancho_columnas(ws):
    """Ajuste aproximado de ancho de columnas."""
    from openpyxl.utils import get_column_letter
    for col_idx in range(1, ws.max_column + 1):
        max_len = 0
        for row in range(1, ws.max_row + 1):
            value = ws.cell(row, col_idx).value
            if value is not None:
                max_len = max(max_len, max((len(line) for line in str(value).splitlines()), default=0))
        columna = ws.column_dimensions[get_column_letter(col_idx)]
        columna.width = max(12, min(max_len + 2, 65))
        # Indicar a Microsoft Excel que el ancho fue calculado para
        # ajustarse al contenido. No ejecuta VBA ni requiere macros.
        columna.bestFit = True


def ajustar_alto_fila_con_listas(ws, fila, columna=2):
    """Ajusta la altura de la fila según el texto real de la celda.

    La plantilla puede traer alturas fijas demasiado grandes o demasiado
    pequeñas. Aquí se estima la altura visual del contenido envuelto en la
    columna B y se asigna una altura exacta, con un margen moderado.
    """
    celda = ws.cell(fila, columna)
    texto = str(celda.value or "").replace("\r\n", "\n").replace("\r", "\n")
    if not texto.strip():
        ws.row_dimensions[fila].height = 15
        return

    ancho = ws.column_dimensions[celda.column_letter].width or 55
    if sum(bool(re.match(r"^\s*(?:[•●▪◦∙]|\d{1,3}[.)])\s+", l))
           for l in texto.split("\n")) >= 5 and ancho < 90:
        ws.column_dimensions[celda.column_letter].width = 90
        ancho = 90

    # Aproximación de caracteres visibles por línea en Calibri 11.
    caracteres_linea = max(18, min(110, int(ancho * 1.05)))
    lineas_visuales = 0
    for linea in texto.split("\n"):
        contenido = linea.expandtabs(4).strip()
        if not contenido:
            lineas_visuales += 1
            continue
        ajuste_bulleta = 2 if re.match(r"^\s*(?:[•●▪◦∙]|\d{1,3}[.)])\s+", linea) else 0
        capacidad = max(12, caracteres_linea - ajuste_bulleta)
        lineas_visuales += max(1, (len(contenido) + capacidad - 1) // capacidad)

    # Altura base + altura por línea; se permite reducir filas sobredimensionadas.
    alto_calculado = min(409.5, max(15, 4 + lineas_visuales * 13.2))
    ws.row_dimensions[fila].height = alto_calculado

def convertir_subindices(texto):
    """Convierte cifras pegadas a letras en subíndice (CO2 → CO₂), evitando expresiones matemáticas."""
    tabla = str.maketrans("0123456789", "₀₁₂₃₄₅₆₇₈₉")
    operadores = "=<>≤≥±×÷≈+-*/^"
    resultado = []
    for linea in str(texto or "").splitlines():
        def reemplazar(match):
            inicio, fin = match.span()
            alrededor = linea[max(0, inicio - 3):min(len(linea), fin + 4)]
            if any(op in alrededor for op in operadores):
                return match.group(1)
            return match.group(1).translate(tabla)

        resultado.append(re.sub(r"(?<=[A-Za-z])(\d+)(?=[A-Za-z]|\b)", reemplazar, linea))
    return "\n".join(resultado)


def quitar_citas_superindices(texto):
    """Elimina índices de citas, sin borrar exponentes científicos ni valores clínicos.

    Algunos PDF representan las referencias voladas como superíndices Unicode
    («texto²⁹»), mientras otros las aplanan («texto29» o «texto.29»).
    Solo se retiran números adjuntos con apariencia de cita; los números
    independientes se conservan porque pueden ser resultados del examen.
    """
    if not texto:
        return ""

    s = str(texto)

    # Superíndices Unicode: suprimir referencias, conservando exponentes
    # matemáticos (10⁶, x²) y unidades (m², cm³, mm³).
    def tratar_superindice(m):
        inicio = m.start()
        antes = s[:inicio]
        if antes:
            if antes[-1].isdigit() or antes[-1] in "^⁻⁺":
                return m.group(0)
            if re.search(r"(?<![\w])(?:m|cm|mm|dm|km|µm|μm|nm|x|y|z)$", antes, re.I):
                return m.group(0)
        return ""

    s = re.sub(r"[⁰¹²³⁴⁵⁶⁷⁸⁹]+", tratar_superindice, s)

    # Referencias expresadas con paréntesis o corchetes: «texto[29]».
    s = re.sub(r"\[\s*\d{1,3}(?:\s*,\s*\d{1,3})*\s*\]", "", s)
    s = re.sub(r"\(\s*\d{1,3}(?:\s*,\s*\d{1,3})*\s*\)", "", s)

    # El punto forma parte de la oración: «(µg de UEF/mL).29» -> «(...).»
    # No modifica decimales (0.5) ni números de versión abreviados (v.2).
    def quitar_despues_de_punto(m):
        prefijo = m.string[:m.start() + 1]
        if prefijo.endswith(").") or prefijo.endswith("]."):
            return "."
        palabra = re.search(r"([A-Za-zÁÉÍÓÚÜÑáéíóúüñ]+)\.$", prefijo)
        return "." if palabra and len(palabra.group(1)) >= 3 else m.group(0)

    s = re.sub(
        r"(?<=[A-Za-zÁÉÍÓÚÜÑáéíóúüñ)\]])\.\s*\d{1,3}"
        r"(?=\s|[.,;:!?\n]|$|[A-ZÁÉÍÓÚÜÑ])",
        quitar_despues_de_punto, s,
    )
    # Referencia a continuación de un paréntesis, sin punto: «(texto)29».
    s = re.sub(r"(?<=[)\]])\d{1,3}(?=\s|[.,;:!?\n]|$)", "", s)

    # PDF con superíndices aplanados: «laboratorio29 indicó…».
    # Se exige palabra previa de 4+ letras para proteger CO2, T3, CD4, etc.
    def quitar_despues_de_palabra(m):
        palabra = re.search(r"([A-Za-zÁÉÍÓÚÜÑáéíóúüñ]+)$", m.string[:m.start()])
        return "" if palabra and len(palabra.group(1)) >= 4 else m.group(0)

    s = re.sub(
        r"(?<=[A-Za-zÁÉÍÓÚÜÑáéíóúüñ])\d{1,3}"
        r"(?=\s|[.,;:!?\n]|$|[A-ZÁÉÍÓÚÜÑ])",
        quitar_despues_de_palabra, s,
    )
    s = re.sub(r"\s+([.,;:])", r"\1", s)
    s = re.sub(r"^\s*,\s*", "", s, flags=re.MULTILINE)
    s = re.sub(r"[ \t]+", " ", s)
    return s.strip()


def es_linea_basura_o_encabezado(s, examen_activo=None):
    """Detecta y filtra números de página, códigos de versión y encabezados repetidos."""
    s_clean = s.strip()
    if not s_clean:
        return True
    if re.fullmatch(r"(?:p[aá]gina\s*)?\d+\s*(?:de|[/])\s*\d+", s_clean, flags=re.I):
        return True
    if re.fullmatch(r"\d{7,}[vV]?\d*(?:\.\d+)*\.?", s_clean):
        return True
    if re.fullmatch(r"\d{4}-\d{2},\s*v\s*[\d\.]+\s*.*", s_clean, flags=re.I):
        return True
    if s_clean.casefold() in ["cobas®", "roche diagnostics", "español", "content"]:
        return True
    if examen_activo:
        term_exam = str(examen_activo).strip().casefold()
        if s_clean.casefold() == term_exam or norm(s_clean) in norm(term_exam):
            return True
    return False


def extraer_solo_valor_referencia(texto):
    """Extrae únicamente el valor o rango numérico/cualitativo principal de los Valores de Referencia, omitiendo explicaciones."""
    if not texto:
        return ""
    t = str(texto).strip()
    if not t or t.casefold() in ["nan", "none", "null"]:
        return ""

    # Cortar párrafos explicativos o notas metodológicas comunes
    patrones_corte = [
        r"\n?\s*Estos valores corresponden.*",
        r"\n?\s*Cada laboratorio deber[ií]a.*",
        r"\n?\s*Datos espec[ií]ficos de funcionamiento.*",
        r"\n?\s*A continuaci[oó]n,\s*se indican.*",
        r"\n?\s*Los resultados de cada laboratorio.*",
        r"\n?\s*Obtenidos a partir de.*",
        r"\n?\s*Consulte la secci[oó]n.*",
        r"\n?\s*Intervalo de referencia.*",
    ]
    for pat in patrones_corte:
        t = re.split(pat, t, flags=re.IGNORECASE | re.DOTALL)[0]

    lineas = [l.strip() for l in t.splitlines() if l.strip()]
    if not lineas:
        return ""

    # Las referencias expresadas en tiempo pueden llevar explicaciones a
    # continuación del número. «s», «seg» y «segundos» son unidades válidas,
    # no encabezados que deban arrastrar el resto del párrafo.
    unidad_tiempo = r"(?:segundos?|seg\.?|s)\b"
    numero = r"(?:[<>≤≥≈=]\s*)?\d+(?:[.,]\d+)?"
    rango = rf"{numero}(?:\s*(?:[-–—]|a|hasta)\s*{numero})?"
    patron_tiempo = re.compile(rf"(?<![\w])({rango}\s*{unidad_tiempo})", re.I)
    primera = lineas[0]
    resultados_tiempo = list(patron_tiempo.finditer(primera))
    if resultados_tiempo:
        # Conservar los grupos identificados por sexo o edad si aparecen
        # varios intervalos en la misma línea, sin copiar la explicación.
        if len(resultados_tiempo) > 1:
            grupos = []
            inicio = 0
            for m in resultados_tiempo:
                previo = primera[inicio:m.start()].strip(" .;,:–—")
                etiqueta = re.search(r"(?:^|[;,.])\s*((?:hombres?|mujeres?|adultos?|niños?|niñas?))\s*:?\s*$", previo, re.I)
                grupos.append(((etiqueta.group(1) + ': ') if etiqueta else '') + m.group(1).strip())
                inicio = m.end()
            return '; '.join(grupos)
        return resultados_tiempo[0].group(1).strip()

    # Si no hay un intervalo de tiempo identificable, conservar el criterio
    # anterior para otras unidades y los resultados cualitativos.
    return primera


def limpiar_valor_excel(valor, examen_activo=None):
    """Quita espacios accidentales, líneas de ruido y citas bibliográficas conservando valores válidos."""
    if valor is None:
        return ""

    valor_str = str(valor).strip()
    if not valor_str or valor_str.casefold() in ["nan", "none", "null"]:
        return ""

    if examen_activo and valor_str.casefold() == str(examen_activo).strip().casefold():
        return valor_str

    valor_sin_citas = quitar_citas_superindices(valor)
    # En algunos PDF el símbolo de viñeta se extrae como un bloque separado
    # de su texto. Unirlos evita que Excel muestre el marcador aislado arriba
    # y la instrucción en otro renglón. No se agregan viñetas inexistentes.
    valor_sin_citas = re.sub(
        r"(?m)^[ \t]*[•●▪◦∙][ \t]*(?:\n[ \t]*)+(?=\S)",
        "  • ", valor_sin_citas,
    )
    # Reconstruir el prefijo de las viñetas tras quitar citas: el limpiador
    # general de citas compacta los espacios, por lo que antes se perdía la
    # sangría y Excel acababa mostrando el marcador pegado al margen.
    # Cada paso/lista conserva su propio renglón y su orden original.
    lineas = []
    for original in valor_sin_citas.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        linea = re.sub(r"[ \t]+", " ", original).strip()
        marca = re.match(r"^([●•▪◦∙]|(?:\d{1,3}|[a-zA-Z])[.)])\s+(.+)$", linea)
        if marca:
            marcador = "•" if marca.group(1) in "●•▪◦∙" else marca.group(1)
            linea = "  " + marcador + " " + marca.group(2).strip()
        lineas.append(linea)
    salida = []
    blancos = 0
    for linea in lineas:
        if not linea:
            blancos += 1
            if blancos == 1 and salida:
                salida.append("")
            continue

        if es_linea_basura_o_encabezado(linea, examen_activo):
            if examen_activo and linea.casefold() == str(examen_activo).strip().casefold() and len(lineas) == 1:
                salida.append(linea)
                continue
            blancos += 1
            if blancos == 1 and salida:
                salida.append("")
            continue

        blancos = 0
        if re.fullmatch(r"https?://\S+|www\.\S+", linea, flags=re.I):
            continue
        salida.append(linea)
    return convertir_subindices("\n".join(salida).strip("\n"))


# --- ALIAS ESPECÍFICOS POR CADA ÁREA DEL LABORATORIO ---

FIELD_ALIASES_BASE = {
    'Código': [],
    'Nombre LIS': [],
    'Laboratorio de procesamiento': [],
    'Recepción de muestras': [],
    'Procesamiento de muestras': [],
    'Plazo entrega resultados': ['tiempo de respuesta', 'plazo entrega resultados', 'plazo de entrega', 'tat', 'tiempo procesamiento'],
    'Preparación del paciente': ['preparacion paciente', 'preparación paciente', 'preparacion del paciente', 'preparación del paciente'],
    'Procedimiento de toma de muestras': ['Procedimiento', 'Procedimiento de toma de muestras', 'Procedimiento de examen', 'Procedimiento del examen'],
    'Muestra requerida y recipiente': ['tipo de muestra', 'muestra requerida', 'muestra', 'tipo y volumen de muestra', 'material biologico', 'material biológico', 'volumen requerido', 'obtencion y preparacion de las muestras', 'obtención y preparación de las muestras'],
    'Rotulación': ['Rotulación', 'Rotulacion'],
    'Estabilidad de la muestra': ['estabilidad muestra', 'estabilidad de la muestra', 'Conservación y estabilidad', 'conservacion y estabilidad'],
    'Condiciones de envío al Laboratorio': ['Condiciones de envío al Laboratorio', 'temperatura de transporte', 'condiciones de envio', 'condiciones de envío', 'envio al laboratorio', 'envío al laboratorio', 'transporte'],
    'Criterios de rechazo': ['Rechazo', 'Criterios de rechazo', 'Criterios rechazo', 'Criterios del rechazo'],
    'Método de procesamiento': ['metodo utilizado', 'método utilizado', 'metodo de procesamiento', 'método de procesamiento', 'metodologia', 'metodología', 'metodo', 'principio del test', 'principio', 'principio de la prueba', 'principio del ensayo'],
    'Test incluidos': ['test incluidos', 'pruebas incluidas', 'examenes incluidos', 'exámenes incluidos'],
    'Valores de referencia': ['valor de referencia', 'valores de referencia', 'valores normales', 'rango de referencia'],
    'Valor crítico': [],
    'Límite de detección': ['limite de deteccion', 'límite de detección', 'limite deteccion', 'límite deteccion', 'límite de cuantificación', 'limite de cuantificacion', 'loq'],
    'Equipamiento utilizado': ['equipamiento utilizado', 'equipo utilizado', 'instrumentacion', 'instrumentación', 'equipo', 'system', 'sistema', 'informacion del sistema', 'información del sistema'],
    'Información clínica': ['caracteristicas', 'características', 'utilidad clinica', 'utilidad clínica', 'informacion clinica', 'información clínica'],
    'Referencias': ['referencias', 'bibliografia', 'bibliografía', 'fuentes'],
    'Actualizado por': [],
    'Fecha actualización': [],
    'Autorizado por': [],
}

FIELD_ALIASES_MICROBIOLOGIA = {
    'Código': [],
    'Nombre LIS': [],
    'Laboratorio de procesamiento': [],
    'Recepción de muestras': [],
    'Procesamiento de muestras': [],
    'Plazo entrega resultados': ['tiempo de respuesta', 'plazo entrega resultados', 'plazo de entrega', 'tat', 'tiempo incubacion', 'dias cultivo'],
    'Preparación del paciente': ['preparacion paciente', 'preparación paciente', 'preparacion del paciente'],
    'Procedimiento de toma de muestras': ['Procedimiento', 'Procedimiento de toma de muestras', 'obtencion muestra microbiologica', 'tecnica hisopado', 'toma de muestra'],
    'Muestra requerida y recipiente': ['tipo de muestra', 'muestra requerida', 'muestra', 'medio de transporte', 'frasco estéril', 'hisopo', 'recipiente'],
    'Rotulación': ['Rotulación', 'Rotulacion'],
    'Estabilidad de la muestra': ['estabilidad muestra', 'estabilidad de la muestra', 'Conservación y estabilidad'],
    'Condiciones de envío al Laboratorio': ['Condiciones de envío al Laboratorio', 'temperatura de transporte', 'condiciones de envio', 'transporte muestra'],
    'Criterios de rechazo': ['Rechazo', 'Criterios de rechazo', 'muestra contaminada', 'hisopo seco'],
    'Método de procesamiento': ['metodo utilizado', 'método utilizado', 'metodo de procesamiento', 'metodologia', 'tintura gram', 'siembra en agar', 'cultivo', 'antibiograma', 'pcr microbiologico'],
    'Test incluidos': ['test incluidos', 'pruebas incluidas', 'identificacion bacteriana', 'sensibilidad antibiotica'],
    'Valores de referencia': ['valor de referencia', 'valores de referencia', 'flora normal', 'sin desarrollo bacteriano', 'resultado esperado'],
    'Valor crítico': [],
    'Límite de detección': ['limite de deteccion', 'límite de detección', 'ufc/ml min', 'sensibilidad analítica'],
    'Equipamiento utilizado': ['equipamiento utilizado', 'equipo utilizado', 'estufa incubadora', 'autoclave', 'vitek', 'maldi-tof', 'sistema automatizado de hemocultivos'],
    'Información clínica': ['caracteristicas', 'características', 'utilidad clinica', 'informacion clinica', 'asociación infecciosa'],
    'Referencias': ['referencias', 'bibliografia', 'clsi', 'eucast'],
    'Actualizado por': [],
    'Fecha actualización': [],
    'Autorizado por': [],
}

FIELD_ALIASES_HEMATOLOGIA = {
    'Código': [],
    'Nombre LIS': [],
    'Laboratorio de procesamiento': [],
    'Recepción de muestras': [],
    'Procesamiento de muestras': [],
    'Plazo entrega resultados': ['tiempo de respuesta', 'plazo entrega resultados', 'plazo de entrega', 'tat', 'tiempo procesamiento'],
    'Preparación del paciente': ['preparacion paciente', 'preparación paciente', 'ayuno hematologia'],
    'Procedimiento de toma de muestras': ['Procedimiento', 'Procedimiento de toma de muestras', 'puncion venosa', 'anticoagulante edta', 'tubo citrato'],
    'Muestra requerida y recipiente': ['tipo de muestra', 'muestra requerida', 'sangre total', 'plasma citratado', 'tubo tapa morada', 'tubo tapa azul'],
    'Rotulación': ['Rotulación', 'Rotulacion'],
    'Estabilidad de la muestra': ['Estabilidad:'],
    'Condiciones de envío al Laboratorio': ['Condiciones de envío al Laboratorio', 'temperatura de transporte', 'envio temperatura ambiente', 'no congelar sangre total'],
    'Criterios de rechazo': ['Rechazo', 'Criterios de rechazo', 'muestra coagulada', 'muestra hemolisada', 'volumen insuficiente', 'relación anticoagulante'],
    'Método de procesamiento': ['Principio del test','metodo utilizado', 'método utilizado', 'metodo de procesamiento', 'metodologia'],
    'Test incluidos': ['test incluidos', 'pruebas incluidas', 'fórmula leucocitaria', 'indices hematimetricos'],
    'Valores de referencia': ['valores teóricos'],
    'Valor crítico': [],
    'Límite de detección': ['limite de detección:', 'límite de detección', 'rango de linealidad'],
    'Equipamiento utilizado': ['equipamiento utilizado', 'equipo utilizado', 'contador hematologico', 'coagulometro', 'sysmex', 'stago'],
    'Información clínica': ['caracteristicas', 'características'],
    'Referencias': ['referencias', 'bibliografia', 'fuentes'],
    'Actualizado por': [],
    'Fecha actualización': [],
    'Autorizado por': [],
}

FIELD_ALIASES_BIOQUIMICA = {
    'Código': [],
    'Nombre LIS': [],
    'Laboratorio de procesamiento': [],
    'Recepción de muestras': [],
    'Procesamiento de muestras': [],
    'Plazo entrega resultados': ['tiempo de respuesta', 'plazo entrega resultados', 'plazo de entrega', 'tat'],
    'Preparación del paciente': ['preparacion paciente', 'preparación paciente', 'ayuno de 8 a 12 horas', 'restricción dietética'],
    'Procedimiento de toma de muestras': ['Procedimiento', 'Procedimiento de toma de muestras', 'puncion venosa', 'tubo suero gel'],
    'Muestra requerida y recipiente': ['tipo de muestra', 'muestra requerida', 'suero', 'plasma', 'orina de 24 horas', 'tubo tapa roja', 'tubo tapa amarilla'],
    'Rotulación': ['Rotulación', 'Rotulacion'],
    'Estabilidad de la muestra': ['estabilidad muestra', 'estabilidad de la muestra', 'estabilidad del suero', 'conservacion refrigerada'],
    'Condiciones de envío al Laboratorio': ['Condiciones de envío al Laboratorio', 'temperatura de transporte', 'suero separado', 'transporte refrigerado'],
    'Criterios de rechazo': ['Rechazo', 'Criterios de rechazo', 'suero hemolisado', 'suero icterico', 'suero lipemico'],
    'Método de procesamiento': ['metodo utilizado', 'método utilizado', 'metodo de procesamiento', 'metodologia', 'espectrofotometria', 'enzimatico', 'colorimetrico', 'electrodo ion selectivo', 'ise'],
    'Test incluidos': ['test incluidos', 'pruebas incluidas', 'perfil hepatico', 'perfil lipidico', 'perfil renal'],
    'Valores de referencia': ['valor de referencia', 'valores de referencia', 'rangos de referencia', 'valores normales'],
    'Valor crítico': [],
    'Límite de detección': ['limite de deteccion', 'límite de detección', 'linealidad analítica', 'loq'],
    'Equipamiento utilizado': ['equipamiento utilizado', 'equipo utilizado', 'analizador bioquimico', 'cobas', 'architect'],
    'Información clínica': ['caracteristicas', 'características', 'utilidad clinica', 'evaluación metabólica', 'función renal'],
    'Referencias': ['referencias', 'bibliografia', 'fuentes'],
    'Actualizado por': [],
    'Fecha actualización': [],
    'Autorizado por': [],
}

FIELD_ALIASES_HORMONAS = {
    'Código': [],
    'Nombre LIS': [],
    'Laboratorio de procesamiento': [],
    'Recepción de muestras': [],
    'Procesamiento de muestras': [],
    'Plazo entrega resultados': ['tiempo de respuesta', 'plazo entrega resultados', 'plazo de entrega', 'tat'],
    'Preparación del paciente': ['preparacion paciente', 'preparación paciente', 'hora de toma de muestra', 'ritmo circadiano', 'ciclo menstrual'],
    'Procedimiento de toma de muestras': ['Procedimiento', 'Procedimiento de toma de muestras', 'puncion venosa'],
    'Muestra requerida y recipiente': ['tipo de muestra', 'muestra requerida', 'suero', 'plasma edta', 'tubo suero con separador'],
    'Rotulación': ['Rotulación', 'Rotulacion'],
    'Estabilidad de la muestra': ['estabilidad muestra', 'estabilidad de la muestra', 'estabilidad hormonas', 'congelación suero'],
    'Condiciones de envío al Laboratorio': ['Condiciones de envío al Laboratorio', 'temperatura de transporte', 'envio congelado', 'cadena de frio'],
    'Criterios de rechazo': ['Rechazo', 'Criterios de rechazo', 'hemolisis', 'descongelamiento previo'],
    'Método de procesamiento': ['metodo utilizado', 'método utilizado', 'metodo de procesamiento', 'metodologia', 'quimioluminiscencia', 'eclia', 'elisa', 'ria', 'electroquimioluminiscencia'],
    'Test incluidos': ['test incluidos', 'pruebas incluidas', 'perfil tiroideo', 'panel hormonal'],
    'Valores de referencia': ['valor de referencia', 'valores de referencia', 'valores por trimestre', 'valores por fase menstrual'],
    'Valor crítico': [],
    'Límite de detección': ['limite de deteccion', 'límite de detección', 'sensibilidad funcional'],
    'Equipamiento utilizado': ['equipamiento utilizado', 'equipo utilizado', 'analizador de inmunoensayo', 'cobas e', 'centaur', 'immulite'],
    'Información clínica': ['caracteristicas', 'características', 'utilidad clinica', 'evaluación endocrina'],
    'Referencias': ['referencias', 'bibliografia', 'fuentes'],
    'Actualizado por': [],
    'Fecha actualización': [],
    'Autorizado por': [],
}

FIELD_ALIASES_DERIVADOS = {
    'Código': [],
    'Nombre LIS': [],
    'Laboratorio de procesamiento': [],
    'Recepción de muestras': [],
    'Procesamiento de muestras': [],
    'Plazo entrega resultados': ['tiempo de respuesta', 'plazo entrega resultados', 'plazo de entrega', 'tat derivación', 'dias hábiles'],
    'Preparación del paciente': ['preparacion paciente', 'preparación paciente', 'indicaciones laboratorio externo'],
    'Procedimiento de toma de muestras': ['Procedimiento', 'Procedimiento de toma de muestras', 'protocolo de derivación'],
    'Muestra requerida y recipiente': ['tipo de muestra', 'muestra requerida', 'alícuota', 'volumen minimo enviado'],
    'Rotulación': ['Rotulación', 'Rotulacion', 'etiqueta derivación'],
    'Estabilidad de la muestra': ['estabilidad muestra', 'estabilidad de la muestra', 'estabilidad en viaje'],
    'Condiciones de envío al Laboratorio': ['Condiciones de envío al Laboratorio', 'temperatura de transporte', 'transporte refrigerado', 'transporte congelado', 'nieve seca'],
    'Criterios de rechazo': ['Rechazo', 'Criterios de rechazo', 'incumplimiento cadena de frio', 'volumen insuficiente'],
    'Método de procesamiento': ['metodo utilizado', 'método utilizado', 'metodo de procesamiento', 'metodologia', 'espectrometria de masas', 'pcr', 'secuenciacion'],
    'Test incluidos': ['test incluidos', 'pruebas incluidas', 'panel genético', 'screening de derivación'],
    'Valores de referencia': ['valor de referencia', 'valores de referencia', 'rangos laboratorio de referencia'],
    'Valor crítico': [],
    'Límite de detección': ['limite de deteccion', 'límite de detección'],
    'Equipamiento utilizado': ['equipamiento utilizado', 'equipo utilizado', 'tecnología laboratorio de origen'],
    'Información clínica': ['caracteristicas', 'características', 'utilidad clinica', 'diagnóstico genético / derivado'],
    'Referencias': ['referencias', 'bibliografia', 'fuentes'],
    'Actualizado por': [],
    'Fecha actualización': [],
    'Autorizado por': [],
}

FIELD_ALIASES_POR_AREA = {
    "Microbiología": FIELD_ALIASES_MICROBIOLOGIA,
    "Hematología": FIELD_ALIASES_HEMATOLOGIA,
    "Hematología y coagulación": FIELD_ALIASES_HEMATOLOGIA,
    "Bioquímica": FIELD_ALIASES_BIOQUIMICA,
    "Hormonas e inmunidad": FIELD_ALIASES_HORMONAS,
    "Derivados": FIELD_ALIASES_DERIVADOS,
}

SECCIONES_DOCUMENTO_MAESTRAS = [
    'uso previsto', 'caracteristicas', 'características',
    'reactivos - soluciones de trabajo', 'reactivos', 'medidas de precaucion y advertencias',
    'medidas de precaucion', 'medidas de prevención', 'preparacion de los reactivos',
    'preparación de los reactivos', 'obtencion y preparacion de las muestras',
    'obtención y preparación de las muestras',
    'material suministrado', 'material requerido adicionalmente', 'material requerido',
    'realizacion del test', 'realización del test', 'calibracion', 'calibración',
    'control de calidad', 'limitaciones del analisis - interferencias', 'limitaciones del analisis',
    'limitaciones del análisis', 'limites e intervalos', 'límites e intervalos', 'valores teoricos',
    'valores teóricos', 'datos especificos de funcionamiento', 'datos específicos de funcionamiento',
    'precision', 'precisión', 'comparacion de metodos', 'comparación de métodos',
    'limite de cuantificacion', 'límite de cuantificación', 'loq',
    'referencias bibliograficas', 'referencias bibliográficas', 'simbolos', 'símbolos'
]


def get_field_aliases(area=None):
    if area and area in FIELD_ALIASES_POR_AREA:
        return FIELD_ALIASES_POR_AREA[area]
    return FIELD_ALIASES_BASE


def norm(s):
    import unicodedata
    s = str(s or '').strip().lower()
    s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    s = re.sub(r'\s*[:：]+\s*$', '', s)
    s = re.sub(r'\s+', ' ', s)
    return s


def nombre_campo_visible(nombre):
    """Retira dos puntos al final de un nombre de campo, no del valor."""
    return re.sub(r"\s*[:：]+\s*$", "", str(nombre or "").strip())


def normalizar_etiquetas_ficha(ws):
    """Presentación uniforme para la hoja del examen, sin alterar B ni SQLite."""
    for fila in range(1, ws.max_row + 1):
        celda = ws.cell(fila, 1)
        if isinstance(celda.value, str):
            celda.value = nombre_campo_visible(celda.value)



def _superindices_visuales_en_pagina(pagina):
    """Recupera como Unicode las citas voladas que PyMuPDF lee como números.

    Para no alterar el flujo de columnas ni todas las líneas del documento,
    reconstruye únicamente los bloques con spans marcados como superíndice.
    Los exponentes legítimos siguen intactos gracias al filtro posterior.
    """
    marcados = {}
    tabla = str.maketrans("0123456789", "⁰¹²³⁴⁵⁶⁷⁸⁹")
    for bloque in pagina.get_text("dict").get("blocks", []):
        if not bloque.get("lines"):
            continue
        modificado = False
        lineas = []
        for linea in bloque["lines"]:
            fragmentos = []
            for span in linea.get("spans", []):
                texto_span = span.get("text", "")
                if span.get("flags", 0) & fitz.TEXT_FONT_SUPERSCRIPT:
                    nuevo = re.sub(r"(?<!\d)\d{1,3}(?!\d)",
                                   lambda m: m.group(0).translate(tabla), texto_span)
                    if nuevo != texto_span:
                        texto_span = nuevo
                        modificado = True
                fragmentos.append(texto_span)
            lineas.append("".join(fragmentos))
        if modificado:
            marcados[bloque["number"]] = "\n".join(lineas)
    return marcados



def _membrete_clinico_hematologia(pagina, numero_pagina):
    """Identifica por POSICIÓN las líneas repetidas de cabecera de una página.

    Se descartan únicamente en la extracción clínica de Hematología: las
    referencias bibliográficas abren de nuevo los bytes originales del PDF.
    Un encabezado solo se reconoce cuando parece un título de portada en la
    zona superior de una página posterior; los títulos de secciones científicas
    y el contenido inferior de tablas nunca se consideran membretes.
    """
    if numero_pagina <= 1:
        return set()
    alto = float(pagina.rect.height)
    ancho = float(pagina.rect.width)
    limite_superior = min(205.0, alto * 0.30)
    lineas = []
    for bloque in pagina.get_text("dict").get("blocks", []):
        if bloque.get("type") != 0:
            continue
        for linea in bloque.get("lines", []):
            spans = [x for x in linea.get("spans", []) if str(x.get("text", "")).strip()]
            if not spans:
                continue
            texto = "".join(x.get("text", "") for x in spans).strip()
            x0, y0, _, _ = linea.get("bbox", (ancho, alto, ancho, alto))
            if not texto or y0 > limite_superior:
                continue
            lineas.append((float(y0), float(x0), max(float(x.get("size", 0)) for x in spans), texto))
    lineas.sort(key=lambda dato: (dato[0], dato[1]))
    if not lineas:
        return set()

    # Códigos editoriales habituales (p. ej. ms_07153414190V5.0), no
    # identificadores analíticos ni resultados de laboratorio.
    patron_codigo = re.compile(r"(?i)^(?:ms[_-]?\d{7,}[a-z0-9._-]*|\d{9,}v\d+(?:\.\d+)+)$")
    codigos = [(y, x, texto) for y, x, _, texto in lineas if patron_codigo.fullmatch(texto)]
    descartadas = {texto for _, _, texto in codigos}
    # La tipografía grande ayuda a distinguir la cabecera del examen de
    # encabezados técnicos válidos como «Límites e intervalos».
    titulos_seccion = {norm(t) for t in SECCIONES_DOCUMENTO_MAESTRAS}
    titulos = [(y, x, texto) for y, x, size, texto in lineas
               if (size >= 17 and x < ancho * 0.65 and 3 <= len(texto) <= 90
                   and norm(texto) not in titulos_seccion
                   and not patron_codigo.fullmatch(texto)
                   and not re.match(r"(?i)^(?:cobas[®™]?|roche diagnostics)", texto))]
    if not titulos:
        return descartadas
    # Evita suprimir el título válido de una sección si no hay un membrete
    # próximo o un título inequívocamente destacado.
    for y, x, titulo in titulos:
        cerca_codigo = any(0 <= y - yc <= 75 for yc, _, _ in codigos)
        if not cerca_codigo and not (y <= min(125, alto * 0.18) and len(titulo) <= 55):
            continue
        descartadas.add(titulo)
        # Los subtítulos ingleses inmediatos («Prothrombin Time») son parte
        # del membrete, pero no una tabla ni otro encabezado técnico.
        for y2, x2, size2, texto2 in lineas:
            if (y < y2 <= y + 37 and abs(x2 - x) < 75
                    and 3 <= len(texto2) <= 75 and size2 <= 16
                    and norm(texto2) not in titulos_seccion
                    and not re.search(r"[:=|\d]|°|µ|μ|%", texto2)):
                descartadas.add(texto2)
        break
    return descartadas


def _limpiar_membrete_ocr_hematologia(texto, numero_pagina):
    """Respaldo prudente para OCR sin coordenadas: requiere un código ms_."""
    if numero_pagina <= 1:
        return texto
    lineas = str(texto).splitlines()
    codigo = re.compile(r"(?i)^ms[_-]?\d{7,}[a-z0-9._-]*$")
    for i, linea in enumerate(lineas[:9]):
        if not codigo.fullmatch(linea.strip()):
            continue
        lineas[i] = ""
        # La portada abreviada y su subtítulo siguen al código en OCR.
        for j in range(i+1, min(i+3, len(lineas))):
            frase = lineas[j].strip()
            if (frase and len(frase) < 70 and not re.search(r"[:=|\d]|°|µ|μ", frase)
                    and norm(frase) not in {norm(t) for t in SECCIONES_DOCUMENTO_MAESTRAS}):
                lineas[j] = ""
            else:
                break
        break
    return "\n".join(lineas)


def extract_pdf_text(uploaded_bytes, filename, examen_activo=None, hematologia=False, progreso=None,
                     rango_paginas=None, manual_procedimiento=False):
    """Extrae cada página independientemente; usa OCR solo donde falta texto.

    Mantiene el contrato anterior: (texto, descripción_del_método).
    Nunca inventa información cuando no es posible leer una página.
    """
    try:
        doc = fitz.open(stream=uploaded_bytes, filetype="pdf")
    except Exception as exc:
        raise RuntimeError(f"No se pudo abrir el PDF '{filename}': {exc}") from exc

    if not doc.page_count:
        doc.close()
        raise RuntimeError("El PDF no contiene páginas.")

    pages_text = []
    paginas_ocr = []
    texto_optico_original = []
    motores_ocr = set()
    engine = None
    try:
        desde, hasta = rango_paginas or (1, doc.page_count)
        desde = max(1, min(int(desde), doc.page_count))
        hasta = max(desde, min(int(hasta), doc.page_count))
        paginas_a_analizar = hasta - desde + 1
        if progreso is not None:
            progreso(0, paginas_a_analizar, False)
        for numero, page in enumerate(doc, start=1):
            if not desde <= numero <= hasta:
                continue
            page_height = page.rect.height
            page_width = page.rect.width
            valid_blocks = []
            superscripts = _superindices_visuales_en_pagina(page)
            # Solo Hematología: limpiar por posición el membrete de la página
            # siguiente sin eliminar la continuación de tablas o párrafos.
            lineas_membrete = (_membrete_clinico_hematologia(page, numero)
                               if hematologia else set())
            for b in page.get_text("blocks"):
                if b[6] != 0:
                    continue
                x0, y0, x1, y1 = b[:4]
                content = superscripts.get(b[5], b[4])
                if lineas_membrete:
                    content = "\n".join(
                        linea for linea in content.splitlines()
                        if linea.strip() not in lineas_membrete)
                content = content.strip()
                if not content:
                    continue
                if y0 < 50 or y1 > page_height - 50:
                    lines = [line.strip() for line in content.splitlines() if line.strip()]
                    if lines and all(es_linea_basura_o_encabezado(line, examen_activo) for line in lines):
                        continue
                if not es_linea_basura_o_encabezado(content, examen_activo):
                    valid_blocks.append((x0, y0, x1, y1, content))

            def sort_key(block):
                x0, y0, x1, _, _ = block
                column = 0 if x1 - x0 > page_width * 0.6 else (1 if x0 < page_width * 0.48 else 2)
                return (column, y0, x0)

            valid_blocks.sort(key=sort_key)
            page_text = "\n".join(b[4] for b in valid_blocks)

            # Un PDF mixto puede contener páginas con texto y páginas escaneadas.
            if manual_procedimiento and len(re.sub(r"\s", "", page_text)) < 50:
                try:
                    import numpy as np
                    # En documentos escaneados usar 160 dpi para reducir la
                    # rasterización. Evitar codificar y decodificar PNG en
                    # memoria: RapidOCR acepta directamente matrices RGB.
                    pix = page.get_pixmap(dpi=160 if manual_procedimiento else 200,
                                          alpha=False, colorspace=fitz.csRGB)
                    imagen_rgb = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
                        pix.height, pix.width, pix.n)
                    # Solo «Documentos escaneados» usa motores OCR.
                    ocr_text, motor_ocr = oe.leer_pagina_escaneada(
                        imagen_rgb, motor_rapid=engine)
                    if len(re.sub(r"\s", "", ocr_text)) > len(re.sub(r"\s", "", page_text)):
                        # Guardar exactamente lo reconocido antes de corregir
                        # palabras inequívocas; no alterar números ni modelos.
                        texto_optico_original.append((numero, ocr_text))
                        import lenguaje_ocr
                        page_text = lenguaje_ocr.corregir_texto_lectura_optica(ocr_text)
                        paginas_ocr.append(numero)
                        motores_ocr.add(motor_ocr)
                except ImportError:
                    # Mantener texto digital disponible; si todo es imagen, informar.
                    pass
                except Exception as exc:
                    # Una página defectuosa no debe descartar las demás.
                    import logging
                    logging.getLogger(__name__).warning("OCR falló en página %s: %s", numero, exc)
            pages_text.append(page_text)
            if progreso is not None:
                progreso(numero - desde + 1, paginas_a_analizar, numero in paginas_ocr)
    finally:
        doc.close()

    text = "\n\n".join(pages_text).strip()
    if manual_procedimiento and "st" in globals():
        st.session_state["texto_optico_original"] = {
            "archivo": filename, "paginas": tuple(texto_optico_original),
        }
    if len(re.sub(r"\s", "", text)) < 10:
        raise RuntimeError("No se detectó texto seleccionable. Si el PDF está escaneado, elige «Documentos escaneados» para utilizar lectura óptica.")
    if paginas_ocr:
        tipo_ocr = "/".join(sorted(motores_ocr)) if manual_procedimiento else "RapidOCR"
        method = "PyMuPDF + " + tipo_ocr + " (páginas: " + ", ".join(map(str, paginas_ocr)) + ")"
    else:
        method = "PyMuPDF (Columna por Columna)"
    return text, method


def clean_value(v):
    v = re.sub(r'[ \t]+', ' ', v).strip()
    return v.strip(' :\t')


def _limpiar_fallback_valores_teoricos(seccion, examen_activo=None):
    """Conserva todo el contenido útil y descarta encabezados/pies de páginas.

    Actúa solo en el texto de «Valores teóricos» y nunca sobre otros campos.
    No elimina unidades científicas ni expresiones matemáticas.
    """
    descartes = (
        r"p[aá]gina\s*[:.]?\s*\d{1,3}(?:\s*(?:de|/)\s*\d{1,3})?",
        r"\d{1,3}\s*(?:de|/)\s*\d{1,3}",
        r"(?:c[oó]digo|cod\.?\s*caract\.?|versi[oó]n|vigencia|fecha)\s*[:：].*",
        r"(?:manual\s+de\s+toma\s+de\s+muestras|laboratorio\s+cl[ií]nico\s+del\s+centro\s+de\s+referencia).{0,120}",
        r"cl[ií]nico\s+manual\s+del\s+de\s+centro.*",
        r"pe[nñ]alol[eé]n\s+cordillera\s+oriente.*",
        r"(?:roche diagnostics|cobas[®]?|www\.\S+|https?://\S+)",
        r"(?:[—–_\-=]{3,}|\|[\s\|]+)",
    )
    salida = []
    lineas_originales = str(seccion).replace("\r", "").split("\n")
    for pos, linea in enumerate(lineas_originales):
        limpia = re.sub(r"[ \t\u00a0]+", " ", linea).strip()
        if not limpia:
            if salida and salida[-1] != "":
                salida.append("")
            continue
        if es_linea_basura_o_encabezado(limpia, examen_activo):
            continue
        if any(re.fullmatch(patron, limpia, re.I) for patron in descartes):
            continue
        # Las cifras aisladas pueden ser valores clínicos. Se eliminan como
        # número de página solo al acompañar a un membrete reconocible.
        if re.fullmatch(r"\d{1,3}", limpia):
            vecinos = " ".join(lineas_originales[max(0, pos-2):pos+4]).casefold()
            if re.search(r"\b(?:c[oó]digo|cod\.?\s*caract|p[aá]gina\s*[:：]|vigencia\s*[:：])", vecinos):
                continue
        # Citas entre palabras o tras puntuación, incluso si el PDF aplana
        # los superíndices a dígitos normales (p. ej., «).29» o «texto²⁹»).
        limpia = quitar_citas_superindices(limpia)
        salida.append(limpia)
    while salida and not salida[-1]:
        salida.pop()
    # Unir líneas físicas de un mismo párrafo; mantener separación entre párrafos.
    parrafos = []
    actual = []
    for linea in salida + [""]:
        if not linea:
            if actual:
                parrafos.append(" ".join(actual).strip())
                actual = []
        else:
            actual.append(linea)
    return "\n\n".join(parrafos).strip()


def extraer_valores_teoricos_hematologia(texto):
    """Prioriza cifras o intervalos en mg/dL; si faltan, copia la sección completa.

    El contenido debe estar delimitado por «Valores teóricos» y el siguiente
    encabezado comprobable. No se copia el resto del PDF sin límite identificado.
    """
    if not texto:
        return ""
    # Translación 1:1 para conservar exactamente los índices del texto original.
    tabla = str.maketrans("áéíóúÁÉÍÓÚ", "aeiouAEIOU")
    original = str(texto)
    normalizado = original.translate(tabla)
    inicio = re.search(r"\bvalores\s+teoricos\b\s*[:：]?", normalizado, re.I)
    if not inicio:
        return ""
    resto = normalizado[inicio.end():]
    limites = []
    # Encabezado que sigue habitualmente a «Valores teóricos».
    fin_principal = re.search(
        r"\bdatos\s+especificos\s+de\s+funcionamiento\s+del\s+test\b",
        resto, re.I)
    if fin_principal:
        limites.append(fin_principal.start())
    # Si otro encabezado maestro aparece antes, no incluirlo ni incluir su cuerpo.
    for encabezado in SECCIONES_DOCUMENTO_MAESTRAS:
        encabezado_norm = encabezado.translate(tabla)
        if encabezado_norm.casefold() in ("valores teoricos", "valores teorico"):
            continue
        partes = [re.escape(x) for x in encabezado_norm.split()]
        patron = r"(?m)^\s*" + r"\s+".join(partes) + r"\b(?:\s*[:：])?\s*(?=\n|$)"
        for titulo in re.finditer(patron, resto, re.I):
            limites.append(titulo.start())
    if not limites:
        return ""
    seccion = original[inicio.end():inicio.end() + min(limites)]
    numero = r"\d+(?:[.,]\d+)?"
    patron = re.compile(
        rf"(?<![\w.,])(?:[<>≤≥]\s*)?{numero}"
        rf"(?:\s*(?:[-‐‑‒–—―−﹣－]|\ba\b|\bhasta\b)\s*{numero})?"
        r"\s*mg\s*/\s*d[lL]\b", re.I
    )
    valores = []
    vistos = set()
    for coincidencia in patron.finditer(seccion):
        valor = re.sub(r"\s+", " ", coincidencia.group()).strip()
        clave = valor.casefold()
        if clave not in vistos:
            vistos.add(clave)
            valores.append(valor)
    if valores:
        return " | ".join(valores)
    return _limpiar_fallback_valores_teoricos(seccion)


def es_alias_hematologia(field_aliases):
    return field_aliases is FIELD_ALIASES_HEMATOLOGIA or (
        isinstance(field_aliases, dict)
        and field_aliases.get("Valores de referencia") == FIELD_ALIASES_HEMATOLOGIA["Valores de referencia"]
    )


def extraer_limite_deteccion_hematologia(texto):
    """Extrae el valor explícito solo del apartado «Límites e intervalos».

    Los insertos Roche pueden expresar la unidad como «µg de UEF/mL» o
    «UEF/mL», y separar la etiqueta, el signo igual y el valor en líneas
    distintas. Nunca se obtiene el dato de linealidad, cuantificación u otras
    secciones del PDF.
    """
    if not texto:
        return ""
    texto = str(texto).replace("\r\n", "\n").replace("\r", "\n")
    lineas = texto.split("\n")
    titulo = re.compile(r"^l[ií]mites\s+e\s+intervalos\s*[:：]?\s*$", re.I)
    etiqueta = r"l[ií]mite\s+de\s+detecci[oó]n"
    numero = r"(?:[<>≤≥≈]\s*)?\d+(?:[.,]\d+)?"
    unidades_basicas = r"(?:UEF|UI|IU|mIU|mU|U|ng|pg|µg|μg|ug|mcg|mg|g|copias)"
    unidad_simple = rf"{unidades_basicas}\s*/\s*(?:mL|dL|L)"
    unidad_compuesta = rf"(?:µg|μg|ug|mcg|ng|pg|mg|g)\s+de\s+{unidad_simple}"
    patron = re.compile(
        rf"{etiqueta}\s*[=:：]\s*(?P<valor>{numero}\s*(?:{unidad_compuesta}|{unidad_simple}))\b",
        re.I,
    )
    # Los títulos de las subsecciones «Intervalo de medición» y «Límites
    # inferiores de medición» pertenecen a este apartado; no lo cierran.
    titulos_que_cierran = {
        norm(encabezado) for encabezado in SECCIONES_DOCUMENTO_MAESTRAS
        if norm(encabezado) != norm("Límites e intervalos")
    }
    titulos_que_cierran.update({
        "valores esperados", "referencias", "especificidad", "interferencias",
        "exactitud", "repetibilidad", "imprecision", "rendimiento analitico",
        "rendimiento del test", "correlacion", "interpretacion de resultados",
        "precision", "valores teoricos", "caracteristicas", "uso previsto",
        "control de calidad", "comparacion de metodos", "referencias bibliograficas",
    })
    for i, linea in enumerate(lineas):
        if not titulo.fullmatch(linea.strip()):
            continue
        seccion = []
        # Las medidas pertenecen a la misma sección; un título de otra sección
        # cierra la búsqueda. El límite también evita recorrer todo el PDF.
        for siguiente in lineas[i + 1: i + 65]:
            if norm(siguiente) in titulos_que_cierran:
                break
            seccion.append(siguiente)
            if sum(len(x) for x in seccion) > 4500:
                break
        coincidencia = patron.search("\n".join(seccion))
        if coincidencia:
            return re.sub(r"\s+", " ", coincidencia.group("valor")).strip()
    return ""


def parse_labeled_text(text, field_aliases=None, examen_activo=None):
    if field_aliases is None:
        field_aliases = FIELD_ALIASES_BASE

    lines = [x.rstrip() for x in text.replace('\r', '').split('\n')]
    found = {}
    current = None

    label_patterns = []
    for field, aliases in field_aliases.items():
        if field == "Código":
            continue
        for a in aliases:
            label_patterns.append((field, a))
    label_patterns.sort(key=lambda x: len(x[1]), reverse=True)

    norm_maestras = [norm(sec) for sec in SECCIONES_DOCUMENTO_MAESTRAS]

    for raw in lines:
        s = raw.strip()
        if not s or es_linea_basura_o_encabezado(s, examen_activo):
            continue

        ns = norm(s)
        matched_field = None
        remainder = ''

        for field, alias in label_patterns:
            na = norm(alias)
            if ns == na:
                matched_field = field
                remainder = ''
                break
            # Buscar separador después de la etiqueta, sin depender de
            # la longitud original de la variante acentuada.
            match = re.match(r"^(.+?)\s*[:：–—-]\s*(.*)$", s)
            if match and norm(match.group(1)) == na:
                matched_field = field
                remainder = match.group(2).strip()
                break

        if matched_field:
            current = matched_field
            if remainder:
                found[current] = (found.get(current, '') + ('\n' if found.get(current) else '') + clean_value(remainder)).strip()
            else:
                found.setdefault(current, '')
            continue

        is_main_doc_section = False
        for sec in norm_maestras:
            if ns == sec or ns.startswith(sec + ':'):
                is_main_doc_section = True
                break

        if is_main_doc_section:
            current = None
            continue

        if current:
            found[current] = (found.get(current, '') + ('\n' if found.get(current) else '') + s).strip()

    for k, v in list(found.items()):
        val_limpio = limpiar_valor_excel(v, examen_activo)
        if k == 'Valores de referencia' and not es_alias_hematologia(field_aliases):
            val_limpio = extraer_solo_valor_referencia(val_limpio)
        found[k] = val_limpio

    if es_alias_hematologia(field_aliases):
        # Esta medida solo se toma cuando el PDF incluye el encabezado exacto;
        # nunca se infiere a partir de «Rango de linealidad» u otros valores.
        limite = extraer_limite_deteccion_hematologia(text)
        if limite:
            found["Límite de detección"] = limite
        else:
            found.pop("Límite de detección", None)
        seccion = extraer_valores_teoricos_hematologia(text)
        if seccion:
            found['Valores de referencia'] = seccion
        else:
            # No usar el texto parcial de un apartado sin límite final verificado.
            found.pop('Valores de referencia', None)

    return found


def infer_name_code(text, data):
    equipos_encontrados = re.findall(r'\bcobas[\s-]*(?:[a-zA-Z]\s*-?\s*)?\d{3,4}\b|\bcobas\s+(?:pure|pro)\b', text, flags=re.I)
    if equipos_encontrados and not data.get('Equipamiento utilizado'):
        equipos_limpios = []
        for nombre_equipo in equipos_encontrados:
            equipo_normalizado = re.sub(r'\s+', ' ', nombre_equipo).strip()
            if eq.clave_equipo(equipo_normalizado) not in [eq.clave_equipo(x) for x in equipos_limpios]:
                equipos_limpios.append(equipo_normalizado)
        data['Equipamiento utilizado'] = "\n".join(equipos_limpios)

    if data.get('Equipamiento utilizado'):
        data['Equipamiento utilizado'] = eq.deduplicar_equipamiento(
            data['Equipamiento utilizado'])

    if not data.get('Método de procesamiento'):
        patrones_metodo = [
            r'\bm[eé]todo\s+de\s+[a-zA-ZáéíóúñÁÉÍÓÚÑ\s]{3,35}(?=\.|\n|,|$)',
            r'\b(?:coagulom[eé]tric[oa]|turbidim[eé]tric[oa]|fotom[eé]tric[oa]|enzim[aá]tic[oa]|espectrofotom[eé]tric[oa]|quimioluminiscencia|ECLIA|ELISA|RIA|inmunoensayo|nefelom[eé]tric[oa]|cromog[eé]nic[oa]|aglutinaci[oó]n|impedancia|citometr[ií]a\s+de\s+flujo|colorim[eé]tric[oa]|punto\s+final|cin[eé]tic[oa])\b'
        ]
        metodos_encontrados = []
        for pat in patrones_metodo:
            coincidencias = re.findall(pat, text, flags=re.I)
            for c in coincidencias:
                c_clean = re.sub(r'\s+', ' ', c).strip()
                if c_clean.casefold() not in [x.casefold() for x in metodos_encontrados]:
                    metodos_encontrados.append(c_clean)
        if metodos_encontrados:
            data['Método de procesamiento'] = "\n".join(metodos_encontrados[:3])

    return data


def apply_to_sheet(ws, data):
    rows = {}
    for r in range(1, ws.max_row + 1):
        label = str(ws.cell(r, 1).value or '').strip()
        if label:
            rows[norm(label)] = r
    updates = []
    for field, value in data.items():
        if not value or str(value).strip().casefold() in ["nan", "none", "null"]:
            continue
        target = field
        r = rows.get(norm(target))
        if r is None:
            for existing, rr in rows.items():
                if existing == norm(target) or norm(target) in existing or existing in norm(target):
                    r = rr
                    break
        if r:
            old = ws.cell(r, 2).value
            celda_destino = ws.cell(r, 2)
            celda_destino.value = value
            if isinstance(value, str):
                celda_destino.data_type = "s"  # El contenido clínico nunca es una fórmula Excel.
            celda_destino.alignment = Alignment(wrap_text=True, vertical='top')
            updates.append((field, old, value, r))
    return updates


def crear_hoja_examen(wb, nombre_examen, field_aliases=None):
    if field_aliases is None:
        field_aliases = FIELD_ALIASES_BASE
    ok, mensaje = nombre_examen_valido(nombre_examen)
    if not ok:
        raise ValueError(mensaje)
    existente = buscar_hoja_por_examen(wb, nombre_examen)
    if existente is not None:
        return existente
    ws = wb.create_sheet(title=titulo_hoja_para_examen(wb, nombre_examen))
    for fila, campo in enumerate(field_aliases, start=1):
        ws.cell(row=fila, column=1, value=campo)
        ws.cell(row=fila, column=2, value=nombre_examen if norm(campo) == norm("Nombre LIS") else "")
    return ws


def consolidar_documentos_pdf(documentos, field_aliases, nombre_examen,
                             manual_procedimiento=False):
    """Extrae cada PDF por separado y completa únicamente los campos vacíos.

    Las diferencias entre valores no se resuelven automáticamente: se devuelven
    para revisión del personal de laboratorio.
    """
    if not documentos:
        raise ValueError("Debes seleccionar al menos un documento PDF.")
    datos = {}
    fuentes = {}
    discrepancias = {}
    textos = []
    metodos = []
    for nombre, contenido in documentos:
        texto, metodo = extract_pdf_text(contenido, nombre, nombre_examen,
                                          hematologia=es_alias_hematologia(field_aliases),
                                          manual_procedimiento=manual_procedimiento)
        encontrados = (mp.extraer_campos_manual(texto) if manual_procedimiento else
                       mp.completar_homologos_manual_digital(
                           texto, infer_name_code(
                               texto, parse_labeled_text(texto, field_aliases, nombre_examen))))
        if (field_aliases is globals().get("FIELD_ALIASES_DERIVADOS") or field_aliases == globals().get("FIELD_ALIASES_DERIVADOS")):
            encontrados.update(de.extraer_ficha_derivados(texto, contenido))
        permitidos_manual = (set(fd.CAMPOS_CONTROLADOS) if
                             (manual_procedimiento or (field_aliases is globals().get("FIELD_ALIASES_DERIVADOS") or field_aliases == globals().get("FIELD_ALIASES_DERIVADOS"))) else
                             {'Muestra requerida y recipiente'}
                             if mp.es_manual_de_procedimiento_textual(texto) else set())
        encontrados = {campo: val for campo, val in encontrados.items()
                       if campo not in fd.CAMPOS_MANUAL
                       and (campo not in fd.CAMPOS_CONTROLADOS or campo in permitidos_manual)}
        textos.append(f"===== {nombre} =====\n{texto}")
        metodos.append(f"{nombre}: {metodo}")
        for campo, valor in encontrados.items():
            valor = str(valor or "").strip()
            if campo == "Equipamiento utilizado":
                valor = eq.deduplicar_equipamiento(valor)
            if not valor or valor.casefold() in ("nan", "none", "null"):
                continue
            if not str(datos.get(campo, "")).strip():
                datos[campo] = valor
                fuentes[campo] = nombre
            elif (not eq.equipamientos_equivalentes(datos[campo], valor)
                  if campo == "Equipamiento utilizado"
                  else norm(str(datos[campo])) != norm(valor) and valor != str(datos[campo])):
                opciones = discrepancias.setdefault(campo, [(fuentes[campo], str(datos[campo]))])
                if (nombre, valor) not in opciones:
                    opciones.append((nombre, valor))
    for campo_manual in fd.CAMPOS_MANUAL:
        datos.pop(campo_manual, None)
        fuentes.pop(campo_manual, None)
        discrepancias.pop(campo_manual, None)
    return datos, fuentes, discrepancias, "\n\n".join(textos), " | ".join(metodos)


def datos_para_ficha_sin_pdf(info_cartera, valores_documentales):
    """Compone únicamente datos respaldados por las fuentes institucionales.

    Las fuentes documentales ya están ordenadas por su política de prioridad;
    no se infieren valores desde una ficha anterior ni desde un PDF.
    """
    datos = {}
    for campo, valor in (info_cartera or {}).items():
        texto = str(valor or "").strip()
        if texto and texto.casefold() not in {"nan", "none", "null"}:
            datos[campo] = texto
    for campo, valor in (valores_documentales or {}).items():
        texto = str(valor or "").strip()
        if texto and texto.casefold() not in {"nan", "none", "null"}:
            # Las Normas y Carteras se consultan con su prioridad habitual.
            datos[campo] = texto
    # Regla operativa: estos campos se completan exclusivamente por el encargado.
    for campo_manual in ("Recepción de muestras", "Procesamiento de muestras", "Valor crítico"):
        datos.pop(campo_manual, None)
    return datos



def homologacion_del_examen(nombre_examen, area, info_cartera=None, nombre_cartera=None):
    """Consulta por identidad institucional, sin inferir sinónimos clínicos."""
    nombre_vinculado = nombre_cartera if nombre_cartera and nombre_cartera != "Sin vincular" else nombre_examen
    return htest.obtener_tests(
        DB_PATH, examen=nombre_vinculado, area=area,
        codigo=(info_cartera or {}).get("Código", ""))


def process(uploaded_bytes, filename, template_bytes, centro='', nombre_examen='', field_aliases=None,
            actualizado_por='', autorizado_por='', datos_sobrescribir=None, documentos_adicionales=None,
            datos_pdf_preextraidos=None, manual_procedimiento=False):
    # Importación local: process también se ejecuta de manera aislada en las
    # comprobaciones y reutiliza el módulo ya cargado por la aplicación.
    import manuales_procedimiento as mp
    campos_solo_manual = ("Recepción de muestras", "Procesamiento de muestras", "Valor crítico")
    if field_aliases is None:
        field_aliases = FIELD_ALIASES_BASE
    if not nombre_examen:
        raise ValueError('Selecciona un examen desde el menú del área antes de analizar el PDF.')
    ok, mensaje = nombre_examen_valido(nombre_examen)
    if not ok:
        raise ValueError(mensaje)

    # La interfaz ya leyó los PDF para detectar discrepancias y fuentes.
    # Reutilizar sus resultados evita ejecutar OCR por duplicado en documentos
    # escaneados (en particular, manuales con numerosas páginas).
    if datos_pdf_preextraidos is not None:
        data, text, method = datos_pdf_preextraidos
        data = dict(data)  # No modificar el resultado original de la revisión.
    elif documentos_adicionales:
        data, _, _, text, method = consolidar_documentos_pdf(
            [(filename, uploaded_bytes)] + documentos_adicionales, field_aliases, nombre_examen,
            manual_procedimiento=manual_procedimiento
        )
    else:
        text, method = extract_pdf_text(uploaded_bytes, filename, nombre_examen,
                                        hematologia=es_alias_hematologia(field_aliases),
                                        manual_procedimiento=manual_procedimiento)
        data = (mp.extraer_campos_manual(text) if manual_procedimiento else
                mp.completar_homologos_manual_digital(
                    text, infer_name_code(
                        text, parse_labeled_text(text, field_aliases, nombre_examen))))

    if (field_aliases is globals().get("FIELD_ALIASES_DERIVADOS") or field_aliases == globals().get("FIELD_ALIASES_DERIVADOS")) and datos_pdf_preextraidos is None:
        data.update(de.extraer_ficha_derivados(text, uploaded_bytes))

    # Ningún OCR ni extractor técnico rellena «Referencias» automáticamente.
    # Su aplicación es exclusivamente una acción explícita del usuario.
    data.pop("Referencias", None)
    for campo_manual in ("Recepción de muestras", "Procesamiento de muestras", "Valor crítico"):
        data.pop(campo_manual, None)
    # Recepción, Procesamiento de muestras y Valor crítico son exclusivos del encargado.
    manual_digital = (not manual_procedimiento and
                      mp.es_manual_de_procedimiento_textual(text))
    if manual_digital:
        data = mp.completar_homologos_manual_digital(text, data)
    manuales_permitidos = (set(fd.CAMPOS_CONTROLADOS) if
                          (manual_procedimiento or (field_aliases is globals().get("FIELD_ALIASES_DERIVADOS") or field_aliases == globals().get("FIELD_ALIASES_DERIVADOS"))) else
                          {'Muestra requerida y recipiente'}
                          if manual_digital else set())
    data = {campo: val for campo, val in data.items()
            if campo not in campos_solo_manual
            and (campo not in fd.CAMPOS_CONTROLADOS or campo in manuales_permitidos)}
    # Los campos entregados por el PDF son una propuesta, no una orden de
    # sobrescritura: conservar su identidad incluso si el texto fue leído
    # fuera de este proceso (OCR reutilizado o varios PDF).
    campos_provenientes_pdf = set(data) if (
        manual_procedimiento or manual_digital or (field_aliases is globals().get("FIELD_ALIASES_DERIVADOS") or field_aliases == globals().get("FIELD_ALIASES_DERIVADOS"))) else set()
    if datos_sobrescribir:
        for campo, val in datos_sobrescribir.items():
            if campo in ("Recepción de muestras", "Procesamiento de muestras", "Valor crítico"):
                continue
            if val and str(val).strip().casefold() not in ["nan", "none", "null"]:
                data[campo] = val

    ZONA_HORARIA = ZoneInfo("America/Santiago")
    fecha_actual = datetime.now(ZONA_HORARIA).strftime('%d/%m/%Y %H:%M')

    data['Nombre LIS'] = nombre_examen.strip()
    data['Laboratorio de procesamiento'] = (centro or '').strip()
    data['Actualizado por'] = (actualizado_por or '').strip()
    data['Autorizado por'] = (autorizado_por or '').strip()
    data['Fecha actualización'] = fecha_actual

    if ('Valores de referencia' in data and data['Valores de referencia']
            and not es_alias_hematologia(field_aliases)):
        data['Valores de referencia'] = extraer_solo_valor_referencia(data['Valores de referencia'])

    wb = load_workbook(io.BytesIO(template_bytes), rich_text=True)

    ws_seleccionada = buscar_hoja_por_examen(wb, nombre_examen)
    if ws_seleccionada is not None:
        ws = ws_seleccionada
        score, reason = 1.0, 'Examen seleccionado por el usuario'
    elif len(wb.sheetnames) == 1 and norm(wb.sheetnames[0]) in ['sheet', 'sheet1', 'hoja1', 'hoja 1']:
        ws = wb.active
        ws.title = titulo_hoja_para_examen(wb, nombre_examen.strip())
        for fila, campo in enumerate(field_aliases, start=1):
            ws.cell(row=fila, column=1, value=campo)
            ws.cell(row=fila, column=2, value=nombre_examen.strip() if norm(campo) == norm("Nombre LIS") else "")
        score, reason = 1.0, 'Ficha renombrada desde el examen seleccionado'
    else:
        ws = crear_hoja_examen(wb, nombre_examen, field_aliases)
        score, reason = 1.0, 'Ficha nueva creada desde el examen seleccionado'

    normalizar_etiquetas_ficha(ws)

    if manual_procedimiento or manual_digital or (field_aliases is globals().get("FIELD_ALIASES_DERIVADOS") or field_aliases == globals().get("FIELD_ALIASES_DERIVADOS")):
        # No alterar ningún valor previamente validado en una ficha. Las
        # fuentes institucionales siguen operando por sus propias reglas.
        datos_documentales = set((datos_sobrescribir or {}).keys())
        existentes = {str(ws.cell(r, 1).value or '').strip(): ws.cell(r, 2).value
                      for r in range(1, ws.max_row + 1)}
        candidatos_pdf = {campo: valor for campo, valor in data.items()
                          if campo in campos_provenientes_pdf
                          and campo not in datos_documentales}
        # No ocultar hallazgos por existir información antigua. El usuario
        # decide qué conservar con «Aplicar»; la Cartera seleccionada se compara.
        data = {campo: valor for campo, valor in data.items()
                if campo not in campos_provenientes_pdf
                or campo in datos_documentales}
        data.update(candidatos_pdf)

    # Fuentes institucionales propuestas para este análisis. Los datos antiguos
    # permanecen como base de comparación, sin aparecer como recién encontrados.
    for fila in range(1, ws.max_row + 1):
        etiqueta = str(ws.cell(fila, 1).value or "").strip()
        # Los campos controlados se calculan en el análisis actual y no deben
        # trasladarse desde una extracción histórica sin un nuevo origen.
        if etiqueta in fd.CAMPOS_CONTROLADOS:
            ws.cell(fila, 2).value = None
    updates = apply_to_sheet(ws, data)

    filas = {norm(str(ws.cell(r, 1).value or '')): r for r in range(1, ws.max_row + 1)}
    for campo, val in [
        ('Nombre LIS', nombre_examen.strip()),
        ('Laboratorio de procesamiento', (centro or '').strip()),
        ('Actualizado por', (actualizado_por or '').strip()),
        ('Autorizado por', (autorizado_por or '').strip()),
        ('Fecha actualización', fecha_actual),
    ]:
        if norm(campo) in filas:
            celda = ws.cell(filas[norm(campo)], 2)
            if not (manual_procedimiento and celda.value is not None
                    and str(celda.value).strip()):
                celda.value = val

    # La tabla solo recibe resultados de ESTE análisis. Los valores anteriores
    # permanecen en el libro/plantilla, pero nunca se hacen pasar por extracción.
    # AG Grid los muestra aparte al desmarcar «Aplicar».

    # Las viñetas documentales ya fueron reconstruidas al consultar SQLite.
    # Su formato multilineal nunca debe pasar de nuevo por limpieza genérica.
    campos_con_vinetas = {
        norm(campo) for campo, valor in (datos_sobrescribir or {}).items()
        if isinstance(valor, str) and any(
            linea.lstrip().startswith(("• ", "● ", "▪ ", "◦ ", "∙ "))
            for linea in valor.splitlines()
        )
    }
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None and isinstance(cell.value, str):
                etiqueta = norm(ws.cell(cell.row, 1).value) if cell.column == 2 else ""
                if etiqueta in campos_con_vinetas:
                    cell.value = cell.value.replace("\r\n", "\n").replace("\r", "\n")
                else:
                    cell.value = limpiar_valor_excel(cell.value, nombre_examen)
            # Evita que una indicación que empieza por "=" genere <f> en XML.
            if cell.column == 2:
                cell.data_type = "s"
            cell.alignment = Alignment(vertical='top', wrap_text=True)
    # El nombre completo de la prestación es un identificador, no texto clínico:
    # evitar que la normalización de subíndices cambie el Nombre LIS.
    if norm('Nombre LIS') in filas:
        ws.cell(filas[norm('Nombre LIS')], 2).value = nombre_examen.strip()
        ws.cell(filas[norm('Nombre LIS')], 2).data_type = "s"
    # Equivalente Python de preajuste del ancho: Excel (.xlsx) no ejecuta VBA.
    # Recalcular las columnas de TODAS las hojas, incluidas las aún sin editar.
    for hoja_del_libro in wb.worksheets:
        ajustar_ancho_columnas(hoja_del_libro)
    for fila in range(1, ws.max_row + 1):
        ajustar_alto_fila_con_listas(ws, fila)

    wb.active = ws

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue(), data, nombre_examen, score, reason, method, text


def actualizar_campo_excel(excel_bytes, nombre_examen, campo, valor):
    """Actualiza un único campo de la ficha sin modificar las demás hojas."""
    wb = load_workbook(io.BytesIO(excel_bytes), rich_text=True)
    ws = buscar_hoja_por_examen(wb, nombre_examen)
    if ws is None:
        raise ValueError("No se encontró la hoja del examen seleccionado.")
    for fila in range(1, ws.max_row + 1):
        if norm(str(ws.cell(fila, 1).value or "")) == norm(campo):
            ws.cell(fila, 2).value = limpiar_valor_excel(valor, nombre_examen)
            ws.cell(fila, 2).data_type = "s"
            break
    else:
        # Las plantillas antiguas pueden no incluir todos los campos del área.
        # Solo se permite añadir un campo reconocido en la configuración actual.
        area_actual = st.session_state.get("area_resultado", st.session_state.get("area_activa"))
        campos_validos = get_field_aliases(area_actual)
        if not any(norm(c) == norm(campo) for c in campos_validos):
            raise ValueError(f"El campo '{campo}' no está configurado para el área seleccionada.")
        fila = ws.max_row + 1
        ws.cell(fila, 1).value = campo
        ws.cell(fila, 2).value = limpiar_valor_excel(valor, nombre_examen)
        ws.cell(fila, 2).data_type = "s"
        ws.cell(fila, 2).alignment = Alignment(vertical='top', wrap_text=True)
        ajustar_ancho_columnas(ws)
    ajustar_ancho_columnas(ws)
    for fila in range(1, ws.max_row + 1):
        ajustar_alto_fila_con_listas(ws, fila)
    salida = io.BytesIO()
    wb.save(salida)
    return salida.getvalue()


def texto_normas_con_negritas(campo, valor):
    """Devuelve texto plano de Normas, sin negritas, para máxima compatibilidad Excel."""
    texto = str(valor or "")
    return texto.replace("\r\n", "\n").replace("\r", "\n")


def aplicar_formato_normas_excel(excel_bytes, nombre_examen, origen_campos):
    """Aplica negritas y viñetas solo donde Normas fue la fuente usada.

    No toca otros exámenes ni sus formatos ni el contenido de SQLite.
    """
    # Algunas plantillas históricas (especialmente Hematología) escriben
    # «Rotulacion» sin tilde, aunque el campo institucional sea «Rotulación».
    # Comparar el nombre normalizado, pero usar el nombre canónico para
    # recuperar los subtítulos y aplicar las negritas del documento.
    campos = {norm(campo): campo for campo, fuente in (origen_campos or {}).items()
              if fuente == "Normas Generales de Toma de Muestras"}
    if not campos:
        return excel_bytes
    libro = load_workbook(io.BytesIO(excel_bytes), rich_text=True)
    hoja = buscar_hoja_por_examen(libro, nombre_examen)
    if hoja is None:
        return excel_bytes
    for fila in range(1, hoja.max_row + 1):
        nombre_plantilla = str(hoja.cell(fila, 1).value or "").strip()
        campo = campos.get(norm(nombre_plantilla))
        if campo and hoja.cell(fila, 2).value:
            limpio = fd.limpiar_formato_normas(campo, hoja.cell(fila, 2).value)
            hoja.cell(fila, 2).value = texto_normas_con_negritas(campo, limpio)
            hoja.cell(fila, 2).data_type = "s"
            # La sangría de cada viñeta forma parte del texto; no desplazar
            # toda la celda con indent=1 (que también desplaza los títulos).
            hoja.cell(fila, 2).alignment = Alignment(vertical="top", wrap_text=True)
            ajustar_alto_fila_con_listas(hoja, fila)
    salida = io.BytesIO()
    libro.save(salida)
    return salida.getvalue()


def aplicar_ediciones_a_excel(excel_bytes, edited_df, nombre_examen, campos_forzar=None):
    wb = load_workbook(io.BytesIO(excel_bytes), rich_text=True)

    ws = buscar_hoja_por_examen(wb, nombre_examen)
    # Una discrepancia validada es una decisión explícita aunque su estado
    # figure como «Encontrado», no «Modificado por usuario».
    forzados = {norm(campo) for campo in (campos_forzar or ())}
    modificado = False
    if ws is not None:
        filas = {}
        for r in range(1, ws.max_row + 1):
            val = str(ws.cell(r, 1).value or '').strip()
            if val:
                filas[norm(val)] = r

        for _, row_data in edited_df.iterrows():
            campo_norm = norm(row_data["Campo"])
            val_raw = row_data["Información extraída"]
            if pd.isna(val_raw) or val_raw is None:
                val_editado = ""
            else:
                val_editado = str(val_raw).strip()
                if val_editado.casefold() in ["nan", "none", "null"]:
                    val_editado = ""
            if campo_norm in filas:
                r_idx = filas[campo_norm]
                celda = ws.cell(r_idx, 2)
                estado_tabla = str(row_data.get("Estado", "") or "")
                # AG Grid puede devolver un texto visualmente aplanado sin que
                # el encargado haya editado la celda. Evitar convertir listas
                # enriquecidas a prosa durante un rerun de Streamlit.
                # El borrado deliberado sí se aplica al quedar el texto vacío.
                if (campo_norm not in forzados
                        and estado_tabla and estado_tabla != "Modificado por usuario"
                        and val_editado):
                    continue
                # Si el texto no cambió, conservar intactos los fragmentos
                # enriquecidos (negritas APA y títulos de Normas generales).
                if str(celda.value or "").strip() == val_editado:
                    continue
                if len(val_editado) > 32767:
                    raise ValueError("Excel admite como máximo 32.767 caracteres por celda.")
                fuente = st.session_state.get("origen_campos_documentales", {}).get(str(row_data["Campo"]).strip())
                if fuente == "Normas Generales de Toma de Muestras" and val_editado:
                    celda.value = texto_normas_con_negritas(str(row_data["Campo"]).strip(), val_editado)
                    celda.alignment = Alignment(vertical='top', wrap_text=True)
                else:
                    celda.value = val_editado
                    celda.alignment = Alignment(vertical='top', wrap_text=True)
                # Ajustar la altura también si el usuario modifica una lista.
                ajustar_alto_fila_con_listas(ws, r_idx)
                celda.data_type = "s"  # Protege texto manual que comience por '='.
                modificado = True

        # Referencias APA: se escriben como texto plano para evitar
        # reparaciones de Excel por formato enriquecido.
        fila_ref = filas.get(norm("Referencias"))
        if fila_ref:
            valor_ref = str(ws.cell(fila_ref, 2).value or "").strip()
            if (valor_ref and valor_ref == st.session_state.get("texto_referencias_auto")
                    and st.session_state.get("fuentes_referencias_aplicadas")):
                referencia_plana = apa.referencias_excel_con_cursiva(
                    st.session_state["fuentes_referencias_aplicadas"])
                if str(ws.cell(fila_ref, 2).value or "") != str(referencia_plana):
                    ws.cell(fila_ref, 2).value = referencia_plana
                    ws.cell(fila_ref, 2).data_type = "s"
                    modificado = True

    if not modificado:
        return excel_bytes  # Evita modificar el contenedor XLSX en cada rerun.
    # Ajuste autónomo en Python: no se requiere Excel de escritorio ni VBA.
    if ws is not None:
        ajustar_ancho_columnas(ws)
        for fila in range(1, ws.max_row + 1):
            ajustar_alto_fila_con_listas(ws, fila)
    out_buf = io.BytesIO()
    wb.save(out_buf)
    return out_buf.getvalue()


def campo_marcado_para_actualizar(valor):
    """Interpreta la casilla de AG Grid sin tratar la cadena 'false' como True."""
    if isinstance(valor, str):
        return valor.strip().casefold() not in ("false", "0", "no", "", "off")
    if valor is None:
        return True
    try:
        if pd.isna(valor):
            return True
    except (TypeError, ValueError):
        pass
    return bool(valor)


def respetar_campos_no_seleccionados(excel_nuevo, excel_anterior, nombre_examen, seleccion):
    """Restituye solo campos desmarcados desde la plantilla previa en Excel.

    Se hace sobre una copia al descargar o confirmar en SQLite: el resultado
    editable conserva todos los valores propuestos por si se vuelve a marcar.
    """
    excluidos = {norm(campo) for campo, marcar in (seleccion or {}).items()
                 if not campo_marcado_para_actualizar(marcar)}
    if not excluidos:
        return excel_nuevo
    libro_nuevo = load_workbook(io.BytesIO(excel_nuevo), rich_text=True)
    hoja_nueva = buscar_hoja_por_examen(libro_nuevo, nombre_examen)
    if hoja_nueva is None:
        return excel_nuevo
    hoja_anterior = None
    if excel_anterior:
        libro_anterior = load_workbook(io.BytesIO(excel_anterior), rich_text=True)
        hoja_anterior = buscar_hoja_por_examen(libro_anterior, nombre_examen)
    origen = {}
    if hoja_anterior is not None:
        origen = {norm(str(hoja_anterior.cell(r, 1).value or '')): hoja_anterior.cell(r, 2)
                  for r in range(1, hoja_anterior.max_row + 1)}
    cambios = False
    for r in range(1, hoja_nueva.max_row + 1):
        campo = norm(str(hoja_nueva.cell(r, 1).value or ''))
        if campo not in excluidos:
            continue
        celda = hoja_nueva.cell(r, 2)
        antigua = origen.get(campo)
        if antigua is not None:
            celda.value = copy.copy(antigua.value)
            # Nunca copiar _style entre libros diferentes: sus índices de
            # fuentes/rellenos pueden diferir e invalidar el XLSX resultante.
            celda.font = copy.copy(antigua.font)
            celda.fill = copy.copy(antigua.fill)
            celda.border = copy.copy(antigua.border)
            celda.alignment = copy.copy(antigua.alignment)
            celda.protection = copy.copy(antigua.protection)
            celda.number_format = antigua.number_format
            celda.hyperlink = copy.copy(antigua.hyperlink) if antigua.hyperlink else None
            celda.comment = copy.copy(antigua.comment) if antigua.comment else None
        else:
            celda.value = None  # No existía valor previo para este examen.
        ajustar_alto_fila_con_listas(hoja_nueva, r)
        cambios = True
    if not cambios:
        return excel_nuevo
    # Restaurar un párrafo largo puede volver a ampliar la columna; la altura
    # de las otras filas depende de ese ancho, por eso se ajustan juntas.
    ajustar_ancho_columnas(hoja_nueva)
    for fila in range(1, hoja_nueva.max_row + 1):
        ajustar_alto_fila_con_listas(hoja_nueva, fila)
    salida = io.BytesIO()
    libro_nuevo.save(salida)
    return salida.getvalue()


def validar_excel_para_descarga(excel_bytes):
    """Comprueba integridad ZIP y XML OOXML antes de entregar el .xlsx.

    Nunca intenta reparar silenciosamente un libro ni altera las fichas guardadas.
    No sustituye la prueba de apertura en Microsoft Excel.
    """
    if not isinstance(excel_bytes, bytes) or not excel_bytes:
        raise ValueError("El libro Excel generado está vacío.")
    try:
        # No convertir manualmente a sharedStrings: Excel rechazaba sus
        # propiedades cuando existían fragmentos enriquecidos vacíos.
        # Compatibilidad usa el serializador de openpyxl solo si detecta
        # cadenas compartidas antiguas o segmentos vacíos heredados.
        excel_bytes = ce.normalizar_texto_enriquecido_para_excel(excel_bytes)
        with zipfile.ZipFile(io.BytesIO(excel_bytes)) as paquete:
            miembros = set(paquete.namelist())
            requeridos = {"[Content_Types].xml", "_rels/.rels", "xl/workbook.xml"}
            if not requeridos.issubset(miembros):
                raise ValueError("Faltan componentes internos esenciales del archivo Excel.")
            danado = paquete.testzip()
            if danado:
                raise ValueError(f"El contenido del archivo Excel está dañado: {danado}.")
            # El límite de 32.767 se aplica a toda la celda enriquecida.
            # La validez XML sola no garantiza que Microsoft Excel la acepte.
            ce.validar_cadenas_ooxml(paquete)
            for nombre in miembros:
                if nombre.endswith((".xml", ".rels")):
                    try:
                        ET.fromstring(paquete.read(nombre))
                    except ET.ParseError as error:
                        raise ValueError(f"XML no válido en {nombre}: {error}.") from error
        libro = load_workbook(io.BytesIO(excel_bytes), read_only=True, data_only=False)
        try:
            if not libro.sheetnames:
                raise ValueError("El libro Excel no contiene hojas.")
        finally:
            libro.close()
    except (zipfile.BadZipFile, OSError, KeyError) as error:
        raise ValueError("El archivo Excel no es un .xlsx íntegro.") from error
    return excel_bytes


def comparar_plantillas_excel(bytes_antiguo, bytes_nuevo, nombre_examen):
    def extraer_datos_fichas(raw_bytes):
        if not raw_bytes:
            return {}, 0, []
        try:
            wb = load_workbook(io.BytesIO(raw_bytes), data_only=True, rich_text=True)
            hojas = wb.sheetnames
            datos_examen = {}
            ws = buscar_hoja_por_examen(wb, nombre_examen)
            if ws is not None:
                for r in range(1, ws.max_row + 1):
                    campo = str(ws.cell(r, 1).value or '').strip()
                    valor = str(ws.cell(r, 2).value or '').strip()
                    if campo:
                        datos_examen[norm(campo)] = (
                            nombre_campo_visible(campo),
                            "" if valor.casefold() in ["nan", "none", "null"] else valor)
            return datos_examen, len(hojas), hojas, ws is not None
        except Exception:
            return {}, 0, [], False

    datos_ant, total_hojas_ant, hojas_ant, existe_ant = extraer_datos_fichas(bytes_antiguo)
    datos_nue, total_hojas_nue, hojas_nue, existe_nue = extraer_datos_fichas(bytes_nuevo)

    campos_completados_ant = sum(1 for _, v in datos_ant.values() if v)
    campos_completados_nue = sum(1 for _, v in datos_nue.values() if v)

    diferencias = []
    todos_campos = list(dict.fromkeys(list(datos_ant.keys()) + list(datos_nue.keys())))

    for campo in todos_campos:
        etiqueta_ant, val_ant = datos_ant.get(campo, (campo, ""))
        etiqueta_nue, val_nue = datos_nue.get(campo, (campo, ""))

        if val_ant != val_nue:
            if not val_ant and val_nue:
                tipo = "Campo nuevo incorporado"
            elif val_ant and not val_nue:
                tipo = "Campo eliminado / borrado"
            else:
                tipo = "Información modificada"

            diferencias.append({
                "Campo": nombre_campo_visible(etiqueta_nue or etiqueta_ant),
                "Valor anterior": val_ant or "(Vacío)",
                "Valor nuevo": val_nue or "(Vacío)",
                "Tipo de cambio": tipo
            })

    es_nueva_hoja = (not existe_ant) and existe_nue

    return {
        "hojas_antes": total_hojas_ant,
        "hojas_despues": total_hojas_nue,
        "campos_completados_antes": campos_completados_ant,
        "campos_completados_despues": campos_completados_nue,
        "es_nueva_hoja": es_nueva_hoja,
        "diferencias": diferencias
    }


# --- ALTA DE USUARIOS DESDE EL LOGIN (SOLO CON CREDENCIALES DE ADMIN) ---
def _cerrar_alta_usuario_login():
    st.session_state["alta_usuario_login_abierta"] = False
    st.session_state.pop("alta_autorizacion_admin", None)
    st.session_state.pop("alta_revision_usuarios", None)


@st.dialog("Crear nuevo usuario", on_dismiss=_cerrar_alta_usuario_login)
def modal_crear_usuario_login():
    autorizacion = st.session_state.get("alta_autorizacion_admin")
    if autorizacion and not rul.autorizacion_vigente(DB_PATH, autorizacion):
        st.session_state.pop("alta_autorizacion_admin", None)
        autorizacion = None
        st.warning("La autorización ha caducado. Identifícate de nuevo como administrador.")

    if not autorizacion:
        st.caption("Para crear una cuenta se requiere autorización del administrador.")
        with st.form("alta_login_verificar_admin", clear_on_submit=True):
            admin_usuario = st.text_input("Usuario del administrador")
            admin_clave = st.text_input("Contraseña del administrador", type="password")
            autorizar = st.form_submit_button("Verificar administrador", type="primary")
        if autorizar:
            ahora = time.monotonic()
            intentos = [v for v in st.session_state.get("alta_intentos_admin", [])
                        if ahora - v < 300]
            if len(intentos) >= 5:
                st.error("Demasiados intentos. Inténtalo más tarde o contacta al administrador.")
            else:
                comprobante = rul.verificar_administrador(DB_PATH, admin_usuario, admin_clave)
                if comprobante:
                    st.session_state["alta_autorizacion_admin"] = comprobante
                    st.session_state["alta_revision_usuarios"] = gp.estado_recurso(
                        DB_PATH, "usuarios:administracion")["revision"]
                    st.session_state["alta_intentos_admin"] = []
                    st.rerun()
                else:
                    st.session_state["alta_intentos_admin"] = intentos + [ahora]
                    st.error("Credenciales de administrador incorrectas o sin autorización.")
    else:
        st.caption("Autorización verificada. La nueva cuenta tendrá el rol **Encargado**.")
        with st.form("alta_login_datos_usuario", clear_on_submit=False):
            usuario_nuevo = st.text_input("Usuario de acceso")
            nombre_nuevo = st.text_input("Nombre visible")
            clave_nueva = st.text_input("Contraseña nueva (mínimo 12 caracteres)", type="password")
            clave_repetida = st.text_input("Repetir contraseña nueva", type="password")
            confirmar = st.checkbox("Confirmo la creación de esta cuenta.")
            registrar = st.form_submit_button("Crear usuario", type="primary")
        if registrar:
            if not confirmar:
                st.error("Confirma la creación de la cuenta.")
            else:
                try:
                    # Validar primero; no consumir respaldos ante claves incorrectas.
                    rul.validar_datos_nueva_cuenta(
                        usuario_nuevo, nombre_nuevo, clave_nueva, clave_repetida)
                    # Respaldos consistentes. Nunca sustituir la base activa.
                    gp.respaldar(DB_PATH)
                    rul.crear_encargado_autorizado(
                        DB_PATH, autorizacion,
                        st.session_state.get("alta_revision_usuarios", -1),
                        usuario=usuario_nuevo, nombre_visible=nombre_nuevo,
                        contrasena=clave_nueva, repetir_contrasena=clave_repetida,
                    )
                    st.session_state["alta_usuario_login_exito"] = (
                        "Usuario creado correctamente. Ya puedes iniciar sesión con tu nueva cuenta."
                    )
                    _cerrar_alta_usuario_login()
                    try:
                        gp.respaldar(DB_PATH)
                    except OSError:
                        st.session_state["alta_usuario_login_exito"] += (
                            " Atención: no se pudo generar el respaldo posterior; revisa los respaldos."
                        )
                    st.rerun()
                except (ValueError, PermissionError, gp.ConflictoConcurrente,
                        sqlite3.Error, OSError) as error:
                    st.error(f"No se creó la cuenta: {error}")
    if st.button("Cancelar", key="cancelar_alta_usuario_login"):
        _cerrar_alta_usuario_login()
        st.rerun()


# --- ACCESO CON CUENTAS INDIVIDUALES EN SQLITE ---
def check_login():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False

    if st.session_state["authenticated"] and st.session_state.get("usuario_acceso"):
        if gp.usuario_habilitado(DB_PATH, st.session_state["usuario_acceso"],
                                 st.session_state.get("revision_sesion", -1)):
            return True
        st.session_state["authenticated"] = False
        st.session_state.pop("usuario_acceso", None)
        st.session_state.pop("nombre_usuario", None)
        st.session_state.pop("revision_sesion", None)

    st.session_state["authenticated"] = False
    st.session_state.pop("modal_area_activa", None)
    st.session_state["mostrar_modal_examenes"] = False

    _, col_login, _ = st.columns([1, 1.2, 1])
    with col_login:
        st.markdown("<h1 style='text-align: center;'>Acceso a herramienta</h1>", unsafe_allow_html=True)
        st.caption("Identifícate con la cuenta asignada por el laboratorio.")
        if st.session_state.pop("alta_usuario_login_exito", None):
            st.success("Usuario creado correctamente. Ya puedes iniciar sesión con tu nueva cuenta.")
        if not gp.cantidad_usuarios(DB_PATH):
            st.error("Aún no existen cuentas. Ejecuta `python crear_usuarios.py` en el servidor para registrar a los tres encargados.")
            return False
        with st.form("login_form", clear_on_submit=False):
            username = st.text_input("Usuario")
            password = st.text_input("Contraseña", type="password")
            submitted = st.form_submit_button("Iniciar sesión", type="primary", use_container_width=True)
        if submitted:
            now = datetime.now(ZoneInfo("America/Santiago"))
            intentos = st.session_state.get("login_intentos", [])
            intentos = [t for t in intentos if (now-t).total_seconds() < 300]
            if len(intentos) >= 5:
                st.error("Demasiados intentos. Cierra esta sesión y contacta al administrador.")
            else:
                credencial = gp.autenticar(DB_PATH, username, password)
                if credencial:
                    st.session_state["authenticated"] = True
                    st.session_state["usuario_acceso"] = credencial["usuario"]
                    st.session_state["nombre_usuario"] = credencial["nombre"]
                    st.session_state["revision_sesion"] = credencial["revision_sesion"]
                    st.session_state["login_intentos"] = []
                    st.session_state["area_activa"] = None
                    gp.registrar_acceso(DB_PATH, credencial["usuario"], "Inicio de sesión")
                    st.rerun()
                else:
                    intentos.append(now)
                    st.session_state["login_intentos"] = intentos
                    st.error("Usuario o contraseña incorrectos.")
        if st.button("Crear nuevo usuario", use_container_width=True,
                     key="abrir_alta_usuario_login"):
            st.session_state["alta_usuario_login_abierta"] = True
            st.rerun()
        if st.session_state.get("alta_usuario_login_abierta"):
            modal_crear_usuario_login()
    return False


def export_imported_workbook(template_bytes, fmt, selected_sheet=None):
    wb = load_workbook(io.BytesIO(template_bytes), data_only=False, rich_text=True)
    sheet_names = wb.sheetnames

    def sheet_to_rows(ws):
        return [[cell.value for cell in row] for row in ws.iter_rows()]

    if fmt == "XLSX":
        return template_bytes, "excel_importado.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    if fmt == "CSV":
        if selected_sheet not in sheet_names:
            raise ValueError("Selecciona una hoja válida para exportar a CSV.")
        ws = wb[selected_sheet]
        output = io.StringIO(newline="")
        writer = csv.writer(output)
        for row in sheet_to_rows(ws):
            writer.writerow(["" if value is None else value for value in row])
        return output.getvalue().encode("utf-8-sig"), "excel_importado.csv", "text/csv; charset=utf-8"

    if fmt == "XLS":
        try:
            import xlwt
        except ImportError as exc:
            raise RuntimeError("Falta la dependencia xlwt. Instálala con: pip install xlwt") from exc
        out_book = xlwt.Workbook(encoding="utf-8")
        used_names = set()
        for ws in wb.worksheets:
            base_name = re.sub(r'[\\/*?:\[\]]', '_', ws.title)[:31] or "Hoja"
            name = base_name
            suffix = 1
            while name.lower() in used_names:
                ending = f"_{suffix}"
                name = base_name[:31 - len(ending)] + ending
                suffix += 1
            used_names.add(name.lower())
            out_ws = out_book.add_sheet(name)
            for row_idx, row in enumerate(ws.iter_rows(), start=0):
                for col_idx, cell in enumerate(row):
                    value = cell.value
                    if value is None:
                        continue
                    if isinstance(value, (str, int, float, bool)):
                        out_ws.write(row_idx, col_idx, str(value)[:32000] if isinstance(value, str) else value)
                    else:
                        out_ws.write(row_idx, col_idx, str(value)[:32000])
        buffer = io.BytesIO()
        out_book.save(buffer)
        return buffer.getvalue(), "resultado_extraccion.xls", "application/vnd.ms-excel"

    if fmt == "JSON":
        payload = {"archivo": "excel_importado.xlsx", "hojas": []}
        for ws in wb.worksheets:
            payload["hojas"].append({
                "nombre": ws.title,
                "filas": sheet_to_rows(ws),
            })
        return json.dumps(payload, ensure_ascii=False, indent=2, default=str).encode(
            "utf-8"), "excel_importado.json", "application/json"

    if fmt == "XML":
        root = ET.Element("libro_excel")
        for ws in wb.worksheets:
            sheet_el = ET.SubElement(root, "hoja", nombre=ws.title)
            for row_idx, row in enumerate(ws.iter_rows(), start=1):
                row_el = ET.SubElement(sheet_el, "fila", numero=str(row_idx))
                for cell in row:
                    if cell.value is None:
                        continue
                    cell_el = ET.SubElement(row_el, "celda", coordenada=cell.coordinate)
                    cell_el.text = str(cell.value)
        xml_bytes = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        return xml_bytes, "excel_importado.xml", "application/xml"

    raise ValueError("Formato de exportación no compatible.")


def guide_tab():
    """Orientaciones para usuarios, no instrucciones para programadores."""
    st.subheader("Guía básica para usuario", anchor=False)
    st.write("Recomendaciones para utilizar la herramienta y revisar los datos antes de guardarlos.")
    with st.expander("1. Seleccionar el examen y sus fuentes", expanded=False):
        st.markdown(
            "- Selecciona el área y el examen correctos antes de empezar.\n"
            "- Verifica el laboratorio de procesamiento y la Cartera asociada.\n"
            "- Las fuentes documentales tienen prioridades; revisa su procedencia cuando haya dudas.\n"
            "- Si solo necesitas fuentes institucionales, utiliza **Omitir análisis** y genera la ficha sin PDF.")
    with st.expander("2. Analizar documentos y revisar resultados"):
        st.markdown(
            "- Utiliza **Fichas técnicas** para PDF con texto y **Documentos escaneados** "
            "cuando corresponda realizar lectura óptica.\n"
            "- Comprueba especialmente las unidades, los códigos, los valores de referencia "
            "y la información clínica.\n"
            "- Un campo no encontrado o pendiente debe verificarse antes de guardarse.")
    with st.expander("3. Resolver discrepancias"):
        st.markdown(
            "- Cuando el documento y una fuente institucional ofrecen valores distintos, "
            "abre **Comparar y resolver discrepancias**.\n"
            "- Lee ambas opciones y selecciona la información que corresponda al examen.\n"
            "- No aceptes un valor solo por el color o por el orden en el que aparece.")
    with st.expander("4. Editar campos y conservar información"):
        st.markdown(
            "- Revisa las propuestas antes de incorporarlas a la ficha.\n"
            "- Edita las celdas que necesiten corrección y utiliza **Deshacer** "
            "cuando corresponda.\n"
            "- Desmarca **Aplicar** si necesitas conservar el valor anterior de un campo.")
    with st.expander("5. Descargar y actualizar fichas"):
        st.markdown(
            "- Comprueba el nombre del examen, los campos y los datos operacionales.\n"
            "- Descarga la ficha para una última revisión antes de usarla.\n"
            "- Actualiza la base de datos solo si la información está confirmada y autorizada.")
    with st.expander("6. Proteger documentos y cuentas"):
        st.markdown(
            "- Utiliza tu cuenta personal y nunca compartas tu contraseña.\n"
            "- Evita cargar documentos con datos personales innecesarios.\n"
            "- Comparte las fichas únicamente con personas autorizadas.\n"
            "- Cierra tu sesión al terminar, especialmente en computadores compartidos.")
    with st.expander("7. Si encuentras un problema"):
        st.markdown(
            "- Anota el nombre del examen, el campo afectado y qué esperabas encontrar.\n"
            "- No guardes una ficha si observas información incorrecta o incompleta.\n"
            "- Informa al encargado o al soporte institucional, sin compartir contraseñas "
            "ni documentación clínica sensible.")


# --- CONFIGURACIÓN DE PÁGINA Y ESTILOS CSS ---
icon_path = os.path.join(os.path.dirname(__file__), "Pestaña.png")

st.set_page_config(
    page_title=APP_TITLE,
    page_icon=icon_path if os.path.exists(icon_path) else "🧪",
    layout="wide"
)

st.markdown("""
<style>
    .ag-theme-alpine ::-webkit-scrollbar {
        width: 8px;
        height: 8px;
    }
    .ag-theme-alpine ::-webkit-scrollbar-track {
        background: #f1f5f9;
        border-radius: 4px;
    }
    .ag-theme-alpine ::-webkit-scrollbar-thumb {
        background-color: #1E293B;
        border-radius: 4px;
    }
    .ag-theme-alpine ::-webkit-scrollbar-thumb:hover {
        background-color: #0f172a;
    }
    .banner-alerta-expiracion {
        background-color: #fee2e2;
        border: 1px solid #fca5a5;
        border-left: 6px solid #dc2626;
        color: #991b1b;
        padding: 12px 16px;
        border-radius: 6px;
        font-size: 0.92rem;
        margin-bottom: 20px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .banner-cola-trabajo {
        background-color: #f0f9ff;
        border: 1px solid #bae6fd;
        border-left: 6px solid #0284c7;
        color: #0369a1;
        padding: 12px 16px;
        border-radius: 6px;
        font-size: 0.95rem;
        margin-bottom: 20px;
    }

    /* Solo el valor del examen tiene color; la etiqueta conserva el color del tema. */
    .sidebar-examen-activo {font-weight:600; line-height:1.55; overflow-wrap:anywhere;}
    .sidebar-examen-activo span {color:#50e89a;}
    .sidebar-examen-sin-seleccionar {font-weight:600; line-height:1.55;}
    .sidebar-examen-sin-seleccionar span {color:#ffb45b;}

    /* Solo en el aviso del manual: nombres de campos blancos sobre azul oscuro.
       El resto de avisos y los colores del modo claro no se modifican. */
    .st-key-aviso_campos_manual [data-testid="stAlert"] strong {
        color: #FFFFFF !important;
        background-color: #245B8A;
        padding: 1px 5px;
        border-radius: 4px;
        font-weight: 700;
        -webkit-box-decoration-break: clone;
        box-decoration-break: clone;
    }


    /* Las tres pestañas principales se presentan como botones independientes.
       No aplica a subpestañas de administración ni altera la selección. */
    .st-key-pestanas_principales [data-testid="stTabs"]:not(
        [data-testid="stTabs"] [data-testid="stTabs"]
    ) [role="tablist"] {
        gap: 9px;
        border-bottom: 0;
        flex-wrap: wrap;
    }
    .st-key-pestanas_principales [data-testid="stTabs"]:not(
        [data-testid="stTabs"] [data-testid="stTabs"]
    ) [role="tab"] {
        border: 1px solid color-mix(in srgb, var(--text-color) 28%, transparent);
        border-radius: 10px;
        padding: 8px 15px;
        background: var(--secondary-background-color);
        color: var(--text-color);
        transition: background-color 180ms ease, color 180ms ease, border-color 180ms ease;
    }
    .st-key-pestanas_principales [data-testid="stTabs"]:not(
        [data-testid="stTabs"] [data-testid="stTabs"]
    ) [role="tab"]:hover {
        background: #ffffff;
        color: #245b8a;
        border-color: #9db8d3;
    }
    .st-key-pestanas_principales [data-testid="stTabs"]:not(
        [data-testid="stTabs"] [data-testid="stTabs"]
    ) [role="tab"][aria-selected="true"] {
        border-color: var(--primary-color);
        font-weight: 700;
    }
    .st-key-pestanas_principales [data-testid="stTabs"]:not(
        [data-testid="stTabs"] [data-testid="stTabs"]
    ) [role="tab"]:focus-visible {
        outline: 2px solid var(--primary-color);
        outline-offset: 3px;
    }
    @media (prefers-reduced-motion: reduce) {
        .st-key-pestanas_principales [role="tab"] { transition: none; }
    }

    /* Botón de avance: azul, independiente del tema del resto de Streamlit. */
    .st-key-contenedor_continuar_examen button {
        background-color: #2563eb !important;
        border: 1px solid #1d4ed8 !important;
        color: #ffffff !important;
    }
    .st-key-contenedor_continuar_examen button:hover {
        background-color: #1d4ed8 !important;
        border-color: #1e40af !important;
        color: #ffffff !important;
    }
    .st-key-contenedor_continuar_examen button:focus-visible {
        box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.28) !important;
    }

    .num-slot {
        height: 37px;
        display: flex;
        align-items: center;
        justify-content: center;
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
        outline: none !important;
        font-weight: bold;
        font-size: 1rem;
        color: inherit !important;
        margin-bottom: 5px;
    }

    iframe[title*="streamlit_sortables"] {
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
        width: 100% !important;
        max-width: 800px;
    }
</style>
""", unsafe_allow_html=True)

if not check_login():
    st.stop()

# El puente preserva el desplazamiento en las acciones de la ficha.
# No usa datos de pacientes ni altera las operaciones del servidor.
iflu.activar()

# Botón flotante: flecha Material (no emoji), sin etiquetas visibles.
# Avisos informativos compactos: se adaptan al texto sin estirar a toda la fila.
st.markdown("""
<style>
  div[data-testid="stAlert"] {
      width: fit-content !important;
      max-width: 100% !important;
      margin-right: auto !important;
      box-sizing: border-box !important;
      overflow-wrap: anywhere;
  }
  /* Hoja de cuaderno: trazar la línea DESPUÉS del contenido de cada campo.
     El patrón de fondo con espacios fijos cortaba los renglones al envolver texto.
     Cada borde se adapta automáticamente a la altura real de su fila. */
  .cuaderno-procedencia {
      padding: 10px 16px 10px 29px;
      border: 1px solid #dce5f0;
      border-radius: 10px;
      border-left: 3px solid #edbbbb;
      background-color: #fcfdff;
  }
  .cuaderno-fila {
      min-height: 53px;
      box-sizing: border-box;
      padding: 7px 0 4px 4px;
      border-bottom: 1px solid #d6e5f7;
  }
  .cuaderno-campo { color: #17365D; font-size: 0.96rem; font-weight: 700; }
  .cuaderno-fuente { color: #556479; font-size: 0.84rem; }
  .cuaderno-fuente-seleccionada { color: #17365D; font-weight: 700; }
  .cuaderno-sin-fuente { color: #b76b71; font-weight: 600; }
  .cuaderno-separador { color: #9ba7b7; margin: 0 0.35rem; }
  .st-key-reinicio_flotante {
      position: fixed !important;
      right: 22px;
      /* Por encima del panel «Manage app» de Streamlit Community Cloud. */
      bottom: 78px;
      z-index: 1001;
      width: 56px !important;
      min-width: 56px !important;
      pointer-events: auto;
  }
  .st-key-reinicio_flotante button {
      width: 56px !important;
      height: 56px !important;
      min-width: 56px !important;
      min-height: 56px !important;
      padding: 0 !important;
      border-radius: 50% !important;
      display: flex !important;
      align-items: center !important;
      justify-content: center !important;
      gap: 0 !important;
      color: #111827 !important;
      background: #f8fafc !important;
      border: 1px solid #e2e8f0 !important;
      box-shadow: 0 3px 12px rgba(20, 40, 60, 0.10) !important;
      transition: background-color 180ms ease, border-color 180ms ease,
                  box-shadow 180ms ease !important;
  }
  .st-key-reinicio_flotante button:hover {
      background: #f1f5f9 !important;
      border-color: #cbd5e1 !important;
  }
  .st-key-reinicio_flotante button:focus-visible {
      outline: 2px solid #64748b !important;
      outline-offset: 3px;
  }
  /* El texto del botón queda accesible, pero fuera del área visible.
     Así la flecha se centra sin texto a un lado. */
  .st-key-reinicio_flotante button [data-testid="stMarkdownContainer"] {
      position: absolute !important;
      width: 1px !important;
      height: 1px !important;
      overflow: hidden !important;
      clip-path: inset(50%) !important;
      white-space: nowrap !important;
      padding: 0 !important;
      margin: 0 !important;
  }
  .st-key-reinicio_flotante button > div {
      gap: 0 !important;
      justify-content: center !important;
      align-items: center !important;
  }
  .st-key-reinicio_flotante button [data-testid="stIconMaterial"] {
      font-size: 30px !important;
      line-height: 1 !important;
      display: inline-flex !important;
      justify-content: center !important;
      align-items: center !important;
      flex: 0 0 30px !important;
      width: 30px !important;
      height: 30px !important;
      margin: 0 !important;
      transform-origin: center center;
      /* La animación existe también en reposo: al pausar conserva el ángulo
         alcanzado, sin volver bruscamente a la posición inicial. */
      animation: reinicio_flecha_giratoria 0.9s linear infinite both !important;
      animation-play-state: paused !important;
  }
  /* Giro limitado a tres segundos en CSS, sin ejecutar Streamlit cada segundo.
     Dos nombres alternos permiten iniciar otra espera tras la expiración. */
  @keyframes reinicio_flecha_giratoria {
      to { transform: rotate(360deg); }
  }
  @keyframes reinicio_flecha_giratoria_alternativa {
      to { transform: rotate(360deg); }
  }
  .st-key-reinicio_flotante:has(.estado-reinicio-confirmar) button [data-testid="stIconMaterial"] {
      animation-duration: 1s !important;
      animation-iteration-count: 3 !important;
      animation-fill-mode: both !important;
      animation-play-state: running !important;
  }
  .st-key-reinicio_flotante:has(.estado-reinicio-confirmar.estado-reinicio-intento-0) button [data-testid="stIconMaterial"] {
      animation-name: reinicio_flecha_giratoria !important;
  }
  .st-key-reinicio_flotante:has(.estado-reinicio-confirmar.estado-reinicio-intento-1) button [data-testid="stIconMaterial"] {
      animation-name: reinicio_flecha_giratoria_alternativa !important;
  }
  /* Confirmación: degradado verde 1,5 s sobre el botón gris.
     El pseudo-elemento permite que la animación sea visible aunque el botón
     nativo de Streamlit defina su fondo con !important. */
  @keyframes reinicio_confirmacion_suave {
      0% { opacity: 0; }
      22% { opacity: 1; }
      70% { opacity: 1; }
      100% { opacity: 0; }
  }
  .st-key-reinicio_flotante button { position: relative !important; isolation: isolate; }
  .st-key-reinicio_flotante:has(.estado-reinicio-completado) button::before {
      content: "";
      position: absolute;
      inset: 0;
      border-radius: inherit;
      background: radial-gradient(circle at 30% 25%, #bbf7d0 0%, #dcfce7 55%, #f0fdf4 100%);
      opacity: 0;
      pointer-events: none;
      z-index: 0;
      animation: reinicio_confirmacion_suave 1.5s ease-in-out both;
  }
  .st-key-reinicio_flotante:has(.estado-reinicio-completado) button [data-testid="stIconMaterial"] {
      position: relative;
      z-index: 1;
  }
  .st-key-reinicio_flotante:has(.estado-reinicio-completado) button,
  .st-key-reinicio_flotante:has(.estado-reinicio-completado) button:hover {
      box-shadow: 0 3px 12px rgba(20, 40, 60, 0.10) !important;
  }
  @media (prefers-reduced-motion: reduce) {
      .st-key-reinicio_flotante:has(.estado-reinicio-confirmar) button [data-testid="stIconMaterial"] {
          animation-duration: 3s !important;
          animation-iteration-count: 1 !important;
      }
  }
  @media (max-width: 640px) {
      .st-key-reinicio_flotante {
          right: 12px; bottom: 96px;
          width: 48px !important; min-width: 48px !important;
      }
      .st-key-reinicio_flotante button {
          width: 48px !important; height: 48px !important;
          min-width: 48px !important; min-height: 48px !important;
      }
  }
  /* Referencias conserva su estado al confirmar (sin animación forzada). */
  [class*="st-key-pdf_ir_izquierda_"] button,
  [class*="st-key-pdf_ir_derecha_"] button {
      border-radius: 8px !important;
      min-width: 2.35rem; min-height: 2.9rem;
      font-size: 1.15rem;
      padding: 0.2rem !important;
  }
  /* Sugerencias: sin espacios excesivos entre botones, sólo dentro de la lista. */
  [class*="st-key-lista_coincidencias_"] [data-testid="stVerticalBlock"],
  [class*="st-key-lista_coincidencias_"][data-testid="stVerticalBlock"] {
      gap: 0.28rem !important;
  }
  /* Tooltip sin duración artificial: Streamlit controla su apertura al pasar
     el mouse. No se permite prolongar la vista moviendo el mouse sobre el
     propio mensaje; al abandonar el activador desaparece sin animación. */
  [data-testid="stTooltipContent"],
  [role="tooltip"] {
      animation: none !important;
      transition: none !important;
      pointer-events: none !important;
  }
</style>
""", unsafe_allow_html=True)

# --- VERIFICACIÓN DE 6 MESES PARA LA CARTERA DE EXÁMENES ---
fecha_cartera, cartera_expirada, msg_cartera = fecha_actualizacion_cartera()

if "mostrar_dialogo_expiracion" not in st.session_state and cartera_expirada:
    st.session_state["mostrar_dialogo_expiracion"] = True

if cartera_expirada:
    st.markdown(f"""
    <div class="banner-alerta-expiracion">
        ATENCIÓN: LA CARTERA DE EXÁMENES SUPERA LOS 6 MESES DE ANTIGÜEDAD<br>
        <span style="font-size: 0.85rem;">
        {msg_cartera if not fecha_cartera else f"La última versión de la Cartera de Exámenes fue registrada el <b>{msg_cartera}</b>. Se recomienda actualizar el Excel en la barra lateral."}
        </span>
    </div>
    """, unsafe_allow_html=True)

st.markdown('<style>\n[data-testid="stHeaderActionElements"],\n[data-testid="stMarkdownContainer"] :is(h1,h2,h3,h4,h5,h6) a[href^="#"] { display:none !important; }\n[data-testid="stElementContainer"]:has([data-testid="stImage"]) [data-testid="stElementToolbar"],\n[data-testid="stImage"] button,\n[data-testid="stImage"] .overlayBtn,\n.st-key-traza_previa_modificaciones [data-testid="stElementToolbar"],\n.st-key-traza_previa_actividad [data-testid="stElementToolbar"] { display:none !important; }\n</style>', unsafe_allow_html=True)
st.title(APP_TITLE, anchor=False)
if "mensaje_cambio_confirmado" in st.session_state:
    st.success(st.session_state.pop("mensaje_cambio_confirmado"))
if "mensaje_respaldo" in st.session_state:
    aviso_respaldo = st.session_state.pop("mensaje_respaldo")
    if aviso_respaldo.startswith("Atención:"):
        st.warning(aviso_respaldo)
    else:
        st.caption(aviso_respaldo)

if "examenes_por_area" not in st.session_state:
    st.session_state["examenes_por_area"] = cargar_examenes_por_area()
if "area_activa" not in st.session_state:
    st.session_state["area_activa"] = None
if "examen_activo" not in st.session_state:
    st.session_state["examen_activo"] = None
if "aggrid_key_counter" not in st.session_state:
    st.session_state["aggrid_key_counter"] = 0
if "cola_pdf_indice" not in st.session_state:
    st.session_state["cola_pdf_indice"] = 0
if "mostrar_modal_examenes" not in st.session_state:
    st.session_state["mostrar_modal_examenes"] = False
# Consultar revisiones compartidas en cada interacción: otras sesiones pueden haber
# cambiado listas en SQLite desde la última carga en este navegador.
with gp.conexion_cerrada(DB_PATH) as con_revisiones:
    id_ultimo_cambio = con_revisiones.execute("SELECT COALESCE(MAX(id),0) FROM auditoria_cambios").fetchone()[0]
if st.session_state.get("id_ultimo_cambio_compartido") != id_ultimo_cambio:
    st.session_state["examenes_por_area"] = cargar_examenes_por_area()
    st.session_state["id_ultimo_cambio_compartido"] = id_ultimo_cambio


# --- TRANSACCIONES COMPARTIDAS Y CONFIRMACIÓN OBLIGATORIA ---
ETIQUETAS_RECURSO = {
    "fuente:cartera_examenes": "Cartera de exámenes",
    "fuente:cartera_acreditacion": "Cartera de acreditación",
    "fuente:homologacion_tests": "Homologació de test",
    "fuente:manual_toma_muestras": "Manual de toma de muestras",  # Trazabilidad histórica
    "fuente:normas_generales": "Normas generales de toma de muestras",
    "catalogo:laboratorios": "Lista de laboratorios",
    "catalogo:encargados": "Lista de encargados",
}


def etiqueta_recurso(recurso):
    for area_ht in htest.AREAS:
        if recurso == htest.recurso_area(area_ht):
            return "Homologación de test — " + area_ht
    if recurso.startswith("area:"):
        return "Área: " + recurso.split(":", 1)[1]
    if recurso.startswith("examenes:"):
        return "Exámenes — " + recurso.split(":", 1)[1]
    return ETIQUETAS_RECURSO.get(recurso, recurso)


def solicitar_cambio(tipo, recurso, archivo=None, nombre_archivo=None, descripcion=None, **datos):
    """La escritura queda pendiente; se realiza SOLO dentro del modal confirmado."""
    if not st.session_state.get("usuario_acceso"):
        raise PermissionError("Debes iniciar sesión para modificar la base de datos.")
    previo = gp.estado_recurso(DB_PATH, recurso)
    st.session_state["cambio_pendiente"] = {
        "tipo": tipo, "recurso": recurso, "revision": previo["revision"],
        "previo": previo, "archivo": archivo, "nombre_archivo": nombre_archivo,
        "descripcion": descripcion or "Actualizar información compartida", "datos": datos,
        "uid": uuid4().hex[:12]
    }
    st.rerun()


def _escribir_catalogo_conexion(con, nombre_tabla, nombres):
    if nombre_tabla not in ("laboratorios", "encargados"):
        raise ValueError("Catálogo desconocido")
    normalizados = _normalizar_nombres(nombres)
    existentes = {x[0].casefold(): x[0] for x in con.execute(f"SELECT nombre FROM {nombre_tabla}")}
    for k, nombre in normalizados.items():
        if k in existentes:
            con.execute(f"UPDATE {nombre_tabla} SET activo=1 WHERE nombre=?", (existentes[k],))
        else:
            con.execute(f"INSERT INTO {nombre_tabla}(nombre,activo) VALUES(?,1)", (nombre,))
    for k, nombre in existentes.items():
        if k not in normalizados:
            con.execute(f"UPDATE {nombre_tabla} SET activo=0 WHERE nombre=?", (nombre,))
    return len(normalizados)


def _escribir_cambio(con, pendiente):
    tipo, datos = pendiente["tipo"], pendiente["datos"]
    archivo, nombre = pendiente.get("archivo"), pendiente.get("nombre_archivo")
    if tipo == "area":
        ok, mensaje = guardar_plantilla_area_db(
            datos["area"], archivo, nombre, conexion=con,
            preservar_fichas_importadas=True)
        if not ok:
            raise ValueError(mensaje)
        return mensaje
    if tipo == "cartera":
        df = pd.read_excel(io.BytesIO(archivo))
        ok, mensaje = guardar_cartera_db(df, nombre, archivo, conexion=con)
        if not ok:
            raise ValueError(mensaje)
        return mensaje
    if tipo == "acreditacion":
        n = fd.guardar_cartera_acreditacion(DB_PATH, archivo, nombre, conexion=con)
        return f"Cartera actualizada: {n} prestaciones."
    if tipo == "homologacion_tests":
        area_ht = datos["area_fuente"]
        meta = htest.guardar(DB_PATH, archivo, nombre, conexion=con,
                            hoja=datos.get("hoja"), columna_examen=datos.get("columna_examen"),
                            columna_codigo=datos.get("columna_codigo"), area_fuente=area_ht)
        return f"Homologación de test de {area_ht} actualizada: {meta['registros']} filas válidas."
    if tipo == "normas":
        n = fd.guardar_normas(DB_PATH, archivo, nombre, conexion=con)
        return f"Normas actualizadas: {n} apartados."
    if tipo == "examenes":
        actualizar_lista_area_conexion(con, datos["area"], datos["nombres"])
        return "Lista de exámenes actualizada."
    if tipo == "catalogo":
        n = _escribir_catalogo_conexion(con, datos["tabla"], datos["nombres"])
        return f"Lista de opciones actualizada: {n} registros activos."
    raise ValueError("Tipo de actualización no reconocido.")


def confirmar_cambio_pendiente(pendiente):
    usuario = st.session_state["usuario_acceso"]
    # Copia consistente ANTES de reemplazar la versión compartida.
    gp.respaldar(DB_PATH)
    resultado = gp.ejecutar_cambio(
        DB_PATH, pendiente["recurso"], usuario, pendiente["revision"],
        pendiente["descripcion"], lambda con: _escribir_cambio(con, pendiente),
        archivo=pendiente.get("nombre_archivo"),
        detalle="Actualización confirmada desde la interfaz"
    )
    if pendiente["tipo"] in ("area", "examenes"):
        st.session_state["examenes_por_area"] = cargar_examenes_por_area()
    if pendiente["tipo"] == "examenes":
        datos = pendiente["datos"]
        st.session_state["mostrar_modal_examenes"] = False
        st.session_state.pop("modal_area_activa", None)
        if datos.get("seleccionar_examen"):
            st.session_state["area_activa"] = datos["area"]
            st.session_state["examen_activo"] = datos["seleccionar_examen"]
            st.session_state.pop("contexto_emparejamiento_cartera", None)
            st.session_state.pop("cartera_seleccion_confirmada_manualmente", None)
        elif datos.get("quitar_examen") == st.session_state.get("examen_activo") and datos["area"] == st.session_state.get("area_activa"):
            st.session_state["examen_activo"] = None
        st.session_state.pop(f"selector_examen_{datos['area']}", None)
    if pendiente["tipo"] == "catalogo":
        for key in ("sel_lab_main", "sel_act_main", "sel_aut_main", "del_lab_select", "del_act_select", "del_aut_select"):
            st.session_state.pop(key, None)
    if pendiente["tipo"] == "acreditacion":
        st.session_state["acreditacion_manual_id"] = None
        st.session_state["_acreditacion_reset_pendiente"] = True
    if pendiente["tipo"] == "cartera":
        st.session_state["mostrar_dialogo_expiracion"] = False
    try:
        destino_respaldo = gp.respaldar(DB_PATH)
        st.session_state["mensaje_respaldo"] = f"Respaldo generado: {os.path.basename(destino_respaldo)}"
    except Exception as error_respaldo:
        st.session_state["mensaje_respaldo"] = f"Atención: cambio guardado, pero falló su respaldo: {error_respaldo}"
    return resultado


# Streamlit normalmente solo oculta el diálogo al pulsar X/Esc; el estado que lo
# abrió debe limpiarse también. Los callbacks se ejecutan ANTES del nuevo rerun.
def _cerrar_confirmacion_por_x():
    st.session_state.pop("cambio_pendiente", None)


def _cerrar_alerta_por_x():
    st.session_state["mostrar_dialogo_expiracion"] = False


def _cerrar_candidatos_por_x():
    st.session_state.pop("candidatos_cartera_pendientes", None)


def _cerrar_discrepancias_por_x():
    st.session_state["modal_discrepancias_unificado"] = False
    # Se conservan las diferencias sin resolver para poder abrirlas manualmente.


def _cerrar_planilla_por_x():
    st.session_state["modal_confirmar_planilla_abierto"] = False
    st.session_state.pop("excel_para_actualizar", None)
    st.session_state.pop("revision_plantilla_pendiente", None)
    st.session_state.pop("metadata_plantilla_pendiente", None)


def _cerrar_examenes_por_x():
    st.session_state["mostrar_modal_examenes"] = False
    st.session_state.pop("modal_area_activa", None)


def _cerrar_administracion_por_x():
    st.session_state["modal_admin_usuarios_abierto"] = False
    st.session_state.pop("revision_admin_usuarios", None)


@st.dialog("Confirmar modificación de la base de datos", on_dismiss=_cerrar_confirmacion_por_x)
def modal_confirmar_cambio_compartido():
    p = st.session_state.get("cambio_pendiente")
    if not p:
        return
    st.subheader(etiqueta_recurso(p["recurso"]))
    st.write(p["descripcion"])
    if p.get("nombre_archivo"):
        st.caption(f"Archivo nuevo: **{p['nombre_archivo']}**")
    anterior = p["previo"]
    st.markdown("**Última modificación registrada:**")
    st.write(f"Fecha: **{gp.mostrar_fecha(anterior['fecha'])}**")
    st.write(f"Usuario: **{anterior['usuario']}**")
    st.caption(f"Revisión registrada: {anterior['revision']}")
    vigente = gp.estado_recurso(DB_PATH, p["recurso"])
    conflicto = vigente["revision"] != p["revision"]
    if conflicto:
        st.error("Esta base cambió mientras preparabas la actualización. Vuelve a abrir la confirmación para revisar la nueva versión. No se guardó nada.")
        st.caption(f"Nueva modificación: {gp.mostrar_fecha(vigente['fecha'])} · {vigente['usuario']}")
    st.warning("Esta operación modifica información compartida por los tres encargados.")
    marcado = st.checkbox("Confirmo que he revisado la información y autorizo la modificación de esta base de datos.",
                          key="confirmacion_" + p["uid"])
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Confirmar y guardar", type="primary", use_container_width=True,
                     disabled=not marcado or conflicto, key="confirmar_" + p["uid"]):
            try:
                resultado = confirmar_cambio_pendiente(p)
            except gp.ConflictoConcurrente as err:
                st.error(str(err))
            except Exception as err:
                st.error(f"No se realizó la actualización. Se conservaron los datos anteriores. Detalle: {err}")
            else:
                st.session_state.pop("cambio_pendiente", None)
                st.session_state["mensaje_cambio_confirmado"] = resultado
                st.rerun()
    with c2:
        if st.button("Cancelar", use_container_width=True, key="cancelar_" + p["uid"]):
            st.session_state.pop("cambio_pendiente", None)
            st.rerun()

# --- DEFINICIÓN DE DIÁLOGOS EMERGENTES DE STREAMLIT (@st.dialog) ---

@st.dialog("Alerta de mantenimiento: Cartera de Exámenes", on_dismiss=_cerrar_alerta_por_x)
def modal_alerta_cartera():
    st.warning("Han transcurrido más de 6 meses desde la última actualización de la Cartera de Exámenes.")
    st.write(
        "Para garantizar la concordancia exacta entre la información del laboratorio y las fichas extraídas, te sugerimos actualizar la base de datos subiendo el archivo Excel más reciente desde el panel lateral."
    )
    if st.button("Entendido / Ocultar notificación emergente", type="primary", use_container_width=True):
        st.session_state["mostrar_dialogo_expiracion"] = False
        st.rerun()


def _conflictos_cartera_actualizados(datos):
    """Compara contra la cartera DESPUÉS de decidir los valores entre PDF."""
    vinculo = st.session_state.get("cartera_examen_emparejado")
    info = obtener_info_cartera_por_nombre(vinculo) if vinculo and vinculo != "Sin vincular" else None
    conflictos = {}
    if info:
        for campo in ("Equipamiento utilizado", "Plazo entrega resultados"):
            pdf = str(datos.get(campo, "") or "").strip()
            cartera = str(info.get(campo, "") or "").strip()
            diferente = (not eq.equipamientos_equivalentes(pdf, cartera)
                         if campo == "Equipamiento utilizado" else
                         pdf.casefold() != cartera.casefold())
            if pdf and cartera and diferente:
                conflictos[campo] = {"pdf": pdf, "cartera": cartera,
                                    "fuente": "Cartera de Exámenes del CRS"}
    if st.session_state.get("area_resultado") == "Derivados":
        fuentes_inst = dict(st.session_state.get("procedencia_valores_documentales", {}))
        origen = st.session_state.get("origen_campos_documentales", {})
        propuestas = st.session_state.get("pdf_para_cartera", {}) or st.session_state.get(
            "procedencia_pdf_campos", {})
        for campo, valor_pdf in propuestas.items():
            if campo == "Código" or campo in conflictos:
                continue
            val_inst = str(fuentes_inst.get(campo, '') or '').strip()
            if de.son_valores_distintos(campo, valor_pdf, val_inst):
                fuente = str(origen.get(campo, '') or '').strip()
                if fuente:
                    conflictos[campo] = {
                        "pdf": str(valor_pdf).strip(), "cartera": val_inst,
                        "fuente": fuente,
                    }
    return conflictos


def _guardar_resolucion_en_ficha(datos, campos_resueltos):
    """La resolución de discrepancias es extracción validada, no edición manual."""
    area = st.session_state.get("area_resultado", st.session_state.get("area_activa"))
    campos = get_field_aliases(area)
    anteriores = set(st.session_state.get("user_edited_fields", set()))
    # Una selección de fuentes PDF/Cartera es validación de extracción.
    # No debe colorearse como una edición directa del usuario.
    anteriores.difference_update(campos_resueltos)
    filas = []
    for campo in campos:
        valor = datos.get(campo, "")
        if valor is None or str(valor).strip().casefold() in ("nan", "none", "null"):
            valor = ""
        filas.append({
            "Campo": campo.strip(),
            "Información extraída": valor,
            "Estado": ("Modificado por usuario" if campo in anteriores
                       else "Encontrado" if valor else "No encontrado"),
        })
    nuevo_excel = aplicar_ediciones_a_excel(
        st.session_state["result_bytes"], pd.DataFrame(filas),
        st.session_state.get("examen_activo", ""),
        campos_forzar=campos_resueltos)
    st.session_state["result_bytes"] = nuevo_excel
    st.session_state["data_extracted"] = datos
    st.session_state["user_edited_fields"] = anteriores
    st.session_state.pop("historial_edicion_tabla", None)
    st.session_state.pop("campos_restaurados_tabla", None)


def _mensaje_discrepancias(cantidad, origen):
    """Mensaje correcto para uno o varios campos."""
    if cantidad == 1:
        return f"Se detectó una discrepancia en 1 campo entre {origen}."
    return f"Se detectaron discrepancias en {cantidad} campos entre {origen}."


def _estilos_tarjeta_discrepancia(clave: str, color: str, seleccionada: bool) -> str:
    """CSS para que cada bloque de discrepancia sea directamente seleccionable."""
    paleta = {
        "azul": ("#193149", "#49a8ff", "#0ea5e9"),
        "verde": ("#163a2b", "#5de398", "#22c55e"),
    }
    fondo, _encabezado, borde = paleta[color]
    titulo_oscuro = "#12518a" if color == "azul" else "#166534"
    fondo_titulo = "#e8f4ff" if color == "azul" else "#e6f8ed"
    borde_real = borde if seleccionada else "#334155"
    return f'''<style>
    .st-key-{clave} button {{
      width: 100%; min-height: 155px !important; height: auto !important;
      background: {fondo} !important; border: 2px solid {borde_real} !important;
      border-radius: 8px; padding: 0.72rem 0.9rem !important;
      display: flex !important; justify-content: flex-start !important; align-items: flex-start !important;
      text-align: left !important; box-shadow: none !important;
    }}
    .st-key-{clave} button > div {{ width: 100% !important; display: flex !important; justify-content: flex-start !important; }}
    .st-key-{clave} button p {{
      color: #ffffff !important; text-align: left !important; white-space: pre-wrap !important;
      width: 100% !important; line-height: 1.52 !important; margin: 0 !important;
      font-weight: 700 !important;
    }}
    .st-key-{clave} button p strong:first-of-type {{
      color: {titulo_oscuro} !important;
      background: {fondo_titulo} !important;
      border-radius: 3px;
      box-decoration-break: clone;
      -webkit-box-decoration-break: clone;
      padding: 0.08rem 0.14rem;
    }}
    .st-key-{clave} button:hover,
    .st-key-{clave} button:focus-visible {{ border-color: {borde} !important; filter: brightness(1.03); }}
    </style>'''


def _boton_valor_discrepancia(titulo, valor, color, clave, seleccionada, on_click, args):
    """Renderiza el bloque azul/verde como un único botón de selección."""
    titulo_seguro = str(titulo or "")
    valor_seguro = str(valor or "(Vacío)")
    etiqueta = f"**{titulo_seguro}**\n\n{valor_seguro}"
    st.markdown(_estilos_tarjeta_discrepancia(clave, color, seleccionada), unsafe_allow_html=True)
    st.button(etiqueta, key=clave, use_container_width=True, on_click=on_click, args=args)


@st.dialog("Comparación y Resolución de Discrepancias", on_dismiss=_cerrar_discrepancias_por_x)
def modal_discrepancias_unificadas():
    """Una ventana y un botón: primero PDF vs PDF, después PDF vs Cartera."""
    pdf_diferencias = st.session_state.get("discrepancias_pdf", {})
    cartera_diferencias = st.session_state.get("conflictos_pendientes", {})
    etapa = "pdf" if pdf_diferencias else "cartera"
    if etapa == "pdf":
        st.markdown(_mensaje_discrepancias(len(pdf_diferencias), "los documentos PDF"))
        st.write("Revisa cada campo y selecciona la información que deseas conservar. **Etapa 1: PDF vs PDF**.")
        diferencias = pdf_diferencias
    else:
        st.markdown(_mensaje_discrepancias(len(cartera_diferencias), "el PDF y las fuentes institucionales"))
        st.write("Por favor, revisa cada campo afectado y selecciona qué información deseas conservar. **PDF vs fuente institucional**.")
        diferencias = cartera_diferencias
    if not diferencias:
        st.info("No hay discrepancias pendientes por resolver.")
        if st.button("Cerrar", key="cerrar_dialogo_unificado"):
            st.session_state["modal_discrepancias_unificado"] = False
            st.rerun()
        return
    st.write("")
    respuestas = {}
    for campo, detalle in diferencias.items():
        st.markdown(f"""<div style="background-color: #f1f5f9; padding: 10px 14px; border-radius: 6px; border-left: 5px solid #0284c7; margin-top: 10px; margin-bottom: 12px;"><span style="font-size: 1.05rem; font-weight: bold; color: #0f172a;">Campo evaluado: <span style="color: #0369a1;">{html.escape(campo)}</span></span></div>""", unsafe_allow_html=True)
        # Ambos tipos de discrepancia comparten dos tarjetas y dos selectores.
        # Cada selector se renderiza DENTRO de la columna de su tarjeta.
        # Los callbacks actualizan el estado sin llamar manualmente a st.rerun().
        if etapa == "pdf":
            opciones = list(detalle)
        else:
            opciones = [
                ("PDF extraído", detalle["pdf"]),
                (str(detalle.get("fuente", "Fuente institucional")), detalle["cartera"]),
            ]

        clave = ("unificado_pdf_" if etapa == "pdf" else "unificado_cartera_") + hashlib.sha256(campo.encode()).hexdigest()[:16]
        if st.session_state.get(clave) not in range(len(opciones)):
            # En lectura óptica el valor del documento es principal, salvo
            # Código o una Cartera seleccionada expresamente.
            es_escaneado_activo = (
                st.session_state.get("procedencia_modalidad") == "Documentos escaneados")
            fuente_elegida = campo in st.session_state.get("fuentes_elegidas_usuario", set())
            st.session_state[clave] = (
                1 if etapa == "cartera" and (
                    not es_escaneado_activo or campo == "Código" or fuente_elegida)
                else 0)

        def _seleccionar_discrepancia(k, indice):
            st.session_state[k] = indice

        columnas = st.columns(len(opciones), gap="medium")
        for indice, (nombre, valor) in enumerate(opciones):
            with columnas[indice]:
                if etapa == "pdf":
                    titulo = f"Valor del PDF ({nombre}):"
                elif indice == 0:
                    titulo = "Valor extraído del PDF:"
                else:
                    titulo = "Valor de " + str(detalle.get("fuente", "Cartera institucional")) + ":"
                clave_tarjeta = f"tarjeta_{etapa}_{hashlib.sha256(campo.encode()).hexdigest()[:12]}_{indice}"
                _boton_valor_discrepancia(
                    titulo, valor, "azul" if indice % 2 == 0 else "verde",
                    clave_tarjeta, st.session_state[clave] == indice,
                    _seleccionar_discrepancia, (clave, indice),
                )
        respuestas[campo] = st.session_state[clave]
        st.divider()
    col1, col2 = st.columns([2, 1])
    with col1:
        etiqueta = "Aplicar y continuar" if etapa == "pdf" else "Aplicar y actualizar ficha"
        if st.button(etiqueta, type="primary", use_container_width=True, key="aplicar_unificado"):
            datos = copy.deepcopy(st.session_state["data_extracted"])
            for campo, eleccion in respuestas.items():
                if etapa == "pdf":
                    datos[campo] = diferencias[campo][eleccion][1]
                else:
                    pdf, cartera = diferencias[campo]["pdf"], diferencias[campo]["cartera"]
                    datos[campo] = pdf if eleccion == 0 else cartera
            try:
                _guardar_resolucion_en_ficha(datos, set(respuestas))
                if etapa == "cartera":
                    elegidas = st.session_state.setdefault("fuentes_elegidas_usuario", set())
                    for campo, eleccion in respuestas.items():
                        if eleccion == 1:
                            elegidas.add(campo)
                            if "Cartera de Exámenes" in diferencias[campo].get("fuente", ""):
                                # Atribuir a la Cartera el valor que el usuario
                                # acaba de elegir, aunque Escáner lo propusiera
                                # inicialmente y la selección fuese posterior.
                                st.session_state.setdefault("procedencia_cartera_campos", {})[
                                    campo] = diferencias[campo]["cartera"]
                                st.session_state["cartera_aporto_campos"] = True
                        else:
                            elegidas.discard(campo)
            except Exception as exc:
                st.error(f"No se pudieron aplicar las selecciones: {exc}")
                return
            if etapa == "pdf":
                # Fuente PDF que el usuario acaba de validar: se utilizará
                # al comparar después con las bases institucionales.
                propuestas_pdf_resueltas = dict(
                    st.session_state.get("procedencia_pdf_campos", {}))
                for campo_pdf, indice_pdf in respuestas.items():
                    propuestas_pdf_resueltas[campo_pdf] = diferencias[campo_pdf][indice_pdf][1]
                st.session_state["pdf_para_cartera"] = propuestas_pdf_resueltas
                st.session_state.pop("discrepancias_pdf", None)
                # No reutilizar conflictos calculados con valores anteriores.
                nuevos = _conflictos_cartera_actualizados(datos)
                if nuevos:
                    st.session_state["conflictos_pendientes"] = nuevos
                    st.session_state["modal_discrepancias_unificado"] = True
                else:
                    st.session_state.pop("conflictos_pendientes", None)
                    st.session_state["modal_discrepancias_unificado"] = False
            else:
                st.session_state.pop("conflictos_pendientes", None)
                st.session_state["modal_discrepancias_unificado"] = False
            st.rerun()
    with col2:
        if st.button("Cancelar", use_container_width=True, key="cancelar_unificado"):
            st.session_state["modal_discrepancias_unificado"] = False
            st.rerun()


@st.dialog("Confirmar actualización en base de datos", width="large", on_dismiss=_cerrar_planilla_por_x)
def modal_confirmar_actualizacion_planilla():
    area_activa = st.session_state.get("area_activa")
    cierre_omision = bool(st.session_state.get("cola_pdf_finalizada_omision"))
    examen_activo = (st.session_state.get("examen_resultado") if cierre_omision
                     else st.session_state.get("examen_activo"))
    comp = st.session_state.get("comparacion_planilla_detalles", {})

    if cierre_omision:
        st.markdown(
            f"Estás a punto de actualizar la **base de datos del área {area_activa}** "
            "con las fichas efectivamente procesadas en esta cola. "
            "Los documentos omitidos no se han analizado ni modificado."
        )
    else:
        st.markdown(f"Estás a punto de reemplazar la **base de datos** para el área **{area_activa}** con los datos recién procesados del examen **{examen_activo}**.")
    st.write("")

    col_m1, col_m2, col_m3 = st.columns(3)
    with col_m1:
        st.metric(
            label="Campos de las fichas" if cierre_omision else "Campos del examen",
            value=f"{comp.get('campos_completados_despues', 0)} completados",
            delta=f"{comp.get('campos_completados_despues', 0) - comp.get('campos_completados_antes', 0)} vs versión anterior"
        )
    with col_m2:
        st.metric(
            label="Exámenes/Hojas en área",
            value=f"{comp.get('hojas_despues', 0)} hojas",
            delta=(f"{comp.get('hojas_despues', 0) - comp.get('hojas_antes', 0):+d} hojas"
                   if cierre_omision else
                   "+1 nuevo examen" if comp.get("es_nueva_hoja") else "Misma cantidad")
        )
    with col_m3:
        st.metric(
            label="Campos con cambios",
            value=f"{len(comp.get('diferencias', []))} modificados"
        )

    st.divider()

    diferencias = comp.get("diferencias", [])
    if diferencias:
        st.markdown("##### Resumen de diferencias detectadas:")
        df_dif = pd.DataFrame(diferencias)
        # El contenido de las fichas puede incluir párrafos completos. Una tabla
        # estática comprime sus columnas dentro del diálogo; el visor desplazable
        # conserva el contenido íntegro y se adapta al ancho disponible.
        st.dataframe(
            df_dif[["Campo", "Tipo de cambio", "Valor anterior", "Valor nuevo"]],
            use_container_width=True,
            hide_index=True,
            height=min(480, max(160, 36 * (len(df_dif) + 2))),
            column_config={
                "Campo": st.column_config.TextColumn("Campo", width=180),
                "Tipo de cambio": st.column_config.TextColumn("Tipo de cambio", width=190),
                "Valor anterior": st.column_config.TextColumn("Valor anterior", width=260),
                "Valor nuevo": st.column_config.TextColumn("Valor nuevo", width=260),
            },
        )
    else:
        st.info("No se detectaron diferencias en los campos. La versión actual y la nueva son idénticas.")

    st.write("")
    st.warning("Esta acción sobreescribirá la Ficha Técnica / Base de Datos del área para que los próximos análisis utilicen este nuevo contenido.")

    recurso = "area:" + area_activa
    # La revisión se capturó al pulsar «Actualizar en base de datos», no al guardar.
    esperada = st.session_state.get("revision_plantilla_pendiente", 0)
    registro = st.session_state.get("metadata_plantilla_pendiente") or gp.estado_recurso(DB_PATH, recurso)
    st.markdown("**Última modificación registrada:**")
    st.write(f"Fecha: **{gp.mostrar_fecha(registro['fecha'])}**")
    st.write(f"Usuario: **{registro['usuario']}**")
    vigente = gp.estado_recurso(DB_PATH, recurso)
    conflicto = vigente["revision"] != esperada
    if conflicto:
        st.error(f"Otro usuario actualizó esta área: {gp.mostrar_fecha(vigente['fecha'])} · {vigente['usuario']}. Cancela y vuelve a revisar el examen.")
    validado = st.checkbox("Confirmo que revisé los cambios y autorizo actualizar esta ficha compartida.",
                            key="confirmar_planilla_generada")
    col_save1, col_save2 = st.columns([2, 1])
    with col_save1:
        if st.button("Confirmar y actualizar base de datos del área", type="primary", use_container_width=True,
                     key="btn_confirmar_guardado_ficha", disabled=not validado or conflicto):
            try:
                nombre_generado = nombre_ficha_tecnica(
                    area_activa, st.session_state.get("fecha_ultima_extraccion_ficha"))
                contenido = st.session_state.get("excel_para_actualizar", st.session_state["result_bytes"])
                def guardar_conexion(con):
                    ok, mensaje = guardar_plantilla_area_db(area_activa, contenido, nombre_generado, conexion=con)
                    if not ok:
                        raise ValueError(mensaje)
                    return mensaje
                gp.respaldar(DB_PATH)
                mensaje = gp.ejecutar_cambio(DB_PATH, recurso, st.session_state["usuario_acceso"],
                    esperada, "Guardar ficha técnica generada para " + (examen_activo or "examen"),
                    guardar_conexion, archivo=nombre_generado, detalle="Actualización desde la ficha generada")
                st.session_state["examenes_por_area"] = cargar_examenes_por_area()
                try:
                    gp.respaldar(DB_PATH)
                except Exception as respaldo_error:
                    st.session_state["mensaje_respaldo"] = f"Atención: cambio guardado, pero falló su respaldo: {respaldo_error}"
                st.session_state["modal_confirmar_planilla_abierto"] = False
                st.session_state.pop("comparacion_planilla_detalles", None)
                st.session_state.pop("excel_para_actualizar", None)
                st.session_state.pop("revision_plantilla_pendiente", None)
                st.session_state.pop("metadata_plantilla_pendiente", None)
                st.session_state["mensaje_cambio_confirmado"] = mensaje
                st.rerun()
            except Exception as err_save:
                st.error(f"No se actualizó la ficha. {err_save}")
    with col_save2:
        if st.button("Cancelar", use_container_width=True, key="btn_cancelar_guardado_ficha"):
            st.session_state["modal_confirmar_planilla_abierto"] = False
            st.session_state.pop("excel_para_actualizar", None)
            st.rerun()


def _limpiar_seleccion_area(area, key_sel):
    """Limpiar widgets en un callback previo al rerun (evita AlreadyInstantiatedError)."""
    st.session_state[key_sel] = None
    st.session_state["area_activa"] = area
    st.session_state["examen_activo"] = None
    st.session_state["cartera_examen_emparejado"] = "Sin vincular"
    st.session_state["acreditacion_manual_id"] = None
    st.session_state["_acreditacion_reset_pendiente"] = True
    st.session_state["_sel_cartera_pendiente"] = None
    st.session_state.pop("candidatos_cartera_pendientes", None)
    st.session_state["mostrar_modal_examenes"] = False
    st.session_state.pop("modal_area_activa", None)


@st.dialog("Seleccionar y administrar exámenes", on_dismiss=_cerrar_examenes_por_x)
def modal_examenes_area(area):
    st.subheader(area, anchor=False)
    examenes = st.session_state["examenes_por_area"].get(area, [])
    actual = st.session_state.get("examen_activo") if st.session_state.get("area_activa") == area else None
    
    key_sel = f"selector_examen_{area}"
    if key_sel not in st.session_state and actual in examenes:
        st.session_state[key_sel] = actual

    seleccionado = st.selectbox(
        "Examen de esta área", 
        examenes, 
        index=None, 
        placeholder="Selecciona un examen",
        key=key_sel,
        help="El nombre completo se conservará en Nombre LIS. Cuando es largo, la herramienta utiliza un título de hoja compatible con Excel."
    )
    
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Usar examen", key=f"usar_examen_{area}", type="primary", use_container_width=True):
            if seleccionado:
                st.session_state["area_activa"] = area
                st.session_state["examen_activo"] = seleccionado
                st.session_state["cartera_examen_emparejado"] = "Sin vincular"
                st.session_state.pop("contexto_emparejamiento_cartera", None)
                st.session_state.pop("cartera_seleccion_confirmada_manualmente", None)
                st.session_state["acreditacion_manual_id"] = None
                st.session_state["_acreditacion_reset_pendiente"] = True
                st.session_state["_sel_cartera_pendiente"] = None
                st.session_state.pop("candidatos_cartera_pendientes", None)
                st.session_state["mostrar_modal_examenes"] = False
                st.session_state.pop("modal_area_activa", None)
                st.rerun()
            else:
                st.warning("Selecciona un examen o agrega uno nuevo.")
    with c2:
        # El callback se ejecuta antes de instanciar selectbox en el siguiente rerun.
        # Es inválido asignar a selector_examen_... en el cuerpo, después del widget.
        if st.button("Limpiar selección", key=f"limpiar_examen_{area}",
                     use_container_width=True, on_click=_limpiar_seleccion_area,
                     args=(area, key_sel)):
            st.rerun()

    st.divider()
    with st.expander("Añadir examen"):
        with st.form(f"form_anadir_examen_{area}", clear_on_submit=True):
            nuevo = st.text_input("Nombre exacto del examen", placeholder="Ej.: Hemograma automatizado")
            agregar = st.form_submit_button("Guardar examen")
        if agregar:
            nuevo = nuevo.strip()
            valido, mensaje = nombre_examen_valido(nuevo)
            actuales = st.session_state["examenes_por_area"].get(area, [])
            if not valido:
                st.warning(mensaje)
            elif nuevo.casefold() in {x.casefold() for x in actuales}:
                st.warning("Ese examen ya existe en esta área. Selecciónalo en el menú.")
            else:
                actuales.append(nuevo)
                solicitar_cambio("examenes", "area:" + area,
                    descripcion=f"Agregar el examen «{nuevo}» al área {area}",
                    area=area, nombres=sorted(actuales, key=clave_orden_alfabetico),
                    seleccionar_examen=nuevo)
    with st.expander("Eliminar examen de la lista"):
        actuales = st.session_state["examenes_por_area"].get(area, [])
        if actuales:
            quitar = st.selectbox(
                "Examen que deseas quitar", 
                actuales, 
                index=None,
                placeholder="Selecciona un examen para eliminar",
                key=f"quitar_examen_{area}"
            )
            confirmar = st.checkbox("Confirmo que quiero quitarlo del menú", key=f"confirmar_quitar_{area}")
            if st.button("Eliminar examen", key=f"eliminar_examen_{area}", disabled=not confirmar or not quitar):
                nombres_nuevos = [x for x in actuales if x.casefold() != quitar.casefold()]
                solicitar_cambio("examenes", "area:" + area,
                    descripcion=f"Eliminar el examen «{quitar}» del área {area}",
                    area=area, nombres=nombres_nuevos, quitar_examen=quitar)
        else:
            st.info("No hay exámenes para eliminar.")


# Únicamente datos temporales de la mesa de trabajo: ni las cuentas de acceso,
# ni los catálogos ni los archivos guardados en SQLite forman parte del reinicio.
CLAVES_TRABAJO_AREA = (
    "examen_activo", "examen_resultado", "area_resultado",
    "cartera_examen_emparejado", "sel_cartera_main", "_sel_cartera_pendiente",
    "acreditacion_manual_id", "acreditacion_manual_selector",
    "contexto_emparejamiento_acreditacion",
    "acreditacion_seleccion_confirmada_manualmente",
    "acreditacion_coincidencia_automatica",
    "_acreditacion_reset_pendiente", "candidatos_cartera_pendientes",
    "contexto_emparejamiento_cartera", "cartera_seleccion_confirmada_manualmente",
    "omitir_emparejamiento_examen", "candidato_cartera_inline",
    "modo_referencias_apa", "modo_referencia_aplicado", "cartera_aporto_campos",
    "pdf_aporto_campos", "pdf_documentos_con_aporte",
    "texto_referencias_auto", "fuentes_referencias_aplicadas",
    "cache_apa_pdf",
    "accion_cartera_sin_coincidencia", "confirmar_reemplazo_referencias",
    "laboratorio_activo", "sel_lab_main", "actualizado_activo", "sel_act_main",
    "autorizado_activo", "sel_aut_main",
    "data_extracted", "user_edited_fields", "result_bytes", "extracted_text",
    "fecha_ultima_extraccion_ficha", "seleccion_campos_excel", "excel_para_actualizar",
    "raw_extracted_text", "fuentes_pdf", "discrepancias_pdf",
    "conflictos_pendientes", "fuentes_elegidas_usuario", "texto_optico_original",
    "origen_campos_documentales",
    "prestacion_acreditacion_vinculada", "estado_vinculacion_acreditacion",
    "bytes_plantilla_inicial", "revision_plantilla_origen",
    "revision_plantilla_pendiente", "metadata_plantilla_pendiente",
    "comparacion_planilla_detalles", "confirmar_planilla_generada",
    "orden_nombres_pdf", "modo_pdf_anterior", "modo_procesamiento_pdf", "modo_solo_base_datos",
    "cola_pdf_omitidos", "cola_pdf_procesados", "cola_pdf_finalizada_omision",
    "historial_edicion_tabla", "campos_restaurados_tabla",
    "pdf_preview_documento_conjunto", "imported_excel_format",
    "imported_excel_sheet", "del_lab_select", "del_act_select",
    "del_aut_select", "nombre_nuevo_laboratorio", "nombre_nuevo_actualizado",
    "nombre_nuevo_autorizado", "mostrar_modal_examenes", "modal_area_activa",
    "modal_confirmar_planilla_abierto", "modal_discrepancias_unificado",
    "modal_conflictos_abierto", "modal_pdf_pdf_abierto",
)


def reiniciar_trabajo_por_cambio_area():
    """Descarta el trabajo temporal al cambiar de área (se ejecuta como callback).

    Streamlit no permite asignar ni borrar el valor de un file_uploader instanciado.
    Por eso se aumenta la generación de sus keys y el navegador renderiza nuevos
    controles vacíos. El callback corre *antes* de instanciar los demás widgets.
    """
    for clave in CLAVES_TRABAJO_AREA:
        st.session_state.pop(clave, None)

    # No debe quedar seleccionada ninguna prestación al regresar a un área previa.
    for area in AREA_CONFIG:
        st.session_state.pop(f"selector_examen_{area}", None)
        st.session_state.pop(f"quitar_examen_{area}", None)
        st.session_state.pop(f"confirmar_quitar_{area}", None)

    st.session_state["cola_pdf_indice"] = 0
    st.session_state["aggrid_key_counter"] = st.session_state.get("aggrid_key_counter", 0) + 1
    st.session_state["generacion_cargas_area"] = st.session_state.get("generacion_cargas_area", 0) + 1
    st.session_state["_acreditacion_reset_pendiente"] = True
    st.session_state["mostrar_modal_examenes"] = False


def progreso_pendiente_cambio_area(estado):
    """Distingue archivos en widgets de campos realmente elegidos o completados."""
    generacion = estado.get("generacion_cargas_area", 0)
    archivos = any(bool(estado.get(clave)) for clave in (
        f"pdf_uploader_{generacion}",
        f"excel_exportacion_uploader_{generacion}",
    ))
    campos = any(bool(estado.get(clave)) for clave in (
        "examen_activo", "laboratorio_activo", "actualizado_activo",
        "autorizado_activo", "data_extracted", "user_edited_fields",
        "result_bytes", "extracted_text",
    ))
    return archivos, campos


def seleccionar_area(area):
    """Solicita confirmación solo si cambiar de área descartaría trabajo temporal."""
    if area not in AREA_CONFIG:
        raise ValueError("Área de laboratorio desconocida")
    actual = st.session_state.get("area_activa")
    if actual == area:
        return  # Pulsar el área ya activa no descarta avances.
    archivos, campos = progreso_pendiente_cambio_area(st.session_state)
    if actual and (archivos or campos):
        st.session_state["cambio_area_pendiente"] = {
            "destino": area, "archivos": archivos, "campos": campos,
        }
        return
    reiniciar_trabajo_por_cambio_area()
    st.session_state["area_activa"] = area


def _cancelar_cambio_area():
    st.session_state.pop("cambio_area_pendiente", None)


@st.dialog("Confirmar cambio de área", on_dismiss=_cancelar_cambio_area)
def modal_confirmar_cambio_area():
    cambio = st.session_state.get("cambio_area_pendiente")
    if not cambio:
        return
    if cambio["archivos"]:
        st.warning("Tienes archivos adjuntados.")
    if cambio["campos"]:
        st.warning("Tienes campos seleccionados o completados.")
    st.write("¿Seguro que quieres cambiar de área? Se borrará el progreso si continúas.")
    st.markdown("""<style>
        .st-key-confirmar_cambio_area button {
            background: #1677c8 !important; border-color: #1677c8 !important;
            color: white !important;
        }
        .st-key-cancelar_cambio_area button {
            background: #c62828 !important; border-color: #c62828 !important;
            color: white !important;
        }
    </style>""", unsafe_allow_html=True)
    col_continuar, col_cancelar = st.columns(2)
    with col_continuar:
        if st.button("Continuar", type="primary", use_container_width=True,
                     key="confirmar_cambio_area"):
            destino = cambio["destino"]
            _cancelar_cambio_area()
            reiniciar_trabajo_por_cambio_area()
            st.session_state["area_activa"] = destino
            st.rerun()
    with col_cancelar:
        if st.button("Cancelar", use_container_width=True,
                     key="cancelar_cambio_area"):
            _cancelar_cambio_area()
            st.rerun()


# Botón flotante: dos pulsaciones, 3 segundos de giro para confirmar;
# sin segunda pulsación cancela silenciosamente, sin X ni rojo.
# Solo descarta datos de st.session_state; nunca modifica SQLite.
SEGUNDOS_CONFIRMACION_REINICIO = 3.0
SEGUNDOS_FLASH_CONFIRMACION = 1.5


def _iniciar_reinicio_flotante():
    st.session_state["estado_reinicio_flotante"] = "confirmar"
    # El límite comienza inmediatamente al hacer clic, permitiendo aceptar
    # incluso un segundo clic que llegue antes de redibujar la flecha.
    st.session_state["vencimiento_reinicio_flotante"] = (
        time.monotonic() + SEGUNDOS_CONFIRMACION_REINICIO
    )
    st.session_state.pop("fin_flash_reinicio_flotante", None)
    # Alterna el nombre CSS de la animación si la confirmación anterior expiró:
    # permite reiniciar el giro sin refrescos automáticos cada segundo.
    st.session_state["reinicio_flotante_intento"] = (
        st.session_state.get("reinicio_flotante_intento", 0) + 1
    ) % 2


def _confirmar_reinicio_flotante():
    """Callback previo a los widgets: los uploaders pueden reiniciarse sin error."""
    if st.session_state.get("estado_reinicio_flotante") != "confirmar":
        return
    limite = st.session_state.get("vencimiento_reinicio_flotante", 0)
    if time.monotonic() >= limite:
        st.session_state["estado_reinicio_flotante"] = "inactivo"
        st.session_state.pop("vencimiento_reinicio_flotante", None)
        return
    reiniciar_trabajo_por_cambio_area()
    st.session_state["estado_reinicio_flotante"] = "completado"
    st.session_state.pop("vencimiento_reinicio_flotante", None)
    # El efecto empieza cuando se dibuja el botón tras la actualización completa;
    # si un PDF ralentiza esa actualización, el verde seguirá viéndose 1,5 s.
    st.session_state["fin_flash_reinicio_flotante"] = None
    # Se actualiza la pantalla completa una única vez tras la confirmación;
    # mientras el usuario no interactúa, no se lanza ningún refresco.
    st.session_state["reinicio_flotante_refrescar_pantalla"] = True


def _pulsar_reinicio_flotante():
    """Misma función y misma key para ambas pulsaciones (también muy rápidas)."""
    estado = _actualizar_tiempo_reinicio_flotante()
    if estado == "inactivo":
        _iniciar_reinicio_flotante()
    elif estado == "confirmar":
        _confirmar_reinicio_flotante()


def _actualizar_tiempo_reinicio_flotante():
    """Transiciones limitadas a la sesión que pulsó el botón."""
    estado = st.session_state.get("estado_reinicio_flotante", "inactivo")
    ahora = time.monotonic()
    limite = st.session_state.get("vencimiento_reinicio_flotante")
    fin_flash = st.session_state.get("fin_flash_reinicio_flotante")
    if estado == "confirmar" and limite is not None and ahora >= limite:
        st.session_state["estado_reinicio_flotante"] = "inactivo"
        st.session_state.pop("vencimiento_reinicio_flotante", None)
    elif estado == "completado" and fin_flash is not None and ahora >= fin_flash:
        st.session_state["estado_reinicio_flotante"] = "inactivo"
        st.session_state.pop("fin_flash_reinicio_flotante", None)
    return st.session_state.get("estado_reinicio_flotante", "inactivo")


@st.fragment
def _reinicio_flotante_rapido():
    """Actualiza el control al pulsarlo, nunca por un temporizador periódico."""
    estado = _actualizar_tiempo_reinicio_flotante()
    if st.session_state.pop("reinicio_flotante_refrescar_pantalla", False):
        st.rerun(scope="app")
    if estado == "completado" and st.session_state.get("fin_flash_reinicio_flotante") is None:
        st.session_state["fin_flash_reinicio_flotante"] = (
            time.monotonic() + SEGUNDOS_FLASH_CONFIRMACION
        )
    with st.container(key="reinicio_flotante"):
        # Marcador CSS sin contenido visible: al comenzar otra confirmación
        # alterna la animación (mismo botón y misma key en Streamlit).
        intento = st.session_state.get("reinicio_flotante_intento", 0)
        st.markdown(
            f'<span class="estado-reinicio-{estado} estado-reinicio-intento-{intento}" '
            'style="display:none" aria-hidden="true"></span>',
            unsafe_allow_html=True,
        )
        st.button(
            "Pulsa dos veces para reiniciar",
            icon=":material/refresh:",
            type="secondary",
            help="Pulsa dos veces para reiniciar",
            key="boton_reinicio_flotante",  # clave CONSTANTE incluso al pulsar muy rápido
            # Durante el destello, el callback ignora pulsaciones; después
            # permite una nueva confirmación sin esperar a otro rerun.
            on_click=_pulsar_reinicio_flotante,
        )


def mostrar_reinicio_flotante():
    """Solo en área activa, sin afectar los datos guardados ni otras sesiones."""
    if st.session_state.get("area_activa"):
        _reinicio_flotante_rapido()


def _tabla_trazabilidad(tipo, limite, desplazamiento=0):
    """Trae solo la página visible, en orden cronológico inverso."""
    if tipo == "modificaciones":
        filas = gp.historial(DB_PATH, limite=limite, desplazamiento=desplazamiento)
        columnas = ["Fecha", "Usuario", "Recurso", "Acción", "Revisión", "Archivo"]
    else:
        filas = gp.historial_actividad(DB_PATH, limite=limite, desplazamiento=desplazamiento)
        columnas = ["Fecha", "Usuario", "Actividad", "Detalle"]
    tabla = pd.DataFrame(filas, columns=columnas)
    if not tabla.empty:
        tabla["Fecha"] = tabla["Fecha"].map(gp.mostrar_fecha)
        if tipo == "modificaciones":
            tabla["Recurso"] = tabla["Recurso"].map(etiqueta_recurso)
    return tabla


def _cerrar_trazabilidad_por_x():
    st.session_state.pop("trazabilidad_pantalla_completa", None)


@st.dialog("Trazabilidad — pantalla completa", width="large", on_dismiss=_cerrar_trazabilidad_por_x)
def modal_trazabilidad_completa():
    tipo = st.session_state.get("trazabilidad_pantalla_completa")
    if tipo not in ("modificaciones", "actividad"):
        return
    st.subheader("Historial de modificaciones" if tipo == "modificaciones"
                 else "Actividad de usuarios", anchor=False)
    total = gp.total_historial(DB_PATH) if tipo == "modificaciones" else gp.total_historial_actividad(DB_PATH)
    cantidad_paginas = max(1, (total + 9) // 10)
    clave = f"traza_pagina_{tipo}"
    pagina = max(1, min(int(st.session_state.get(clave, 1)), cantidad_paginas))
    st.session_state[clave] = pagina
    tabla = _tabla_trazabilidad(tipo, 10, (pagina - 1) * 10)
    if not tabla.empty:
        st.dataframe(tabla, hide_index=True, use_container_width=True, height=420)
    else:
        st.info("No existen registros en este historial.")
    # Páginas debajo de los registros; cambiar de página solo actualiza el modal.
    izquierda, medio, derecha = st.columns([1, 2, 1])
    with izquierda:
        if st.button("Anterior", key=f"traza_anterior_{tipo}",
                     disabled=pagina <= 1, use_container_width=True):
            st.session_state[clave] = pagina - 1
            st.rerun(scope="fragment")
    with medio:
        st.markdown(f"<p style='text-align:center'>Página <b>{pagina}</b> de <b>{cantidad_paginas}</b> · {total} registros</p>",
                    unsafe_allow_html=True)
    with derecha:
        if st.button("Siguiente", key=f"traza_siguiente_{tipo}",
                     disabled=pagina >= cantidad_paginas, use_container_width=True):
            st.session_state[clave] = pagina + 1
            st.rerun(scope="fragment")
    if st.button("Cerrar", key=f"traza_cerrar_{tipo}"):
        st.session_state.pop("trazabilidad_pantalla_completa", None)
        st.rerun()


# --- ADMINISTRACIÓN DE USUARIOS DESDE ACCESO ---
@st.dialog("Administración de usuarios", width="large", on_dismiss=_cerrar_administracion_por_x)
def modal_administracion_usuarios():
    actor = st.session_state.get("usuario_acceso")
    if not gp.es_administrador(DB_PATH, actor):
        st.error("No tienes permisos para administrar usuarios.")
        return
    recurso = "usuarios:administracion"
    vigente = gp.estado_recurso(DB_PATH, recurso)
    revision_vista = st.session_state.get("revision_admin_usuarios", vigente["revision"])
    st.caption(f"Última modificación: **{gp.mostrar_fecha(vigente['fecha'])}** · "
               f"Responsable: **{vigente['usuario']}**")
    if vigente["revision"] != revision_vista:
        st.warning("Otro administrador modificó las cuentas. Actualiza la lista antes de continuar.")
        if st.button("Actualizar listado de usuarios", key="admin_recargar_lista"):
            st.session_state["revision_admin_usuarios"] = vigente["revision"]
            st.rerun()
        return

    try:
        cuentas = gp.listar_usuarios(DB_PATH, actor)
    except PermissionError as exc:
        st.error(str(exc))
        return
    # Tabla a ancho completo del modal para que se vean sus cinco encabezados.
    if cuentas:
        vista = pd.DataFrame([{
            "Usuario": c["usuario"], "Nombre visible": c["nombre_visible"],
            "Rol": "Administrador" if c["rol"] == "administrador" else "Encargado",
            "Estado": "Activo" if c["activo"] else "Inactivo",
            "Creado": gp.mostrar_fecha(c["creado_en"]),
        } for c in cuentas])
        st.dataframe(
            vista, use_container_width=True, hide_index=True,
            column_config={
                "Usuario": st.column_config.TextColumn("Usuario", width=115),
                "Nombre visible": st.column_config.TextColumn("Nombre visible", width=235),
                "Rol": st.column_config.TextColumn("Rol", width=135),
                "Estado": st.column_config.TextColumn("Estado", width=110),
                "Creado": st.column_config.TextColumn("Creado", width=170),
            },
        )
    st.caption("Las contraseñas no se pueden consultar. Solo se pueden restablecer.")

    def confirmar_operacion(operacion, objetivo, **parametros):
        try:
            # Una copia antes y otra después permiten recuperar operaciones fallidas.
            gp.respaldar(DB_PATH)
            resultado = gp.administrar_cuenta(
                DB_PATH, actor, revision_vista, operacion, objetivo, **parametros)
            try:
                gp.respaldar(DB_PATH)
            except OSError as exc_respaldo:
                st.session_state["mensaje_respaldo"] = (
                    f"Atención: cuenta actualizada, pero falló su respaldo: {exc_respaldo}")
            st.session_state["modal_admin_usuarios_abierto"] = False
            st.session_state.pop("revision_admin_usuarios", None)
            st.session_state["mensaje_cambio_confirmado"] = resultado
            st.rerun()
        except (ValueError, PermissionError, gp.ConflictoConcurrente,
                sqlite3.Error, OSError) as exc:
            st.error(f"No se modificó la cuenta: {exc}")

    tab_crear, tab_editar, tab_clave, tab_estado = st.tabs(
        ["Crear usuario", "Editar usuario", "Editar contraseña", "Activar/desactivar"])
    roles = {"Encargado": "encargado", "Administrador": "administrador"}
    with tab_crear:
        _margen_crear_izq, contenido_crear, _margen_crear_der = st.columns([0.13, 0.74, 0.13], gap="small")
        with contenido_crear:
            with st.form("admin_crear_usuario"):
                col1, col2 = st.columns(2, gap="medium")
                with col1:
                    usuario_nuevo = st.text_input("Usuario de acceso", key="admin_nuevo_acceso")
                with col2:
                    nombre_nuevo = st.text_input("Nombre visible", key="admin_nuevo_nombre")
                col3, col4 = st.columns(2, gap="medium")
                with col3:
                    clave_nueva = st.text_input("Contraseña (mínimo 12 caracteres)", type="password", key="admin_nueva_clave")
                with col4:
                    repite_nueva = st.text_input("Repetir contraseña", type="password", key="admin_repite_nueva")
                rol_nuevo = st.selectbox("Rol (opcional, por defecto Encargado)", list(roles),
                                        index=None, placeholder="Selecciona un rol", key="admin_nuevo_rol")
                aceptar_nuevo = st.checkbox("Confirmo la creación de esta cuenta.", key="admin_crear_ok")
                guardar_nuevo = st.form_submit_button("Crear usuario", type="primary")
            if guardar_nuevo:
                if not aceptar_nuevo:
                    st.error("Marca la casilla de confirmación.")
                elif clave_nueva != repite_nueva:
                    st.error("Las contraseñas no coinciden.")
                else:
                    confirmar_operacion("crear", usuario_nuevo, nombre_visible=nombre_nuevo,
                                        rol=roles.get(rol_nuevo, "encargado"), clave=clave_nueva)

    nombres = [x["usuario"] for x in cuentas]
    detalle = {x["usuario"]: x for x in cuentas}
    with tab_editar:
        _margen_editar_izq, contenido_editar, _margen_editar_der = st.columns([0.13, 0.74, 0.13], gap="small")
        with contenido_editar:
            sel_editar = st.selectbox("Cuenta que deseas editar", nombres,
                                     index=None, placeholder="Selecciona un usuario", key="admin_sel_editar")
            if sel_editar:
                actual = detalle[sel_editar]
                st.warning("Si cambias el nombre de acceso, las sesiones abiertas de esa cuenta dejarán de ser válidas. Deberá iniciar sesión con el nuevo usuario.")
                with st.form(f"admin_editar_{sel_editar}"):
                    col1, col2 = st.columns(2, gap="medium")
                    with col1:
                        usuario_editado = st.text_input("Usuario de acceso", value=actual["usuario"])
                    with col2:
                        nombre_editado = st.text_input("Nombre visible", value=actual["nombre_visible"])
                    st.caption(f"Rol actual: {'Administrador' if actual['rol'] == 'administrador' else 'Encargado'}")
                    rol_editado = st.selectbox("Cambiar rol (opcional)", list(roles),
                        index=None, placeholder="Selecciona un rol")
                    aceptar_editar = st.checkbox("Confirmo los cambios de esta cuenta.")
                    guardar_editar = st.form_submit_button("Guardar cambios", type="primary")
                if guardar_editar:
                    if aceptar_editar:
                        confirmar_operacion("editar", sel_editar, nombre_visible=nombre_editado,
                                            rol=roles.get(rol_editado, actual["rol"]),
                                            nuevo_usuario=usuario_editado)
                    else:
                        st.error("Marca la casilla de confirmación.")

    with tab_clave:
        _margen_clave_izq, contenido_clave, _margen_clave_der = st.columns([0.13, 0.74, 0.13], gap="small")
        with contenido_clave:
            sel_clave = st.selectbox("Cuenta cuya contraseña deseas restablecer", nombres,
                                    index=None, placeholder="Selecciona un usuario", key="admin_sel_clave")
            if sel_clave:
                st.warning("Al cambiar la contraseña se cerrarán las sesiones actuales de esta cuenta. Será necesario volver a iniciar sesión.")
                with st.form(f"admin_clave_{sel_clave}"):
                    col1, col2 = st.columns(2, gap="medium")
                    with col1:
                        clave1 = st.text_input("Nueva contraseña (mínimo 12 caracteres)", type="password")
                    with col2:
                        clave2 = st.text_input("Repetir nueva contraseña", type="password")
                    aceptar_clave = st.checkbox("Confirmo el restablecimiento de la contraseña.")
                    cambiar_clave = st.form_submit_button("Restablecer contraseña", type="primary")
                if cambiar_clave:
                    if not aceptar_clave:
                        st.error("Marca la casilla de confirmación.")
                    elif clave1 != clave2:
                        st.error("Las contraseñas no coinciden.")
                    else:
                        confirmar_operacion("clave", sel_clave, clave=clave1)

    with tab_estado:
        _margen_estado_izq, contenido_estado, _margen_estado_der = st.columns([0.13, 0.74, 0.13], gap="small")
        with contenido_estado:
            sel_estado = st.selectbox("Cuenta que deseas activar o desactivar", nombres,
                                     index=None, placeholder="Selecciona un usuario", key="admin_sel_estado")
            if sel_estado:
                c = detalle[sel_estado]
                texto = "Desactivar" if c["activo"] else "Activar"
                st.info(f"Estado actual: {'Activo' if c['activo'] else 'Inactivo'}")
                with st.form(f"admin_estado_{sel_estado}"):
                    aceptar_estado = st.checkbox(f"Confirmo que deseo {texto.lower()} la cuenta {sel_estado}.")
                    ejecutar_estado = st.form_submit_button(f"{texto} usuario", type="primary")
                if ejecutar_estado:
                    if aceptar_estado:
                        confirmar_operacion("estado", sel_estado, activo=not bool(c["activo"]))
                    else:
                        st.error("Marca la casilla de confirmación.")
    if st.button("Cerrar administración", key="admin_cerrar_modal"):
        st.session_state["modal_admin_usuarios_abierto"] = False
        st.session_state.pop("revision_admin_usuarios", None)
        st.rerun()


# --- GESTOR DE DIÁLOGOS EMERGENTES ---
def ejecutar_gestor_dialogos():
    """Ejecuta como máximo UN solo diálogo emergente por ciclo para evitar errores de Streamlit."""
    if st.session_state.get("trazabilidad_pantalla_completa"):
        modal_trazabilidad_completa()
    elif st.session_state.get("cambio_area_pendiente"):
        modal_confirmar_cambio_area()
    elif st.session_state.get("modal_admin_usuarios_abierto"):
        modal_administracion_usuarios()
    elif st.session_state.get("cambio_pendiente"):
        modal_confirmar_cambio_compartido()
    elif st.session_state.get("modal_confirmar_planilla_abierto"):
        st.session_state["modal_confirmar_planilla_abierto"] = False
        modal_confirmar_actualizacion_planilla()
    elif st.session_state.get("modal_discrepancias_unificado"):
        modal_discrepancias_unificadas()
    elif st.session_state.get("mostrar_modal_examenes") and st.session_state.get("modal_area_activa") in AREA_CONFIG:
        modal_examenes_area(st.session_state["modal_area_activa"])
    elif st.session_state.get("mostrar_dialogo_expiracion"):
        modal_alerta_cartera()



def _cambiar_selector_cartera():
    """Registra el cambio deliberado, incluso cuando el encargado quita el vínculo."""
    st.session_state["cartera_seleccion_confirmada_manualmente"] = True
    st.session_state["cartera_examen_emparejado"] = (
        st.session_state.get("sel_cartera_main") or "Sin vincular"
    )


def _vincular_candidato_cartera(nombre_examen):
    """Vincula una alternativa sin deshacerla en el siguiente rerun automático."""
    if nombre_examen:
        st.session_state["sel_cartera_main"] = nombre_examen
        st.session_state["cartera_seleccion_confirmada_manualmente"] = True
        st.session_state["cartera_examen_emparejado"] = nombre_examen


def _cambiar_selector_acreditacion():
    """Marca los cambios explícitos: escoger una prestación o dejarla vacía."""
    st.session_state["acreditacion_seleccion_confirmada_manualmente"] = True
    st.session_state["acreditacion_manual_id"] = st.session_state.get(
        "acreditacion_manual_selector")
    st.session_state["acreditacion_coincidencia_automatica"] = None


def _seleccionar_candidato_acreditacion(identificador):
    """Confirma una sugerencia mediante callback previo al selectbox existente."""
    st.session_state["_acreditacion_reset_pendiente"] = False
    st.session_state["acreditacion_manual_selector"] = identificador
    st.session_state["acreditacion_manual_id"] = identificador
    st.session_state["acreditacion_seleccion_confirmada_manualmente"] = True
    st.session_state["acreditacion_coincidencia_automatica"] = None


def _fuente_pdf_apa_cacheada(pdf, examen, texto_reconocido=""):
    """Evita volver a abrir un PDF solo para su ficha bibliográfica en cada rerun.

    La clave incluye huella del archivo, nombre y examen. Un archivo actualizado
    incluso con igual nombre se vuelve a analizar; el contenido de extracción
    del PDF no se guarda ni se modifica aquí. Hasta 12 entradas por sesión.
    """
    contenido = pdf.getvalue()
    # Versión del analizador: evita reutilizar referencias antiguas después
    # de introducir el reconocimiento del título y los datos editoriales Roche.
    clave = ("bibliografia_v4", pdf.name, hashlib.sha256(contenido).hexdigest(),
             examen or "", hashlib.sha256(texto_reconocido.encode("utf-8")).hexdigest())
    cache = st.session_state.setdefault("cache_apa_pdf", {})
    if clave not in cache:
        if len(cache) >= 12:
            cache.pop(next(iter(cache)))
        if texto_reconocido:
            cache[clave] = apa.fuente_pdf(pdf.name, contenido, examen,
                                              texto_reconocido=texto_reconocido)
        else:
            cache[clave] = apa.fuente_pdf(pdf.name, contenido, examen)
    return cache[clave]


def _texto_optico_para_referencia(texto_global, nombre_pdf):
    """En análisis conjunto, no atribuir a un PDF la fecha de otro PDF."""
    texto_global = str(texto_global or "")
    cabecera = re.compile(r"(?m)^===== (.+?) =====\n")
    bloques = list(cabecera.finditer(texto_global))
    if not bloques:
        return texto_global
    for indice, marca in enumerate(bloques):
        if marca.group(1) == nombre_pdf:
            fin = bloques[indice + 1].start() if indice + 1 < len(bloques) else len(texto_global)
            return texto_global[marca.end():fin].strip()
    return ""  # Nunca tomar la fecha de otro documento.


def _fuentes_para_referencias(pdfs_activos, indice_actual, examen, datos, origen_campos, usó_cartera):
    """Referencias de documentos realmente utilizados, nunca solo por estar en SQLite."""
    pdf_seleccionados = ([pdfs_activos[indice_actual]] if pdfs_activos else [])
    # Excluir documentos sin campos extraídos, aunque estén cargados en la cola.
    nombres_con_aporte = set(st.session_state.get("pdf_documentos_con_aporte", []))
    pdf_fuentes = [_fuente_pdf_apa_cacheada(
        pdf, examen, _texto_optico_para_referencia(
            st.session_state.get("extracted_text", ""), pdf.name)
        if st.session_state.get("procedencia_modalidad") == "Documentos escaneados"
        else "")
        for pdf in pdf_seleccionados if pdf.name in nombres_con_aporte]
    todas = list(pdf_fuentes)
    id_origen = {
        "Normas Generales de Toma de Muestras": "normas_generales",
        "Cartera de Exámenes del CRS · Indicaciones": "cartera_examenes",
        htest.ETIQUETA: htest.FUENTE_ID,
    }
    ids = set()
    for campo, origen in origen_campos.items():
        if not str(datos.get(campo, "") or "").strip():
            continue
        # Citar un documento solo cuando el valor vigente realmente proviene
        # de él, y no porque también estaba disponible durante el análisis.
        valor_institucional = st.session_state.get(
            "procedencia_valores_documentales", {}).get(campo)
        if str(valor_institucional or "").strip() != str(datos.get(campo, "") or "").strip():
            continue
        if (st.session_state.get("procedencia_modalidad") == "Documentos escaneados"
                and campo not in st.session_state.get("fuentes_elegidas_usuario", set())
                and str(st.session_state.get("procedencia_pdf_campos", {}).get(campo, "") or "").strip()
                    == str(datos.get(campo, "") or "").strip()):
            continue
        if origen in id_origen:
            ids.add(id_origen[origen])
        elif origen.startswith("Cartera de Acreditación"):
            ids.add("cartera_acreditacion")
    if usó_cartera:
        nombre_cartera = st.session_state.get("cartera_examen_emparejado")
        info_cartera_real = obtener_info_cartera_por_nombre(nombre_cartera)
        if info_cartera_real and any(
            str(datos.get(k, "") or "").strip() == str(info_cartera_real.get(k, "") or "").strip()
            for k in ("Código", "Preparación del paciente", "Equipamiento utilizado", "Plazo entrega resultados")
            if str(info_cartera_real.get(k, "") or "").strip()
        ):
            ids.add("cartera_examenes")
    cache_fuentes = st.session_state.setdefault("cache_apa_fuentes", {})
    for fuente_id in sorted(ids):
        metadatos = fd.documento_fuente(DB_PATH, fuente_id, incluir_archivo=False)
        if not metadatos and fuente_id == "cartera_examenes":
            # Bases antiguas no siempre incluyen el Excel original: citar su
            # registro institucional de forma provisional, sin inventar año.
            with sqlite3.connect(DB_PATH) as con:
                fila_nombre = con.execute(
                    "SELECT valor FROM cartera_metadata WHERE clave='nombre_archivo'"
                ).fetchone()
            nombre = fila_nombre[0] if fila_nombre and fila_nombre[0] else "Cartera de exámenes del CRS"
            metadatos = {"sha256": "sin-documento:" + nombre, "nombre_archivo": nombre}
        if not metadatos:
            continue
        cache = cache_fuentes.get(fuente_id)
        if cache and cache[0] == metadatos["sha256"]:
            fuente = cache[1]
        else:
            # Solo ante cambio del documento: leer el original en PDF para
            # encontrar fecha bibliográfica; jamás usar nombre/fecha de carga.
            documento_original = fd.documento_fuente(DB_PATH, fuente_id, incluir_archivo=True)
            fuente = apa.fuente_institucional(fuente_id, documento_original or metadatos)
            cache_fuentes[fuente_id] = (metadatos["sha256"], fuente)
        todas.append(fuente)
    return pdf_fuentes, todas


def _limpiar_examen_para_siguiente_pdf():
    """Limpia solo el examen actual, sin perder las fichas ya consolidadas."""
    st.session_state["examen_activo"] = None
    st.session_state["acreditacion_manual_id"] = None
    st.session_state["_acreditacion_reset_pendiente"] = True
    st.session_state["cartera_examen_emparejado"] = "Sin vincular"
    st.session_state["_sel_cartera_pendiente"] = None
    for clave in (
        "candidatos_cartera_pendientes", "data_extracted", "user_edited_fields",
        "extracted_text", "contexto_emparejamiento_cartera",
        "omitir_emparejamiento_examen", "modo_referencia_aplicado",
        "seleccion_campos_excel", "excel_para_actualizar",
        "historial_edicion_tabla", "campos_restaurados_tabla", "marca_historial_consumida",
    ):
        st.session_state.pop(clave, None)
    st.session_state["aggrid_key_counter"] = st.session_state.get("aggrid_key_counter", 0) + 1
    st.session_state["mostrar_modal_examenes"] = False
    st.session_state.pop("modal_area_activa", None)


def _omitir_pdf_actual(indice, total, nombre):
    """Omitir no extrae, no altera result_bytes y no guarda en SQLite."""
    cf.omitir_documento(st.session_state, indice, total, nombre)
    _limpiar_examen_para_siguiente_pdf()


def _aplicar_historial_tabla(accion, actualizar_grid=True):
    """Restaura valores en la sesión y los fuerza también en el Excel.

    El botón reconstruye la tabla (un rerun habitual de Streamlit). Ctrl+Z
    restaura visualmente las celdas en AG Grid y solo sincroniza el backend:
    no incrementa la clave ni fuerza un segundo st.rerun completo.
    """
    historial = st.session_state.get("historial_edicion_tabla", {})
    datos = st.session_state.get("data_extracted")
    if datos is None:
        return False
    editados = st.session_state.setdefault("user_edited_fields", set())
    campos = ht.aplicar_accion(historial, datos, editados, accion)
    if not campos:
        return False
    st.session_state.setdefault("campos_restaurados_tabla", set()).update(campos)
    if actualizar_grid:
        st.session_state["aggrid_key_counter"] = st.session_state.get("aggrid_key_counter", 0) + 1
    return True


# --- SIDEBAR ---
with st.sidebar:
    logo_path = os.path.join(os.path.dirname(__file__), "CRSCO.png")
    if os.path.exists(logo_path):
        st.image(logo_path, use_container_width=True)
    st.caption(f"Sesión: **{st.session_state.get('nombre_usuario', 'Sin identificar')}** · {st.session_state.get('usuario_acceso','')}")
    with st.expander("Acceso", expanded=False):
        rol_actual = gp.rol_usuario(DB_PATH, st.session_state.get("usuario_acceso"))
        st.caption(f"Usuario: **{st.session_state.get('usuario_acceso', '')}**")
        st.caption(f"Rol: **{'Administrador' if rol_actual == 'administrador' else 'Encargado'}**")
        if rol_actual == "administrador":
            if st.button("Administrar usuarios", use_container_width=True,
                         key="btn_administracion_usuarios"):
                st.session_state["revision_admin_usuarios"] = gp.estado_recurso(
                    DB_PATH, "usuarios:administracion")["revision"]
                st.session_state["modal_admin_usuarios_abierto"] = True
                st.rerun()
        else:
            st.caption("La administración de cuentas está disponible solo para administradores.")
    st.divider()
    st.header("Áreas del laboratorio", anchor=False)
    for area in AREA_CONFIG:
        st.button(
            area,
            key=f"abrir_area_{area}",
            use_container_width=True,
            type="primary" if st.session_state.get("area_activa") == area else "secondary",
            on_click=seleccionar_area,
            args=(area,),
        )
    st.divider()
    st.markdown(f"**Área activa:** {st.session_state.get('area_activa') or 'Sin seleccionar'}")
    examen_en_uso = st.session_state.get("examen_activo")
    if examen_en_uso:
        st.markdown(
            '<div class="sidebar-examen-activo">'
            '<strong>Examen activo:</strong> '
            f'<span>{html.escape(str(examen_en_uso))}</span></div>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            '<div class="sidebar-examen-sin-seleccionar">'
            '<strong>Examen activo:</strong> <span>Sin seleccionar</span></div>',
            unsafe_allow_html=True,
        )
    st.divider()

    # --- SECCIÓN BASE DE DATOS (SQLITE) SIN EMOJIS ---
    st.header("Base de datos", anchor=False)

    # Áreas del laboratorio almacenadas en SQLite (mantener el orden de AREA_CONFIG)
    for area_item in AREA_CONFIG:
        b_bytes_area, fecha_act_area = obtener_plantilla_area_db(area_item)
        with st.expander(f"Área: {area_item}", expanded=False):
            st.caption(f"Última actualización: **{fecha_act_area}**")
            st.caption(f"Modificado por: **{gp.estado_recurso(DB_PATH, 'area:' + area_item)['usuario']}**")
            nombre_plantilla = obtener_nombre_plantilla_area_db(area_item)
            if nombre_plantilla:
                st.caption(f"Documento almacenado: **{nombre_plantilla}**")
            elif b_bytes_area:
                st.caption("Documento de origen: no registrado en esta versión anterior de la base de datos.")

            if b_bytes_area:
                st.download_button(
                    f"Descargar ficha técnica ({area_item})",
                    data=b_bytes_area,
                    file_name=nombre_descarga_excel_generado(
                        nombre_plantilla, area=area_item, fecha_registro=fecha_act_area),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key=f"dl_area_btn_{area_item}", on_click="ignore"
                )
            else:
                st.info("No hay Ficha Técnica/Plantilla registrada en la base de datos para esta área.")

            area_file = st.file_uploader(f"Cargar nueva plantilla de {area_item} (.xlsx)", type=["xlsx"], key=f"uploader_area_{area_item}")
            if area_file is not None:
                if st.button(f"Guardar {area_item} en Base de Datos", type="primary", use_container_width=True, key=f"btn_save_area_{area_item}"):
                    solicitar_cambio("area", "area:" + area_item,
                                    archivo=area_file.getvalue(), nombre_archivo=area_file.name,
                                    descripcion="Reemplazar la ficha técnica almacenada del área",
                                    area=area_item)

    # Fuentes documentales al final del menú de la base de datos.
    # Cada origen conserva su documento, fecha, descarga y actualización en SQLite.
    with st.expander("Cartera de exámenes y acreditación", expanded=cartera_expirada):
        tab_acreditacion, tab_examenes = st.tabs(["Acreditación", "Exámenes"])

        with tab_acreditacion:
            meta_acr = fd.documento_fuente(DB_PATH, fd.ACREDITACION_ID, incluir_archivo=True)
            if meta_acr:
                st.caption(f"Última actualización: **{formatear_fecha_fuente(meta_acr['fecha_actualizacion'])}**")
                st.caption(f"Modificado por: **{gp.estado_recurso(DB_PATH, 'fuente:cartera_acreditacion')['usuario']}**")
                st.caption(f"Documento almacenado: **{meta_acr['nombre_archivo']}** · {meta_acr['registros']} prestaciones")
                st.download_button(
                    "Descargar cartera de acreditación original",
                    data=meta_acr["archivo"],file_name=meta_acr["nombre_archivo"],
                    mime=meta_acr["tipo_mime"],use_container_width=True,
                    key="descargar_acreditacion_original", on_click="ignore")
            else:
                st.warning("No está cargada la Cartera de Acreditación.")
            nuevo_acr = st.file_uploader(
                "Cargar nueva Cartera de Acreditación (.xlsx)",type=["xlsx"],
                key="actualizar_acreditacion_archivo")
            if nuevo_acr is not None and st.button(
                    "Guardar nueva Cartera de Acreditación",key="guardar_acreditacion_nueva",
                    type="primary",use_container_width=True):
                solicitar_cambio("acreditacion", "fuente:cartera_acreditacion",
                                archivo=nuevo_acr.getvalue(), nombre_archivo=nuevo_acr.name,
                                descripcion="Reemplazar la Cartera de acreditación y su índice de prestaciones")


        with tab_examenes:
            st.caption(f"Última actualización: **{msg_cartera}**")
            st.caption(f"Modificado por: **{gp.estado_recurso(DB_PATH, 'fuente:cartera_examenes')['usuario']}**")
            meta_cartera_original = fd.documento_fuente(DB_PATH, "cartera_examenes", incluir_archivo=True)
            if meta_cartera_original:
                st.caption(f"Documento almacenado: **{meta_cartera_original['nombre_archivo']}**")
            else:
                st.caption("Documento de origen: no registrado en esta versión anterior de la base de datos.")

            b_bytes_cartera = (meta_cartera_original["archivo"] if meta_cartera_original
                               else obtener_cartera_excel_bytes())
            if b_bytes_cartera:
                ZONA_HORARIA = ZoneInfo("America/Santiago")
                fecha_cartera_dl = fecha_hoy_nombre_archivo()
                st.download_button(
                    "Descargar cartera de exámenes",
                    data=b_bytes_cartera,
                    file_name=(meta_cartera_original['nombre_archivo'] if meta_cartera_original
                               else f"Cartera_de_Exámenes_Actualizada_{fecha_cartera_dl}.xlsx"),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="dl_cartera_btn", on_click="ignore"
                )
            else:
                st.info("No hay exámenes registrados en la Cartera de la base de datos.")

            cartera_file = st.file_uploader("Cargar nueva Cartera (.xlsx)", type=["xlsx"], key="cartera_uploader_sidebar")
            if cartera_file is not None:
                if st.button("Guardar en Base de Datos", type="primary", use_container_width=True, key="btn_save_cartera"):
                    solicitar_cambio("cartera", "fuente:cartera_examenes",
                                    archivo=cartera_file.getvalue(), nombre_archivo=cartera_file.name,
                                    descripcion="Reemplazar la Cartera de exámenes y su documento original")


    # Homologación: cinco documentos separados, uno por área. Los archivos y
    # sus índices se reemplazan de forma independiente y sin cruces de exámenes.
    with st.expander("Homologación de test", expanded=False):
        st.caption("Cada área conserva su propio archivo. La herramienta consulta "
                   "únicamente la homologación del área del examen seleccionado.")
        pestanas_ht = st.tabs(list(htest.AREAS))
        for area_ht, pestana_ht in zip(htest.AREAS, pestanas_ht):
            with pestana_ht:
                clave_ht = htest.id_fuente_area(area_ht).split(":", 1)[1]
                recurso_ht = htest.recurso_area(area_ht)
                meta_ht = fd.documento_fuente(DB_PATH, htest.id_fuente_area(area_ht), incluir_archivo=True)
                if st.session_state.get("area_activa") == area_ht:
                    st.caption("Área de trabajo actual: **" + area_ht + "**")
                if meta_ht:
                    st.caption(f"Última actualización: **{formatear_fecha_fuente(meta_ht['fecha_actualizacion'])}**")
                    usuario_ht = gp.estado_recurso(DB_PATH, recurso_ht)['usuario']
                    if usuario_ht:
                        st.caption(f"Modificado por: **{usuario_ht}**")
                    st.caption(f"Documento almacenado: **{meta_ht['nombre_archivo']}** · {meta_ht['registros']} registros")
                    st.download_button(
                        "Descargar homologación original de " + area_ht,
                        data=meta_ht['archivo'], file_name=meta_ht['nombre_archivo'],
                        mime=meta_ht['tipo_mime'], use_container_width=True,
                        key=f"descargar_homologacion_{clave_ht}", on_click="ignore")
                else:
                    st.info("Todavía no se ha cargado un archivo de homologación para " + area_ht + ".")
                    # El Excel de Hematología se distribuye sin base de datos.
                    # La incorporación se confirma mediante el mismo circuito
                    # auditado que cualquier reemplazo institucional.
                    if area_ht == "Hematología y coagulación":
                        incluido_ht = htest.archivo_inicial_hematologia()
                        if incluido_ht:
                            try:
                                registros_incluidos, meta_incluida = htest.leer_excel(incluido_ht)
                                incompatibles_incluidas = [fila[6] for fila in registros_incluidos
                                    if fila[6] and not htest._area_celda_compatible(fila[6], area_ht)]
                                if incompatibles_incluidas:
                                    st.warning("El documento incluido necesita revisión de sus grupos antes de importarlo.")
                                else:
                                    st.caption("Documento de Hematología adjuntado al proyecto: "
                                               f"**{meta_incluida['registros']} exámenes**. "
                                               "No se instalará sin confirmación.")
                                    if st.button("Incorporar Excel de Hematología incluido",
                                                 use_container_width=True,
                                                 key="incorporar_homologacion_incluida"):
                                        solicitar_cambio(
                                            "homologacion_tests", recurso_ht,
                                            archivo=incluido_ht,
                                            nombre_archivo=htest.ARCHIVO_INICIAL_HEMATOLOGIA.name,
                                            descripcion="Incorporar homologación de test de Hematología desde el archivo incluido",
                                            area_fuente=area_ht,
                                            hoja=meta_incluida["hoja"],
                                            columna_examen=meta_incluida["columna_examen"],
                                            columna_codigo=meta_incluida["columna_codigo"])
                            except ValueError as exc_ht:
                                st.warning("No se pudo verificar el documento inicial de Hematología: " + str(exc_ht))
                nuevo_ht = st.file_uploader(
                    "Cargar o reemplazar homologación de " + area_ht + " (.xlsx)", type=["xlsx"],
                    key=f"archivo_homologacion_{clave_ht}")
                if nuevo_ht is not None:
                    try:
                        tablas_ht = htest.inspeccionar_encabezados(nuevo_ht.getvalue())
                        if not tablas_ht:
                            st.error('No se encontró la columna «¿QUÉ EXAMEN O RESULTADOS INCLUYE?».')
                        else:
                            titulos_tabla = [f"{hoja} (encabezado fila {fila})"
                                             for hoja, fila, _ in tablas_ht]
                            indice_tabla = st.selectbox(
                                "Tabla que contiene los resultados", list(range(len(tablas_ht))),
                                format_func=lambda i: titulos_tabla[i],
                                key=f"tabla_homologacion_{clave_ht}")
                            hoja_ht, _, columnas_ht = tablas_ht[indice_tabla]
                            opciones_ht = [f"{i + 1}. {nombre or '(sin nombre)'}"
                                           for i, nombre in enumerate(columnas_ht)]
                            col_examen_ht = st.selectbox(
                                "Columna que identifica el examen (opcional si se utiliza código)",
                                [None] + list(range(len(columnas_ht))),
                                format_func=lambda i: "Detectar automáticamente" if i is None else opciones_ht[i],
                                key=f"examen_homologacion_{clave_ht}")
                            col_codigo_ht = st.selectbox(
                                "Código comparable con la Cartera del CRS (no seleccionar Código FONASA)",
                                [None] + list(range(len(columnas_ht))),
                                format_func=lambda i: "Detectar automáticamente" if i is None else opciones_ht[i],
                                key=f"codigo_homologacion_{clave_ht}")
                            try:
                                filas_ht, detalle_ht = htest.leer_excel(
                                    nuevo_ht.getvalue(), hoja=hoja_ht,
                                    columna_examen=col_examen_ht, columna_codigo=col_codigo_ht)
                                ajenas_ht = [fila[6] for fila in filas_ht if fila[6]
                                            and not htest._area_celda_compatible(fila[6], area_ht)]
                                if ajenas_ht:
                                    st.error("El documento contiene una identificación de área diferente "
                                             f"({ajenas_ht[0]}). Selecciona la pestaña correcta antes de guardarlo.")
                                else:
                                    st.caption(f"Se reconocerán **{detalle_ht['registros']} filas** para {area_ht}.")
                                    if st.button("Guardar homologación de " + area_ht, type="primary",
                                                 use_container_width=True, key=f"guardar_homologacion_{clave_ht}"):
                                        solicitar_cambio(
                                            "homologacion_tests", recurso_ht,
                                            archivo=nuevo_ht.getvalue(), nombre_archivo=nuevo_ht.name,
                                            descripcion="Reemplazar homologación de test de " + area_ht,
                                            area_fuente=area_ht, hoja=hoja_ht,
                                            columna_examen=col_examen_ht,
                                            columna_codigo=col_codigo_ht)
                            except ValueError as error_excel:
                                st.warning(str(error_excel))
                                st.caption("Selecciona explícitamente la columna de examen o código correcta.")
                    except ValueError as error_hojas:
                        st.error(str(error_hojas))

    # Normas generales permanece independiente y disponible para reemplazo.
    with st.expander("Normas generales de toma de muestras", expanded=False):
        meta_normas = fd.documento_fuente(DB_PATH, fd.NORMAS_ID, incluir_archivo=True)
        if meta_normas:
            st.caption(f"Última actualización: **{formatear_fecha_fuente(meta_normas['fecha_actualizacion'])}**")
            st.caption(f"Modificado por: **{gp.estado_recurso(DB_PATH, 'fuente:normas_generales')['usuario']}**")
            st.caption(f"Documento almacenado: **{meta_normas['nombre_archivo']}** · 4 campos predeterminados")
            st.download_button(
                "Descargar documento original de normas",
                data=meta_normas["archivo"],
                file_name=meta_normas["nombre_archivo"],
                mime=meta_normas["tipo_mime"],
                use_container_width=True, key="descargar_normas_original", on_click="ignore")
        else:
            st.warning("No están cargadas las normas generales.")
        nuevo_normas = st.file_uploader(
            "Cargar nuevas Normas Generales (.docx o .pdf)", type=["docx", "pdf"],
            key="actualizar_normas_archivo")
        if nuevo_normas is not None and st.button(
                "Guardar nuevas normas en la base de datos", key="guardar_normas_nuevas",
                type="primary", use_container_width=True):
            solicitar_cambio("normas", "fuente:normas_generales",
                            archivo=nuevo_normas.getvalue(), nombre_archivo=nuevo_normas.name,
                            descripcion="Reemplazar las Normas generales y sus cuatro campos")

    st.divider()
    st.header("Trazabilidad", anchor=False)
    with st.expander("Historial de modificaciones", expanded=False):
        tabla = _tabla_trazabilidad("modificaciones", 5)
        if not tabla.empty:
            with st.container(key="traza_previa_modificaciones"):
                st.dataframe(tabla, hide_index=True, use_container_width=True)
            if st.button("⛶ Fullscreen", key="traza_full_modificaciones", use_container_width=True):
                st.session_state["trazabilidad_pantalla_completa"] = "modificaciones"
                st.session_state["traza_pagina_modificaciones"] = 1
                st.rerun()
        else:
            st.info("Aún no hay modificaciones confirmadas registradas.")

    with st.expander("Actividad de usuarios", expanded=False):
        tabla = _tabla_trazabilidad("actividad", 5)
        if not tabla.empty:
            with st.container(key="traza_previa_actividad"):
                st.dataframe(tabla, hide_index=True, use_container_width=True)
            if st.button("⛶ Fullscreen", key="traza_full_actividad", use_container_width=True):
                st.session_state["trazabilidad_pantalla_completa"] = "actividad"
                st.session_state["traza_pagina_actividad"] = 1
                st.rerun()
        else:
            st.info("No hay actividad registrada.")

    st.divider()
    if st.button("Cerrar sesión"):
        try:
            gp.registrar_acceso(DB_PATH, st.session_state.get("usuario_acceso", "Desconocido"), "Cierre de sesión")
        except sqlite3.Error:
            pass
        st.session_state.pop("usuario_acceso", None)
        st.session_state.pop("nombre_usuario", None)
        st.session_state.pop("revision_sesion", None)
        st.session_state["authenticated"] = False
        st.session_state.pop("result_bytes", None)
        st.session_state.pop("modal_area_activa", None)
        st.session_state["mostrar_modal_examenes"] = False
        st.session_state["area_activa"] = None
        st.session_state["cola_pdf_indice"] = 0
        st.rerun()


# --- VISTA PRINCIPAL ---
if st.session_state.get("area_activa") is None:
    st.write("")
    st.info("### ¡Bienvenido/a a la Herramienta de Extracción de Datos!")
    
    st.markdown(
        """
        Para comenzar a trabajar, selecciona un **área del laboratorio** en el panel lateral de la izquierda.

        #### Pasos para trabajar:
        1. **Selecciona un Área del laboratorio** en el menú de la izquierda (*Microbiología, Hematología y coagulación, Bioquímica, Hormonas e inmunidad, Derivados*).
        2. **Administra y selecciona el examen** que deseas procesar usando el botón **Administrar exámenes**.
        3. **Carga las fichas técnicas en PDF** (lote de hasta 5 documentos simultáneos) y organiza la cola de procesamiento si lo deseas.
        4. **Define los parámetros operacionales** (*Laboratorio de procesamiento, Actualizado por, Autorizado por*) y verifica el emparejamiento con la Cartera de Exámenes.
        5. **Haz clic en "Analizar documento"** para realizar la extracción automática de texto, lectura óptica y cruce con la Base Maestra de Cartera.
        6. **Revisa y edita la información** en la tabla interactiva, resuelve discrepancias si existen, y actualiza la planilla del área o descarga el Excel consolidado.
        """
    )
    
    st.divider()
    
    # --- MÓDULO DE ESTADO DE LA CARTERA SQLITE EN LA PÁGINA DE BIENVENIDA ---
    st.markdown("#### Estado de la Cartera de Exámenes:")
    lista_cartera_home = obtener_lista_cartera_db()
    total_examenes_cartera = len(lista_cartera_home)
    fecha_cartera_home, cartera_expirada_home, msg_cartera_home = fecha_actualizacion_cartera()

    if fecha_cartera_home is None or total_examenes_cartera == 0:
        st.warning("Cartera desactualizada: No se ha subido ninguna Cartera de Exámenes a la base de datos. Se sugiere cargar el archivo Excel en el menú lateral.")
    elif cartera_expirada_home:
        st.warning(f"Cartera desactualizada (más de 6 meses de antigüedad): Registrada el **{msg_cartera_home}**. Se recomienda actualizar la base de datos subiendo el archivo Excel desde el panel lateral.")
    else:
        st.success(f"Cartera de Exámenes cargada y lista: Última actualización registrada el **{msg_cartera_home}**.")

    col_stat1, col_stat2 = st.columns(2)
    with col_stat1:
        st.metric(
            label="Exámenes registrados en la Cartera Base",
            value=f"{total_examenes_cartera} exámenes"
        )
    with col_stat2:
        st.metric(
            label="Fecha de última actualización",
            value=msg_cartera_home if fecha_cartera_home else "Sin registro"
        )

else:
    with st.container(key="pestanas_principales"):
        tab_pdf, tab_export, tab_guide = st.tabs([
            "Extraer información desde un PDF",
            "Exportar a otro formato",
            "Guía básica para usuario",
        ])

    with tab_pdf:
        st.subheader("Extraer información desde fichas técnicas (PDF)")
        area_activa = st.session_state.get("area_activa")
        examen_activo = st.session_state.get("examen_activo")

        # CABECERA CON BOTÓN PARA ABRIR LA VENTANA DE ADMINISTRAR EXÁMENES
        col_info_area, col_btn_admin = st.columns([3, 1])
        with col_info_area:
            st.info(f"Área seleccionada: **{area_activa}** · Examen activo: **{examen_activo or 'sin seleccionar'}**")
        with col_btn_admin:
            if st.button("Administrar exámenes", type="primary", use_container_width=True, key="btn_abrir_admin_examenes"):
                st.session_state["modal_area_activa"] = area_activa
                st.session_state["mostrar_modal_examenes"] = True
                st.rerun()

        # Un solo selector: fichas técnicas conserva la extracción habitual;
        # documentos escaneados habilita OCR optimizado y secciones numeradas.
        modalidades_validas = ("Fichas técnicas", "Documentos escaneados")
        if st.session_state.get("modo_procesamiento_pdf") not in (*modalidades_validas, None):
            st.session_state.pop("modo_procesamiento_pdf", None)
        # El selector secundario de versiones anteriores se retira de sesión.
        st.session_state.pop("tipo_documento_extraccion", None)
        modo_procesamiento = st.radio(
            "Modalidad de análisis", modalidades_validas, horizontal=True,
            key="modo_procesamiento_pdf",
            help="Fichas técnicas: análisis habitual. Documentos escaneados: lectura de PDF escaneados y manuales."
        )
        es_manual_procedimiento = modo_procesamiento == "Documentos escaneados"
        # La modalidad de varios PDF por una misma ficha ya no se ofrece.
        # Se mantiene la cola individual; los archivos se procesan uno por uno.
        modo_conjunto = False
        if st.session_state.get("modo_pdf_anterior") != modo_procesamiento:
            st.session_state["modo_pdf_anterior"] = modo_procesamiento
            st.session_state["cola_pdf_indice"] = 0
            for clave in ("cola_pdf_omitidos", "cola_pdf_procesados", "cola_pdf_finalizada_omision",
                          "historial_edicion_tabla", "campos_restaurados_tabla"):
                st.session_state.pop(clave, None)
            for clave in ("data_extracted", "result_bytes", "extracted_text", "user_edited_fields", "discrepancias_pdf", "fuentes_pdf", "modal_pdf_pdf_abierto", "modal_discrepancias_unificado", "manual_ocr_diagnostico", "procedencia_pdf_campos", "procedencia_cartera_campos", "procedencia_valores_documentales", "manual_pdf_digital_detectado"):
                st.session_state.pop(clave, None)

        # Una ficha institucional puede generarse SIN PDF. El botón no omite
        # una posición en la cola: cambia el origen de ESTA ficha y no ejecuta
        # extracción digital ni lectura óptica.
        modo_solo_bd = bool(st.session_state.get("modo_solo_base_datos", False))
        if not modo_solo_bd:
            if st.button(
                "Omitir análisis y trabajar solo con la base de datos",
                type="secondary", key="btn_omitir_analisis_completo",
                help="No necesitas adjuntar un PDF: se utilizarán únicamente las fuentes institucionales vinculadas.",
            ):
                st.session_state["modo_solo_base_datos"] = True
                for clave in ("data_extracted", "result_bytes", "extracted_text", "user_edited_fields",
                              "procedencia_pdf_campos", "procedencia_cartera_campos", "procedencia_valores_documentales",
                              "conflictos_pendientes", "discrepancias_pdf", "fuentes_pdf", "pdf_documentos_con_aporte",
                              "seleccion_campos_excel", "historial_edicion_tabla", "campos_restaurados_tabla",
                              "bytes_plantilla_inicial", "manual_ocr_diagnostico", "modo_referencia_aplicado"):
                    st.session_state.pop(clave, None)
                st.rerun()
        else:
            st.info("**Solo base de datos:** no se realizará análisis de documentos. "
                    "Selecciona las Carteras y revisa los campos propuestos antes de guardar.")
            if st.button("Volver al análisis de documentos", type="secondary",
                         key="btn_volver_a_documentos"):
                st.session_state.pop("modo_solo_base_datos", None)
                for clave in ("data_extracted", "result_bytes", "extracted_text", "user_edited_fields",
                              "procedencia_pdf_campos", "procedencia_cartera_campos", "procedencia_valores_documentales",
                              "conflictos_pendientes", "discrepancias_pdf", "fuentes_pdf", "pdf_documentos_con_aporte",
                              "seleccion_campos_excel", "historial_edicion_tabla", "campos_restaurados_tabla",
                              "bytes_plantilla_inicial", "manual_ocr_diagnostico", "modo_referencia_aplicado"):
                    st.session_state.pop(clave, None)
                st.rerun()

        # CARGA DE ARCHIVOS PDF Y CONTROL DE COLA
        if modo_solo_bd:
            pdfs = []
        else:
            pdfs = st.file_uploader(
                (f"Cargar documentos escaneados (hasta {MAX_PDFS_PERMITIDOS} PDF)"
                 if es_manual_procedimiento else
                 f"Cargar fichas técnicas (hasta {MAX_PDFS_PERMITIDOS} PDF para procesar en cola)"),
                type=["pdf"],
                accept_multiple_files=True,
                help=f"Puedes cargar hasta {MAX_PDFS_PERMITIDOS} archivos PDF para procesarlos de forma individual uno tras otro.",
                key=f"pdf_uploader_{st.session_state.get('generacion_cargas_area', 0)}"
            )

        exceso_archivos = pdfs is not None and len(pdfs) > MAX_PDFS_PERMITIDOS

        if exceso_archivos:
            st.error(f"Has seleccionado **{len(pdfs)} archivos**. El límite máximo simultáneo es de **{MAX_PDFS_PERMITIDOS} documentos PDF** para evitar saturar el sistema.")

        if (pdfs and not exceso_archivos
                and st.session_state.get("cola_pdf_finalizada_omision")
                and len(pdfs) > 1
                and set(st.session_state.get("orden_nombres_pdf", [])) == {p.name for p in pdfs}):
            # No volver al último PDF: fue omitido expresamente. La descarga y
            # la actualización deben seguir disponibles si hubo extracciones.
            total_pdfs_cola = len(pdfs)
            resumen = cf.resumen_cola(st.session_state, total_pdfs_cola)
            st.success(
                f"¡Has completado la cola! {len(resumen['procesados'])} documento(s) "
                f"procesado(s) y {len(resumen['omitidos'])} omitido(s)."
            )
            if resumen["omitidos"]:
                st.caption("Omitidos sin cambios: " + ", ".join(resumen["omitidos"]))
            if resumen["puede_guardar"]:
                # Centrar únicamente este encabezado; conservar el estilo nativo
                # del subheader y no afectar otros títulos de la aplicación.
                st.html("<style>.st-key-titulo_guardado_omision [data-testid='stHeadingWithActionElements'] { justify-content: center; width: 100%; } .st-key-titulo_guardado_omision h3 { text-align: center; width: 100%; }</style>")
                with st.container(key="titulo_guardado_omision"):
                    st.subheader("Guardar avances de la cola", anchor=False)
                # Centrar el aviso sobre la franja que ocupan los dos botones,
                # conservando el estilo azul original de st.info.
                st.html("<style>.st-key-aviso_guardado_omision [data-testid='stAlert'] p { text-align: center; }</style>")
                _, col_aviso_guardado, _ = st.columns([1, 4, 1], gap="small")
                with col_aviso_guardado:
                    with st.container(key="aviso_guardado_omision"):
                        st.info("Los exámenes procesados siguen en la ficha consolidada. Omitir un PDF no altera ni borra los avances anteriores.")
                try:
                    excel_final = validar_excel_para_descarga(st.session_state["result_bytes"])
                except ValueError as error_excel:
                    excel_final = None
                    st.error(f"No se pudo validar el Excel consolidado: {error_excel}")
                # En el guardado habitual cada botón ocupa una de tres columnas.
                # Reservamos medio ancho de columna en cada extremo para
                # centrar los dos botones de omisión sin ensancharlos.
                _, col_final_descarga, col_final_actualiza, _ = st.columns(
                    [1, 2, 2, 1], gap="small"
                )
                with col_final_descarga:
                    if excel_final is not None:
                        st.download_button(
                            "Descargar Excel consolidado (.xlsx)",
                            data=excel_final,
                            file_name=nombre_ficha_tecnica(
                                area_activa, st.session_state.get("fecha_ultima_extraccion_ficha")),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="secondary", use_container_width=True,
                            key="descarga_cola_tras_omision", on_click="ignore",
                        )
                with col_final_actualiza:
                    hay_discrepancias = bool(
                        st.session_state.get("discrepancias_pdf")
                        or st.session_state.get("conflictos_pendientes")
                    )
                    if st.button(
                        "Actualizar en base de datos", type="primary", use_container_width=True,
                        key="guardar_cola_tras_omision",
                        disabled=(excel_final is None or hay_discrepancias),
                        help=("Deshabilitado por discrepancias sin analizar."
                              if hay_discrepancias else
                              "No hay un Excel válido para actualizar."
                              if excel_final is None else None),
                    ):
                        anterior, _ = obtener_plantilla_area_db(area_activa)
                        st.session_state["excel_para_actualizar"] = excel_final
                        st.session_state["comparacion_planilla_detalles"] = cf.comparar_libros_completos(
                            anterior, excel_final)
                        revision = gp.estado_recurso(DB_PATH, "area:" + area_activa)
                        st.session_state["revision_plantilla_pendiente"] = revision["revision"]
                        st.session_state["metadata_plantilla_pendiente"] = revision
                        st.session_state["modal_confirmar_planilla_abierto"] = True
            else:
                st.info("Todos los documentos fueron omitidos. No se creó ninguna ficha nueva para descargar o actualizar.")

        elif (pdfs or modo_solo_bd) and not exceso_archivos:
            if not modo_solo_bd:
                mapa_pdfs = {pdf.name: pdf for pdf in pdfs}
                nombres_actuales = [pdf.name for pdf in pdfs]

                if ("orden_nombres_pdf" not in st.session_state or 
                    set(st.session_state["orden_nombres_pdf"]) != set(nombres_actuales)):
                
                    orden_previo = st.session_state.get("orden_nombres_pdf", [])
                    nuevo_orden_filtrado = [nombre for nombre in orden_previo if nombre in mapa_pdfs]
                
                    for nombre in nombres_actuales:
                        if nombre not in nuevo_orden_filtrado:
                            nuevo_orden_filtrado.append(nombre)
                
                    st.session_state["orden_nombres_pdf"] = nuevo_orden_filtrado
                    for clave in ("cola_pdf_omitidos", "cola_pdf_procesados", "cola_pdf_finalizada_omision",
                                  "historial_edicion_tabla", "campos_restaurados_tabla"):
                        st.session_state.pop(clave, None)
                    st.session_state["cola_pdf_indice"] = 0
                    st.session_state.pop("data_extracted", None)
                    st.session_state.pop("user_edited_fields", None)
                    st.session_state.pop("extracted_text", None)
                    st.session_state.pop("result_bytes", None)
                    st.session_state.pop("discrepancias_pdf", None)
                    st.session_state.pop("modo_referencia_aplicado", None)
                    st.session_state.pop("seleccion_campos_excel", None)

                hash_archivos = "_".join(sorted(nombres_actuales))
                key_sortable_dinamico = f"sortable_queue_{st.session_state.get('generacion_cargas_area', 0)}_{hash_archivos}"

                if len(nombres_actuales) > 1:
                    st.markdown("##### Organiza el orden de la cola de procesamiento:")
                
                    col_posicion, col_fichas = st.columns([1, 14])

                    with col_posicion:
                        st.caption("Orden")
                        for i in range(1, len(st.session_state["orden_nombres_pdf"]) + 1):
                            st.markdown(f'<div class="num-slot">{i}.</div>', unsafe_allow_html=True)

                    with col_fichas:
                        st.caption("Documentos (arrastrar para ordenar)")
                        nuevo_orden_nombres = sort_items(
                            st.session_state["orden_nombres_pdf"],
                            direction="vertical",
                            key=key_sortable_dinamico
                        )

                    if nuevo_orden_nombres and nuevo_orden_nombres != st.session_state["orden_nombres_pdf"]:
                        st.session_state["orden_nombres_pdf"] = nuevo_orden_nombres
                        for clave in ("cola_pdf_omitidos", "cola_pdf_procesados", "cola_pdf_finalizada_omision",
                                      "historial_edicion_tabla", "campos_restaurados_tabla"):
                            st.session_state.pop(clave, None)
                        st.session_state["cola_pdf_indice"] = 0
                        st.session_state.pop("data_extracted", None)
                        st.session_state.pop("user_edited_fields", None)
                        st.session_state.pop("extracted_text", None)
                        st.session_state.pop("result_bytes", None)
                        st.session_state.pop("bytes_plantilla_inicial", None)
                        st.session_state.pop("fecha_ultima_extraccion_ficha", None)
                        st.rerun()

                pdfs_ordenados = [mapa_pdfs[nombre] for nombre in st.session_state["orden_nombres_pdf"] if nombre in mapa_pdfs]

                total_pdfs = len(pdfs_ordenados)
                idx_actual = st.session_state.get("cola_pdf_indice", 0)

                if idx_actual >= total_pdfs:
                    st.session_state["cola_pdf_indice"] = 0
                    idx_actual = 0

                pdf_activo = pdfs_ordenados[idx_actual]
                try:
                    clase_pdf, paginas_escaneadas, total_paginas = previsualizar_tipo_pdf(
                        pdf_activo.getvalue())
                    if clase_pdf == 'escaneado' and not es_manual_procedimiento:
                        st.warning('Este PDF parece escaneado y no contiene texto seleccionable. '
                                   'Para leerlo, elige «Documentos escaneados».')
                    elif clase_pdf == 'mixto':
                        st.info(f'Este PDF combina texto e imágenes: {paginas_escaneadas} de '
                                f'{total_paginas} páginas podrían necesitar lectura óptica. '
                                'Comprueba que la modalidad elegida sea adecuada.')
                    elif clase_pdf == 'texto' and es_manual_procedimiento:
                        st.info('Este PDF contiene texto seleccionable. Puedes utilizar '
                                '«Fichas técnicas» si corresponde a una ficha de fabricante; '
                                'mantén «Documentos escaneados» si necesitas interpretar un manual.')
                except Exception:
                    # La clasificación es solo orientativa: nunca impedir la carga.
                    pass

                if modo_conjunto:
                    st.info(f"**Análisis conjunto:** {total_pdfs} documento(s) para el examen **{examen_activo or 'por seleccionar'}**. Se generará una única ficha. Los valores contradictorios requieren revisión.")
                else:
                    st.markdown(f"""
                    <div class="banner-cola-trabajo">
                        <b>COLA DE PROCESAMIENTO ACTIVA:</b> Documento <b>{idx_actual + 1}</b> de <b>{total_pdfs}</b><br>
                        <span style="font-size: 0.90rem;">Archivo en revisión actual: <b>{pdf_activo.name}</b></span>
                    </div>
                    """, unsafe_allow_html=True)
    
                    st.progress((idx_actual + 1) / total_pdfs)
                    # La omisión solo existe en colas de varios exámenes; se ofrece
                    # ANTES de analizar, también para el último documento.
                    if total_pdfs > 1 and "data_extracted" not in st.session_state:
                        etiqueta_omitir = (
                            "Omitir último examen y finalizar cola"
                            if idx_actual == total_pdfs - 1 else
                            "Omitir este examen y pasar al siguiente"
                        )
                        if st.button(
                            etiqueta_omitir, type="secondary", key="btn_omitir_pdf_cola",
                            help="Este PDF no se analizará ni modificará las fichas ya procesadas.",
                        ):
                            _omitir_pdf_actual(idx_actual, total_pdfs, pdf_activo.name)
                            st.rerun()


            else:
                mapa_pdfs = {}
                pdfs_ordenados = []
                idx_actual = 0
                pdf_activo = None
                st.caption("Se generará la ficha con la información institucional disponible para el examen seleccionado.")

            st.divider()

            # --- EMPAREJAMIENTO INSTITUCIONAL: AMBAS CARTERAS CON EL MISMO FLUJO ---
            st.markdown("##### Emparejamiento con **Cartera de Exámenes del CRS**")
            lista_cartera = obtener_lista_cartera_db()
            seleccion_manual_cartera = None
            if not lista_cartera:
                st.warning("La Cartera de Exámenes del CRS no está cargada. Puedes actualizarla en Base de datos.")
            else:
                # Sincronizar ANTES de instanciar el selectbox. No basta comprobar
                # si cambió el examen: un estado antiguo puede mantener vacío
                # el widget aunque la Cartera tenga una coincidencia 100 %.
                revision_cartera = gp.estado_recurso(
                    DB_PATH, "fuente:cartera_examenes")["revision"]
                crs.sincronizar_vinculo_en_sesion(
                    st.session_state, DB_PATH, examen_activo, area_activa,
                    lista_cartera, revision=revision_cartera,
                )
                if "_sel_cartera_pendiente" in st.session_state:
                    pendiente = st.session_state.pop("_sel_cartera_pendiente")
                    if pendiente in lista_cartera:
                        st.session_state["sel_cartera_main"] = pendiente
                        st.session_state["cartera_seleccion_confirmada_manualmente"] = True
                    elif pendiente is not None:
                        st.session_state["sel_cartera_main"] = None
                seleccion_manual_cartera = st.selectbox(
                    "Examen vinculado en la Cartera",
                    options=lista_cartera,
                    index=None,
                    placeholder="Selecciona un examen (opcional)",
                    key="sel_cartera_main",
                    on_change=_cambiar_selector_cartera,
                    help="Selecciona un examen de la Cartera o confirma una coincidencia sugerida."
                )
                st.session_state["cartera_examen_emparejado"] = seleccion_manual_cartera or "Sin vincular"
                # Buscar las cinco mejores DESPUÉS de aplicar el umbral de
                # similitud. Si se limita primero, pueden mostrarse solo las
                # variantes por nombre aunque haya más opciones pertinentes.
                sugeridas_crs = (buscar_top_coincidencias_cartera(
                    examen_activo, area_activa, top_n=5, similitud_minima=0.55)
                    if examen_activo else [])
                if seleccion_manual_cartera:
                    with sqlite3.connect(DB_PATH) as con:
                        fila_lab = con.execute("SELECT laboratorio FROM cartera_examenes WHERE nombre_examen=?",
                                               (seleccion_manual_cartera,)).fetchone()
                    lab_vinculado = fila_lab[0] if fila_lab else ""
                    nombre_exacto = fd.normalizar(seleccion_manual_cartera) == fd.normalizar(examen_activo)
                    if nombre_exacto:
                        descripcion_vinculo = "exacta"
                    else:
                        descripcion_vinculo = "manual"
                    st.caption(f"Vinculada: **{seleccion_manual_cartera}** "
                               f"({lab_vinculado or 'Laboratorio sin registrar'}) · {descripcion_vinculo}")
                    # La selección automática del examen exacto no oculta los
                    # derivados cuyo nombre incluye el examen original completo.
                    # No se fuerza ninguna vinculación de candidatos difusos.
                    if nombre_exacto:
                        sugeridas_crs = crs.candidatos_variantes_cartera(
                            DB_PATH, examen_activo, area_activa, limite=5)
                    else:
                        sugeridas_crs = []
                elif examen_activo:
                    # Si el encargado quitó voluntariamente un vínculo exacto,
                    # no afirmar incorrectamente que no existe tal coincidencia.
                    if crs.coincidencia_exacta_cartera(DB_PATH, examen_activo, area_activa) is None:
                        st.warning("No hay coincidencia exacta por nombre y laboratorio. Selecciona la prestación adecuada para evitar datos de otro examen.")

                if sugeridas_crs:
                    st.markdown("<u><strong>Posibles coincidencias:</strong></u>", unsafe_allow_html=True)
                    columna_sugerencias, _ = st.columns([1, 1], gap="small")
                    with columna_sugerencias:
                        with st.container(key="lista_coincidencias_crs"):
                            for nombre_sugerido, laboratorio_sugerido, puntaje in sugeridas_crs:
                                mostrar_coincidencia_seleccionable(
                                    "crs", nombre_sugerido, nombre_sugerido,
                                    laboratorio_sugerido or "Laboratorio sin registrar", puntaje,
                                    _vincular_candidato_cartera, nombre_sugerido)
                elif examen_activo and not seleccion_manual_cartera:
                    st.caption("No se identificaron sugerencias suficientemente similares.")

            # Separación entre los dos bloques completos de emparejamiento.
            st.divider()
            # Vinculación independiente a «Prestación» (no se confunde con la cartera antigua).
            st.markdown("##### Emparejamiento con **Cartera de acreditación**")
            prestaciones_ui = fd.listar_prestaciones_acreditacion(DB_PATH)
            if prestaciones_ui:
                etiqueta_prestaciones = {r[0]: f"{r[1]} · {r[2]}" for r in prestaciones_ui}
                opciones_prestaciones = [r[0] for r in prestaciones_ui]
                # Igual que en la Cartera del CRS: actualizar el valor del
                # selector ANTES de instanciarlo, nunca después. Así la
                # prestación exacta aparece efectivamente seleccionada.
                if st.session_state.pop("_acreditacion_reset_pendiente", False):
                    st.session_state["acreditacion_manual_selector"] = None
                    st.session_state["acreditacion_manual_id"] = None
                    st.session_state["acreditacion_seleccion_confirmada_manualmente"] = False
                    st.session_state.pop("contexto_emparejamiento_acreditacion", None)
                revision_acreditacion = gp.estado_recurso(
                    DB_PATH, "fuente:cartera_acreditacion")["revision"]
                fd.sincronizar_vinculo_acreditacion_en_sesion(
                    st.session_state, DB_PATH, examen_activo, area_activa,
                    opciones_prestaciones, revision=revision_acreditacion,
                )
                identificador_manual = st.selectbox(
                    "Prestación vinculada en la Cartera de Acreditación",
                    options=opciones_prestaciones,
                    index=None,
                    placeholder="Selecciona una prestación (opcional)",
                    format_func=lambda id_p: etiqueta_prestaciones.get(id_p,"Desconocida"),
                    key="acreditacion_manual_selector",
                    on_change=_cambiar_selector_acreditacion,
                    help="Se utiliza «Prestación» como nombre del examen. Si no coincide exactamente, selecciona la prestación correcta manualmente."
                )
                st.session_state["acreditacion_manual_id"] = identificador_manual
                _, _, seleccion_previa, detalle_match = fd.obtener_valores_fuentes(
                    DB_PATH, examen_activo, area_activa, identificador_manual)
                if seleccion_previa:
                    # La función extractora considera 'manual' todo ID que
                    # recibe por parámetro; para el vínculo preseleccionado
                    # conservar la etiqueta exacta (o exacta por área).
                    detalle_match = (st.session_state.get("acreditacion_coincidencia_automatica")
                                     or detalle_match)
                    st.caption(f"Vinculada: **{seleccion_previa['prestacion']}** ({seleccion_previa['area']}) · {detalle_match}")
                elif examen_activo:
                    st.warning("No hay coincidencia exacta. Selecciona la prestación adecuada para evitar datos de otro examen.")
                    # Considerar también el examen que el usuario vinculó
                    # explícitamente en la Cartera CRS. Su nombre puede ser
                    # diferente al nombre administrativo del examen activo.
                    # El vínculo en Acreditación sigue requiriendo confirmación.
                    sugeridas = fd.candidatos_acreditacion(
                        DB_PATH, examen_activo, area_activa,
                        limite=5, similitud_minima=0.55,
                        nombre_referencia=seleccion_manual_cartera,
                    )
                    if sugeridas:
                        st.markdown("<u><strong>Posibles coincidencias:</strong></u>", unsafe_allow_html=True)
                        columna_sugerencias, _ = st.columns([1, 1], gap="small")
                        with columna_sugerencias:
                            with st.container(key="lista_coincidencias_acreditacion"):
                                for identificador, nombre_sugerido, area_sugerida, puntaje in sugeridas:
                                    mostrar_coincidencia_seleccionable(
                                        "acreditacion", identificador, nombre_sugerido,
                                        area_sugerida or "Área sin registrar", puntaje,
                                        _seleccionar_candidato_acreditacion, identificador)
                    else:
                        st.caption("No se identificaron sugerencias suficientemente similares.")
            else:
                st.warning("Aún no hay prestaciones cargadas en la Cartera de Acreditación.")

            st.divider()
            st.markdown("##### Parámetros operacionales de la ficha")

            c_lab, c_act, c_aut = st.columns(3)

            with c_lab:
                laboratorios = cargar_laboratorios()
                actual_lab = st.session_state.get("laboratorio_activo")
                key_lab = "sel_lab_main"
                if key_lab not in st.session_state and actual_lab in laboratorios:
                    st.session_state[key_lab] = actual_lab

                sel_lab = st.selectbox(
                    "Laboratorio de procesamiento", 
                    laboratorios, 
                    index=None,
                    placeholder="Selecciona laboratorio", 
                    key=key_lab
                )
                st.session_state["laboratorio_activo"] = sel_lab

            with c_act:
                encargados_act = cargar_encargados("actualizado")
                actual_act = st.session_state.get("actualizado_activo")
                key_act = "sel_act_main"
                if key_act not in st.session_state and actual_act in encargados_act:
                    st.session_state[key_act] = actual_act

                sel_act = st.selectbox(
                    "Actualizado por", 
                    encargados_act, 
                    index=None,
                    placeholder="Selecciona quién actualiza", 
                    key=key_act
                )
                st.session_state["actualizado_activo"] = sel_act

            with c_aut:
                encargados_aut = cargar_encargados("autorizado")
                actual_aut = st.session_state.get("autorizado_activo")
                key_aut = "sel_aut_main"
                if key_aut not in st.session_state and actual_aut in encargados_aut:
                    st.session_state[key_aut] = actual_aut

                sel_aut = st.selectbox(
                    "Autorizado por", 
                    encargados_aut, 
                    index=None,
                    placeholder="Selecciona quién autoriza", 
                    key=key_aut
                )
                st.session_state["autorizado_activo"] = sel_aut

            with st.expander("Administrar parámetros operacionales"):
                tab_m_lab, tab_m_act = st.tabs(["Laboratorios", "Encargados"])

                with tab_m_lab:
                    col_a, col_b = st.columns(2)
                    with col_a:
                        with st.form("form_add_lab_inline", clear_on_submit=True):
                            new_lab = st.text_input("Añadir nuevo laboratorio", key="nombre_nuevo_laboratorio")
                            if st.form_submit_button("Guardar laboratorio"):
                                nuevo_nombre = " ".join(new_lab.strip().split())
                                actuales = cargar_laboratorios()
                                if nuevo_nombre and nuevo_nombre.casefold() not in {x.casefold() for x in actuales}:
                                    solicitar_cambio("catalogo", "catalogo:laboratorios",
                                        descripcion=f"Añadir el laboratorio «{nuevo_nombre}» a la lista",
                                        tabla="laboratorios", nombres=actuales + [nuevo_nombre])
                                else:
                                    st.warning("Nombre no válido o ya existente.")
                    with col_b:
                        labs_del = cargar_laboratorios()
                        if labs_del:
                            with st.form("form_del_lab_inline"):
                                target_del = st.selectbox(
                                    "Eliminar laboratorio", 
                                    labs_del, 
                                    index=None, 
                                    placeholder="Selecciona laboratorio a eliminar",
                                    key="del_lab_select"
                                )
                                if st.form_submit_button("Eliminar seleccionado"):
                                    if target_del:
                                        updated = [x for x in labs_del if x.casefold() != target_del.casefold()]
                                        solicitar_cambio("catalogo", "catalogo:laboratorios",
                                            descripcion=f"Desactivar el laboratorio «{target_del}»",
                                            tabla="laboratorios", nombres=updated)

                with tab_m_act:
                    col_a, col_b = st.columns(2)
                    with col_a:
                        with st.form("form_add_act_inline", clear_on_submit=True):
                            new_act = st.text_input("Añadir encargado", key="nombre_nuevo_actualizado")
                            if st.form_submit_button("Guardar encargado"):
                                curr = cargar_encargados("actualizado")
                                if new_act.strip() and new_act.strip().casefold() not in {x.casefold() for x in curr}:
                                    curr.append(new_act.strip())
                                    solicitar_cambio("catalogo", "catalogo:encargados",
                                        descripcion=f"Añadir el encargado «{new_act.strip()}»",
                                        tabla="encargados", nombres=curr)
                                else:
                                    st.warning("Nombre no válido o ya existente.")
                    with col_b:
                        acts_del = cargar_encargados("actualizado")
                        if acts_del:
                            with st.form("form_del_act_inline"):
                                target_del = st.selectbox(
                                    "Eliminar de la lista", 
                                    acts_del, 
                                    index=None, 
                                    placeholder="Selecciona encargado a eliminar",
                                    key="del_act_select"
                                )
                                if st.form_submit_button("Eliminar seleccionado"):
                                    if target_del:
                                        updated = [x for x in acts_del if x.casefold() != target_del.casefold()]
                                        solicitar_cambio("catalogo", "catalogo:encargados",
                                            descripcion=f"Desactivar el encargado «{target_del}»",
                                            tabla="encargados", nombres=updated)

            st.divider()

            if not modo_solo_bd:
                # Permitir omitir portadas/anexos en documentos escaneados extensos.
                # No se omite ninguna página por defecto; el usuario elige el rango.
                rango_paginas_manual = None
                if es_manual_procedimiento:
                    if not modo_conjunto:
                        try:
                            with fitz.open(stream=pdf_activo.getvalue(), filetype="pdf") as doc_paginas:
                                cantidad_paginas_manual = doc_paginas.page_count
                            if cantidad_paginas_manual > 1:
                                rango_paginas_manual = st.slider(
                                    "Páginas del documento por analizar con lectura óptica",
                                    min_value=1, max_value=cantidad_paginas_manual,
                                    value=(1, cantidad_paginas_manual),
                                    key="paginas_manual_" + hashlib.sha256(
                                        pdf_activo.getvalue()).hexdigest()[:12],
                                    help="Para acelerar las pruebas puedes acotar las páginas; "
                                         "las páginas fuera del rango no aportarán información.")
                        except Exception:
                            pass

                # Vista previa aislada: las flechas navegan sin reconstruir toda la interfaz.
                with st.expander("Vista previa interactiva del documento PDF en revisión", expanded=True):
                    _visor_pdf_rapido(pdf_activo, idx_actual)

                # BOTÓN DE ANÁLISIS
                st.write("")
                col_btn_analisis, col_sp1 = st.columns([1.5, 2])
                with col_btn_analisis:
                    iniciar_analisis = st.button("Analizar documento", type="primary", key="analyze_pdf", disabled=exceso_archivos, use_container_width=False)

                if iniciar_analisis:
                    if not examen_activo:
                        st.warning("Selecciona un examen antes de continuar")
                    elif not st.session_state.get("laboratorio_activo"):
                        st.warning("Selecciona un laboratorio de procesamiento antes de analizar.")
                    else:
                        try:
                            # Una extracción nueva inicia con propuestas nuevas. Las
                            # correcciones anteriores se conservan en el Excel de
                            # origen, visible al desmarcar «Aplicar», no se reciclan.
                            ediciones_previas = {}
                            st.session_state.pop("data_extracted", None)
                            st.session_state.pop("user_edited_fields", None)
                            st.session_state.pop("extracted_text", None)
                            st.session_state.pop("conflictos_pendientes", None)

                            st.session_state["aggrid_key_counter"] += 1
                            st.session_state.pop("seleccion_campos_excel", None)

                            aliases_area = get_field_aliases(area_activa)
                            plantilla_bytes_db, _ = obtener_plantilla_area_db(area_activa)

                            if "result_bytes" in st.session_state and idx_actual > 0:
                                template_bytes_init = st.session_state["result_bytes"]
                            elif plantilla_bytes_db:
                                template_bytes_init = plantilla_bytes_db
                            else:
                                wb_vacio = openpyxl.Workbook()
                                ws_vacio = wb_vacio.active
                                ws_vacio.title = examen_activo
                                for fila, campo in enumerate(aliases_area, start=1):
                                    ws_vacio.cell(row=fila, column=1, value=campo)
                                    ws_vacio.cell(row=fila, column=2, value=examen_activo if norm(campo) == norm("Nombre LIS") else "")
                                buf = io.BytesIO()
                                wb_vacio.save(buf)
                                template_bytes_init = buf.getvalue()

                            st.session_state["bytes_plantilla_inicial"] = template_bytes_init
                            st.session_state["valores_anteriores_ficha"] = rv.valores_anteriores(
                                template_bytes_init, examen_activo, norm)
                            if "revision_plantilla_origen" not in st.session_state or idx_actual == 0:
                                st.session_state["revision_plantilla_origen"] = gp.estado_recurso(DB_PATH, "area:" + area_activa)

                            nombre_vinc = st.session_state.get("cartera_examen_emparejado")
                            info_cartera = obtener_info_cartera_por_nombre(nombre_vinc) if (nombre_vinc and nombre_vinc != "Sin vincular") else None

                            datos_sobrescribir = {}
                            conflictos = {}

                            if info_cartera and info_cartera.get("Código"):
                                datos_sobrescribir["Código"] = info_cartera["Código"]

                            # Los campos de equipamiento/plazo conservan la comparación con el PDF.
                            # «Indicaciones» del CRS ahora tiene máxima prioridad para preparación.
                            campos_comparables = ["Equipamiento utilizado", "Plazo entrega resultados"]

                            adicionales = [(pdf.name, pdf.getvalue()) for pdf in pdfs_ordenados[1:]] if modo_conjunto else []
                            if modo_conjunto:
                                parsed_pdf_data, fuentes, diferencias, raw_text, metodo_lectura_pdf = consolidar_documentos_pdf(
                                    [(pdf.name, pdf.getvalue()) for pdf in pdfs_ordenados], aliases_area, examen_activo,
                                    manual_procedimiento=es_manual_procedimiento
                                )
                                st.session_state["fuentes_pdf"] = fuentes
                                st.session_state["discrepancias_pdf"] = diferencias
                            else:
                                # Mostrar avance real de las páginas OCR, sin refrescar toda la aplicación.
                                estado_lectura_pdf = st.empty()
                                def _actualizar_progreso_pdf(pagina, total, uso_ocr):
                                    descripcion = (f"Leyendo PDF: página {pagina} de {total}"
                                                   if pagina else f"Iniciando lectura de {total} páginas")
                                    if uso_ocr:
                                        descripcion += " · Lectura óptica"
                                    estado_lectura_pdf.progress(
                                        pagina / max(total, 1), text=descripcion)
                                try:
                                    raw_text, metodo_lectura_pdf = extract_pdf_text(
                                        pdf_activo.getvalue(), pdf_activo.name, examen_activo,
                                        hematologia=es_alias_hematologia(aliases_area),
                                        progreso=_actualizar_progreso_pdf,
                                        rango_paginas=rango_paginas_manual,
                                        manual_procedimiento=es_manual_procedimiento)
                                finally:
                                    estado_lectura_pdf.empty()
                                # Conservar diagnóstico de lectura: si no se encuentra
                                # información, distinguir OCR legible de secciones no reconocidas.
                                diagnostico_manual_actual = (mp.analizar_manual(raw_text)
                                                             if es_manual_procedimiento else None)
                                manual_digital_reconocido = (
                                    not es_manual_procedimiento and
                                    mp.es_manual_de_procedimiento_textual(raw_text))
                                parsed_pdf_data = (dict(diagnostico_manual_actual["campos"])
                                                   if es_manual_procedimiento else
                                                   mp.completar_homologos_manual_digital(
                                                       raw_text, infer_name_code(
                                                           raw_text, parse_labeled_text(
                                                               raw_text, aliases_area, examen_activo))))
                                if area_activa == "Derivados":
                                    parsed_pdf_data.update(de.extraer_ficha_derivados(
                                        raw_text, pdf_activo.getvalue()))
                                # «Referencias» solo entra en la ficha tras pulsar
                                # expresamente el botón del panel bibliográfico.
                                parsed_pdf_data.pop("Referencias", None)
                                for campo_manual in fd.CAMPOS_MANUAL:
                                    parsed_pdf_data.pop(campo_manual, None)
                                if parsed_pdf_data.get("Equipamiento utilizado"):
                                    parsed_pdf_data["Equipamiento utilizado"] = eq.deduplicar_equipamiento(
                                        parsed_pdf_data["Equipamiento utilizado"])
                                manuales_permitidos = (
                                    set(fd.CAMPOS_CONTROLADOS) if
                                    (es_manual_procedimiento or area_activa == "Derivados") else
                                    {'Muestra requerida y recipiente'}
                                    if manual_digital_reconocido else set())
                                parsed_pdf_data = {campo: val for campo, val in parsed_pdf_data.items()
                                                   if campo not in fd.CAMPOS_MANUAL
                                                   and (campo not in fd.CAMPOS_CONTROLADOS
                                                        or campo in manuales_permitidos)}
                                st.session_state.pop("fuentes_pdf", None)
                                st.session_state.pop("discrepancias_pdf", None)

                            if info_cartera:
                                for campo in campos_comparables:
                                    val_pdf = str(parsed_pdf_data.get(campo, "")).strip()
                                    val_cartera = str(info_cartera.get(campo, "")).strip()
                                    if campo == "Equipamiento utilizado":
                                        val_pdf = eq.deduplicar_equipamiento(val_pdf)
                                        val_cartera = eq.deduplicar_equipamiento(val_cartera)
                                        if val_pdf:
                                            parsed_pdf_data[campo] = val_pdf

                                    if val_cartera:
                                        datos_sobrescribir[campo] = val_cartera
                                        if (val_pdf and (not eq.equipamientos_equivalentes(val_pdf, val_cartera)
                                                if campo == "Equipamiento utilizado" else
                                                ' '.join(val_pdf.casefold().split()) !=
                                                ' '.join(val_cartera.casefold().split()))):
                                            conflictos[campo] = {
                                                "pdf": val_pdf, "cartera": val_cartera,
                                                "fuente": "Cartera de Exámenes del CRS",
                                            }

                            # Rastrear los PDF que aportaron datos reales a esta ficha.
                            if modo_conjunto:
                                nombres_que_aportaron = set(st.session_state.get("fuentes_pdf", {}).values())
                                for opciones_conflicto in st.session_state.get("discrepancias_pdf", {}).values():
                                    nombres_que_aportaron.update(nombre for nombre, _ in opciones_conflicto)
                            else:
                                nombres_que_aportaron = {pdf_activo.name} if any(
                                    str(valor or "").strip() for campo, valor in parsed_pdf_data.items()
                                    if campo not in fd.CAMPOS_CONTROLADOS
                                    and campo not in ("Nombre LIS", "Fecha actualización", "Referencias")
                                ) else set()
                            st.session_state["pdf_documentos_con_aporte"] = sorted(nombres_que_aportaron)
                            # Registrar si el PDF aportó campos distintos de los documentales.
                            st.session_state["pdf_aporto_campos"] = any(
                                str(valor or "").strip() for campo, valor in parsed_pdf_data.items()
                                if campo not in fd.CAMPOS_CONTROLADOS
                                and campo not in ("Nombre LIS", "Fecha actualización", "Referencias")
                            )
                            # Solo registrar la Cartera tradicional como fuente si aportó valores.
                            st.session_state["cartera_aporto_campos"] = bool(
                                info_cartera and (datos_sobrescribir or conflictos)
                            )
                            st.session_state.pop("modo_referencia_aplicado", None)
                            st.session_state.pop("texto_referencias_auto", None)
                            st.session_state.pop("fuentes_referencias_aplicadas", None)

                            # Preparación: Indicaciones del CRS > Acreditación > Normas.
                            # Los otros campos con fuentes mantienen sus prioridades;
                            # Recepción y Procesamiento permanecen sin origen automático.
                            # Ningún PDF técnico reemplaza estos campos documentales.
                            valores_doc, origen_doc, prestacion_doc, estado_vinculo = fd.obtener_valores_fuentes(
                                DB_PATH, examen_activo, area_activa,
                                st.session_state.get("acreditacion_manual_id"),
                                indicaciones_cartera=(info_cartera or {}).get("Preparación del paciente", ""),
                            )
                            datos_sobrescribir.update(valores_doc)
                            # «Test incluidos» se nutre de Homologació de test
                            # solo cuando código/nombre/área identifican una fila sin dudas.
                            test_homologado, estado_homologacion = homologacion_del_examen(
                                examen_activo, area_activa, info_cartera, nombre_vinc)
                            if test_homologado:
                                datos_sobrescribir["Test incluidos"] = test_homologado
                                origen_doc["Test incluidos"] = htest.ETIQUETA
                                valores_doc["Test incluidos"] = test_homologado
                                test_pdf = str(parsed_pdf_data.get("Test incluidos", "") or "").strip()
                                if test_pdf and de.son_valores_distintos(
                                        "Test incluidos", test_pdf, test_homologado):
                                    conflictos["Test incluidos"] = {
                                        "pdf": test_pdf, "cartera": test_homologado,
                                        "fuente": htest.ETIQUETA,
                                    }
                            elif estado_homologacion.get("estado") == "requiere_revision":
                                st.warning("Homologació de test contiene registros ambiguos para "
                                           "este examen; Test incluidos no se completará desde esa fuente.")
                            if area_activa == "Derivados":
                                # Comparar los candidatos de una fila documentada con
                                # cualquier fuente institucional habilitada y con valor.
                                # La prioridad define la propuesta, pero NO oculta una
                                # diferencia: siempre se ofrecen ambas opciones.
                                for campo, propuesto in parsed_pdf_data.items():
                                    if campo == 'Código' or campo in conflictos:
                                        continue
                                    valor_inst = str(datos_sobrescribir.get(campo, '') or '').strip()
                                    if not de.son_valores_distintos(campo, propuesto, valor_inst):
                                        continue
                                    fuente_inst = str(origen_doc.get(campo, '') or '')
                                    if not fuente_inst and campo in ('Equipamiento utilizado', 'Plazo entrega resultados'):
                                        fuente_inst = 'Cartera de Exámenes del CRS'
                                    if not fuente_inst:
                                        continue
                                    conflictos[campo] = {
                                        'pdf': str(propuesto).strip(), 'cartera': valor_inst,
                                        'fuente': fuente_inst,
                                    }
                            if es_manual_procedimiento:
                                # Escáner primero SOLO donde esta procedencia existe;
                                # Código conserva Cartera prioritaria. Una fuente
                                # elegida por el usuario tiene preferencia explícita.
                                eleccion_crs = bool(st.session_state.get(
                                    "cartera_seleccion_confirmada_manualmente"))
                                eleccion_acreditacion = bool(st.session_state.get(
                                    "acreditacion_seleccion_confirmada_manualmente"))
                                elegidos_expresamente = set()
                                for campo, valor_ocr_bruto in parsed_pdf_data.items():
                                    if (campo == "Código" or campo in (
                                            "Nombre LIS", "Laboratorio de procesamiento",
                                            "Actualizado por", "Autorizado por",
                                            "Fecha actualización", "Referencias")):
                                        continue
                                    valor_ocr = str(valor_ocr_bruto or "").strip()
                                    if not valor_ocr:
                                        continue
                                    origen_institucional = str(origen_doc.get(campo, ""))
                                    es_expresa = (
                                        (origen_institucional.startswith("Cartera de Acreditación")
                                         and eleccion_acreditacion)
                                        or (origen_institucional.startswith("Cartera de Exámenes del CRS")
                                            and eleccion_crs)
                                        or (campo in ("Equipamiento utilizado", "Plazo entrega resultados")
                                            and eleccion_crs and bool(info_cartera))
                                    )
                                    if es_expresa:
                                        elegidos_expresamente.add(campo)
                                        valor_institucional = str(datos_sobrescribir.get(campo, "") or "").strip()
                                        diferente = (not eq.equipamientos_equivalentes(valor_ocr, valor_institucional)
                                                     if campo == "Equipamiento utilizado" else
                                                     ' '.join(valor_ocr.casefold().split()) !=
                                                     ' '.join(valor_institucional.casefold().split()))
                                        if valor_institucional and diferente:
                                            conflictos[campo] = {
                                                "pdf": valor_ocr, "cartera": valor_institucional,
                                                "fuente": ("Cartera de Acreditación" if
                                                           origen_institucional.startswith("Cartera de Acreditación")
                                                           else "Cartera de Exámenes del CRS"),
                                            }
                                    else:
                                        # Dejar que process() aplique el texto óptico
                                        # como propuesta en lugar del dato institucional.
                                        datos_sobrescribir.pop(campo, None)
                                st.session_state["fuentes_elegidas_usuario"] = elegidos_expresamente
                            else:
                                st.session_state["fuentes_elegidas_usuario"] = set()
                            # La Cartera se considera utilizada solo en los
                            # campos cuya propuesta final procede de ella.
                            st.session_state["cartera_aporto_campos"] = bool(info_cartera and any(
                                str(info_cartera.get(campo, "") or "").strip()
                                and str(info_cartera.get(campo, "") or "").strip() ==
                                    str(datos_sobrescribir.get(campo, "") or "").strip()
                                for campo in ("Código", "Equipamiento utilizado",
                                              "Plazo entrega resultados", "Preparación del paciente")
                            ))
                            st.session_state["origen_campos_documentales"] = origen_doc
                            st.session_state["procedencia_valores_documentales"] = dict(valores_doc)
                            st.session_state["prestacion_acreditacion_vinculada"] = prestacion_doc
                            st.session_state["estado_vinculacion_acreditacion"] = estado_vinculo

                            out_bytes, data, target, score, reason, method, text = process(
                                pdf_activo.getvalue(), pdf_activo.name, template_bytes_init,
                                st.session_state.get("laboratorio_activo", ""),
                                examen_activo, aliases_area,
                                st.session_state.get("actualizado_activo", ""),
                                st.session_state.get("autorizado_activo", ""),
                                datos_sobrescribir=datos_sobrescribir,
                                documentos_adicionales=adicionales,
                                datos_pdf_preextraidos=(parsed_pdf_data, raw_text, metodo_lectura_pdf),
                                manual_procedimiento=es_manual_procedimiento
                            )

                            if data.get("Equipamiento utilizado"):
                                data["Equipamiento utilizado"] = eq.deduplicar_equipamiento(
                                    data["Equipamiento utilizado"])
                            if es_manual_procedimiento:
                                # No confundir "PDF leído" con "PDF que aportó valores".
                                # En un examen ya validado puede haber coincidencias
                                # que no se aplican porque la celda está ocupada.
                                diag = diagnostico_manual_actual
                                st.session_state["manual_ocr_diagnostico"] = diag
                                if diag["campos"]:
                                    # El fondo del aviso lo proporciona el tema de Streamlit.
                                    # Cada nombre se resalta de forma legible en ambos modos,
                                    # sin alterar los otros mensajes informativos de la aplicación.
                                    campos_destacados = ", ".join(
                                        "**" + re.sub(r"([\\*_[\]`])", r"\\\1", str(campo)) + "**"
                                        for campo in diag["campos"]
                                    )
                                    with st.container(key="aviso_campos_manual"):
                                        st.info("El manual permitió reconocer los siguientes campos: "
                                                + campos_destacados + ". "
                                                "Revisa los datos propuestos antes de guardarlos.")
                                else:
                                    st.warning("La lectura óptica terminó, pero no se identificó información suficiente para los campos de la ficha.")
                            else:
                                st.session_state.pop("manual_ocr_diagnostico", None)

                            # Las ediciones manuales confirmadas conservan la prioridad máxima.
                            if ediciones_previas:
                                data.update(ediciones_previas)
                                for campo, valor in ediciones_previas.items():
                                    out_bytes = actualizar_campo_excel(out_bytes, examen_activo, campo, valor)
                            # Reconstruir el formato documental que no está
                            # disponible en una celda de texto plano de AG Grid.
                            # La fuente se comprueba campo por campo: otras Carteras
                            # y las modificaciones manuales no se reinterpretan.
                            origen_para_excel = dict(origen_doc)
                            for campo in ediciones_previas:
                                origen_para_excel.pop(campo, None)
                            out_bytes = aplicar_formato_normas_excel(
                                out_bytes, examen_activo, origen_para_excel)
                            st.session_state["result_bytes"] = out_bytes
                            # La fecha del nombre corresponde a la extracción exitosa,
                            # no al día en que se descargue el documento.
                            st.session_state["fecha_ultima_extraccion_ficha"] = fecha_hoy_nombre_archivo()
                            st.session_state["data_extracted"] = copy.deepcopy(data)
                            st.session_state["manual_pdf_digital_detectado"] = (
                                not es_manual_procedimiento and
                                mp.es_manual_de_procedimiento_textual(raw_text))
                            st.session_state["procedencia_pdf_campos"] = dict(parsed_pdf_data)
                            st.session_state["procedencia_cartera_campos"] = {
                                campo: valor for campo, valor in datos_sobrescribir.items()
                                if campo not in fd.CAMPOS_CONTROLADOS
                            }
                            st.session_state["procedencia_modalidad"] = modo_procesamiento
                            st.session_state["user_edited_fields"] = set(ediciones_previas)
                            st.session_state.pop("historial_edicion_tabla", None)
                            st.session_state.pop("campos_restaurados_tabla", None)
                            if not modo_conjunto:
                                procesados = st.session_state.setdefault("cola_pdf_procesados", [])
                                if pdf_activo.name not in procesados:
                                    procesados.append(pdf_activo.name)
                            st.session_state["examen_resultado"] = examen_activo
                            st.session_state["extracted_text"] = text
                            st.session_state["area_resultado"] = area_activa
                            gp.registrar_acceso(DB_PATH, st.session_state["usuario_acceso"],
                                "Ficha generada", f"Área: {area_activa}; examen: {examen_activo}")

                            st.session_state.pop("conflictos_pendientes", None)
                            st.session_state.pop("pdf_para_cartera", None)
                            st.session_state["modal_conflictos_abierto"] = False
                            st.session_state["modal_discrepancias_unificado"] = False
                            if conflictos:
                                # Cada análisis ofrece inicialmente la propuesta de
                                # la Cartera seleccionada, sin heredar una elección
                                # de discrepancia de un examen anterior.
                                for campo_conflicto in conflictos:
                                    clave_preferencia = ("unificado_cartera_" +
                                         hashlib.sha256(campo_conflicto.encode()).hexdigest()[:16])
                                    st.session_state.pop(clave_preferencia, None)
                                st.session_state["conflictos_pendientes"] = conflictos
                                # PDF vs Cartera se abre solo por su propio botón,
                                # nunca simultáneamente con PDF vs PDF.

                            # El clic ya inició una ejecución completa. La revisión se
                            # renderiza justo debajo en esta misma ejecución: forzar
                            # st.rerun() reconstruía de nuevo la página y podía saltar
                            # la vista desde el botón hasta el final de la tabla.
                            st.success(f"¡Se procesó correctamente el archivo «{pdf_activo.name}»!")

                        except Exception as e:
                            st.error(f"No se pudo analizar el documento: {e}")
            else:
                st.divider()
                if st.button("Generar ficha con datos institucionales", type="primary",
                             key="generar_ficha_solo_bd"):
                    if not examen_activo:
                        st.warning("Selecciona un examen antes de continuar.")
                    elif not st.session_state.get("laboratorio_activo"):
                        st.warning("Selecciona el laboratorio de procesamiento antes de continuar.")
                    else:
                        try:
                            aliases_area = get_field_aliases(area_activa)
                            plantilla_bytes_db, _ = obtener_plantilla_area_db(area_activa)
                            if plantilla_bytes_db:
                                plantilla = plantilla_bytes_db
                            else:
                                wb_vacio = openpyxl.Workbook()
                                ws_vacio = wb_vacio.active
                                ws_vacio.title = titulo_hoja_para_examen(wb_vacio, examen_activo)
                                for fila, campo in enumerate(aliases_area, start=1):
                                    ws_vacio.cell(fila, 1, campo)
                                    ws_vacio.cell(fila, 2, examen_activo if norm(campo) == norm("Nombre LIS") else "")
                                buf = io.BytesIO()
                                wb_vacio.save(buf)
                                plantilla = buf.getvalue()
                            nombre_vinculado = st.session_state.get("cartera_examen_emparejado")
                            info_cartera = (obtener_info_cartera_por_nombre(nombre_vinculado)
                                            if nombre_vinculado and nombre_vinculado != "Sin vincular" else None)
                            valores_doc, origen_doc, prestacion_doc, estado_vinculo = fd.obtener_valores_fuentes(
                                DB_PATH, examen_activo, area_activa,
                                st.session_state.get("acreditacion_manual_id"),
                                indicaciones_cartera=(info_cartera or {}).get("Preparación del paciente", ""),
                            )
                            test_homologado, estado_homologacion = homologacion_del_examen(
                                examen_activo, area_activa, info_cartera, nombre_vinculado)
                            if test_homologado:
                                valores_doc["Test incluidos"] = test_homologado
                                origen_doc["Test incluidos"] = htest.ETIQUETA
                            elif estado_homologacion.get("estado") == "requiere_revision":
                                st.warning("Homologació de test contiene una asociación ambigua; "
                                           "revisa el examen antes de completar Test incluidos.")
                            institucional = datos_para_ficha_sin_pdf(info_cartera, valores_doc)
                            if not institucional:
                                st.warning("No hay datos institucionales vinculados para este examen. "
                                           "Selecciona una prestación o carga las Carteras desde Base de datos.")
                            else:
                                st.session_state["bytes_plantilla_inicial"] = plantilla
                                st.session_state["valores_anteriores_ficha"] = rv.valores_anteriores(
                                    plantilla, examen_activo, norm)
                                st.session_state["revision_plantilla_origen"] = gp.estado_recurso(
                                    DB_PATH, "area:" + area_activa)
                                st.session_state["aggrid_key_counter"] += 1
                                st.session_state.pop("seleccion_campos_excel", None)
                                salida, datos, *_ = process(
                                    b"", "", plantilla,
                                    st.session_state.get("laboratorio_activo", ""),
                                    examen_activo, aliases_area,
                                    st.session_state.get("actualizado_activo", ""),
                                    st.session_state.get("autorizado_activo", ""),
                                    datos_sobrescribir=institucional,
                                    datos_pdf_preextraidos=({}, "", "Sin análisis de PDF"),
                                    manual_procedimiento=False,
                                )
                                salida = aplicar_formato_normas_excel(
                                    salida, examen_activo, origen_doc)
                                st.session_state["result_bytes"] = salida
                                st.session_state["data_extracted"] = copy.deepcopy(datos)
                                st.session_state["fecha_ultima_extraccion_ficha"] = fecha_hoy_nombre_archivo()
                                st.session_state["extracted_text"] = ""
                                st.session_state["examen_resultado"] = examen_activo
                                st.session_state["area_resultado"] = area_activa
                                st.session_state["user_edited_fields"] = set()
                                st.session_state["procedencia_modalidad"] = "Solo base de datos"
                                st.session_state["origen_campos_documentales"] = dict(origen_doc)
                                st.session_state["procedencia_valores_documentales"] = dict(valores_doc)
                                st.session_state["procedencia_pdf_campos"] = {}
                                st.session_state["procedencia_cartera_campos"] = {
                                    campo: valor for campo, valor in institucional.items()
                                    if campo not in fd.CAMPOS_CONTROLADOS
                                }
                                st.session_state["prestacion_acreditacion_vinculada"] = prestacion_doc
                                st.session_state["estado_vinculacion_acreditacion"] = estado_vinculo
                                st.session_state["cartera_aporto_campos"] = bool(info_cartera and any(
                                    str(info_cartera.get(c, "") or "").strip()
                                    == str(institucional.get(c, "") or "").strip()
                                    for c in ("Código", "Equipamiento utilizado", "Plazo entrega resultados",
                                              "Preparación del paciente")
                                    if str(info_cartera.get(c, "") or "").strip()
                                ))
                                st.session_state["pdf_aporto_campos"] = False
                                st.session_state["pdf_documentos_con_aporte"] = []
                                st.session_state["fuentes_elegidas_usuario"] = set()
                                st.session_state["manual_pdf_digital_detectado"] = False
                                st.session_state["discrepancias_pdf"] = {}
                                st.session_state["conflictos_pendientes"] = {}
                                st.session_state.pop("historial_edicion_tabla", None)
                                st.session_state.pop("campos_restaurados_tabla", None)
                                st.session_state.pop("texto_referencias_auto", None)
                                gp.registrar_acceso(DB_PATH, st.session_state["usuario_acceso"],
                                    "Ficha generada", f"Área: {area_activa}; examen: {examen_activo}; origen: institucional")
                                st.success("Ficha generada con información de las fuentes institucionales. "
                                           "Revisa cada campo antes de guardarla.")
                        except Exception as e:
                            st.error(f"No se pudo generar la ficha desde la base de datos: {e}")
        elif not pdfs:
            st.session_state.pop("result_bytes", None)
            st.session_state.pop("data_extracted", None)
            st.session_state.pop("manual_ocr_diagnostico", None)
            st.session_state.pop("manual_pdf_digital_detectado", None)
            st.session_state.pop("orden_nombres_pdf", None)
            for clave in ("cola_pdf_omitidos", "cola_pdf_procesados", "cola_pdf_finalizada_omision",
                          "historial_edicion_tabla", "campos_restaurados_tabla"):
                st.session_state.pop(clave, None)
            st.info("Carga uno o más PDF para comenzar la extracción en cola.")

        if "data_extracted" in st.session_state and "result_bytes" in st.session_state and (pdfs or modo_solo_bd):
            data = st.session_state.get("data_extracted", {})
            user_edited = st.session_state.get("user_edited_fields", set())
            area_res = st.session_state.get("area_resultado", area_activa)
            aliases_area = get_field_aliases(area_res)

            pdfs_ordenados_rev = [mapa_pdfs[nombre] for nombre in st.session_state.get("orden_nombres_pdf", []) if nombre in mapa_pdfs] if 'mapa_pdfs' in locals() else pdfs
            pdf_actual_nombre = pdfs_ordenados_rev[st.session_state.get('cola_pdf_indice', 0)].name if pdfs_ordenados_rev else "fuentes institucionales"

            st.subheader("Revisión de fuentes institucionales" if modo_solo_bd else
                         f"Revisión e información extraída de «{pdf_actual_nombre}»")
            diferencias_pdf = st.session_state.get("discrepancias_pdf", {})
            diferencias_cartera = st.session_state.get("conflictos_pendientes", {})
            if diferencias_pdf or diferencias_cartera:
                if diferencias_pdf:
                    st.warning(_mensaje_discrepancias(len(diferencias_pdf), "los documentos PDF"))
                if diferencias_cartera:
                    st.warning(_mensaje_discrepancias(len(diferencias_cartera), "el PDF y las fuentes institucionales"))
                if st.button("Comparar y resolver discrepancias", type="primary", key="abrir_discrepancias_unificadas"):
                    st.session_state["modal_discrepancias_unificado"] = True
                    st.rerun()
            elif modo_conjunto:
                st.success("No hay discrepancias pendientes de revisión.")

            with st.expander("Procedencia de campos documentales", expanded=False):
                st.caption("Fuentes de mayor a menor prioridad. Se destaca la utilizada en el examen.")
                origen_registrado = st.session_state.get("origen_campos_documentales", {})
                filas_cuaderno = []
                # Exactamente el mismo orden de la tabla de revisión.
                for campo in aliases_area:
                    procedencia = origen_registrado.get(campo, "")
                    prioridades, fuente_aplicada = pc.procedencia_campo(
                        campo, data.get(campo, ""), procedencia,
                        st.session_state.get("procedencia_pdf_campos", {}),
                        st.session_state.get("procedencia_cartera_campos", {}),
                        st.session_state.get("user_edited_fields", set()),
                        es_manual_procedimiento,
                        st.session_state.get("procedencia_valores_documentales", {}),
                        st.session_state.get("texto_referencias_auto", ""),
                        st.session_state.get("fuentes_elegidas_usuario", set()),
                        manual_digital=st.session_state.get("manual_pdf_digital_detectado", False),
                        es_derivados=(area_res == "Derivados"),
                    )
                    nombres = []
                    for nombre_fuente in prioridades:
                        seleccionada = nombre_fuente == fuente_aplicada
                        clase = "cuaderno-fuente-seleccionada" if seleccionada else ""
                        sufijo = " (utilizada)" if seleccionada else ""
                        nombres.append(
                            f'<span class="{clase}" title="{html.escape(str(procedencia or fuente_aplicada or "SIN FUENTE APLICADA"), quote=True)}">'
                            f'{html.escape(nombre_fuente + sufijo)}</span>'
                        )
                    prioridades_html = '<span class="cuaderno-separador">&gt;</span>'.join(nombres)
                    if not fuente_aplicada and campo not in fd.CAMPOS_MANUAL:
                        aviso_sin_fuente = '<span class="cuaderno-sin-fuente">SIN FUENTE APLICADA</span>'
                        prioridades_html = ((prioridades_html + ' · ') if prioridades_html else '') + aviso_sin_fuente
                    filas_cuaderno.append(
                        '<div class="cuaderno-fila">'
                        f'<div class="cuaderno-campo">{html.escape(campo)}</div>'
                        f'<div class="cuaderno-fuente">{prioridades_html}</div>'
                        '</div>'
                    )
                st.markdown('<div class="cuaderno-procedencia">' + ''.join(filas_cuaderno) + '</div>',
                            unsafe_allow_html=True)

            # --- REFERENCIAS APA CON DOS VISTAS PREVIAS, SIN DIÁLOGOS ---
            if "Referencias" in aliases_area:
                st.subheader("Referencias", anchor=False)
                # Mantener la identidad del contenedor para conservar el menú
                # abierto o cerrado tal como lo dejó el usuario al confirmar.
                with st.container(key="contenedor_referencias"):
                    with st.expander("Referencias bibliográficas", expanded=True):
                      refs_pdf, refs_todas = _fuentes_para_referencias(
                          pdfs_ordenados_rev, st.session_state.get("cola_pdf_indice", 0),
                          examen_activo, data,
                          st.session_state.get("origen_campos_documentales", {}),
                          st.session_state.get("cartera_aporto_campos", False),
                      )
                      ref_pdf_texto = apa.ordenar_y_unir(refs_pdf)
                      ref_todas_texto = apa.ordenar_y_unir(refs_todas)
                      ref_pdf_html = apa.ordenar_y_unir_html(refs_pdf)
                      ref_todas_html = apa.ordenar_y_unir_html(refs_todas)
                      col_ref1, col_ref2 = st.columns(2)
                      with col_ref1:
                          if modo_solo_bd:
                              st.markdown("**Sin PDF técnico (análisis omitido)**")
                          else:
                              st.markdown("**Opción 1 · Solo PDF técnico**")
                          st.markdown(f"<div style='border:1px solid #dbe2ea;border-radius:10px;padding:0.85rem 0.95rem;min-height:170px;background:#0f172a08;white-space:pre-wrap'>{ref_pdf_html or 'Sin fuentes técnicas identificadas'}</div>", unsafe_allow_html=True)
                      with col_ref2:
                          st.markdown("**Fuentes institucionales utilizadas**" if modo_solo_bd
                                      else "**Opción 2 · Todas las fuentes utilizadas**")
                          st.markdown(f"<div style='border:1px solid #dbe2ea;border-radius:10px;padding:0.85rem 0.95rem;min-height:170px;background:#0f172a08;white-space:pre-wrap'>{ref_todas_html or 'Sin fuentes identificadas'}</div>", unsafe_allow_html=True)
                      modo_referencia = st.radio(
                          "¿Qué referencias deseas incorporar a la ficha?",
                          (["Todas las fuentes utilizadas"] if modo_solo_bd else
                           ["Solo PDF técnico", "Todas las fuentes utilizadas"]),
                          key="modo_referencias_apa", horizontal=True,
                      )
                      seleccionadas = refs_todas if modo_referencia == "Todas las fuentes utilizadas" else refs_pdf
                      texto_elegido = ref_todas_texto if modo_referencia == "Todas las fuentes utilizadas" else ref_pdf_texto
                      avisos = apa.advertencias_fuentes(seleccionadas)
                      if avisos:
                          st.warning("Los metadatos bibliográficos de algunas fuentes son incompletos. Comprueba los datos antes de aprobar la ficha.")
                          # Avisos visibles sin crear un segundo desplegable.
                          # El usuario sigue viendo los datos que debe revisar.
                          for aviso in avisos:
                              st.caption("• " + aviso)
                      elif texto_elegido:
                          st.success("Las fuentes seleccionadas cuentan con autor, fecha y título identificados.")
                      texto_actual_ref = str(data.get("Referencias", "") or "").strip()
                      editar_existente = ("Referencias" in user_edited and bool(texto_actual_ref))
                      permitir_reemplazo_ref = True
                      if editar_existente:
                          st.warning("El campo Referencias ya fue editado manualmente. No se reemplazará sin tu autorización.")
                          permitir_reemplazo_ref = st.checkbox("Confirmo reemplazar mis referencias editadas", key="confirmar_reemplazo_referencias")
                      if st.button("Aplicar opción seleccionada a Referencias", type="primary",
                                   disabled=(not bool(texto_elegido) or not permitir_reemplazo_ref),
                                   key="btn_aplicar_referencias_apa"):
                          st.session_state["data_extracted"]["Referencias"] = texto_elegido
                          st.session_state["texto_referencias_auto"] = texto_elegido
                          st.session_state["fuentes_referencias_aplicadas"] = seleccionadas
                          st.session_state["result_bytes"] = actualizar_campo_excel(
                              st.session_state["result_bytes"], examen_activo, "Referencias", texto_elegido)
                          st.session_state.setdefault("user_edited_fields", set()).discard("Referencias")
                          st.session_state.pop("historial_edicion_tabla", None)
                          st.session_state.pop("campos_restaurados_tabla", None)
                          st.session_state["aggrid_key_counter"] += 1
                          st.session_state["modo_referencia_aplicado"] = modo_referencia
                          st.rerun()

            # Dibujar el contador DESPUÉS de recibir la edición del grid.
            # Un metric renderizado aquí mostraba el valor de la ejecución
            # anterior tras borrar una celda con Suprimir.
            espacio_campos_encontrados = st.empty()

            # El historial vive en la sesión de Streamlit para que sobreviva
            # a la reconstrucción del grid al editar o borrar con Suprimir.
            historial_ediciones = st.session_state.setdefault("historial_edicion_tabla", {})
            ht.contexto_nuevo(
                historial_ediciones,
                (area_res, st.session_state.get("examen_resultado"),
                 st.session_state.get("cola_pdf_indice")),
            )

            rows = []
            for field in aliases_area:
                value = data.get(field, "")
                if pd.isna(value) or str(value).strip().casefold() in ["nan", "none", "null"]:
                    value = ""
                if field in user_edited:
                    state = "Modificado por usuario"
                elif field in ["Nombre LIS", "Laboratorio de procesamiento", "Actualizado por", "Autorizado por"]:
                    state = "Ingresado por usuario" if value else "Pendiente de ingreso"
                elif field == "Fecha actualización":
                    state = "Asignado por sistema" if value else "Pendiente de asignación"
                else:
                    state = "Encontrado" if value else "No encontrado"
                rows.append({
                    "Actualizar": bool(st.session_state.get("seleccion_campos_excel", {}).get(field, True)),
                    "Campo": field.strip(),
                    "Información extraída": value,
                    "Valor anterior": st.session_state.get("valores_anteriores_ficha", {}).get(norm(field), ""),
                    "Estado": state,
                    ht.COLUMNA_COMANDO: "",
                })
            df_base = pd.DataFrame(rows)

            st.markdown(
                """
                <div style="font-size: 0.85rem; color: #6c757d; line-height: 1.6; margin-bottom: 12px;">
                    <span style="background-color: #B7F7C1; color: #124D20; padding: 2px 8px; border-radius: 4px; font-weight: bold;">Verde:</span> Campo completado correctamente.
                    &nbsp;&nbsp;
                    <span style="background-color: #BAE6FD; color: #0369A1; padding: 2px 8px; border-radius: 4px; font-weight: bold;">Azul:</span> Campo modificado por usuario.
                    &nbsp;&nbsp;
                    <span style="background-color: #FFF0B3; color: #6B4D00; padding: 2px 8px; border-radius: 4px; font-weight: bold;">Amarillo:</span> Campo pendiente de revisión / no encontrado.
                    &nbsp;&nbsp;
                    <span style="background-color: #D99B60; color: #472509; padding: 2px 8px; border-radius: 4px; font-weight: bold;">Naranja:</span> Campo completado con valor anterior.
                    <br><br>
                    <ul style="margin: 0; padding-left: 1.25rem; list-style-type: disc;">
                        <li>Las casillas de la columna <b>Aplicar</b> indican qué campos reemplazar en el Excel al descargar. Desmarca un campo para conservar el valor anterior.</li>
                        <li>Haz doble clic sobre cualquier celda de la columna <b>Información extraída</b> para editar su contenido directamente en la tabla. El editor admite textos extensos.</li>
                        <li>Presionar <b>Ctrl + Z</b> o hacer clic en <b>Deshacer</b> permitirá revertir los cambios.</li>
                    </ul>
                </div>
                """,
                unsafe_allow_html=True
            )

            # Altura fija antes de recibir el grid: aunque el botón se
            # complete después del evento (para habilitar el PRIMER Deshacer),
            # el componente AG Grid nunca ocupa temporalmente su posición.
            with st.container(height=62, border=False, key="zona_deshacer_estable"):
                espacio_boton_deshacer = st.empty()

            get_row_style = JsCode("""
            function(params) {
                if (!params.data) return null;
                if (params.data['Actualizar'] === false) {
                    if (!String(params.data['Valor anterior'] || '').trim()) {
                        return {'backgroundColor': '#FFF0B3', 'color': '#6B4D00', 'fontWeight': '500'};
                    }
                    return {'backgroundColor': '#D99B60', 'color': '#472509', 'fontWeight': '600'};
                }
                var val = params.data['Información extraída'];
                var estado = params.data['Estado'];
                var valStr = val != null ? val.toString().trim() : '';
                if (valStr.toLowerCase() === 'nan' || valStr.toLowerCase() === 'none' || valStr.toLowerCase() === 'null') {
                    valStr = '';
                }
                
                if (valStr !== '' && estado === 'Modificado por usuario') {
                    return {
                        'backgroundColor': '#BAE6FD',
                        'color': '#0369A1',
                        'fontWeight': '600'
                    };
                } else if (valStr !== '') {
                    return {
                        'backgroundColor': '#B7F7C1',
                        'color': '#124D20',
                        'fontWeight': '600'
                    };
                } else {
                    return {
                        'backgroundColor': '#FFF0B3',
                        'color': '#6B4D00',
                        'fontWeight': '500'
                    };
                }
            }
            """)

            dark_header_style = {
                'backgroundColor': '#1E293B',
                'color': '#FFFFFF',
                'fontWeight': 'bold',
                'fontSize': '0.95rem'
            }

            gb = GridOptionsBuilder.from_dataframe(df_base)
            gb.configure_default_column(
                headerStyle=dark_header_style,
                resizable=False,
                suppressMovable=True,
            )

            # Casilla interactiva por campo, independiente de seleccionar una fila.
            casilla_actualizacion = JsCode("""
            class CasillaActualizacion {
                init(params) {
                    this.params = params;
                    this.elemento = document.createElement('input');
                    this.elemento.type = 'checkbox';
                    this.elemento.checked = params.value !== false;
                    this.elemento.setAttribute('aria-label', 'Actualizar campo ' + params.data['Campo']);
                    this.onChange = (evento) => {
                        params.node.setDataValue('Actualizar', !!evento.target.checked);
                        // La vista alterna entre propuesta y valor anterior,
                        // conservando la propuesta intacta para Excel y Ctrl+Z.
                        params.api.redrawRows({rowNodes: [params.node]});
                    };
                    this.elemento.addEventListener('change', this.onChange);
                }
                getGui() { return this.elemento; }
                destroy() { this.elemento.removeEventListener('change', this.onChange); }
            }
            """)
            gb.configure_column(
                "Actualizar", header_name="Aplicar", editable=False, cellRenderer=casilla_actualizacion,
                pinned="left", width=108, minWidth=96, maxWidth=118, flex=0,
                sortable=False, filter=False, resizable=False,
                headerTooltip="Marcado: reemplazar valor; desmarcado: conservar el valor anterior.",
                headerClass='ag-header-centered',
                cellStyle={'textAlign': 'center'}
            )
            gb.configure_column(
                "Campo",
                editable=False,
                flex=2,
                headerClass='ag-header-left',
                cellStyle={'fontWeight': 'bold', 'fontSize': '0.9rem'}
            )
            gb.configure_column(
                "Información extraída",
                editable=JsCode("function(params) { return params.data && params.data['Actualizar'] !== false && params.data['Campo'] !== 'Referencias'; }"),
                valueFormatter=JsCode("""function(params) {
                    if (params.data && params.data['Actualizar'] === false) {
                        const anterior = params.data['Valor anterior'];
                        return anterior == null || String(anterior).trim() === ''
                            ? 'Sin valor anterior' : String(anterior);
                    }
                    return params.value == null ? '' : String(params.value);
                }"""),
                flex=5,
                headerClass='ag-header-left',
                cellEditor='agLargeTextCellEditor',
                cellEditorPopup=True,
                cellEditorParams={'maxLength': 32767, 'rows': 18, 'cols': 85},
                autoHeight=True
            )
            gb.configure_column("Valor anterior", hide=True, editable=False,
                                sortable=False, filter=False, suppressColumnsToolPanel=True)
            gb.configure_column(
                "Estado",
                editable=False,
                valueFormatter=JsCode("""function(params) {
                    if (params.data && params.data['Actualizar'] === false)
                        return String(params.data['Valor anterior'] || '').trim()
                            ? 'Valor anterior' : 'Sin valor anterior';
                    return params.value == null ? '' : String(params.value);
                }"""),
                flex=2,
                headerClass='ag-header-centered',
                cellStyle={'textAlign': 'center', 'fontSize': '0.85rem'}
            )
            # Columna oculta: Ctrl+Z envía una sola orden al servidor. La
            # operación anterior también se refleja inmediatamente en el
            # propio grid; no es necesario forzar un segundo rerun de Streamlit.
            gb.configure_column(ht.COLUMNA_COMANDO, hide=True, editable=False,
                                sortable=False, filter=False, suppressColumnsToolPanel=True)

            grid_options = gb.build()
            grid_options['getRowStyle'] = get_row_style
            grid_options['suppressMovableColumns'] = True
            # El historial nativo se pierde cuando Streamlit reconstruye la
            # tabla tras cualquier edición. Conservamos la última operación
            # de la sesión únicamente como contexto EFÍMERO del grid.
            grid_options['undoRedoCellEditing'] = False
            grid_options['context'] = {
                # En un remontaje, recuperar hasta veinte acciones del servidor.
                # Las nuevas ediciones se registran también en el navegador,
                # para no depender de la latencia del rerun para Ctrl+Z.
                # st_aggrid.shared.walk_gridOptions no admite listas de listas
                # dentro de gridOptions: provocaba TypeError (índice dict en list)
                # al existir al menos un historial de ediciones. Enviar JSON
                # como cadena y deserializarlo únicamente dentro del navegador.
                'pilaDeshacerJson': json.dumps(
                    historial_ediciones.get('deshacer', [])[-ht.LIMITE:],
                    ensure_ascii=False,
                ),
                'ultimaEdicionReversible': (
                    historial_ediciones.get('deshacer', [])[-1]
                    if historial_ediciones.get('deshacer') else []
                ),
            }
            # Una ÚNICA ruta de deshacer: tanto el evento del documento como
            # el de la celda llaman al mismo manejador, sin contextos obsoletos
            # ni doble envío. La pila local incluye el primer Suprimir incluso
            # antes de que termine el rerun que actualiza Excel en Python.
            grid_options['onGridReady'] = JsCode("""
            function(params) {
                const api = params.api;
                const opciones = params.context || {};
                let desdeServidor = [];
                try {
                    const datos = JSON.parse(opciones.pilaDeshacerJson || '[]');
                    if (Array.isArray(datos)) {
                        desdeServidor = datos.filter(Array.isArray)
                            .map(grupo => grupo.map(c => ({...c})));
                    }
                } catch (_) {
                    // Un contexto obsoleto no debe impedir abrir la tabla.
                    desdeServidor = [];
                }
                const estado = api.__fichasUndoState || { pila: desdeServidor.slice(-20) };
                api.__fichasUndoState = estado;

                const deshacer = (evento) => {
                    if (!evento || evento.__fichasDeshacerAtendido || evento.repeat ||
                        !(evento.ctrlKey || evento.metaKey) || evento.altKey || evento.shiftKey ||
                        String(evento.key || '').toLowerCase() !== 'z') return;
                    // Dentro del editor largo o un input sigue funcionando
                    // el deshacer de texto nativo, sin alterar el historial.
                    if (api.getEditingCells && api.getEditingCells().length) return;
                    if (evento.target && evento.target.closest &&
                        evento.target.closest('input, textarea, [contenteditable="true"]')) return;
                    const grupo = estado.pila[estado.pila.length - 1];
                    if (!Array.isArray(grupo) || !grupo.length) return;
                    const anteriores = new Map(grupo.map(c => [String(c.campo), c]));
                    const nodos = [];
                    let primerNodo = null;
                    api.forEachNode(node => {
                        if (!primerNodo) primerNodo = node;
                        const cambio = anteriores.get(String(node.data['Campo']));
                        if (!cambio) return;
                        const valor = cambio.antes == null ? '' : String(cambio.antes);
                        node.data['Información extraída'] = valor;
                        const campo = String(cambio.campo);
                        node.data['Estado'] = cambio.antes_editado && valor ? 'Modificado por usuario' :
                            ['Nombre LIS', 'Laboratorio de procesamiento', 'Actualizado por',
                             'Autorizado por'].includes(campo) ?
                                (valor ? 'Ingresado por usuario' : 'Pendiente de ingreso') :
                            campo === 'Fecha actualización' ?
                                (valor ? 'Asignado por sistema' : 'Pendiente de asignación') :
                                (valor ? 'Encontrado' : 'No encontrado');
                        nodos.push(node);
                    });
                    if (!nodos.length || !primerNodo) return;
                    evento.__fichasDeshacerAtendido = true;
                    evento.preventDefault();
                    evento.stopPropagation();
                    if (evento.stopImmediatePropagation) evento.stopImmediatePropagation();
                    estado.pila.pop();
                    api.refreshCells({rowNodes: nodos, force: true});
                    // Se cambia solo la marca, no el valor restaurado:
                    // una única comunicación al backend, sin repetir eventos.
                    primerNodo.setDataValue('__historial_accion',
                        'deshacer:' + Date.now() + ':' + Math.random());
                };
                api.__fichasUndoTecla = deshacer;
                // Las teclas se escuchan en el iframe de la tabla. Al
                // reconstruir el grid, soltar la escucha anterior.
                if (window.__fichasCtrlZHandler) {
                    document.removeEventListener('keydown', window.__fichasCtrlZHandler, true);
                }
                window.__fichasCtrlZHandler = deshacer;
                document.addEventListener('keydown', deshacer, true);
            }
            """)
            grid_options['onCellKeyDown'] = JsCode("""
            function(params) {
                // Respaldo del evento nativo de AG Grid cuando el navegador
                // entrega la tecla a una celda y no al documento del iframe.
                // Ambos caminos ejecutan exactamente la misma función.
                const accion = params.api.__fichasUndoTecla;
                if (typeof accion === 'function') accion(params.event);
            }
            """)

            grid_options['onCellValueChanged'] = JsCode("""
            function(params) {
                if (params.column && params.column.colDef && params.column.colDef.field === 'Información extraída') {
                    var newVal = params.newValue != null ? params.newValue.toString().trim() : '';
                    if (newVal.toLowerCase() === 'nan' || newVal.toLowerCase() === 'none' || newVal.toLowerCase() === 'null') {
                        newVal = '';
                    }
                    // La pila en AG Grid registra la edición EN EL INSTANTE
                    // en que cambia la celda; no espera la respuesta de
                    // Streamlit (y conserva el primer borrado con Suprimir).
                    const antes = params.oldValue == null ? '' : String(params.oldValue).trim();
                    if (antes !== newVal) {
                        const cambio = {
                            campo: String(params.data['Campo']),
                            antes: antes,
                            despues: newVal,
                            antes_editado: params.data['Estado'] === 'Modificado por usuario',
                            despues_editado: !!newVal,
                        };
                        const estadoUndo = params.api.__fichasUndoState ||
                            (params.api.__fichasUndoState = {pila: []});
                        estadoUndo.pila.push([cambio]);
                        if (estadoUndo.pila.length > 20) estadoUndo.pila.shift();
                        if (params.context) params.context.ultimaEdicionReversible = [cambio];
                    }

                    if (newVal === '') {
                        var campo = params.data['Campo'];
                        if (['Nombre LIS', 'Laboratorio de procesamiento', 'Actualizado por', 'Autorizado por'].includes(campo)) {
                            params.data['Estado'] = 'Pendiente de ingreso';
                        } else if (campo === 'Fecha actualización') {
                            params.data['Estado'] = 'Pendiente de asignación';
                        } else {
                            params.data['Estado'] = 'No encontrado';
                        }
                    } else {
                        params.data['Estado'] = 'Modificado por usuario';
                    }
                    // Actualizar la apariencia sin emitir otro evento de
                    // valor modificado: evita reruns duplicados de Streamlit.
                    params.api.refreshCells({rowNodes: [params.node], force: true});
                }
            }
            """)

            filas_cant = len(df_base)
            altura_tabla = filas_cant * 42 + 55

            grid_response = AgGrid(
                df_base,
                gridOptions=grid_options,
                allow_unsafe_jscode=True,
                update_mode=GridUpdateMode.VALUE_CHANGED,  # Evitar reruns por el refresco del modelo.
                reload_data=False,
                fit_columns_on_grid_load=True,
                theme="alpine",
                height=altura_tabla,
                custom_css={
                    ".ag-header-centered .ag-header-cell-label": {"justify-content": "center", "width": "100%"},
                    ".ag-header-centered .ag-header-cell-text": {"text-align": "center", "width": "100%"},
                    ".ag-header-left .ag-header-cell-label": {"justify-content": "flex-start", "width": "100%"},
                    ".ag-header-left .ag-header-cell-text": {"text-align": "left", "width": "100%"},
                },
                key=f"aggrid_data_table_{st.session_state['aggrid_key_counter']}"
            )

            edited_df = pd.DataFrame(grid_response['data'])
            filas_de_grid = edited_df.to_dict('records')
            # El grid puede retener la última marca oculta al actualizarse
            # sin volver a montar el componente. Una misma marca se consume
            # UNA sola vez para evitar deshacer en cascada durante otros reruns.
            marca_historial = ht.marca_desde_filas(filas_de_grid)
            if (marca_historial
                    and marca_historial != st.session_state.get("marca_historial_consumida")):
                st.session_state["marca_historial_consumida"] = marca_historial
                comando_historial = ht.comando_desde_filas(filas_de_grid)
            else:
                comando_historial = None
            if comando_historial == "deshacer":
                # El navegador ya restauró visualmente las celdas. Sincronizar
                # el historial Python y el Excel en este MISMO rerun, evitando
                # st.rerun() (causaba dos reconstrucciones por atajo).
                acciones_pendientes = historial_ediciones.get("deshacer") or []
                campos_anteriores = {
                    cambio["campo"] for cambio in (acciones_pendientes[-1]
                                                    if acciones_pendientes else [])
                }
                if _aplicar_historial_tabla("deshacer", actualizar_grid=False):
                    editados_recuperados = st.session_state.get("user_edited_fields", set())
                    for indice, fila_rec in edited_df.iterrows():
                        campo_rec = str(fila_rec["Campo"])
                        if campo_rec not in campos_anteriores:
                            continue
                        valor_rec = st.session_state["data_extracted"].get(campo_rec, "")
                        edited_df.at[indice, "Información extraída"] = valor_rec
                        if campo_rec in editados_recuperados and valor_rec:
                            estado_rec = "Modificado por usuario"
                        elif campo_rec in ("Nombre LIS", "Laboratorio de procesamiento",
                                           "Actualizado por", "Autorizado por"):
                            estado_rec = "Ingresado por usuario" if valor_rec else "Pendiente de ingreso"
                        elif campo_rec == "Fecha actualización":
                            estado_rec = "Asignado por sistema" if valor_rec else "Pendiente de asignación"
                        else:
                            estado_rec = "Encontrado" if valor_rec else "No encontrado"
                        edited_df.at[indice, "Estado"] = estado_rec
            else:
                ht.registrar_cambios(
                    historial_ediciones,
                    st.session_state.get("data_extracted", {}),
                    st.session_state.get("user_edited_fields", set()),
                    filas_de_grid,
                )
            seleccion_actual = st.session_state.setdefault("seleccion_campos_excel", {})
            for _, fila_check in edited_df.iterrows():
                campo_check = str(fila_check["Campo"])
                seleccion_actual[campo_check] = campo_marcado_para_actualizar(
                    fila_check.get("Actualizar", seleccion_actual.get(campo_check, True)))

            if "data_extracted" in st.session_state:
                for idx, row in edited_df.iterrows():
                    field = row["Campo"]
                    val_raw = row["Información extraída"]
                    if pd.isna(val_raw) or val_raw is None:
                        val_editado = ""
                    else:
                        val_editado = str(val_raw).strip()
                        if val_editado.casefold() in ["nan", "none", "null"]:
                            val_editado = ""
                    st.session_state["data_extracted"][field] = val_editado
                    
                    if row["Estado"] == "Modificado por usuario" and val_editado:
                        st.session_state.setdefault("user_edited_fields", set()).add(field)
                    else:
                        st.session_state.get("user_edited_fields", set()).discard(field)

            # Contar lo que realmente devolvió AG Grid, no los valores
            # anteriores al borrado. El placeholder conserva la ubicación
            # original del indicador, sin forzar otro st.rerun.
            espacio_campos_encontrados.metric(
                "Campos encontrados",
                ht.contar_campos_encontrados({
                    campo: valor for campo, valor in st.session_state.get("data_extracted", {}).items()
                    if seleccion_actual.get(campo, True)
                }),
            )

            # Este bloque se renderiza arriba de la tabla gracias al espacio
            # reservado. Su estado ya incluye el PRIMER cambio de esta ejecución.
            with espacio_boton_deshacer.container():
                col_deshacer, _ = st.columns([1, 4])
                with col_deshacer:
                    st.button(
                        "↶ Deshacer", key="btn_deshacer_tabla", use_container_width=True,
                        disabled=not bool(historial_ediciones.get("deshacer")),
                        on_click=_aplicar_historial_tabla, args=("deshacer",),
                    )

            updated_excel_bytes = aplicar_ediciones_a_excel(
                st.session_state["result_bytes"],
                edited_df,
                examen_activo or "",
                campos_forzar=st.session_state.pop("campos_restaurados_tabla", None),
            )
            st.session_state["result_bytes"] = updated_excel_bytes
            # La descarga y la escritura SQLite usan la MISMA versión filtrada.
            # El libro editable sigue guardando todas las propuestas para permitir
            # que el usuario cambie nuevamente sus casillas sin perder información.
            excel_con_confirmaciones = respetar_campos_no_seleccionados(
                updated_excel_bytes,
                st.session_state.get("bytes_plantilla_inicial"),
                examen_activo or "",
                seleccion_actual,
            )

            st.write("")
            st.subheader("Opciones de guardado y avance de cola", anchor=False)

            col_down, col_update, col_next = st.columns(3)

            with col_down:
                nombre_archivo_excel = nombre_ficha_tecnica(
                    area_activa, st.session_state.get("fecha_ultima_extraccion_ficha"))

                try:
                    excel_para_descarga = validar_excel_para_descarga(
                        excel_con_confirmaciones)
                    # No sobrescribir el libro editable: si el usuario vuelve a
                    # marcar un campo, su texto recién extraído debe recuperarse.
                except ValueError as error_excel:
                    excel_para_descarga = None
                    st.error(f"No se pudo validar el Excel. No se descargará un archivo dañado: {error_excel}")
                if excel_para_descarga is not None:
                    st.download_button(
                        "Descargar Excel consolidado (.xlsx)", type="secondary",
                        data=excel_para_descarga,
                        file_name=nombre_archivo_excel,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="download_updated_xlsx",
                        on_click="ignore",  # Descargar no reinicia Streamlit ni mueve el scroll.
                    )

            with col_update:
                discrepancias_sin_revisar = bool(
                    st.session_state.get("discrepancias_pdf")
                    or st.session_state.get("conflictos_pendientes")
                )
                if st.button(
                    "Actualizar en base de datos", type="primary",
                    use_container_width=True, key="btn_trigger_save_template",
                    disabled=discrepancias_sin_revisar,
                    help=("Comprobar discrepancias antes de actualizar."
                          if discrepancias_sin_revisar else None),
                ):
                    bytes_iniciales = st.session_state.get("bytes_plantilla_inicial")
                    st.session_state["excel_para_actualizar"] = validar_excel_para_descarga(
                        excel_con_confirmaciones)
                    st.session_state["comparacion_planilla_detalles"] = comparar_plantillas_excel(
                        bytes_iniciales,
                        st.session_state["excel_para_actualizar"],
                        examen_activo or ""
                    )
                    registro_modificacion = st.session_state.get("revision_plantilla_origen") or gp.estado_recurso(DB_PATH, "area:" + area_activa)
                    st.session_state["revision_plantilla_pendiente"] = registro_modificacion["revision"]
                    st.session_state["metadata_plantilla_pendiente"] = registro_modificacion
                    st.session_state["modal_confirmar_planilla_abierto"] = True
                    # El clic de Streamlit ya provocó el rerun necesario;
                    # otro rerun aquí desplazaba la página sin aportar nada.

            with col_next:
                if modo_conjunto:
                    st.info("Análisis conjunto finalizado. Descarga o actualiza la ficha única.")
                idx_actual = st.session_state.get("cola_pdf_indice", 0)
                total_pdfs = len(pdfs_ordenados_rev) if 'pdfs_ordenados_rev' in locals() and pdfs_ordenados_rev else 0
                
                if not modo_conjunto and idx_actual < total_pdfs - 1:
                    with st.container(key="contenedor_continuar_examen"):
                        avanzar_cola = st.button(
                            f"Continuar al siguiente examen ({idx_actual + 1}/{total_pdfs})",
                            type="primary", use_container_width=True,
                            key="btn_continuar_siguiente_examen", help=None,
                        )
                    if avanzar_cola:
                        # Conservar únicamente los campos marcados y no perder
                        # ninguna ficha ya procesada al pasar de PDF.
                        st.session_state["result_bytes"] = excel_con_confirmaciones
                        st.session_state["cola_pdf_indice"] = idx_actual + 1
                        _limpiar_examen_para_siguiente_pdf()
                        st.rerun()
                elif not modo_conjunto:
                    if total_pdfs > 1:
                        st.success("¡Has completado la cola de exámenes! Descarga o actualiza la documentación.")
                    elif total_pdfs == 1:
                        st.success("¡Has completado el examen! Descarga o actualiza la documentación.")

            if not modo_solo_bd:
                with st.expander("Ver texto bruto extraído"):
                    st.text_area("Texto", st.session_state.get("extracted_text", ""), height=350, key="raw_extracted_text")

    with tab_export:
        st.subheader("Exportar a otro formato", anchor=False)
        st.write("Convierte un archivo Excel a otro formato desde esta pestaña.")
        template_exportacion = st.file_uploader(
            "Cargar Excel para exportar", type=["xlsx"],
            key=f"excel_exportacion_uploader_{st.session_state.get('generacion_cargas_area', 0)}",
        )
        if template_exportacion:
            try:
                imported_wb = load_workbook(io.BytesIO(template_exportacion.getvalue()), read_only=True, data_only=False, rich_text=True)
                imported_sheet_names = imported_wb.sheetnames
                imported_wb.close()
                import_col1, import_col2 = st.columns([1, 2])
                with import_col1:
                    imported_format = st.selectbox(
                        "Formato de exportación",
                        ["XLSX", "CSV", "XLS", "JSON", "XML"],
                        index=None,
                        placeholder="Selecciona formato",
                        key="imported_excel_format",
                    )
                with import_col2:
                    selected_import_sheet = None
                    if imported_format == "CSV":
                        selected_import_sheet = st.selectbox(
                            "Hoja que se exportará a CSV",
                            imported_sheet_names,
                            index=None,
                            placeholder="Selecciona una hoja",
                            key="imported_excel_sheet",
                        )
                    else:
                        st.caption("Se exportarán todas las hojas cuando el formato lo permita.")
                
                if imported_format:
                    imported_bytes, imported_name, imported_mime = export_imported_workbook(
                        template_exportacion.getvalue(), imported_format, selected_import_sheet
                    )
                    st.download_button(
                        f"Descargar Excel importado en {imported_format}", type="primary",
                        data=imported_bytes,
                        file_name=imported_name,
                        mime=imported_mime,
                        key="download_imported_excel",
                        on_click="ignore",
                    )
            except Exception as e:
                st.error(f"No se pudo preparar la exportación del Excel importado: {e}")
        else:
            st.info("Carga un archivo Excel en esta pestaña para habilitar la exportación.")

    with tab_guide:
        guide_tab()

# --- EJECUCIÓN FINAL DEL GESTOR DE DIÁLOGOS EMERGENTES ---
ejecutar_gestor_dialogos()

# Control flotante al extremo inferior derecho (visible solo en un área activa).
mostrar_reinicio_flotante()
