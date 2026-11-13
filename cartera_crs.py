"""Área y emparejamiento de la Cartera del CRS, sin dependencia de Streamlit.

Migración ADITIVA: recupera LABORATORIO desde el Excel que ya existe como
BLOB en documentos_fuente; no reemplaza datos de exámenes ni cuentas.
"""
from __future__ import annotations

import io
import re
import sqlite3
from difflib import SequenceMatcher
import pandas as pd
import fuentes_documentales as fd


def _norm(texto):
    return fd.normalizar(texto)


def _columnas_cartera(df):
    cols = {_norm(c): c for c in df.columns}
    nombre = cols.get("descripcion")
    laboratorio = cols.get("laboratorio")
    return nombre, laboratorio


def migrar_laboratorio_cartera(db_path):
    """Agregar columna y rellenar solo valores vacíos usando el Excel almacenado.

    Transacción única; si falta el documento original no inventa el área.
    Al volver a ejecutar no sustituye valores ya guardados.
    """
    with sqlite3.connect(db_path, timeout=30) as con:
        con.execute("PRAGMA busy_timeout=30000")
        # La mayoría de las ejecuciones solo consultan la estructura: no
        # bloquear el escritor SQLite en cada rerun de tres usuarios.
        cols = {r[1] for r in con.execute("PRAGMA table_info(cartera_examenes)")}
        if "laboratorio" in cols:
            pendientes = con.execute("""SELECT COUNT(*) FROM cartera_examenes
                                      WHERE laboratorio IS NULL OR TRIM(laboratorio)=''""").fetchone()[0]
            if not pendientes:
                return 0
        fila = con.execute("SELECT archivo FROM documentos_fuente WHERE fuente_id='cartera_examenes'").fetchone()
        registros = {}
        if fila and fila[0]:
            try:
                df = pd.read_excel(io.BytesIO(fila[0]))
            except (ValueError, OSError, TypeError, ImportError):
                df = None  # Un documento antiguo corrupto no impide iniciar sesión.
            if df is not None:
                col_nombre, col_lab = _columnas_cartera(df)
                if col_nombre is not None and col_lab is not None:
                    for _, row in df.iterrows():
                        nombre, lab = str(row[col_nombre]).strip(), str(row[col_lab]).strip()
                        if lab and lab.casefold() not in ("nan", "none", "null"):
                            registros.setdefault(_norm(nombre), lab)
        # Sin Excel original, solo se necesita añadir la columna la primera vez.
        if "laboratorio" in cols and not registros:
            return 0
        con.execute("BEGIN IMMEDIATE")
        # Volver a comprobar tras adquirir la exclusión: otra sesión pudo migrar.
        cols = {r[1] for r in con.execute("PRAGMA table_info(cartera_examenes)")}
        if "laboratorio" not in cols:
            con.execute("ALTER TABLE cartera_examenes ADD COLUMN laboratorio TEXT")
        cantidad = 0
        for nombre, in con.execute("""SELECT nombre_examen FROM cartera_examenes
                                    WHERE laboratorio IS NULL OR TRIM(laboratorio)=''""").fetchall():
            lab = registros.get(_norm(nombre))
            if lab:
                con.execute("UPDATE cartera_examenes SET laboratorio=? WHERE nombre_examen=?",
                            (lab, nombre))
                cantidad += 1
        return cantidad


def es_variante_con_nombre_incluido(nombre_examen, candidato):
    """Detecta un nombre ampliado, p. ej. FIBRINOGENO – HOSMET.

    Exige que el nombre original aparezca completo, delimitado por palabras:
    «FIBRINOGENO» no puede activar «FIBRINOGENOS» por accidente.
    """
    objetivo = _norm(nombre_examen)
    variante = _norm(candidato)
    if not objetivo or not variante or objetivo == variante:
        return False
    return bool(re.search(r"(?<!\w)" + re.escape(objetivo) + r"(?!\w)", variante))


def coincidencia_exacta_cartera(db_path, nombre_examen, area):
    """Prioriza el nombre exacto de área compatible; no vincula nombres difusos.

    Excepción acotada: si existe UN único nombre exacto, pero el área no
    coincide o todavía no figura en la base, se acepta ese nombre cuando
    también hay otra prestación ampliada que contiene el nombre íntegro.
    Los candidatos ampliados permanecen disponibles para selección manual.
    Si hay varios nombres exactos en distintas áreas sin forma de resolverlos,
    la vinculación automática sigue bloqueada.
    """
    objetivo = _norm(nombre_examen)
    if not objetivo:
        return None
    with sqlite3.connect(db_path) as con:
        filas = con.execute("SELECT nombre_examen, laboratorio FROM cartera_examenes").fetchall()
    exactas = [(nombre, lab) for nombre, lab in filas if _norm(nombre) == objetivo]
    compatibles = [(nombre, lab) for nombre, lab in exactas
                   if area and lab and fd.area_equivalente(lab) == fd.area_equivalente(area)]
    if len(compatibles) == 1:
        return compatibles[0][0]
    if len(compatibles) > 1 or len(exactas) != 1:
        return None
    if any(es_variante_con_nombre_incluido(nombre_examen, nombre)
           for nombre, _ in filas):
        return exactas[0][0]
    return None


def sincronizar_vinculo_en_sesion(estado, db_path, examen, area, opciones, revision=0):
    """Sincroniza el selector de Streamlit ANTES de crear el widget.

    Los valores del widget pueden quedar vacíos tras actualizar el programa o
    tras cambiar opciones sin que haya cambiado el nombre del examen. En esos
    casos se recupera la coincidencia exacta. Una elección o limpieza hecha
    *explícitamente* por el usuario se conserva hasta cambiar de examen/área.

    ``estado`` puede ser st.session_state o un diccionario para probarlo sin UI.
    No escribe en SQLite ni altera las preferencias de otros usuarios.
    """
    disponibles = set(opciones)
    nuevo_contexto = (str(examen or ""), str(area or ""), int(revision or 0))
    contexto_previo = estado.get("contexto_emparejamiento_cartera")
    eleccion_actual = estado.get("sel_cartera_main")
    manual = bool(estado.get("cartera_seleccion_confirmada_manualmente", False))

    # Se calcula siempre desde la cartera activa: evita que un 'None' obsoleto
    # bloquee un nombre que ahora tiene coincidencia exacta.
    exacto = coincidencia_exacta_cartera(db_path, examen, area)
    if exacto not in disponibles:
        exacto = None

    if contexto_previo != nuevo_contexto:
        # Si solo cambió la versión de la cartera, respetar una vinculación
        # manual aún disponible. Si cambió el examen o el área, empezar limpio.
        mismo_examen_area = (isinstance(contexto_previo, tuple)
                            and contexto_previo[:2] == nuevo_contexto[:2])
        conservar_manual = (manual and mismo_examen_area
                           and (eleccion_actual is None or eleccion_actual in disponibles))
        # Estado heredado de versiones anteriores: no sobrescribir un vínculo
        # diferente que el usuario ya había seleccionado manualmente.
        vinculo_heredado = (isinstance(contexto_previo, str)
                           and contexto_previo == examen
                           and eleccion_actual in disponibles
                           and bool(eleccion_actual)
                           and eleccion_actual != exacto)
        if conservar_manual or vinculo_heredado:
            estado["cartera_seleccion_confirmada_manualmente"] = True
        else:
            estado["sel_cartera_main"] = exacto
            estado["cartera_seleccion_confirmada_manualmente"] = False
        estado["contexto_emparejamiento_cartera"] = nuevo_contexto
    elif eleccion_actual not in disponibles and eleccion_actual is not None:
        # Si se eliminó un nombre de la cartera, nunca dejar un vínculo roto.
        estado["sel_cartera_main"] = None if manual else exacto
    elif not manual and eleccion_actual != exacto:
        # Reparar selectbox vacío/obsoleto aunque el examen activo NO cambie.
        estado["sel_cartera_main"] = exacto

    seleccionado = estado.get("sel_cartera_main")
    estado["cartera_examen_emparejado"] = seleccionado or "Sin vincular"
    return seleccionado



def _prioridad_nombre_examen(nombre_buscar, nombre_candidato):
    """Misma prioridad clínica que Acreditación para ordenar candidatos."""
    return fd.similitud_sugerencia_clinica(nombre_buscar, nombre_candidato)[0]


def similitud_nombre_cartera(nombre_examen, candidato):
    """Porcentaje textual común a CRS y Acreditación, no equivalencia médica.

    Los nombres completos compuestos (p. ej. «AC. LUPICO») no pierden
    prioridad por prefijos largos de la cartera; un parecido solo por letras
    no puede superar automáticamente un nombre íntegro compartido.
    """
    return fd.similitud_sugerencia_clinica(nombre_examen, candidato)[1]


def candidatos_cartera(db_path, nombre_examen, area, limite=8, similitud_minima=0.0):
    """Prioriza el nombre clínico y completa con alternativas por similitud.

    Filtra por similitud ANTES de limitar el resultado; de lo contrario, las
    variantes de un nombre y otros exámenes de la misma área pueden ocupar
    los primeros puestos y dejar sin mostrar coincidencias difusas válidas.
    El área solo desempata porcentajes iguales y nunca altera los valores
    del cálculo textual ni sustituye la selección manual del encargado.
    """
    objetivo = _norm(nombre_examen)
    if not objetivo:
        return []
    area_norm = fd.area_equivalente(area)
    with sqlite3.connect(db_path) as con:
        filas = con.execute("SELECT nombre_examen, laboratorio FROM cartera_examenes").fetchall()
    encontrados = []
    for nombre, lab in filas:
        nom_norm = _norm(nombre)
        prioridad_nombre, ratio = fd.similitud_sugerencia_clinica(objetivo, nom_norm)
        # Aplicar el umbral antes del top-N. Las alternativas difusas siguen
        # visibles solo cuando su semejanza textual es suficiente.
        if ratio < similitud_minima:
            continue
        misma_area = bool(lab and area_norm == fd.area_equivalente(lab))
        encontrados.append((nombre, lab or "", ratio, prioridad_nombre, misma_area))
    encontrados.sort(key=lambda item: (-item[3], -item[2], -int(item[4]), _norm(item[0])))
    return [(nombre, lab, ratio) for nombre, lab, ratio, _, _ in encontrados[:limite]]


def candidatos_variantes_cartera(db_path, nombre_examen, area, limite=5):
    """Alternativas cuyo nombre contiene íntegro el examen exacto.

    No dependen del límite de sugerencias difusas. Esto garantiza que una
    variante como «FIBRINOGENO – HOSMET» siga visible bajo «FIBRINOGENO»,
    aunque existan numerosos nombres parecidos en el catálogo.
    """
    return [c for c in candidatos_cartera(db_path, nombre_examen, area, limite=100000)
            if c[2] >= 0.55 and es_variante_con_nombre_incluido(nombre_examen, c[0])][:limite]
