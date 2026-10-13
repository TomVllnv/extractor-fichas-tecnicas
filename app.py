import io, os, re, tempfile, subprocess, shutil, csv, json, hashlib, xml.etree.ElementTree as ET
from datetime import datetime
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

APP_TITLE = 'Herramienta de Extracción de Datos de Fichas Técnicas'

FIELD_ALIASES = {
    'Código': ['codigo examen','codigo','código examen','código', 'Código del Examen', 'Código BK'],
    'Nombre LIS': ['Derivación','nombre del examen','examen','nombre examen', 'Nombres del Examen'],
    'Laboratorio de procesamiento': ['laboratorio de procesamiento','laboratorio'],
    'Recepción de muestras': ['recepcion de muestras','recepción de muestras','recepcion'],
    'Procesamiento de muestras': ['procesamiento de muestras','procesamiento de muetsras','procesamiento'],
    'Plazo entrega resultados': ['tiempo de respuesta','plazo entrega resultados','plazo de entrega','tat','tiempo procesamiento'],
    'Preparación del paciente': ['preparacion paciente','preparación paciente','preparacion del paciente','preparación del paciente'],
    'Muestra requerida': ['tipo de muestra','muestra requerida','muestra','tipo y volumen de muestra','material biologico','material biológico','volumen requerido'],
    'Estabilidad de la muestra': ['estabilidad muestra','estabilidad de la muestra','estabilidad'],
    'Condiciones de envío al Laboratorio': ['temperatura de transporte','condiciones de envio','condiciones de envío','envio al laboratorio','envío al laboratorio','transporte'],
    'Método de procesamiento': ['metodo utilizado','método utilizado','metodo de procesamiento','método de procesamiento','metodologia','metodología','metodo'],
    'Test incluidos ': ['test incluidos','pruebas incluidas','examenes incluidos','exámenes incluidos'],
    'Valores de referencia': ['valor de referencia','valores de referencia','valores normales','rango de referencia'],
    'Valor crítico': ['valor critico','valor crítico','valores criticos','valores críticos'],
    'Límite de detección ': ['limite de deteccion','límite de detección','limite deteccion','límite deteccion'],
    'Equipamiento utilizado': ['equipamiento utilizado','equipo utilizado','instrumentacion','instrumentación','equipo'],
    'Información clínica': ['utilidad clinica','utilidad clínica','informacion clinica','información clínica','informacion clínica'],
    'Referencias ': ['referencias','bibliografia','bibliografía','fuentes'],
    'Actualizado por:': ['actualizado por'],
    'Fecha actualización:': ['fecha actualizacion','fecha actualización','fecha de actualización'],
    'Autorizado por:': ['autorizado por'],
}

def norm(s):
    import unicodedata
    s = str(s or '').strip().lower()
    s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    s = re.sub(r'\s+', ' ', s)
    return s

def extract_pdf_text(uploaded_bytes, filename):
    try:
        from pypdf import PdfReader
        reader = PdfReader(io.BytesIO(uploaded_bytes))
        pages = [(p.extract_text() or '') for p in reader.pages]
        text = '\n\n'.join(pages)
        if len(re.sub(r'\s','',text)) >= 50:
            return text, 'PDF'
    except Exception:
        pass
    # Optional OCR fallback: PyMuPDF + pytesseract
    try:
        import fitz, pytesseract
        doc = fitz.open(stream=uploaded_bytes, filetype='pdf')
        parts=[]
        for page in doc:
            pix = page.get_pixmap(matrix=fitz.Matrix(2,2), alpha=False)
            from PIL import Image
            img = Image.open(io.BytesIO(pix.tobytes('png')))
            parts.append(pytesseract.image_to_string(img, lang='spa+eng'))
        return '\n\n'.join(parts), 'OCR'
    except Exception as e:
        raise RuntimeError('No se pudo extraer texto del PDF. Instala pypdf; para PDFs escaneados agrega pymupdf, pytesseract y Tesseract OCR.') from e

def clean_value(v):
    v = re.sub(r'[ \t]+', ' ', v).strip()
    return v.strip(' :\t')

def parse_labeled_text(text):
    lines = [x.rstrip() for x in text.replace('\r','').split('\n')]
    found = {}
    current = None
    label_patterns=[]
    for field, aliases in FIELD_ALIASES.items():
        for a in aliases:
            label_patterns.append((field, a))
    # Longest aliases first to avoid partial matches
    label_patterns.sort(key=lambda x: len(x[1]), reverse=True)
    for raw in lines:
        s = raw.strip()
        if not s:
            if current and found.get(current):
                found[current] += '\n'
            continue
        ns = norm(s)
        matched = None
        remainder = ''
        for field, alias in label_patterns:
            na = norm(alias)
            if ns == na:
                matched = field; remainder=''; break
            if ns.startswith(na + ':'):
                matched = field; remainder=s[s.lower().find(':')+1:].strip(); break
        if matched:
            current = matched
            if remainder:
                found[current] = clean_value(remainder)
            else:
                found.setdefault(current, '')
            continue
        # Special generic colon labels, useful for unknown variants
        if ':' in s and len(s.split(':',1)[0]) < 55:
            head, tail = s.split(':',1)
            nh=norm(head)
            for field, alias in label_patterns:
                if nh == norm(alias):
                    current=field; found[current]=clean_value(tail); break
            else:
                if current:
                    found[current] = (found.get(current,'') + ('\n' if found.get(current) else '') + s).strip()
        elif current:
            found[current] = (found.get(current,'') + ('\n' if found.get(current) else '') + s).strip()
    # normalize whitespace but preserve paragraph breaks
    for k,v in list(found.items()):
        found[k] = '\n'.join(clean_value(x) for x in v.split('\n') if x.strip())
    return found

def infer_name_code(text, data):
    code = data.get('Código','')
    if not code:
        m=re.search(r'\b\d{6,}(?:[-.]\d+)?\b', text)
        if m: data['Código']=m.group(0)
    if not data.get('Derivación'):
        for line in text.splitlines():
            s=line.strip()
            if s and not ':' in s and len(s)<100:
                data['Derivación']=s
                break
    return data

def workbook_fields(wb):
    fields=[]
    for ws in wb.worksheets:
        for row in range(1, ws.max_row+1):
            label=ws.cell(row,1).value
            if label and str(label).strip() not in fields:
                fields.append(str(label).strip())
    return fields

def find_sheet(wb, data):
    code=norm(data.get('Código',''))
    name=norm(data.get('Derivación',''))
    # Exact code first
    if code:
        for ws in wb.worksheets:
            if norm(ws['B1'].value)==code:
                return ws, 1.0, 'Código exacto'
            for r in range(1, min(ws.max_row,5)+1):
                if norm(ws.cell(r,2).value)==code:
                    return ws, 1.0, 'Código exacto'
    # Exact/containment by name against sheet name and B2
    best=None
    for ws in wb.worksheets:
        candidates=[norm(ws.title), norm(ws['B2'].value or '')]
        score=0
        for c in candidates:
            if name and (name==c): score=max(score,0.98)
            elif name and name in c or c and c in name: score=max(score,0.85)
        if score and (best is None or score>best[1]): best=(ws,score,'Nombre')
    return best if best else (None,0,'No encontrado')

def apply_to_sheet(ws, data):
    # Build label->row mapping dynamically, tolerating shifted/duplicate rows.
    rows={}
    for r in range(1, ws.max_row+1):
        label=str(ws.cell(r,1).value or '').strip()
        if label:
            rows[norm(label)]=r
    updates=[]
    for field,value in data.items():
        if not value: continue
        target=field
        r=rows.get(norm(target))
        if r is None:
            # match aliases to existing workbook labels
            for existing, rr in rows.items():
                if existing==norm(target) or norm(target) in existing or existing in norm(target):
                    r=rr; break
        if r:
            old=ws.cell(r,2).value
            ws.cell(r,2).value=value
            ws.cell(r,2).alignment=Alignment(wrap_text=True, vertical='top')
            updates.append((field, old, value, r))
    return updates

def style_review_sheet(ws):
    for cell in ws[1]:
        cell.font=Font(bold=True)
        cell.fill=PatternFill('solid', fgColor='D9EAF7')
    for row in ws.iter_rows():
        for c in row:
            c.alignment=Alignment(vertical='top', wrap_text=True)
    ws.freeze_panes='A2'
    widths={'A':28,'B':24,'C':70,'D':18}
    for col,w in widths.items(): ws.column_dimensions[col].width=w

def process(uploaded_bytes, filename, template_bytes, centro=''):
    text, method = extract_pdf_text(uploaded_bytes, filename)
    data = infer_name_code(text, parse_labeled_text(text))
    wb=load_workbook(io.BytesIO(template_bytes))
    ws,score,reason=find_sheet(wb,data)
    status='Coincidencia encontrada' if ws else 'Sin coincidencia'
    updates=[]
    if ws:
        updates=apply_to_sheet(ws,data)
        target=ws.title
    else:
        target='Sin modificación'
    # Always add a review/import log sheet so nothing is lost.
    if 'Nueva Derivación' in wb.sheetnames:
        del wb['Nueva Derivación']
    log=wb.create_sheet('Nueva Derivación')
    log.append(['Campo','Valor extraído','Estado','Documento','Centro de salud'])
    updated_fields={u[0] for u in updates}
    for field in FIELD_ALIASES:
        val=data.get(field,'')
        if not val:
            state='No encontrado'
        elif field in updated_fields:
            state='Incorporado'
        else:
            state='Encontrado / revisar'
        log.append([field,val,state,filename,centro])
    style_review_sheet(log)
    log['G1']='Resultado'; log['H1']='Método'; log['G2']=status; log['H2']=method
    # Save
    out=io.BytesIO(); wb.save(out); out.seek(0)
    return out.getvalue(), data, target, score, reason, method, text


# --- Acceso básico (adecuado solo para prototipo/pruebas) ---
DEFAULT_USERNAME = "CRSCODUOC"
DEFAULT_PASSWORD = "CRSCODUOC"

def check_login():
    """Muestra la pantalla de acceso y devuelve True cuando la sesión está autenticada."""
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False

    if st.session_state["authenticated"]:
        return True

    st.title("Acceso a la herramienta")
    st.caption("Ingresa tus credenciales para continuar.")
    with st.form("login_form", clear_on_submit=False):
        username = st.text_input("Usuario")
        password = st.text_input("Contraseña", type="password")
        submitted = st.form_submit_button("Iniciar sesión", type="primary")
    if submitted:
        # En producción, configurar credenciales privadas mediante st.secrets y usar
        # un sistema de autenticación robusto. Estas credenciales son de demostración.
        try:
            expected_user = st.secrets.get("APP_USERNAME", DEFAULT_USERNAME)
            expected_password = st.secrets.get("APP_PASSWORD", DEFAULT_PASSWORD)
        except Exception:
            expected_user, expected_password = DEFAULT_USERNAME, DEFAULT_PASSWORD
        if (hashlib.sha256(username.encode("utf-8")).digest() ==
                hashlib.sha256(expected_user.encode("utf-8")).digest() and
                hashlib.sha256(password.encode("utf-8")).digest() ==
                hashlib.sha256(expected_password.encode("utf-8")).digest()):
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Usuario o contraseña incorrectos.")
    st.info("Acceso de prueba: usuario y contraseña iniciales son CRSCODUOC.")
    return False

def build_export_bytes(fmt, data, filename, centro, target, score, reason, method):
    """Crea una exportación estructurada de los campos extraídos."""
    rows = []
    for field in FIELD_ALIASES:
        value = data.get(field, "")
        rows.append({
            "Campo": field.strip(),
            "Información extraída": value,
            "Estado": "Encontrado" if value else "No encontrado",
            "Documento": filename,
            "Centro de salud": centro,
            "Ficha identificada": target,
            "Método de extracción": method,
        })

    if fmt == "CSV":
        buffer = io.StringIO(newline="")
        writer = csv.DictWriter(buffer, fieldnames=list(rows[0].keys()), delimiter=",")
        writer.writeheader()
        writer.writerows(rows)
        return buffer.getvalue().encode("utf-8-sig"), "resultado_extraccion.csv", "text/csv; charset=utf-8"

    if fmt == "JSON":
        payload = {
            "documento": filename,
            "centro_de_salud": centro,
            "ficha_identificada": target,
            "coincidencia": {"porcentaje": round(score * 100, 1), "motivo": reason},
            "metodo_extraccion": method,
            "campos": rows,
        }
        return json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8"), "resultado_extraccion.json", "application/json"

    if fmt == "XML":
        root = ET.Element("resultado_extraccion")
        for key, value in [
            ("documento", filename), ("centro_de_salud", centro),
            ("ficha_identificada", target), ("metodo_extraccion", method),
            ("coincidencia_porcentaje", str(round(score * 100, 1))),
            ("motivo_coincidencia", reason),
        ]:
            node = ET.SubElement(root, key)
            node.text = str(value or "")
        fields_node = ET.SubElement(root, "campos")
        for row in rows:
            field_node = ET.SubElement(fields_node, "campo")
            for key, value in row.items():
                # XML tags cannot contain spaces or accented characters.
                tag = {
                    "Campo": "nombre", "Información extraída": "informacion_extraida",
                    "Estado": "estado", "Documento": "documento",
                    "Centro de salud": "centro_de_salud",
                    "Ficha identificada": "ficha_identificada",
                    "Método de extracción": "metodo_extraccion",
                }[key]
                child = ET.SubElement(field_node, tag)
                child.text = str(value or "")
        xml_bytes = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        return xml_bytes, "resultado_extraccion.xml", "application/xml"

    if fmt == "XLS":
        try:
            import xlwt
        except ImportError as exc:
            raise RuntimeError("Falta la dependencia xlwt. Instálala con: pip install xlwt") from exc
        book = xlwt.Workbook(encoding="utf-8")
        sheet = book.add_sheet("Extraccion")
        header_style = xlwt.easyxf("font: bold on; pattern: pattern solid, fore_colour ice_blue;")
        headers = list(rows[0].keys())
        for col, header in enumerate(headers):
            sheet.write(0, col, header, header_style)
            sheet.col(col).width = 5000 if col != 1 else 12000
        for row_idx, row in enumerate(rows, start=1):
            for col_idx, header in enumerate(headers):
                sheet.write(row_idx, col_idx, str(row.get(header, ""))[:32000])
        buffer = io.BytesIO()
        book.save(buffer)
        return buffer.getvalue(), "resultado_extraccion.xls", "application/vnd.ms-excel"

    raise ValueError("Formato de exportación no compatible.")


def export_imported_workbook(template_bytes, fmt, selected_sheet=None):
    """Exporta los datos del Excel que el usuario importó, independientemente del PDF."""
    wb = load_workbook(io.BytesIO(template_bytes), data_only=False)
    sheet_names = wb.sheetnames

    def sheet_to_rows(ws):
        return [[cell.value for cell in row] for row in ws.iter_rows()]

    if fmt == "XLSX":
        # Devuelve el archivo original exactamente como fue cargado.
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
            # Excel XLS limita el nombre de la hoja a 31 caracteres y no admite ciertos símbolos.
            base_name = re.sub(r'[\\/*?:\[\]]', '_', ws.title)[:31] or "Hoja"
            name = base_name
            suffix = 1
            while name.lower() in used_names:
                ending = f"_{suffix}"
                name = base_name[:31-len(ending)] + ending
                suffix += 1
            used_names.add(name.lower())
            out_ws = out_book.add_sheet(name)
            for row_idx, row in enumerate(ws.iter_rows(), start=0):
                for col_idx, cell in enumerate(row):
                    value = cell.value
                    if value is None:
                        continue
                    # XLS tiene un límite de longitud de celda cercano a 32 mil caracteres.
                    if isinstance(value, (str, int, float, bool)):
                        out_ws.write(row_idx, col_idx, str(value)[:32000] if isinstance(value, str) else value)
                    else:
                        out_ws.write(row_idx, col_idx, str(value)[:32000])
        buffer = io.BytesIO()
        out_book.save(buffer)
        return buffer.getvalue(), "excel_importado.xls", "application/vnd.ms-excel"

    if fmt == "JSON":
        payload = {"archivo": "excel_importado.xlsx", "hojas": []}
        for ws in wb.worksheets:
            payload["hojas"].append({
                "nombre": ws.title,
                "filas": sheet_to_rows(ws),
            })
        return json.dumps(payload, ensure_ascii=False, indent=2, default=str).encode("utf-8"), "excel_importado.json", "application/json"

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
    st.subheader("Guía de desarrollo seguro")
    st.write(
        "Esta guía reúne prácticas recomendadas para desarrollar, probar y mantener la "
        "herramienta de extracción de fichas técnicas. La presencia de esta guía no "
        "certifica que la aplicación sea segura ni sustituye una evaluación técnica."
    )
    with st.expander("1. Protección de documentos y datos", expanded=True):
        st.markdown(
            "- Utilizar documentos ficticios o anonimizados durante las pruebas.\n"
            "- No cargar datos personales o clínicos identificables en un prototipo público.\n"
            "- Compartir resultados solo con personas autorizadas.\n"
            "- Evitar guardar documentos o resultados más tiempo del necesario."
        )
    with st.expander("2. Usuarios, contraseñas y configuración"):
        st.markdown(
            "- Las credenciales iniciales `CRSCODUOC` / `CRSCODUOC` son solo para pruebas.\n"
            "- No reutilizar estas credenciales en sistemas reales.\n"
            "- Para una versión de producción, usar autenticación robusta, contraseñas individuales, "
            "control de intentos y gestión de sesiones.\n"
            "- Guardar secretos en la configuración privada de Streamlit (`st.secrets`), nunca en GitHub."
        )
    with st.expander("3. Manejo seguro de archivos"):
        st.markdown(
            "- Aceptar solo los tipos de archivo necesarios y validar su contenido.\n"
            "- Establecer límites de tamaño y manejar errores sin mostrar información interna.\n"
            "- Mantener las bibliotecas actualizadas y revisar su origen.\n"
            "- No ejecutar archivos cargados ni tratar su contenido como código."
        )
    with st.expander("4. Calidad y revisión de datos"):
        st.markdown(
            "- Marcar claramente los campos no encontrados y los que requieren revisión.\n"
            "- No completar datos clínicos por suposición.\n"
            "- Comparar los resultados con el documento original antes de utilizarlos.\n"
            "- Probar con documentos de distintos formatos y con PDFs escaneados."
        )
    with st.expander("5. GitHub y publicación"):
        st.markdown(
            "- Revisar los cambios antes de hacer commit.\n"
            "- No subir documentos clínicos, contraseñas, tokens ni claves al repositorio.\n"
            "- Restringir el acceso al repositorio y a la aplicación según las necesidades.\n"
            "- Revisar la configuración de Streamlit Cloud y las dependencias antes de publicar."
        )
    st.warning(
        "Importante: el login incluido en este prototipo es una barrera básica de acceso, "
        "no un sistema de autenticación de nivel productivo. No usar con información clínica "
        "sensible hasta implementar y validar controles de seguridad apropiados."
    )


st.set_page_config(page_title=APP_TITLE, page_icon="🧪", layout="wide")

if not check_login():
    st.stop()

st.title(APP_TITLE)
st.caption("Prototipo: documento de derivación → extracción → revisión → exportación")

with st.sidebar:
    col_izq, col_logo, col_der = st.columns([0.5, 4, 0.5])
    with col_logo:
        st.image("CRSCO.png", width=240)
    st.divider()
    st.header("📁 Archivos")
    template = st.file_uploader("Excel de fichas técnicas", type=["xlsx"])
    centro = st.text_input("Centro de salud")
    st.info("La herramienta no genera un PDF. Usa el PDF como fuente y actualiza una copia del Excel.")
    if st.button("Cerrar sesión"):
        st.session_state["authenticated"] = False
        st.session_state.pop("result", None)
        st.rerun()

tab_extractor, tab_guide = st.tabs(["Extractor de fichas", "Guía de desarrollo seguro"])

with tab_extractor:
    st.subheader("Exportar el Excel importado")
    st.write("Esta función exporta la información que ya existe en el Excel, aunque todavía no hayas cargado un PDF. XLSX conserva el libro original; CSV exporta una hoja seleccionada; XLS, JSON y XML incluyen todas las hojas del libro.")
    if template:
        try:
            imported_wb = load_workbook(io.BytesIO(template.getvalue()), read_only=True, data_only=False)
            imported_sheet_names = imported_wb.sheetnames
            imported_wb.close()
            import_col1, import_col2 = st.columns([1, 2])
            with import_col1:
                imported_format = st.selectbox(
                    "Formato del Excel importado",
                    ["XLSX", "CSV", "XLS", "JSON", "XML"],
                    key="imported_excel_format",
                )
            with import_col2:
                selected_import_sheet = None
                if imported_format == "CSV":
                    selected_import_sheet = st.selectbox(
                        "Hoja que se exportará a CSV",
                        imported_sheet_names,
                        key="imported_excel_sheet",
                    )
                else:
                    st.caption("Se exportarán todas las hojas cuando el formato lo permita.")
            imported_bytes, imported_name, imported_mime = export_imported_workbook(
                template.getvalue(), imported_format, selected_import_sheet
            )
            st.download_button(
                f"⬇️ Descargar Excel importado en {imported_format}",
                data=imported_bytes,
                file_name=imported_name,
                mime=imported_mime,
                key="download_imported_excel",
            )
        except Exception as e:
            st.error(f"No se pudo preparar la exportación del Excel importado: {e}")
    else:
        st.info("Carga primero un Excel de fichas técnicas en la barra lateral para habilitar esta exportación.")

    st.divider()
    st.subheader("Extraer información desde un PDF")
    pdf = st.file_uploader(
        "1. Carga la derivación",
        type=["pdf"],
        help="PDF con la información del examen.",
        key="pdf_uploader",
    )

    if pdf and template:
        if st.button("🔎 Analizar documento", type="primary"):
            try:
                result = process(pdf.getvalue(), pdf.name, template.getvalue(), centro)
                st.session_state["result"] = result
                st.session_state["filename"] = pdf.name
                st.session_state["centro"] = centro
                st.session_state["target_template_name"] = template.name
            except Exception as e:
                st.error(f"No se pudo analizar el documento: {e}")
    elif not template or not pdf:
        st.info("Carga el Excel de fichas técnicas y una derivación PDF para comenzar.")

    if "result" in st.session_state:
        out, data, target, score, reason, method, extracted_text = st.session_state["result"]
        filename = st.session_state.get("filename", "documento.pdf")
        centro_result = st.session_state.get("centro", "")
        st.subheader("2. Resultado de extracción")
        c1, c2, c3 = st.columns(3)
        c1.metric("Campos encontrados", sum(bool(v) for v in data.values()))
        c2.metric("Método", method)
        c3.metric("Ficha identificada", target[:25] if target else "No encontrada")
        st.write("**Coincidencia:**", reason, f"({score:.0%})" if score else "")

        rows = []
        for field in FIELD_ALIASES:
            value = data.get(field, "")
            rows.append({
                "Campo": field.strip(),
                "Información extraída": value,
                "Estado": "✓ Encontrado" if value else "⚠ No encontrado",
            })
        df = pd.DataFrame(rows)

        def highlight_found(row):
            if row["Estado"] == "✓ Encontrado":
                return ["background-color: #B7F7C1; color: #124D20; font-weight: bold"] * len(row)
            return ["background-color: #FFF0B3; color: #6B4D00"] * len(row)

        st.caption("Verde: se encontró información. Amarillo: no se encontró información; revisa el documento original.")
        st.dataframe(
            df.style.apply(highlight_found, axis=1),
            use_container_width=True,
            hide_index=True,
        )

        st.subheader("3. Exportar resultados")
        st.write("Selecciona el formato que necesitas. XLSX descarga una copia del Excel actualizado; los demás formatos descargan una tabla con los campos extraídos y su estado.")
        export_format = st.selectbox(
            "Formato de exportación",
            ["XLSX", "CSV", "XLS", "JSON", "XML"],
            index=0,
        )
        if export_format == "XLSX":
            st.download_button(
                "⬇️ Descargar Excel actualizado (.xlsx)",
                data=out,
                file_name="Fichas_tecnicas_actualizadas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_xlsx",
            )
        else:
            try:
                export_bytes, export_name, export_mime = build_export_bytes(
                    export_format, data, filename, centro_result, target, score, reason, method
                )
                st.download_button(
                    f"⬇️ Descargar resultado (.{export_format.lower()})",
                    data=export_bytes,
                    file_name=export_name,
                    mime=export_mime,
                    key=f"download_{export_format.lower()}",
                )
            except Exception as e:
                st.error(f"No se pudo generar el archivo {export_format}: {e}")

        with st.expander("Ver texto bruto extraído"):
            st.text_area("Texto", extracted_text, height=350)

with tab_guide:
    guide_tab()
