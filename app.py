import io, os, re, tempfile, subprocess, shutil
from datetime import datetime
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

APP_TITLE = 'Extractor de Fichas Técnicas'

FIELD_ALIASES = {
    'Código': ['codigo examen','codigo','código examen','código'],
    'Nombre LIS': ['nombre lis','nombre del examen','examen','nombre examen'],
    'Laboratorio de procesamiento': ['laboratorio de procesamiento','laboratorio'],
    'Recepción de muestras:': ['recepcion de muestras','recepción de muestras','recepcion'],
    'Procesamiento de muestras:': ['procesamiento de muestras','procesamiento de muetsras','procesamiento'],
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
            return text, 'Texto PDF'
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
    if not data.get('Nombre LIS'):
        for line in text.splitlines():
            s=line.strip()
            if s and not ':' in s and len(s)<100:
                data['Nombre LIS']=s
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
    name=norm(data.get('Nombre LIS',''))
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
        target='No se modificó ninguna ficha existente'
    # Always add a review/import log sheet so nothing is lost.
    if 'LOG IMPORTACION' in wb.sheetnames:
        del wb['LOG IMPORTACION']
    log=wb.create_sheet('LOG IMPORTACION')
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

st.set_page_config(page_title=APP_TITLE, page_icon='🧪', layout='wide')
st.title('🧪 Extractor de Fichas Técnicas')
st.caption('Prototipo: documento de derivación → extracción → revisión → Excel')

with st.sidebar:
    st.header('Archivos')
    template=st.file_uploader('Excel de fichas técnicas', type=['xlsx'])
    centro=st.text_input('Centro de salud (opcional)')
    st.info('La herramienta no genera un PDF. Usa el PDF/documento como fuente y actualiza una copia del Excel.')

pdf=st.file_uploader('1. Carga la derivación', type=['pdf'], help='PDF con la información del examen.')

if pdf and template:
    if st.button('🔎 Analizar documento', type='primary'):
        try:
            result=process(pdf.getvalue(),pdf.name,template.getvalue(),centro)
            out,data,target,score,reason,method,text=result
            st.session_state['result']=result
            st.session_state['filename']=pdf.name
        except Exception as e:
            st.error(str(e))

if 'result' in st.session_state:
    out,data,target,score,reason,method,text=st.session_state['result']
    st.subheader('2. Resultado de extracción')
    c1,c2,c3=st.columns(3)
    c1.metric('Campos encontrados',sum(bool(v) for v in data.values()))
    c2.metric('Método',method)
    c3.metric('Ficha identificada',target[:25] if target else 'No encontrada')
    st.write('**Coincidencia:**', reason, f'({score:.0%})' if score else '')
    rows=[]
    for field in FIELD_ALIASES:
        value=data.get(field,'')
        rows.append({'Campo':field,'Información extraída':value,'Estado':'✓ Encontrado' if value else '⚠ No encontrado'})
    st.dataframe(rows, use_container_width=True, hide_index=True)
    st.subheader('3. Incorporar al Excel')
    st.download_button('⬇️ Descargar Excel actualizado', data=out, file_name='Fichas_tecnicas_actualizadas.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    with st.expander('Ver texto bruto extraído'):
        st.text_area('Texto',text,height=350)
else:
    st.info('Carga el Excel y una derivación PDF para comenzar.')
