# Extractor de Fichas Técnicas

Prototipo funcional para transformar información de documentos de derivación en datos estructurados para el Excel de fichas técnicas.

## Qué hace
- Carga un PDF de derivación.
- Extrae texto del PDF.
- Reconoce distintas formas de nombrar campos (por ejemplo, `TIPO DE MUESTRA`, `MUESTRA REQUERIDA` o `MATERIAL BIOLÓGICO`).
- Identifica código y nombre del examen cuando están disponibles.
- Busca la ficha correspondiente en el Excel por código o nombre.
- Actualiza una **copia** del Excel, conservando la estructura del archivo original.
- Crea una hoja `LOG IMPORTACION` con lo encontrado, lo incorporado y lo que requiere revisión.
- Si no encuentra una ficha existente, no inventa ni modifica una ficha equivocada; deja el resultado como `Sin coincidencia` para revisión.
- No genera PDF.

## Instalación
Se recomienda Python 3.11 o superior.

```bash
pip install -r requirements.txt
```

## Ejecutar
En Windows:

```bash
streamlit run app.py
```

Se abrirá una página local en el navegador.

## Uso
1. Cargar el Excel `Fichas tecnicas examenes microbiologicos 2025 V01 (2).xlsx`.
2. Opcionalmente indicar el centro de salud.
3. Cargar el PDF de derivación.
4. Pulsar `Analizar documento`.
5. Revisar los campos extraídos.
6. Descargar `Fichas_tecnicas_actualizadas.xlsx`.

## Limitación importante
Esta primera versión funciona mejor con PDF que contienen texto seleccionable. Para PDFs escaneados se puede ampliar con OCR (PyMuPDF + Tesseract) en una segunda iteración.

## Seguridad y calidad
El programa no debe inventar valores clínicos. Si un dato no aparece, se marca como `No encontrado`. Los campos encontrados deben ser revisados por la persona responsable antes de utilizar la ficha.
