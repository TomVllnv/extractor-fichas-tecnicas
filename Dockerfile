# Herramienta OCR: Python/Streamlit y motores de lectura óptica.
# El código se instala en la imagen; la base y los respaldos viven en /datos,
# que Docker Compose monta fuera del contenedor.
FROM python:3.12-slim-bookworm

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    HOME=/home/laboratorio \
    HERRAMIENTA_DB_PATH=/datos/base_de_datos.db \
    HERRAMIENTA_BACKUP_DIR=/datos/respaldos

WORKDIR /app

# tzdata garantiza que ZoneInfo('America/Santiago') funcione en Linux.
# El idioma español de Tesseract se instala a nivel de sistema.
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
       ca-certificates tzdata libgl1 libglib2.0-0 libgomp1 \
       tesseract-ocr tesseract-ocr-spa \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt /app/requirements.txt
RUN python -m pip install --no-cache-dir -r /app/requirements.txt

# No incluir .db, .env ni secrets.toml en la imagen: ver .dockerignore.
RUN groupadd --gid 10001 laboratorio \
    && useradd --uid 10001 --gid laboratorio --create-home --shell /usr/sbin/nologin laboratorio \
    && mkdir -p /datos \
    && chown -R 10001:10001 /datos /home/laboratorio

COPY --chown=10001:10001 . /app

USER 10001:10001
EXPOSE 8501

# Comprobación de salud oficial de Streamlit, sin añadir curl al contenedor.
HEALTHCHECK --interval=30s --timeout=5s --start-period=90s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8501/_stcore/health', timeout=3).read()" || exit 1

CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0", "--server.headless=true", "--server.fileWatcherType=none", "--browser.gatherUsageStats=false"]
