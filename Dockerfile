# Despliegue opcional en servidor institucional. Revisar primero con TI.
FROM python:3.12-slim
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1 \
    HERRAMIENTA_DB_PATH=/datos/base_de_datos.db \
    HERRAMIENTA_BACKUP_DIR=/datos/respaldos
WORKDIR /app
RUN apt-get update && apt-get install -y --no-install-recommends \
    libgl1 libglib2.0-0 libgomp1 && rm -rf /var/lib/apt/lists/*
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt
COPY . /app
RUN useradd --uid 10001 --create-home --shell /usr/sbin/nologin laboratorio \
    && mkdir -p /datos && chown 10001:10001 /datos
USER 10001
EXPOSE 8501
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0", "--server.headless=true"]
