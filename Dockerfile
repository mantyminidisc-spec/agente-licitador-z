# =============================================================================
# LICITAIA / GOVILICIT.IA - Dockerfile v5.4 (para Coolify)
# =============================================================================
# Este archivo asume que el agente corregido se llama agente_licitador_z.py
# (recomendado: reemplaza el archivo antiguo conservando el mismo nombre).
# Si le cambias el nombre, ajusta las 2 líneas COPY/CMD del final.
#
# Cambios frente al Dockerfile anterior:
#   1. curl_cffi instalado (sin él, Mercado Público responde 403 por TLS).
#   2. PYTHONUTF8=1 y LICITAIA_DATA_DIR=/data (datos fuera del contenedor).
#   3. TZ=America/Santiago (si no, la fecha del CRM y de los documentos sale en UTC).
#   4. VOLUME /data: sin esto el CRM, el historial y la cola se borran al redeploy.
#   5. El contenedor queda vivo (sleep infinity) para poder ejecutar el agente
#      a mano o con un "Scheduled Task" de Coolify.
#   6. Las credenciales NO se hornean: se cargan desde Coolify > Environment.
#
# NOTA: en Docker los comentarios deben ir en su propia línea. NO se puede
# escribir  COPY archivo.py .   # comentario  porque Docker intenta copiar
# los pedazos "#", "<<<", "comentario"... como si fueran archivos.
# =============================================================================
FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PYTHONUTF8=1 \
    TZ=America/Santiago \
    LICITAIA_DATA_DIR=/data \
    GEMINI_MODEL=gemini-2.5-flash \
    PAUSA_ENTRE_PROYECTOS=15

WORKDIR /app

# Dependencias de sistema para Chromium (Playwright) y pdfplumber
RUN apt-get update && apt-get install -y --no-install-recommends \
    wget curl ca-certificates gnupg tzdata fonts-liberation \
    libglib2.0-0 libnss3 libnspr4 libatk1.0-0 libatk-bridge2.0-0 \
    libcups2 libdrm2 libdbus-1-3 libexpat1 libxcb1 libxkbcommon0 \
    libx11-6 libxcomposite1 libxdamage1 libxext6 libxfixes3 libxrandr2 \
    libgbm1 libpango-1.0-0 libcairo2 libasound2 \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Navegador real para el buscador de Mercado Público (SPA detrás de CloudFront)
RUN playwright install chromium --with-deps

# --- Estos dos COPY deben apuntar al nombre real del agente ----------------
COPY agente_licitador_z.py .
COPY probar_conexion.py .
# Logo corporativo (se usa en DOC 1 y DOC 2). Si no lo quieres, bórralo del repo.
COPY logo.png .

RUN mkdir -p /data
VOLUME ["/data"]

# El contenedor queda vivo: el agente se ejecuta desde la terminal de Coolify
# o con un Scheduled Task (recomendado para producción).
CMD ["sleep", "infinity"]

# -----------------------------------------------------------------------------
# ALTERNATIVA: que el agente corra solo al arrancar el contenedor.
# Comenta la línea CMD de arriba y descomenta esta:
#
#   CMD ["python", "agente_licitador_z.py"]
#
# (Ojo: con esa opción, cada reinicio del contenedor lanza una corrida nueva.)
# -----------------------------------------------------------------------------
