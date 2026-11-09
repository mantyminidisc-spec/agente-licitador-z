# =============================================================================
# LICITAIA · Dockerfile v6.4 (build rápido y sin dependencia de CDN externo)
# =============================================================================
# Cambios respecto a la versión anterior:
#   1. Chromium YA NO se instala por defecto: la descarga desde cdn.playwright.dev
#      es el punto más frágil del build (timeouts de 30s x 5 intentos).
#      Los adjuntos de Compra Ágil vienen del servicio oficial (REST) y los de
#      licitaciones de la ficha + visor, así que el navegador es solo respaldo.
#      Si algún día lo necesitas:  INSTALAR_NAVEGADOR=1
#   2. Se quitaron las librerías de sistema de Chromium (278 MB menos de build).
#      Si activas el navegador, "--with-deps" las instala solo.
#   3. curl_cffi, TZ Chile, volumen /data y contenedor vivo: igual que antes.
#
# IMPORTANTE: no pongas secretos (API keys) como Build Arguments en Coolify;
# se hornean en la imagen. Van en Environment, no en el build.
# =============================================================================
FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PYTHONUTF8=1 \
    TZ=America/Santiago \
    LICITAIA_DATA_DIR=/data \
    GEMINI_MODEL=gemini-3.8-flash \
    PAUSA_ENTRE_PROYECTOS=15

WORKDIR /app

# Solo lo mínimo: certificados, zona horaria, fuentes y OCR local (Tesseract)
RUN apt-get update && apt-get install -y --no-install-recommends \
        antiword \
        catdoc
    ca-certificates curl tzdata fonts-liberation \
    tesseract-ocr tesseract-ocr-spa \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Navegador OPCIONAL (apagado por defecto). No rompe el build si el CDN falla.
ARG INSTALAR_NAVEGADOR=0
RUN if [ "$INSTALAR_NAVEGADOR" = "1" ]; then \
        playwright install chromium --with-deps ; \
    else \
        echo "Navegador omitido (INSTALAR_NAVEGADOR=0). El agente funciona sin él." ; \
    fi

COPY agente_licitador_z.py .
COPY probar_conexion.py .
# Portal de entregables (sirve las páginas privadas y las descargas)
COPY portal.py .
# Planificador interno: corre el agente a las horas configuradas (sin Coolify)
COPY planificador.py .
# Logo corporativo (DOC 1 y DOC 2). Si no lo quieres, bórralo del repo.
COPY logo.png .
# Perfiles de clientes (multi-cliente). Se pueden sobrescribir en /data/clientes
COPY clientes ./clientes

RUN mkdir -p /data
VOLUME ["/data"]

EXPOSE 8090

# El contenedor queda vivo: el portal atiende en 8090 y el agente se ejecuta
# desde la Terminal de Coolify o con un Scheduled Task.
CMD ["sh", "-c", "python portal.py & python planificador.py & exec sleep infinity"]

# -----------------------------------------------------------------------------
# ALTERNATIVA: que el agente corra solo al arrancar el contenedor.
# Cambia la línea CMD de arriba por:
#
#   CMD ["sh", "-c", "python portal.py & exec python agente_licitador_z.py"]
# -----------------------------------------------------------------------------
