# -*- coding: utf-8 -*-
"""
===============================================================================
LICITAIA · PORTAL DE ENTREGABLES (v7.0)
===============================================================================
Sirve una página privada por proyecto, con enlace con token y descarga del
paquete completo en ZIP. Sin dependencias externas (solo librería estándar).

Rutas
    GET /salud                     estado del servicio (JSON)
    GET /p/<token>                 página del proyecto
    GET /p/<token>/zip             descarga todo el paquete
    GET /p/<token>/f/<archivo>     descarga un archivo

Variables de entorno
    LICITAIA_DATA_DIR   carpeta de datos (por defecto /data)
    PORTAL_SECRETO      secreto para firmar los tokens (cámbialo)
    PORTAL_PUERTO       puerto de escucha (por defecto 8090)
    PORTAL_URL_BASE     URL pública, ej. http://1.2.3.4:8090
===============================================================================
"""
import hashlib
import hmac
import json
import os
import socket
import sys
import zipfile
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from io import BytesIO
from urllib.parse import quote, unquote

RAIZ = os.getenv("LICITAIA_DATA_DIR") or os.getcwd()
ENTREGABLES = os.path.join(RAIZ, "Entregables_Drive_GoviLicit")
SECRETO = os.getenv("PORTAL_SECRETO", "cambia-este-secreto-licitaia").encode("utf-8")
PUERTO = int(os.getenv("PORTAL_PUERTO", "8090"))
URL_BASE = os.getenv("PORTAL_URL_BASE", "").rstrip("/")
MAX_ZIP = 80 * 1024 * 1024

_cache = {"t": 0.0, "proyectos": {}}


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------
def _token(codigo):
    """Token estable y no adivinable para un proyecto."""
    firma = hmac.new(SECRETO, str(codigo).encode("utf-8"), hashlib.sha256).hexdigest()
    return firma[:20]


def _formato_clp(valor):
    try:
        return "$" + f"{int(valor):,}".replace(",", ".")
    except Exception:
        return "$0"


def _proyectos(recargar=False):
    """Mapa token -> manifest, escaneando los manifest.json del disco."""
    ahora = datetime.now().timestamp()
    if not recargar and (ahora - _cache["t"]) < 20 and _cache["proyectos"]:
        return _cache["proyectos"]

    encontrados = {}
    for raiz, _, archivos in os.walk(ENTREGABLES):
        if "manifest.json" not in archivos:
            continue
        ruta = os.path.join(raiz, "manifest.json")
        try:
            with open(ruta, "r", encoding="utf-8") as f:
                datos = json.load(f)
        except Exception:
            continue
        codigo = str(datos.get("codigo") or os.path.basename(raiz))
        datos["_carpeta"] = raiz
        datos["_codigo"] = codigo
        encontrados[_token(codigo)] = datos

    _cache["t"] = ahora
    _cache["proyectos"] = encontrados
    return encontrados


def _archivos(carpeta):
    salida = []
    for nombre in sorted(os.listdir(carpeta)):
        ruta = os.path.join(carpeta, nombre)
        if not os.path.isfile(ruta) or nombre == "manifest.json":
            continue
        if nombre.startswith("."):
            continue
        salida.append({"nombre": nombre, "bytes": os.path.getsize(ruta)})
    return salida


def _adjuntos(carpeta):
    """Archivos originales del portal (subcarpeta adjuntos/)."""
    salida = []
    ruta_adj = os.path.join(carpeta, "adjuntos")
    if not os.path.isdir(ruta_adj):
        return salida
    for raiz, _, archivos in os.walk(ruta_adj):
        for nombre in sorted(archivos):
            if nombre.startswith("."):
                continue
            ruta = os.path.join(raiz, nombre)
            relativo = os.path.relpath(ruta, ruta_adj).replace("\\", "/")
            salida.append({"nombre": relativo, "bytes": os.path.getsize(ruta)})
    return salida


def _zip_proyecto(carpeta):
    buffer = BytesIO()
    total = 0
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for raiz, _, archivos in os.walk(carpeta):
            for nombre in sorted(archivos):
                if nombre.startswith("."):
                    continue
                ruta = os.path.join(raiz, nombre)
                relativo = os.path.relpath(ruta, carpeta)
                total += os.path.getsize(ruta)
                if total > MAX_ZIP:
                    break
                zf.write(ruta, relativo)
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# HTML
# ---------------------------------------------------------------------------
PLANTILLA = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow">
<title>{codigo} · LICITAIA</title>
<style>
 :root {{ --azul:#1f2a44; --gris:#6b7280; --linea:#e5e7eb; --fondo:#f7f8fa; }}
 * {{ box-sizing:border-box; }}
 body {{ margin:0; font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;
        color:#111827; background:var(--fondo); }}
 header {{ background:var(--azul); color:#fff; padding:22px 24px; }}
 header .marca {{ font-size:13px; letter-spacing:.14em; opacity:.75; text-transform:uppercase; }}
 header h1 {{ margin:8px 0 4px; font-size:22px; font-weight:600; line-height:1.25; }}
 header .meta {{ font-size:13px; opacity:.8; }}
 main {{ max-width:880px; margin:0 auto; padding:24px; }}
 .tarjeta {{ background:#fff; border:1px solid var(--linea); border-radius:10px; padding:20px; margin-bottom:16px; }}
 .kpis {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(150px,1fr)); gap:12px; margin-bottom:16px; }}
 .kpi {{ background:#fff; border:1px solid var(--linea); border-radius:10px; padding:14px 16px; }}
 .kpi .et {{ font-size:10px; letter-spacing:.1em; text-transform:uppercase; color:var(--gris); }}
 .kpi .va {{ font-size:19px; font-weight:600; color:var(--azul); margin-top:4px; }}
 .kpi .no {{ font-size:11px; color:var(--gris); }}
 h2 {{ font-size:12px; letter-spacing:.1em; text-transform:uppercase; color:var(--azul);
      border-bottom:1px solid var(--azul); padding-bottom:6px; margin:0 0 12px; }}
 table {{ width:100%; border-collapse:collapse; font-size:13px; }}
 td {{ padding:7px 4px; border-bottom:1px solid var(--linea); }}
 td.n {{ text-align:right; color:var(--gris); white-space:nowrap; }}
 a.boton {{ display:inline-block; background:var(--azul); color:#fff; text-decoration:none;
           padding:12px 20px; border-radius:8px; font-weight:600; font-size:14px; }}
 a.boton:hover {{ opacity:.9; }}
 a.link {{ color:#1d4ed8; text-decoration:none; font-size:13px; }}
 footer {{ text-align:center; color:var(--gris); font-size:11px; padding:24px; }}
 .aviso {{ background:#fff7ed; border:1px solid #fed7aa; color:#9a3412; font-size:12px;
          padding:10px 12px; border-radius:8px; }}
</style>
</head>
<body>
<header>
  <div class="marca">LICITAIA · entregables</div>
  <h1>{nombre}</h1>
  <div class="meta">{organismo} · Proceso {codigo} · {tipo}</div>
</header>
<main>
  <div class="kpis">
    <div class="kpi"><div class="et">Presupuesto</div><div class="va">{presupuesto}</div><div class="no">IVA incluido</div></div>
    <div class="kpi"><div class="et">Oferta recomendada</div><div class="va">{oferta}</div><div class="no">con IVA</div></div>
    <div class="kpi"><div class="et">Cierre</div><div class="va" style="font-size:15px">{cierre}</div><div class="no">plazo fatal</div></div>
  </div>

  <div class="tarjeta">
    <h2>Documentos del proyecto</h2>
    <table>
      {filas}
    </table>
    <p style="margin:18px 0 0"><a class="boton" href="/p/{token}/zip">Descargar todo (ZIP)</a></p>
  </div>

  {bloque_adjuntos}

  <div class="tarjeta">
    <h2>Cómo se usa</h2>
    <p style="font-size:13px; line-height:1.6; margin:0">
      Descarga el ZIP, revisa el <b>DOC 2</b> (cotización) y ajusta lo que necesites en Word.
      El <b>DOC 1</b> es tu auditoría interna y el <b>Excel</b> el respaldo financiero.
      Si quieres otro precio, pídelo por WhatsApp y se regeneran los documentos.
    </p>
  </div>

  <div class="aviso">
    Enlace privado. No lo compartas: da acceso a la información comercial de este proyecto.
  </div>
</main>
<footer>Generado por LICITAIA · documento válido como respaldo de postulación</footer>
</body>
</html>
"""


def _pagina(proyecto):
    token = _token(proyecto["_codigo"])
    archivos = _archivos(proyecto["_carpeta"])
    adjuntos = _adjuntos(proyecto["_carpeta"])
    filas = "".join(
        f'<tr><td><a class="link" href="/p/{token}/f/{quote(a["nombre"])}">{a["nombre"]}</a></td>'
        f'<td class="n">{a["bytes"] // 1024} KB</td></tr>'
        for a in archivos
    ) or '<tr><td colspan="2">Todavía no hay documentos en esta carpeta.</td></tr>'
    filas_adj = "".join(
        f'<tr><td><a class="link" href="/p/{token}/a/{quote(a["nombre"])}">{a["nombre"]}</a></td>'
        f'<td class="n">{a["bytes"] // 1024} KB</td></tr>'
        for a in adjuntos
    )
    bloque_adj = ""
    if filas_adj:
        bloque_adj = (
            '<div class="tarjeta"><h2>Adjuntos originales del portal (' + str(len(adjuntos)) + ')</h2>'
            '<p style="font-size:12px;color:#6b7280;margin:0 0 10px">'
            'Documentos oficiales tal como los publicó el organismo (PDF, DOCX, XLSX).</p>'
            '<table>' + filas_adj + '</table></div>'
        )
    return PLANTILLA.format(
        token=token,
        codigo=proyecto.get("codigo", ""),
        nombre=str(proyecto.get("nombre", "Proyecto"))[:140],
        organismo=str(proyecto.get("organismo", ""))[:90],
        tipo="Compra Ágil" if proyecto.get("tipo_portal") == "COMPRA_AGIL" else "Licitación pública",
        presupuesto=_formato_clp(proyecto.get("presupuesto_clp", 0)),
        oferta=_formato_clp(proyecto.get("oferta_bruta_clp", 0)),
        cierre=proyecto.get("fecha_cierre", "ver ficha"),
        filas=filas,
        bloque_adjuntos=bloque_adj,
    )


# ---------------------------------------------------------------------------
# Servidor
# ---------------------------------------------------------------------------
class Manejador(BaseHTTPRequestHandler):
    server_version = "LicitaiaPortal/7.0"

    def log_message(self, formato, *args):
        sys.stdout.write("%s - %s\n" % (self.address_string(), formato % args))

    def _enviar(self, codigo, cuerpo, tipo="text/html; charset=utf-8", extra=None):
        if isinstance(cuerpo, str):
            cuerpo = cuerpo.encode("utf-8")
        self.send_response(codigo)
        self.send_header("Content-Type", tipo)
        self.send_header("Content-Length", str(len(cuerpo)))
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Referrer-Policy", "no-referrer")
        for clave, valor in (extra or {}).items():
            self.send_header(clave, valor)
        self.end_headers()
        self.wfile.write(cuerpo)

    def do_GET(self):
        ruta = self.path.split("?")[0]
        partes = [p for p in ruta.split("/") if p]

        if ruta == "/salud" or ruta == "/":
            datos = {"servicio": "licitaia-portal", "estado": "ok",
                     "proyectos": len(_proyectos()),
                     "carpeta": ENTREGABLES}
            return self._enviar(200, json.dumps(datos, ensure_ascii=False), "application/json")

        if len(partes) >= 2 and partes[0] == "p":
            token = partes[1]
            proyecto = _proyectos().get(token)
            if not proyecto:
                _proyectos(recargar=True)
                proyecto = _proyectos().get(token)
            if not proyecto:
                return self._enviar(404, "<h1>Enlace no válido</h1>"
                                         "<p>El proyecto no existe o el enlace fue revocado.</p>")

            # /p/<token>/zip
            if len(partes) == 3 and partes[2] == "zip":
                datos = _zip_proyecto(proyecto["_carpeta"])
                nombre = f"{proyecto['_codigo']}.zip"
                return self._enviar(200, datos, "application/zip",
                                    {"Content-Disposition": f'attachment; filename="{nombre}"'})

            # /p/<token>/f/<archivo>
            if len(partes) == 4 and partes[2] == "f":
                pedido = unquote(partes[3])
                seguro = os.path.basename(pedido)
                ruta_archivo = os.path.join(proyecto["_carpeta"], seguro)
                if not os.path.isfile(ruta_archivo):
                    return self._enviar(404, "<h1>Archivo no encontrado</h1>")
                with open(ruta_archivo, "rb") as f:
                    datos = f.read()
                tipo = "application/octet-stream"
                if seguro.lower().endswith(".pdf"):
                    tipo = "application/pdf"
                return self._enviar(200, datos, tipo,
                                    {"Content-Disposition": f'attachment; filename="{seguro}"'})

            # /p/<token>/a/<adjunto>  (originales del portal, en adjuntos/)
            if len(partes) == 4 and partes[2] == "a":
                pedido = unquote(partes[3])
                seguro = os.path.basename(pedido)
                ruta_archivo = os.path.join(proyecto["_carpeta"], "adjuntos", seguro)
                if not os.path.isfile(ruta_archivo):
                    return self._enviar(404, "<h1>Adjunto no encontrado</h1>")
                with open(ruta_archivo, "rb") as f:
                    datos = f.read()
                tipo = "application/octet-stream"
                if seguro.lower().endswith(".pdf"):
                    tipo = "application/pdf"
                return self._enviar(200, datos, tipo,
                                    {"Content-Disposition": f'attachment; filename="{seguro}"'})

            return self._enviar(200, _pagina(proyecto))

        return self._enviar(404, "<h1>No encontrado</h1>")


def main():
    os.makedirs(ENTREGABLES, exist_ok=True)
    print(f"LICITAIA PORTAL · carpeta: {ENTREGABLES}")
    print(f"LICITAIA PORTAL · proyectos indexados: {len(_proyectos())}")
    print(f"LICITAIA PORTAL · escuchando en 0.0.0.0:{PUERTO}"
          + (f" · URL base: {URL_BASE}" if URL_BASE else ""))
    servidor = ThreadingHTTPServer(("0.0.0.0", PUERTO), Manejador)
    try:
        servidor.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
