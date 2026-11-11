# -*- coding: utf-8 -*-
"""
===============================================================================
LICITAIA · PORTAL DE ENTREGABLES (v2)
===============================================================================
Sirve una página privada por proyecto, con enlace con token, descarga en ZIP y
SUBIDA de anexos (para los documentos que el portal oficial no permite bajar:
EETT, BAE, formularios) + botón para regenerar la auditoría con ellos.

Rutas
    GET  /salud                       estado del servicio (JSON)
    GET  /p/<token>                   página del proyecto
    GET  /p/<token>/zip               descarga todo el paquete
    GET  /p/<token>/f/<archivo>       descarga un archivo
    GET  /p/<token>/a/<adjunto>       descarga un adjunto original
    POST /p/<token>/subir             sube anexos (multipart)
    POST /p/<token>/reprocesar        relanza el agente con los anexos subidos

Variables de entorno
    LICITAIA_DATA_DIR   carpeta de datos (por defecto /data)
    PORTAL_SECRETO      secreto para firmar los tokens
    PORTAL_PUERTO       puerto de escucha (por defecto 8090)
    PORTAL_URL_BASE     URL pública, ej. https://licitaia.duckdns.org
===============================================================================
"""
import hashlib
import hmac
import json
import os
import re
import shutil
import subprocess
import sys
import time
import zipfile

try:
    import portal_ui
except Exception:
    portal_ui = None
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
MAX_SUBIDA = 60 * 1024 * 1024
BLOQUEO = os.path.join(RAIZ, "portal_reproceso.lock")
RUTA_LOGO = os.path.join(RAIZ, "logo.png")

_cache = {"t": 0.0, "proyectos": {}}

# Qué documentos conviene subir para una auditoría técnica completa.
# Cada institución entrega las bases con nombres distintos, así que se aceptan
# muchas variantes (TR, TDR, anexos técnicos, especificaciones, etc.).
CHECKLIST = [
    ("Bases administrativas (BAE / BAG / DAE)",
     ("bae", "bag", "bases", "administrativ", "dae")),
    ("Especificaciones o bases técnicas (EETT / TR / TDR)",
     ("eett", "tecnica", "técnica", "tecnicos", "técnicos", "especificacion",
      "especificación", "termino", "término", "tr_", "tdr", "referencia")),
    ("Formularios, anexos y económicos",
     ("formato", "anexo", "declaracion", "declaración", "carta", "economico", "económico")),
]


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------
def _token(codigo):
    firma = hmac.new(SECRETO, str(codigo).encode("utf-8"), hashlib.sha256).hexdigest()
    return firma[:20]


def _formato_clp(valor):
    try:
        return "$" + f"{int(valor):,}".replace(",", ".")
    except Exception:
        return "$0"


def _proyectos(recargar=False):
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


def _leer_manifest(carpeta):
    try:
        with open(os.path.join(carpeta, "manifest.json"), encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _guardar_manifest(carpeta, datos):
    try:
        with open(os.path.join(carpeta, "manifest.json"), "w", encoding="utf-8") as f:
            json.dump(datos, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


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
    salida = []
    ruta_adj = os.path.join(carpeta, "adjuntos")
    if not os.path.isdir(ruta_adj):
        return salida
    for raiz, _, archivos in os.walk(ruta_adj):
        for nombre in sorted(archivos):
            if nombre.startswith("."):
                continue
            # Archivos internos del agente: no son adjuntos del portal
            if nombre.upper().startswith(("FICHA_TEXTO_", "CORREO_PENDIENTE", "MANIFEST")):
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
                total += os.path.getsize(ruta)
                if total > MAX_ZIP:
                    break
                zf.write(ruta, os.path.relpath(ruta, carpeta))
    return buffer.getvalue()


def _leer_multipart(cuerpo, tipo_contenido):
    """Parser multipart mínimo (sin dependencias)."""
    m = re.search(r"boundary=([^;]+)", tipo_contenido or "")
    if not m:
        return []
    limite = m.group(1).strip().strip('"').encode()
    archivos = []
    for parte in cuerpo.split(b"--" + limite):
        if b"filename=" not in parte:
            continue
        cabecera, _, datos = parte.partition(b"\r\n\r\n")
        mf = re.search(rb'filename="([^"]*)"', cabecera)
        if not mf:
            continue
        nombre = os.path.basename(mf.group(1).decode("utf-8", "ignore"))
        datos = datos.rstrip(b"\r\n")
        if datos[-2:] == b"--":
            datos = datos[:-2].rstrip(b"\r\n")
        if nombre and datos:
            archivos.append((nombre, datos))
    return archivos


def _agente_python():
    for ruta in ("/app/agente_licitador_z.py",
                 os.path.join(os.path.dirname(os.path.abspath(__file__)), "agente_licitador_z.py")):
        if os.path.isfile(ruta):
            return ruta
    return ""


def _lanzar_reproceso(codigo):
    """Lanza el agente en segundo plano. Devuelve (ok, mensaje)."""
    if os.path.exists(BLOQUEO):
        try:
            with open(BLOQUEO, encoding="utf-8") as f:
                pid = int((f.read() or "0").strip())
            os.kill(pid, 0)
            return False, "Ya hay un procesamiento en curso."
        except Exception:
            pass
    script = _agente_python()
    if not script:
        return False, "No se encontró el agente en el contenedor."
    registro = os.path.join(RAIZ, f"reproceso_{codigo}.log")
    try:
        with open(registro, "ab") as salida:
            proceso = subprocess.Popen(
                [sys.executable, script, "--reprocesar", codigo],
                cwd=os.path.dirname(script), stdout=salida, stderr=subprocess.STDOUT,
            )
        with open(BLOQUEO, "w", encoding="utf-8") as f:
            f.write(str(proceso.pid))
        return True, f"Reproceso iniciado (PID {proceso.pid}). Revisa el registro {registro}"
    except Exception as e:
        return False, f"No se pudo iniciar: {e}"


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
 a.boton, button.boton {{ display:inline-block; background:var(--azul); color:#fff; text-decoration:none;
           padding:12px 20px; border-radius:8px; font-weight:600; font-size:14px; border:none; cursor:pointer; }}
 a.boton:hover, button.boton:hover {{ opacity:.9; }}
 a.link {{ color:#1d4ed8; text-decoration:none; font-size:13px; }}
 footer {{ text-align:center; color:var(--gris); font-size:11px; padding:24px; }}
 .aviso {{ background:#fff7ed; border:1px solid #fed7aa; color:#9a3412; font-size:12px;
          padding:10px 12px; border-radius:8px; }}
 .ok {{ background:#f0fdf4; border:1px solid #bbf7d0; color:#166534; font-size:12px;
       padding:10px 12px; border-radius:8px; margin-bottom:12px; }}
 .falta {{ background:#fffbeb; border:1px solid #fde68a; color:#92400e; font-size:12px;
          padding:10px 12px; border-radius:8px; margin-bottom:12px; }}
 .lista-check {{ list-style:none; padding:0; margin:0 0 14px; font-size:13px; }}
 .lista-check li {{ padding:5px 0; }}
 .si {{ color:#166534; font-weight:600; }}
 .no {{ color:#b45309; font-weight:600; }}
 input[type=file] {{ font-size:13px; margin:8px 0 12px; width:100%; }}
 .ayuda {{ font-size:12px; color:var(--gris); }}
</style>
</head>
<body>
<header>
  <div class="marca">LICITAIA · entregables</div>
  <h1>{nombre}</h1>
  <div class="meta">{organismo} · Proceso {codigo} · {tipo}</div>
</header>
<main>
  {aviso}
  <div class="kpis">
    <div class="kpi"><div class="et">Presupuesto</div><div class="va">{presupuesto}</div><div class="no">IVA incluido</div></div>
    <div class="kpi"><div class="et">Oferta recomendada</div><div class="va">{oferta}</div><div class="no">con IVA</div></div>
    <div class="kpi"><div class="et">Cierre</div><div class="va" style="font-size:15px">{cierre}</div><div class="no">plazo fatal</div></div>
  </div>

  {boton_portal}

  {bloque_logo}

  <div class="tarjeta">
    <h2>Documentos del proyecto</h2>
    <table>
      {filas}
    </table>
    <p style="margin:18px 0 0"><a class="boton" href="/p/{token}/zip">Descargar todo (ZIP)</a></p>
  </div>

  {bloque_adjuntos}

  {bloque_anexos}

  <div class="tarjeta">
    <h2>Cómo se usa</h2>
    <p style="font-size:13px; line-height:1.6; margin:0">
      Descarga el ZIP y revisa el <b>DOC 2</b> (cotización). El <b>DOC 1</b> es tu
      auditoría interna y el <b>Excel</b> el respaldo financiero. Si quieres otro
      precio, pídelo por WhatsApp y se regeneran los documentos.
    </p>
  </div>

  <div class="aviso">
    Enlace privado. No lo compartas: da acceso a la información comercial de este proyecto.
  </div>
</main>
<footer>Generado por LICITAIA · documento válido como respaldo de postulación</footer>

<div id="listo" style="display:none;position:fixed;left:50%;top:18px;transform:translateX(-50%);
     background:#0f766e;color:#fff;padding:14px 22px;border-radius:12px;font-weight:600;
     box-shadow:0 10px 30px rgba(0,0,0,.25);z-index:99;text-align:center">
  ✅ Análisis actualizado. Recargando…
</div>
<script>
(function () {{
  var inicial = "{marca_inicial}";
  var objetivo = "/p/{token}/estado";
  var mostrado = false;
  setInterval(function () {{
    if (mostrado) return;
    fetch(objetivo, {{ cache: "no-store" }})
      .then(function (r) {{ return r.json(); }})
      .then(function (d) {{
        if (String(d.doc1) !== inicial && d.doc1) {{
          mostrado = true;
          document.getElementById("listo").style.display = "block";
          setTimeout(function () {{ location.href = "/p/{token}"; }}, 1800);
        }}
      }})
      .catch(function () {{}});
  }}, 12000);
}})();
</script>
</body>
</html>
"""


def _estado_analisis(carpeta):
    """Para que la página avise sola cuando el análisis quedó actualizado."""
    doc1 = 0
    generado = ""
    try:
        for nombre in os.listdir(carpeta):
            if nombre.upper().startswith("DOC1_") and nombre.upper().endswith(".DOCX"):
                doc1 = int(os.path.getmtime(os.path.join(carpeta, nombre)))
    except Exception:
        doc1 = 0
    try:
        with open(os.path.join(carpeta, "manifest.json"), encoding="utf-8") as f:
            generado = str((json.load(f) or {}).get("generado") or "")
    except Exception:
        generado = ""
    return {"doc1": doc1, "generado": generado}


def _pagina(proyecto, aviso=""):
    token = _token(proyecto["_codigo"])
    carpeta = proyecto["_carpeta"]
    archivos = _archivos(carpeta)
    adjuntos = _adjuntos(carpeta)
    if portal_ui is None:
        return "<h1>Panel no disponible</h1><p>Falta portal_ui.py en la imagen.</p>"
    return portal_ui.pagina(
        proyecto=proyecto,
        token=token,
        archivos=archivos,
        adjuntos=adjuntos,
        aviso=aviso,
        hay_logo=os.path.isfile(RUTA_LOGO),
        logo_url=f"/p/{token}/logo",
        doc1_marca=_estado_analisis(carpeta).get("doc1", 0),
    )


class Manejador(BaseHTTPRequestHandler):
    server_version = "LicitaiaPortal/2.0"

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

    def _redirigir(self, destino):
        self.send_response(303)
        self.send_header("Location", destino)
        self.send_header("Content-Length", "0")
        self.end_headers()

    def _buscar(self, token):
        proyecto = _proyectos().get(token)
        if not proyecto:
            _proyectos(recargar=True)
            proyecto = _proyectos().get(token)
        return proyecto

    def do_GET(self):
        ruta = self.path.split("?")[0]
        consulta = self.path.split("?")[1] if "?" in self.path else ""
        partes = [p for p in ruta.split("/") if p]

        if ruta in ("/salud", "/"):
            datos = {"servicio": "licitaia-portal", "estado": "ok", "version": "2.0",
                     "proyectos": len(_proyectos()), "carpeta": ENTREGABLES}
            return self._enviar(200, json.dumps(datos, ensure_ascii=False), "application/json")

        if len(partes) >= 2 and partes[0] == "p":
            token = partes[1]
            proyecto = self._buscar(token)
            if not proyecto:
                return self._enviar(404, "<h1>Enlace no válido</h1>"
                                         "<p>El proyecto no existe o el enlace fue revocado.</p>")

            if len(partes) == 3 and partes[2] == "logo":
                if not os.path.isfile(RUTA_LOGO):
                    return self._enviar(404, "<h1>Sin logo</h1>")
                with open(RUTA_LOGO, "rb") as f:
                    datos = f.read()
                tipo = "image/png" if RUTA_LOGO.lower().endswith(".png") else "image/jpeg"
                return self._enviar(200, datos, tipo, {"Cache-Control": "no-cache"})

            if len(partes) == 3 and partes[2] == "estado":
                datos = _estado_analisis(proyecto["_carpeta"])
                return self._enviar(200, json.dumps(datos, ensure_ascii=False),
                                    "application/json")

            if len(partes) == 3 and partes[2] == "zip":
                datos = _zip_proyecto(proyecto["_carpeta"])
                nombre = f"{proyecto['_codigo']}.zip"
                return self._enviar(200, datos, "application/zip",
                                    {"Content-Disposition": f'attachment; filename="{nombre}"'})

            if len(partes) == 4 and partes[2] in ("ver", "va"):
                seguro = os.path.basename(unquote(partes[3]))
                base = proyecto["_carpeta"] if partes[2] == "ver" else os.path.join(
                    proyecto["_carpeta"], "adjuntos")
                ruta_archivo = os.path.join(base, seguro)
                if not os.path.isfile(ruta_archivo):
                    return self._enviar(404, "<h1>Archivo no encontrado</h1>")
                contenido, es_html = (portal_ui.vista_archivo(ruta_archivo, seguro)
                                     if portal_ui else (None, False))
                if contenido is None:
                    with open(ruta_archivo, "rb") as f:
                        datos = f.read()
                    return self._enviar(200, datos, "application/pdf",
                                        {"Content-Disposition": f'inline; filename="{seguro}"'})
                marco = ("<!DOCTYPE html><html lang='es'><head><meta charset='utf-8'>"
                         "<meta name='viewport' content='width=device-width,initial-scale=1'>"
                         f"<title>{seguro}</title><style>body{{margin:0;padding:20px;"
                         "font-family:-apple-system,'Segoe UI',Roboto,sans-serif;font-size:14px;"
                         "color:#1c2536;line-height:1.55}}table{{border-collapse:collapse;width:100%}}"
                         "th,td{{border:1px solid #e3e7ee;padding:6px 8px;font-size:13px}}"
                         "th{{background:#f7f9fc;text-align:left}}h3{{color:#12213a}}"
                         "h4{{color:#1e3a5f}}pre{{white-space:pre-wrap}}"
                         "caption{{text-align:left;font-weight:700;color:#5b6674;padding:6px 0}}"
                         "</style></head><body>"
                         f"{contenido}</body></html>")
                return self._enviar(200, marco, "text/html; charset=utf-8")

            if len(partes) == 4 and partes[2] in ("f", "a"):
                seguro = os.path.basename(unquote(partes[3]))
                base = proyecto["_carpeta"] if partes[2] == "f" else os.path.join(
                    proyecto["_carpeta"], "adjuntos")
                ruta_archivo = os.path.join(base, seguro)
                if not os.path.isfile(ruta_archivo):
                    return self._enviar(404, "<h1>Archivo no encontrado</h1>")
                with open(ruta_archivo, "rb") as f:
                    datos = f.read()
                tipo = "application/pdf" if seguro.lower().endswith(".pdf") else "application/octet-stream"
                return self._enviar(200, datos, tipo,
                                    {"Content-Disposition": f'attachment; filename="{seguro}"'})

            aviso = ""
            if "subido=" in consulta:
                aviso = '<div class="ok">Anexos recibidos correctamente.</div>'
            if "run=1" in consulta:
                aviso = ('<div class="ok">Reproceso iniciado. En unos minutos el informe '
                         'estará actualizado en esta misma página.</div>')
            return self._enviar(200, _pagina(proyecto, aviso))

        return self._enviar(404, "<h1>No encontrado</h1>")

    def do_POST(self):
        ruta = self.path.split("?")[0]
        partes = [p for p in ruta.split("/") if p]
        if len(partes) != 3 or partes[0] != "p" or partes[2] not in ("subir", "reprocesar", "logo"):
            return self._enviar(404, "<h1>No encontrado</h1>")
        proyecto = self._buscar(partes[1])
        if not proyecto:
            return self._enviar(404, "<h1>Enlace no válido</h1>")

        carpeta = proyecto["_carpeta"]
        codigo = proyecto["_codigo"]

        if partes[2] == "logo":
            largo = int(self.headers.get("Content-Length") or 0)
            if largo <= 0 or largo > 5 * 1024 * 1024:
                return self._enviar(413, "<h1>Archivo demasiado grande (máx 5 MB)</h1>")
            cuerpo = self.rfile.read(largo)
            archivos = _leer_multipart(cuerpo, self.headers.get("Content-Type", ""))
            if not archivos:
                return self._enviar(400, "<h1>No se recibió el logo</h1>")
            nombre, datos = archivos[0]
            extension = os.path.splitext(nombre)[1].lower()
            if extension not in (".png", ".jpg", ".jpeg"):
                return self._enviar(400, "<h1>Formato no válido: usa PNG o JPG</h1>")
            try:
                os.makedirs(os.path.dirname(RUTA_LOGO) or ".", exist_ok=True)
                if os.path.exists(RUTA_LOGO):
                    try:
                        os.replace(RUTA_LOGO, RUTA_LOGO + ".anterior")
                    except Exception:
                        pass
                with open(RUTA_LOGO, "wb") as f:
                    f.write(datos)
                print(f"   [PORTAL] logo del cliente actualizado: {nombre} ({len(datos)} bytes)")
            except Exception as e:
                print(f"   [PORTAL] error guardando el logo: {e}")
                return self._enviar(500, "<h1>No se pudo guardar el logo</h1>")
            return self._redirigir(f"/p/{partes[1]}?logo=1")

        if partes[2] == "reprocesar":
            ok, mensaje = _lanzar_reproceso(codigo)
            print(f"   [PORTAL] reproceso {codigo}: {mensaje}")
            return self._redirigir(f"/p/{partes[1]}" + ("?run=1" if ok else ""))

        largo = int(self.headers.get("Content-Length") or 0)
        if largo <= 0 or largo > MAX_SUBIDA:
            return self._enviar(413, "<h1>Archivo demasiado grande</h1>")
        cuerpo = self.rfile.read(largo)
        archivos = _leer_multipart(cuerpo, self.headers.get("Content-Type", ""))
        if not archivos:
            return self._enviar(400, "<h1>No se recibió ningún archivo</h1>")

        destino = os.path.join(carpeta, "adjuntos")
        os.makedirs(destino, exist_ok=True)
        guardados = []
        for nombre, datos in archivos:
            try:
                ruta_final = os.path.join(destino, nombre)
                with open(ruta_final, "wb") as f:
                    f.write(datos)
                guardados.append(nombre)
                print(f"   [PORTAL] subido {codigo}: {nombre} ({len(datos)} bytes)")
            except Exception as e:
                print(f"   [PORTAL] error guardando {nombre}: {e}")

        manifest = _leer_manifest(carpeta)
        previos = manifest.get("anexos_subidos") or []
        manifest["anexos_subidos"] = sorted(set(previos + guardados))
        manifest["estado"] = "anexos_subidos"
        manifest["subido_en"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        _guardar_manifest(carpeta, manifest)
        _cache["t"] = 0
        return self._redirigir(f"/p/{partes[1]}?subido={len(guardados)}")


def main():
    os.makedirs(ENTREGABLES, exist_ok=True)
    print(f"LICITAIA PORTAL v2 · carpeta: {ENTREGABLES}")
    print(f"LICITAIA PORTAL v2 · proyectos indexados: {len(_proyectos())}")
    print(f"LICITAIA PORTAL v2 · escuchando en 0.0.0.0:{PUERTO}"
          + (f" · URL base: {URL_BASE}" if URL_BASE else ""))
    servidor = ThreadingHTTPServer(("0.0.0.0", PUERTO), Manejador)
    try:
        servidor.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
