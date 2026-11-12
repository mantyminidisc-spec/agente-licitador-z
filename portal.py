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
def _version_agente():
    """Lee la versión del agente desde el archivo del contenedor."""
    import re as _re
    for ruta in ("/app/agente_licitador_z.py",
                 os.path.join(os.path.dirname(os.path.abspath(__file__)), "agente_licitador_z.py")):
        try:
            with open(ruta, encoding="utf-8") as f:
                for linea in f:
                    if linea.startswith("VERSION_LICITAIA"):
                        return linea.split("=")[-1].strip().strip('\"').strip(chr(39))
        except Exception:
            continue
    return "desconocida"


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
        if nombre.lower().endswith((".json", ".log", ".tmp", ".lock", ".html", ".eml")):
            continue
        if "CORREO_PENDIENTE" in nombre.upper():
            continue
        if nombre.upper().startswith("PLANILLA_SIMULADA"):
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


def _ruta_cotizacion(carpeta):
    return os.path.join(carpeta, "cotizacion_editada.json")


def _leer_cotizacion(carpeta):
    try:
        with open(_ruta_cotizacion(carpeta), encoding="utf-8") as f:
            return json.load(f) or {}
    except Exception:
        return {}


def _guardar_cotizacion(carpeta, datos):
    ruta = _ruta_cotizacion(carpeta)
    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False, indent=2)
    return ruta


def _solo_digitos(texto, defecto=0):
    """Convierte '5.200.000', '5,200,000' o '$5.200.000' en 5200000."""
    try:
        limpio = "".join(c for c in str(texto) if c.isdigit())
        return int(limpio) if limpio else int(defecto)
    except Exception:
        return int(defecto)


def _monto_recomendado(proyecto):
    """Monto con IVA que el cliente eligió para la cotización."""
    try:
        cot = _leer_cotizacion(proyecto["_carpeta"])
        elegido = str(cot.get("escenario_cotizacion") or "recomendado").lower()
        if elegido == "mi_precio":
            costo_cot = _solo_digitos(cot.get("costo") or 0)
            try:
                margen_cot = float(cot.get("margen_objetivo") or 0.35)
            except Exception:
                margen_cot = 0.35
            if margen_cot > 1:
                margen_cot = margen_cot / 100.0
            neto_preferido = _solo_digitos(cot.get("neto") or 0)
            if neto_preferido > 0:
                iva_txt3 = cot.get("iva") or 0.19
                try:
                    iva_v3 = float(iva_txt3)
                except Exception:
                    iva_v3 = 0.19
                if iva_v3 > 1:
                    iva_v3 = iva_v3 / 100.0
                return int(round(neto_preferido * (1 + iva_v3)))
            if costo_cot > 0 and 0 < margen_cot < 1:
                neto_calc = int(round(costo_cot / (1 - margen_cot)))
                iva_txt2 = cot.get("iva") or 0.19
                try:
                    iva_v2 = float(iva_txt2)
                except Exception:
                    iva_v2 = 0.19
                if iva_v2 > 1:
                    iva_v2 = iva_v2 / 100.0
                return int(round(neto_calc * (1 + iva_v2)))
            neto = _solo_digitos(cot.get("neto") or 0)
            iva_txt = cot.get("iva") or 0.19
            try:
                iva_v = float(iva_txt)
            except Exception:
                iva_v = 0.19
            if iva_v > 1:
                iva_v = iva_v / 100.0
            if neto:
                return int(round(neto * (1 + iva_v)))
        monto_elegido = cot.get(f"monto_{elegido}")
        if monto_elegido:
            return _solo_digitos(monto_elegido)
        if cot.get("monto"):
            return _solo_digitos(cot.get("monto"))
        recomendado = _recomendado_del_estudio(proyecto["_carpeta"])
        if recomendado > 0:
            return recomendado
        valores = _valores_planilla(proyecto)
        return int(round(int(valores.get("presupuesto") or 0) * float(valores.get("pct2") or 0.8)))
    except Exception:
        return int(proyecto.get("oferta_bruta_clp") or 0)


def _sensibilidad(presupuesto, iva, costo, desde=0.70, hasta=0.99, paso=0.02,
                  pct_competencia=0.85, peso_precio=0.40):
    """Tabla de sensibilidad igual a la del Excel: puntaje y probabilidad por precio."""
    import math
    filas = []
    desde = max(0.50, float(desde or 0.70))
    hasta = min(0.99, max(desde + 0.02, float(hasta or 0.99)))
    pct = desde
    while pct <= hasta + 0.0001:
        bruto = int(round(presupuesto * pct))
        neto = int(round(bruto / (1 + iva))) if (1 + iva) else 0
        ganancia = neto - costo
        margen = (ganancia / neto * 100) if neto else 0
        if pct <= 0:
            pct = 0.01
        puntaje = (peso_precio * 100) * min(1.0, pct_competencia / pct) + (1 - peso_precio) * 100
        x = (pct_competencia - pct) / 0.035
        try:
            prob = 1.0 / (1.0 + math.exp(-x))
        except Exception:
            prob = 0.0
        filas.append({"pct": pct, "bruto": bruto, "neto": neto, "ganancia": ganancia,
                      "margen": margen, "puntaje": puntaje, "prob": prob * 100})
        pct = round(pct + paso, 4)
    return filas


def _pcts_coherentes(p1, p2, p3):
    """Los tres escenarios deben ir en orden ascendente y dentro de rango razonable."""
    try:
        valores = [float(p1), float(p2), float(p3)]
    except Exception:
        return False
    if not all(0.50 <= v <= 0.99 for v in valores):
        return False
    if not (valores[0] < valores[1] < valores[2]):
        return False
    if (valores[2] - valores[0]) < 0.02:
        return False
    return True


def _pct_seguro(valor, defecto):
    """Devuelve una fracción válida; si el dato guardado es basura, usa el valor del documento."""
    try:
        v = float(valor)
    except Exception:
        return round(float(defecto), 4)
    if v > 1.5:
        v = v / 100.0
    if v > 1.0 or v < 0.40:
        return round(float(defecto), 4)
    return round(v, 4)


def _pcts_del_estudio(carpeta):
    """Lee los porcentajes de tope del estudio financiero del agente."""
    try:
        datos = _leer_datos_informe(carpeta)
        estudio = (datos.get("informes") or {}).get("estudio_financiero_escenarios") or {}
    except Exception:
        return {}
    claves = ["escenario_1_competitivo_agresivo",
              "escenario_2_equilibrado_recomendado",
              "escenario_3_premium_alta_gama"]
    salida = {}
    for pos, clave in enumerate(claves, 1):
        bloque = estudio.get(clave)
        if not isinstance(bloque, dict):
            continue
        try:
            pct = float(bloque.get("porcentaje_tope") or 0)
        except Exception:
            continue
        if pct > 1.5:
            pct = pct / 100.0
        if 0.50 <= pct <= 1.0:
            salida["pct%d" % pos] = round(pct, 4)
            total = bloque.get("total_bruto_clp")
            if total:
                salida["bruto%d" % pos] = _solo_digitos(total)
    return salida


def _recomendado_del_estudio(carpeta):
    """Monto bruto del escenario recomendado según el estudio del agente."""
    try:
        datos = _leer_datos_informe(carpeta)
        estudio = (datos.get("informes") or {}).get("estudio_financiero_escenarios") or {}
        bloque = estudio.get("escenario_2_equilibrado_recomendado") or {}
        monto = _solo_digitos(bloque.get("total_bruto_clp") or 0)
        return monto
    except Exception:
        return 0


def _valores_planilla(proyecto):
    """Datos de entrada de la planilla: ficha + lo que el cliente ya editó."""
    cot = _leer_cotizacion(proyecto["_carpeta"])
    try:
        presupuesto = int(cot.get("presupuesto") or proyecto.get("presupuesto_clp") or 0)
    except Exception:
        presupuesto = int(proyecto.get("presupuesto_clp") or 0)
    try:
        iva_guardado = float(cot.get("iva") or 0.19)
    except Exception:
        iva_guardado = 0.19
    if iva_guardado > 1:
        iva_guardado = iva_guardado / 100.0
    oferta = int(cot.get("monto") or proyecto.get("oferta_bruta_clp") or 0)
    base = (oferta / presupuesto) if presupuesto else 0.80
    if base > 1.0 or base < 0.50:
        base = 0.80  # el monto guardado no es coherente: se usa el escenario estándar
    horas = _solo_digitos(cot.get("horas") or 0)
    valor_hora = _solo_digitos(cot.get("valor_hora") or 0)
    cloud = _solo_digitos(cot.get("cloud") or 0)
    gestion = _solo_digitos(cot.get("gestion") or 0)
    costo_desglose = horas * valor_hora + cloud + gestion
    try:
        costo_guardado = _solo_digitos(cot.get("costo") or 0)
    except Exception:
        costo_guardado = 0
    # v28.3: los costos parten en cero; el cliente los completa
    if costo_desglose > 0:
        costo_guardado = costo_desglose
    else:
        costo_guardado = 0
        horas = horas or 0
        valor_hora = valor_hora or 0
        cloud = cloud or 0
        gestion = gestion or 0
    defecto_1 = min(0.98, max(0.50, round(base - 0.02, 2)))
    defecto_2 = min(0.99, max(0.55, round(base, 2)))
    defecto_3 = min(0.99, max(0.60, round(base + 0.05, 2)))

    # si el agente dejó un estudio financiero, esas son las referencias válidas
    del_estudio = _pcts_del_estudio(proyecto["_carpeta"])
    if del_estudio:
        defecto_1 = del_estudio.get("pct1", defecto_1)
        defecto_2 = del_estudio.get("pct2", defecto_2)
        defecto_3 = del_estudio.get("pct3", defecto_3)

    p1 = _pct_seguro(cot.get("pct1"), defecto_1)
    p2 = _pct_seguro(cot.get("pct2"), defecto_2)
    p3 = _pct_seguro(cot.get("pct3"), defecto_3)
    if del_estudio:
        # el estudio del agente tiene prioridad sobre lo guardado
        p1 = defecto_1
        p2 = defecto_2
        p3 = defecto_3
    if not _pcts_coherentes(p1, p2, p3):
        print(f"   [PORTAL] escenarios incoherentes ({p1}/{p2}/{p3}): "
              f"se usan los del documento {defecto_1}/{defecto_2}/{defecto_3}")
        p1, p2, p3 = defecto_1, defecto_2, defecto_3

    try:
        neto_guardado = _solo_digitos(cot.get("neto") or 0)
    except Exception:
        neto_guardado = 0
    try:
        margen_obj = float(cot.get("margen_objetivo") or 0.35)
    except Exception:
        margen_obj = 0.35
    if margen_obj > 1:
        margen_obj = margen_obj / 100.0
    if not (0 < margen_obj < 1):
        margen_obj = 0.35

    return {
        "presupuesto": presupuesto,
        "iva": iva_guardado,
        "costo": costo_guardado,
        "horas": horas,
        "valor_hora": valor_hora,
        "cloud": cloud,
        "gestion": gestion,
        "neto": neto_guardado,
        "margen_objetivo": margen_obj,
        "pct1": p1,
        "pct2": p2,
        "pct3": p3,
        "escenario_cotizacion": str(cot.get("escenario_cotizacion") or "recomendado").lower(),
    }


def _valores_simulador(proyecto):
    cot = _leer_cotizacion(proyecto["_carpeta"])
    presupuesto = int(proyecto.get("presupuesto_clp") or 0)
    oferta = int(cot.get("monto") or proyecto.get("oferta_bruta_clp") or 0)
    costo = int(cot.get("costo") or (oferta * 0.42) or 0)
    base = (oferta / presupuesto) if presupuesto else 0.80
    return {
        "presupuesto": presupuesto,
        "iva": 0.19,
        "costo": costo,
        "pct1": round(max(0.70, base - 0.02), 2),
        "pct2": round(base, 2),
        "pct3": round(min(0.99, base + 0.05), 2),
    }


def _leer_datos_informe(carpeta):
    try:
        with open(os.path.join(carpeta, "datos_informe.json"), encoding="utf-8") as f:
            return json.load(f) or {}
    except Exception:
        return {}


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
        cotizacion=_leer_cotizacion(carpeta),
        # el KPI de oferta se calcula igual que la cotización
        oferta_vigente=_monto_recomendado(proyecto),
        datos_informe=_leer_datos_informe(carpeta),
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
        try:
            return self._do_get_interno()
        except Exception as e:
            import traceback
            traceback.print_exc()
            self._enviar(500, f"<h1>Error en la vista</h1><pre>{e}</pre>")

    def _do_get_interno(self):
        ruta = self.path.split("?")[0]
        consulta = self.path.split("?")[1] if "?" in self.path else ""
        partes = [p for p in ruta.split("/") if p]

        if ruta in ("/salud", "/"):
            datos = {"servicio": "licitaia-portal", "estado": "ok", "version": "2.1",
                     "version_agente": _version_agente(),
                     "ia_principal": os.getenv("IA_MODELO", ""),
                     "ia_premium": os.getenv("IA_PREMIUM_MODELO", ""),
                     "proyectos": len(_proyectos()), "carpeta": ENTREGABLES}
            return self._enviar(200, json.dumps(datos, ensure_ascii=False), "application/json")

        if len(partes) >= 2 and partes[0] == "p":
            token = partes[1]
            proyecto = self._buscar(token)
            if not proyecto:
                return self._enviar(404, "<h1>Enlace no válido</h1>"
                                         "<p>El proyecto no existe o el enlace fue revocado.</p>")

            if len(partes) == 4 and partes[2] == "docu":
                if portal_ui is None:
                    return self._enviar(500, "<h1>UI no disponible</h1>")
                datos = _leer_datos_informe(proyecto["_carpeta"])
                nombre_doc = os.path.basename(unquote(partes[3]))
                titulo = portal_ui.nombre_legible(nombre_doc)
                if nombre_doc.upper().startswith("DOC3"):
                    valores_fin = _valores_planilla(proyecto)
                    costo_fin = int(valores_fin.get("costo") or 0)
                    filas_fin = portal_ui.calcular_escenarios(
                        valores_fin["presupuesto"], valores_fin["iva"], costo_fin,
                        [(u"Competitivo", valores_fin["pct1"]),
                         (u"Recomendado", valores_fin["pct2"]),
                         (u"Conservador", valores_fin["pct3"])])
                    return self._enviar(200, portal_ui.informe_financiero_html(
                        proyecto, valores_fin, filas_fin,
                        elegido=str(valores_fin.get("escenario_cotizacion") or "recomendado"),
                        monto_cotizacion=_monto_recomendado(proyecto),
                        logo_url=f"/p/{partes[1]}/logo",
                        hay_logo=os.path.isfile(RUTA_LOGO),
                        estudio=(datos.get("informes") or {}).get(
                            "estudio_financiero_escenarios")), "text/html; charset=utf-8")
                html_doc = portal_ui.documento_html(
                    titulo, datos,
                    logo_url=f"/p/{partes[1]}/logo",
                    hay_logo=os.path.isfile(RUTA_LOGO),
                    meta_extra=str(proyecto.get("organismo") or ""))
                return self._enviar(200, html_doc, "text/html; charset=utf-8")

            if len(partes) == 3 and partes[2] == "planilla":
                if portal_ui is None:
                    return self._enviar(500, "<h1>UI no disponible</h1>")
                valores = _valores_planilla(proyecto)
                costo_desglose_get = (valores["horas"] * valores["valor_hora"]
                                      + valores["cloud"] + valores["gestion"])
                costo = costo_desglose_get if costo_desglose_get > 0 else int(
                    valores.get("costo") or 0)
                valores["costo"] = costo
                filas = portal_ui.calcular_escenarios(
                    valores["presupuesto"], valores["iva"], costo,
                    [(u"Competitivo", valores["pct1"]),
                     (u"Recomendado", valores["pct2"]),
                     (u"Conservador", valores["pct3"])])
                sens = _sensibilidad(valores["presupuesto"], valores["iva"], costo,
                                     valores["pct1"] - 0.08, min(0.99, valores["pct3"] + 0.10))
                enlace = (f"/p/{partes[1]}/f/Planilla_Simulada.xlsx"
                          if os.path.isfile(os.path.join(proyecto["_carpeta"],
                                                         "Planilla_Simulada.xlsx")) else "")
                return self._enviar(200, portal_ui.planilla_html(
                    proyecto, partes[1], valores, filas, enlace_excel=enlace,
                    sensibilidad=sens), "text/html; charset=utf-8")

            if len(partes) == 3 and partes[2] == "simulador":
                if portal_ui is None:
                    return self._enviar(500, "<h1>UI no disponible</h1>")
                valores = _valores_simulador(proyecto)
                return self._enviar(200, portal_ui.simulador_html(
                    proyecto, partes[1], valores), "text/html; charset=utf-8")

            if len(partes) == 3 and partes[2] == "cotizacion":
                datos = _leer_cotizacion(proyecto["_carpeta"])
                datos["monto"] = _monto_recomendado(proyecto)
                if portal_ui is None:
                    return self._enviar(500, "<h1>UI no disponible</h1>")
                html_cot = portal_ui.cotizacion_html(
                    proyecto, partes[1], datos,
                    hay_logo=os.path.isfile(RUTA_LOGO),
                    logo_url=f"/p/{partes[1]}/logo",
                    cot=datos)
                return self._enviar(200, html_cot, "text/html; charset=utf-8")

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
        try:
            return self._do_post_interno()
        except Exception as e:
            import traceback
            traceback.print_exc()
            self._enviar(500, f"<h1>Error al procesar</h1><pre>{e}</pre>")

    def _do_post_interno(self):
        ruta = self.path.split("?")[0]
        partes = [p for p in ruta.split("/") if p]
        if len(partes) != 3 or partes[0] != "p" or partes[2] not in ("subir", "reprocesar", "logo", "cotizacion", "simular", "planilla"):
            return self._enviar(404, "<h1>No encontrado</h1>")
        proyecto = self._buscar(partes[1])
        if not proyecto:
            return self._enviar(404, "<h1>Enlace no válido</h1>")

        carpeta = proyecto["_carpeta"]
        codigo = proyecto["_codigo"]

        if partes[2] == "planilla":
            import re as _re_reset
            if _re_reset.search(r"(^|&)reset=1", self.path.split("?")[1] if "?" in self.path else ""):
                try:
                    ruta_cot = _ruta_cotizacion(carpeta)
                    if os.path.isfile(ruta_cot):
                        os.remove(ruta_cot)
                    print(f"   [PORTAL] planilla {codigo}: valores restablecidos del documento")
                except Exception as e_reset:
                    print(f"   [PORTAL] no se pudo restablecer: {e_reset}")
                return self._redirigir(f"/p/{partes[1]}/planilla")

            largo = int(self.headers.get("Content-Length") or 0)
            if largo <= 0 or largo > 512 * 1024:
                return self._enviar(413, "<h1>Datos inválidos</h1>")
            from urllib.parse import parse_qs
            campos = parse_qs(self.rfile.read(largo).decode("utf-8", "ignore"))

            def _num(clave, defecto=0):
                crudo = (campos.get(clave) or [""])[0]
                try:
                    return float(str(crudo).replace(".", "").replace(",", ".") or defecto)
                except Exception:
                    return float(defecto)

            def _pct(clave, defecto):
                v = _num(clave, defecto * 100) or (defecto * 100)
                return v / 100.0 if v > 1.5 else v

            # boton "poner costos en cero"
            accion_previa = str((campos.get("accion") or [""])[0]).lower()
            if accion_previa == "cero_costos":
                try:
                    cot_cero = _leer_cotizacion(carpeta)
                    cot_cero.update({"horas": 0, "valor_hora": 0, "cloud": 0,
                                     "gestion": 0, "costo": 0})
                    _guardar_cotizacion(carpeta, cot_cero)
                    print(f"   [PORTAL] planilla {codigo}: costos en cero")
                except Exception as e_cero:
                    print(f"   [PORTAL] no se pudo limpiar costos: {e_cero}")
                return self._redirigir(f"/p/{partes[1]}/planilla")

            crudos = {c: _num(c, 0) for c in ("pct1", "pct2", "pct3")}
            montos_libres = [_solo_digitos((campos.get(f"monto{i}") or [""])[0], 0)
                             for i in (1, 2, 3)]
            anteriores = _valores_planilla(proyecto)
            for c, defecto in (("pct1", anteriores["pct1"]), ("pct2", anteriores["pct2"]),
                               ("pct3", anteriores["pct3"])):
                if crudos[c] <= 0:
                    crudos[c] = defecto * 100
            invalidos = [c for c, v in crudos.items() if not (1 <= v <= 100)]
            if not invalidos:
                _v = [crudos["pct1"] / 100.0, crudos["pct2"] / 100.0, crudos["pct3"] / 100.0]
                if not _pcts_coherentes(*_v):
                    invalidos = ["orden"]
            if invalidos:
                valores_prev = _valores_planilla(proyecto)
                costo_prev = (valores_prev["horas"] * valores_prev["valor_hora"]
                              + valores_prev["cloud"] + valores_prev["gestion"])
                filas_prev = portal_ui.calcular_escenarios(
                    valores_prev["presupuesto"], valores_prev["iva"], costo_prev,
                    [(u"Competitivo", valores_prev["pct1"]),
                     (u"Recomendado", valores_prev["pct2"]),
                     (u"Conservador", valores_prev["pct3"])])
                if invalidos == ["orden"]:
                    aviso = ("Los escenarios deben ir en orden ascendente y con al menos 3 puntos "
                             "de diferencia (por ejemplo 78, 80, 85). No se guardó nada.")
                else:
                    aviso = ("El porcentaje debe estar entre 1 y 100. Corrige los campos marcados "
                             "y vuelve a guardar (no se guardó nada).")
                return self._enviar(200, portal_ui.planilla_html(
                    proyecto, partes[1], valores_prev, filas_prev, mensaje=aviso,
                    sensibilidad=_sensibilidad(valores_prev["presupuesto"], valores_prev["iva"],
                                               costo_prev, valores_prev["pct1"] - 0.08,
                                               min(0.99, valores_prev["pct3"] + 0.10))),
                    "text/html; charset=utf-8")

            anteriores_post = _valores_planilla(proyecto)
            valores = {
                "presupuesto": _solo_digitos((campos.get("presupuesto") or ["0"])[0],
                                             anteriores_post["presupuesto"]),
                "iva": (lambda v: v / 100.0 if v > 1 else v)(_num("iva", 19) or 19),
                "costo": _solo_digitos((campos.get("costo") or ["0"])[0],
                                       anteriores_post["costo"]),
                "horas": _solo_digitos((campos.get("horas") or ["0"])[0],
                                       anteriores_post.get("horas") or 0),
                "valor_hora": _solo_digitos((campos.get("valor_hora") or ["0"])[0],
                                            anteriores_post.get("valor_hora") or 0),
                "cloud": _solo_digitos((campos.get("cloud") or ["0"])[0],
                                       anteriores_post.get("cloud") or 0),
                "gestion": _solo_digitos((campos.get("gestion") or ["0"])[0],
                                         anteriores_post.get("gestion") or 0),
                "neto": _solo_digitos((campos.get("neto") or ["0"])[0],
                                      anteriores_post.get("neto") or 0),
                "margen_objetivo": (lambda v: v / 100.0 if v > 1 else v)(
                    _num("margen_objetivo", 35) or 35),
                "pct1": anteriores_post["pct1"],
                "pct2": anteriores_post["pct2"],
                "pct3": anteriores_post["pct3"],
            }
            costo_desglose_post = (valores["horas"] * valores["valor_hora"]
                                   + valores["cloud"] + valores["gestion"])
            if costo_desglose_post > 0:
                costo = costo_desglose_post
            else:
                costo = _solo_digitos((campos.get("costo") or ["0"])[0],
                                      anteriores_post["costo"]) or anteriores_post["costo"]
            # los escenarios salen del presupuesto por los porcentajes del motor
            if False and any(m > 0 for m in montos_libres):
                filas = []
                for pos, (nombre_esc, monto_bruto) in enumerate(
                        zip((u"Competitivo", u"Recomendado", u"Conservador"), montos_libres)):
                    bruto = monto_bruto or int(round(valores["presupuesto"]
                                                     * (0.78, 0.80, 0.85)[pos]))
                    neto = int(round(bruto / (1 + valores["iva"]))) if (1 + valores["iva"]) else 0
                    iva_m = bruto - neto
                    ganancia = neto - costo
                    margen = (ganancia / neto * 100) if neto else 0
                    roi = (ganancia / costo * 100) if costo else 0
                    pct = (bruto / valores["presupuesto"]) if valores["presupuesto"] else 0
                    filas.append({"escenario": nombre_esc, "pct": pct, "bruto": bruto,
                                  "neto": neto, "iva": iva_m, "costo": costo,
                                  "ganancia": ganancia, "margen": margen, "roi": roi})
                for pos, clave in enumerate(("pct1", "pct2", "pct3")):
                    valores[clave] = round(filas[pos]["pct"], 4)
            else:
                filas = portal_ui.calcular_escenarios(
                    valores["presupuesto"], valores["iva"], costo,
                    [(u"Competitivo", valores["pct1"]),
                     (u"Recomendado", valores["pct2"]),
                     (u"Conservador", valores["pct3"])])
            try:
                destino = os.path.join(carpeta, "Planilla_Simulada.xlsx")
                portal_ui.escribir_xlsx_simulado(destino, {
                    "presupuesto": valores["presupuesto"],
                    "iva": valores["iva"],
                    "horas": valores["horas"],
                    "valor_hora": valores["valor_hora"],
                    "cloud": valores["cloud"],
                    "gestion": valores["gestion"],
                    "costo": costo,
                    "costo": costo}, filas)
                # la cotización toma el monto del escenario recomendado
                cot = _leer_cotizacion(carpeta)
                accion = str((campos.get("accion") or [""])[0]).lower()
                if accion == "mi_precio":
                    elegido = "mi_precio"
                elif accion == "opcion":
                    elegido = str((campos.get("escenario_cotizacion")
                                   or ["recomendado"])[0]).lower()
                else:
                    elegido = str((campos.get("escenario_cotizacion")
                                   or ["recomendado"])[0]).lower()
                if elegido not in ("competitivo", "recomendado", "conservador", "mi_precio"):
                    elegido = "recomendado"
                if elegido == "mi_precio":
                    margen_v = float(valores.get("margen_objetivo") or 0.35)
                    neto_escrito = int(valores.get("neto") or 0)
                    if neto_escrito > 0:
                        neto_final = neto_escrito
                    elif 0 < margen_v < 1 and costo > 0:
                        neto_final = int(round(costo / (1 - margen_v)))
                    else:
                        neto_final = costo
                    valores["neto"] = neto_final
                    monto_final = int(round(neto_final * (1 + valores["iva"])))
                else:
                    indice = {"competitivo": 0, "recomendado": 1, "conservador": 2}[elegido]
                    monto_final = filas[indice]["bruto"]
                cot.update({
                    "monto": monto_final,
                    "neto": valores["neto"],
                    "margen_objetivo": valores["margen_objetivo"],
                    "monto_competitivo": filas[0]["bruto"],
                    "monto_recomendado": filas[1]["bruto"],
                    "monto_conservador": filas[2]["bruto"],
                    "escenario_cotizacion": elegido,
                    "presupuesto": valores["presupuesto"],
                    "iva": valores["iva"],
                    "pct1": valores["pct1"],
                    "pct2": valores["pct2"],
                    "pct3": valores["pct3"],
                    "costo": costo,
                })
                _guardar_cotizacion(carpeta, cot)
                print(f"   [PORTAL] planilla {codigo}: monto recomendado "
                      f"{filas[1]['bruto']} y cotizacion actualizada")
            except Exception as e_plan:
                print(f"   [PORTAL] error guardando la planilla: {e_plan}")
            sens = _sensibilidad(valores["presupuesto"], valores["iva"], costo,
                                 valores["pct1"] - 0.08, min(0.99, valores["pct3"] + 0.10))
            html_plan = portal_ui.planilla_html(
                proyecto, partes[1], valores, filas,
                mensaje="Planilla recalculada. La cotización formal ya tiene el nuevo monto.",
                enlace_excel=f"/p/{partes[1]}/f/Planilla_Simulada.xlsx",
                sensibilidad=sens)
            return self._enviar(200, html_plan, "text/html; charset=utf-8")

        if partes[2] == "simular":
            largo = int(self.headers.get("Content-Length") or 0)
            if largo <= 0 or largo > 512 * 1024:
                return self._enviar(413, "<h1>Datos inválidos</h1>")
            from urllib.parse import parse_qs
            campos = parse_qs(self.rfile.read(largo).decode("utf-8", "ignore"))
            valores = {}
            for clave in ("presupuesto", "iva", "costo", "pct1", "pct2", "pct3"):
                crudo = (campos.get(clave) or [""])[0]
                try:
                    valores[clave] = float(str(crudo).replace(".", "").replace(",", ".") or 0)
                except Exception:
                    valores[clave] = 0
            for clave in ("pct1", "pct2", "pct3"):
                if valores[clave] > 1.5:
                    valores[clave] = valores[clave] / 100.0
            if valores.get("iva", 0) > 1:
                valores["iva"] = valores["iva"] / 100.0
            pcts = [(u"Escenario 1", valores.get("pct1") or 0.78),
                    (u"Escenario 2", valores.get("pct2") or 0.80),
                    (u"Escenario 3", valores.get("pct3") or 0.85)]
            filas = portal_ui.calcular_escenarios(
                valores.get("presupuesto"), valores.get("iva") or 0.19,
                valores.get("costo"), pcts)
            destino = os.path.join(carpeta, "Planilla_Simulada.xlsx")
            try:
                portal_ui.escribir_xlsx_simulado(destino, valores, filas)
                print(f"   [PORTAL] planilla simulada {codigo}: "
                      f"{int(valores.get('presupuesto') or 0)}")
            except Exception as e_xls:
                print(f"   [PORTAL] no se pudo escribir la planilla: {e_xls}")
            enlace = f"/p/{partes[1]}/f/Planilla_Simulada.xlsx"
            html_sim = portal_ui.simulador_html(
                proyecto, partes[1], valores, filas,
                mensaje="Simulación recalculada y Excel generado.",
                enlace_excel=enlace)
            return self._enviar(200, html_sim, "text/html; charset=utf-8")

        if partes[2] == "cotizacion":
            largo = int(self.headers.get("Content-Length") or 0)
            if largo <= 0 or largo > 1024 * 1024:
                return self._enviar(413, "<h1>Datos demasiado grandes</h1>")
            cuerpo = self.rfile.read(largo).decode("utf-8", "ignore")
            from urllib.parse import parse_qs
            campos = parse_qs(cuerpo)
            datos = {clave: (valores[0] if valores else "")
                     for clave, valores in campos.items()}
            try:
                datos["monto"] = int(str(datos.get("monto") or "0").replace(".", "").strip() or 0)
            except Exception:
                datos["monto"] = 0
            try:
                _guardar_cotizacion(carpeta, datos)
                print(f"   [PORTAL] cotización editada {codigo}: monto {datos.get('monto')}")
            except Exception as e:
                print(f"   [PORTAL] error guardando cotización: {e}")
                return self._enviar(500, "<h1>No se pudo guardar</h1>")
            return self._redirigir(f"/p/{partes[1]}#cotizacion")

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
                # Regenerar los documentos para que salgan con el logo nuevo
                try:
                    ok_repro, msg_repro = _lanzar_reproceso(codigo)
                    print(f"   [PORTAL] regenerando con el logo nuevo: {msg_repro}")
                except Exception as e_repro:
                    print(f"   [PORTAL] no se pudo regenerar: {str(e_repro)[:70]}")
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
