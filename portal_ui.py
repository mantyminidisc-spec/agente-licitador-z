# -*- coding: utf-8 -*-
"""
LICITAIA · Panel del proyecto (mini ERP)
===============================================================================
· Menú lateral izquierdo con los documentos y secciones
· Vista previa dentro del panel: PDF nativo, DOCX y XLSX convertidos a HTML
· Cotización editable (formulario) con versión imprimible para guardar en PDF
Sin dependencias externas ni servicios de terceros.
===============================================================================
"""
import html
import json
import os
from datetime import datetime

TIPOS_PREVIA = (".pdf", ".docx", ".xlsx", ".txt", ".csv")

NOMBRES = (
    ("DOC1", "Informe de auditoría"),
    ("DOC2", "Cotización formal"),
    ("DOC3", "Informe financiero"),
    ("PLANILLA", "Planilla financiera"),
    ("DOC4", "Plan de contrato"),
)


def _esc(t):
    return html.escape(str(t or ""), quote=True)


def _clp(v):
    try:
        return "$" + f"{int(v):,}".replace(",", ".")
    except Exception:
        return "$0"


def _peso(b):
    try:
        kb = int(b) / 1024
        return f"{kb/1024:.1f} MB" if kb > 1024 else f"{kb:.0f} KB"
    except Exception:
        return ""


def _fecha(t):
    try:
        v = datetime.strptime(str(t)[:16], "%Y-%m-%d %H:%M")
        meses = ("ene", "feb", "mar", "abr", "may", "jun",
                 "jul", "ago", "sep", "oct", "nov", "dic")
        return f"{v.day} {meses[v.month-1]} {v.year} · {v:%H:%M}"
    except Exception:
        return str(t or "por definir")


def nombre_legible(nombre):
    alto = nombre.upper()
    for prefijo, etiqueta in NOMBRES:
        if alto.startswith(prefijo):
            return etiqueta
    if alto.startswith("PLANILLA"):
        return "Planilla financiera"
    return nombre


def es_previble(nombre):
    return str(nombre).lower().endswith(TIPOS_PREVIA)


# ---------------------------------------------------------------------------
# Conversión de documentos a HTML
# ---------------------------------------------------------------------------
def docx_a_html(ruta):
    """Convierte un Word a HTML respetando el orden real: párrafos, tablas e imágenes."""
    try:
        from docx import Document
        from docx.table import Table
        from docx.text.paragraph import Paragraph
        from docx.oxml.ns import qn
    except Exception as e:
        return f"<p>No se pudo abrir el Word: {_esc(e)[:120]}</p>"
    try:
        documento = Document(ruta)
    except Exception as e:
        return f"<p>No se pudo abrir el Word: {_esc(e)[:120]}</p>"

    def imagen_de_run(tramo):
        """Devuelve un <img> en base64 si el tramo tiene una imagen incrustada."""
        try:
            blips = tramo._element.findall(".//" + qn("a:blip"))
            for blip in blips:
                rid = blip.get(qn("r:embed"))
                if not rid:
                    continue
                parte = documento.part.related_parts.get(rid)
                if not parte:
                    continue
                tipo = str(getattr(parte, "content_type", ""))
                if "emf" in tipo or "wmf" in tipo:
                    continue
                import base64
                datos = base64.b64encode(parte.blob).decode("ascii")
                return (f'<img class="doc-img" src="data:{tipo or "image/png"};base64,{datos}" '
                        f'alt="imagen">')
        except Exception:
            return ""
        return ""

    def parrafo_a_html(parrafo):
        texto = parrafo.text.strip()
        imagenes = "".join(imagen_de_run(r) for r in parrafo.runs)
        if not texto and not imagenes:
            return ""
        estilo = (parrafo.style.name or "").lower()
        if texto and (estilo.startswith("heading") or estilo.startswith("título")):
            nivel = "h3" if ("1" in estilo or "2" in estilo) else "h4"
            return f"<{nivel}>{_esc(texto)}</{nivel}>{imagenes}"
        marcado = ""
        for tramo in parrafo.runs:
            if not tramo.text:
                continue
            trozo = _esc(tramo.text)
            if tramo.bold:
                trozo = f"<strong>{trozo}</strong>"
            if tramo.italic:
                trozo = f"<em>{trozo}</em>"
            marcado += trozo
        alineado = ""
        try:
            if parrafo.alignment is not None and str(parrafo.alignment).startswith("CENTER"):
                alineado = ' class="centrado"'
        except Exception:
            pass
        cuerpo = f"<p{alignedo}>{marcado}</p>" if marcado else ""
        return f"{imagenes}{cuerpo}"

    def tabla_a_html(tabla):
        filas = []
        for pos, fila in enumerate(tabla.rows):
            etiqueta = "th" if pos == 0 else "td"
            celdas = []
            for celda in fila.cells:
                contenido = "".join(parrafo_a_html(p) for p in celda.paragraphs).strip()
                celdas.append(f"<{etiqueta}>{contenido or '&nbsp;'}</{etiqueta}>")
            filas.append("<tr>" + "".join(celdas) + "</tr>")
        return f'<div class="scroll-x"><table class="tabla-doc">{"".join(filas)}</table></div>'

    partes = []
    try:
        for elemento in documento.element.body.iterchildren():
            etiqueta = elemento.tag.split("}")[-1]
            if etiqueta == "p":
                html_par = parrafo_a_html(Paragraph(elemento, documento))
                if html_par:
                    partes.append(html_par)
            elif etiqueta == "tbl":
                partes.append(tabla_a_html(Table(elemento, documento)))
    except Exception as e:
        return f"<p>No se pudo convertir el documento: {_esc(e)[:120]}</p>"

    return "\n".join(partes) or "<p>El documento no tiene contenido.</p>"


def xlsx_a_html(ruta, max_filas=400):
    try:
        import openpyxl
        libro = openpyxl.load_workbook(ruta, data_only=True, read_only=True)
    except Exception as e:
        return f"<p>No se pudo abrir el Excel: {_esc(e)[:120]}</p>"
    partes = []
    for hoja in libro.worksheets:
        filas = []
        for numero, fila in enumerate(hoja.iter_rows(values_only=True), 1):
            if numero > max_filas:
                filas.append("<tr><td colspan='12'>… (vista truncada)</td></tr>")
                break
            if not any(v not in (None, "") for v in fila):
                continue
            etiqueta = "th" if numero == 1 else "td"
            celdas = "".join(f"<{etiqueta}>{_esc('' if v is None else v)}</{etiqueta}>"
                             for v in fila)
            filas.append(f"<tr>{celdas}</tr>")
        partes.append(f'<h4>{_esc(hoja.title)}</h4><div class="scroll-x">'
                      f'<table class="tabla-doc">{"".join(filas)}</table></div>')
    try:
        libro.close()
    except Exception:
        pass
    return "\n".join(partes) or "<p>La planilla no tiene datos.</p>"


def vista_archivo(ruta, nombre):
    ext = os.path.splitext(nombre)[1].lower()
    if ext == ".pdf":
        return None, False
    if ext == ".docx":
        return docx_a_html(ruta), True
    if ext == ".xlsx":
        return xlsx_a_html(ruta), True
    if ext in (".txt", ".csv"):
        try:
            with open(ruta, encoding="utf-8", errors="ignore") as f:
                return f"<pre>{_esc(f.read()[:200000])}</pre>", True
        except Exception:
            return "<p>No se pudo leer el archivo.</p>", True
    return "<p>Formato sin vista previa. Descárgalo para abrirlo.</p>", True


# ---------------------------------------------------------------------------
# CSS compartido
# ---------------------------------------------------------------------------
CSS = """
:root{--azul:#12213a;--azul2:#1e3a5f;--texto:#1c2536;--gris:#5b6674;--linea:#e3e7ee;
--fondo:#f5f7fb;--blanco:#fff;--verde:#0f766e;--radio:14px}
*{box-sizing:border-box}
body{margin:0;background:var(--fondo);color:var(--texto);
font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;
font-size:15px;line-height:1.55}
a{color:var(--azul2)}
.shell{display:grid;grid-template-columns:250px 1fr;min-height:100vh}
aside{background:var(--azul);color:#fff;padding:16px 12px;position:sticky;top:0;height:100vh;
overflow-y:auto}
.marca{display:flex;align-items:center;gap:10px;padding:6px 6px 14px;border-bottom:1px solid #ffffff22;
margin-bottom:12px}
.marca img{height:36px;max-width:100px;object-fit:contain;background:#fff;border-radius:8px;padding:4px}
.marca .ini{width:36px;height:36px;border-radius:8px;background:#ffffff22;display:flex;
align-items:center;justify-content:center;font-weight:700}
.marca small{display:block;opacity:.7;font-size:11px}
.menu{display:flex;flex-direction:column;gap:2px}
.menu .titulo{font-size:10px;letter-spacing:.1em;text-transform:uppercase;opacity:.55;
padding:14px 8px 6px}
.menu a,.menu button{display:flex;align-items:center;gap:9px;width:100%;text-align:left;
background:transparent;border:0;color:#dbe3ee;padding:9px 10px;border-radius:9px;
font-size:13.5px;font-weight:550;cursor:pointer;text-decoration:none;font-family:inherit}
.menu a:hover,.menu button:hover{background:#ffffff14;color:#fff}
.menu .activo{background:#ffffff1f;color:#fff}
.menu .chip{margin-left:auto;background:#ffffff22;border-radius:20px;padding:0 7px;font-size:10.5px}
main{padding:22px 26px 60px;max-width:1000px}
.cab{display:flex;align-items:flex-start;gap:14px;flex-wrap:wrap;margin-bottom:16px}
.cab h1{font-size:20px;margin:0;line-height:1.25;color:var(--azul)}
.cab .sub{color:var(--gris);font-size:12.5px;margin-top:3px}
.cab .der{margin-left:auto;display:flex;gap:8px;flex-wrap:wrap}
.btn{display:inline-block;background:var(--azul);color:#fff;border:0;border-radius:10px;
padding:9px 16px;font-weight:650;font-size:13.5px;text-decoration:none;cursor:pointer;
font-family:inherit}
.btn-sec{background:#eef2f8;color:var(--azul2)}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin-bottom:16px}
.kpi{background:var(--blanco);border:1px solid var(--linea);border-radius:var(--radio);padding:12px 14px}
.kpi .et{font-size:10.5px;text-transform:uppercase;letter-spacing:.08em;color:var(--gris);font-weight:700}
.kpi .va{font-size:19px;font-weight:700;margin-top:2px;color:var(--azul)}
.kpi .no{font-size:11.5px;color:var(--gris)}
.caja{background:var(--blanco);border:1px solid var(--linea);border-radius:var(--radio);
margin-bottom:16px;overflow:hidden}
.caja>h2{font-size:14.5px;margin:0;padding:13px 16px;border-bottom:1px solid var(--linea);
display:flex;align-items:center;gap:10px}
.caja>h2 .der{margin-left:auto;display:flex;gap:8px}
.caja .cuerpo{padding:16px}
.oculto{display:none!important}
iframe.visor{width:100%;height:70vh;min-height:420px;border:0;display:block;background:#fff}
.scroll-x{overflow-x:auto}
.tabla-doc{border-collapse:collapse;font-size:13px;width:100%}
.tabla-doc caption{text-align:left;font-weight:700;font-size:12px;color:var(--gris);padding:8px 0}
.tabla-doc th,.tabla-doc td{border:1px solid var(--linea);padding:6px 8px;vertical-align:top}
.tabla-doc th{background:#f7f9fc;text-align:left;font-weight:650}
.doc-img{max-width:100%;height:auto;display:block;margin:8px 0}
.centrado{text-align:center}
.cuerpo-doc,iframe.visor{background:#fff}
.tabla-doc td{font-size:12.5px}
table.lista{width:100%;border-collapse:collapse;font-size:13.5px}
table.lista th{text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.06em;
color:var(--gris);padding:10px 8px;border-bottom:1px solid var(--linea)}
table.lista td{padding:10px 8px;border-bottom:1px solid #f1f3f7;word-break:break-word}
.der{text-align:right}
.btn-mini{display:inline-block;background:#eef2f8;color:var(--azul2);border-radius:8px;
padding:5px 11px;font-size:12.5px;font-weight:600;text-decoration:none;border:0;cursor:pointer}
.ayuda{color:var(--gris);font-size:13px;margin:0 0 10px}
.campo{display:flex;flex-direction:column;gap:4px;margin-bottom:12px}
.campo label{font-size:11.5px;text-transform:uppercase;letter-spacing:.06em;color:var(--gris);font-weight:700}
.campo input,.campo textarea,.campo select{border:1px solid var(--linea);border-radius:10px;
padding:9px 11px;font-size:14px;font-family:inherit;color:var(--texto);background:#fff;width:100%}
.campo textarea{min-height:74px;resize:vertical}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:0 16px}
.aviso{background:#ecfdf5;border:1px solid #a7f3d0;border-radius:12px;padding:10px 14px;
font-size:13px;color:#065f46;margin-bottom:14px}
.vacio{color:var(--gris);font-size:13px}
@media (max-width:860px){
 .shell{grid-template-columns:1fr}
 aside{position:static;height:auto}
 .menu{flex-direction:row;flex-wrap:wrap}
 .menu .titulo{display:none}
 main{padding:16px 14px 40px}
 .grid2{grid-template-columns:1fr}
 iframe.visor{height:60vh;min-height:320px}
}
"""

JS = """
(function(){
  var items = document.querySelectorAll('[data-seccion]');
  var paneles = document.querySelectorAll('.seccion');
  function activar(id){
    if(!document.getElementById('sec-'+id)) return;
    items.forEach(function(b){ b.classList.toggle('activo', b.dataset.seccion === id); });
    paneles.forEach(function(p){ p.classList.toggle('oculto', p.id !== 'sec-'+id); });
    if(history.replaceState) history.replaceState(null,'','#'+id);
  }
  items.forEach(function(b){ b.addEventListener('click', function(e){ e.preventDefault(); activar(b.dataset.seccion); }); });
  var inicial = (location.hash||'').replace('#','') || 'resumen';
  activar(inicial);
})();
"""


# ---------------------------------------------------------------------------
# Página principal
# ---------------------------------------------------------------------------
def pagina(proyecto, token, archivos, adjuntos, aviso="", hay_logo=True, logo_url="",
           doc1_marca=0, cotizacion=None):
    codigo = proyecto.get("codigo", "")
    nombre = str(proyecto.get("nombre") or "Proyecto")
    organismo = str(proyecto.get("organismo") or "")
    es_ca = proyecto.get("tipo_portal") == "COMPRA_AGIL"
    presupuesto = int(proyecto.get("presupuesto_clp") or 0)
    oferta = int(proyecto.get("oferta_bruta_clp") or 0)
    estado = str(proyecto.get("estado") or "")

    cot = dict(cotizacion or {})
    datos = {
        "razon_social": cot.get("razon_social") or "Mantytech SpA",
        "rut": cot.get("rut") or "78.479.059-2",
        "representante": cot.get("representante") or "Daniel Andrés Salazar Gavilán",
        "email": cot.get("email") or "contacto@mantytech.cl",
        "telefono": cot.get("telefono") or "+56 9 5090 1627",
        "monto": cot.get("monto") or oferta,
        "plazo": cot.get("plazo") or proyecto.get("plazo_entrega_organismo") or "",
        "validez": cot.get("validez") or "60 días",
        "observaciones": cot.get("observaciones") or "",
    }

    # --- documentos con nombre legible
    orden = ("DOC1", "DOC2", "DOC3", "PLANILLA", "DOC4")
    def prioridad(a):
        alto = a["nombre"].upper()
        for i, pre in enumerate(orden):
            if alto.startswith(pre):
                return (i, alto)
        return (99, alto)

    menu_docs = []
    secciones_docs = []
    for archivo in sorted(archivos, key=prioridad):
        nom = archivo["nombre"]
        ident = "doc-" + "".join(c if c.isalnum() else "-" for c in nom)[:32]
        etiqueta = nombre_legible(nom)
        menu_docs.append(
            f'<button data-seccion="{ident}">{_esc(etiqueta)}'
            f'<span class="chip">{_peso(archivo["bytes"])}</span></button>')
        cuerpo = (f'<iframe class="visor" src="/p/{token}/ver/{_esc(nom)}" loading="lazy" '
                  f'title="{_esc(nom)}"></iframe>' if es_previble(nom)
                  else '<div class="cuerpo"><p class="ayuda">Este formato no se puede mostrar aquí. '
                       'Descárgalo para abrirlo.</p></div>')
        secciones_docs.append(
            f'<section class="seccion oculto" id="sec-{ident}">'
            f'<div class="caja"><h2>{_esc(etiqueta)}'
            f'<span class="der"><a class="btn-mini" href="/p/{token}/f/{_esc(nom)}" download>Descargar</a></span>'
            f'</h2>{cuerpo}</div></section>')

    # --- anexos
    filas = []
    for adj in sorted(adjuntos, key=lambda a: a["nombre"]):
        nom = adj["nombre"]
        acc = [f'<a class="btn-mini" href="/p/{token}/a/{_esc(nom)}" download>Descargar</a>']
        if es_previble(nom):
            acc.insert(0, f'<a class="btn-mini" href="/p/{token}/va/{_esc(nom)}" '
                          f'target="_blank" rel="noopener">Ver</a>')
        filas.append(f'<tr><td>{_esc(nom)}</td><td class="der">{_peso(adj["bytes"])}</td>'
                     f'<td class="der">{" ".join(acc)}</td></tr>')
    tabla_anexos = ('<div class="scroll-x"><table class="lista"><thead><tr><th>Documento</th>'
                    '<th class="der">Tamaño</th><th class="der">Acciones</th></tr></thead><tbody>'
                    + "".join(filas) + '</tbody></table></div>') if filas else \
                   '<p class="vacio">Este proceso no trae adjuntos del organismo.</p>'

    # --- cotización editable
    opciones_plazo = ["", "1 día", "2 días", "3 días", "5 días hábiles", "10 días hábiles",
                      "15 días hábiles", "30 días corridos", "60 días corridos"]
    if datos["plazo"] and datos["plazo"] not in opciones_plazo:
        opciones_plazo.insert(1, datos["plazo"])
    selects_plazo = "".join(
        f'<option value="{_esc(o)}"{" selected" if o == datos["plazo"] else ""}>'
        f'{_esc(o) or "— elegir —"}</option>' for o in opciones_plazo)

    form_cotizacion = f"""
<form method="post" action="/p/{token}/cotizacion" class="cuerpo">
  <p class="ayuda">Edita los datos y el monto. Al guardar se actualiza esta cotización;
  con <b>Descargar PDF</b> se abre la versión imprimible (elige «Guardar como PDF»).</p>
  <div class="grid2">
    <div class="campo"><label>Razón social</label>
      <input name="razon_social" value="{_esc(datos['razon_social'])}"></div>
    <div class="campo"><label>RUT</label>
      <input name="rut" value="{_esc(datos['rut'])}"></div>
    <div class="campo"><label>Representante legal</label>
      <input name="representante" value="{_esc(datos['representante'])}"></div>
    <div class="campo"><label>Correo de contacto</label>
      <input name="email" value="{_esc(datos['email'])}"></div>
    <div class="campo"><label>Teléfono</label>
      <input name="telefono" value="{_esc(datos['telefono'])}"></div>
    <div class="campo"><label>Monto total (CLP, IVA incluido)</label>
      <input name="monto" type="number" value="{int(datos['monto'] or 0)}"></div>
    <div class="campo"><label>Plazo de entrega comprometido</label>
      <select name="plazo">{selects_plazo}</select></div>
    <div class="campo"><label>Validez de la oferta</label>
      <input name="validez" value="{_esc(datos['validez'])}"></div>
  </div>
  <div class="campo"><label>Observaciones para el organismo</label>
    <textarea name="observaciones">{_esc(datos['observaciones'])}</textarea></div>
  <div style="display:flex;gap:10px;flex-wrap:wrap">
    <button class="btn" type="submit">Guardar cotización</button>
    <a class="btn btn-sec" href="/p/{token}/cotizacion" target="_blank" rel="noopener">Descargar PDF</a>
  </div>
</form>"""

    aviso_html = f'<div class="aviso">{aviso}</div>' if aviso else ""
    logo_html = (f'<img src="{_esc(logo_url)}" alt="Logo">' if hay_logo
                 else '<div class="ini">L</div>')

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow">
<title>{_esc(nombre[:60])} · LICITAIA</title>
<style>{CSS}</style>
</head>
<body>
<div class="shell">
<aside>
  <div class="marca">{logo_html}
    <div><strong>LICITAIA</strong><small>Panel del proyecto</small></div>
  </div>
  <nav class="menu">
    <div class="titulo">Proyecto</div>
    <button data-seccion="resumen" class="activo">Resumen</button>
    <button data-seccion="datos">Datos del proceso</button>

    <div class="titulo">Documentos</div>
    {"".join(menu_docs) if menu_docs else '<button data-seccion="resumen">Sin documentos</button>'}
    <button data-seccion="cotizacion">Editar cotización</button>

    <div class="titulo">Anexos</div>
    <button data-seccion="anexos">Anexos<span class="chip">{len(adjuntos)}</span></button>
    <button data-seccion="subir">Subir anexos</button>

    <div class="titulo">Empresa</div>
    <button data-seccion="logo">Logo</button>
    <div style="padding:10px 8px">
      <a class="btn-sec btn" style="width:100%;text-align:center" href="/p/{token}/zip">Descargar todo</a>
    </div>
  </nav>
</aside>

<main>
  <div class="cab">
    <div>
      <h1>{_esc(nombre[:96])}</h1>
      <div class="sub">{_esc(organismo[:70])} · {_esc(codigo)} · {"Compra Ágil" if es_ca else "Licitación pública"}</div>
    </div>
    <div class="der">
      {f'<a class="btn btn-sec" href="{_esc(proyecto.get("url_portal") or "")}" target="_blank" rel="noopener">Ver en Mercado Público</a>' if proyecto.get("url_portal") else ""}
      <a class="btn" href="/p/{token}/zip">Descargar todo</a>
    </div>
  </div>

  {aviso_html}

  <div class="kpis">
    <div class="kpi"><div class="et">Presupuesto</div><div class="va">{_clp(presupuesto)}</div><div class="no">IVA incluido</div></div>
    <div class="kpi"><div class="et">Oferta</div><div class="va">{_clp(datos["monto"])}</div><div class="no">editable</div></div>
    <div class="kpi"><div class="et">Cierre</div><div class="va" style="font-size:15px">{_esc(_fecha(proyecto.get("fecha_cierre", "")))}</div><div class="no">plazo fatal</div></div>
    <div class="kpi"><div class="et">Documentos</div><div class="va">{len(archivos)}</div><div class="no">{len(adjuntos)} anexos</div></div>
  </div>

  <section class="seccion" id="sec-resumen">
    <div class="caja"><h2>Documentos del proyecto</h2><div class="cuerpo">
      <div class="scroll-x"><table class="lista">
        <thead><tr><th>Documento</th><th class="der">Tamaño</th><th class="der">Acción</th></tr></thead>
        <tbody>{"".join(f'<tr><td>{_esc(nombre_legible(a["nombre"]))}</td><td class="der">{_peso(a["bytes"])}</td><td class="der"><a class="btn-mini" href="/p/{token}/f/{_esc(a["nombre"])}" download>Descargar</a></td></tr>' for a in sorted(archivos, key=prioridad))}</tbody>
      </table></div>
      <p class="ayuda" style="margin-top:12px">Abre cada documento desde el menú de la izquierda: se muestran
      aquí mismo, sin descargar.</p>
    </div></div>
  </section>

  <section class="seccion oculto" id="sec-datos">
    <div class="caja"><h2>Datos del proceso</h2><div class="cuerpo">
      <div class="scroll-x"><table class="lista">
        <tbody>
          <tr><th>Código</th><td>{_esc(codigo)}</td></tr>
          <tr><th>Tipo</th><td>{"Compra Ágil" if es_ca else "Licitación pública"}</td></tr>
          <tr><th>Organismo</th><td>{_esc(organismo)}</td></tr>
          <tr><th>Presupuesto oficial</th><td>{_clp(presupuesto)}</td></tr>
          <tr><th>Tipo de presupuesto</th><td>{_esc(proyecto.get("tipo_presupuesto") or "no informado")}</td></tr>
          <tr><th>Plazo exigido</th><td>{_esc(proyecto.get("plazo_entrega_organismo") or "no informado")}</td></tr>
          <tr><th>Lugar de entrega</th><td>{_esc(proyecto.get("direccion_entrega") or "no informado")}</td></tr>
          <tr><th>Cierre</th><td>{_esc(_fecha(proyecto.get("fecha_cierre", "")))}</td></tr>
        </tbody></table></div>
    </div></div>
  </section>

  {''.join(secciones_docs)}

  <section class="seccion oculto" id="sec-cotizacion">
    <div class="caja"><h2>Editar cotización</h2>{form_cotizacion}</div>
  </section>

  <section class="seccion oculto" id="sec-anexos">
    <div class="caja"><h2>Anexos del organismo</h2><div class="cuerpo">{tabla_anexos}</div></div>
  </section>

  <section class="seccion oculto" id="sec-subir">
    <div class="caja"><h2>Subir anexos</h2><div class="cuerpo">
      {('<p class="ayuda">Compra Ágil: los documentos se descargan solos, no necesitas subir nada.</p>' if es_ca else
        '<p class="ayuda">El portal oficial no siempre permite bajar las bases. Súbelas aquí (PDF, DOCX, XLSX, ZIP) y el agente las incorpora al informe.</p>'
        f'<form method="post" action="/p/{token}/subir" enctype="multipart/form-data" style="display:flex;gap:10px;flex-wrap:wrap;align-items:center">'
        '<input type="file" name="archivos" multiple><button class="btn" type="submit">Subir anexos</button></form>'
        f'<form method="post" action="/p/{token}/reprocesar" style="margin-top:10px">'
        '<button class="btn btn-sec" type="submit">Regenerar auditoría con estos anexos</button></form>')}
    </div></div>
  </section>

  <section class="seccion oculto" id="sec-logo">
    <div class="caja"><h2>Logo de tu empresa</h2><div class="cuerpo">
      <p class="ayuda">Aparece en la cotización, el informe y el financiero. PNG con fondo transparente,
      ideal 600 px de ancho.</p>
      <div style="display:flex;align-items:center;gap:16px;flex-wrap:wrap">
        {logo_html}
        <form method="post" action="/p/{token}/logo" enctype="multipart/form-data"
              style="display:flex;gap:10px;flex-wrap:wrap;align-items:center">
          <input type="file" name="logo" accept=".png,.jpg,.jpeg">
          <button class="btn" type="submit">Subir logo</button>
        </form>
      </div>
    </div></div>
  </section>
</main>
</div>
<script>{JS}</script>
<script>
(function(){{
  var marca = "{doc1_marca}";
  var mostrado = false;
  setInterval(function(){{
    if (mostrado) return;
    fetch("/p/{token}/estado", {{cache:"no-store"}}).then(function(r){{return r.json();}})
      .then(function(d){{
        if (d && d.doc1 && String(d.doc1) !== marca) {{
          mostrado = true;
          var aviso = document.createElement('div');
          aviso.className = 'aviso';
          aviso.textContent = 'Análisis actualizado. Recargando…';
          document.querySelector('main').insertBefore(aviso, document.querySelector('.kpis'));
          setTimeout(function(){{ location.reload(); }}, 1600);
        }}
      }}).catch(function(){{}});
  }}, 20000);
}})();
</script>
</body>
</html>"""


# ---------------------------------------------------------------------------
# Cotización imprimible
# ---------------------------------------------------------------------------
def cotizacion_html(proyecto, token, datos, hay_logo=True, logo_url="", cot=None):
    cot = dict(cot or {})
    codigo = _esc(proyecto.get("codigo", ""))
    organismo = _esc(proyecto.get("organismo", ""))
    nombre = _esc(str(proyecto.get("nombre") or "Servicio"))
    monto = int(cot.get("monto") or proyecto.get("oferta_bruta_clp") or 0)
    neto = int(round(monto / 1.19))
    iva = monto - neto
    plazo = _esc(cot.get("plazo") or proyecto.get("plazo_entrega_organismo") or "según bases")
    validez = _esc(cot.get("validez") or "60 días")
    obs = _esc(cot.get("observaciones") or "")
    logo = f'<img src="{_esc(logo_url)}" alt="Logo" style="height:52px;max-width:170px;object-fit:contain">' if hay_logo else ""

    return f"""<!DOCTYPE html>
<html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow">
<title>Cotización {codigo}</title>
<style>
@page{{size:A4;margin:16mm}}
body{{margin:0;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;
color:#1c2536;font-size:12.5px;line-height:1.5}}
.hoja{{max-width:780px;margin:0 auto;padding:18px}}
.cab{{display:flex;justify-content:space-between;align-items:flex-start;gap:16px;
border-bottom:2px solid #12213a;padding-bottom:12px;margin-bottom:18px}}
.cab h1{{font-size:17px;margin:0;color:#12213a;text-transform:uppercase;letter-spacing:.04em}}
.cab .meta{{text-align:right;font-size:11.5px;color:#5b6674}}
h2{{font-size:12px;text-transform:uppercase;letter-spacing:.07em;color:#1e3a5f;
margin:18px 0 6px;border-bottom:1px solid #e3e7ee;padding-bottom:4px}}
table{{width:100%;border-collapse:collapse;font-size:12px}}
th,td{{border:1px solid #d9dee7;padding:6px 8px;text-align:left;vertical-align:top}}
th{{background:#f3f6fa;font-weight:600}}
.der{{text-align:right}}
.tot{{background:#f8fafc;font-weight:700}}
ul{{margin:4px 0 0 18px;padding:0}}
li{{margin-bottom:3px}}
.pie{{margin-top:22px;font-size:10.5px;color:#5b6674;text-align:center}}
.barra{{text-align:center;margin:18px 0}}
.barra button{{background:#12213a;color:#fff;border:0;border-radius:9px;padding:10px 20px;
font-size:13px;font-weight:650;cursor:pointer}}
@media print{{.barra{{display:none}}body{{font-size:11.5px}}}}
</style></head><body>
<div class="barra"><button onclick="window.print()">Imprimir / Guardar como PDF</button></div>
<div class="hoja">
  <div class="cab">
    <div>{logo}
      <div style="margin-top:8px;font-weight:700">{_esc(cot.get("razon_social") or "")}</div>
      <div style="font-size:11px;color:#5b6674">RUT {_esc(cot.get("rut") or "")}<br>
      {_esc(cot.get("email") or "")} · {_esc(cot.get("telefono") or "")}</div>
    </div>
    <div class="meta">
      <strong>COTIZACIÓN</strong><br>
      Proceso: {codigo}<br>
      Fecha: {datetime.now():%d-%m-%Y}
    </div>
  </div>

  <h2>Antecedentes del proceso</h2>
  <table>
    <tr><th style="width:32%">Organismo comprador</th><td>{organismo}</td></tr>
    <tr><th>Proceso</th><td>{codigo}</td></tr>
    <tr><th>Objeto</th><td>{nombre}</td></tr>
    <tr><th>Representante legal</th><td>{_esc(cot.get("representante") or "")}</td></tr>
  </table>

  <h2>Oferta económica</h2>
  <table>
    <tr><th>Descripción</th><th class="der" style="width:22%">Monto (CLP)</th></tr>
    <tr><td>{nombre}</td><td class="der">{_clp(neto)}</td></tr>
    <tr><td>IVA 19%</td><td class="der">{_clp(iva)}</td></tr>
    <tr class="tot"><td>TOTAL CON IVA INCLUIDO</td><td class="der">{_clp(monto)}</td></tr>
  </table>
  <p style="font-size:11px;color:#5b6674;margin-top:6px">
    Son {_clp(monto)} pesos con IVA incluido. La presente oferta no supera el presupuesto
    máximo informado por el organismo requirente.</p>

  <h2>Condiciones comerciales</h2>
  <ul>
    <li>Plazo de entrega: <strong>{plazo}</strong>, conforme a lo exigido en la ficha del proceso.</li>
    <li>Validez de la oferta: {validez}.</li>
    <li>Forma de pago: 30 días contra recepción conforme y factura electrónica (Ley 21.131).</li>
    <li>Garantía técnica de 12 meses desde la recepción, con soporte preferente.</li>
    {f"<li>{obs}</li>" if obs else ""}
  </ul>

  <h2>Declaración</h2>
  <p>El oferente declara conocer y aceptar las bases administrativas y técnicas del proceso,
  y que la información entregada es fidedigna.</p>

  <div style="margin-top:34px;display:flex;gap:60px">
    <div style="flex:1;border-top:1px solid #9aa3ae;padding-top:6px;font-size:11px;color:#5b6674">
      {_esc(cot.get("representante") or "")}<br>Representante legal
    </div>
    <div style="flex:1;border-top:1px solid #9aa3ae;padding-top:6px;font-size:11px;color:#5b6674">
      Fecha y firma
    </div>
  </div>

  <div class="pie">Documento generado con LICITAIA · {codigo}</div>
</div>
</body></html>"""
