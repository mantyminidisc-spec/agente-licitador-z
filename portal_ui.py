# -*- coding: utf-8 -*-
"""
LICITAIA · Interfaz del portal (mini CRM)
===============================================================================
Genera la página del proyecto con pestañas y vista previa de los documentos:
  · PDF     → se muestra con el visor del navegador
  · DOCX    → se convierte a HTML en el servidor (python-docx)
  · XLSX    → se convierte a tabla HTML (openpyxl)
  · .doc/.xls antiguos → solo descarga
Sin dependencias externas ni servicios de terceros.
===============================================================================
"""
import html
import os
import zipfile
from datetime import datetime
from io import BytesIO

CODIFICACION = "utf-8"

TIPOS_PREVIA = (".pdf", ".docx", ".xlsx", ".txt", ".csv")


def _esc(texto):
    return html.escape(str(texto or ""), quote=True)


def _clp(valor):
    try:
        return "$" + f"{int(valor):,}".replace(",", ".")
    except Exception:
        return "$0"


def _peso(bytes_):
    try:
        kb = int(bytes_) / 1024
        return f"{kb/1024:.1f} MB" if kb > 1024 else f"{kb:.0f} KB"
    except Exception:
        return ""


def _fecha( texto):
    try:
        valor = datetime.strptime(str(texto)[:16], "%Y-%m-%d %H:%M")
        meses = ("ene", "feb", "mar", "abr", "may", "jun",
                 "jul", "ago", "sep", "oct", "nov", "dic")
        return f"{valor.day} {meses[valor.month-1]} {valor.year} · {valor:%H:%M}"
    except Exception:
        return str(texto or "por definir")


# ---------------------------------------------------------------------------
# Vista previa
# ---------------------------------------------------------------------------
def es_previble(nombre):
    return str(nombre).lower().endswith(TIPOS_PREVIA)


def docx_a_html(ruta):
    """Convierte un Word a HTML legible (párrafos, negritas y tablas)."""
    try:
        from docx import Document
    except Exception:
        return "<p>No se puede previsualizar este Word.</p>"
    try:
        documento = Document(ruta)
    except Exception as e:
        return f"<p>No se pudo abrir el Word: {_esc(e)[:120]}</p>"

    partes = []
    for parrafo in documento.paragraphs:
        texto = parrafo.text.strip()
        if not texto:
            continue
        estilo = (parrafo.style.name or "").lower()
        if estilo.startswith("heading") or estilo.startswith("título"):
            nivel = "h3" if "1" in estilo or "2" in estilo else "h4"
            partes.append(f"<{nivel}>{_esc(texto)}</{nivel}>")
            continue
        marcado = ""
        for tramo in parrafo.runs:
            if not tramo.text:
                continue
            contenido = _esc(tramo.text)
            if tramo.bold:
                contenido = f"<strong>{contenido}</strong>"
            if tramo.italic:
                contenido = f"<em>{contenido}</em>"
            marcado += contenido
        partes.append(f"<p>{marcado or _esc(texto)}</p>")

    for indice, tabla in enumerate(documento.tables, 1):
        filas = []
        for pos, fila in enumerate(tabla.rows):
            celdas = "".join(
                f"<{'th' if pos == 0 else 'td'}>{_esc(c.text.strip())}</{'th' if pos == 0 else 'td'}>"
                for c in fila.cells)
            filas.append(f"<tr>{celdas}</tr>")
        partes.append(f'<div class="tabla-scroll"><table class="tabla-doc">'
                      f'<caption>Tabla {indice}</caption>{"".join(filas)}</table></div>')

    return "\n".join(partes) or "<p>El documento no tiene contenido de texto.</p>"


def xlsx_a_html(ruta, max_filas=400):
    """Convierte un Excel en tablas HTML (una por hoja)."""
    try:
        import openpyxl
    except Exception:
        return "<p>No se puede previsualizar este Excel.</p>"
    try:
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
            if not any(valor not in (None, "") for valor in fila):
                continue
            celdas = "".join(
                f"<{'th' if numero == 1 else 'td'}>{_esc('' if v is None else v)}</{'th' if numero == 1 else 'td'}>"
                for v in fila)
            filas.append(f"<tr>{celdas}</tr>")
        partes.append(
            f'<h4>{_esc(hoja.title)}</h4>'
            f'<div class="tabla-scroll"><table class="tabla-doc">{"".join(filas)}</table></div>')
    try:
        libro.close()
    except Exception:
        pass
    return "\n".join(partes) or "<p>La planilla no tiene datos visibles.</p>"


def vista_archivo(ruta, nombre):
    """Devuelve (contenido_html, es_html). Si es PDF, se sirve tal cual."""
    extension = os.path.splitext(nombre)[1].lower()
    if extension == ".pdf":
        return None, False
    if extension == ".docx":
        return docx_a_html(ruta), True
    if extension == ".xlsx":
        return xlsx_a_html(ruta), True
    if extension in (".txt", ".csv"):
        try:
            with open(ruta, encoding="utf-8", errors="ignore") as f:
                return f"<pre>{_esc(f.read()[:200000])}</pre>", True
        except Exception:
            return "<p>No se pudo leer el archivo.</p>", True
    return ("<p>Este formato no se puede previsualizar. Descárgalo para abrirlo.</p>", True)


# ---------------------------------------------------------------------------
# Página
# ---------------------------------------------------------------------------
def pagina(proyecto, token, archivos, adjuntos, aviso="", hay_logo=True, logo_url="",
           doc1_marca=0):
    codigo = proyecto.get("codigo", "")
    nombre = str(proyecto.get("nombre") or "Proyecto")
    organismo = str(proyecto.get("organismo") or "")
    es_ca = proyecto.get("tipo_portal") == "COMPRA_AGIL"
    presupuesto = _clp(proyecto.get("presupuesto_clp", 0))
    oferta = _clp(proyecto.get("oferta_bruta_clp", 0))
    cierre = _fecha(proyecto.get("fecha_cierre", ""))
    estado = str(proyecto.get("estado") or "")
    enlace_oficial = str(proyecto.get("url_portal") or "").strip()

    # --- Documentos con vista previa
    pestañas = []
    paneles = []
    orden = ("DOC1", "DOC2", "DOC3", "Planilla", "DOC4")
    def clave_orden(n):
        for i, pref in enumerate(orden):
            if n.upper().startswith(pref.upper()):
                return (i, n)
        return (99, n)
    for archivo in sorted(archivos, key=lambda a: clave_orden(a["nombre"])):
        nom = archivo["nombre"]
        ident = "art-" + "".join(c if c.isalnum() else "-" for c in nom)[:40]
        etiqueta = nom.split("_")[0].replace("DOC", "DOC ").replace("Planilla", "Excel")
        if nom.lower().startswith("planilla"):
            etiqueta = "Excel"
        pestañas.append(
            f'<button class="pestana" data-panel="{ident}">{_esc(etiqueta)}'
            f'<span class="chip">{_peso(archivo["bytes"])}</span></button>')
        if es_previble(nom):
            cuerpo = (f'<iframe class="visor" src="/p/{token}/ver/{_esc(nom)}" '
                      f'loading="lazy" title="{_esc(nom)}"></iframe>')
        else:
            cuerpo = ('<div class="sin-vista">Este formato no se puede mostrar aquí. '
                      'Descárgalo para abrirlo.</div>')
        paneles.append(
            f'<section class="panel" id="{ident}">'
            f'<div class="panel-cab"><h3>{_esc(nom)}</h3>'
            f'<a class="btn-mini" href="/p/{token}/f/{_esc(nom)}" download>Descargar</a></div>'
            f'{cuerpo}</section>')

    # --- Anexos
    filas_adj = []
    for adj in sorted(adjuntos, key=lambda a: a["nombre"]):
        nom = adj["nombre"]
        acciones = [f'<a class="btn-mini" href="/p/{token}/a/{_esc(nom)}" download>Descargar</a>']
        if es_previble(nom):
            acciones.insert(0, f'<a class="btn-mini" href="/p/{token}/va/{_esc(nom)}" '
                               f'target="_blank" rel="noopener">Ver</a>')
        filas_adj.append(
            f'<tr><td>{_esc(nom)}</td><td class="der">{_peso(adj["bytes"])}</td>'
            f'<td class="der">{" ".join(acciones)}</td></tr>')
    tabla_adj = (
        '<div class="tabla-scroll"><table class="tabla-lista">'
        '<thead><tr><th>Documento</th><th class="der">Tamaño</th><th class="der">Acciones</th></tr></thead>'
        f'<tbody>{"".join(filas_adj)}</tbody></table></div>'
        if filas_adj else
        '<p class="vacio">Este proceso no trae adjuntos del organismo.</p>')

    # --- Checklist / subida
    bloque_subida = ""
    if not es_ca:
        faltan = [x for x in ("bases administrativas", "bases técnicas", "formularios")
                  if x not in " ".join(a["nombre"].lower() for a in adjuntos)]
        bloque_subida = (
            '<div class="bloque">'
            '<h3>Anexos para el análisis técnico</h3>'
            '<p class="ayuda">El portal oficial no siempre permite bajar las bases. '
            'Súbelas aquí (PDF, DOCX, XLSX, ZIP) y el agente las incorpora al informe.</p>'
            f'<form method="post" action="/p/{token}/subir" enctype="multipart/form-data" class="fila-form">'
            '<input type="file" name="archivos" multiple>'
            '<button class="btn" type="submit">Subir anexos</button></form>'
            + (f'<form method="post" action="/p/{token}/reprocesar" class="fila-form">'
               '<button class="btn btn-sec" type="submit">Regenerar auditoría con estos anexos</button>'
               '</form>' if estado == "anexos_subidos" or adjuntos else "")
            + '</div>')
    else:
        bloque_subida = (
            '<div class="bloque ok-verde"><h3>Compra Ágil: los documentos se descargan solos</h3>'
            '<p class="ayuda">No necesitas subir nada: el agente obtiene los archivos desde el '
            'servicio oficial de ChileCompra.</p></div>')

    # --- Logo
    vista_logo = (f'<img src="{_esc(logo_url)}" alt="Logo" class="logo">' if hay_logo
                  else '<div class="logo-vacio">?</div>')

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow">
<title>{_esc(nombre[:60])} · LICITAIA</title>
<style>
:root{{--azul:#12213a;--azul2:#1e3a5f;--texto:#1c2536;--gris:#5b6674;--linea:#e3e7ee;
--fondo:#f4f6fa;--blanco:#fff;--verde:#0f766e;--ambar:#b45309;--radio:14px}}
*{{box-sizing:border-box}}
body{{margin:0;background:var(--fondo);color:var(--texto);
font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;
font-size:15px;line-height:1.55}}
a{{color:var(--azul2)}}
.app{{max-width:1120px;margin:0 auto;padding:18px 16px 60px}}
.barra{{display:flex;align-items:center;gap:14px;background:var(--azul);color:#fff;
border-radius:var(--radio);padding:14px 18px;flex-wrap:wrap}}
.logo{{height:44px;max-width:130px;object-fit:contain;background:#fff;border-radius:10px;padding:5px}}
.logo-vacio{{width:44px;height:44px;border-radius:10px;background:#fff3;display:flex;
align-items:center;justify-content:center;font-weight:700;font-size:20px}}
.barra h1{{font-size:16px;margin:0;font-weight:650;line-height:1.25}}
.barra .sub{{font-size:12px;opacity:.75;margin-top:2px}}
.barra .acciones{{margin-left:auto;display:flex;gap:8px;flex-wrap:wrap}}
.btn{{display:inline-block;background:var(--blanco);color:var(--azul);border:0;border-radius:10px;
padding:9px 16px;font-weight:650;font-size:13.5px;text-decoration:none;cursor:pointer}}
.btn-sec{{background:#ffffff22;color:#fff;border:1px solid #ffffff44}}
.btn-mini{{display:inline-block;background:#eef2f8;color:var(--azul2);border-radius:8px;
padding:5px 11px;font-size:12.5px;font-weight:600;text-decoration:none;border:0;cursor:pointer}}
.kpis{{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin:14px 0}}
.kpi{{background:var(--blanco);border:1px solid var(--linea);border-radius:var(--radio);padding:12px 14px}}
.kpi .et{{font-size:10.5px;text-transform:uppercase;letter-spacing:.08em;color:var(--gris);font-weight:700}}
.kpi .va{{font-size:19px;font-weight:700;margin-top:3px;color:var(--azul)}}
.kpi .no{{font-size:11.5px;color:var(--gris)}}
.pestanas{{display:flex;gap:6px;overflow-x:auto;padding:4px 2px 0;margin-top:6px;
border-bottom:1px solid var(--linea);-webkit-overflow-scrolling:touch}}
.pestana{{background:transparent;border:0;border-bottom:3px solid transparent;padding:10px 12px;
font-size:13.5px;font-weight:650;color:var(--gris);cursor:pointer;white-space:nowrap;
display:flex;align-items:center;gap:7px}}
.pestana.activa{{color:var(--azul);border-bottom-color:var(--azul)}}
.chip{{background:#eef2f8;color:var(--gris);border-radius:20px;padding:1px 8px;font-size:11px;font-weight:600}}
.panel{{display:none;background:var(--blanco);border:1px solid var(--linea);border-radius:var(--radio);
margin-top:14px;overflow:hidden}}
.panel.activa{{display:block}}
.panel-cab{{display:flex;align-items:center;gap:10px;padding:12px 16px;border-bottom:1px solid var(--linea)}}
.panel-cab h3{{font-size:14px;margin:0;flex:1;word-break:break-word}}
.visor{{width:100%;height:72vh;min-height:420px;border:0;display:block;background:#fff}}
.sin-vista{{padding:26px;color:var(--gris)}}
.bloque{{background:var(--blanco);border:1px solid var(--linea);border-radius:var(--radio);
padding:16px;margin-top:14px}}
.bloque h3{{font-size:14px;margin:0 0 6px}}
.ok-verde{{background:#ecfdf5;border-color:#a7f3d0}}
.ayuda{{color:var(--gris);font-size:13px;margin:0 0 10px}}
.fila-form{{display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin-top:8px}}
input[type=file]{{font-size:13px;max-width:100%}}
.tabla-scroll{{overflow-x:auto}}
.tabla-lista{{width:100%;border-collapse:collapse;font-size:13.5px}}
.tabla-lista th{{text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.06em;
color:var(--gris);padding:10px 8px;border-bottom:1px solid var(--linea)}}
.tabla-lista td{{padding:10px 8px;border-bottom:1px solid #f1f3f7;word-break:break-word}}
.der{{text-align:right;white-space:nowrap}}
.tabla-doc{{border-collapse:collapse;font-size:13px;width:100%}}
.tabla-doc caption{{text-align:left;font-weight:700;font-size:12px;color:var(--gris);padding:8px 0}}
.tabla-doc th,.tabla-doc td{{border:1px solid var(--linea);padding:6px 8px;vertical-align:top}}
.tabla-doc th{{background:#f7f9fc;text-align:left}}
.aviso{{background:#fffbeb;border:1px solid #fde68a;border-radius:12px;padding:10px 14px;
font-size:13px;color:#7c4a03;margin-top:12px}}
.aviso-ok{{background:#ecfdf5;border-color:#a7f3d0;color:#065f46}}
.cuerpo-doc{{padding:18px 20px;max-height:70vh;overflow:auto}}
.cuerpo-doc p{{margin:0 0 8px}}
.cuerpo-doc h3{{margin:18px 0 6px;font-size:15px;color:var(--azul)}}
.cuerpo-doc h4{{margin:14px 0 4px;font-size:13.5px;color:var(--azul2)}}
.cuerpo-doc pre{{white-space:pre-wrap;font-size:12.5px}}
.vacio{{color:var(--gris);font-size:13px}}
.pie{{text-align:center;color:var(--gris);font-size:11.5px;margin-top:26px}}
@media (max-width:640px){{
 .app{{padding:12px 10px 40px}} .kpi .va{{font-size:17px}}
 .visor{{height:64vh;min-height:340px}} .barra{{padding:12px}}
 .barra h1{{font-size:14.5px}}
}}
</style>
</head>
<body>
<div class="app">

<div class="barra">
  {vista_logo}
  <div>
    <h1>{_esc(nombre[:90])}</h1>
    <div class="sub">{_esc(organismo[:70])} · {_esc(codigo)} · {"Compra Ágil" if es_ca else "Licitación"}</div>
  </div>
  <div class="acciones">
    <a class="btn" href="/p/{token}/zip">Descargar todo</a>
    {f'<a class="btn btn-sec" href="{_esc(enlace_oficial)}" target="_blank" rel="noopener">Ver en Mercado Público</a>' if enlace_oficial else ''}
  </div>
</div>

{aviso}

<div class="kpis">
  <div class="kpi"><div class="et">Presupuesto</div><div class="va">{presupuesto}</div><div class="no">IVA incluido</div></div>
  <div class="kpi"><div class="et">Oferta sugerida</div><div class="va">{oferta}</div><div class="no">con IVA</div></div>
  <div class="kpi"><div class="et">Cierre</div><div class="va" style="font-size:15px">{_esc(cierre)}</div><div class="no">plazo fatal</div></div>
  <div class="kpi"><div class="et">Documentos</div><div class="va">{len(archivos)}</div><div class="no">{len(adjuntos)} anexos del organismo</div></div>
</div>

<div class="pestanas" id="pestanas">
  <button class="pestana activa" data-panel="panel-resumen">Resumen</button>
  {"".join(pestañas)}
  <button class="pestana" data-panel="panel-anexos">Anexos<span class="chip">{len(adjuntos)}</span></button>
</div>

<section class="panel activa" id="panel-resumen">
  <div class="cuerpo-doc">
    <h3>Cómo usar este panel</h3>
    <p>Los documentos se ven <strong>aquí mismo</strong>, sin descargarlos. Usa las pestañas de arriba:
    DOC 1 es la auditoría interna, DOC 2 la cotización para el organismo, DOC 3 el informe financiero
    y Excel la planilla con fórmulas.</p>
    <p>Si necesitas el archivo para firmar o editar, cada pestaña tiene su botón <em>Descargar</em>.</p>
    <h3>Documentos de este proyecto</h3>
    <div class="tabla-scroll"><table class="tabla-lista">
      <thead><tr><th>Documento</th><th class="der">Tamaño</th><th class="der">Acción</th></tr></thead>
      <tbody>
      {"".join(f'<tr><td>{_esc(a["nombre"])}</td><td class="der">{_peso(a["bytes"])}</td>'
               f'<td class="der"><a class="btn-mini" href="/p/{token}/f/{_esc(a["nombre"])}" download>Descargar</a></td></tr>'
               for a in sorted(archivos, key=lambda x: clave_orden(x["nombre"])))}
      </tbody></table></div>
  </div>
</section>

{"".join(paneles)}

<section class="panel" id="panel-anexos">
  <div class="panel-cab"><h3>Anexos originales del organismo</h3></div>
  <div style="padding:12px 16px">{tabla_adj}</div>
</section>

{bloque_subida}

<div class="bloque">
  <h3>Logo de tu empresa</h3>
  <p class="ayuda">Aparece en la cotización, el informe y el financiero. PNG con fondo transparente,
  ideal 600 px de ancho.</p>
  <div class="fila-form">
    {vista_logo}
    <form method="post" action="/p/{token}/logo" enctype="multipart/form-data" class="fila-form">
      <input type="file" name="logo" accept=".png,.jpg,.jpeg">
      <button class="btn" type="submit">Subir logo</button>
    </form>
  </div>
</div>

<div class="pie">LICITAIA · enlace privado de {_esc(organismo[:40] or "tu empresa")} · no compartir</div>
</div>

<script>
(function () {{
  var botones = document.querySelectorAll('.pestana');
  var paneles = document.querySelectorAll('.panel');
  function activar(ident) {{
    botones.forEach(function (b) {{ b.classList.toggle('activa', b.dataset.panel === ident); }});
    paneles.forEach(function (p) {{ p.classList.toggle('activa', p.id === ident); }});
    if (history.replaceState) history.replaceState(null, '', '#' + ident);
  }}
  botones.forEach(function (b) {{
    b.addEventListener('click', function () {{ activar(b.dataset.panel); }});
  }});
  if (location.hash) {{
    var inicial = location.hash.replace('#', '');
    if (document.getElementById(inicial)) activar(inicial);
  }}

  // aviso automático cuando el análisis terminó
  var marcaDoc = "{doc1_marca}";
  var objetivo = "/p/{token}/estado";
  var mostrado = false;
  setInterval(function () {{
    if (mostrado) return;
    fetch(objetivo, {{ cache: "no-store" }}).then(function (r) {{ return r.json(); }})
      .then(function (d) {{
        if (d && d.doc1 && String(d.doc1) !== marcaDoc) {{
          mostrado = true;
          var barra = document.createElement('div');
          barra.className = 'aviso aviso-ok';
          barra.textContent = 'Análisis actualizado. Recargando…';
          document.querySelector('.app').insertBefore(barra, document.querySelector('.kpis'));
          setTimeout(function () {{ location.reload(); }}, 1600);
        }}
      }}).catch(function () {{}});
  }}, 20000);
}})();
</script>
</body>
</html>"""
