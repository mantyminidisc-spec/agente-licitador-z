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
        estilo_p = ""
        try:
            if parrafo.alignment is not None and "CENTER" in str(parrafo.alignment):
                estilo_p = ' style="text-align:center"'
        except Exception:
            estilo_p = ""
        cuerpo = f"<p{estilo_p}>{marcado}</p>" if marcado else ""
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



# ---------------------------------------------------------------------------
# Documento renderizado (Informe / Financiero) desde datos_informe.json
# ---------------------------------------------------------------------------
TITULOS = {
    "resumen_ejecutivo": "Resumen ejecutivo",
    "objeto": "Objeto del requerimiento",
    "objeto_alcance": "Objeto y alcance",
    "analisis_consultor": "Análisis del consultor",
    "admisibilidad": "Admisibilidad de la oferta",
    "requisitos": "Requisitos obligatorios",
    "excluyentes": "Causales de inadmisibilidad",
    "documentos_obligatorios": "Documentos que debe presentar",
    "criterios_evaluacion": "Criterios de evaluación",
    "plazos": "Plazos y hitos",
    "garantias": "Garantías exigidas",
    "multas": "Multas y sanciones",
    "riesgos": "Riesgos detectados",
    "observaciones": "Cláusulas observadas",
    "preguntas": "Preguntas para el foro",
    "auditoria_multas_ds661": "Auditoría de multas (DS 661/2024)",
    "analisis_por_documento": "Análisis documental por anexo",
    "escenarios": "Escenarios de precio",
    "plan_contractual": "Plan contractual",
    "recomendaciones": "Recomendaciones",
    "conclusion": "Conclusión",
    "estrategia": "Estrategia comercial",
    "justificacion_precio": "Justificación del precio",
}

TECNICOS = ("informe_forense", "texto_extraido", "adjuntos_locales", "resumen_item",
            "nombres_adjuntos", "url_portal", "keyword_match")


def _formato_celda(v):
    if v is None:
        return ""
    if isinstance(v, float):
        if abs(v) < 1 and v != 0:
            return f"{v*100:.1f}%"
        return f"{v:,.0f}".replace(",", ".")
    if isinstance(v, int):
        return f"{v:,}".replace(",", ".")
    if isinstance(v, (list, dict)):
        return ""
    return str(v)


def _titulo(clave):
    if clave in TITULOS:
        return TITULOS[clave]
    return str(clave).replace("_", " ").strip().capitalize()


def _valor_html(valor, nivel=0):
    """Convierte cualquier estructura en HTML legible."""
    if valor is None or valor == "" or valor == [] or valor == {}:
        return ""
    if isinstance(valor, bool):
        return f"<p>{'Sí' if valor else 'No'}</p>"
    if isinstance(valor, (int, float)):
        return f"<p>{valor:,}</p>".replace(",", ".")
    if isinstance(valor, str):
        texto = valor.strip()
        if not texto:
            return ""
        lineas = [l.strip("•-* ").strip() for l in texto.split("\n") if l.strip()]
        if len(lineas) > 1:
            return "<ul>" + "".join(f"<li>{_esc(l)}</li>" for l in lineas) + "</ul>"
        return f"<p>{_esc(texto)}</p>"
    if isinstance(valor, list):
        if all(isinstance(v, (str, int, float)) for v in valor):
            return "<ul>" + "".join(f"<li>{_esc(v)}</li>" for v in valor) + "</ul>"
        partes = []
        for elemento in valor:
            if isinstance(elemento, dict):
                titulo = (elemento.get("archivo") or elemento.get("titulo")
                          or elemento.get("nombre") or elemento.get("escenario") or "")
                partes.append('<div class="subbloque">')
                if titulo:
                    partes.append(f'<h4>{_esc(titulo)}</h4>')
                for clave, contenido in elemento.items():
                    if clave in ("archivo", "titulo", "nombre", "escenario", "_extractivo"):
                        continue
                    if str(clave).startswith("_"):
                        continue
                    trozo = _valor_html(contenido, nivel + 1)
                    if trozo:
                        partes.append(f'<div class="campo-doc"><span>{_esc(_titulo(clave))}</span>'
                                      f'{trozo}</div>')
                partes.append("</div>")
            else:
                partes.append(_valor_html(elemento, nivel + 1))
        return "".join(partes)
    if isinstance(valor, dict):
        # diccionario de diccionarios (escenarios, estudio financiero) -> tabla
        if valor and all(isinstance(v, dict) for v in valor.values()):
            columnas = []
            for sub in valor.values():
                for k in sub.keys():
                    if k not in columnas and not str(k).startswith("_"):
                        columnas.append(k)
            cabecera = "<tr><th>Concepto</th>" + "".join(
                f"<th>{_esc(_titulo(c))}</th>" for c in columnas) + "</tr>"
            filas = []
            for nombre_fila, sub in valor.items():
                celdas = "".join(
                    f"<td>{_esc(_formato_celda(sub.get(c)))}</td>" for c in columnas)
                filas.append(f"<tr><td><strong>{_esc(_titulo(nombre_fila))}</strong></td>"
                             f"{celdas}</tr>")
            return ('<div class="scroll-x"><table class="tabla-doc">'
                    + cabecera + "".join(filas) + "</table></div>")
        partes = []
        for clave, contenido in valor.items():
            if clave in TECNICOS:
                continue
            trozo = _valor_html(contenido, nivel + 1)
            if trozo:
                partes.append(f'<div class="campo-doc"><span>{_esc(_titulo(clave))}</span>'
                              f'{trozo}</div>')
        return "".join(partes)
    return f"<p>{_esc(valor)}</p>"


def documento_html(titulo_doc, datos, logo_url="", hay_logo=False, meta_extra=""):
    proyecto = datos or {}
    informes = proyecto.get("informes") or {}
    codigo = _esc(proyecto.get("codigo", ""))
    organismo = _esc(proyecto.get("organismo", ""))
    nombre = _esc(proyecto.get("nombre", ""))
    logo = (f'<img src="{_esc(logo_url)}" alt="Logo">' if hay_logo else "")

    claves_financieras = ("escenarios", "estudio_financiero", "dictamen", "recomendaciones",
                          "conclusion", "roi", "margen", "costos", "flujo_caja",
                          "justificacion_precio", "precio_recomendado")
    es_financiero = str(proyecto.get("_tipo_doc") or "").lower() == "financiero"
    cuerpo = []
    for clave, contenido in informes.items():
        if clave in TECNICOS:
            continue
        if es_financiero and clave not in claves_financieras:
            continue
        trozo = _valor_html(contenido)
        if not trozo:
            continue
        cuerpo.append(f'<section class="bloque-doc"><h3>{_esc(_titulo(clave))}</h3>{trozo}</section>')

    if not cuerpo:
        cuerpo.append('<p class="vacio">Este documento todavía no tiene datos guardados. '
                      'Vuelve a generar el proyecto para verlo aquí.</p>')

    return f"""<!DOCTYPE html>
<html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow">
<title>{_esc(titulo_doc)} · {codigo}</title>
<style>
@page{{size:A4;margin:16mm}}
body{{margin:0;background:#eef1f6;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",
Roboto,Arial,sans-serif;color:#1c2536;font-size:13px;line-height:1.6}}
.hoja{{max-width:820px;margin:18px auto;background:#fff;border-radius:14px;
box-shadow:0 8px 30px rgba(18,33,58,.10);padding:34px 40px 44px}}
.cab{{display:flex;justify-content:space-between;align-items:flex-start;gap:18px;
border-bottom:2px solid #12213a;padding-bottom:14px;margin-bottom:20px}}
.cab img{{height:52px;max-width:180px;object-fit:contain}}
.cab h1{{font-size:16px;margin:0;color:#12213a;text-transform:uppercase;letter-spacing:.06em}}
.cab .sub{{font-size:11.5px;color:#5b6674;margin-top:4px}}
.cab .der{{text-align:right;font-size:11.5px;color:#5b6674}}
h3{{font-size:12.5px;text-transform:uppercase;letter-spacing:.07em;color:#1e3a5f;
margin:22px 0 8px;padding-bottom:5px;border-bottom:1px solid #e3e7ee}}
h4{{font-size:12.5px;margin:12px 0 4px;color:#12213a}}
p{{margin:0 0 8px}}
ul{{margin:4px 0 10px 18px;padding:0}}
li{{margin-bottom:4px}}
.subbloque{{border-left:3px solid #e3e7ee;padding:2px 0 2px 12px;margin:10px 0}}
.campo-doc{{margin:6px 0}}
.campo-doc>span{{display:block;font-size:10.5px;text-transform:uppercase;letter-spacing:.07em;
color:#8b95a4;font-weight:700;margin-bottom:2px}}
.vacio{{color:#5b6674}}
.barra{{max-width:820px;margin:14px auto 0;display:flex;justify-content:flex-end;gap:10px}}
.barra a,.barra button{{background:#12213a;color:#fff;border:0;border-radius:9px;padding:9px 16px;
font-size:13px;font-weight:650;text-decoration:none;cursor:pointer}}
.pie{{margin-top:26px;border-top:1px solid #e3e7ee;padding-top:10px;font-size:10.5px;
color:#8b95a4;text-align:center}}
@media print{{body{{background:#fff}}.hoja{{box-shadow:none;margin:0;padding:0;border-radius:0}}
.barra{{display:none}}}}
@media (max-width:640px){{.hoja{{margin:0;border-radius:0;padding:20px 16px}}}}
</style></head><body>
<div class="barra">
  <button onclick="window.print()">Imprimir / Guardar PDF</button>
</div>
<div class="hoja">
  <div class="cab">
    <div>{logo}<div class="sub">{meta_extra}</div></div>
    <div class="der">
      <strong>{_esc(titulo_doc)}</strong><br>
      {codigo}<br>{organismo}
    </div>
  </div>
  <h1 style="font-size:15px;margin-bottom:2px">{nombre}</h1>
  {''.join(cuerpo)}
  <div class="pie">Documento generado con LICITAIA · {codigo}</div>
</div>
</body></html>"""


def xlsx_a_html(ruta, max_filas=200, max_columnas=14):
    """Muestra el Excel como una grilla real: celdas combinadas, colores y alineación."""
    try:
        import openpyxl
        from openpyxl.utils import column_index_from_string
        libro = openpyxl.load_workbook(ruta, data_only=True)
    except Exception as e:
        return f"<p>No se pudo abrir el Excel: {_esc(e)[:120]}</p>"

    def _valor(celda):
        v = celda.value
        if v is None:
            return ""
        if isinstance(v, float):
            formato = str(celda.number_format or "")
            if "%" in formato:
                return f"{v*100:.0f}%"
            if abs(v) < 1 and v != 0:
                return f"{v*100:.0f}%"
            return f"{v:,.0f}".replace(",", ".")
        if isinstance(v, int):
            return f"{v:,}".replace(",", ".")
        return _esc(v)

    partes = []
    for hoja in libro.worksheets:
        combinadas = {}
        cubiertas = set()
        for rango in hoja.merged_cells.ranges:
            filas_span = rango.max_row - rango.min_row + 1
            cols_span = rango.max_col - rango.min_col + 1
            combinadas[(rango.min_row, rango.min_col)] = (filas_span, cols_span)
            for f in range(rango.min_row, rango.max_row + 1):
                for c in range(rango.min_col, rango.max_col + 1):
                    if (f, c) != (rango.min_row, rango.min_col):
                        cubiertas.add((f, c))

        ocultas = set()
        for letra, dim in hoja.column_dimensions.items():
            if getattr(dim, "hidden", False):
                try:
                    ocultas.add(column_index_from_string(letra))
                except Exception:
                    pass

        filas_html = []
        total_filas = min(hoja.max_row or 0, max_filas)
        total_cols = min(hoja.max_column or 0, max_columnas)
        for f in range(1, total_filas + 1):
            celdas_html = []
            for c in range(1, total_cols + 1):
                if c in ocultas or (f, c) in cubiertas:
                    continue
                celda = hoja.cell(row=f, column=c)
                span = ""
                if (f, c) in combinadas:
                    fs, cs = combinadas[(f, c)]
                    if fs > 1 or cs > 1:
                        span = f' rowspan="{fs}" colspan="{cs}"'
                estilos = []
                color_fondo = ""
                try:
                    relleno = celda.fill
                    if relleno and relleno.fgColor and relleno.patternType:
                        rgb = str(relleno.fgColor.rgb or "")
                        if rgb and rgb not in ("00000000", "None"):
                            color_fondo = "#" + rgb[-6:]
                except Exception:
                    pass
                if color_fondo:
                    estilos.append(f"background:{color_fondo}")
                    # si el fondo es oscuro, el texto debe ser blanco
                    try:
                        r = int(color_fondo[1:3], 16)
                        g = int(color_fondo[3:5], 16)
                        b = int(color_fondo[5:7], 16)
                        oscuro = (r * 0.299 + g * 0.587 + b * 0.114) < 140
                    except Exception:
                        oscuro = False
                    color_texto = ""
                    try:
                        if celda.font and celda.font.color and celda.font.color.rgb:
                            rgb_texto = str(celda.font.color.rgb)
                            if len(rgb_texto) >= 6 and rgb_texto not in ("000000",):
                                color_texto = "#" + rgb_texto[-6:]
                    except Exception:
                        color_texto = ""
                    if color_texto:
                        estilos.append(f"color:{color_texto}")
                    elif oscuro:
                        estilos.append("color:#ffffff")
                if celda.font and celda.font.bold:
                    estilos.append("font-weight:700")
                if isinstance(celda.value, (int, float)):
                    estilos.append("text-align:right")
                etiqueta = "td"
                if f <= 3 and celda.font and celda.font.bold:
                    etiqueta = "td"
                estilo = f' style="{";".join(estilos)}"' if estilos else ""
                celdas_html.append(f"<{etiqueta}{span}{estilo}>{_valor(celda)}</{etiqueta}>")
            filas_html.append("<tr>" + "".join(celdas_html) + "</tr>")

        nombre_hoja = _esc(hoja.title).replace("_", " ")
        partes.append(
            f'<h4 class="hoja-titulo">{nombre_hoja}</h4>'
            '<div class="scroll-x"><table class="tabla-xls">'
            + "".join(filas_html) + "</table></div>")
    try:
        libro.close()
    except Exception:
        pass
    return "\n".join(partes) or "<p>La planilla no tiene datos.</p>"



# ---------------------------------------------------------------------------
# Simulador financiero (planilla editable)
# ---------------------------------------------------------------------------
def _normaliza_pct(valor, defecto=0.0, minimo=0.40, maximo=1.00):
    """Acepta 78 (=78%) o 0.78 (=78%). Si queda fuera de rango, usa el valor por defecto."""
    try:
        v = float(valor)
    except Exception:
        return defecto
    if v <= 0:
        return defecto
    if v > 1.5:
        v = v / 100.0
    if v > maximo or v < minimo:
        return defecto
    return v


def calcular_escenarios(presupuesto, iva_pct, costo, pcts):
    """Repite el cálculo del motor de precios: bruto, neto, IVA, ganancia y ROI."""
    presupuesto = int(presupuesto or 0)
    costo = int(costo or 0)
    iva_pct = _normaliza_pct(iva_pct, 0.19)
    pcts = [(nombre, _normaliza_pct(pct, 0.80)) for nombre, pct in pcts]
    filas = []
    for nombre, pct in pcts:
        bruto = int(round(presupuesto * pct))
        neto = int(round(bruto / (1 + iva_pct))) if (1 + iva_pct) else 0
        iva = bruto - neto
        ganancia = neto - costo
        margen = (ganancia / neto * 100) if neto else 0
        roi = (ganancia / costo * 100) if costo else 0
        filas.append({"escenario": nombre, "pct": pct, "bruto": bruto, "neto": neto,
                      "iva": iva, "costo": costo, "ganancia": ganancia,
                      "margen": margen, "roi": roi})
    return filas


def _tabla_escenarios(filas):
    cabecera = ("<tr><th>Escenario</th><th>% del tope</th><th>Oferta neta</th><th>IVA</th>"
                "<th>Total bruto</th><th>Costo interno</th><th>Ganancia</th>"
                "<th>Margen</th><th>ROI</th></tr>")
    cuerpo = []
    for f in filas:
        cuerpo.append(
            "<tr>"
            f"<td>{_esc(f['escenario'])}</td>"
            f"<td class='der'>{f['pct']*100:.0f}%</td>"
            f"<td class='der'>{_clp(f['neto'])}</td>"
            f"<td class='der'>{_clp(f['iva'])}</td>"
            f"<td class='der'><strong>{_clp(f['bruto'])}</strong></td>"
            f"<td class='der'>{_clp(f['costo'])}</td>"
            f"<td class='der'>{_clp(f['ganancia'])}</td>"
            f"<td class='der'>{f['margen']:.1f}%</td>"
            f"<td class='der'>{f['roi']:.1f}%</td>"
            "</tr>")
    return ('<div class="scroll-x"><table class="tabla-doc">' + cabecera
            + "".join(cuerpo) + "</table></div>")



# ---------------------------------------------------------------------------
# Planilla financiera editable (dentro del panel)
# ---------------------------------------------------------------------------
def planilla_html(proyecto, token, valores, filas=None, mensaje="", enlace_excel="",
                  sensibilidad=None):
    presupuesto = int(valores.get("presupuesto") or 0)
    iva = float(valores.get("iva") or 0.19)
    horas = int(valores.get("horas") or 0)
    valor_hora = int(valores.get("valor_hora") or 0)
    cloud = int(valores.get("cloud") or 0)
    gestion = int(valores.get("gestion") or 0)
    costo = horas * valor_hora + cloud + gestion
    elegido = str(valores.get("escenario_cotizacion") or "recomendado").lower()
    codigo = _esc(proyecto.get("codigo", ""))
    nombre = _esc(str(proyecto.get("nombre") or "")[:90])
    aviso = f'<div class="aviso">{_esc(mensaje)}</div>' if mensaje else ""
    filas = filas or []

    claves = ["competitivo", "recomendado", "conservador"]
    cuerpo = []
    for pos, f in enumerate(filas):
        clave = claves[pos] if pos < len(claves) else ""
        marca = " class='fila-ok'" if clave == elegido else ""
        etiqueta = _esc(f["escenario"])
        if clave == elegido:
            etiqueta += " <span class='et'>· va a la cotización</span>"
        cuerpo.append(
            f"<tr{marca}><td>{etiqueta}</td>"
            f"<td class='der'>{f['pct']*100:.0f}%</td>"
            f"<td class='der'>{_clp(f['bruto'])}</td>"
            f"<td class='der'>{_clp(f['neto'])}</td>"
            f"<td class='der'>{_clp(f['iva'])}</td>"
            f"<td class='der'>{_clp(f['ganancia'])}</td>"
            f"<td class='der'>{f['margen']:.1f}%</td></tr>")
    tabla = ("<div class='scroll-x'><table class='planilla'>"
             "<thead><tr><th>Opción</th><th class='der'>% del tope</th>"
             "<th class='der'>Total bruto</th><th class='der'>Neto</th>"
             "<th class='der'>IVA</th><th class='der'>Ganancia</th>"
             "<th class='der'>Margen</th></tr></thead><tbody>"
             + "".join(cuerpo) + "</tbody></table></div>")

    # campos de monto: el cliente puede escribir lo que quiera
    montos = []
    etiquetas = ("Opción 1 · competitiva", "Opción 2 · recomendada", "Opción 3 · conservadora")
    for pos, etiqueta in enumerate(etiquetas):
        valor = filas[pos]["bruto"] if pos < len(filas) else round(presupuesto * (0.78 + pos * 0.035))
        pct_txt = (f"({filas[pos]['pct']*100:.0f}% sugerido)" if pos < len(filas) else "")
        montos.append(
            f'<div class="campo"><label>{etiqueta} <span class="pista">{pct_txt}</span></label>'
            f'<input name="monto{pos+1}" type="text" inputmode="numeric" '
            f'value="{valor:,}".replace(",", ".")></div>')

    opciones = ""
    for clave, etiqueta in (("competitivo", "Opción 1 · competitiva"),
                            ("recomendado", "Opción 2 · recomendada"),
                            ("conservador", "Opción 3 · conservadora")):
        sel = " selected" if clave == elegido else ""
        opciones += f'<option value="{clave}"{sel}>{etiqueta}</option>'

    return f"""<!DOCTYPE html>
<html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow">
<title>Planilla financiera · {codigo}</title>
<style>
:root{{--azul:#12213a;--azul2:#1e3a5f;--texto:#1c2536;--gris:#5b6674;--linea:#e3e7ee;
--fondo:#f5f7fb;--blanco:#fff;--ambar:#fef3c7;--verde:#ecfdf5}}
*{{box-sizing:border-box}}
body{{margin:0;background:var(--fondo);color:var(--texto);font-size:14px;line-height:1.5;
font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif}}
.wrap{{max-width:940px;margin:0 auto;padding:18px 16px 50px}}
.cab{{display:flex;align-items:flex-end;gap:12px;flex-wrap:wrap;margin-bottom:14px}}
.cab h1{{font-size:17px;margin:0;color:var(--azul)}}
.cab .sub{{font-size:12px;color:var(--gris)}}
.cab .der{{margin-left:auto;display:flex;gap:8px}}
.btn{{display:inline-block;background:var(--azul);color:#fff;border:0;border-radius:9px;
padding:9px 16px;font-weight:650;font-size:13px;text-decoration:none;cursor:pointer;font-family:inherit}}
.btn-sec{{background:#eef2f8;color:var(--azul2)}}
.tarjeta{{background:var(--blanco);border:1px solid var(--linea);border-radius:14px;
padding:18px 20px;margin-bottom:14px}}
h2{{font-size:12px;text-transform:uppercase;letter-spacing:.07em;color:var(--azul2);
margin:0 0 12px;padding-bottom:6px;border-bottom:1px solid var(--linea)}}
.grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:14px 18px}}
.campo{{display:flex;flex-direction:column;gap:5px}}
.campo label{{font-size:10.5px;text-transform:uppercase;letter-spacing:.06em;
color:var(--gris);font-weight:700}}
.campo .pista{{font-weight:500;text-transform:none;letter-spacing:0;color:#8b95a4}}
.campo input,.campo select{{border:1px solid #d7dde7;border-radius:9px;padding:10px 12px;
font-size:15px;font-family:inherit;background:var(--ambar);color:var(--texto);width:100%}}
.campo input:focus,.campo select:focus{{outline:2px solid #93b4e6;background:#fff}}
.planilla{{width:100%;border-collapse:collapse;font-size:13px}}
.planilla th{{background:var(--azul);color:#fff;text-align:left;padding:9px 10px;
font-weight:600;font-size:11.5px;text-transform:uppercase;letter-spacing:.04em;white-space:nowrap}}
.planilla td{{border:1px solid #c9d2de;padding:8px 10px}}
.planilla .fila-ok td{{background:var(--verde);font-weight:600}}
.planilla .et{{font-size:10px;color:#0f766e;font-weight:700}}
.der{{text-align:right}}
.scroll-x{{overflow-x:auto}}
.ayuda{{color:var(--gris);font-size:12.5px;margin:0 0 12px}}
.aviso{{background:var(--verde);border:1px solid #a7f3d0;border-radius:10px;padding:10px 14px;
font-size:12.5px;color:#065f46;margin-bottom:14px}}
.total{{margin-top:14px;padding:12px 14px;background:#f3f6fa;border-radius:10px;
display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px}}
.total strong{{font-size:16px;color:var(--azul)}}
.acciones{{display:flex;gap:10px;flex-wrap:wrap;margin-top:16px}}
@media (max-width:640px){{.wrap{{padding:12px 10px 40px}}.tarjeta{{padding:14px}}}}
</style></head><body>
<div class="wrap">
  <div class="cab">
    <div><h1>Planilla financiera</h1>
      <div class="sub">{nombre} · {codigo}</div></div>
    <div class="der">
      {f'<a class="btn btn-sec" href="{_esc(enlace_excel)}" download>Descargar Excel</a>' if enlace_excel else ''}
      <a class="btn btn-sec" href="/p/{token}">Volver al panel</a>
    </div>
  </div>
  {aviso}
  <form method="post" action="/p/{token}/planilla">
    <div class="tarjeta">
      <h2>1. Presupuesto y costos</h2>
      <div class="grid">
        <div class="campo"><label>Presupuesto oficial (con IVA)</label>
          <input name="presupuesto" type="text" inputmode="numeric" value="{presupuesto:,}".replace(",", ".")></div>
        <div class="campo"><label>IVA (%)</label>
          <input name="iva" type="text" inputmode="decimal" value="{iva*100:g}"></div>
        <div class="campo"><label>Horas de trabajo</label>
          <input name="horas" type="text" inputmode="numeric" value="{horas}"></div>
        <div class="campo"><label>Valor por hora</label>
          <input name="valor_hora" type="text" inputmode="numeric" value="{valor_hora:,}".replace(",", ".")></div>
        <div class="campo"><label>Cloud y licencias</label>
          <input name="cloud" type="text" inputmode="numeric" value="{cloud:,}".replace(",", ".")></div>
        <div class="campo"><label>Gestión e imprevistos</label>
          <input name="gestion" type="text" inputmode="numeric" value="{gestion:,}".replace(",", ".")></div>
      </div>
      <div class="total"><span>Costo interno total</span><strong>{_clp(costo)}</strong></div>
    </div>

    <div class="tarjeta">
      <h2>2. Cómo queda cada opción</h2>
      {tabla}
      <div class="grid" style="margin-top:16px">
        <div class="campo" style="grid-column:1/-1">
          <label>¿Cuál opción va a la cotización formal?</label>
          <select name="escenario_cotizacion">{opciones}</select>
        </div>
      </div>
    </div>

    <div class="acciones">
      <button class="btn" type="submit">Guardar y actualizar la cotización</button>
      <a class="btn btn-sec" href="/p/{token}/cotizacion" target="_blank" rel="noopener">
        Ver cotización</a>
      <button class="btn btn-sec" type="submit" formaction="/p/{token}/planilla?reset=1"
              formnovalidate>Restablecer del documento</button>
    </div>
  </form>
</div>
</body></html>"""


def simulador_html(proyecto, token, valores, filas=None, mensaje="", enlace_excel=""):
    presupuesto = int(valores.get("presupuesto") or 0)
    iva = float(valores.get("iva") or 0.19)
    costo = int(valores.get("costo") or 0)
    pct1 = float(valores.get("pct1") or 0.78)
    pct2 = float(valores.get("pct2") or 0.80)
    pct3 = float(valores.get("pct3") or 0.85)
    nombre = _esc(str(proyecto.get("nombre") or "")[:90])
    codigo = _esc(proyecto.get("codigo", ""))
    resultado = _tabla_escenarios(filas) if filas else (
        '<p class="ayuda">Cambia los datos y aprieta <b>Recalcular</b> para ver la simulación.</p>')
    aviso = f'<div class="aviso">{_esc(mensaje)}</div>' if mensaje else ""
    boton_excel = (f'<a class="btn" href="{_esc(enlace_excel)}" download>Descargar Excel simulado</a>'
                   if enlace_excel else "")

    return f"""<!DOCTYPE html>
<html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow">
<title>Simulador financiero · {codigo}</title>
<style>
@page{{size:A4;margin:14mm}}
body{{margin:0;background:#eef1f6;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",
Roboto,Arial,sans-serif;color:#1c2536;font-size:13.5px;line-height:1.55}}
.hoja{{max-width:900px;margin:18px auto;background:#fff;border-radius:14px;
box-shadow:0 8px 30px rgba(18,33,58,.10);padding:26px 30px 34px}}
h1{{font-size:16px;margin:0 0 2px;color:#12213a}}
.sub{{color:#5b6674;font-size:12px;margin-bottom:16px}}
h3{{font-size:12px;text-transform:uppercase;letter-spacing:.06em;color:#1e3a5f;
margin:20px 0 8px;border-bottom:1px solid #e3e7ee;padding-bottom:4px}}
.grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:12px 16px}}
.campo{{display:flex;flex-direction:column;gap:4px}}
.campo label{{font-size:10.5px;text-transform:uppercase;letter-spacing:.06em;color:#5b6674;font-weight:700}}
.campo input{{border:1px solid #dfe4ec;border-radius:9px;padding:9px 11px;font-size:14px;
font-family:inherit;background:#fffbeb}}
.tabla-doc{{border-collapse:collapse;width:100%;font-size:12.5px}}
.tabla-doc th,.tabla-doc td{{border:1px solid #c3ccd9;padding:7px 9px}}
.tabla-doc th{{background:#12213a;color:#fff;text-align:left;font-weight:600}}
.tabla-doc tr:nth-child(even) td{{background:#f8fafc}}
.der{{text-align:right}}
.btn{{display:inline-block;background:#12213a;color:#fff;border:0;border-radius:9px;
padding:10px 18px;font-weight:650;font-size:13.5px;text-decoration:none;cursor:pointer;
font-family:inherit}}
.btn-sec{{background:#eef2f8;color:#1e3a5f}}
.acciones{{display:flex;gap:10px;flex-wrap:wrap;margin-top:18px}}
.aviso{{background:#ecfdf5;border:1px solid #a7f3d0;border-radius:10px;padding:10px 14px;
font-size:12.5px;color:#065f46;margin-bottom:14px}}
.ayuda{{color:#5b6674;font-size:12.5px}}
@media (max-width:640px){{.hoja{{margin:0;border-radius:0;padding:18px 14px}}}}
</style></head><body>
<div class="hoja">
  <h1>{nombre}</h1>
  <div class="sub">{codigo} · simulador de la planilla financiera</div>
  {aviso}
  <form method="post" action="/p/{token}/simular">
    <h3>Datos de entrada (puedes cambiarlos)</h3>
    <div class="grid">
      <div class="campo"><label>Presupuesto oficial (con IVA)</label>
        <input name="presupuesto" type="number" value="{presupuesto}"></div>
      <div class="campo"><label>IVA (%)</label>
        <input name="iva" type="number" step="0.01" value="{iva}"></div>
      <div class="campo"><label>Costo interno total</label>
        <input name="costo" type="number" value="{costo}"></div>
      <div class="campo"><label>Escenario 1 (%)</label>
        <input name="pct1" type="number" step="0.01" value="{pct1}"></div>
      <div class="campo"><label>Escenario 2 (%)</label>
        <input name="pct2" type="number" step="0.01" value="{pct2}"></div>
      <div class="campo"><label>Escenario 3 (%)</label>
        <input name="pct3" type="number" step="0.01" value="{pct3}"></div>
    </div>
    <div class="acciones">
      <button class="btn" type="submit">Recalcular</button>
      <a class="btn btn-sec" href="/p/{token}">Volver al panel</a>
      {boton_excel}
    </div>
  </form>
  <h3>Resultado</h3>
  {resultado}
  <p class="ayuda" style="margin-top:10px">El cálculo sigue la misma regla que el documento:
  el total bruto no puede superar el presupuesto oficial, y el IVA se descuenta para obtener la
  oferta neta.</p>
</div>
</body></html>"""


def escribir_xlsx_simulado(ruta, valores, filas):
    """Genera un Excel con los datos usados y el resultado de la simulación."""
    try:
        import openpyxl
        from openpyxl.styles import Alignment, Font, PatternFill
    except Exception:
        return False
    libro = openpyxl.Workbook()
    hoja = libro.active
    hoja.title = "1_Escenarios_Oferta"
    azul = PatternFill("solid", start_color="12213A")
    amarillo = PatternFill("solid", start_color="FFFBEB")
    blanco = Font(color="FFFFFF", bold=True)

    hoja.merge_cells("A1:I1")
    hoja["A1"] = "SIMULACIÓN DE ESCENARIOS"
    hoja["A1"].font = Font(color="FFFFFF", bold=True, size=13)
    hoja["A1"].fill = azul
    hoja["A1"].alignment = Alignment(horizontal="center")

    entradas = [
        ("Presupuesto oficial (tope con IVA)", int(valores.get("presupuesto") or 0)),
        ("IVA", float(valores.get("iva") or 0.19)),
        ("Costo interno total", int(valores.get("costo") or 0)),
    ]
    for i, (etiqueta, valor) in enumerate(entradas, 3):
        hoja.cell(row=i, column=1, value=etiqueta).font = Font(bold=True)
        celda = hoja.cell(row=i, column=3, value=valor)
        celda.fill = amarillo
        if etiqueta == "IVA":
            celda.number_format = "0%"

    encabezados = ["Escenario", "% del tope", "Oferta neta", "IVA", "Total bruto",
                   "Costo interno", "Ganancia", "Margen", "ROI"]
    for i, h in enumerate(encabezados, 1):
        c = hoja.cell(row=7, column=i, value=h)
        c.font = blanco
        c.fill = azul

    for f, fila in enumerate(filas, 8):
        valores_fila = [fila["escenario"], fila["pct"], fila["neto"], fila["iva"],
                        fila["bruto"], fila["costo"], fila["ganancia"],
                        fila["margen"] / 100, fila["roi"] / 100]
        for i, v in enumerate(valores_fila, 1):
            c = hoja.cell(row=f, column=i, value=v)
            if i == 2:
                c.number_format = "0%"
            elif i in (8, 9):
                c.number_format = "0.0%"
            elif i >= 3:
                c.number_format = "#,##0"

    hoja.column_dimensions["A"].width = 24
    for letra in "BCDEFGHI":
        hoja.column_dimensions[letra].width = 14
    hoja.freeze_panes = "A8"

    libro.save(ruta)
    return True


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
.doc-img{max-width:100%;height:auto;display:block;margin:8px 0}\n.grid-entrada{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));\ngap:12px 16px;margin-top:6px}\n.tabla-edit{border-collapse:collapse;width:100%;font-size:12.5px;margin-top:8px}\n.tabla-edit th{background:#12213a;color:#fff;text-align:left;padding:8px 9px;\nfont-weight:600;font-size:11.5px}\n.tabla-edit td{border:1px solid #c3ccd9;padding:7px 9px}\n.tabla-edit tr:nth-child(even) td{background:#f8fafc}\n.tabla-edit .der{text-align:right}\n.tabla-xls{border-collapse:collapse;font-size:12px;background:#fff}\n.tabla-xls{border-collapse:collapse;table-layout:fixed;font-size:12px;background:#fff}\n.tabla-xls td{border:1px solid #c3ccd9;padding:6px 8px;white-space:nowrap;\nmin-width:64px;color:#1c2536;height:26px}\n.tabla-xls tr:first-child td{background:#f3f6fa;font-weight:650}\n.hoja-titulo{font-size:12.5px;color:#1e3a5f;margin:16px 0 6px}
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
           doc1_marca=0, cotizacion=None, datos_informe=None, oferta_vigente=None):
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
        "monto": oferta_vigente or cot.get("monto") or oferta,
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
        es_planilla = nom.upper().startswith("PLANILLA")
        es_cotizacion = nom.upper().startswith("DOC2")
        renderizable = nom.upper().startswith(("DOC1", "DOC3")) and datos_informe
        if es_planilla:
            cuerpo = (f'<iframe class="visor" src="/p/{token}/planilla" loading="lazy" '
                      f'title="Planilla financiera"></iframe>')
        elif es_cotizacion:
            cuerpo = (f'<iframe class="visor" src="/p/{token}/cotizacion" loading="lazy" '
                      f'title="Cotización formal"></iframe>')
        elif renderizable:
            cuerpo = (f'<iframe class="visor" src="/p/{token}/docu/{_esc(nom)}" loading="lazy" '
                      f'title="{_esc(etiqueta)}"></iframe>')
        elif es_previble(nom):
            cuerpo = (f'<iframe class="visor" src="/p/{token}/ver/{_esc(nom)}" loading="lazy" '
                      f'title="{_esc(nom)}"></iframe>')
        else:
            cuerpo = ('<div class="cuerpo"><p class="ayuda">Este formato no se puede mostrar aquí. '
                      'Descárgalo para abrirlo.</p></div>')
        secciones_docs.append(
            f'<section class="seccion oculto" id="sec-{ident}">'
            f'<div class="caja"><h2>{_esc(etiqueta)}'
            f'<span class="der">{"<a class=\"btn-mini\" href=\"#cotizacion\" data-ir=\"cotizacion\">Editar</a> " if es_cotizacion else ""}'\
            f'<a class="btn-mini" href="/p/{token}/f/{_esc(nom)}" download>Descargar</a></span>'
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
    <a href="/p/{token}/simulador" target="_blank" rel="noopener"
       style="display:flex;align-items:center;gap:9px;color:#dbe3ee;padding:9px 10px;
       border-radius:9px;font-size:13.5px;font-weight:550;text-decoration:none">
       Simulador financiero</a>

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
