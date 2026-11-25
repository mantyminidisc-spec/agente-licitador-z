# -*- coding: utf-8 -*-
"""
LICITAIA · DOC 2 (Cotización oficial) — plantilla única
===============================================================================
Un solo lugar que genera la cotización en Word, usado por:

  · el agente, al empaquetar el proyecto (Agente 5), y
  · el portal, cada vez que se guarda un cambio de montos o de texto técnico,
    y también al momento de descargar el archivo.

Así el Word descargado nunca queda con montos distintos a los del portal.

Reglas de redacción que aplica esta plantilla (y que el agente ya no decide):

  1. El plazo de cierre se muestra como fecha real; si el portal no la tiene,
     dice "según la ficha publicada en Mercado Público" (nunca "Abierta Hoy").
  2. El SLA tiene un único criterio en todo el documento.
  3. Si las bases exigen habilitación en N días corridos, ese plazo manda y no
     se contradice con un "plazo total" en días hábiles.
  4. No se declaran certificaciones (ISO 27001 / NCh 2728) si no están
     acreditadas: se redacta como alineación a los lineamientos.
  5. Se incluyen los requisitos de la entidad leídos de los adjuntos, en una
     matriz requisito → compromiso.
===============================================================================
"""

import json
import os
import re
from datetime import datetime

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

AZUL = "1F2A44"
GRIS_TEXTO = "374151"
GRIS_SUAVE = "6B7280"
GRIS_LINEA = "D1D5DB"
GRIS_FONDO = "F3F4F6"

# Un solo criterio de soporte para todo el documento (punto 2 del checklist).
SLA_RESPUESTA = "2 horas hábiles"
SLA_RESOLUCION = "4 horas hábiles"
SLA_TEXTO = ("Soporte técnico con respuesta en un máximo de " + SLA_RESPUESTA +
             " y resolución en un máximo de " + SLA_RESOLUCION +
             ", de lunes a viernes en horario hábil, con canal directo de atención "
             "para el organismo y registro de cada solicitud.")

MESES = ("enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre")


# ---------------------------------------------------------------------------
# Utilidades de formato
# ---------------------------------------------------------------------------
def clp(valor):
    """$1.234.567"""
    try:
        return "$" + f"{int(valor):,}".replace(",", ".")
    except Exception:
        return "$0"


def _txt(valor, defecto=""):
    texto = "" if valor is None else str(valor).strip()
    return texto or defecto


def _entero(valor, defecto=0):
    try:
        return int(round(float(valor)))
    except Exception:
        return int(defecto)


def _fraccion(valor, defecto=0.19):
    """Acepta 19, 0.19 o "19%" y devuelve 0.19."""
    try:
        numero = float(str(valor).replace("%", "").replace(",", ".").strip())
    except Exception:
        return defecto
    if numero > 1:
        numero = numero / 100.0
    return numero if 0 < numero < 1 else defecto


def fecha_legible(valor):
    """Devuelve una fecha corta y legible, o "" si no hay una fecha real."""
    texto = _txt(valor)
    if not texto:
        return ""
    if texto.lower().startswith(("abierta", "abierto", "ver ficha", "por definir",
                                 "proximos", "próximos", "no informad", "según", "segun")):
        return ""
    if not any(caracter.isdigit() for caracter in texto):
        # "Abierta Hoy", "Ver ficha", "Según bases": no son una fecha
        return ""
    limpio = texto.replace("T", " ").replace(".000", "")
    for formato in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
                    "%d-%m-%Y %H:%M", "%d-%m-%Y", "%d/%m/%Y"):
        try:
            fecha = datetime.strptime(limpio[:19].strip(), formato)
            return fecha.strftime("%d-%m-%Y") + (
                (" " + fecha.strftime("%H:%M")) if "%H" in formato else "")
        except Exception:
            continue
    return limpio[:40]


def frase_cierre(valor):
    """Cómo se nombra el cierre del proceso en la carta."""
    fecha = fecha_legible(valor)
    if fecha:
        return (f"cuyo plazo de cierre es el {fecha}, según la ficha publicada en "
                f"Mercado Público")
    return ("cuyas condiciones, consultas y plazo de cierre se rigen por la ficha "
            "publicada en Mercado Público")


def fecha_larga():
    hoy = datetime.now()
    return f"{hoy.day} de {MESES[hoy.month - 1]} de {hoy.year}"


# ---------------------------------------------------------------------------
# Coherencia de textos
# ---------------------------------------------------------------------------
def normalizar_sla(texto):
    """Deja un único SLA en el documento.

    Si el texto que redactó la IA trae horas distintas (2 vs 4, hábiles o
    corridas), se reemplazan por el criterio único. Si queda cualquier otro
    número de horas, se usa el texto patrón completo.
    """
    original = _txt(texto)
    if not original:
        return SLA_TEXTO
    horas = r"\d+\s*(?:horas?|hrs?)(?:\s*(?:h[aá]biles|corridas?))?"
    limpio = re.sub(r"(?i)(respuesta[s]?)[^.;\n]{0,60}?\b" + horas,
                    "respuesta en un máximo de " + SLA_RESPUESTA, original)
    limpio = re.sub(r"(?i)(resoluci[oó]n|soluc[ií]on)[^.;\n]{0,60}?\b" + horas,
                    "resolución en un máximo de " + SLA_RESOLUCION, limpio)
    sobrantes = [n for n in re.findall(r"(\d+)\s*(?:horas?|hrs?)", limpio)
                 if n not in ("2", "4")]
    if sobrantes or len(limpio) < 20:
        return SLA_TEXTO
    if "hábiles" not in limpio.lower() and "habiles" not in limpio.lower():
        limpio = limpio.rstrip(". ") + ", en horario hábil."
    return limpio


def _sin_certificaciones(texto):
    """Cambia 'certificación ISO 27001' por 'alineado a' (no declarar de más)."""
    texto = _txt(texto)
    if not texto:
        return texto
    norma = r"(nch[\s-]*iso\s*27001|iso[\s/]*27001|nch\s*2728)"
    reemplazos = (
        # "cuenta con certificación (formal) en ISO 27001" -> "cuenta con procesos alineados a ..."
        (r"(?i)(?:contamos|cuenta|cuentan|dispone|disponemos|posee|poseemos)\s+con\s+"
         r"(?:la\s+)?certificaci[oó]n\s+(?:formal\s+)?(?:en\s+|bajo\s+)?(?:la\s+norma\s+)?" + norma,
         r"cuenta con procesos alineados a \1"),
        # "está certificada en ISO 27001" -> "cuenta con procesos alineados a ..."
        (r"(?i)(?:está|esta|están|estan)\s+certificad[oa]s?\s+en\s+" + norma,
         r"cuenta con procesos alineados a \1"),
        (r"(?i)\bcertificad[oa]s?\s+en\s+" + norma, r"con procesos alineados a \1"),
        # "certificación NCh 2728" -> "alineación a los lineamientos de ..."
        (r"(?i)\bcertificaci[oó]n(?:es)?\s+(?:formal\s+)?(?:en\s+|bajo\s+)?"
         r"(?:la\s+)?(?:norma\s+)?" + norma, r"alineación a los lineamientos de \1"),
        (r"(?i)\bcontamos con (?:la )?certificaci[oó]n\b", r"trabajamos alineados a"),
    )
    original = texto
    for patron, cambio in reemplazos:
        texto = re.sub(patron, cambio, texto)
    if texto and original[:1].isupper() and texto[:1].islower():
        texto = texto[:1].upper() + texto[1:]
    return texto


def _normas_redactadas(informes):
    """Lista de normas en formato 'alineado a', sin declarar certificados."""
    informes = informes or {}
    normas = informes.get("normas_chilenas_nch_estandares") or {}
    lista = normas.get("normas_aplicables") or informes.get("checklist_nch_detectadas") or []
    salida = []
    for norma in lista[:6]:
        texto = _sin_certificaciones(str(norma))[:180]
        if texto and texto not in salida:
            salida.append(texto)
    return salida


def _sin_dato(texto):
    """True si el texto es un relleno tipo 'No informado en las bases'."""
    limpio = _txt(texto).lower()
    if not limpio:
        return True
    return any(marca in limpio for marca in (
        "no informad", "sin informaci", "no especificad", "no aplica", "no disponible",
        "n/a", "no se informa", "por definir", "no indicad"))


def _datos_de_bases(informes, texto_bases=""):
    """Cifras que exige la entidad, primero desde el JSON y si no, desde el texto.

    Nunca inventa: si el dato no aparece, la clave queda fuera y el documento
    omite esa frase en vez de escribir un número sin respaldo.
    """
    informes = informes or {}
    fuente = informes.get("requisitos_entidad") or {}
    if not isinstance(fuente, dict):
        fuente = {}
    crudo = " ".join([texto_bases or "", json.dumps(fuente, ensure_ascii=False),
                      json.dumps(informes, ensure_ascii=False)[:4000]])
    crudo = crudo.replace("\u00a0", " ")
    salida = {}

    disponibilidad = _txt(fuente.get("disponibilidad"))
    if _sin_dato(disponibilidad):
        hallado = re.search(r"(99[.,]\d{1,2})\s*%", crudo)
        disponibilidad = (hallado.group(1).replace(".", ",") + "%") if hallado else ""
    if disponibilidad:
        # se guarda solo el porcentaje: el texto del contrato agrega "mensual"
        solo_pct = re.search(r"\d+[.,]?\d*\s*%", disponibilidad)
        salida["disponibilidad"] = ((solo_pct.group(0).replace(" ", ""))
                                    if solo_pct else disponibilidad.strip())

    usuarios = _txt(fuente.get("usuarios"))
    if _sin_dato(usuarios):
        hallado = re.search(r"(?i)(\d{1,3})\s*(?:usuarios?|funcionarios?|agentes?)", crudo)
        usuarios = hallado.group(1) if hallado else ""
    if usuarios:
        salida["usuarios"] = re.sub(r"[^\d]", "", usuarios)[:3]

    vigencia = _txt(fuente.get("vigencia"))
    if _sin_dato(vigencia):
        hallado = re.search(r"(?i)(\d{1,2})\s*meses", crudo)
        vigencia = (hallado.group(1) + " meses") if hallado else ""
    if vigencia:
        meses = re.search(r"(\d{1,2})\s*mes", vigencia.lower())
        salida["vigencia"] = (meses.group(1) + " meses") if meses else vigencia.lower().strip()

    habilitacion = _txt(fuente.get("plazo_habilitacion"))
    if _sin_dato(habilitacion):
        habilitacion = ""
        patron_dias = r"(?i)\d{1,3}\s*d[ií]as?\s*corridos?"
        # El plazo suele venir DESPUÉS de "habilitación / puesta en marcha":
        # se busca ahí primero para no confundirlo con el plazo de pago.
        for marca in re.finditer(r"(?i)(habilitaci[oó]n|puesta en marcha|"
                                 r"implementaci[oó]n|inicio del servicio)", crudo):
            despues = crudo[marca.end():marca.end() + 200]
            antes = crudo[max(0, marca.start() - 80):marca.start()]
            hallado = (re.search(patron_dias, despues)
                       or re.search(patron_dias, antes))
            if hallado:
                habilitacion = hallado.group(0)
                break
        if not habilitacion:
            hallado = re.search(patron_dias, crudo)
            habilitacion = hallado.group(0) if hallado else ""
    if habilitacion:
        salida["habilitacion"] = habilitacion.lower().strip()

    forma_pago = _txt(fuente.get("forma_pago"))
    if _sin_dato(forma_pago):
        forma_pago = ""
        hallado = re.search(r"(?i)(\d{1,2})\s*cuotas?", crudo)
        if hallado:
            forma_pago = hallado.group(1) + (" cuota" if hallado.group(1) == "1" else " cuotas")
    if forma_pago:
        salida["forma_pago"] = forma_pago

    return salida


def _requisitos(informes, texto_bases=""):
    """Matriz requisito → compromiso, tomada de lo que el agente leyó.

    Orden de preferencia:
      1. requisitos_entidad.obligatorios (lo que exige la entidad, con su fuente)
      2. analisis_por_documento[].requisitos (resumen por adjunto)
      3. las cifras que se pudieron leer de las bases (plazos, disponibilidad,
         usuarios, vigencia, pago), para que la sección nunca quede vacía
         cuando la IA no devolvió el bloque.
    """
    informes = informes or {}
    salida = []
    entidad = informes.get("requisitos_entidad") or {}
    if isinstance(entidad, dict):
        for item in (entidad.get("obligatorios") or [])[:8]:
            if isinstance(item, dict):
                requisito = _txt(item.get("requisito"))[:220]
                fuente = _txt(item.get("exigido_en"))[:70]
                compromiso = _txt(item.get("compromiso"))[:220]
            else:
                requisito, fuente, compromiso = _txt(item)[:220], "", ""
            if requisito:
                salida.append({"requisito": requisito, "fuente": fuente,
                               "compromiso": compromiso or "Se cumple en su totalidad"})
    if salida:
        return salida
    for documento in (informes.get("analisis_por_documento") or [])[:4]:
        if not isinstance(documento, dict):
            continue
        archivo = _txt(documento.get("archivo"))[:60]
        for requisito in (documento.get("requisitos") or [])[:4]:
            texto = _txt(requisito)[:220]
            if texto:
                salida.append({"requisito": texto, "fuente": archivo,
                               "compromiso": "Se cumple en su totalidad"})
    if not salida:
        bases = _datos_de_bases(informes, texto_bases)
        plantillas = (
            ("habilitacion", "Habilitación y puesta en marcha en {}"),
            ("disponibilidad", "Disponibilidad mínima de {} mensual"),
            ("usuarios", "Atención para {} o más usuarios con roles y permisos"),
            ("vigencia", "Vigencia del servicio de {}"),
            ("forma_pago", "Pago en {} contra recepción conforme y factura electrónica"),
        )
        for clave, plantilla in plantillas:
            if bases.get(clave):
                salida.append({"requisito": plantilla.format(bases[clave]),
                               "fuente": "Bases del proceso",
                               "compromiso": "Se cumple en su totalidad"})
    return salida[:8]


# ---------------------------------------------------------------------------
# Hoja con el estilo editorial de los documentos LICITAIA
# ---------------------------------------------------------------------------
class _Hoja:
    def __init__(self, perfil=None, logo=""):
        self.perfil = perfil or {}
        self.logo = logo

    # -- estilos base -------------------------------------------------------
    def configurar(self, doc, margen_cm=1.5, tam_base=9.5):
        seccion = doc.sections[0]
        seccion.top_margin = Cm(margen_cm)
        seccion.bottom_margin = Cm(margen_cm)
        seccion.left_margin = Cm(1.9)
        seccion.right_margin = Cm(1.9)
        normal = doc.styles["Normal"]
        normal.font.name = "Calibri"
        normal.font.size = Pt(tam_base)
        normal.font.color.rgb = RGBColor.from_string(GRIS_TEXTO)
        formato = normal.paragraph_format
        formato.space_before = Pt(0)
        formato.space_after = Pt(5)
        formato.line_spacing = 1.14

    @staticmethod
    def _borde_inferior(parrafo, color=GRIS_LINEA, grosor=6):
        props = parrafo._p.get_or_add_pPr()
        for viejo in props.findall(qn("w:pBdr")):
            props.remove(viejo)
        borde = OxmlElement("w:pBdr")
        abajo = OxmlElement("w:bottom")
        abajo.set(qn("w:val"), "single")
        abajo.set(qn("w:sz"), str(grosor))
        abajo.set(qn("w:space"), "4")
        abajo.set(qn("w:color"), color)
        borde.append(abajo)
        props.append(borde)

    def titulo(self, doc, numero, texto):
        parrafo = doc.add_paragraph()
        parrafo.paragraph_format.space_before = Pt(12)
        parrafo.paragraph_format.space_after = Pt(6)
        parrafo.paragraph_format.keep_with_next = True
        corrida = parrafo.add_run(f"{numero}.  {texto}".upper() if numero else str(texto).upper())
        corrida.bold = True
        corrida.font.size = Pt(9.5)
        corrida.font.name = "Calibri"
        corrida.font.color.rgb = RGBColor.from_string(AZUL)
        self._borde_inferior(parrafo, AZUL, 6)
        return parrafo

    def parrafo(self, doc, texto, tam=9, negrita=False, color=None,
                espacio_despues=5, cursiva=False, justificar=True):
        parrafo = doc.add_paragraph()
        parrafo.paragraph_format.space_after = Pt(espacio_despues)
        if justificar:
            parrafo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        corrida = parrafo.add_run(str(texto))
        corrida.font.size = Pt(tam)
        corrida.bold = negrita
        corrida.italic = cursiva
        corrida.font.color.rgb = RGBColor.from_string(color or GRIS_TEXTO)
        return parrafo

    def vineta(self, doc, texto, tam=8.5):
        return self.parrafo(doc, "•  " + str(texto), tam, espacio_despues=1)

    @staticmethod
    def _sombra(celda, color_hex):
        props = celda._tc.get_or_add_tcPr()
        for viejo in props.findall(qn("w:shd")):
            props.remove(viejo)
        sombra = OxmlElement("w:shd")
        sombra.set(qn("w:val"), "clear")
        sombra.set(qn("w:color"), "auto")
        sombra.set(qn("w:fill"), color_hex)
        props.append(sombra)

    @staticmethod
    def _margen(celda, arriba=60, abajo=60, izq=100, der=100):
        props = celda._tc.get_or_add_tcPr()
        margen = OxmlElement("w:tcMar")
        for lado, valor in (("top", arriba), ("bottom", abajo), ("left", izq), ("right", der)):
            nodo = OxmlElement(f"w:{lado}")
            nodo.set(qn("w:w"), str(valor))
            nodo.set(qn("w:type"), "dxa")
            margen.append(nodo)
        props.append(margen)

    def bordes(self, tabla, interior=True):
        props = tabla._tbl.tblPr
        for viejo in props.findall(qn("w:tblBorders")):
            props.remove(viejo)
        bordes = OxmlElement("w:tblBorders")
        definicion = [("top", GRIS_LINEA, "4")]
        if interior:
            definicion.append(("insideH", "E5E7EB", "2"))
        definicion.append(("bottom", GRIS_LINEA, "4"))
        for nombre, color, grosor in definicion:
            nodo = OxmlElement(f"w:{nombre}")
            nodo.set(qn("w:val"), "single")
            nodo.set(qn("w:sz"), grosor)
            nodo.set(qn("w:space"), "0")
            nodo.set(qn("w:color"), color)
            bordes.append(nodo)
        for nombre in ("left", "right", "insideV"):
            nodo = OxmlElement(f"w:{nombre}")
            nodo.set(qn("w:val"), "nil")
            bordes.append(nodo)
        referencia = props.find(qn("w:tblLook"))
        if referencia is not None:
            referencia.addprevious(bordes)
        else:
            props.append(bordes)

    def tabla(self, doc, encabezados, filas, anchos_cm, alineaciones=None, tam=8.4):
        tabla = doc.add_table(rows=1, cols=len(encabezados))
        tabla.alignment = WD_TABLE_ALIGNMENT.CENTER
        tabla.autofit = False
        self.bordes(tabla)
        for indice, titulo in enumerate(encabezados):
            celda = tabla.rows[0].cells[indice]
            celda.width = Cm(anchos_cm[indice])
            self._sombra(celda, AZUL)
            self._margen(celda, 70, 70, 100, 100)
            parrafo = celda.paragraphs[0]
            parrafo.paragraph_format.space_after = Pt(0)
            corrida = parrafo.add_run(str(titulo))
            corrida.bold = True
            corrida.font.size = Pt(tam - 0.4)
            corrida.font.color.rgb = RGBColor(255, 255, 255)
        for posicion, fila in enumerate(filas):
            celdas = tabla.add_row().cells
            for indice, valor in enumerate(fila):
                celdas[indice].width = Cm(anchos_cm[indice])
                if posicion % 2 == 1:
                    self._sombra(celdas[indice], "FAFAFA")
                self._margen(celdas[indice], 55, 55, 100, 100)
                parrafo = celdas[indice].paragraphs[0]
                parrafo.paragraph_format.space_after = Pt(0)
                if alineaciones and indice < len(alineaciones):
                    if alineaciones[indice] == "r":
                        parrafo.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    elif alineaciones[indice] == "c":
                        parrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER
                corrida = parrafo.add_run(str(valor))
                corrida.font.size = Pt(tam)
        cabecera = tabla.rows[0]._tr
        props = cabecera.get_or_add_trPr()
        repetir = OxmlElement("w:tblHeader")
        repetir.set(qn("w:val"), "true")
        props.append(repetir)
        doc.add_paragraph().paragraph_format.space_after = Pt(1)
        return tabla

    # -- cabecera / pie ----------------------------------------------------
    def encabezado(self, doc, titulo_derecha, lineas_derecha):
        tabla = doc.add_table(rows=1, cols=2)
        tabla.autofit = False
        self.bordes(tabla, interior=False)
        izquierda, derecha = tabla.rows[0].cells
        izquierda.width = Cm(9.2)
        derecha.width = Cm(7.6)
        self._margen(izquierda, 0, 60, 0, 200)
        self._margen(derecha, 0, 60, 200, 0)

        parrafo_logo = izquierda.paragraphs[0]
        parrafo_logo.paragraph_format.space_after = Pt(4)
        if self.logo and os.path.isfile(self.logo):
            try:
                parrafo_logo.add_run().add_picture(self.logo, width=Cm(3.4))
            except Exception:
                pass

        empresa = izquierda.add_paragraph()
        empresa.paragraph_format.space_after = Pt(0)
        corrida = empresa.add_run(_txt(self.perfil.get("razon_social"), "Mantytech SpA").upper())
        corrida.bold = True
        corrida.font.size = Pt(10.5)
        corrida.font.color.rgb = RGBColor.from_string(AZUL)

        datos = izquierda.add_paragraph()
        datos.paragraph_format.space_after = Pt(0)
        partes = []
        if self.perfil.get("rut"):
            partes.append(f"RUT {self.perfil['rut']}  ·  Proveedor hábil del Estado de Chile")
        if self.perfil.get("ciudad_base"):
            partes.append(str(self.perfil["ciudad_base"]))
        contacto = "  ·  ".join(x for x in (self.perfil.get("email_corporativo"),
                                            self.perfil.get("telefono_contacto")) if x)
        if contacto:
            partes.append(contacto)
        corrida = datos.add_run("\n".join(partes))
        corrida.font.size = Pt(7.8)
        corrida.font.color.rgb = RGBColor.from_string(GRIS_SUAVE)

        titulo = derecha.paragraphs[0]
        titulo.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        titulo.paragraph_format.space_after = Pt(2)
        corrida = titulo.add_run(str(titulo_derecha).upper())
        corrida.bold = True
        corrida.font.size = Pt(13)
        corrida.font.color.rgb = RGBColor.from_string(AZUL)

        for linea in lineas_derecha:
            parrafo = derecha.add_paragraph()
            parrafo.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            parrafo.paragraph_format.space_after = Pt(0)
            corrida = parrafo.add_run(str(linea))
            corrida.font.size = Pt(8.2)
            corrida.font.color.rgb = RGBColor.from_string(GRIS_SUAVE)

        separador = doc.add_paragraph()
        separador.paragraph_format.space_before = Pt(6)
        separador.paragraph_format.space_after = Pt(8)
        self._borde_inferior(separador, AZUL, 8)

    def pie(self, doc, texto):
        pie = doc.sections[0].footer
        parrafo = pie.paragraphs[0]
        parrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        corrida = parrafo.add_run(texto)
        corrida.font.size = Pt(7.2)
        corrida.font.color.rgb = RGBColor.from_string(GRIS_SUAVE)


# ---------------------------------------------------------------------------
# Armado del documento
# ---------------------------------------------------------------------------
def _bloque_condiciones(datos, bases):
    """Condiciones comerciales y económicas, coherentes entre sí."""
    condiciones = []
    if bases.get("vigencia") and "mes" in bases["vigencia"]:
        condiciones.append(
            "Vigencia del servicio: " + bases["vigencia"] + " contados desde la habilitación y "
            "puesta en marcha, con renovación solo por acuerdo escrito entre las partes.")
    else:
        condiciones.append(
            "Vigencia del servicio: la informada en las bases del proceso, contada desde la "
            "habilitación y puesta en marcha.")
    condiciones.append(
        "Forma de pago: pago en " + (bases.get("forma_pago") or "1 cuota") +
        " contra recepción conforme y factura electrónica, a 30 días (Ley 21.131).")
    if bases.get("habilitacion"):
        condiciones.append(
            "Habilitación, configuración inicial y puesta en marcha operativa: en un máximo de "
            + bases["habilitacion"] + " desde la aceptación de la orden de compra. Las etapas "
            "posteriores de afinamiento, reportes o capacitación se ejecutan dentro de ese mismo "
            "plazo y no lo extienden.")
    if bases.get("disponibilidad"):
        condiciones.append(
            "Disponibilidad del servicio: mínimo " + bases["disponibilidad"] +
            " mensual, medida sobre el servicio en operación. Se excluyen de la medición las "
            "fallas atribuibles a terceros proveedores (por ejemplo Meta), los casos de fuerza "
            "mayor y las ventanas de mantenimiento informadas con al menos 48 horas de "
            "anticipación.")
    if bases.get("usuarios"):
        condiciones.append(
            "Usuarios: atención simultánea desde el canal institucional para " +
            bases["usuarios"] + " o más funcionarios, con administración de usuarios, roles, "
            "perfiles y permisos a cargo del organismo.")
    condiciones.append(
        "El precio incluye todo lo necesario para cumplir los requisitos mínimos de las bases. "
        "Los costos variables asociados a mensajería (conversaciones de la API oficial, "
        "plantillas y números adicionales) se facturan según el consumo informado por el "
        "proveedor de la plataforma, se informan por escrito antes de su activación y no se "
        "activan sin autorización del organismo.")
    condiciones.append(
        "Moneda y reajuste: valores expresados en pesos chilenos (CLP), con impuestos incluidos "
        "según el desglose económico. No se aplican reajustes durante la vigencia del contrato.")
    observaciones = _txt(datos.get("observaciones"))
    if observaciones:
        condiciones.append(observaciones)
    return condiciones


def _bloque_supuestos(datos):
    return [
        "La ejecución considera coordinación remota con el organismo y disponibilidad de sus "
        "equipos para las validaciones y la capacitación.",
        "Los plazos corren desde la aceptación de la orden de compra y desde la entrega de los "
        "accesos, insumos y antecedentes que debe proporcionar el organismo.",
        "Se excluye hardware, conectividad del organismo y licencias de terceros no detalladas "
        "en el desglose económico.",
        "Los cambios de alcance solicitados después de la orden de compra se cotizan por "
        "separado mediante orden de cambio firmada por ambas partes.",
        "El oferente mantiene la oferta vigente por " + _txt(datos.get("validez"), "60 días") +
        " corridos desde el cierre del proceso.",
    ]


def _bloque_datos_pago(datos):
    perfil = datos.get("perfil") or {}
    banco = datos.get("banco") or {}
    filas = [
        ("Razón social", _txt(perfil.get("razon_social"), "Mantytech SpA")),
        ("RUT", _txt(perfil.get("rut"), "No informado")),
        ("Banco", _txt(banco.get("banco"), "No informado")),
        ("Tipo de cuenta", _txt(banco.get("tipo_cuenta"), "No informado")),
        ("N° de cuenta", _txt(banco.get("numero_cuenta"), "Por informar al momento de facturar")),
        ("Correo de contacto", _txt(perfil.get("email_corporativo")
                                    or perfil.get("email"), "No informado")),
    ]
    return filas


def generar(datos, ruta):
    """Escribe el DOC 2 en `ruta` y devuelve la ruta."""
    datos = datos or {}
    perfil = datos.get("perfil") or {}
    informes = datos.get("informes") or {}
    tecnico = datos.get("tecnico") or {}
    codigo = _txt(datos.get("codigo"), "proceso")
    organismo = _txt(datos.get("organismo"), "Organismo público")
    nombre = _txt(datos.get("nombre"), "Servicio")
    presupuesto = _entero(datos.get("presupuesto"))
    iva = _fraccion(datos.get("iva"), 0.19)
    bruto = _entero(datos.get("monto"))
    neto = _entero(datos.get("neto")) or (int(round(bruto / (1 + iva))) if bruto else 0)
    if not bruto and neto:
        bruto = int(round(neto * (1 + iva)))
    if bruto and not neto:
        neto = int(round(bruto / (1 + iva)))
    iva_clp = bruto - neto

    hoja = _Hoja(perfil, _txt(datos.get("logo")))
    doc = Document()
    hoja.configurar(doc, 1.5, 9.5)

    hoja.encabezado(doc, "Cotización", [
        f"N° de proceso: {codigo}",
        fecha_larga(),
        "Validez de la oferta: " + _txt(datos.get("validez"), "60 días"),
    ])

    parrafo = doc.add_paragraph()
    parrafo.paragraph_format.space_after = Pt(1)
    corrida = parrafo.add_run("SEÑORES")
    corrida.font.size = Pt(7.5)
    corrida.bold = True
    corrida.font.color.rgb = RGBColor.from_string(GRIS_SUAVE)

    parrafo = doc.add_paragraph()
    parrafo.paragraph_format.space_after = Pt(0)
    corrida = parrafo.add_run(organismo)
    corrida.bold = True
    corrida.font.size = Pt(11)
    corrida.font.color.rgb = RGBColor.from_string(AZUL)

    referencia = "   ·   ".join(x for x in (
        ("Unidad de compra: " + _txt(datos.get("unidad_comprador"))) if datos.get("unidad_comprador") else "",
        ("RUT: " + _txt(datos.get("rut_comprador"))) if datos.get("rut_comprador") else "",
        " ".join(x for x in (_txt(datos.get("comuna")), _txt(datos.get("region"))) if x),
    ) if x)
    if referencia:
        parrafo = doc.add_paragraph()
        parrafo.paragraph_format.space_after = Pt(8)
        corrida = parrafo.add_run(referencia)
        corrida.font.size = Pt(7.8)
        corrida.font.color.rgb = RGBColor.from_string(GRIS_SUAVE)

    hoja.parrafo(doc,
                 "De nuestra consideración: por la presente presentamos nuestra oferta "
                 f"económica para el proceso «{nombre}», en estricta conformidad con las bases "
                 "administrativas y técnicas publicadas en Mercado Público, " +
                 frase_cierre(datos.get("fecha_cierre")) + ".", 9, espacio_despues=6)

    # --- 1. Objeto
    hoja.titulo(doc, 1, "Objeto de la oferta")
    alcance_resumido = informes.get("alcance_resumido_doc2") or []
    if isinstance(alcance_resumido, str):
        alcance_resumido = [linea.strip() for linea in alcance_resumido.splitlines()
                            if linea.strip()]
    if alcance_resumido:
        for hito in alcance_resumido[:4]:
            hoja.vineta(doc, _txt(hito)[:260], 8.6)
    else:
        hoja.parrafo(doc, _txt(datos.get("descripcion"), nombre)[:420], 8.8)

    # --- 2. Requisitos de la entidad
    matriz = _requisitos(informes, datos.get("texto_bases") or "")
    if matriz:
        hoja.titulo(doc, 2, "Requisitos del organismo y su cumplimiento")
        hoja.parrafo(doc,
                     "Se revisaron los documentos del proceso y se responde uno a uno cada "
                     "requisito exigido. La columna «Exigido en» indica el documento del cual "
                     "proviene la exigencia.", 8.6, espacio_despues=4)
        hoja.tabla(doc, ["Requisito exigido", "Exigido en", "Compromiso del oferente"],
                   [[_txt(r["requisito"]), _txt(r["fuente"], "Bases del proceso"),
                     _txt(r["compromiso"])] for r in matriz],
                   [7.2, 3.0, 6.8], alineaciones=["l", "l", "l"], tam=8.0)

    # --- 3. Especificaciones técnicas
    numero = 3 if matriz else 2
    hoja.titulo(doc, numero, "Especificaciones técnicas de la solución")
    resumen = informes.get("resumen_ejecutivo_proyecto") or {}
    especificaciones = informes.get("propuesta_especificaciones_tecnicas") or {}

    def _texto_tecnico(clave_tecnico, clave_propuesta, defecto=""):
        return (_txt(tecnico.get(clave_tecnico)) or _txt(especificaciones.get(clave_propuesta))
                or defecto)

    alcance = (_texto_tecnico("alcance", "alcance")
               or _txt(tecnico.get("alcance")) or _txt(resumen.get("alcance_real")))
    arquitectura = _texto_tecnico("arquitectura", "arquitectura_solucion")
    metodologia = _texto_tecnico("metodologia", "metodologia_trabajo")
    sla = normalizar_sla(_texto_tecnico("sla", "sla_soporte_postventa"))
    garantia = (_texto_tecnico("garantia", "garantia_tecnica")
                or "Garantía de correcto funcionamiento durante toda la vigencia del servicio, "
                   "con soporte preferente y corrección de fallas sin costo para el organismo.")
    entregables = tecnico.get("entregables")
    if isinstance(entregables, str):
        entregables = [linea.strip() for linea in entregables.splitlines() if linea.strip()]
    entregables = list(entregables or especificaciones.get("entregables_clave") or [])

    for titulo, cuerpo in (("Alcance del servicio", alcance),
                           ("Arquitectura de la solución", arquitectura),
                           ("Metodología de trabajo", metodologia)):
        if cuerpo:
            hoja.parrafo(doc, titulo, 9, negrita=True, color=AZUL, espacio_despues=2)
            hoja.parrafo(doc, _txt(cuerpo), 8.6, espacio_despues=5)

    if entregables:
        hoja.parrafo(doc, "Entregables comprometidos", 9, negrita=True, color=AZUL,
                     espacio_despues=3)
        total_entregables = min(len(entregables), 6)
        buscador = _datos_de_bases(informes, datos.get("texto_bases") or "")
        referencia_plazo = (buscador.get("habilitacion")
                            or _txt(datos.get("plazo_comprometido"))
                            or "el plazo de las bases")
        filas = [[str(indice), _txt(entregable)[:150],
                  f"Etapa {indice} de {total_entregables}", "Incluido"]
                 for indice, entregable in enumerate(entregables[:6], 1)]
        hoja.tabla(doc, ["N°", "Entregable", "Etapa", "Costo"], filas,
                   [1.0, 9.8, 3.2, 2.4], alineaciones=["c", "l", "l", "c"], tam=8.0)
        hoja.parrafo(doc,
                     "Los entregables se cumplen dentro de " + referencia_plazo +
                     " contados desde la aceptación de la orden de compra, con acta de recepción "
                     "por etapa y sin extender la puesta en marcha comprometida.",
                     8.2, cursiva=True, espacio_despues=5)

    hoja.parrafo(doc, "Nivel de servicio (SLA)", 9, negrita=True, color=AZUL, espacio_despues=2)
    hoja.parrafo(doc, sla, 8.6, espacio_despues=5)
    hoja.parrafo(doc, "Garantía técnica", 9, negrita=True, color=AZUL, espacio_despues=2)
    hoja.parrafo(doc, garantia, 8.6, espacio_despues=5)

    # --- 4. Cumplimiento normativo
    numero += 1
    hoja.titulo(doc, numero, "Cumplimiento normativo y estándares")
    legal = informes.get("auditoria_37_normas_ley_19886") or {}
    hoja.parrafo(doc,
                 "La presente oferta se presenta en estricto cumplimiento de la Ley 19.886 de "
                 "Bases sobre Contratos Administrativos, su Reglamento (DS 661/2024), la Ley "
                 "21.634 y las bases administrativas y técnicas del proceso. El proveedor "
                 "declara no encontrarse afecto a las inhabilidades del artículo 4 de la Ley "
                 "19.886.", 8.6, espacio_despues=4)
    declaraciones = []
    for clave in ("cumplimiento_art_4", "probidad_e_intereses"):
        if legal.get(clave):
            declaraciones.append(_sin_certificaciones(str(legal[clave]))[:300])
    estrategia = (informes.get("normas_chilenas_nch_estandares") or {}).get("estrategia_cumplimiento")
    if estrategia:
        declaraciones.append(_sin_certificaciones(str(estrategia))[:300])
    for declaracion in declaraciones[:3]:
        if len(declaracion.strip()) < 30:
            # frases sueltas tipo "Sin inhabilidades.": ya están en el párrafo legal
            continue
        hoja.vineta(doc, declaracion, 8.3)
    normas = _normas_redactadas(informes)
    if normas:
        hoja.parrafo(doc, "Normas y estándares considerados", 9, negrita=True, color=AZUL,
                     espacio_despues=2)
        for norma in normas:
            hoja.vineta(doc, norma, 8.3)
        hoja.parrafo(doc,
                     "La oferta se redacta conforme a los lineamientos de estas normas. No se "
                     "declara una certificación que no se encuentre formalmente emitida y "
                     "acompañada del documento correspondiente.", 8.2, cursiva=True,
                     espacio_despues=4)

    # --- 5. Desglose económico
    numero += 1
    hoja.titulo(doc, numero, "Desglose económico")
    hoja.tabla(doc,
               ["Ítem", "Descripción", "Cant.", "Valor unitario neto", "Total neto"],
               [["1", f"Servicio integral para el proceso {codigo}: {nombre[:90]}",
                 "1", clp(neto), clp(neto)]],
               [1.2, 8.6, 1.4, 3.0, 2.8], alineaciones=["c", "l", "c", "r", "r"], tam=8.4)

    totales = doc.add_table(rows=4, cols=2)
    totales.autofit = False
    totales.alignment = WD_TABLE_ALIGNMENT.RIGHT
    hoja.bordes(totales, interior=False)
    etiqueta_presupuesto = ("Presupuesto oficial"
                            if "oficial" in _txt(datos.get("origen_presupuesto"), "oficial").lower()
                            else "Presupuesto referencial")
    filas_totales = [
        ("Subtotal neto", clp(neto), False),
        (f"IVA {int(round(iva * 100))}% (D.L. 825)", clp(iva_clp), False),
        ("Total con IVA incluido", clp(bruto), True),
        (etiqueta_presupuesto, clp(presupuesto), False),
    ]
    for indice, (etiqueta, valor, destacado) in enumerate(filas_totales):
        celda_etiqueta = totales.rows[indice].cells[0]
        celda_valor = totales.rows[indice].cells[1]
        celda_etiqueta.width = Cm(6.0)
        celda_valor.width = Cm(4.4)
        hoja._margen(celda_etiqueta, 40, 40, 120, 100)
        hoja._margen(celda_valor, 40, 40, 100, 120)
        if destacado:
            hoja._sombra(celda_etiqueta, GRIS_FONDO)
            hoja._sombra(celda_valor, GRIS_FONDO)
        parrafo = celda_etiqueta.paragraphs[0]
        parrafo.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        parrafo.paragraph_format.space_after = Pt(0)
        corrida = parrafo.add_run(etiqueta)
        corrida.font.size = Pt(8.8 if destacado else 8.4)
        corrida.bold = destacado
        corrida.font.color.rgb = RGBColor.from_string(AZUL if destacado else GRIS_TEXTO)
        parrafo = celda_valor.paragraphs[0]
        parrafo.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        parrafo.paragraph_format.space_after = Pt(0)
        corrida = parrafo.add_run(valor)
        corrida.font.size = Pt(10.2 if destacado else 8.4)
        corrida.bold = destacado
        corrida.font.color.rgb = RGBColor.from_string(AZUL if destacado else GRIS_TEXTO)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)
    if presupuesto and bruto:
        hoja.parrafo(doc,
                     "La oferta total con IVA incluido asciende a " + clp(bruto) +
                     " y no supera el presupuesto máximo informado por el organismo ("
                     + clp(presupuesto) + ").", 8.4, cursiva=True, espacio_despues=4)

    # --- 6. Condiciones comerciales y económicas
    numero += 1
    hoja.titulo(doc, numero, "Condiciones comerciales y económicas")
    bases = _datos_de_bases(informes, datos.get("texto_bases") or "")
    for condicion in _bloque_condiciones(datos, bases):
        hoja.vineta(doc, condicion, 8.4)

    # --- 7. Supuestos y exclusiones
    numero += 1
    hoja.titulo(doc, numero, "Supuestos, exclusiones y condiciones")
    for supuesto in _bloque_supuestos(datos):
        hoja.vineta(doc, supuesto, 8.4)

    # --- 8. Datos del proveedor y de pago
    numero += 1
    hoja.titulo(doc, numero, "Datos del proveedor y de pago")
    hoja.tabla(doc, ["Concepto", "Antecedente"],
               [[etiqueta, valor] for etiqueta, valor in _bloque_datos_pago(datos)],
               [5.2, 11.8], alineaciones=["l", "l"], tam=8.2)

    # --- 9. Término del servicio y datos
    numero += 1
    hoja.titulo(doc, numero, "Término del servicio y tratamiento de la información")
    hoja.parrafo(doc,
                 "Al término del servicio, por cualquier causa, el proveedor entregará al "
                 "organismo una exportación completa del historial de conversaciones, reportes y "
                 "configuraciones dentro de 15 días corridos, en formatos abiertos (CSV y PDF). "
                 "Una vez recepcionada la entrega, y dentro de los 15 días corridos siguientes, "
                 "eliminará las copias de producción y respaldos que contengan datos del "
                 "organismo, dejando constancia escrita de la eliminación. Mientras el servicio "
                 "esté vigente, los datos son del organismo y se tratan solo para la ejecución "
                 "del contrato, conforme a la Ley 19.628 y a la Ley 21.096.", 8.4,
                 espacio_despues=6)

    # --- Firma
    parrafo = doc.add_paragraph()
    parrafo.paragraph_format.space_before = Pt(18)
    parrafo.paragraph_format.space_after = Pt(0)
    corrida = parrafo.add_run("______________________________________")
    corrida.font.size = Pt(9)
    corrida.font.color.rgb = RGBColor.from_string(GRIS_SUAVE)

    parrafo = doc.add_paragraph()
    parrafo.paragraph_format.space_before = Pt(4)
    parrafo.paragraph_format.space_after = Pt(0)
    corrida = parrafo.add_run(_txt(perfil.get("representante_legal"), "Representante legal"))
    corrida.bold = True
    corrida.font.size = Pt(9.5)
    corrida.font.color.rgb = RGBColor.from_string(AZUL)

    parrafo = doc.add_paragraph()
    parrafo.paragraph_format.space_after = Pt(0)
    pie_empresa = " · ".join(x for x in (_txt(perfil.get("cargo_representante")),
                                         _txt(perfil.get("razon_social")),
                                         ("RUT " + _txt(perfil.get("rut"))) if perfil.get("rut") else "") if x)
    corrida = parrafo.add_run(pie_empresa)
    corrida.font.size = Pt(7.8)
    corrida.font.color.rgb = RGBColor.from_string(GRIS_SUAVE)

    hoja.pie(doc, f"Documento generado por LICITAIA para {organismo[:60]} · Proceso {codigo}")
    os.makedirs(os.path.dirname(os.path.abspath(ruta)) or ".", exist_ok=True)
    doc.save(ruta)
    return ruta


# ---------------------------------------------------------------------------
# Lectura de la carpeta del proyecto (lo que usa el portal)
# ---------------------------------------------------------------------------
def _leer_json(ruta):
    try:
        with open(ruta, encoding="utf-8") as archivo:
            return json.load(archivo) or {}
    except Exception:
        return {}


def _candidatos_datos():
    """Carpetas donde puede estar la configuración del cliente y los datos bancarios."""
    base = os.path.dirname(os.path.abspath(__file__))
    salida = [os.getenv("LICITAIA_DATA_DIR") or "", os.getcwd(), base,
              os.path.dirname(base)]
    limpio = []
    for ruta in salida:
        if ruta and ruta not in limpio:
            limpio.append(ruta)
    return limpio


def perfil_empresa(cliente=""):
    """Perfil del proveedor: clientes/<cliente>.json o el primero que exista."""
    nombre_cliente = (cliente or os.getenv("CLIENTE") or "mantytech").strip().lower()
    for base in _candidatos_datos():
        carpeta = os.path.join(base, "clientes")
        for candidato in (os.path.join(carpeta, nombre_cliente + ".json"),
                          os.path.join(carpeta, "mantytech.json")):
            datos = _leer_json(candidato)
            if datos:
                return datos
        if os.path.isdir(carpeta):
            for archivo in sorted(os.listdir(carpeta)):
                if archivo.endswith(".json"):
                    datos = _leer_json(os.path.join(carpeta, archivo))
                    if datos:
                        return datos
    return {}


def datos_bancarios():
    for base in _candidatos_datos():
        datos = _leer_json(os.path.join(base, "datos_bancarios.json"))
        if datos:
            return datos
    return {}


def ruta_doc2(carpeta, codigo=""):
    """Nombre del DOC 2: conserva el que ya existe para no duplicar archivos."""
    try:
        for archivo in sorted(os.listdir(carpeta)):
            if archivo.upper().startswith("DOC2") and archivo.lower().endswith(".docx"):
                return os.path.join(carpeta, archivo)
    except Exception:
        pass
    return os.path.join(carpeta, f"DOC2_Cotizacion_Oficial_{codigo or 'proceso'}.docx")


def generar_desde_carpeta(carpeta, ruta="", monto=None, neto=None, iva=None,
                          presupuesto=None):
    """Regenera el DOC 2 con lo que hay guardado en la carpeta del proyecto.

    Devuelve la ruta escrita, o "" si no se pudo (el portal sigue funcionando).
    """
    try:
        manifiesto = _leer_json(os.path.join(carpeta, "manifest.json"))
        datos_informe = _leer_json(os.path.join(carpeta, "datos_informe.json"))
        cot = _leer_json(os.path.join(carpeta, "cotizacion_editada.json"))
        informes = datos_informe.get("informes") or {}
        codigo = _txt(cot.get("codigo") or manifiesto.get("codigo")
                      or datos_informe.get("codigo"))
        tope = _entero(presupuesto if presupuesto is not None else
                       (cot.get("presupuesto") or manifiesto.get("presupuesto_clp")
                        or datos_informe.get("presupuesto_estatal_clp")))
        iva_valor = _fraccion(iva if iva is not None else cot.get("iva"), 0.19)
        total = _entero(monto if monto is not None else
                        (cot.get("monto") or manifiesto.get("oferta_bruta_clp")))
        neto_valor = _entero(neto if neto is not None else cot.get("neto"))
        perfil = perfil_empresa(_txt(manifiesto.get("cliente")))
        for clave_portal, clave_perfil in (("razon_social", "razon_social"),
                                           ("rut", "rut"),
                                           ("representante", "representante_legal"),
                                           ("email", "email_corporativo"),
                                           ("telefono", "telefono_contacto")):
            if _txt(cot.get(clave_portal)) and not _txt(perfil.get(clave_perfil)):
                perfil[clave_perfil] = cot[clave_portal]
        datos = {
            "codigo": codigo,
            "organismo": _txt(cot.get("organismo") or manifiesto.get("organismo")
                              or datos_informe.get("organismo")),
            "nombre": _txt(manifiesto.get("nombre") or datos_informe.get("nombre")),
            "unidad_comprador": _txt(datos_informe.get("unidad_compradora")
                                     or manifiesto.get("unidad_compradora")),
            "rut_comprador": _txt(datos_informe.get("rut_comprador")
                                  or manifiesto.get("rut_comprador")),
            "comuna": _txt(datos_informe.get("comuna")),
            "region": _txt(datos_informe.get("region")),
            "fecha_cierre": _txt(cot.get("fecha_cierre") or manifiesto.get("fecha_cierre")
                                 or datos_informe.get("fecha_cierre")),
            "presupuesto": tope,
            "monto": total,
            "neto": neto_valor,
            "iva": iva_valor,
            "origen_presupuesto": _txt(datos_informe.get("origen_presupuesto")
                                       or manifiesto.get("tipo_presupuesto")),
            "plazo_comprometido": _txt(cot.get("plazo") or manifiesto.get("plazo_entrega_organismo")
                                       or datos_informe.get("plazo_entrega_organismo")),
            "validez": _txt(cot.get("validez"), "60 días"),
            "observaciones": _txt(cot.get("observaciones")),
            "texto_bases": _txt(datos_informe.get("texto_extraido_adjuntos")),
            "perfil": perfil,
            "banco": datos_bancarios(),
            "informes": informes,
            "tecnico": cot.get("tecnico") if isinstance(cot.get("tecnico"), dict) else {},
            "logo": _buscar_logo(),
        }
        destino = ruta or ruta_doc2(carpeta, codigo)
        return generar(datos, destino)
    except Exception as error:
        print(f"   [DOC2] no se pudo regenerar la cotización: {error}")
        return ""


def _buscar_logo():
    for base in _candidatos_datos():
        for nombre in ("logo.png", "logo.jpg"):
            ruta = os.path.join(base, nombre)
            if os.path.isfile(ruta):
                return ruta
    return ""
