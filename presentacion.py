"""
Presentación financiera del mes en PowerPoint, con el mismo fondo y la misma tipografía de la presentación de la
compañía (plantilla_presentacion.pptx: franja naranja, logo y título en Futura).

Abre con un resumen del mes (indicadores clave y lo más importante) y sigue con cada tema en diapositivas aparte:
presupuesto, nómina, flujo de caja, combustible, balance e indicadores, y estado de resultados. Cada diapositiva
lleva un titular con su mensaje principal, tarjetas de resumen y gráficas de variaciones (verde a favor, rojo en
contra). Tablas, gráficas y comentarios quedan editables en PowerPoint (los comentarios se redactan solos con las
mayores variaciones del mes y son un punto de partida).

    datos, avisos = presentacion.generar(leer, "ALDC", 2026, 9, "8+4")
"""
import copy
from io import BytesIO
from pathlib import Path

import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION, XL_TICK_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

import cargas
import datos
import eeff
import nomina
import presupuestos as P

PLANTILLA = Path(__file__).with_name("plantilla_presentacion.pptx")
SECCIONES = ["Portada", "Resumen del mes", "Presupuesto", "Nómina", "Flujo de caja", "Combustible", "Facturación de aseo",
             "Corte de césped", "Consumo operativo", "Operaciones por ítem", "Mantenimiento de flota", "Flota por placa", "Balance e indicadores", "Estado de resultados"]
MESES = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"]
MESES_L = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre",
           "noviembre", "diciembre"]
EMPRESAS = {"ALDC": "Área Limpia D.C.", "ALSM": "Área Limpia Servicios Medioambientales"}

# tipografía y colores de la presentación
F_TIT, F_TXT = "Futura Ultra-Bold", "Futura Light"
NAVY, AZUL, NARANJA, ROJO, VERDE = "1F3864", "0070C0", "F38E28", "A5281B", "2E7D32"
GRIS, NEGRO, BLANCO, AZUL_TXT, GRIS_OSC = "F4F5F7", "000000", "FFFFFF", "003A5D", "A6A6A6"
VERDE_BAR, ROJO_BAR = "4CAF50", "E53935"
VERDE_CL, ROJO_CL = "A5E6A8", "FFB4A9"            # sobre fondo azul oscuro
ANCHO, ALTO = 18288000, 10287000
Y0 = 2000000                                      # donde empieza el contenido (debajo del titular)


# ------------------------------------------------------------------ formatos
def _num(v, div: float = 1.0, dec: int = 0) -> str:
    if v is None or pd.isna(v):
        return ""
    v = v / div
    if abs(v) < 0.5 * 10 ** -dec:
        return "-"
    s = f"{abs(v):,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("-" if v < 0 else "") + s


def _M(v, signo: bool = False) -> str:
    """Millones con un decimal: $49,2M (con signo: +34,8M)."""
    if v is None or pd.isna(v):
        return "—"
    s = f"{abs(v) / 1e6:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")
    if signo:
        return ("+" if v >= 0 else "-") + s + "M"
    return ("-" if v < 0 else "") + "$" + s + "M"


def _pct(p, dec: int = 1, signo: bool = False) -> str:
    if p is None or pd.isna(p):
        return "—"
    s = f"{abs(p) * 100:.{dec}f}".replace(".", ",") + "%"
    return (("+" if p >= 0 else "-") if signo else ("-" if p < 0 else "")) + s


def _ent(v) -> str:
    """Entero con signo: +44, -8."""
    return ("+" if v > 0 else "") + _num(v) if v else "0"


def _pp(d) -> str:
    return "—" if d is None or pd.isna(d) else f"{d * 100:+.1f}".replace(".", ",") + " pp"


def _x(v) -> str:
    return "—" if v is None or pd.isna(v) else f"{v:.2f}".replace(".", ",") + "x"


def _dias(v) -> str:
    return "—" if v is None or pd.isna(v) else f"{v:,.0f}".replace(",", ".") + " días"


def _div(a, b):
    return a / b if b and not pd.isna(b) and abs(b) > 1e-9 and not pd.isna(a) else float("nan")


def _corto(t: str, n: int = 34) -> str:
    t = " ".join(str(t).split())
    return t if len(t) <= n else t[: n - 1].rstrip() + "…"


def _lista(pares, n: int = 3) -> str:
    return ", ".join(f"{_corto(k, 40).lower()} ({_M(v, True)})" for k, v in pares[:n])


# ------------------------------------------------------------------ piezas de la diapositiva
def _fuente(run, nombre, tam, negrita=False, color=NEGRO):
    f = run.font
    f.name, f.size, f.bold = nombre, Pt(tam), negrita
    f.color.rgb = RGBColor.from_string(color)


def _clonar(prs, base, titulo: str, mensaje: str = ""):
    """Nueva diapositiva con el mismo fondo de la base (franja, logo y título) y, debajo del título, el titular con
    el mensaje principal de la diapositiva."""
    s = prs.slides.add_slide(base.slide_layout)
    for sh in list(s.shapes):
        sh._element.getparent().remove(sh._element)
    mapa = {}
    for rid, rel in base.part.rels.items():
        if "slideLayout" in rel.reltype or "notesSlide" in rel.reltype or rel.is_external:
            continue
        mapa[rid] = s.part.relate_to(rel._target, rel.reltype)
    for sh in base.shapes:
        el = copy.deepcopy(sh._element)
        for nodo in el.iter():
            for k, v in list(nodo.attrib.items()):
                if k.startswith("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}") \
                        and v in mapa:
                    nodo.set(k, mapa[v])
        s.shapes._spTree.append(el)
    for sh in s.shapes:
        if sh.name == "TextBox 8":
            sh.left, sh.width = Emu(571500), Emu(ANCHO - 2 * 571500)
            r = sh.text_frame.paragraphs[0].runs[0]
            r.text = titulo.upper()
            r.font.size = Pt(48 if len(titulo) <= 32 else (40 if len(titulo) <= 42 else 32))
    if mensaje:
        tb = s.shapes.add_textbox(Emu(1200000), Emu(1480000), Emu(ANCHO - 2400000), Emu(420000))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        _fuente(p.add_run(), F_TXT, 19 if len(mensaje) <= 120 else 16, True, AZUL_TXT)
        p.runs[0].text = mensaje
    return s


def _tarjeta(s, x, y, w, h, etiqueta, valor, sub="", color_sub=ROJO, borde=False, tam=24):
    """Tarjeta gris (o blanca con borde azul) como las de la presentación: etiqueta, valor grande y nota."""
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Emu(x), Emu(y), Emu(w), Emu(h))
    sh.adjustments[0] = 0.08
    sh.fill.solid()
    sh.fill.fore_color.rgb = RGBColor.from_string(BLANCO if borde else GRIS)
    if borde:
        sh.line.color.rgb = RGBColor.from_string(NAVY)
        sh.line.width = Pt(1.5)
    else:
        sh.line.fill.background()
    sh.shadow.inherit = False
    tf = sh.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Emu(45720)
    for i, (txt, fn, tm, col) in enumerate([(etiqueta, F_TXT, round(tam * .72), AZUL_TXT), (valor, F_TIT, tam, NEGRO),
                                            (sub, F_TXT, round(tam * .58), color_sub)]):
        if not txt:
            continue
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        _fuente(p.add_run(), fn, tm, True, col)
        p.runs[0].text = str(txt)
    return sh


def _kpi(s, x, y, w, h, valor, etiqueta, sub="", bueno=None, tam=30):
    """Recuadro azul oscuro como los indicadores de la presentación (valor grande en blanco y su nombre abajo)."""
    sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(x), Emu(y), Emu(w), Emu(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = RGBColor.from_string(NAVY)
    sh.line.fill.background()
    tf = sh.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = Emu(60000)
    tf.margin_top = tf.margin_bottom = Emu(30000)
    col_sub = BLANCO if bueno is None else (VERDE_CL if bueno else ROJO_CL)
    for i, (txt, fn, tm, col) in enumerate([(valor, F_TIT, tam, BLANCO), (etiqueta, F_TXT, round(tam * .5), BLANCO),
                                            (sub, F_TXT, round(tam * .45), col_sub)]):
        if not txt:
            continue
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        _fuente(p.add_run(), fn, tm, True, col)
        p.runs[0].text = str(txt)
    return sh


def _texto(s, x, y, w, h, parrafos, tam=18, titulo=None, alinear=PP_ALIGN.JUSTIFY, viñeta=False):
    """parrafos: lista de textos o de (negrita, texto)."""
    tb = s.shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    primero = True
    if titulo:
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        _fuente(p.add_run(), F_TXT, tam, True)
        p.runs[0].text = titulo
        p.space_after = Pt(tam * .6)
        primero = False
    for par in parrafos:
        if not par:
            continue
        p = tf.paragraphs[0] if primero else tf.add_paragraph()
        primero = False
        p.alignment = alinear
        p.space_after = Pt(tam * .4)
        neg, txt = par if isinstance(par, tuple) else ("", par)
        if viñeta:
            _fuente(p.add_run(), F_TXT, tam, True, NARANJA)
            p.runs[-1].text = "■  "
        if neg:
            _fuente(p.add_run(), F_TXT, tam, True)
            p.runs[-1].text = neg + " "
        _fuente(p.add_run(), F_TXT, tam)
        p.runs[-1].text = txt
    return tb


def _nota(s, x, y, txt, w=6000000):
    return _texto(s, x, y, w, 300000, [txt], tam=12, alinear=PP_ALIGN.LEFT)


def _tabla(s, x, y, w, encabezados, filas, anchos, alto_fila=235000, tam=11, grupos=None):
    """filas: lista de dict(celdas=[...], tipo='' | 'total' | 'seccion' | 'sub', rojo={índices}, verde={índices}).
    grupos: fila superior opcional [(texto, columnas que abarca), ...]."""
    n_enc = 2 if grupos else 1
    gf = s.shapes.add_table(len(filas) + n_enc, len(encabezados), Emu(x), Emu(y), Emu(w),
                            Emu(alto_fila * (len(filas) + n_enc)))
    tbl = gf.table
    tblPr = gf._element.graphic.graphicData.tbl.tblPr
    for a in ("firstRow", "bandRow"):
        tblPr.set(a, "0")
    estilo = tblPr.find(qn("a:tableStyleId"))
    if estilo is not None:
        estilo.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"      # sin estilo, sin cuadrícula
    tot = float(sum(anchos))
    for i, a in enumerate(anchos):
        tbl.columns[i].width = Emu(int(w * a / tot))

    def celda(c, txt, fondo, color, negrita, alin, tm=tam):
        c.fill.solid()
        c.fill.fore_color.rgb = RGBColor.from_string(fondo)
        c.margin_left = c.margin_right = Emu(54000)
        c.margin_top = c.margin_bottom = Emu(12000)
        c.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = c.text_frame.paragraphs[0]
        p.alignment = alin
        p.font.size = Pt(tm)                                   # las celdas vacías no agrandan la fila
        _fuente(p.add_run(), F_TXT, tm, negrita, color)
        p.runs[0].text = " " if txt is None or txt == "" else str(txt)

    fila0 = 0
    if grupos:
        j = 0
        for txt, n in grupos:
            if n > 1:
                tbl.cell(0, j).merge(tbl.cell(0, j + n - 1))
            celda(tbl.cell(0, j), txt, NAVY, BLANCO, True, PP_ALIGN.CENTER)
            j += n
        fila0 = 1
    for j, e in enumerate(encabezados):
        celda(tbl.cell(fila0, j), e, NAVY, BLANCO, True, PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER)
    for i, f in enumerate(filas, start=n_enc):
        tipo = f.get("tipo", "")
        fondo = NAVY if tipo == "total" else ("DCE6F2" if tipo == "sub" else BLANCO)
        for j, v in enumerate(f["celdas"]):
            if tipo == "total":
                color = BLANCO
            else:
                color = ROJO if j in f.get("rojo", ()) else (VERDE if j in f.get("verde", ()) else NEGRO)
            celda(tbl.cell(i, j), v, fondo, color, tipo in ("total", "seccion", "sub"),
                  PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT)
    for r in tbl.rows:
        r.height = Emu(alto_fila)
    return gf


def _estilo_ejes(ch, formato, tam):
    va = ch.value_axis
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = RGBColor.from_string("E3E3E3")
    va.format.line.fill.background()
    va.tick_labels.number_format, va.tick_labels.number_format_is_linked = formato, False
    va.tick_labels.font.size = Pt(tam - 2)
    ch.category_axis.tick_labels.font.size = Pt(tam)
    ch.category_axis.format.line.color.rgb = RGBColor.from_string("BFBFBF")


def _titulo_grafico(ch, titulo, tam):
    ch.has_title = bool(titulo)
    if titulo:
        ch.chart_title.text_frame.text = titulo
        _fuente(ch.chart_title.text_frame.paragraphs[0].runs[0], F_TXT, tam + 2, True, AZUL_TXT)


def _grafico(s, x, y, w, h, categorias, series, titulo=None, resaltar_ultimo=True, formato="#,##0", tam=12,
             colores=(AZUL, NARANJA, NAVY), desde_cero=False):
    """Columnas agrupadas. series: {nombre: [valores]}. Con una sola serie, la última barra va en naranja."""
    cd = CategoryChartData()
    cd.categories = list(categorias)
    for nombre, vals in series.items():
        cd.add_series(nombre, [None if v is None or pd.isna(v) else float(v) for v in vals])
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Emu(x), Emu(y), Emu(w), Emu(h), cd)
    ch = gf.chart
    ch.font.name, ch.font.size = F_TXT, Pt(tam)
    _titulo_grafico(ch, titulo, tam)
    ch.has_legend = len(series) > 1
    if ch.has_legend:
        ch.legend.position = XL_LEGEND_POSITION.BOTTOM
        ch.legend.include_in_layout = False
        ch.legend.font.size = Pt(tam)
    pl = ch.plots[0]
    pl.gap_width = 60
    pl.overlap = -10 if len(series) > 1 else 0
    pl.has_data_labels = True
    dl = pl.data_labels
    dl.number_format, dl.number_format_is_linked = formato, False
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    dl.font.size, dl.font.name = Pt(tam - 1), F_TXT
    for i, ser in enumerate(pl.series):
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = RGBColor.from_string(colores[i % len(colores)])
        ser.invert_if_negative = False
        if len(series) == 1 and resaltar_ultimo and categorias:
            pt = ser.points[len(categorias) - 1]
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = RGBColor.from_string(NARANJA)
    _estilo_ejes(ch, formato, tam)
    if desde_cero:                                      # diferencias pequeñas no deben verse como enormes
        ch.value_axis.minimum_scale = 0
    return gf


def _barras_var(s, x, y, w, h, pares, titulo=None, tam=12, formato='+#,##0.0;-#,##0.0;0', bueno_si_sube=True):
    """Barras horizontales de variación (millones), de mayor a menor: verde lo que va a favor y rojo lo que va en
    contra (bueno_si_sube=False para gastos: subir es rojo)."""
    pares = [(n, v) for n, v in pares if v is not None and not pd.isna(v)]
    if not pares:
        return None
    cd = CategoryChartData()
    cd.categories = [_corto(n, 32) for n, _ in pares]
    cd.add_series("Variación", [float(v) / 1e6 for _, v in pares])
    gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Emu(x), Emu(y), Emu(w), Emu(h), cd)
    ch = gf.chart
    ch.font.name, ch.font.size = F_TXT, Pt(tam)
    _titulo_grafico(ch, titulo, tam)
    ch.has_legend = False
    pl = ch.plots[0]
    pl.gap_width = 45
    pl.has_data_labels = True
    dl = pl.data_labels
    dl.number_format, dl.number_format_is_linked = formato, False
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    dl.font.size, dl.font.name = Pt(tam - 1), F_TXT
    ser = pl.series[0]
    ser.invert_if_negative = False
    for i, (_, v) in enumerate(pares):
        pt = ser.points[i]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = RGBColor.from_string(VERDE_BAR if (v >= 0) == bueno_si_sube else ROJO_BAR)
    _estilo_ejes(ch, '#,##0', tam)
    vals = [float(v) / 1e6 for _, v in pares]
    lo, hi = min(0.0, min(vals)), max(0.0, max(vals))
    holgura = (hi - lo) * 0.28 or 1.0
    va = ch.value_axis
    va.minimum_scale = lo - (holgura if lo < 0 else 0)
    va.maximum_scale = hi + (holgura if hi > 0 else 0)
    va.has_major_gridlines = False
    va.tick_label_position = XL_TICK_LABEL_POSITION.NONE
    ca = ch.category_axis
    ca.reverse_order = True                              # el primero arriba
    ca.tick_label_position = XL_TICK_LABEL_POSITION.LOW  # nombres siempre a la izquierda
    return gf


def _fila_tarjetas(s, y, tarjetas, x0=500000, x1=ANCHO - 500000, alto=1180000, tam=22, sep=180000):
    """Fila de tarjetas grises del mismo ancho: [(etiqueta, valor, nota, bueno), ...]."""
    n = len(tarjetas)
    w = int((x1 - x0 - sep * (n - 1)) / n)
    for i, (et, val, sub, bueno) in enumerate(tarjetas):
        _tarjeta(s, x0 + i * (w + sep), y, w, alto, et, val, sub,
                 AZUL_TXT if bueno is None else (VERDE if bueno else ROJO), tam=tam)


# ------------------------------------------------------------------ presupuesto
def _nota_version(R, seccion: str, usada):
    """Si una sección no pudo usar exactamente el presupuesto elegido (no está cargado en ese módulo, o se cargó sin versión), lo deja anotado
    para avisarlo al terminar: la presentación nunca compara en silencio contra otro presupuesto."""
    v = R["version"]
    if usada == v or usada is None:
        return
    R["notas"].append(f"{seccion}: el presupuesto cargado no tiene versión; se comparó contra él como {P.etiqueta(v)}." if usada == ""
                      else f"{seccion}: no hay {P.etiqueta(v)} cargado en este módulo; se comparó contra el {P.etiqueta(usada)}.")


def _datos_ppto(leer, empresa, anio, mes, version):
    ppto = leer("SELECT codigo, sub_rubro, mes, valor FROM presupuesto WHERE anio=? AND version=? AND empresa=?",
                (anio, version, empresa))
    ejec = leer("SELECT codigo, sub_rubro, mes, valor FROM ejecucion WHERE anio=? AND empresa=? AND codigo<>''",
                (anio, empresa))
    if ppto.empty or ejec.empty or not (ejec["mes"] == mes).any():
        return None
    ejec = ejec[ejec["codigo"].isin(list(cargas.CLASES))]
    ejec = datos.homologar(ejec, ppto, leer("SELECT * FROM homologacion"), version)
    return ppto, ejec


def _cuadro(ppto, ejec, mes, por):
    """Presupuesto y ejecución del mes y acumulados, en signo de resultado (ingreso +, gasto −)."""
    def g(df, acum):
        d = df[df["mes"] <= mes] if acum else df[df["mes"] == mes]
        return -d.groupby(por)["valor"].sum()
    t = pd.DataFrame({"em": g(ejec, False), "ea": g(ejec, True), "pm": g(ppto, False), "pa": g(ppto, True),
                      "anual": -ppto.groupby(por)["valor"].sum()}).fillna(0.0)
    return t


def _fila_ppto(nombre, r, tipo="", signo=None):
    """Variación como en la presentación: en ingresos y subtotales, ejecutado − presupuesto; en gastos, cuánto más
    (o menos) se gastó. En rojo lo que va en contra, en verde lo que va a favor."""
    if signo is None:
        ref = r["anual"] if abs(r["anual"]) > 1 else (r["pa"] if abs(r["pa"]) > 1 else r["ea"])
        signo = 1 if ref >= 0 else -1
    vm, va = signo * (r["em"] - r["pm"]), signo * (r["ea"] - r["pa"])
    rojo = {j for j, d in ((5, r["em"] - r["pm"]), (6, r["ea"] - r["pa"])) if d < -0.5}
    verde = {j for j, d in ((5, r["em"] - r["pm"]), (6, r["ea"] - r["pa"])) if d > 0.5}
    return dict(celdas=[nombre, _num(r["em"]), _num(r["ea"]), _num(r["pm"]), _num(r["pa"]), _num(vm), _num(va)],
                tipo=tipo, rojo=rojo, verde=verde)


def _comentarios_ppto(ppto, ejec, mes, clases, n=5):
    """Las clases con mayor variación del mes y los sub-rubros que la explican."""
    sub = _cuadro(ppto, ejec, mes, ["codigo", "sub_rubro"])
    sub["d"] = sub["em"] - sub["pm"]                       # + a favor del resultado
    out = []
    orden = clases.assign(d=clases["em"] - clases["pm"])
    orden = orden[orden["d"].abs() > 1e5].reindex(orden["d"].abs().sort_values(ascending=False).index).dropna()
    for cod, r in orden.head(n).iterrows():
        es_ing = r["anual"] >= 0 if abs(r["anual"]) > 1 else r["em"] >= 0
        d = sub.loc[cod, "d"] if cod in sub.index.get_level_values(0) else pd.Series(dtype=float)
        d = d[d.abs() > 5e4]
        nat = d if es_ing else -d                           # en gastos: + es mayor gasto
        total = r["d"] if es_ing else -r["d"]
        mismos = nat[nat * total > 0].sort_values(key=abs, ascending=False)
        contra = nat[nat * total < 0].sort_values(key=abs, ascending=False)
        txt = f"{_M(total, True)} frente al presupuesto del mes"
        if len(mismos):
            txt += ": " + _lista(list(mismos.items()))
            if len(contra):
                txt += ", compensados con " + _lista(list(contra.items()), 2)
        out.append((cargas.CLASES.get(cod, cod) + ".", txt + "."))
    return out


def _diap_presupuesto(prs, base, leer, empresa, anio, mes, R):
    version = R["version"]
    d = _datos_ppto(leer, empresa, anio, mes, version)
    if d is None:
        return f"Presupuesto: no hay presupuesto {version} o ejecución de {MESES_L[mes - 1]} {anio} para {empresa}."
    ppto, ejec = d
    ml = MESES_L[mes - 1]
    t = _cuadro(ppto, ejec, mes, "codigo")
    t = t[(t[["em", "ea", "pm", "pa"]].abs() > 0.5).any(axis=1)]
    t = t.reindex(sorted(t.index, key=cargas.orden_clase))

    def total(cods):
        return t[t.index.isin(cods)].sum()
    ing = t.loc["1.0"] if "1.0" in t.index else pd.Series(0.0, index=t.columns)
    ebi = total(cargas.SUBTOTALES["EBITDA"])
    cyg = total([c for c in cargas.EBITDA if c != "1.0"])
    neto = total(list(cargas.CLASES))
    me_m, me_a = _div(ebi["em"], ing["em"]), _div(ebi["ea"], ing["ea"])
    mp_m, mp_a = _div(ebi["pm"], ing["pm"]), _div(ebi["pa"], ing["pa"])
    R["kpi"]["ing"] = (_M(ing["em"]), f"Ingresos {ml}", f"{_pct(_div(ing['em'], ing['pm']) - 1, signo=True)} vs ppto.",
                       ing["em"] >= ing["pm"])
    R["kpi"]["ebitda"] = (_pct(me_m), f"Margen EBITDA {ml}", f"{_pp(me_m - mp_m)} vs ppto. ({_pct(mp_m)})",
                          me_m >= mp_m if not pd.isna(me_m - mp_m) else None)
    dif_e = ebi["em"] - ebi["pm"]
    R["dest"].append(("Presupuesto.",
                      f"EBITDA de {ml} {_M(ebi['em'])} ({_pct(me_m)} de margen), {_M(dif_e, True)} frente al "
                      f"presupuesto {version}; en el acumulado {_M(ebi['ea'])} frente a {_M(ebi['pa'])} "
                      f"({_M(ebi['ea'] - ebi['pa'], True)})."))

    # 1) tabla comparativa
    filas, hecho = [], set()
    for cod, r in t.iterrows():
        for limite, nombre in (("4", "EBITDA"), ("7", "Resultado antes de impuestos")):
            if cod >= limite and nombre not in hecho:
                hecho.add(nombre)
                filas.append(_fila_ppto(nombre if nombre == "EBITDA" else "Resultado antes de imp.",
                                        total(cargas.SUBTOTALES[nombre]), "total", 1))
        filas.append(_fila_ppto("Obra ejecutada / Ventas" if cod == "1.0" else cargas.CLASES.get(cod, cod), r))
    for nombre in ("EBITDA", "Resultado antes de impuestos"):
        if nombre not in hecho:
            filas.append(_fila_ppto(nombre, total(cargas.SUBTOTALES[nombre]), "total", 1))
    filas.append(_fila_ppto("Resultado del ejercicio", neto, "total", 1))

    m = MESES[mes - 1]
    msg = (f"EBITDA de {ml}: {_M(ebi['em'])}, margen {_pct(me_m)} frente a {_pct(mp_m)} presupuestado "
           f"({_pp(me_m - mp_m)})")
    s = _clonar(prs, base, "Presupuesto comparativo", msg)
    alto_f = min(285000, int(5550000 / (len(filas) + 2)))
    _tabla(s, 500000, Y0, 10500000, ["RUBRO", "MES", "ACUMULADO", "MES", "ACUMULADO", "MES", "ACUMULADO"], filas,
           [2.5, 1.45, 1.6, 1.45, 1.6, 1.3, 1.4], alto_fila=alto_f, tam=11,
           grupos=[("", 1), ("EJECUCIÓN", 2), (f"PRESUPUESTO {version}", 2), ("VARIACIÓN", 2)])
    y_t = Y0 + alto_f * (len(filas) + 2) + 150000
    if "1.0" in t.index:
        dm, da = me_m - mp_m, me_a - mp_a
        _tarjeta(s, 500000, y_t, 5150000, 1300000, "MARGEN EBITDA EJECUTADO",
                 f"{m}. {_pct(me_m)}   ACUM. {_pct(me_a)}", f"Ejecutado vs Ppto mes:  {_pp(dm)}",
                 ROJO if dm < 0 else VERDE, borde=True, tam=22)
        _tarjeta(s, 5850000, y_t, 5150000, 1300000, "MARGEN EBITDA PRESUPUESTADO",
                 f"{m}. {_pct(mp_m)}   ACUM. {_pct(mp_a)}", f"Ejec. vs Ppto acum.:  {_pp(da)}",
                 ROJO if da < 0 else VERDE, borde=True, tam=22)
    _texto(s, 11400000, Y0, 6350000, 6000000, _comentarios_ppto(ppto, ejec, mes, t), tam=16,
           titulo="COMENTARIOS MES")

    # 2) principales variaciones: tarjetas + qué rubros y sub-rubros explican la diferencia
    def tj(nombre, r, gasto=False):
        ej, pp_ = (-r["em"], -r["pm"]) if gasto else (r["em"], r["pm"])
        dif = ej - pp_
        bueno = dif <= 0 if gasto else dif >= 0
        return (nombre, _M(ej), f"Ppto. {_M(pp_)}  ·  {_M(dif, True)} ({_pct(_div(ej, pp_) - 1, signo=True)})", bueno)
    s2 = _clonar(prs, base, f"Variaciones del presupuesto {ml}",
                 f"Los ingresos van {_pct(_div(ing['em'], ing['pm']) - 1, signo=True)} y los costos y gastos "
                 f"{_pct(_div(cyg['em'], cyg['pm']) - 1, signo=True)} frente al presupuesto {version} del mes")
    _fila_tarjetas(s2, Y0, [tj("INGRESOS", ing), tj("COSTOS Y GASTOS", cyg, True), tj("EBITDA", ebi),
                            tj("RESULTADO DEL EJERCICIO", neto)], tam=22)
    clases = [(cargas.CLASES.get(c, c), r["em"] - r["pm"]) for c, r in t.iterrows() if abs(r["em"] - r["pm"]) > 1e5]
    clases = sorted(clases, key=lambda p: -p[1])
    _barras_var(s2, 500000, Y0 + 1400000, 8500000, 5300000, clases,
                titulo="Variación por rubro (millones, + a favor del resultado)")
    sub = _cuadro(ppto, ejec, mes, ["codigo", "sub_rubro"])
    sub["d"] = sub["em"] - sub["pm"]
    sub = sub[sub["d"].abs() > 1e5]
    top = pd.concat([sub["d"].nlargest(5), sub["d"].nsmallest(5)])
    top = top[~top.index.duplicated()].sort_values(ascending=False)
    _barras_var(s2, 9250000, Y0 + 1400000, 8550000, 5300000,
                [(f"{sr}", v) for (c, sr), v in top.items()],
                titulo="Sub-rubros que más mueven el resultado del mes")
    _nota(s2, 600000, Y0 + 6750000, "Verde: mejora el resultado frente al presupuesto (más ingreso o menos gasto). "
                                    "Rojo: lo empeora.", w=12000000)

    # 3) ingresos por concepto
    ti = _cuadro(ppto[ppto["codigo"] == "1.0"], ejec[ejec["codigo"] == "1.0"], mes, "sub_rubro")
    ti = ti[(ti[["em", "ea", "pm", "pa"]].abs() > 0.5).any(axis=1)].sort_values("ea", ascending=False)
    if len(ti) > 1:
        tot = ti.sum()
        s3 = _clonar(prs, base, "Ingresos presupuesto comparativo",
                     f"Ingresos de {ml} {_M(tot['em'])}: {_pct(_div(tot['em'], tot['pm']), 1)} del presupuesto · "
                     f"acumulado {_pct(_div(tot['ea'], tot['pa']), 1)}")
        f2 = [_fila_ppto(_corto(k, 46), r, signo=1) for k, r in ti.head(14).iterrows()]
        if len(ti) > 14:
            f2.append(_fila_ppto("Otros ingresos", ti.iloc[14:].sum(), signo=1))
        f2.append(_fila_ppto("Total ingresos", tot, "total", 1))
        alto2 = min(280000, int(4300000 / (len(f2) + 2)))
        _tabla(s3, 500000, Y0, 11000000,
               ["CONCEPTO", "MES", "ACUMULADO", "MES", "ACUMULADO", "MES", "ACUMULADO"], f2,
               [3.2, 1.45, 1.6, 1.45, 1.6, 1.3, 1.4], alto_fila=alto2, tam=11,
               grupos=[("", 1), ("EJECUCIÓN", 2), (f"PRESUPUESTO {version}", 2), ("VARIACIÓN", 2)])
        dv = ti["em"] - ti["pm"]
        _barras_var(s3, 11800000, Y0, 6000000, 4600000,
                    sorted([(k, v) for k, v in dv.items() if abs(v) > 5e4], key=lambda p: -p[1])[:10],
                    titulo="Variación del mes por concepto (millones)")
        sube, baja = dv[dv > 5e4].sort_values(ascending=False), dv[dv < -5e4].sort_values()
        txt = (f"Ingresos de {ml} por {_M(tot['em'])}, {_pct(_div(tot['em'], tot['pm']) - 1, signo=True)} frente "
               f"al presupuesto ({_M(tot['pm'])}). Acumulado {_M(tot['ea'])} frente a {_M(tot['pa'])} presupuestado "
               f"({_pct(_div(tot['ea'], tot['pa']) - 1, signo=True)}).")
        if len(sube):
            txt += " Por encima del presupuesto: " + _lista(list(sube.items())) + "."
        if len(baja):
            txt += " Por debajo: " + _lista(list(baja.items())) + "."
        _texto(s3, 1300000, Y0 + 4900000, 13300000, 1900000, [txt], tam=17, titulo="COMENTARIOS MES")
    return None


# ------------------------------------------------------------------ nómina
_CONC_CORTO = {"Sueldos y salarios": "Sueldos", "Horas extra y recargos": "Horas extra", "Auxilio de transporte":
               "Aux. transporte", "Bonificación de cumplimiento": "Bono cumpl.",
               "Auxilios y bonificaciones extralegales": "Extralegales",
               "Aportes (salud, pensión, riesgos, caja, parafiscales)": "Aportes",
               "Prestaciones (cesantías, intereses, primas)": "Prestaciones", "Vacaciones": "Vacaciones",
               "Otros conceptos": "Otros"}


def _diap_nomina(prs, base, leer, empresa, anio, mes, R):
    import mod_nomina
    ppto, ejec, pers = mod_nomina.preparar(leer, anio, empresa, R["version"])
    if ejec.empty or not (ejec["mes"] == mes).any():
        return f"Nómina: no hay nómina ejecutada de {MESES_L[mes - 1]} {anio} para {empresa}."
    _nota_version(R, "Nómina", ppto.attrs.get("version_usada"))
    ml = MESES_L[mes - 1]
    ejec = ejec.assign(cc_=ejec["concepto"].map(nomina.concepto_sap))
    pp, ej = ppto[ppto["mes"] == mes], ejec[ejec["mes"] == mes]
    ant = ejec[ejec["mes"] == mes - 1]
    nombres = [n for _, n in nomina.CONCEPTOS]
    con_det = (not pp.empty) and all(c in pp.columns for c in nomina.COLS_CONCEPTO) and \
        bool(pp[nomina.COLS_CONCEPTO].notna().any().any())
    if con_det:
        pc = pd.Series({n: pp[c].fillna(0).sum() for c, n in nomina.CONCEPTOS})
    else:
        pc = pd.Series({n: 0.0 for n in nombres})
        pc["Otros conceptos"] = pp["valor"].sum() if not pp.empty else 0.0
    t = pd.DataFrame({"p": pc, "e": ej.groupby("cc_")["valor"].sum().reindex(nombres),
                      "a": ant.groupby("cc_")["valor"].sum().reindex(nombres)}).fillna(0.0)
    t = t[(t[["p", "e"]].abs() > 0.5).any(axis=1)]
    t["v"] = t["e"] - t["p"]
    P, Ej, A = t["p"].sum(), t["e"].sum(), t["a"].sum()
    n_real = float(pers.loc[pers["mes"] == mes, "personas"].sum()) if not pers.empty else 0.0
    n_ppto = float(pp["personas"].sum()) if "personas" in pp.columns and not pp.empty else 0.0
    n_ant = float(pers.loc[pers["mes"] == mes - 1, "personas"].sum()) if not pers.empty else 0.0
    prom, prom_p = _div(Ej, n_real), (_div(P, n_ppto) if P else float("nan"))

    R["kpi"]["nomina"] = (_M(Ej), f"Nómina {ml}", f"{_pct(_div(Ej, P) - 1, signo=True)} vs ppto." if P else
                          (f"{_pct(_div(Ej, A) - 1, signo=True)} vs {MESES_L[mes - 2]}" if A else ""),
                          (Ej <= P) if P else None)
    if n_real:
        R["kpi"]["personas"] = (_num(n_real), "Personas en nómina",
                                f"{_ent(n_real - n_ppto)} vs ppto. ({_num(n_ppto)})" if n_ppto else "",
                                (n_real <= n_ppto) if n_ppto else None)
    txt_r = f"Nómina de {ml} {_M(Ej)}"
    if P:
        txt_r += f", {_pct(abs(_div(Ej, P) - 1))} {'por encima' if Ej > P else 'por debajo'} del presupuesto"
        ab = t[t["v"] < -5e4].sort_values("v")
        ar = t[t["v"] > 5e4].sort_values("v", ascending=False)
        if Ej <= P and len(ab):
            txt_r += f"; el mayor ahorro está en {ab.index[0].lower()} ({_M(ab['v'].iloc[0], True)})"
        elif Ej > P and len(ar):
            txt_r += f"; la mayor sobre-ejecución está en {ar.index[0].lower()} ({_M(ar['v'].iloc[0], True)})"
    if n_real:
        txt_r += f". {_num(n_real)} personas" + (f" frente a {_num(n_ppto)} presupuestadas" if n_ppto else "")
    R["dest"].append(("Nómina.", txt_r + "."))

    # 1) variación por concepto
    msg = (f"Nómina de {ml} {_M(Ej)}: " + (f"{_pct(_div(Ej, P) - 1, signo=True)} frente al presupuesto"
                                            if P else "sin presupuesto cargado")
           + (f" y {_pct(_div(Ej, A) - 1, signo=True)} frente a {MESES_L[mes - 2]}" if A else ""))
    s = _clonar(prs, base, f"Variación de nómina {ml}", msg)
    _fila_tarjetas(s, Y0, [
        ("PRESUPUESTO MES", _M(P) if P else "—", ml.capitalize(), None),
        ("NÓMINA EJECUTADA", _M(Ej), f"{_M(Ej - P, True)} vs ppto." if P else ml.capitalize(),
         (Ej <= P) if P else None),
        ("PERSONAS", _num(n_real) if n_real else "—",
         f"Ppto. {_num(n_ppto)}  ·  {_ent(n_real - n_ppto)}" if n_ppto else "", (n_real <= n_ppto) if n_ppto else None),
        ("COSTO POR PERSONA", _M(prom) if not pd.isna(prom) else "—",
         f"Ppto. {_M(prom_p)}" if not pd.isna(prom_p) else "", (prom <= prom_p) if not pd.isna(prom_p) else None),
        (f"VS {MESES_L[mes - 2].upper()}" if mes > 1 else "MES ANTERIOR", _M(Ej - A, True) if A else "—",
         f"{_pct(_div(Ej, A) - 1, signo=True)}" + (f"  ·  {_ent(n_real - n_ant)} personas" if n_ant else "")
         if A else "", (Ej <= A) if A else None),
    ], tam=20, alto=1100000)
    series = {"Presupuesto": (t["p"] / 1e6).tolist(), "Ejecutado": (t["e"] / 1e6).tolist()} if P else \
        {"Ejecutado": (t["e"] / 1e6).tolist()}
    _grafico(s, 450000, Y0 + 1250000, 10300000, 4250000, [_CONC_CORTO.get(n, n) for n in t.index], series,
             resaltar_ultimo=False, tam=12, titulo="Presupuesto vs ejecutado por concepto (millones)")
    if P:
        _barras_var(s, 10950000, Y0 + 1250000, 6900000, 4250000,
                    sorted([(_CONC_CORTO.get(n, n), v) for n, v in t["v"].items() if abs(v) > 5e4],
                           key=lambda p: -p[1]), titulo="Ejecutado − presupuesto (rojo: sobrecosto)",
                    bueno_si_sube=False)
    txt = f"En {ml} la nómina ejecutada cerró en {_M(Ej)}"
    if P:
        txt += (f", un {_pct(abs(_div(Ej, P) - 1))} {'por encima' if Ej > P else 'por debajo'} del presupuesto "
                f"({_M(P)}; variación {_M(Ej - P, True)}).")
        arriba, abajo = t[t["v"] > 5e4].sort_values("v", ascending=False), t[t["v"] < -5e4].sort_values("v")
        if len(abajo):
            txt += " Por debajo del presupuesto: " + _lista(list(abajo["v"].items())) + "."
        if len(arriba):
            txt += " Por encima: " + _lista(list(arriba["v"].items())) + "."
    else:
        txt += "."
    if not ant.empty:
        dm = (t["e"] - t["a"])
        dm = dm[dm.abs() > 5e4].sort_values(key=abs, ascending=False)
        txt += (f" Frente a {MESES_L[mes - 2]} ({_M(A)}) la nómina {'aumentó' if Ej > A else 'disminuyó'} "
                f"{_pct(abs(_div(Ej, A) - 1))}" + (", principalmente por " + _lista(list(dm.items())) if len(dm) else "")
                + ".")
    _texto(s, 900000, Y0 + 5650000, 13500000, 1700000, [txt], tam=16)

    # 2) sueldos y salarios por área y cargo (dónde bajó o subió)
    conc = "Sueldos y salarios"
    if con_det and P and conc in t.index:
        col = dict((n, c) for c, n in nomina.CONCEPTOS)[conc]
        ec = ej[ej["cc_"] == conc]
        re_mes = pers[pers["mes"] == mes]
        k = ["area", "comun"]
        d = pd.DataFrame({"p": pp.groupby(k)[col].sum(), "e": ec.groupby(k)["valor"].sum(),
                          "pp": pp.groupby(k)["personas"].sum(), "pr": re_mes.groupby(k)["personas"].sum()}).fillna(0.0)
        d["v"] = d["e"] - d["p"]
        a = d.groupby(level=0)[["p", "e", "v", "pp", "pr"]].sum().sort_values("v")
        a = a[(a[["p", "e"]].abs() > 0.5).any(axis=1)]
        if len(a) > 13:
            resto = a.iloc[6:-6].sum()
            a = pd.concat([a.iloc[:6], pd.DataFrame([resto], index=["Otras áreas"]), a.iloc[-6:]])
        f = [dict(celdas=[_corto(n, 34), _num(r.p), _num(r.e), _num(r.v), _num(r.pp), _num(r.pr), _num(r.pr - r.pp)],
                  rojo={3} if r.v > 0.5 else (), verde={3} if r.v < -0.5 else ()) for n, r in a.iterrows()]
        tt = d[["p", "e", "v", "pp", "pr"]].sum()
        f.append(dict(celdas=["Total", _num(tt.p), _num(tt.e), _num(tt.v), _num(tt.pp), _num(tt.pr),
                              _num(tt.pr - tt.pp)], tipo="total"))
        s2 = _clonar(prs, base, "Detalle variación sueldos y salarios",
                     f"Sueldos de {ml} {_M(tt.e)} frente a {_M(tt.p)} presupuestados ({_M(tt.v, True)}) · "
                     f"{_num(tt.pr)} personas frente a {_num(tt.pp)}")
        alto = min(290000, int(4200000 / (len(f) + 2)))
        _tabla(s2, 500000, Y0, 10900000,
               ["ÁREA", "PRESUPUESTO", "EJECUTADO", "VARIACIÓN", "PPTO.", "REAL", "DIF."], f,
               [3.0, 1.6, 1.6, 1.5, .8, .8, .8], alto_fila=alto, tam=11,
               grupos=[("", 1), (f"SUELDOS Y SALARIOS {MESES[mes - 1]}", 3), ("PERSONAS", 3)])
        top = pd.concat([d["v"].nsmallest(5), d["v"].nlargest(3)])
        top = top[(top.abs() > 5e4) & ~top.index.duplicated()]
        _barras_var(s2, 500000, Y0 + alto * (len(f) + 2) + 150000, 10900000,
                    max(2400000, 7250000 - alto * (len(f) + 2) - 150000),
                    sorted([(f"{_corto(c, 26)} · {_corto(ar, 18)}", v) for (ar, c), v in top.items()],
                           key=lambda p: -p[1]), titulo="Cargos que más se alejan del presupuesto (rojo: sobrecosto)",
                    tam=11, bueno_si_sube=False)

        def cargos(sub):
            return "; ".join(f"{_corto(c, 34).lower()} en {_corto(ar, 26)} ({_M(v, True)})"
                             for (ar, c), v in sub["v"].items())
        baj, sub_ = d[d["v"] < -5e4].sort_values("v").head(4), d[d["v"] > 5e4].sort_values("v", ascending=False).head(4)
        par = [("", f"Sueldos y salarios de {ml}: {_M(tt.e)} frente a {_M(tt.p)} presupuestado "
                    f"({_M(tt.v, True)}), con {_num(tt.pr)} personas en nómina frente a {_num(tt.pp)} presupuestadas.")]
        if len(baj):
            par.append(("Disminuyen.", cargos(baj) + "."))
        if len(sub_):
            par.append(("Aumentan.", cargos(sub_) + "."))
        _texto(s2, 11750000, Y0, 6000000, 6500000, par, tam=15, titulo="COMENTARIOS MES")

    # 3) evolución mensual: nómina total y horas extra
    meses = [m_ for m_ in range(1, mes + 1) if (ejec["mes"] == m_).any()]
    if len(meses) > 1:
        tot_m = ejec.groupby("mes")["valor"].sum()
        he = ejec[ejec["cc_"] == "Horas extra y recargos"].groupby("mes")["valor"].sum()
        hay_he = bool(he.abs().sum() > 0.5)
        h1, h0 = he.get(mes, 0.0), he.get(mes - 1, 0.0)
        prom_m = tot_m.reindex(meses).mean()
        s3 = _clonar(prs, base, "Detalle nómina y horas extra",
                     f"Nómina de {ml} {_pct(_div(Ej, prom_m) - 1, signo=True)} frente al promedio del año"
                     + (f" · horas extra {_pct(_div(h1, Ej))} de la nómina" if hay_he else ""))
        cats = [MESES[m_ - 1] for m_ in meses]
        _grafico(s3, 450000, Y0, 8650000 if hay_he else 13000000, 4300000, cats,
                 {"Nómina": [tot_m.get(m_, 0) / 1e6 for m_ in meses]}, titulo="Nómina ejecutada (millones)", tam=12)
        if hay_he:
            _grafico(s3, 9250000, Y0, 8550000, 4300000, cats,
                     {"Horas extra": [he.get(m_, 0) / 1e6 for m_ in meses]},
                     titulo="Horas extra y recargos (millones)", formato="#,##0.0", tam=12)
        tarj = [("ACUMULADO AÑO", _M(tot_m.reindex(meses).sum()), f"{MESES[meses[0] - 1]} - {MESES[mes - 1]}", None),
                ("PROMEDIO MES", _M(prom_m), f"{len(meses)} meses", None)]
        if hay_he:
            tarj += [("HORAS EXTRA MES", _M(h1), f"{_pct(_div(h1, h0) - 1, signo=True)} vs {MESES_L[mes - 2]}"
                      if h0 else "", (h1 <= h0) if h0 else None),
                     ("% HORAS EXTRA", _pct(_div(h1, Ej)), f"Promedio año {_pct(_div(he.sum(), tot_m.sum()))}",
                      _div(h1, Ej) <= _div(he.sum(), tot_m.sum()))]
        _fila_tarjetas(s3, Y0 + 4450000, tarj, alto=1050000, tam=20)
        txt = f"En {ml} la nómina ejecutada fue de {_M(Ej)}"
        if A:
            txt += f", {_pct(_div(Ej, A) - 1, signo=True)} frente a {MESES_L[mes - 2]} ({_M(A)})"
        if hay_he:
            txt += f". Las horas extra y recargos cerraron en {_M(h1)}"
        if hay_he and h0:
            txt += (f", {'un aumento' if h1 > h0 else 'una reducción'} del {_pct(abs(_div(h1, h0) - 1))} frente a "
                    f"{MESES_L[mes - 2]} ({_M(h0)})")
        pico = he.reindex(meses).fillna(0)
        if hay_he and pico.idxmax() != mes:
            txt += f"; el mes más alto del año es {MESES_L[int(pico.idxmax()) - 1]} ({_M(pico.max())})"
        _texto(s3, 1300000, Y0 + 5700000, 13300000, 1400000, [txt + "."], tam=17)
    return None


# ------------------------------------------------------------------ flujo de caja
def _diap_flujo(prs, base, leer, empresa, anio, mes, R):
    lin = leer("SELECT * FROM flujo_lineas WHERE anio=? AND empresa=? ORDER BY orden, mes", (anio, empresa))
    if lin.empty or not ((lin["mes"] == mes) & (lin["valor"] != 0)).any():
        return f"Flujo de caja: no hay flujo de {MESES_L[mes - 1]} {anio} para {empresa}."
    ml = MESES_L[mes - 1]
    info = lin.drop_duplicates("orden").set_index("orden")[["seccion", "padre", "concepto", "tipo"]]
    info["padre"] = info["padre"].fillna("")
    ancho = lin.pivot_table(index="orden", columns="mes", values="valor", aggfunc="sum") \
        .reindex(columns=range(1, 13)).fillna(0.0)
    meses = list(range(1, mes + 1))
    por_nombre = {info.at[o, "concepto"]: o for o in info.index}

    def serie(c):
        o = por_nombre.get(c)
        return ancho.loc[o] if o is not None else pd.Series(0.0, index=range(1, 13))

    def total(o):
        k, sr = str(info.at[o, "concepto"]).lower(), ancho.loc[o]
        if k.startswith("saldo inicial"):
            return sr[meses[0]]
        if k.startswith("saldo final") or k.startswith("flujo de caja"):
            return sr[meses[-1]]
        return sr[meses].sum()

    si, se, sf, s0 = serie("Total Ingresos"), serie("Total Egresos"), serie("Flujo de caja operación"), \
        serie("Saldo inicial")
    egr3 = [-se[k] for k in range(max(1, mes - 2), mes + 1)]
    prom3 = sum(egr3) / len(egr3) if egr3 else 0.0
    dias = _div(sf[mes], prom3 / 30) if prom3 else float("nan")
    neto = si[mes] + se[mes]
    R["kpi"]["caja"] = (_M(sf[mes]), f"Caja operativa fin de {ml}",
                        f"{_M(sf[mes] - s0[mes], True)} en el mes" + (f"  ·  {_dias(dias)}" if not pd.isna(dias) else ""),
                        sf[mes] >= s0[mes])
    R["dest"].append(("Caja.", f"La caja operativa pasó de {_M(s0[mes])} a {_M(sf[mes])} en {ml} (flujo neto "
                               f"{_M(neto, True)}): ingresaron {_M(si[mes])} y se pagaron {_M(-se[mes])}"
                      + (f"; alcanza para {_dias(dias)} al ritmo de pagos de los últimos 3 meses" if not pd.isna(dias)
                         else "") + "."))

    # 1) tabla
    filas, sec = [], None
    top = info[(info["padre"] == "") & info["tipo"].isin(["linea", "total"])]
    for o, r in top.iterrows():
        if not (ancho.loc[o, meses].abs() > 0.5).any():
            continue
        if r["tipo"] == "linea" and r["seccion"] != sec:
            sec = r["seccion"]
            filas.append(dict(celdas=[sec] + [""] * (len(meses) + 1), tipo="seccion"))
        filas.append(dict(celdas=[("   " if r["tipo"] == "linea" else "") + _corto(r["concepto"], 36)]
                          + [_num(ancho.at[o, m_], 1e3) for m_ in meses] + [_num(total(o), 1e3)],
                          tipo="total" if r["tipo"] == "total" else ""))
    s = _clonar(prs, base, "Flujo de caja ejecutado",
                f"Saldo final de {ml} {_M(sf[mes])} · flujo neto del mes {_M(neto, True)}"
                + (f" · {_dias(dias)} de caja" if not pd.isna(dias) else ""))
    n = len(filas) + 1
    alto = min(290000, int(6750000 / n))
    _tabla(s, 900000, Y0, 16500000, [""] + [MESES[m_ - 1] for m_ in meses] + ["Total"], filas,
           [3.2] + [1.0] * len(meses) + [1.15], alto_fila=alto, tam=11 if n <= 28 else (10 if n <= 34 else 9))
    _nota(s, 900000, Y0 + alto * n + 60000, "Cifras en miles de cops")

    # 2) variación del mes
    s2 = _clonar(prs, base, f"Variación flujo de caja {ml}",
                 f"Ingresos {_pct(_div(si[mes], si[mes - 1]) - 1, signo=True)} y egresos "
                 f"{_pct(_div(se[mes], se[mes - 1]) - 1, signo=True)} frente a {MESES_L[mes - 2]}" if mes > 1 else
                 f"Flujo neto de {ml}: {_M(neto, True)}")
    _grafico(s2, 450000, Y0, 12700000, 4400000, [MESES[m_ - 1] for m_ in meses],
             {"Ingresos": [si[m_] / 1e6 for m_ in meses], "Egresos": [-se[m_] / 1e6 for m_ in meses]},
             resaltar_ultimo=False, tam=12, titulo="Ingresos y egresos por mes (millones)")
    y = Y0
    for et, val, sub_, col in [("SALDO INICIAL", _M(s0[mes]), f"Inicio de {ml}", AZUL_TXT),
                               ("INGRESOS DEL MES", _M(si[mes]),
                                f"{_pct(_div(si[mes], si[mes - 1]) - 1, signo=True)} vs {MESES_L[mes - 2]}"
                                if mes > 1 else "", VERDE if mes > 1 and si[mes] >= si[mes - 1] else ROJO),
                               ("EGRESOS DEL MES", _M(-se[mes]),
                                f"{_pct(_div(se[mes], se[mes - 1]) - 1, signo=True)} vs {MESES_L[mes - 2]}"
                                if mes > 1 else "", ROJO if mes > 1 and -se[mes] > -se[mes - 1] else VERDE),
                               ("SALDO FINAL", _M(sf[mes]), f"{_M(sf[mes] - s0[mes], True)} en el mes",
                                VERDE if sf[mes] >= s0[mes] else ROJO),
                               ("DÍAS DE CAJA", _dias(dias), "Ritmo de pagos últimos 3 meses",
                                ROJO if not pd.isna(dias) and dias < 30 else VERDE)]:
        _tarjeta(s2, 13450000, y, 4350000, 1000000, et, val, sub_, col, tam=20)
        y += 1090000
    txt = (f"En {ml} ingresaron {_M(si[mes])} y se pagaron {_M(-se[mes])}, para un flujo neto de "
           f"{_M(neto)}. La caja operativa pasó de {_M(s0[mes])} a {_M(sf[mes])}.")
    lin_top = top[top["tipo"] == "linea"]
    if mes > 1:
        d = pd.Series({r["concepto"]: ancho.at[o, mes] - ancho.at[o, mes - 1] for o, r in lin_top.iterrows()})
        d = d[d.abs() > 5e4]
        mas, menos = d[d > 0].sort_values(ascending=False), d[d < 0].sort_values()
        txt += f" Frente a {MESES_L[mes - 2]}"
        if len(mas):
            txt += ", mejoran la caja: " + _lista(list(mas.items()))
        if len(menos):
            txt += ("; " if len(mas) else ", ") + "la reducen: " + _lista(list(menos.items()))
        txt += "."
    _texto(s2, 900000, Y0 + 4600000, 12300000, 2300000, [txt], tam=17)

    # 3) en qué se fue la caja: egresos por concepto y mayores pagos del mes
    egr = sorted([(r["concepto"], -ancho.at[o, mes]) for o, r in lin_top.iterrows()
                  if r["seccion"] == "Egresos" and -ancho.at[o, mes] > 0.5], key=lambda p: -p[1])
    if egr:
        tot_e = sum(v for _, v in egr)
        s3 = _clonar(prs, base, f"Egresos y principales pagos {ml}",
                     f"{_corto(egr[0][0], 40)} concentra el {_pct(_div(egr[0][1], tot_e), 0)} de los egresos del mes")
        _grafico(s3, 450000, Y0, 9300000, 5200000, [_corto(n, 22) for n, _ in egr[:9]],
                 {"Egresos": [v / 1e6 for _, v in egr[:9]]}, resaltar_ultimo=False, tam=11,
                 titulo=f"Egresos de {ml} por concepto (millones)")
        egr_n = {" ".join(str(c).lower().split()) for c, _ in egr if "traslado" not in str(c).lower()}
        mov = leer("SELECT concepto, tercero, valor FROM flujo_mov WHERE anio=? AND empresa=? AND mes=?",
                   (anio, empresa, mes))
        if not mov.empty:
            mov = mov[(mov["valor"] < 0) & mov["concepto"].fillna("").map(
                lambda c: " ".join(str(c).lower().split())).isin(egr_n)]
            mov = mov[mov["tercero"].fillna("").str.strip().ne("") & mov["tercero"].ne("(Sin tercero)")]
        if not mov.empty:
            pt = mov.groupby("tercero").agg(v=("valor", "sum"), c=("concepto", "first")).assign(v=lambda x: -x["v"])
            pt = pt.nlargest(10, "v")
            tot_m = -mov["valor"].sum()
            f = [dict(celdas=[_corto(tr, 34), _corto(r.c, 22), _num(r.v / 1e6, dec=1), _pct(_div(r.v, tot_m))])
                 for tr, r in pt.iterrows()]
            f.append(dict(celdas=["Total 10 mayores", "", _num(pt["v"].sum() / 1e6, dec=1),
                                  _pct(_div(pt["v"].sum(), tot_m))], tipo="total"))
            _tabla(s3, 9950000, Y0, 7900000, ["TERCERO", "CONCEPTO", "M$", "%"], f, [3.2, 2.1, 1.0, .8],
                   alto_fila=330000, tam=11, grupos=[(f"MAYORES PAGOS DE {MESES[mes - 1]}", 4)])
        txt = (f"Los egresos de {ml} suman {_M(tot_e)}: " + _lista([(n, v) for n, v in egr], 4).replace("+", "")
               + ".")
        _texto(s3, 900000, Y0 + 5450000, 13000000, 1500000, [txt], tam=17)
    return None


# ------------------------------------------------------------------ combustible
def _diap_combustible(prs, base, leer, empresa, anio, mes, R):
    b = leer("SELECT mes, fuente, placa, valor, galones FROM combustible WHERE anio=? AND empresa=?", (anio, empresa))
    if b.empty or not (b["mes"] == mes).any():
        return f"Combustible: no hay tanqueos de {MESES_L[mes - 1]} {anio} para {empresa}."
    ml = MESES_L[mes - 1]
    b = b[b["mes"] <= mes]
    pm = b.groupby("mes")[["valor", "galones"]].sum()
    meses = [m_ for m_ in range(1, mes + 1) if m_ in pm.index]
    cats = [MESES[m_ - 1] for m_ in meses]
    c1, g1 = pm.at[mes, "valor"], pm.at[mes, "galones"]
    c0, g0 = (pm.at[mes - 1, "valor"], pm.at[mes - 1, "galones"]) if mes - 1 in pm.index else (None, None)
    p1, p0 = _div(c1, g1), (_div(c0, g0) if c0 else float("nan"))
    prom = pm["valor"].mean()
    R["kpi"]["comb"] = (_M(c1), f"Combustible {ml}",
                        f"{_pct(_div(c1, c0) - 1, signo=True)} vs {MESES_L[mes - 2]}" if c0 else "",
                        (c1 <= c0) if c0 else None)
    if c0:
        R["dest"].append(("Combustible.", f"{_M(c1)} en {ml} ({_pct(_div(c1, c0) - 1, signo=True)} frente a "
                                          f"{MESES_L[mes - 2]}): galones {_pct(_div(g1, g0) - 1, signo=True)} y precio "
                                          f"por galón {_pct(_div(p1, p0) - 1, signo=True)}."))
    s = _clonar(prs, base, "Consumo de combustible",
                f"{ml.capitalize()}: {_M(c1)}, {_pct(_div(c1, prom) - 1, signo=True)} frente al promedio del año"
                + (f" · precio por galón {_pct(_div(p1, p0) - 1, signo=True)}" if not pd.isna(p0) else ""))
    _grafico(s, 450000, Y0, 8650000, 4200000, cats, {"Costo": [pm.at[m_, "valor"] / 1e6 for m_ in meses]},
             titulo="Costo de combustible (millones de cops)", formato="#,##0.0", tam=12)
    # valor por proveedor: tabla del mes y, debajo, cada proveedor mes a mes
    b["fuente"] = b["fuente"].fillna("").replace("", "Sin proveedor")
    pv = b[b["mes"] == mes].groupby("fuente")[["valor", "galones"]].sum().sort_values("valor", ascending=False)
    pv_a = b[b["mes"] == mes - 1].groupby("fuente")["valor"].sum()
    pv_ac = b.groupby("fuente")["valor"].sum()
    fp = []
    for prov, r in pv.iterrows():
        d = r.valor - pv_a.get(prov, 0.0)
        fp.append(dict(celdas=[_corto(prov, 26), _num(r.galones), _num(r.valor / 1e6, dec=1), _pct(_div(r.valor, c1)),
                               _num(d / 1e6, dec=1) if mes > 1 else "", "$" + _num(_div(r.valor, r.galones)),
                               _num(pv_ac.get(prov, 0.0) / 1e6, dec=1)],
                       rojo={4} if d > 1e3 else (), verde={4} if d < -1e3 else ()))
    fp.append(dict(celdas=["Total", _num(g1), _num(c1 / 1e6, dec=1), "100%", _num((c1 - (c0 or 0)) / 1e6, dec=1)
                           if mes > 1 else "", "$" + _num(p1), _num(pm["valor"].sum() / 1e6, dec=1)], tipo="total"))
    alto_p = min(300000, int(2300000 / (len(fp) + 2)))
    _tabla(s, 9250000, Y0, 8550000, ["PROVEEDOR", "GALONES", "VALOR M$", "%", f"VS {MESES[mes - 2]}" if mes > 1
                                     else "VS MES ANT.", "PRECIO GAL.", "ACUM. M$"], fp,
           [2.4, 1.1, 1.1, .8, 1.0, 1.15, 1.1], alto_fila=alto_p, tam=11,
           grupos=[(f"VALOR POR PROVEEDOR {MESES[mes - 1]}", 7)])
    y_g = Y0 + alto_p * (len(fp) + 2) + 120000
    h_g = Y0 + 4200000 - y_g
    if len(pv) > 1:
        provs = list(pv.index[:5])
        pm_p = b[b["fuente"].isin(provs)].pivot_table(index="mes", columns="fuente", values="valor", aggfunc="sum")
        _grafico(s, 9250000, y_g, 8550000, h_g, cats,
                 {p_: [pm_p.at[m_, p_] / 1e6 if m_ in pm_p.index and not pd.isna(pm_p.at[m_, p_]) else 0.0
                       for m_ in meses] for p_ in provs},
                 titulo="Valor por proveedor cada mes (millones)", formato="#,##0.0", resaltar_ultimo=False, tam=10,
                 colores=(AZUL, NARANJA, NAVY, GRIS_OSC, VERDE_BAR))
    else:
        _grafico(s, 9250000, y_g, 8550000, h_g, cats, {"Galones": [pm.at[m_, "galones"] for m_ in meses]},
                 titulo="Galones consumidos", tam=10)
    rango = f"Acum. {MESES[meses[0] - 1].capitalize()} - {MESES[mes - 1].capitalize()}"
    _fila_tarjetas(s, Y0 + 4350000, [
        ("COSTO DEL MES", _M(c1), f"{_pct(_div(c1, c0) - 1, signo=True)} vs {MESES_L[mes - 2]}" if c0 else "",
         (c1 <= c0) if c0 else None),
        ("GALONES DEL MES", _num(g1), f"{_pct(_div(g1, g0) - 1, signo=True)} vs {MESES_L[mes - 2]}" if g0 else "",
         (g1 <= g0) if g0 else None),
        ("PRECIO GALÓN", "$" + _num(p1), f"{_pct(_div(p1, p0) - 1, signo=True)} vs {MESES_L[mes - 2]}"
         if not pd.isna(p0) else "", (p1 <= p0) if not pd.isna(p0) else None),
        ("TOTAL CONSUMO", _M(pm["valor"].sum()), rango, None),
        ("PROMEDIO MES", _M(prom), f"{len(meses)} meses", None),
    ], alto=1050000, tam=20)
    txt = f"El consumo de combustible en {ml} cerró en {_M(c1)} ({_num(g1)} galones)"
    if c0:
        txt += (f", {'un incremento' if c1 > c0 else 'una reducción'} del {_pct(abs(_div(c1, c0) - 1))} frente a "
                f"{MESES_L[mes - 2]} ({_M(c0)}): los galones variaron {_pct(_div(g1, g0) - 1, signo=True)} y el precio "
                f"promedio por galón {_pct(_div(p1, p0) - 1, signo=True)} (${_num(p1)} frente a ${_num(p0)})")
    txt += (f". El acumulado {MESES_L[meses[0] - 1]} - {ml} suma {_M(pm['valor'].sum())}, con un "
            f"promedio mensual de {_M(prom)}.")
    if len(pv) > 1:
        txt += (" Por proveedor: " + ", ".join(f"{_corto(p_, 26)} {_M(r.valor)} ({_pct(_div(r.valor, c1), 0)})"
                                               for p_, r in pv.head(4).iterrows()) + ".")
        dpv = (pv["valor"] - pv_a.reindex(pv.index).fillna(0.0))
        dpv = dpv[dpv.abs() > 1e5]
        if mes > 1 and len(dpv):
            k = dpv.abs().idxmax()
            txt += f" El mayor cambio frente a {MESES_L[mes - 2]} es {_corto(k, 26)} ({_M(dpv[k], True)})."
    _texto(s, 1300000, Y0 + 5600000, 13300000, 1500000, [txt], tam=16)

    # 2) por vehículo
    pl = b[b["mes"] == mes].groupby("placa")[["valor", "galones"]].sum().sort_values("valor", ascending=False)
    if len(pl) >= 2:
        pa = b[b["mes"] == mes - 1].groupby("placa")["valor"].sum()
        top10 = pl.head(10)
        s2 = _clonar(prs, base, f"Consumo por vehículo {ml}",
                     f"{len(pl)} vehículos · los 5 de mayor consumo suman el "
                     f"{_pct(_div(pl['valor'].head(5).sum(), c1), 0)} del costo del mes")
        _grafico(s2, 450000, Y0, 10300000, 4700000, list(top10.index), {"Costo": (top10["valor"] / 1e6).tolist()},
                 titulo="10 vehículos de mayor consumo del mes (millones)", formato="#,##0.0", resaltar_ultimo=False,
                 tam=11)
        var = (pl["valor"] - pa.reindex(pl.index).fillna(0)) if not pa.empty else pd.Series(dtype=float)
        var = var[var.abs() > 1e4]
        if len(var):
            top = pd.concat([var.nlargest(4), var.nsmallest(4)])
            top = top[~top.index.duplicated()]
            _barras_var(s2, 10950000, Y0, 6900000, 4700000, sorted(top.items(), key=lambda q: -q[1]),
                        titulo=f"Variación vs {MESES_L[mes - 2]} (rojo: subió)", tam=11, bueno_si_sube=False)
        _fila_tarjetas(s2, Y0 + 4850000, [
            ("VEHÍCULOS", _num(len(pl)), f"{_num(pa.size)} en {MESES_L[mes - 2]}" if not pa.empty else "", None),
            ("COSTO POR VEHÍCULO", _M(_div(c1, len(pl))), "Promedio del mes", None),
            ("GALONES POR VEHÍCULO", _num(_div(g1, len(pl))), "Promedio del mes", None),
        ], x0=500000, x1=13300000, alto=1000000, tam=20)
        txt = "Mayor consumo del mes: " + ", ".join(f"{p} ({_M(r.valor)})" for p, r in pl.head(3).iterrows()) + "."
        if len(var):
            sube = var[var > 0].nlargest(3)
            if len(sube):
                txt += " Más aumentaron frente a " + MESES_L[mes - 2] + ": " + \
                    ", ".join(f"{p} ({_M(v, True)})" for p, v in sube.items()) + "."
        _texto(s2, 900000, Y0 + 6000000, 13000000, 1100000, [txt], tam=16)
    return None


# ------------------------------------------------------------------ estados financieros
def _datos_eeff(leer, empresa, anio, mes):
    lin = leer("SELECT estado, idx, grupo, subgrupo, linea, anio, mes, valor FROM eeff WHERE empresa=?", (empresa,))
    if lin.empty or not ((lin["anio"] == anio) & (lin["mes"] == mes)).any():
        return None
    r = eeff.resumen(lin)
    ref = (anio - 1, mes) if (anio - 1, mes) in r.index else ((anio - 1, 12) if (anio - 1, 12) in r.index else None)
    return lin, r, ref


def _diap_balance(prs, base, leer, empresa, anio, mes, R):
    d = _datos_eeff(leer, empresa, anio, mes)
    if d is None:
        return f"Balance: no hay estados financieros de {MESES_L[mes - 1]} {anio} para {empresa}."
    lin, r, ref = d
    esf = lin[lin["estado"] == "ESF"]
    f = r.loc[(anio, mes)]
    fb = r.loc[ref] if ref else None
    et_a = f"{MESES[mes - 1]} {anio}"
    et_b = f"{MESES[ref[1] - 1]} {ref[0]}" if ref else ""
    vs = f" vs {et_b.lower()}" if ref else ""

    def corte(k):
        e = esf[(esf["anio"] == k[0]) & (esf["mes"] == k[1])]
        return e.groupby(["grupo", "subgrupo"], sort=False)["valor"].sum()
    act = corte((anio, mes))
    ant = corte(ref) if ref else pd.Series(dtype=float)
    t = pd.DataFrame({"a": act, "b": ant}).fillna(0.0)
    orden = esf.drop_duplicates(["grupo", "subgrupo"]).set_index(["grupo", "subgrupo"])["idx"]
    t = t.reindex(sorted(t.index, key=lambda k: orden.get(k, 999)))
    umbral = 0.01 * max(abs(f["ACTIVO"]), 1.0)                    # partidas de menos del 1% del activo: "Otros"

    def fila(nombre, a, b, tipo=""):
        return dict(celdas=[nombre, _num(a / 1e6, dec=1), _num(b / 1e6, dec=1) if ref else "",
                            _num((a - b) / 1e6, dec=1) if ref else "",
                            _pct(_div(a, b) - 1, signo=True) if ref and abs(b) > 1 else ""], tipo=tipo)
    filas, tot = [], {}
    for g in eeff.GRUPOS_ESF:
        sub = t[t.index.get_level_values(0) == g]
        sub = sub[(sub[["a", "b"]].abs() > 0.5).any(axis=1)]
        grandes = sub[(sub[["a", "b"]].abs() >= umbral).any(axis=1)]
        chicos = sub.drop(grandes.index)
        for (_, sg), v in grandes.iterrows():
            filas.append(fila("   " + _corto(sg.capitalize(), 50), v["a"], v["b"]))
        if len(chicos):
            filas.append(fila("   Otros", chicos["a"].sum(), chicos["b"].sum()))
        tot[g] = sub[["a", "b"]].sum()
        filas.append(fila(g, tot[g]["a"], tot[g]["b"], "sub"))
        if g == "ACTIVO NO CORRIENTE":
            ta = tot["ACTIVO CORRIENTE"] + tot[g]
            filas.append(fila("TOTAL ACTIVO", ta["a"], ta["b"], "total"))
        if g == "PASIVO NO CORRIENTE":
            tp = tot["PASIVO CORRIENTE"] + tot[g]
            filas.append(fila("TOTAL PASIVO", tp["a"], tp["b"], "total"))
    tpp = tp + tot.get("PATRIMONIO", 0)
    filas.append(fila("PASIVO + PATRIMONIO", tpp["a"], tpp["b"], "total"))
    R["kpi"]["ktnc"] = (_M(f["ktnc"]), "Capital de trabajo",
                        f"Razón corriente {_x(f['razon_corriente'])}" + (f" ({_x(fb['razon_corriente'])}{vs})"
                                                                       if fb is not None else ""),
                        (f["ktnc"] >= fb["ktnc"]) if fb is not None else None)
    if fb is not None:
        R["dest"].append(("Balance.", f"Activo {_M(f['ACTIVO'])} ({_pct(_div(f['ACTIVO'], fb['ACTIVO']) - 1, signo=True)}"
                                      f"{vs}), endeudamiento {_pct(f['endeudamiento'])} "
                                      f"({_pp(f['endeudamiento'] - fb['endeudamiento'])}) y capital de trabajo "
                                      f"{_M(f['ktnc'])}."))

    s = _clonar(prs, base, f"Balance general {et_a}" + (f" - {ref[0]}" if ref else ""),
                f"Activo {_M(f['ACTIVO'])} ({_pct(_div(f['ACTIVO'], fb['ACTIVO']) - 1, signo=True)}{vs}) · "
                f"endeudamiento {_pct(f['endeudamiento'])}" if fb is not None else
                f"Activo {_M(f['ACTIVO'])} · endeudamiento {_pct(f['endeudamiento'])}")
    n = len(filas) + 1
    alto = min(290000, int(6900000 / n))
    _tabla(s, 500000, Y0, 10900000, ["MILLONES", et_a, et_b, "VAR", "%"], filas, [4.3, 1.6, 1.6, 1.5, .8],
           alto_fila=alto, tam=11 if n <= 30 else 10)

    def dif(c, modo):
        if fb is None or pd.isna(fb[c]) or pd.isna(f[c]):
            return ""
        if modo == "pp":
            return _pp(f[c] - fb[c]) + vs
        if modo == "M":
            return _M(f[c] - fb[c], True) + vs
        return f"{f[c] - fb[c]:+.2f}".replace(".", ",") + vs
    x, w = 11800000, 1950000
    for i, (et, val, c, modo, mejor_si_sube) in enumerate([
            ("CAPITAL DE TRABAJO", _M(f["ktnc"]), "ktnc", "M", True),
            ("RAZÓN CORRIENTE", _x(f["razon_corriente"]), "razon_corriente", "x", True),
            ("ENDEUDAMIENTO", _pct(f["endeudamiento"]), "endeudamiento", "pp", False)]):
        bueno = None if fb is None or pd.isna(fb[c]) else ((f[c] >= fb[c]) == mejor_si_sube)
        _tarjeta(s, x + i * (w + 100000), Y0, w, 1250000, et, val, dif(c, modo),
                 AZUL_TXT if bueno is None else (VERDE if bueno else ROJO), tam=18)
    par = []
    if ref:
        t["d"] = t["a"] - t["b"]
        for nombre, grupos in (("Activo.", ["ACTIVO CORRIENTE", "ACTIVO NO CORRIENTE"]),
                               ("Pasivo.", ["PASIVO CORRIENTE", "PASIVO NO CORRIENTE"]),
                               ("Patrimonio.", ["PATRIMONIO"])):
            sub = t[t.index.get_level_values(0).isin(grupos)]
            dt = sub["d"].sum()
            dd = sub["d"][sub["d"].abs() > 1e6].sort_values(key=abs, ascending=False)
            par.append((nombre, f"{_M(sub['a'].sum())}, {_M(dt, True)} "
                                f"({_pct(_div(sub['a'].sum(), sub['b'].sum()) - 1, signo=True)}){vs}"
                                + (": " + _lista([(k[1], v) for k, v in dd.items()], 3) if len(dd) else "") + "."))
    else:
        par.append(("", f"Activo {_M(ta['a'])}, pasivo {_M(tp['a'])} y patrimonio {_M(tot['PATRIMONIO']['a'])} "
                        f"al cierre de {MESES_L[mes - 1]} {anio}."))
    _texto(s, 11800000, Y0 + 1450000, 6050000, 5200000, par, tam=15, titulo="COMENTARIOS MES")

    # 2) indicadores financieros, como el tablero de la presentación
    s2 = _clonar(prs, base, f"Indicadores financieros {et_a}",
                 f"Margen EBITDA {_pct(f['m_ebitda'])} · ROE {_pct(f['roe'])} · razón corriente "
                 f"{_x(f['razon_corriente'])} · endeudamiento {_pct(f['endeudamiento'])}")
    ind = [("ktnc", "Capital de trabajo", _M, "M", True), ("razon_corriente", "Razón corriente", _x, "x", True),
           ("m_ebitda", "% EBITDA", _pct, "pp", True), ("roa", "ROA", _pct, "pp", True),
           ("roe", "ROE", _pct, "pp", True), ("endeudamiento", "Endeudamiento", _pct, "pp", False),
           ("deuda_ebitda", "Deuda neta / EBITDA", _x, "x", False), ("dias_cartera", "Días de cartera", _dias, "d",
                                                                    False),
           ("dias_proveedores", "Días de proveedores", _dias, "d", True),
           ("ciclo_efectivo", "Ciclo de efectivo", _dias, "d", False)]
    w, sep, x0 = 3200000, 170000, 500000
    for i, (c, nombre, fmt, modo, mejor_si_sube) in enumerate(ind):
        sub_, bueno = "", None
        if fb is not None and not pd.isna(fb[c]) and not pd.isna(f[c]):
            dv = f[c] - fb[c]
            sub_ = {"pp": _pp(dv), "M": _M(dv, True), "x": f"{dv:+.2f}".replace(".", ",") + "x",
                    "d": f"{dv:+,.0f}".replace(",", ".") + " días"}[modo] + f" vs {ref[0]}"
            bueno = None if abs(dv) < (0.0005 if modo == "pp" else 0.005) else (dv >= 0) == mejor_si_sube
        _kpi(s2, x0 + (i % 5) * (w + sep), Y0 + (i // 5) * 1500000, w, 1350000, fmt(f[c]), nombre, sub_, bueno,
             tam=28)
    m_act = [m_ for m_ in range(1, mes + 1) if (anio, m_) in r.index]
    serie_a = [r.at[(anio, m_), "ebitda_m"] / r.at[(anio, m_), "ing_m"] if r.at[(anio, m_), "ing_m"] else None
               for m_ in m_act]
    series = {}
    if all((anio - 1, m_) in r.index for m_ in m_act):
        series[str(anio - 1)] = [_div(r.at[(anio - 1, m_), "ebitda_m"], r.at[(anio - 1, m_), "ing_m"]) for m_ in m_act]
    series[str(anio)] = serie_a
    _grafico(s2, 500000, Y0 + 3150000, 11500000, 3900000, [MESES[m_ - 1] for m_ in m_act], series,
             titulo="Margen EBITDA de cada mes", formato="0.0%", resaltar_ultimo=False, tam=11,
             colores=(GRIS_OSC, AZUL) if len(series) > 1 else (AZUL,))
    par = [("Liquidez.", f"Capital de trabajo {_M(f['ktnc'])} y razón corriente {_x(f['razon_corriente'])}"
                         + (f" ({_x(fb['razon_corriente'])} en {ref[0]})" if fb is not None else "") + "."),
           ("Rentabilidad.", f"Margen EBITDA acumulado {_pct(f['m_ebitda'])}, ROA {_pct(f['roa'])} y ROE "
                             f"{_pct(f['roe'])} anualizados."),
           ("Ciclo.", f"Cartera a {_dias(f['dias_cartera'])} y proveedores a {_dias(f['dias_proveedores'])}: "
                      f"ciclo de efectivo de {_dias(f['ciclo_efectivo'])}.")]
    _texto(s2, 12300000, Y0 + 3250000, 5500000, 3800000, par, tam=15, titulo="LECTURA")
    return None


_PYG = [("INGRESOS", "ing", "", 1), ("COSTOS", "cos", "", -1), ("RESULTADO BRUTO", "bruto", "sub", 1),
        ("GASTOS DE ADMINISTRACIÓN", "adm", "", -1), ("EBITDA", "ebitda", "total", 1), ("DETERIORO", "det", "", -1),
        ("DEPRECIACIONES", "dep", "", -1), ("AMORTIZACIONES", "amo", "", -1),
        ("RESULTADO OPERACIONAL", "operacional", "sub", 1), ("OTROS INGRESOS", "oi", "", 1),
        ("OTROS GASTOS", "og", "", -1), ("INGRESOS FINANCIEROS", "if", "", 1), ("COSTOS FINANCIEROS", "cf", "", -1),
        ("RESULTADO ANTES DE IMPUESTOS", "rai", "sub", 1), ("IMPUESTO CORRIENTE", "impc", "", -1),
        ("IMPUESTO DIFERIDO", "impd", "", -1), ("RESULTADO DEL EJERCICIO", "neto", "total", 1)]


def _diap_pyg(prs, base, leer, empresa, anio, mes, R):
    d = _datos_eeff(leer, empresa, anio, mes)
    if d is None:
        return f"Estado de resultados: no hay estados financieros de {MESES_L[mes - 1]} {anio} para {empresa}."
    _, r, _ = d
    f = r.loc[(anio, mes)]
    k_ant = (anio - 1, mes)
    fb = r.loc[k_ant] if k_ant in r.index else None
    m = MESES[mes - 1]
    R["kpi"]["neto"] = (_M(f["neto"]), f"Utilidad acum. ene - {MESES_L[mes - 1][:3]}",
                        f"{_pct(_div(f['neto'], fb['neto']) - 1, signo=True)} vs {anio - 1}" if fb is not None else
                        f"Margen neto {_pct(f['m_neto'])}", (f["neto"] >= fb["neto"]) if fb is not None else None)
    if fb is not None:
        R["dest"].append(("Resultados.", f"Ingresos acumulados {_M(f['ing'])} "
                                         f"({_pct(_div(f['ing'], fb['ing']) - 1, signo=True)} vs {anio - 1}), EBITDA "
                                         f"{_M(f['ebitda'])} (margen {_pct(f['m_ebitda'])}) y utilidad {_M(f['neto'])} "
                                         f"({_pct(_div(f['neto'], fb['neto']) - 1, signo=True)})."))
    filas = []
    for nombre, c, tipo, sig in _PYG:
        cel = [nombre, _num(f[c + "_m"] / 1e6, dec=1)]
        rojo, verde = set(), set()
        if fb is not None:
            dm, da = f[c + "_m"] - fb[c + "_m"], f[c] - fb[c]
            cel += [_num(fb[c + "_m"] / 1e6, dec=1), _num(dm / 1e6, dec=1), _num(f[c] / 1e6, dec=1),
                    _num(fb[c] / 1e6, dec=1), _num(da / 1e6, dec=1),
                    _pct(_div(f[c], fb[c]) - 1, signo=True) if abs(fb[c]) > 1 else ""]
            rojo = {j for j, v in ((3, dm), (6, da), (7, da)) if v * sig < -0.5}
            verde = {j for j, v in ((3, dm), (6, da), (7, da)) if v * sig > 0.5}
        else:
            cel += [_num(f[c] / 1e6, dec=1)]
        filas.append(dict(celdas=cel, tipo=tipo, rojo=rojo, verde=verde))
    msg = (f"Utilidad acumulada {_M(f['neto'])}, {_pct(_div(f['neto'], fb['neto']) - 1, signo=True)} frente a "
           f"{anio - 1} · margen EBITDA {_pct(f['m_ebitda'])} ({_pp(f['m_ebitda'] - fb['m_ebitda'])})"
           if fb is not None else f"Utilidad acumulada {_M(f['neto'])} · margen EBITDA {_pct(f['m_ebitda'])}")
    s = _clonar(prs, base, f"Estado de resultados {m} {anio}" + (f" - {anio - 1}" if fb is not None else ""), msg)
    if fb is not None:
        _tabla(s, 500000, Y0, 11300000, ["MILLONES", str(anio), str(anio - 1), "VAR", str(anio), str(anio - 1), "VAR",
                                         "%"],
               filas, [3.3, 1.35, 1.35, 1.25, 1.45, 1.45, 1.35, .8], alto_fila=290000, tam=11,
               grupos=[("", 1), (f"MES {m}", 3), (f"ACUMULADO ENE - {m}", 4)])
    else:
        _tabla(s, 500000, Y0, 9000000, ["MILLONES", f"MES {m}", f"ACUM. ENE - {m}"], filas, [3.3, 1.6, 1.7],
               alto_fila=310000, tam=12)
    y = Y0 + 290000 * (len(filas) + 2) + 180000
    for i, (et, c) in enumerate([("% EBITDA", "m_ebitda"), ("% OPERACIONAL", "m_operacional"), ("% NETO", "m_neto"),
                                 ("ROE", "roe")]):
        sub_, col = "", AZUL_TXT
        if fb is not None and not pd.isna(fb[c]) and not pd.isna(f[c]):
            dpp = f[c] - fb[c]
            sub_, col = f"{_pp(dpp)} vs {anio - 1}", (AZUL_TXT if abs(dpp) < 0.0005 else (VERDE if dpp >= 0 else ROJO))
        _tarjeta(s, 500000 + i * 2850000, y, 2700000, 1100000, et, _pct(f[c]), sub_, col, tam=20)
    par = []
    if fb is not None:
        def frase(c, nombre):
            return (f"{nombre} acumulado {_M(f[c])} frente a {_M(fb[c])} en {anio - 1} "
                    f"({_pct(_div(f[c], fb[c]) - 1, signo=True)}); en el mes {_M(f[c + '_m'])} frente a "
                    f"{_M(fb[c + '_m'])}.")
        par.append(("Ingresos.", frase("ing", "Ingreso")))
        par.append(("Costos y gastos.", f"Costos {_M(f['cos'])} ({_pct(_div(f['cos'], fb['cos']) - 1, signo=True)}) y "
                                        f"gastos de administración {_M(f['adm'])} "
                                        f"({_pct(_div(f['adm'], fb['adm']) - 1, signo=True)}) frente a {anio - 1}."))
        par.append(("EBITDA.", frase("ebitda", "EBITDA") + f" Margen {_pct(f['m_ebitda'])} frente a "
                                                           f"{_pct(fb['m_ebitda'])}."))
        par.append(("Resultado.", frase("neto", "Resultado del ejercicio")))
    else:
        par.append(("", f"Ingresos acumulados {_M(f['ing'])}, EBITDA {_M(f['ebitda'])} (margen {_pct(f['m_ebitda'])}) "
                        f"y resultado del ejercicio {_M(f['neto'])}. No hay cifras de {anio - 1} para comparar."))
    _texto(s, 12200000, Y0, 5600000, 5700000, par, tam=15, titulo="COMENTARIOS MES")

    # 2) evolución mes a mes contra el año anterior
    m_act = [m_ for m_ in range(1, mes + 1) if (anio, m_) in r.index]
    if len(m_act) > 1:
        comp = all((anio - 1, m_) in r.index for m_ in m_act)

        def ser(c):
            out = {}
            if comp:
                out[str(anio - 1)] = [r.at[(anio - 1, m_), c] / 1e6 for m_ in m_act]
            out[str(anio)] = [r.at[(anio, m_), c] / 1e6 for m_ in m_act]
            return out
        cats = [MESES[m_ - 1] for m_ in m_act]
        col = (GRIS_OSC, AZUL) if comp else (AZUL,)
        s2 = _clonar(prs, base, "Evolución de resultados",
                     f"{m}: ingresos {_M(f['ing_m'])} y EBITDA {_M(f['ebitda_m'])}"
                     + (f" frente a {_M(fb['ing_m'])} y {_M(fb['ebitda_m'])} en {m} {anio - 1}" if fb is not None
                        else ""))
        _grafico(s2, 450000, Y0, 8650000, 4700000, cats, ser("ing_m"), titulo="Ingresos de cada mes (millones)",
                 resaltar_ultimo=False, tam=11, colores=col)
        _grafico(s2, 9250000, Y0, 8550000, 4700000, cats, ser("ebitda_m"), titulo="EBITDA de cada mes (millones)",
                 resaltar_ultimo=False, tam=11, colores=col)
        em = pd.Series({m_: r.at[(anio, m_), "ebitda_m"] for m_ in m_act})
        im = pd.Series({m_: r.at[(anio, m_), "ing_m"] for m_ in m_act})
        txt = (f"El mejor mes del año en EBITDA es {MESES_L[int(em.idxmax()) - 1]} ({_M(em.max())}) y el más bajo "
               f"{MESES_L[int(em.idxmin()) - 1]} ({_M(em.min())}); {MESES_L[mes - 1]} queda en {_M(em[mes])}, "
               f"{_pct(_div(em[mes], em.mean()) - 1, signo=True)} frente al promedio del año. Los ingresos del mes "
               f"van {_pct(_div(im[mes], im.mean()) - 1, signo=True)} frente a su promedio")
        if comp:
            sube = sum(r.at[(anio, m_), "ing_m"] > r.at[(anio - 1, m_), "ing_m"] for m_ in m_act)
            txt += f" y superan a {anio - 1} en {sube} de {len(m_act)} meses"
        _texto(s2, 1300000, Y0 + 4950000, 13300000, 1700000, [txt + "."], tam=17)
    return None


# ------------------------------------------------------------------ resumen del mes
def _diap_resumen(prs, base, empresa, anio, mes, R):
    if not R["kpi"] and not R["dest"]:
        return "Resumen del mes: no hay información del mes para resumir."
    s = _clonar(prs, base, f"Resumen del mes {MESES_L[mes - 1]} {anio}",
                f"{EMPRESAS.get(empresa, empresa)} · cifras principales y lo más importante del mes · frente al {P.etiqueta(R['version'])}")
    orden = ["ing", "ebitda", "neto", "caja", "nomina", "personas", "comb", "ktnc"]
    tiles = [R["kpi"][k] for k in orden if k in R["kpi"]]
    por_fila = 4 if len(tiles) in (4, 7, 8) else (3 if len(tiles) > 2 else max(len(tiles), 1))
    sep, x0 = 220000, 500000
    w = int((ANCHO - 2 * x0 - sep * (por_fila - 1)) / por_fila)
    for i, (val, et, sub_, bueno) in enumerate(tiles):
        _kpi(s, x0 + (i % por_fila) * (w + sep), Y0 + (i // por_fila) * 1600000, w, 1450000, val, et, sub_, bueno,
             tam=30)
    filas_t = (len(tiles) + por_fila - 1) // por_fila
    y = Y0 + filas_t * 1600000 + 150000
    if R["dest"]:
        _texto(s, 700000, y, 13900000, 8900000 - y, R["dest"], tam=16, titulo="LO MÁS IMPORTANTE DEL MES",
               alinear=PP_ALIGN.LEFT, viñeta=True)
    return None


# ------------------------------------------------------------------ portada y diapositivas de operación
def _cmp_txt(v, ref, nombre: str = "el presupuesto") -> str:
    """'11% por debajo del presupuesto' / '8% por encima del presupuesto'."""
    d = _div(v, ref) - 1
    if pd.isna(d):
        return ""
    return f"{_pct(abs(d), 0 if abs(d) >= .01 else 1)} {'por encima' if d > 0 else 'por debajo'} de {nombre}"


def _diap_portada(prs, base, empresa, anio, mes):
    s = _clonar(prs, base, "Resultados área admin y finanzas",
                f"{EMPRESAS.get(empresa, empresa)} · {MESES_L[mes - 1].capitalize()} {anio}")
    return s


def _diap_corte(prs, base, leer, empresa, anio, mes, R):
    if empresa != "ALDC":
        return "Corte de césped: el módulo Corte y poda no distingue compañía, así que solo se incluye en la presentación de ALDC."
    import mod_corte_poda as CP
    pp = leer("SELECT * FROM cp_ppto WHERE anio=?", (anio,))
    ej = leer("SELECT * FROM cp_ejec WHERE anio=?", (anio,))
    ml = MESES_L[mes - 1]
    pp, usada = P.filtrar(pp, R["version"])
    if pp.empty:
        return (f"Corte de césped: no hay presupuesto de corte de {anio} (Administración → Corte y poda)." if usada is not None or not len(leer("SELECT 1 FROM cp_ppto WHERE anio=? LIMIT 1", (anio,)))
                else f"Corte de césped: no hay presupuesto de corte {R['version']} cargado para {anio} (Administración → Corte y poda).")
    _nota_version(R, "Corte de césped", usada)
    t = CP.tabla_mensual(pp, ej if not ej.empty else pd.DataFrame(columns=["anio", "mes", "m2", "valor_corte", "arboles", "valor_poda"]))
    fila = t[t["mes"] == mes]
    if fila.empty or pd.isna(fila.iloc[0]["m2_ej"]):
        return f"Corte de césped: no hay ejecución de {ml} {anio} cargada."
    f = fila.iloc[0]
    hasta = t[(t["mes"] <= mes) & t["m2_ej"].notna()]
    m2_p, m2_e = f["m2"], f["m2_ej"]
    v_p, v_e = f["valor_corte_mta"], f["corte_ej_mta"]
    dm2 = m2_e - m2_p
    ant = t[(t["mes"] == mes - 1) & t["m2_ej"].notna()]
    ac_m2p, ac_m2e = hasta["m2"].sum(), hasta["m2_ej"].sum()
    ac_vp, ac_ve = hasta["valor_corte_mta"].sum(min_count=1), hasta["corte_ej_mta"].sum(min_count=1)
    R["dest"].append(("Corte de césped.", f"{_num(m2_e)} m² en {ml} ({_ent(dm2)} m² frente al presupuesto), "
                                          f"{_M(v_e)} frente a {_M(v_p)} presupuestados."))
    R["kpi"]["corte"] = (_M(v_e), f"Corte césped {ml}", f"{_pct(_div(v_e, v_p) - 1, signo=True)} vs ppto"
                         if not pd.isna(v_p) else "", (v_e <= v_p) if not pd.isna(v_p) else None)
    s = _clonar(prs, base, "Detalle costo corte césped",
                f"{ml.capitalize()}: {_num(m2_e)} m² cortados, {_ent(dm2)} m² frente al presupuesto · "
                f"{_M(v_e)} ejecutados vs {_M(v_p)}")
    cats = [MESES[int(m_) - 1] for m_ in t["mes"]]
    _grafico(s, 450000, Y0, 17400000, 3300000, cats,
             {"Presupuesto": (t["m2"] / 1e3).tolist(), "Ejecución": (t["m2_ej"] / 1e3).tolist()},
             titulo="Área cortada por mes (miles de m²)", formato="#,##0", resaltar_ultimo=False, tam=11,
             colores=(AZUL, NARANJA), desde_cero=True)
    ver = t[t["mes"] <= mes]
    var = ver["corte_ej_mta"] - ver["valor_corte_mta"]
    filas = [
        dict(celdas=["PPTO m²"] + [_num(v) for v in ver["m2"]] + [_num(ac_m2p)]),
        dict(celdas=["EJECUCIÓN m²"] + [_num(v) for v in ver["m2_ej"]] + [_num(ac_m2e)]),
        dict(celdas=["PPTO $M"] + [_num(v / 1e6, dec=1) for v in ver["valor_corte_mta"]] + [_num(ac_vp / 1e6, dec=1)]),
        dict(celdas=["EJECUCIÓN $M"] + [_num(v / 1e6, dec=1) for v in ver["corte_ej_mta"]] + [_num(ac_ve / 1e6, dec=1)]),
        dict(celdas=["VARIACIÓN $M"] + [_num(v / 1e6, dec=1) for v in var] + [_num((ac_ve - ac_vp) / 1e6, dec=1)],
             tipo="total"),
    ]
    _tabla(s, 450000, Y0 + 3400000, 17400000, ["", *[MESES[int(m_) - 1] for m_ in ver["mes"]], "ACUM."], filas,
           [1.9] + [1.0] * len(ver) + [1.25], alto_fila=290000, tam=11)
    y = Y0 + 3400000 + 290000 * 6 + 130000
    _fila_tarjetas(s, y, [
        ("ÁREA DEL MES", f"{_num(m2_e)} m²", f"{_ent(dm2)} m² vs ppto ({_pct(_div(m2_e, m2_p) - 1, signo=True)})",
         dm2 <= 0),
        ("COSTO DEL MES", _M(v_e), f"{_pct(_div(v_e, v_p) - 1, signo=True)} vs ppto" if not pd.isna(v_p) else "",
         (v_e <= v_p) if not pd.isna(v_p) else None),
        ("ACUMULADO", _M(ac_ve), f"{_M(ac_ve - ac_vp, True)} vs ppto acum." if not pd.isna(ac_vp) else "",
         (ac_ve <= ac_vp) if not pd.isna(ac_vp) else None),
        ("COSTO POR m²", "$" + _num(_div(v_e, m2_e)), f"ppto ${_num(_div(v_p, m2_p))}" if not pd.isna(v_p) else "", None),
    ], alto=900000, tam=18)
    txt = (f"En {ml} se cortaron {_num(m2_e)} m² de césped, {_num(abs(dm2))} m² {'más' if dm2 > 0 else 'menos'} de los "
           f"{_num(m2_p)} m² presupuestados ({_pct(_div(m2_e, m2_p) - 1, signo=True)}). El costo fue {_M(v_e)} frente a "
           f"{_M(v_p)} del presupuesto ({_cmp_txt(v_e, v_p)})")
    if len(ant):
        txt += f"; frente a {MESES_L[mes - 2]} el área cambió {_pct(_div(m2_e, ant.iloc[0]['m2_ej']) - 1, signo=True)}"
    txt += (f". Acumulado a {ml}: {_num(ac_m2e)} m² y {_M(ac_ve)} ({_cmp_txt(ac_ve, ac_vp)}).")
    _texto(s, 1300000, y + 1000000, 13500000, 900000, [txt], tam=15)
    return None


def _diap_operativo(prs, base, leer, empresa, anio, mes, R):
    ctx = R.get("ctx")
    ml = MESES_L[mes - 1]
    if ctx is None:
        return "Consumo operativo: faltó la conexión a la base para armarlo."
    import operaciones as OP
    series, esc = OP.serie_mensual(ctx, empresa, anio, mes, R["version"])
    if series is None:
        return f"Consumo operativo: {esc}."
    nombres = {"Bolsas": "Bolsas", "Herramientas": "Herramientas"}
    hojas = [h for h in nombres if h in series]
    ok = [h for h in hojas if (series[h].loc[series[h]["mes"] == mes, "ejecutado"].fillna(0) > 0).any()]
    if not ok:
        return (f"Consumo operativo: no hay ejecución de {ml} {anio} para {empresa} (se arma con el SISPO guardado en "
                "Administración → Operaciones).")
    s = _clonar(prs, base, "Consumo operativo", "")
    msgs, tarj, textos = [], [], []
    w = 8650000
    for i, h in enumerate(hojas):
        t = series[h].copy()
        t.loc[t["ejecutado"].fillna(0) <= 0, "ejecutado"] = float("nan")     # sin ejecución cargada: no es cero
        cats = [MESES[int(m_) - 1] for m_ in t["mes"]]
        _grafico(s, 450000 + i * (w + 100000), Y0 - 50000, w, 3700000, cats,
                 {"Presupuesto": (t["presupuesto"] / 1e6).tolist(), "Ejecución": (t["ejecutado"] / 1e6).tolist()},
                 titulo=f"{nombres[h]} · millones de cops ({esc})", formato="#,##0.0", resaltar_ultimo=False, tam=11,
                 colores=(AZUL, NARANJA), desde_cero=True)
        fila = t[t["mes"] == mes].iloc[0]
        hasta = t[(t["mes"] <= mes) & t["ejecutado"].notna()]       # presupuesto y ejecución cubren los mismos meses
        ac_e, ac_p = hasta["ejecutado"].sum(), hasta["presupuesto"].sum()
        e1, p1 = fila["ejecutado"], fila["presupuesto"]
        ant = t[(t["mes"] == mes - 1) & t["ejecutado"].notna()]
        e0 = ant.iloc[0]["ejecutado"] if len(ant) else float("nan")
        tarj.append((f"{nombres[h].upper()} {MESES[mes - 1]}", _M(e1), f"{_cmp_txt(e1, p1, 'el ppto')}" if p1 else "",
                     (e1 <= p1) if p1 else None))
        tarj.append((f"{nombres[h].upper()} ACUM.", _M(ac_e), f"{_M(ac_e - ac_p, True)} vs ppto", ac_e <= ac_p))
        tx = (f"En {nombres[h].lower()}, {ml} cerró en {_M(e1)}, {_cmp_txt(e1, p1)} ({_M(p1)})"
              if p1 else f"En {nombres[h].lower()}, {ml} cerró en {_M(e1)} (sin presupuesto mensual)")
        if not pd.isna(e0):
            tx += f" y {_pct(abs(_div(e1, e0) - 1), 0)} {'más' if e1 > e0 else 'menos'} que {MESES_L[mes - 2]} ({_M(e0)})"
        tx += f". Acumulado {_M(ac_e)} frente a {_M(ac_p)} presupuestados."
        textos.append(tx)
        R["dest"].append((f"{nombres[h]}.", tx))
        msgs.append(f"{nombres[h].lower()} {_M(e1)}")
    tb = s.shapes.add_textbox(Emu(1200000), Emu(1480000), Emu(ANCHO - 2400000), Emu(420000))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    _fuente(p.add_run(), F_TXT, 19, True, AZUL_TXT)
    p.runs[0].text = f"{ml.capitalize()}: " + " · ".join(msgs)
    y = Y0 + 3550000
    _fila_tarjetas(s, y, tarj, x0=1500000, x1=ANCHO - 1500000, alto=1000000, tam=18)
    _texto(s, 1300000, y + 1150000, 13500000, 2100000, textos, tam=15)
    return None


def _flota_mensual(ctx, empresa, anio, version=None):
    """Presupuesto y ejecución del sub-rubro «Flota y Equipo de transporte» por mes (misma base que la pantalla de Mantenimiento), con el
    presupuesto de la versión elegida (si no está cargada, la más cercana)."""
    import logica_m as L
    import mod_mantenimiento as MM
    version = P.version_para(ctx.leer, empresa, anio, version)
    ppto, ejec = MM._base(ctx, empresa, anio, version)
    k = cargas.norm(L.SUBRUBRO_FLOTA)
    p = ppto[ppto["_sub"] == k].groupby("mes")["valor"].sum() if len(ppto) else pd.Series(dtype=float)
    e = ejec[ejec["_sub"] == k].groupby("mes")["valor"].sum() if len(ejec) else pd.Series(dtype=float)
    return p, e, version


def _diap_flota(prs, base, leer, empresa, anio, mes, R):
    ctx = R.get("ctx")
    ml = MESES_L[mes - 1]
    if ctx is None:
        return "Mantenimiento de flota: faltó la conexión a la base para armarlo."
    p, e, version = _flota_mensual(ctx, empresa, anio, R["version"])
    _nota_version(R, "Mantenimiento de flota", version)
    e = e[e.index <= mes]
    if e.empty or mes not in e.index or abs(e.get(mes, 0.0)) < 1:
        return f"Mantenimiento de flota: no hay ejecución de {ml} {anio} para {empresa}."
    meses = [m_ for m_ in range(1, mes + 1) if m_ in e.index]
    cats = [MESES[m_ - 1] for m_ in meses]
    c1 = float(e[mes])
    c0 = float(e[mes - 1]) if (mes - 1) in e.index and abs(e[mes - 1]) > 1 else float("nan")
    total, prom = float(e.sum()), float(e.mean())
    mx = e.idxmax()
    con_ppto = len(p) and float(p.reindex(meses).fillna(0).sum()) > 0
    p1 = float(p.get(mes, 0.0)) if con_ppto else float("nan")
    ac_p = float(p.reindex(meses).fillna(0).sum()) if con_ppto else float("nan")
    R["kpi"]["flota"] = (_M(c1), f"Flota {ml}", f"{_pct(_div(c1, c0) - 1, signo=True)} vs {MESES_L[mes - 2]}"
                         if not pd.isna(c0) else "", (c1 <= c0) if not pd.isna(c0) else None)
    s = _clonar(prs, base, "Mantenimiento de flota",
                f"{ml.capitalize()}: {_M(c1)}" + (f", {_pct(_div(c1, c0) - 1, signo=True)} frente a {MESES_L[mes - 2]}"
                                                  if not pd.isna(c0) else "")
                + (f" · {_cmp_txt(c1, p1)}" if con_ppto and p1 > 0 else ""))
    series = {"Ejecución": [e[m_] / 1e6 for m_ in meses]}
    if con_ppto:
        series = {"Presupuesto": [float(p.get(m_, 0.0)) / 1e6 for m_ in meses], **series}
    _grafico(s, 450000, Y0, 17400000, 4300000, cats, series, titulo="Consumo de mantenimiento de flota (millones de cops)",
             formato="#,##0.0", tam=12, colores=(AZUL, NARANJA) if con_ppto else (AZUL,),
             resaltar_ultimo=not con_ppto, desde_cero=True)
    tarj = [("CONSUMO DEL MES", _M(c1), f"{_pct(_div(c1, c0) - 1, signo=True)} vs {MESES_L[mes - 2]}"
             if not pd.isna(c0) else "", (c1 <= c0) if not pd.isna(c0) else None),
            ("TOTAL ACUMULADO", _M(total), f"Ene - {MESES[mes - 1].capitalize()}", None),
            ("PROMEDIO MES", _M(prom), f"{len(meses)} {'mes' if len(meses) == 1 else 'meses'} · mayor: {MESES_L[mx - 1]}", None)]
    if con_ppto:
        tarj.append(("VS PRESUPUESTO ACUM.", _M(total - ac_p, True), _cmp_txt(total, ac_p, "el ppto"), total <= ac_p))
    _fila_tarjetas(s, Y0 + 4450000, tarj, x0=1500000, x1=ANCHO - 1500000, alto=1000000, tam=19)
    txt = f"El consumo operativo de mantenimiento de vehículos en {ml} cerró en {_M(c1)}"
    if not pd.isna(c0):
        txt += (f", {'un incremento' if c1 > c0 else 'una reducción'} del {_pct(abs(_div(c1, c0) - 1))} frente a "
                f"{MESES_L[mes - 2]} ({_M(c0)})")
    if con_ppto and p1 > 0:
        txt += f" y {_cmp_txt(c1, p1)} del mes ({_M(p1)})"
    txt += (f". {MESES_L[mx - 1].capitalize()} " + ("es el mes de mayor consumo del año" if mx != mes else
            "es el mes de mayor consumo del año hasta ahora") + f" ({_M(e[mx])}). El acumulado enero - {ml} suma "
            f"{_M(total)}, con un promedio mensual de {_M(prom)}"
            + (f"; frente al presupuesto acumulado ({_M(ac_p)}) la diferencia es {_M(total - ac_p, True)}" if con_ppto else "")
            + ".")
    R["dest"].append(("Flota.", f"{_M(c1)} en {ml}" + (f" ({_pct(_div(c1, c0) - 1, signo=True)} vs {MESES_L[mes - 2]})"
                                                       if not pd.isna(c0) else "") + f"; acumulado {_M(total)}."))
    _texto(s, 1300000, Y0 + 5600000, 13500000, 1600000, [txt], tam=16)
    _nota(s, 1300000, Y0 + 7250000, "Fuente: SAP, sub-rubro Flota y Equipo de transporte (presupuesto " + version + ").",
          w=12000000)
    return None


# ------------------------------------------------------------------ facturación del servicio de aseo (digitada a mano)
FACT_COMPONENTES = ["Barrido y limpieza", "Limpieza urbana", "Comercialización", "Recolección y transporte",
                    "Comercialización no aprovechable", "Servicios especiales", "Intereses"]
FACT_TOTAL = "Facturación total"
FACT_DDL = ("CREATE TABLE IF NOT EXISTS facturacion_mes(empresa TEXT, anio INT, mes INT, componente TEXT, valor DOUBLE PRECISION, "
            "actualizado TEXT)")


# Facturación enero - septiembre 2026 de ALDC tomada de la presentación «Faltante» (componentes en millones de pesos
# (exactos a $1.000); facturación total en pesos). Solo se usa para precargar los meses que aún no se han digitado.
FACT_HISTORICO = {("ALDC", 2026): {
    "componentes_millones": {
        "Barrido y limpieza": [6286.291, 7514.331, 6911.685, 7315.183, 7237.413, 7612.388, 7488.201, 7567.089, 7595.451],
        "Limpieza urbana": [1246.77, 1471.696, 1366.021, 1459.282, 1448.744, 1481.421, 1450.083, 1463.053, 1483.85],
        "Comercialización": [1437.879, 1496.518, 1385.406, 1479.46, 1479.076, 1515.805, 1483.22, 1505.641, 1542.251],
        "Recolección y transporte": [3279.238, 3500.305, 3189.561, 3433.381, 3440.782, 3528.775, 3436.023, 3425.62, 3436.195],
        "Comercialización no aprovechable": [267.418, 278.344, 257.681, 275.172, 275.106, 281.935, 275.864, 280.042, 286.862],
        "Servicios especiales": [16.355, 39.208, 24.481, 2.281, 56.333, 35.105, 21.459, 43.913, 30.098],
        "Intereses": [31.919, 49.417, 29.867, 36.294, 37.356, 38.458, 38.516, 41.202, 39.225],
    },
    "total": [
        18794125955.58, 21436230866.82, 20316834705.95,
        20203736805.24, 20981169644.43, 21275829230.96,
        20917666379.78, 21242363101.0, 22065377980.88,
    ],
}}


def _apilado(s, x, y, w, h, categorias, series, titulo=None, tam=11):
    """Columnas apiladas (composición de cada mes). series: {nombre: [valores]}."""
    cd = CategoryChartData()
    cd.categories = list(categorias)
    for nombre, vals in series.items():
        cd.add_series(nombre, [None if v is None or pd.isna(v) else float(v) for v in vals])
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, Emu(x), Emu(y), Emu(w), Emu(h), cd)
    ch = gf.chart
    ch.font.name, ch.font.size = F_TXT, Pt(tam)
    _titulo_grafico(ch, titulo, tam)
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch.legend.include_in_layout = False
    ch.legend.font.size = Pt(tam - 1)
    pl = ch.plots[0]
    pl.gap_width = 50
    paleta = (NAVY, AZUL, "8FAADC", "B4C7E7", NARANJA, GRIS_OSC, VERDE_BAR)
    for i, ser in enumerate(pl.series):
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = RGBColor.from_string(paleta[i % len(paleta)])
    _estilo_ejes(ch, "#,##0", tam)
    ch.value_axis.minimum_scale = 0
    return gf


def _barras_part(s, x, y, w, h, pares, titulo=None, tam=11):
    """Barras horizontales de participación (0-1), de mayor a menor."""
    cd = CategoryChartData()
    cd.categories = [_corto(n, 30) for n, _ in pares]
    cd.add_series("Participación", [float(v) for _, v in pares])
    gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Emu(x), Emu(y), Emu(w), Emu(h), cd)
    ch = gf.chart
    ch.font.name, ch.font.size = F_TXT, Pt(tam)
    _titulo_grafico(ch, titulo, tam)
    ch.has_legend = False
    pl = ch.plots[0]
    pl.gap_width = 45
    pl.has_data_labels = True
    pl.data_labels.number_format, pl.data_labels.number_format_is_linked = "0.0%", False
    pl.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    pl.data_labels.font.size = Pt(tam - 1)
    pl.series[0].format.fill.solid()
    pl.series[0].format.fill.fore_color.rgb = RGBColor.from_string(NARANJA)
    _estilo_ejes(ch, "0%", tam)
    ch.value_axis.minimum_scale = 0
    ch.value_axis.maximum_scale = max(0.1, max(v for _, v in pares) * 1.25)
    ch.value_axis.has_major_gridlines = False
    ch.value_axis.tick_label_position = XL_TICK_LABEL_POSITION.NONE
    ch.category_axis.reverse_order = True
    return gf


def _diap_facturacion(prs, base, leer, empresa, anio, mes, R):
    ml = MESES_L[mes - 1]
    try:
        d = leer("SELECT mes, componente, valor FROM facturacion_mes WHERE empresa=? AND anio=? AND mes<=?",
                 (empresa, anio, mes))
    except Exception:
        d = pd.DataFrame(columns=["mes", "componente", "valor"])
    if d.empty or not (d["mes"] == mes).any():
        return (f"Facturación servicio de aseo: no hay valores de {ml} {anio} para {empresa}. Se digitan en "
                "Administración → Presentación financiera → Facturación del servicio de aseo.")
    d["valor"] = pd.to_numeric(d["valor"], errors="coerce")
    pv = d.pivot_table(index="mes", columns="componente", values="valor", aggfunc="sum")
    comps = [c for c in FACT_COMPONENTES if c in pv.columns and pv[c].fillna(0).abs().sum() > 0]
    if not comps:
        return f"Facturación servicio de aseo: {ml} {anio} solo trae la facturación total, sin componentes."
    meses = [m_ for m_ in range(1, mes + 1) if m_ in pv.index]
    cats = [MESES[m_ - 1] for m_ in meses]
    comp_m = pv[comps].fillna(0).sum(axis=1)                       # suma de componentes de cada mes
    tot_m = pv[FACT_TOTAL] if FACT_TOTAL in pv.columns else pd.Series(dtype=float)
    tot_m = tot_m.dropna()
    c1 = float(comp_m.get(mes, 0.0))
    t1 = float(tot_m.get(mes, float("nan")))
    c0 = float(comp_m.get(mes - 1, float("nan"))) if (mes - 1) in comp_m.index else float("nan")
    ac_c = float(comp_m.reindex(meses).sum())
    m_con_total = [m_ for m_ in meses if m_ in tot_m.index]
    ac_t = float(tot_m.reindex(m_con_total).sum()) if m_con_total else float("nan")
    ac_c_t = float(comp_m.reindex(m_con_total).sum()) if m_con_total else float("nan")   # mismos meses que el total
    part = _div(ac_c_t, ac_t)
    part1 = _div(c1, t1)
    R["kpi"]["fact"] = (_M(c1), f"Componentes {ml}", f"{_pct(_div(c1, c0) - 1, signo=True)} vs {MESES_L[mes - 2]}"
                        if not pd.isna(c0) else "", (c1 >= c0) if not pd.isna(c0) else None)
    s = _clonar(prs, base, "Facturación servicio de aseo",
                f"{ml.capitalize()}: componentes {_M(c1)}" + (f" · {_pct(part1)} de la facturación total"
                                                               if not pd.isna(part1) else ""))
    series = {"Componentes": [comp_m.get(m_, 0.0) / 1e6 for m_ in meses]}
    if len(tot_m):
        series["Facturación total"] = [tot_m.get(m_, float("nan")) / 1e6 for m_ in meses]
    _grafico(s, 450000, Y0, 17400000, 4300000, cats, series, titulo="Componentes vs facturación total (millones de cops)",
             formato="#,##0", tam=12, resaltar_ultimo=False, colores=(AZUL, NARANJA), desde_cero=True)
    _fila_tarjetas(s, Y0 + 4450000, [
        ("COMPONENTES " + empresa, _M(ac_c), f"Acum. Ene - {MESES[mes - 1].capitalize()}", None),
        ("FACTURACIÓN TOTAL", _M(ac_t) if not pd.isna(ac_t) else "—",
         f"Acum. de {len(m_con_total)} {'mes' if len(m_con_total) == 1 else 'meses'}" if m_con_total else "sin digitar", None),
        ("% PARTICIPACIÓN", _pct(part) if not pd.isna(part) else "—", "Componentes / facturación", None),
    ], x0=2500000, x1=ANCHO - 2500000, alto=1050000, tam=22)
    txt = f"Los componentes de {empresa} sumaron {_M(c1)} en {ml}"
    if not pd.isna(c0):
        txt += f" ({_pct(_div(c1, c0) - 1, signo=True)} frente a {MESES_L[mes - 2]}, {_M(c0)})"
    if not pd.isna(t1):
        txt += f", equivalentes al {_pct(part1)} de la facturación total del mes ({_M(t1)})"
    txt += f". Acumulado enero - {ml}: {_M(ac_c)}" + (f"; sobre la facturación total acumulada de {_M(ac_t)} la participación es "
                                                       f"{_pct(part)}" if not pd.isna(part) else "") + "."
    if pd.isna(t1):
        txt += " La facturación total del mes no está digitada: la participación del mes no se calcula."
    _texto(s, 1300000, Y0 + 5700000, 13500000, 1400000, [txt], tam=16)
    R["dest"].append(("Facturación.", f"Componentes {_M(c1)} en {ml}" + (f", {_pct(part1)} de la facturación total"
                                                                       if not pd.isna(part1) else "") + "."))

    # 2) composición por componente
    s2 = _clonar(prs, base, "Facturación servicio de aseo",
                 f"Composición de los componentes de {empresa} · enero - {ml}")
    _apilado(s2, 450000, Y0, 9300000, 3800000, cats,
             {c: [pv[c].get(m_, 0.0) / 1e6 if not pd.isna(pv[c].get(m_, float("nan"))) else 0.0 for m_ in meses]
              for c in comps}, titulo="Composición de ingresos por componente (millones de cops)", tam=11)
    ac_comp = {c: float(pv[c].reindex(meses).fillna(0).sum()) for c in comps}
    pares = sorted(((c, _div(v, ac_c)) for c, v in ac_comp.items()), key=lambda q: -q[1])
    _barras_part(s2, 9900000, Y0, 7950000, 3800000, pares, titulo="Participación promedio por componente (%)")
    filas = []
    for c in comps:
        filas.append(dict(celdas=[c] + [_num(pv[c].get(m_, 0.0) / 1e6, dec=1) for m_ in meses] + [_num(ac_comp[c] / 1e6, dec=1)]))
    filas.append(dict(celdas=["Total componentes"] + [_num(comp_m.get(m_, 0.0) / 1e6, dec=1) for m_ in meses]
                      + [_num(ac_c / 1e6, dec=1)], tipo="total"))
    alto = min(300000, int(3000000 / (len(filas) + 1)))
    _tabla(s2, 450000, Y0 + 3950000, 17400000, ["COMPONENTE (M$)", *cats, "ACUM."], filas,
           [2.6] + [1.0] * len(meses) + [1.2], alto_fila=alto, tam=10)
    mayor = pares[0]
    txt2 = (f"{mayor[0]} es el componente de mayor peso, con el {_pct(mayor[1], 0)} de lo facturado en componentes en el acumulado "
            f"({_M(ac_comp[mayor[0]])}).")
    if len(pares) > 1:
        txt2 += f" Le siguen {pares[1][0].lower()} ({_pct(pares[1][1], 0)})" + (f" y {pares[2][0].lower()} ({_pct(pares[2][1], 0)})"
                                                                                if len(pares) > 2 else "") + "."
    if mes - 1 in pv.index:
        dif = {c: pv[c].get(mes, 0.0) - pv[c].get(mes - 1, 0.0) for c in comps}
        k = max(dif, key=lambda c: abs(dif[c]))
        if abs(dif[k]) > 1e5:
            txt2 += f" El mayor cambio frente a {MESES_L[mes - 2]} es {k.lower()} ({_M(dif[k], True)})."
    _texto(s2, 1300000, Y0 + 3950000 + alto * (len(filas) + 1) + 120000, 13500000, 900000, [txt2], tam=14)
    return None


# ------------------------------------------------------------------ flota por placa y operaciones por ítem
def _diap_flota_placas(prs, base, leer, empresa, anio, mes, R):
    ctx = R.get("ctx")
    ml = MESES_L[mes - 1]
    if ctx is None:
        return "Flota por placa: faltó la conexión a la base para armarlo."
    import logica_m as L
    import mod_mantenimiento as MM
    v, esc = MM.datos_placas(ctx, empresa, anio, mes, R["version"])
    if v is None:
        return f"Flota por placa: {esc}."
    nodos = L.arbol_placas(v)
    reales = [x for x in nodos if x["placa"] != L.SIN_PLACA]
    tot_e = sum(float(x["ejecutado"]) for x in reales)
    if tot_e < 1:
        return f"Flota por placa: no hay ejecución por placa de {ml} {anio} para {empresa}."
    tot_p = sum(float(x["presupuesto"]) for x in reales)
    con_ej = [x for x in reales if float(x["ejecutado"]) > 0]
    sobre = [x for x in con_ej if float(x["ejecutado"]) > float(x["presupuesto"]) + 1]
    sin_ppto = [x for x in con_ej if float(x["presupuesto"]) <= 0]
    top = MM._top_desviacion(nodos, 10)
    top5 = sum(sorted((float(x["ejecutado"]) for x in con_ej), reverse=True)[:5])
    s = _clonar(prs, base, "Flota por placa",
                f"{ml.capitalize()}: {len(sobre)} de {len(con_ej)} placas con ejecución superan su presupuesto · "
                f"las 10 de mayor diferencia suman {_M(sum(float(x['ejecutado']) - float(x['presupuesto']) for x in top), True)}")

    def concepto_principal(x):
        h = x["hijos"]
        h = h[h["ejecutado"] > 0].sort_values("ejecutado", ascending=False)
        return (h.iloc[0]["concepto"], float(h.iloc[0]["ejecutado"])) if len(h) else ("", 0.0)

    pares = [(x["placa"], float(x["ejecutado"]) - float(x["presupuesto"])) for x in top]
    _barras_var(s, 450000, Y0, 6500000, 4650000, pares, titulo=f"Diferencia vs presupuesto (millones) · rojo: se pasó",
                tam=11, bueno_si_sube=False)
    filas = []
    for x in top:
        cn, cv = concepto_principal(x)
        d = float(x["ejecutado"]) - float(x["presupuesto"])
        filas.append(dict(celdas=[x["placa"], _corto(str(x.get("tipo") or ""), 18), _num(float(x["presupuesto"]) / 1e6, dec=1),
                                  _num(float(x["ejecutado"]) / 1e6, dec=1), _num(d / 1e6, dec=1),
                                  f"{_corto(cn, 24)} ({_num(cv / 1e6, dec=1)})" if cn else ""],
                         rojo={4} if d > 1 else (), verde={4} if d < -1 else ()))
    filas.append(dict(celdas=["Total placas", "", _num(tot_p / 1e6, dec=1), _num(tot_e / 1e6, dec=1),
                              _num((tot_e - tot_p) / 1e6, dec=1), ""], tipo="total"))
    alto = min(330000, int(4400000 / (len(filas) + 1)))
    _tabla(s, 7100000, Y0, 10750000, ["PLACA", "TIPO", "PPTO M$", "EJEC. M$", "DIF. M$", "CONCEPTO PRINCIPAL (M$)"], filas,
           [1.0, 1.5, .9, .9, .9, 3.0], alto_fila=alto, tam=10)
    _fila_tarjetas(s, Y0 + 4800000, [
        ("PLACAS CON EJECUCIÓN", _num(len(con_ej)), f"{len(sobre)} sobre presupuesto ({_pct(_div(len(sobre), len(con_ej)), 0)})",
         len(sobre) <= len(con_ej) / 2),
        ("EJECUTADO EN PLACAS", _M(tot_e), f"{_M(tot_e - tot_p, True)} vs ppto {_M(tot_p)}", tot_e <= tot_p),
        ("CONCENTRACIÓN TOP 5", _pct(_div(top5, tot_e), 0), "de la ejecución del mes", None),
        ("SIN PRESUPUESTO", str(len(sin_ppto)), f"{_M(sum(float(x['ejecutado']) for x in sin_ppto))} ejecutados" if sin_ppto else "ninguna", None
         if not sin_ppto else False),
    ], alto=950000, tam=18)
    p1 = top[0]
    cn1, cv1 = concepto_principal(p1)
    txt = (f"{p1['placa']} es la placa con mayor diferencia: {_M(float(p1['ejecutado']))} ejecutados frente a "
           f"{_M(float(p1['presupuesto']))} presupuestados ({_M(float(p1['ejecutado']) - float(p1['presupuesto']), True)})"
           + (f", sobre todo por {cn1.lower()} ({_M(cv1)})" if cn1 else "") + ".")
    if len(top) > 1:
        txt += " Le siguen " + ", ".join(f"{x['placa']} ({_M(float(x['ejecutado']) - float(x['presupuesto']), True)})"
                                         for x in top[1:3]) + "."
    txt += f" Las 5 placas de mayor gasto concentran el {_pct(_div(top5, tot_e), 0)} de la ejecución en placas."
    _texto(s, 1300000, Y0 + 5900000, 13500000, 1300000, [txt], tam=15)
    R["dest"].append(("Placas.", f"{len(sobre)} de {len(con_ej)} placas superan su presupuesto; la mayor diferencia es {p1['placa']} "
                                 f"({_M(float(p1['ejecutado']) - float(p1['presupuesto']), True)})."))

    # 2) por concepto de gasto
    ppc = v["por_placa_concepto"]
    ppc = ppc[ppc["placa"] != L.SIN_PLACA]
    pc = ppc.groupby("concepto")[["presupuesto", "ejecutado"]].sum()
    pc = pc[(pc["presupuesto"] != 0) | (pc["ejecutado"] != 0)].sort_values("ejecutado", ascending=False)
    if len(pc) >= 2:
        pend = float(pc.loc[L.PENDIENTE, "ejecutado"]) if L.PENDIENTE in pc.index else 0.0
        clasif = pc.drop(index=L.PENDIENTE, errors="ignore")
        clasif = clasif if len(clasif) else pc
        s2 = _clonar(prs, base, "Flota por concepto de gasto",
                     f"{ml.capitalize()}: {_corto(clasif.index[0], 30).lower()} es el concepto de mayor ejecución "
                     f"({_pct(_div(clasif['ejecutado'].iloc[0], pc['ejecutado'].sum()), 0)} de lo ejecutado en placas)")
        g = pc.head(8)
        _grafico(s2, 450000, Y0, 10600000, 4700000, [_corto(c, 20) for c in g.index],
                 {"Presupuesto": (g["presupuesto"] / 1e6).tolist(), "Ejecución": (g["ejecutado"] / 1e6).tolist()},
                 titulo="Presupuesto vs ejecución por concepto de gasto (millones)", formato="#,##0.0", resaltar_ultimo=False,
                 tam=10, colores=(AZUL, NARANJA), desde_cero=True)
        dif = (pc["ejecutado"] - pc["presupuesto"])
        dif = dif[dif.abs() > 1e5]
        top_d = pd.concat([dif.nlargest(4), dif.nsmallest(4)])
        top_d = top_d[~top_d.index.duplicated()]
        _barras_var(s2, 11200000, Y0, 6650000, 4700000, sorted(top_d.items(), key=lambda q: -q[1]),
                    titulo="Diferencia vs presupuesto (millones)", tam=10, bueno_si_sube=False)
        pasan = dif.drop(index=L.PENDIENTE, errors="ignore")
        pasan = pasan[pasan > 0].sort_values(ascending=False)
        txt2 = (f"En {ml}, {clasif.index[0].lower()} fue el concepto de mayor ejecución ({_M(clasif['ejecutado'].iloc[0])}"
                + (f"; presupuesto {_M(clasif['presupuesto'].iloc[0])}" if clasif["presupuesto"].iloc[0] else "") + ").")
        if pend > 1:
            txt2 += (f" Atención: {_M(pend)} ({_pct(_div(pend, pc['ejecutado'].sum()), 0)} de lo ejecutado) está pendiente de "
                     "clasificar por concepto; mientras tanto la comparación por concepto es parcial.")
        if len(pasan):
            txt2 += " Superan su presupuesto: " + ", ".join(f"{c.lower()} ({_M(x, True)})" for c, x in pasan.head(3).items()) + "."
        sin_p = clasif[(clasif["presupuesto"] <= 0) & (clasif["ejecutado"] > 1)]
        if len(sin_p):
            txt2 += (f" Sin presupuesto en el mes: {', '.join(c.lower() for c in sin_p.index[:3])} "
                     f"({_M(sin_p['ejecutado'].sum())}).")
        _texto(s2, 1300000, Y0 + 4950000, 13500000, 1400000, [txt2], tam=15)
    return None


def _diap_operativo_items(prs, base, leer, empresa, anio, mes, R):
    """Un solo slide: Herramientas por ítem (donde el detalle aporta) y, en una franja, los ítems de Bolsas."""
    ctx = R.get("ctx")
    ml = MESES_L[mes - 1]
    if ctx is None:
        return "Operaciones por ítem: faltó la conexión a la base para armarlo."
    import operaciones as OP
    its, esc = OP.items_mensual(ctx, empresa, anio, mes, R["version"])
    if its is None:
        return f"Operaciones por ítem: {esc}."

    def limpio(h):
        g = its.get(h)
        if g is None or g.empty:
            return None
        g = g[(g["presupuesto"] != 0) | (g["ejecutado"] != 0)].copy()
        g["item"] = g["item"].astype(str).str.replace(" · sin ítem", "", regex=False)
        g["item"] = g["item"].str.strip().replace({"": "(ítem sin nombre)", "nan": "(ítem sin nombre)"})
        return g if g["ejecutado"].sum() >= 1 else None

    g, gb = limpio("Herramientas"), limpio("Bolsas")
    if g is None:
        return f"Operaciones por ítem: no hay ejecución de herramientas por ítem de {ml} {anio} para {empresa}."
    tot_e, tot_p = g["ejecutado"].sum(), g["presupuesto"].sum()
    s = _clonar(prs, base, "Herramientas por ítem",
                f"{ml.capitalize()}: {_M(tot_e)} ejecutados frente a {_M(tot_p)} presupuestados ({_cmp_txt(tot_e, tot_p)})")
    var = g.set_index("item")["variacion"]
    var = var[var.abs() > 1e4]
    top_v = pd.concat([var.nlargest(5), var.nsmallest(5)])
    top_v = top_v[~top_v.index.duplicated()]
    _barras_var(s, 450000, Y0, 7600000, 4000000, sorted(top_v.items(), key=lambda q: -q[1]),
                titulo="Variación vs presupuesto por ítem (millones) · rojo: se pasó", tam=10, bueno_si_sube=False)
    t = g.sort_values("ejecutado", ascending=False).head(8)
    filas = [dict(celdas=[_corto(r.item, 30), _num(r.presupuesto / 1e6, dec=2), _num(r.ejecutado / 1e6, dec=2),
                          _num(r.variacion / 1e6, dec=2), _pct(r.pct, 0) if not pd.isna(r.pct) else "—"],
                  rojo={3} if r.variacion > 1e4 else (), verde={3} if r.variacion < -1e4 else ())
             for r in t.itertuples()]
    resto = g.drop(t.index)
    if len(resto):
        filas.append(dict(celdas=[f"Otros {len(resto)} ítems", _num(resto["presupuesto"].sum() / 1e6, dec=2),
                                  _num(resto["ejecutado"].sum() / 1e6, dec=2), _num(resto["variacion"].sum() / 1e6, dec=2), ""],
                          tipo="sub"))
    filas.append(dict(celdas=["Total", _num(tot_p / 1e6, dec=2), _num(tot_e / 1e6, dec=2), _num((tot_e - tot_p) / 1e6, dec=2),
                              _pct(_div(tot_e, tot_p), 0) if tot_p else "—"], tipo="total"))
    alto = min(330000, int(3900000 / (len(filas) + 1)))
    _tabla(s, 8250000, Y0, 9600000, ["ÍTEM", "PPTO M$", "EJEC. M$", "VAR. M$", "%"], filas, [3.4, 1.0, 1.0, 1.0, .7],
           alto_fila=alto, tam=10)
    if gb is not None:                                   # Bolsas: pocos ítems, van en tarjetas
        tb = gb.sort_values("ejecutado", ascending=False).head(4)
        _fila_tarjetas(s, Y0 + 4150000, [
            (f"BOLSAS · {_corto(r.item, 22).upper()}", _M(r.ejecutado),
             f"{_cmp_txt(r.ejecutado, r.presupuesto, 'el ppto')}" if r.presupuesto > 0 else "sin presupuesto",
             (r.ejecutado <= r.presupuesto) if r.presupuesto > 0 else False) for r in tb.itertuples()],
            x0=1500000, x1=ANCHO - 1500000, alto=950000, tam=18)
    sin_ppto = g[(g["presupuesto"] <= 0) & (g["ejecutado"] > 1)]
    sin_ejec = g[(g["presupuesto"] > 0) & (g["ejecutado"] <= 0)]
    mayor = g.loc[g["variacion"].idxmax()]
    txt = (f"En herramientas, el ítem que más se pasó del presupuesto es {mayor['item'].lower()} ({_M(mayor['variacion'], True)}). "
           if mayor["variacion"] > 0 else f"En herramientas, ningún ítem superó su presupuesto en {ml}. ")
    if len(sin_ppto):
        txt += (f"{len(sin_ppto)} ítems se ejecutaron sin presupuesto en el mes ({_M(sin_ppto['ejecutado'].sum())}), entre ellos "
                + ", ".join(_corto(i, 34).lower() for i in sin_ppto.sort_values("ejecutado", ascending=False)["item"].head(3)) + ". ")
    if len(sin_ejec):
        txt += f"{len(sin_ejec)} ítems presupuestados no tuvieron consumo ({_M(sin_ejec['presupuesto'].sum())} sin ejecutar). "
    if gb is not None:
        txt += (f"Bolsas cerró en {_M(gb['ejecutado'].sum())} ({_cmp_txt(gb['ejecutado'].sum(), gb['presupuesto'].sum())}).")
    _texto(s, 1300000, Y0 + 5300000, 13500000, 1500000, [txt.strip()], tam=15)
    R["dest"].append(("Herramientas por ítem.", txt.split(". ")[0] + "."))
    return None


# ------------------------------------------------------------------ armado
_ARMAR = {"Presupuesto": _diap_presupuesto, "Nómina": _diap_nomina, "Flujo de caja": _diap_flujo,
          "Combustible": _diap_combustible, "Facturación de aseo": _diap_facturacion,
          "Corte de césped": _diap_corte, "Consumo operativo": _diap_operativo, "Operaciones por ítem": _diap_operativo_items,
          "Mantenimiento de flota": _diap_flota, "Flota por placa": _diap_flota_placas, "Balance e indicadores": _diap_balance,
          "Estado de resultados": _diap_pyg}


def generar(leer, empresa: str, anio: int, mes: int, version: str = P.DEFECTO, secciones=None, ctx=None) -> tuple[bytes, list[str]]:
    """Presentación del mes para una compañía. Devuelve el archivo .pptx y la lista de secciones que no se pudieron
    armar (sin datos de ese mes). `ctx` (el contexto de la aplicación) lo piden las diapositivas de Operaciones y Flota,
    que usan los mismos cálculos de sus pantallas."""
    secciones = list(secciones or SECCIONES)
    anio, mes = int(anio), int(mes)
    prs = Presentation(str(PLANTILLA))
    base = prs.slides[0]
    R = {"version": version, "kpi": {}, "dest": [], "ctx": ctx, "notas": []}
    avisos = []
    con_resumen = "Resumen del mes" in secciones
    borrador = Presentation(str(PLANTILLA)) if con_resumen else None   # temas no elegidos: solo para el resumen
    for sec, armar in _ARMAR.items():
        elegido = sec in secciones
        if not elegido and not con_resumen:
            continue
        destino = prs if elegido else borrador
        try:
            aviso = armar(destino, destino.slides[0], leer, empresa, anio, mes, R)
        except Exception as e:                          # una sección con problemas no tumba las demás
            aviso = f"{sec}: no se pudo armar ({type(e).__name__}: {e})."
        if aviso and elegido:
            avisos.append(aviso)
    ids = prs.slides._sldIdLst
    if con_resumen:
        aviso = _diap_resumen(prs, base, empresa, anio, mes, R)
        if aviso:
            avisos.append(aviso)
        else:                                           # el resumen va de primero
            el = list(ids)[-1]
            ids.remove(el)
            ids.insert(1, el)
    if "Portada" in secciones and len(ids) > 1:         # la portada va de primera, antes del resumen
        _diap_portada(prs, base, empresa, anio, mes)
        el = list(ids)[-1]
        ids.remove(el)
        ids.insert(1, el)
    # quitar la diapositiva base de la plantilla
    primero = list(ids)[0]
    prs.part.drop_rel(primero.rId)
    ids.remove(primero)
    avisos += [n for n in R["notas"] if n not in avisos]        # presupuestos que no son exactamente el elegido
    if not len(ids):
        return b"", avisos
    buf = BytesIO()
    prs.save(buf)
    return buf.getvalue(), avisos


# diapositivas agrupadas por tema, para elegirlas en pantalla (el orden final del archivo lo fija `_ARMAR`)
GRUPOS = {
    "Inicio": ["Portada", "Resumen del mes"],
    "Finanzas": ["Presupuesto", "Flujo de caja", "Balance e indicadores", "Estado de resultados"],
    "Personal y operación": ["Nómina", "Combustible", "Facturación de aseo", "Corte de césped", "Consumo operativo", "Operaciones por ítem"],
    "Mantenimiento": ["Mantenimiento de flota", "Flota por placa"],
}


def disponibilidad(leer, empresa: str, anio: int, mes: int, version: str) -> dict:
    """Qué diapositivas tienen datos para ese mes: {nombre: (hay datos, motivo si no)}. Son consultas rápidas de existencia; no arman nada.
    Una consulta que falla (tabla que aún no existe) cuenta como «sin datos»."""
    def hay(sql, par) -> bool:
        try:
            return bool(len(leer(sql + " LIMIT 1", par)))
        except Exception:
            return False
    e, a, m = empresa, int(anio), int(mes)
    con_ppto = hay("SELECT 1 FROM presupuesto WHERE anio=? AND version=? AND empresa=?", (a, version, e))
    con_ejec = hay("SELECT 1 FROM ejecucion WHERE anio=? AND mes=? AND empresa=? AND codigo<>''", (a, m, e))
    res = {"Portada": (True, ""), "Resumen del mes": (True, "")}
    res["Presupuesto"] = (con_ppto and con_ejec, "" if con_ppto and con_ejec else
                          (f"no hay {P.etiqueta(version)} cargado" if not con_ppto else "no hay ejecución del mes"))
    res["Nómina"] = (hay("SELECT 1 FROM nomina_ejec WHERE anio=? AND mes=? AND COALESCE(empresa,'ALDC')=?", (a, m, e)), "no hay nómina del mes")
    res["Flujo de caja"] = (hay("SELECT 1 FROM flujo_mov WHERE anio=? AND mes=? AND COALESCE(empresa,'ALDC')=?", (a, m, e)), "no hay movimientos de caja del mes")
    res["Combustible"] = (hay("SELECT 1 FROM combustible WHERE anio=? AND mes=? AND empresa=?", (a, m, e)), "no hay tanqueos del mes")
    res["Facturación de aseo"] = (hay("SELECT 1 FROM facturacion_mes WHERE empresa=? AND anio=? AND mes=?", (e, a, m)) or (e, a) in FACT_HISTORICO,
                                  "no hay facturación del mes")
    res["Corte de césped"] = (e == "ALDC" and hay("SELECT 1 FROM cp_ejec WHERE anio=? AND mes=?", (a, m)),
                              "solo ALDC" if e != "ALDC" else "no hay ejecución del mes")
    sispo = hay("SELECT 1 FROM oper_sispo_periodo WHERE empresa=? AND periodo=?", (e, a * 100 + m))
    res["Consumo operativo"] = (sispo, "no hay SISPO del mes")
    res["Operaciones por ítem"] = (sispo, "no hay SISPO del mes")
    res["Mantenimiento de flota"] = (con_ejec, "no hay ejecución del mes")
    res["Flota por placa"] = (con_ejec, "no hay ejecución del mes")
    res["Balance e indicadores"] = (hay("SELECT 1 FROM eeff WHERE empresa=? AND anio=? AND mes=?", (e, a, m)), "no hay estados financieros del mes")
    res["Estado de resultados"] = res["Balance e indicadores"]
    return res


def nombre_archivo(empresa: str, anio: int, mes: int) -> str:
    return f"{mes:02d}_{empresa}_PRESENTACION_FINANCIERA_{anio}.pptx"
