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

PLANTILLA = Path(__file__).with_name("plantilla_presentacion.pptx")
SECCIONES = ["Resumen del mes", "Presupuesto", "Nómina", "Flujo de caja", "Combustible", "Balance e indicadores",
             "Estado de resultados"]
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


def _cap(t: str) -> str:
    return t[:1].upper() + t[1:] if t else t


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
             colores=(AZUL, NARANJA, NAVY)):
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
    ppto, ejec, pers = mod_nomina.preparar(leer, anio, empresa)
    if ejec.empty or not (ejec["mes"] == mes).any():
        return f"Nómina: no hay nómina ejecutada de {MESES_L[mes - 1]} {anio} para {empresa}."
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
                f"{EMPRESAS.get(empresa, empresa)} · cifras principales y lo más importante del mes")
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


# ------------------------------------------------------------------ armado
_ARMAR = {"Presupuesto": _diap_presupuesto, "Nómina": _diap_nomina, "Flujo de caja": _diap_flujo,
          "Combustible": _diap_combustible, "Balance e indicadores": _diap_balance,
          "Estado de resultados": _diap_pyg}


def generar(leer, empresa: str, anio: int, mes: int, version: str = "8+4", secciones=None) -> tuple[bytes, list[str]]:
    """Presentación del mes para una compañía. Devuelve el archivo .pptx y la lista de secciones que no se pudieron
    armar (sin datos de ese mes)."""
    secciones = list(secciones or SECCIONES)
    anio, mes = int(anio), int(mes)
    prs = Presentation(str(PLANTILLA))
    base = prs.slides[0]
    R = {"version": version, "kpi": {}, "dest": []}
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
    # quitar la diapositiva base de la plantilla
    primero = list(ids)[0]
    prs.part.drop_rel(primero.rId)
    ids.remove(primero)
    if not len(ids):
        return b"", avisos
    buf = BytesIO()
    prs.save(buf)
    return buf.getvalue(), avisos


def nombre_archivo(empresa: str, anio: int, mes: int) -> str:
    return f"{mes:02d}_{empresa}_PRESENTACION_FINANCIERA_{anio}.pptx"
