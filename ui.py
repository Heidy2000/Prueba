"""Constantes y ayudas visuales compartidas por todos los módulos."""
import pandas as pd
import streamlit as st

import estilo as E

MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
MESES_L = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
           "septiembre", "octubre", "noviembre", "diciembre"]
ADMIN = "Financiera"
SIN_DUENO = "Sin asignar"
# compañías del grupo (sigla -> nombre). La primera es la que se muestra por defecto.
EMPRESAS = {"ALDC": "Área Limpia D.C.", "ALSM": "Área Limpia Servicios Medioambientales"}
EMPRESA_DEFECTO = "ALDC"


def empresa_de_nombre(nombre: str) -> str:
    """Compañía según el nombre del archivo (ALSM / 'servicios medio ambientales'); si no dice, ALDC."""
    n = " ".join(str(nombre).upper().replace("_", " ").replace("-", " ").split())
    return "ALSM" if "ALSM" in n or "MEDIO AMBIENT" in n or "MEDIOAMBIENT" in n else EMPRESA_DEFECTO


def selector_empresa(contenedor, disponibles, key: str, label: str = "Compañía"):
    """Selector de compañía; si solo hay una con datos, no se muestra nada y se devuelve esa."""
    disponibles = [e for e in EMPRESAS if e in set(disponibles)] or [EMPRESA_DEFECTO]
    if len(disponibles) == 1:
        return disponibles[0]
    return contenedor.selectbox(label, disponibles, key=key, format_func=lambda e: f"{e} · {EMPRESAS[e]}")


def ahora():
    """Fecha y hora de Colombia (en la nube el servidor está en otra zona horaria)."""
    from datetime import datetime, timedelta, timezone
    try:
        from zoneinfo import ZoneInfo
        return datetime.now(ZoneInfo("America/Bogota")).replace(tzinfo=None)
    except Exception:                      # Windows sin base de zonas horarias: Colombia es UTC-5 todo el año
        return datetime.now(timezone(timedelta(hours=-5))).replace(tzinfo=None)


def estilo_grafico(fig, alto=320):
    fig.update_layout(height=alto, margin=dict(t=10, b=10, l=10, r=10),
                      paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                      font=dict(family=E.FUENTE, color=E.TEXTO, size=12),
                      legend=dict(orientation="h", y=-0.18, x=0, font=dict(size=11)),
                      hoverlabel=dict(font_family=E.FUENTE), separators=",.")
    fig.update_xaxes(showgrid=False, linecolor=E.BORDE)
    fig.update_yaxes(gridcolor="#eef1f6", zeroline=False)
    return fig


def titulo_tarjeta(t, s=""):
    st.html(f'<p class="nf-card-t">{t}</p>' + (f'<p class="nf-card-s">{s}</p>' if s else ""))


class Periodo:
    """Selección de periodo común: acumulado hasta un mes o solo ese mes."""

    def __init__(self, modo: str, mes: int):
        self.solo = modo == "Solo el mes"
        self.mes = mes
        self.mes_txt = MESES_L[mes - 1]
        self.suf = f"de {self.mes_txt}" if self.solo else "al corte"
        self.texto = f"solo {self.mes_txt}" if self.solo else f"enero a {self.mes_txt}"

    def filtrar(self, df: pd.DataFrame) -> pd.DataFrame:
        return df[df["mes"] == self.mes] if self.solo else df[df["mes"] <= self.mes]

    def meses(self) -> list[int]:
        return [self.mes] if self.solo else list(range(1, self.mes + 1))


def selector_periodo(col_modo, col_mes, ultimo_mes: int, clave: str) -> Periodo:
    modo = col_modo.segmented_control("Periodo", ["Acumulado", "Solo el mes"], default="Acumulado",
                                      key=f"{clave}_modo") or "Acumulado"
    mes = col_mes.selectbox("Mes" if modo == "Solo el mes" else "Hasta", list(range(1, 13)),
                            index=ultimo_mes - 1, format_func=lambda m: MESES_L[m - 1].capitalize(),
                            key=f"{clave}_mes")
    return Periodo(modo, mes)


def a_excel(df: pd.DataFrame) -> bytes:
    from io import BytesIO
    buf = BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


def _partir(texto: str, ancho: int = 12) -> str:
    """Parte una etiqueta larga en dos o tres renglones (eje x de los gráficos)."""
    lineas, actual = [], ""
    for w in str(texto).split():
        if actual and len(actual) + 1 + len(w) > ancho:
            lineas.append(actual)
            actual = w
        else:
            actual = f"{actual} {w}".strip()
    lineas.append(actual)
    return "<br>".join(lineas[:3])


def cascada(pasos: list[dict], alto: int = 300):
    """
    Gráfico de cascada (puente) en millones: del presupuesto a lo ejecutado.
    pasos: dict(nombre, valor (pesos), tipo='total'|'delta', color). Los totales parten de cero; los deltas suben
    o bajan desde el nivel anterior. El eje arranca cerca del menor nivel para que se vean las diferencias.
    """
    import plotly.graph_objects as go
    x, base, alto_b, colores, textos, niveles = [], [], [], [], [], []
    nivel = 0.0
    for p in pasos:
        v = p["valor"] / 1e6
        nom = str(p["nombre"])
        if nom.isupper():                     # en mayúsculas ocupa más ancho
            nom = nom.capitalize()
        etq = _partir(nom if len(nom) <= 26 else nom[:24].rstrip() + "…")
        while etq in x:                       # dos etiquetas iguales se juntarían en una sola barra
            etq += "\u200b"
        x.append(etq)
        if p.get("tipo", "delta") == "total":
            base.append(0.0)
            alto_b.append(v)
            nivel = v
            textos.append(f"<b>{v:,.0f}</b>".replace(",", "."))
        else:
            base.append(min(nivel, nivel + v))
            alto_b.append(abs(v))
            nivel += v
            textos.append(f"<b>{'+' if v >= 0 else '−'}{abs(v):,.0f}</b>".replace(",", "."))
        niveles.append(nivel)
        colores.append(p.get("color", E.AZUL))
    tops = [b + h for b, h in zip(base, alto_b)]
    piso_datos = min([b for b, p in zip(base, pasos) if p.get("tipo") != "total"] + niveles)
    techo = max(tops) or 1
    piso = max(0.0, piso_datos - (techo - piso_datos) * 1.2)
    piso = 0.0 if piso < techo * 0.25 else piso
    fig = go.Figure()
    base_vis = [max(b, piso) for b in base]
    fig.add_bar(x=x, y=[t - bv for t, bv in zip(tops, base_vis)], base=base_vis, marker_color=colores,
                marker_line_width=0, text=textos, textposition="outside", cliponaxis=False,
                textfont=dict(color=E.TEXTO, size=12), showlegend=False, constraintext="none",
                customdata=[a for a in alto_b], hovertemplate="%{x}: %{customdata:,.1f} M<extra></extra>")
    for i in range(len(x) - 1):          # conectores punteados entre barras
        fig.add_shape(type="line", x0=i + 0.3, x1=i + 0.7, y0=niveles[i], y1=niveles[i], xref="x", yref="y",
                      line=dict(color="#9aa7b8", width=1, dash="dot"))
    fig.update_layout(bargap=.38,
                      yaxis=dict(range=[piso, techo * 1.12 if piso == 0 else techo + (techo - piso) * .15],
                                 tickformat=",.0f", ticksuffix=" M", showgrid=True))
    fig = estilo_grafico(fig, alto)
    fig.update_layout(margin=dict(t=24, b=10, l=10, r=10))
    fig.update_xaxes(tickangle=0, tickfont=dict(size=11))
    return fig, piso


def dona_doble(p1: float, p2: float | None, t1: str, t2: str, alto: int = 210):
    """Dona de cumplimiento: anillo exterior p1 (p. ej. vs presupuesto) e interior p2 (vs otra base)."""
    import plotly.graph_objects as go

    def anillo(p, color, dom, hueco):
        p = 0.0 if p is None or pd.isna(p) else float(p)
        lleno = max(0.0, min(p, 1.0))
        return go.Pie(values=[lleno, 1 - lleno], hole=hueco, sort=False, direction="clockwise", rotation=0,
                      marker_colors=[color, "#e9eef5"], textinfo="none", hoverinfo="skip", domain=dom)
    fig = go.Figure(anillo(p1, E.AZUL, dict(x=[0, 1], y=[0, 1]), .82))
    if p2 is not None and not pd.isna(p2):
        fig.add_trace(anillo(p2, "#7fa8e6", dict(x=[.13, .87], y=[.13, .87]), .78))
    centro = "—" if p1 is None or pd.isna(p1) else f"{p1 * 100:.0f}%"
    fig.add_annotation(text=f"<b style='font-size:28px'>{centro}</b><br><span style='font-size:11px'>{t1}</span>",
                       showarrow=False, font=dict(family=E.FUENTE, color=E.TEXTO))
    fig.update_layout(showlegend=False)
    return estilo_grafico(fig, alto)
