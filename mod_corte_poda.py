"""Módulo Corte y poda: ejecución de área cortada (m²) y de árboles podados contra el presupuesto, valorada con las
dos tarifas del presupuesto (MTA y NMT) lado a lado. El presupuesto sale del libro «SERVICIOS PROYECTADOS CORTE Y PODA»
y la ejecución de un archivo mensual; ambos se cargan en Administración → Corte y poda."""
import io
import re
import unicodedata

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import estilo as E
from ui import MESES_L, a_excel, estilo_grafico, selector_periodo, titulo_tarjeta

TABLAS = """
CREATE TABLE IF NOT EXISTS cp_ppto(anio INT, mes INT, m2 REAL, tarifa_corte_mta REAL, tarifa_corte_nmt REAL,
    valor_corte_mta REAL, valor_corte_nmt REAL, arboles REAL, valor_poda_mta REAL, valor_poda_nmt REAL, archivo TEXT);
CREATE TABLE IF NOT EXISTS cp_ejec(anio INT, mes INT, m2 REAL, arboles REAL, archivo TEXT)
"""
COLOR_PPTO, COLOR_EJEC = "#9aa7b8", E.AZUL


def _norm(x) -> str:
    t = unicodedata.normalize("NFKD", str(x if x is not None else "")).encode("ascii", "ignore").decode().upper()
    return re.sub(r"[^A-Z0-9]", "", t)


_MES_NUM = {_norm(m): i + 1 for i, m in enumerate(MESES_L)}
_MES_NUM.update({k[:3]: v for k, v in list(_MES_NUM.items())})          # ENE, FEB… también valen
_MES_NUM["SEPT"] = 9


def _mes(x) -> int | None:
    if isinstance(x, (int, float)) and not pd.isna(x) and 1 <= int(x) <= 12:
        return int(x)
    if hasattr(x, "month"):
        return int(x.month)
    return _MES_NUM.get(_norm(x))


def _num(x):
    return float(x) if isinstance(x, (int, float)) and not isinstance(x, bool) and not pd.isna(x) else None


def leer_presupuesto(archivo, anio_inicio: int) -> tuple[pd.DataFrame, list[str]]:
    """Presupuesto del libro de corte y poda: una fila por mes (año según el orden: sube al pasar de diciembre a enero).
    Columnas desde «MES»: m², $/m² MTA, total MTA, $/m² NMT, total NMT, árboles, poda MTA, poda NMT."""
    from openpyxl import load_workbook
    wb = load_workbook(archivo, data_only=True, read_only=True)
    ws = next((w for w in wb.worksheets if "PODA" in _norm(w.title)), wb.worksheets[0])
    filas = [list(r) for r in ws.iter_rows(values_only=True)]
    cab = next(((i, j) for i, r in enumerate(filas) for j, v in enumerate(r) if _norm(v) == "MES"), None)
    if cab is None:
        raise ValueError("No encontré la columna «MES» en el libro.")
    i0, j = cab
    avisos, out, anio, previo = [], [], anio_inicio, None
    for r in filas[i0 + 1:]:
        m = _mes(r[j]) if j < len(r) else None
        if m is None or not isinstance(r[j], str):
            if out:
                break
            continue
        if previo is not None and m < previo:
            anio += 1
        previo = m
        g = lambda k: _num(r[j + k]) if j + k < len(r) else None
        out.append(dict(anio=anio, mes=m, m2=g(1), tarifa_corte_mta=g(2), valor_corte_mta=g(3), tarifa_corte_nmt=g(4),
                        valor_corte_nmt=g(5), arboles=g(6), valor_poda_mta=g(7), valor_poda_nmt=g(8)))
    d = pd.DataFrame(out)
    if d.empty:
        raise ValueError("No encontré meses debajo de la columna «MES».")
    for esc in ("mta", "nmt"):                                           # sin tarifa no hay valor: no se inventa
        sin = d[f"tarifa_corte_{esc}"].isna()
        if sin.any():
            d.loc[sin, f"valor_corte_{esc}"] = None
            avisos.append(f"{esc.upper()}: sin tarifa de corte en {int(sin.sum())} mes(es) (el total del libro se ignora).")
        ok = d["m2"].notna() & d[f"tarifa_corte_{esc}"].notna() & d[f"valor_corte_{esc}"].notna()
        dif = (d["m2"] * d[f"tarifa_corte_{esc}"] - d[f"valor_corte_{esc}"]).abs()
        mal = ok & (dif > d[f"valor_corte_{esc}"].abs() * 0.001)
        if mal.any():
            avisos.append(f"{esc.upper()}: m² × tarifa no coincide con el total del libro en {int(mal.sum())} mes(es).")
    sin_arb = d["arboles"].isna() & d["valor_poda_mta"].notna()
    if sin_arb.any():
        avisos.append(f"Poda: {int(sin_arb.sum())} mes(es) tienen valor pero no cantidad de árboles "
                      f"({', '.join(f'{MESES_L[m - 1][:3]} {a}' for a, m in d.loc[sin_arb, ['anio', 'mes']].itertuples(index=False))}): "
                      "ahí la poda se compara solo en árboles ejecutados, sin valorizar.")
    return d, avisos


def leer_ejecucion(archivo, anio_defecto: int) -> tuple[pd.DataFrame, list[str]]:
    """Ejecución: columnas reconocidas por nombre (Año, Mes o Fecha, m², Árboles). Varias filas del mismo mes se suman."""
    d = pd.read_excel(archivo, dtype=object)
    claves = {_norm(c): c for c in d.columns}
    def col(*prefijos):
        return next((c for k, c in claves.items() if any(k.startswith(p) for p in prefijos)), None)
    c_f, c_m, c_a = col("FECHA"), col("MES"), col("ANO", "ANIO")
    c_m2 = col("M2", "AREA", "METROS")
    c_ar = col("ARBOL", "PODA")
    if (c_f is None and c_m is None) or (c_m2 is None and c_ar is None):
        raise ValueError("El archivo necesita una columna de mes (o fecha) y al menos una de m² / árboles.")
    avisos = []
    if c_f is not None:
        f = pd.to_datetime(d[c_f], errors="coerce")
        anio, mes = f.dt.year, f.dt.month
    else:
        mes = d[c_m].map(_mes)
        anio = pd.to_numeric(d[c_a], errors="coerce") if c_a else pd.Series(anio_defecto, index=d.index)
    out = pd.DataFrame({"anio": anio, "mes": mes,
                        "m2": pd.to_numeric(d[c_m2], errors="coerce") if c_m2 else float("nan"),
                        "arboles": pd.to_numeric(d[c_ar], errors="coerce") if c_ar else float("nan")})
    malas = out["anio"].isna() | out["mes"].isna()
    if malas.any():
        avisos.append(f"{int(malas.sum())} fila(s) sin mes/año válido no se cargan.")
    out = out[~malas].astype({"anio": int, "mes": int})
    out = out.groupby(["anio", "mes"], as_index=False).agg(m2=("m2", lambda s: s.sum(min_count=1)),
                                                           arboles=("arboles", lambda s: s.sum(min_count=1)))
    return out, avisos


def plantilla_ejecucion() -> bytes:
    return a_excel(pd.DataFrame({"Año": [2026, 2026], "Mes": ["SEPTIEMBRE", "OCTUBRE"],
                                 "m² ejecutados": [None, None], "Árboles podados": [None, None]}))


# ------------------------------------------------------------------ cálculo
def tabla_mensual(ppto: pd.DataFrame, ejec: pd.DataFrame) -> pd.DataFrame:
    """Una fila por mes: presupuesto, ejecución (vacía si no hay dato: no es cero) y ejecución valorizada con cada tarifa."""
    ppto = ppto.apply(pd.to_numeric, errors="coerce")
    t = ppto.merge(ejec.rename(columns={"m2": "m2_ej", "arboles": "arboles_ej"})[["anio", "mes", "m2_ej", "arboles_ej"]],
                   on=["anio", "mes"], how="left")
    for esc in ("mta", "nmt"):
        t[f"corte_ej_{esc}"] = t["m2_ej"] * t[f"tarifa_corte_{esc}"]
        por_arbol = t[f"valor_poda_{esc}"] / t["arboles"].where(t["arboles"] > 0)
        t[f"poda_ej_{esc}"] = t["arboles_ej"] * por_arbol
    return t.sort_values(["anio", "mes"]).reset_index(drop=True)


def _pct(a, b):
    return a / b if b and not pd.isna(b) and not pd.isna(a) else float("nan")


def _fmt_pct(v):
    return "—" if pd.isna(v) else f"{v * 100:,.1f}%".replace(",", "X").replace(".", ",").replace("X", ".")


def _n(v, dec=0):
    return "—" if pd.isna(v) else f"{v:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def resumen(t: pd.DataFrame) -> dict:
    """Totales de los meses CON ejecución cargada: así presupuesto y ejecución cubren el mismo periodo."""
    m2 = t[t["m2_ej"].notna()]
    ar = t[t["arboles_ej"].notna()]
    r = dict(m2_p=m2["m2"].sum(), m2_e=m2["m2_ej"].sum(), ar_p=ar["arboles"].sum(), ar_e=ar["arboles_ej"].sum())
    for esc in ("mta", "nmt"):
        r[f"corte_p_{esc}"] = m2[f"valor_corte_{esc}"].sum(min_count=1)
        r[f"corte_e_{esc}"] = m2[f"corte_ej_{esc}"].sum(min_count=1)
        v = ar[ar["arboles"] > 0]
        r[f"poda_p_{esc}"] = v[f"valor_poda_{esc}"].sum(min_count=1)
        r[f"poda_e_{esc}"] = v[f"poda_ej_{esc}"].sum(min_count=1)
    return r


# ------------------------------------------------------------------ administración
def admin_cargas(ctx):
    con, leer = ctx.con, ctx.leer
    with st.container(key="card_adm_cp"):
        titulo_tarjeta("Corte y poda", "Presupuesto (libro de servicios proyectados) y ejecución mensual de m² y árboles.")
        anio0 = st.number_input("Año del primer mes del libro", 2024, 2035, 2026, key="cp_anio0",
                                help="El año sube solo al pasar de diciembre a enero.")
        a, b = st.tabs(["Presupuesto", "Ejecución"])
        with a:
            f = st.file_uploader("Libro de presupuesto", type=["xlsx", "xlsm"], key="cp_f_ppto", label_visibility="collapsed")
            if f is not None:
                try:
                    d, avisos = leer_presupuesto(f, int(anio0))
                except Exception as err:
                    st.error(f"No pude leer el libro: {err}")
                else:
                    for x in avisos:
                        st.warning(x)
                    st.dataframe(d, hide_index=True, width="stretch")
                    if st.button("Guardar presupuesto", type="primary", key="cp_g_ppto"):
                        for anio in d["anio"].unique():
                            con.execute("DELETE FROM cp_ppto WHERE anio=?", (int(anio),))
                        d2 = d.assign(archivo=f.name)
                        con.anexar(d2, "cp_ppto"); con.commit()
                        ctx.registrar_carga(ctx.usuario, "Corte y poda · presupuesto", f.name, f"{len(d)} meses")
                        st.success("Presupuesto guardado.")
        with b:
            st.download_button("Descargar plantilla", plantilla_ejecucion(), "plantilla_ejecucion_corte_poda.xlsx",
                               icon=":material/download:", key="cp_plantilla")
            f = st.file_uploader("Archivo de ejecución", type=["xlsx", "xlsm"], key="cp_f_ejec", label_visibility="collapsed")
            if f is not None:
                try:
                    d, avisos = leer_ejecucion(f, int(anio0))
                except Exception as err:
                    st.error(f"No pude leer el archivo: {err}")
                else:
                    for x in avisos:
                        st.warning(x)
                    st.dataframe(d, hide_index=True, width="stretch")
                    if st.button("Guardar ejecución", type="primary", key="cp_g_ejec"):
                        for r in d.itertuples(index=False):
                            con.execute("DELETE FROM cp_ejec WHERE anio=? AND mes=?", (int(r.anio), int(r.mes)))
                        con.anexar(d.assign(archivo=f.name), "cp_ejec"); con.commit()
                        ctx.registrar_carga(ctx.usuario, "Corte y poda · ejecución", f.name, f"{len(d)} meses")
                        st.success("Ejecución guardada.")


# ------------------------------------------------------------------ pestaña
def mostrar(ctx):
    leer = ctx.leer
    ppto, ejec = leer("SELECT * FROM cp_ppto"), leer("SELECT * FROM cp_ejec")
    zona = st.container()
    if ppto.empty:
        zona.html(E.encabezado("Corte y poda", "Ejecución de área y de árboles contra el presupuesto", []))
        st.info("Aún no hay presupuesto de corte y poda. Se carga en Administración → Corte y poda.")
        return
    with st.container(key="card_filtros"):
        st.html('<div class="nf-ftit">Filtros del informe</div>')
        c = st.columns([0.8, 1.6, 0.9], vertical_alignment="bottom")
        anios = sorted(ppto["anio"].unique())
        anio = c[0].selectbox("Año", anios, index=len(anios) - 1 if not ejec.empty and ejec["anio"].max() >= anios[-1] else 0, key="cp_anio")
        ult = int(ejec.loc[ejec["anio"] == anio, "mes"].max()) if (ejec["anio"] == anio).any() else int(ppto.loc[ppto["anio"] == anio, "mes"].min())
        per = selector_periodo(c[1], c[2], ult, "cp")
    t = tabla_mensual(ppto[ppto["anio"] == anio], ejec[ejec["anio"] == anio])
    t = per.filtrar(t)
    chips = [(f"{'Mes' if per.solo else 'Corte'}: {MESES_L[per.mes - 1].capitalize()} {anio}", "")]
    chips.append((f"Ejecución cargada hasta {MESES_L[ult - 1].capitalize()}", "ok") if (ejec["anio"] == anio).any()
                 else ("Sin ejecución cargada", "warn"))
    zona.html(E.encabezado("Corte y poda", f"Área cortada y árboles podados · {per.texto}", chips))
    if t["m2_ej"].isna().all() and t["arboles_ej"].isna().all():
        st.info("No hay ejecución cargada para este periodo.")
        return
    r = resumen(t)
    st.html(E.kpis([
        dict(label="Área cortada (m²)", valor=_n(r["m2_e"]), sub=f"{_fmt_pct(_pct(r['m2_e'], r['m2_p']))} de {_n(r['m2_p'])} m²",
             icono="receipt", color=E.AZUL, fondo=E.AZUL_SUAVE),
        dict(label="Árboles podados", valor=_n(r["ar_e"]), sub=(f"{_fmt_pct(_pct(r['ar_e'], r['ar_p']))} de {_n(r['ar_p'])}" if r["ar_p"] else "sin meta de árboles"),
             icono="calendar", color=E.AZUL, fondo=E.AZUL_SUAVE),
        dict(label="Corte valorizado MTA", valor=E.millones(r["corte_e_mta"]), sub=f"{_fmt_pct(_pct(r['corte_e_mta'], r['corte_p_mta']))} de {E.millones(r['corte_p_mta'])}",
             icono="receipt", color=E.AZUL, fondo=E.AZUL_SUAVE),
        dict(label="Corte valorizado NMT", valor=E.millones(r["corte_e_nmt"]), sub=f"{_fmt_pct(_pct(r['corte_e_nmt'], r['corte_p_nmt']))} de {E.millones(r['corte_p_nmt'])}",
             icono="receipt", color=E.AZUL, fondo=E.AZUL_SUAVE)]))

    mes_txt = [MESES_L[m - 1].capitalize() for m in t["mes"]]
    g1, g2 = st.columns(2)
    for col, titulo, p, e in ((g1, "Área (m²) por mes", "m2", "m2_ej"), (g2, "Árboles por mes", "arboles", "arboles_ej")):
        with col:
            titulo_tarjeta(titulo)
            fig = go.Figure([go.Bar(x=mes_txt, y=t[p], name="Presupuesto", marker_color=COLOR_PPTO),
                             go.Bar(x=mes_txt, y=t[e], name="Ejecutado", marker_color=COLOR_EJEC)])
            fig.update_layout(barmode="group")
            st.plotly_chart(estilo_grafico(fig, 280), width="stretch")

    corte = pd.DataFrame({"Mes": mes_txt, "m² presup.": t["m2"], "m² ejec.": t["m2_ej"],
                          "% m²": [_pct(a, b) for a, b in zip(t["m2_ej"], t["m2"])],
                          "$ presup. MTA": t["valor_corte_mta"], "$ ejec. MTA": t["corte_ej_mta"],
                          "$ presup. NMT": t["valor_corte_nmt"], "$ ejec. NMT": t["corte_ej_nmt"]})
    poda = pd.DataFrame({"Mes": mes_txt, "Árboles presup.": t["arboles"], "Árboles ejec.": t["arboles_ej"],
                         "% árboles": [_pct(a, b) for a, b in zip(t["arboles_ej"], t["arboles"])],
                         "$ presup. MTA": t["valor_poda_mta"], "$ ejec. MTA": t["poda_ej_mta"],
                         "$ presup. NMT": t["valor_poda_nmt"], "$ ejec. NMT": t["poda_ej_nmt"]})
    cfg_cols = {c: st.column_config.NumberColumn(format="%.0f") for c in corte.columns[1:] if c not in ("% m²",)}
    cfg_cols["% m²"] = st.column_config.NumberColumn(format="percent")
    with st.container(key="card_cp_corte"):
        titulo_tarjeta("Corte de césped (área)", "Ejecución valorizada = m² ejecutados × tarifa del mes. Mes sin ejecución cargada queda vacío.")
        st.dataframe(corte, hide_index=True, width="stretch", column_config=cfg_cols)
    cfg_p = {c: st.column_config.NumberColumn(format="%.0f") for c in poda.columns[1:] if c != "% árboles"}
    cfg_p["% árboles"] = st.column_config.NumberColumn(format="percent")
    with st.container(key="card_cp_poda"):
        titulo_tarjeta("Poda de árboles", "Valor por árbol = valor mensual de poda ÷ árboles del presupuesto; sin árboles presupuestados no se valoriza.")
        st.dataframe(poda, hide_index=True, width="stretch", column_config=cfg_p)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf) as w:
        corte.to_excel(w, "Corte", index=False)
        poda.to_excel(w, "Poda", index=False)
    st.download_button("Descargar Excel", buf.getvalue(), f"corte_poda_{anio}_{per.mes:02d}.xlsx", icon=":material/download:", key="cp_dl")
