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
from ui import MESES_L, ahora, estilo_grafico, selector_periodo, titulo_tarjeta

TABLAS = """
CREATE TABLE IF NOT EXISTS cp_ppto(anio INT, mes INT, m2 REAL, tarifa_corte_mta REAL, tarifa_corte_nmt REAL,
    valor_corte_mta REAL, valor_corte_nmt REAL, arboles REAL, valor_poda_mta REAL, valor_poda_nmt REAL, archivo TEXT);
CREATE TABLE IF NOT EXISTS cp_ejec(anio INT, mes INT, m2 REAL, tarifa REAL, valor_corte REAL, arboles REAL, valor_poda REAL, actualizado TEXT)
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
    Formato nuevo, columnas desde «MES»: m², $/m² MTA, $/m² NMT, árboles, poda MTA, poda NMT (el total de corte = m² × tarifa
    lo calcula el aplicativo). Formato antiguo: m², $/m² MTA, total MTA, $/m² NMT, total NMT, árboles, poda MTA, poda NMT."""
    from openpyxl import load_workbook
    wb = load_workbook(archivo, data_only=True, read_only=True)
    ws = next((w for w in wb.worksheets if "PODA" in _norm(w.title)), wb.worksheets[0])
    filas = [list(r) for r in ws.iter_rows(values_only=True)]
    cab = next(((i, j) for i, r in enumerate(filas) for j, v in enumerate(r) if _norm(v) == "MES"), None)
    if cab is None:
        raise ValueError("No encontré la columna «MES» en el libro.")
    i0, j = cab
    avisos, out, anio, previo = [], [], anio_inicio, None
    cuerpo = []                                                          # filas de meses (texto de mes en la columna MES)
    for r in filas[i0 + 1:]:
        m = _mes(r[j]) if j < len(r) else None
        if m is None or not isinstance(r[j], str):
            if cuerpo:
                break
            continue
        cuerpo.append((m, r))
    g0 = lambda r, k: _num(r[j + k]) if j + k < len(r) else None
    # Formato del libro: el antiguo trae el total de corte (m² × tarifa) junto a cada tarifa; el nuevo solo m² y tarifas.
    rotulos = " ".join(_norm(v) for v in filas[i0][j + 1:j + 7] if v is not None)
    antiguo = "TOTAL" in rotulos
    if not antiguo:
        for _, r in cuerpo:
            m2_, t_, v_ = g0(r, 1), g0(r, 2), g0(r, 3)
            if m2_ and t_ and v_:
                antiguo = abs(m2_ * t_ - v_) <= abs(v_) * 0.001
                break
    for m, r in cuerpo:
        if previo is not None and m < previo:
            anio += 1
        previo = m
        g = lambda k: g0(r, k)
        if antiguo:   # m², $/m² MTA, total MTA, $/m² NMT, total NMT, árboles, poda MTA, poda NMT
            out.append(dict(anio=anio, mes=m, m2=g(1), tarifa_corte_mta=g(2), valor_corte_mta=g(3), tarifa_corte_nmt=g(4),
                            valor_corte_nmt=g(5), arboles=g(6), valor_poda_mta=g(7), valor_poda_nmt=g(8)))
        else:         # m², $/m² MTA, $/m² NMT, árboles, poda MTA, poda NMT: los totales de corte se calculan
            m2_ = g(1)
            out.append(dict(anio=anio, mes=m, m2=m2_, tarifa_corte_mta=g(2),
                            valor_corte_mta=(m2_ * g(2) if m2_ is not None and g(2) is not None else None),
                            tarifa_corte_nmt=g(3),
                            valor_corte_nmt=(m2_ * g(3) if m2_ is not None and g(3) is not None else None),
                            arboles=g(4), valor_poda_mta=g(5), valor_poda_nmt=g(6)))
    d = pd.DataFrame(out).apply(pd.to_numeric, errors="coerce")        # columnas sin datos quedan numéricas (NaN)
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


# ------------------------------------------------------------------ cálculo
def tabla_mensual(ppto: pd.DataFrame, ejec: pd.DataFrame) -> pd.DataFrame:
    """Una fila por mes: presupuesto, ejecución (vacía si no hay dato: no es cero) y ejecución valorizada con cada tarifa."""
    ppto = ppto.apply(pd.to_numeric, errors="coerce")
    ejec = ejec.apply(pd.to_numeric, errors="coerce")
    t = ppto.merge(ejec.rename(columns={"m2": "m2_ej", "arboles": "arboles_ej", "valor_corte": "fact_corte", "valor_poda": "fact_poda"})
                   [["anio", "mes", "m2_ej", "arboles_ej", "fact_corte", "fact_poda"]], on=["anio", "mes"], how="left")
    for esc in ("mta", "nmt"):
        t[f"corte_ej_{esc}"] = t["m2_ej"] * t[f"tarifa_corte_{esc}"]
        por_arbol = t[f"valor_poda_{esc}"] / t["arboles"].where(t["arboles"] > 0)
        t[f"poda_ej_{esc}"] = t["arboles_ej"] * por_arbol
    # Lo facturado es la ejecución real en pesos: manda sobre la valorización con tarifa en la columna MTA (vigente hoy).
    t["corte_ej_mta"] = t["fact_corte"].where(t["fact_corte"].notna(), t["corte_ej_mta"])
    t["poda_ej_mta"] = t["fact_poda"].where(t["fact_poda"].notna(), t["poda_ej_mta"])
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
        titulo_tarjeta("Corte y poda", "Presupuesto (libro de servicios proyectados) y ejecución mensual digitada: m², árboles y lo facturado."
                   if MOSTRAR_PODA else "Presupuesto (libro de servicios proyectados) y ejecución mensual digitada: m² y lo facturado.")
        anio0 = st.number_input("Año del primer mes del libro", 2024, 2035, 2026, key="cp_anio0",
                                help="El año sube solo al pasar de diciembre a enero.")
        a, b = st.tabs(["Presupuesto", "Ejecución (histórico manual)"])
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
            _digitar_mes(ctx)
            with st.expander("Histórico completo (corregir o borrar meses)"):
                _editor_historico(ctx)


def _digitar_mes(ctx):
    """Digitación de un mes: solo corte o solo poda. Guarda únicamente lo que se digita y conserva lo demás del mes."""
    con, leer = ctx.con, ctx.leer
    c = st.columns([1.2, 0.8, 1.2], vertical_alignment="bottom")
    tipo = c[0].segmented_control("Ejecución de", ["Corte", "Poda"], default="Corte", key="cp_tipo") or "Corte"
    hoy = ahora()
    anio = int(c[1].number_input("Año", 2024, 2040, hoy.year, key="cp_f_anio"))
    mes = c[2].selectbox("Mes", list(range(1, 13)), index=hoy.month - 1, key="cp_f_mes", format_func=lambda m: MESES_L[m - 1].capitalize())
    ya = leer("SELECT * FROM cp_ejec WHERE anio=? AND mes=?", (anio, mes))
    ya = ya.iloc[0] if len(ya) else None
    prev = lambda k: None if ya is None or pd.isna(ya[k]) else float(ya[k])
    if ya is not None:
        st.caption(f"{MESES_L[mes - 1].capitalize()} {anio} ya tiene datos: lo que guardes los actualiza.")
    k = f"{anio}_{mes}_{tipo}"                       # al cambiar de mes o tipo, los campos toman el valor guardado
    if tipo == "Corte":
        f = st.columns(3)
        m2 = f[0].number_input("m² ejecutados", 0.0, value=prev("m2"), step=1000.0, format="%.0f", key=f"cp_m2_{k}")
        tar = f[1].number_input("$/m²", 0.0, value=prev("tarifa"), step=0.1, format="%.2f", key=f"cp_tar_{k}")
        t0 = prev("valor_corte")             # si el total guardado es m² × $/m², se deja en blanco para que se recalcule al cambiar m²
        if t0 is not None and prev("m2") is not None and prev("tarifa") is not None and abs(t0 - prev("m2") * prev("tarifa")) < 1:
            t0 = None
        tot = f[2].number_input("Total corte ($)", 0.0, value=t0, step=1000.0, format="%.0f", key=f"cp_tot_{k}",
                                help="En blanco = m² × $/m²")
        nuevo = dict(m2=m2, tarifa=tar, valor_corte=tot if tot is not None else (round(m2 * tar) if m2 is not None and tar is not None else None))
        vacio = m2 is None and tot is None
    else:
        f = st.columns(2)
        arb = f[0].number_input("Árboles podados", 0.0, value=prev("arboles"), step=1.0, format="%.0f", key=f"cp_arb_{k}")
        tot = f[1].number_input("Total poda ($)", 0.0, value=prev("valor_poda"), step=1000.0, format="%.0f", key=f"cp_totp_{k}")
        nuevo = dict(arboles=arb, valor_poda=tot)
        vacio = arb is None and tot is None
    if st.button(f"Guardar {tipo.lower()} de {MESES_L[mes - 1]} {anio}", type="primary", key="cp_g_mes", disabled=vacio):
        fila = {c_: (None if ya is None or pd.isna(ya[c_]) else ya[c_]) for c_ in ("m2", "tarifa", "valor_corte", "arboles", "valor_poda")}
        fila.update(nuevo)
        con.execute("DELETE FROM cp_ejec WHERE anio=? AND mes=?", (anio, mes))
        con.anexar(pd.DataFrame([dict(anio=anio, mes=mes, **fila, actualizado=hoy.strftime("%Y-%m-%d %H:%M"))]), "cp_ejec"); con.commit()
        ctx.registrar_carga(ctx.usuario, f"Corte y poda · ejecución {tipo.lower()}", "manual", f"{MESES_L[mes - 1]} {anio}")
        st.success(f"Guardado: {tipo.lower()} de {MESES_L[mes - 1]} {anio}.")
        st.rerun()


def _editor_historico(ctx):
    """Histórico manual de ejecución: se agregan, corrigen o borran meses y se guarda todo junto. Valores en blanco: el total de
    corte se calcula como m² × $/m² y la poda queda sin valor."""
    con, leer = ctx.con, ctx.leer
    h = leer("SELECT anio, mes, m2, tarifa, valor_corte, arboles, valor_poda FROM cp_ejec ORDER BY anio, mes")
    primera = h.empty
    if primera:                                      # primera vez: lo ya facturado, para revisar y guardar
        h = pd.DataFrame({"anio": [2026, 2026], "mes": [9, 10], "m2": [9382754.0, 9342238.0], "tarifa": [151.2, 151.2],
                          "valor_corte": [1418672405.0, 1412546386.0], "arboles": [None, None], "valor_poda": [None, None]})
        st.caption("Precargado con lo facturado de corte de septiembre y octubre; revísalo y guarda.")
    h = h.assign(mes=h["mes"].map(lambda m: MESES_L[int(m) - 1].capitalize()))
    h = h.rename(columns={"anio": "Año", "mes": "Mes", "m2": "m² ejecutados", "tarifa": "$/m²", "valor_corte": "Total corte ($)",
                          "arboles": "Árboles podados", "valor_poda": "Total poda ($)"})
    ed = st.data_editor(h, num_rows="dynamic", hide_index=True, width="stretch", key="cp_editor", column_config={
        "Año": st.column_config.NumberColumn(format="%d", min_value=2024, max_value=2040, step=1, required=True),
        "Mes": st.column_config.SelectboxColumn(options=[m.capitalize() for m in MESES_L], required=True),
        "m² ejecutados": st.column_config.NumberColumn(format="%.0f", min_value=0),
        "$/m²": st.column_config.NumberColumn(format="%.2f", min_value=0),
        "Total corte ($)": st.column_config.NumberColumn(format="%.0f", help="En blanco = m² × $/m²", min_value=0),
        "Árboles podados": st.column_config.NumberColumn(format="%.0f", min_value=0),
        "Total poda ($)": st.column_config.NumberColumn(format="%.0f", min_value=0)})
    if st.button("Guardar histórico", type="primary", key="cp_g_ejec"):
        d = ed.dropna(subset=["Año", "Mes"]).copy()
        d["mes"] = d["Mes"].map(lambda m: [x.capitalize() for x in MESES_L].index(m) + 1)
        d["anio"] = d["Año"].astype(int)
        if d.duplicated(["anio", "mes"]).any():
            dup = d[d.duplicated(["anio", "mes"], keep=False)]
            st.error("Hay meses repetidos: " + ", ".join(sorted({f"{MESES_L[m - 1]} {a}" for a, m in zip(dup["anio"], dup["mes"])})))
            return
        d["valor_corte"] = d["Total corte ($)"].where(d["Total corte ($)"].notna(), (d["m² ejecutados"] * d["$/m²"]).round(0))
        out = d.assign(m2=d["m² ejecutados"], tarifa=d["$/m²"], arboles=d["Árboles podados"], valor_poda=d["Total poda ($)"],
                       actualizado=ahora().strftime("%Y-%m-%d %H:%M"))[
            ["anio", "mes", "m2", "tarifa", "valor_corte", "arboles", "valor_poda", "actualizado"]]
        con.execute("DELETE FROM cp_ejec")
        con.anexar(out, "cp_ejec"); con.commit()
        ctx.registrar_carga(ctx.usuario, "Corte y poda · ejecución", "manual", f"{len(out)} meses")
        st.success(f"Histórico guardado: {len(out)} meses.")


MOSTRAR_PODA = False        # de momento la página solo muestra el corte; la poda se sigue cargando y guardando en Administración


# ------------------------------------------------------------------ pestaña
def mostrar(ctx):
    leer = ctx.leer
    ppto, ejec = leer("SELECT * FROM cp_ppto"), leer("SELECT * FROM cp_ejec")
    zona = st.container()
    if ppto.empty:
        zona.html(E.encabezado("Corte y poda", "Ejecución de área y de árboles contra el presupuesto" if MOSTRAR_PODA else "Ejecución de área cortada contra el presupuesto", []))
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
    zona.html(E.encabezado("Corte y poda", (f"Área cortada y árboles podados · {per.texto}" if MOSTRAR_PODA else f"Área cortada · {per.texto}"), chips))
    if t["m2_ej"].isna().all() and (not MOSTRAR_PODA or t["arboles_ej"].isna().all()):
        st.info("No hay ejecución cargada para este periodo.")
        return
    r = resumen(t)
    kpis = [
        dict(label="Área cortada (m²)", valor=_n(r["m2_e"]), sub=f"{_fmt_pct(_pct(r['m2_e'], r['m2_p']))} de {_n(r['m2_p'])} m²",
             icono="receipt", color=E.AZUL, fondo=E.AZUL_SUAVE),
        dict(label="Árboles podados", valor=_n(r["ar_e"]), sub=(f"{_fmt_pct(_pct(r['ar_e'], r['ar_p']))} de {_n(r['ar_p'])}" if r["ar_p"] else "sin meta de árboles"),
             icono="calendar", color=E.AZUL, fondo=E.AZUL_SUAVE),
        dict(label="Corte valorizado MTA", valor=E.millones(r["corte_e_mta"]), sub=f"{_fmt_pct(_pct(r['corte_e_mta'], r['corte_p_mta']))} de {E.millones(r['corte_p_mta'])}",
             icono="receipt", color=E.AZUL, fondo=E.AZUL_SUAVE),
        dict(label="Corte valorizado NMT", valor=E.millones(r["corte_e_nmt"]), sub=f"{_fmt_pct(_pct(r['corte_e_nmt'], r['corte_p_nmt']))} de {E.millones(r['corte_p_nmt'])}",
             icono="receipt", color=E.AZUL, fondo=E.AZUL_SUAVE)]
    if not MOSTRAR_PODA:
        kpis = [k for k in kpis if k["label"] != "Árboles podados"]
    st.html(E.kpis(kpis))

    mes_txt = [MESES_L[m - 1].capitalize() for m in t["mes"]]
    graficos = [("Área (m²) por mes", "m2", "m2_ej")] + ([("Árboles por mes", "arboles", "arboles_ej")] if MOSTRAR_PODA else [])
    for col, (titulo, p, e) in zip(st.columns(len(graficos)), graficos):
        with col:
            titulo_tarjeta(titulo, "Barras: m² presupuestados y ejecutados. Línea (eje derecho): precio unitario por m² (tarifa MTA)." if p == "m2" else "")
            fmt = lambda s_: s_.map(lambda v: "" if pd.isna(v) else f"{v:,.0f}".replace(",", "."))
            fig = go.Figure([go.Bar(x=mes_txt, y=t[p], name="Presupuesto", marker_color=COLOR_PPTO, marker_line_width=0,
                                    text=fmt(t[p]), textposition="outside", cliponaxis=False),
                             go.Bar(x=mes_txt, y=t[e], name="Ejecución", marker_color=COLOR_EJEC, marker_line_width=0,
                                    text=fmt(t[e]), textposition="outside", cliponaxis=False)])
            fig.update_layout(barmode="group", bargap=.3, yaxis=dict(visible=False, rangemode="tozero"), legend=dict(orientation="h", y=1.12, x=0))
            tar = t["tarifa_corte_mta"] if p == "m2" else None
            if tar is not None and tar.notna().any():
                lo, hi = float(tar.min()), float(tar.max())
                holgura = max((hi - lo) * 0.5, hi * 0.05)
                fig.add_trace(go.Scatter(x=mes_txt, y=tar, name="Precio unitario ($/m²)", yaxis="y2", mode="lines+markers+text",
                                         line=dict(color="#e8912d", width=2), marker=dict(size=6), connectgaps=False,
                                         text=tar.map(lambda v: "" if pd.isna(v) else f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")),
                                         textposition="top center", cliponaxis=False))
                fig.update_layout(yaxis2=dict(overlaying="y", side="right", showgrid=False, range=[lo - holgura * 3, hi + holgura],
                                              tickformat=",.0f", title=dict(text="$/m²", font=dict(size=11))))
            st.plotly_chart(estilo_grafico(fig, 300), width="stretch", config={"displayModeBar": False})

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
    if MOSTRAR_PODA:
        cfg_p = {c: st.column_config.NumberColumn(format="%.0f") for c in poda.columns[1:] if c != "% árboles"}
        cfg_p["% árboles"] = st.column_config.NumberColumn(format="percent")
        with st.container(key="card_cp_poda"):
            titulo_tarjeta("Poda de árboles", "Valor por árbol = valor mensual de poda ÷ árboles del presupuesto; sin árboles presupuestados no se valoriza.")
            st.dataframe(poda, hide_index=True, width="stretch", column_config=cfg_p)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf) as w:
        corte.to_excel(w, sheet_name="Corte", index=False)
        if MOSTRAR_PODA:
            poda.to_excel(w, sheet_name="Poda", index=False)
    st.download_button("Descargar Excel", buf.getvalue(), f"corte_poda_{anio}_{per.mes:02d}.xlsx", icon=":material/download:", key="cp_dl")
