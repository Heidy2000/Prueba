"""Módulo Estados financieros: balance, estado de resultados e indicadores financieros mes a mes (desde SAP)."""
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import eeff
import estilo as E
from ui import EMPRESAS, MESES, MESES_L, a_excel, estilo_grafico, titulo_tarjeta


@st.cache_data(show_spinner=False, max_entries=8)
def _datos(_leer, empresa: str, version: str):
    lin = _leer("SELECT estado, idx, grupo, subgrupo, linea, anio, mes, valor FROM eeff WHERE empresa=?", (empresa,))
    return lin, eeff.resumen(lin)


def _var(a, b):
    return (a / b - 1) if b and not pd.isna(b) and abs(b) > 1e-6 and not pd.isna(a) else float("nan")


def mostrar(ctx):
    leer = ctx.leer
    zona = st.container()
    hay = set(leer("SELECT DISTINCT empresa FROM eeff")["empresa"])
    with st.container(key="card_filtros"):
        st.html('<div class="nf-ftit">Filtros del informe</div>')
        c = st.columns([0.9, 0.7, 1.0, 1.4, 1.6], vertical_alignment="bottom")
        empresa = c[0].selectbox("Compañía", list(EMPRESAS), key="ef_empresa",
                                 help=" · ".join(f"{e}: {n}" for e, n in EMPRESAS.items()))
        if empresa not in hay:
            zona.html(E.encabezado(f"Estados financieros · {empresa}", EMPRESAS[empresa], []))
            st.info(f"Aún no hay estados financieros de **{empresa}**. Financiera carga el histórico y el balance de "
                    "SAP de cada mes en Administración → Estados financieros.", icon=":material/info:")
            return
        v = leer("SELECT COUNT(*) n, SUM(valor) s FROM eeff WHERE empresa=?", (empresa,))
        lin, r = _datos(leer, empresa, f"{v['n'][0]}-{round(float(v['s'][0] or 0))}")
        periodos = list(r.index)
        anios = sorted({a for a, _ in periodos}, reverse=True)
        anio = c[1].selectbox("Año", anios, key=f"ef_anio_{empresa}")
        meses = sorted(m for a, m in periodos if a == anio)
        mes = c[2].selectbox("Mes de corte", meses, index=len(meses) - 1, key=f"ef_mes_{empresa}_{anio}",
                             format_func=lambda m: MESES_L[m - 1].capitalize())
        vista = c[3].segmented_control("Estado de resultados", ["Acumulado", "Solo el mes"], default="Acumulado",
                                       key="ef_vista") or "Acumulado"
    SOLO = vista == "Solo el mes"
    k, k_ant = (anio, mes), (anio - 1, mes)
    k_dic = (anio - 1, 12)
    fila = r.loc[k]
    ant = r.loc[k_ant] if k_ant in r.index else None
    fuente = leer("SELECT DISTINCT fuente FROM eeff WHERE empresa=? AND anio=? AND mes=?", (empresa, anio, mes))
    chips = [(f"Corte: {MESES[mes - 1]} {anio}", ""),
             ("Balance de SAP" if "sap" in set(fuente["fuente"]) else "Histórico", "ok")]
    descuadre = fila["ACTIVO"] - (fila["PASIVO"] + fila["PAT"])
    if abs(descuadre) > 1000:
        chips.append((f"Descuadre balance: {E.millones(descuadre)}", "warn"))
    zona.html(E.encabezado(f"Estados financieros · {empresa}",
                           f"{EMPRESAS[empresa]} · cifras en millones de pesos", chips))

    t1, t2, t3, t4 = st.tabs(["Indicadores", "Estado de resultados", "Balance", "Tendencias"],
                             on_change="rerun", key="ef_tabs")
    suf = "_m" if SOLO else ""
    txt_per = f"{MESES_L[mes - 1]} {anio}" if SOLO else f"enero a {MESES_L[mes - 1]} {anio}"

    # ================= INDICADORES
    with t1:
        if t1.open:
            def kpi(nombre, valor_actual, valor_ant, icono, fmt="m", margen=None, bueno_si_sube=True):
                var = _var(valor_actual, valor_ant) if fmt == "m" else (
                    valor_actual - valor_ant if valor_ant is not None and not pd.isna(valor_ant) else float("nan"))
                sube = (var > 0) if not pd.isna(var) else None
                color = E.AZUL if sube is None else (E.VERDE if sube == bueno_si_sube else E.ROJO)
                val = E.millones(valor_actual) if fmt == "m" else (
                    E.pct_txt(valor_actual) if fmt == "p" else f"{valor_actual:,.2f}".replace(".", ","))
                sub = ""
                if margen is not None:
                    sub += f"Margen {E.pct_txt(margen)} · "
                if not pd.isna(var):
                    sub += (f'<b style="color:{color}">{"+" if var > 0 else ""}'
                            + (E.pct_txt(var) if fmt == "m" else
                               (f"{var * 100:.1f} pp".replace(".", ",") if fmt == "p" else
                                f"{var:.2f}".replace(".", ","))) + "</b> vs año anterior")
                else:
                    sub += "Sin dato del año anterior"
                return dict(label=nombre, valor=val, sub=sub, icono=icono, color=color,
                            fondo=E.VERDE_SUAVE if color == E.VERDE else (E.ROJO_SUAVE if color == E.ROJO
                                                                           else E.AZUL_SUAVE))
            g = lambda col: (ant[col] if ant is not None else None)                     # noqa: E731
            st.html(E.kpis([
                kpi(f"Ingresos {'del mes' if SOLO else 'acumulados'}", fila["ing" + suf], g("ing" + suf), "wallet"),
                kpi("EBITDA", fila["ebitda" + suf], g("ebitda" + suf), "piggy",
                    margen=fila["ebitda" + suf] / fila["ing" + suf] if fila["ing" + suf] else None),
                kpi("Utilidad neta", fila["neto" + suf], g("neto" + suf), "check",
                    margen=fila["neto" + suf] / fila["ing" + suf] if fila["ing" + suf] else None),
                kpi("Razón corriente", fila["razon_corriente"], g("razon_corriente"), "calendar", fmt="x"),
            ]))
            with st.container(key="card_ef_ind"):
                titulo_tarjeta("Indicadores financieros",
                               f"Corte {MESES_L[mes - 1]} {anio} frente al mismo mes del año anterior y al cierre del "
                               "año anterior. Márgenes y rotaciones con el acumulado del año.")
                cols_ref = [k, k_ant, k_dic]
                nombres = [f"{MESES[mes - 1]} {anio}", f"{MESES[mes - 1]} {anio - 1}", f"Dic {anio - 1}"]

                def vals(col):
                    out = [r.loc[x, col] if x in r.index else None for x in cols_ref]
                    out.append(out[0] - out[1] if out[1] is not None and not pd.isna(out[1]) and
                               not pd.isna(out[0]) else None)
                    return out
                secciones = [
                    ("Rentabilidad", "p", [("Margen bruto", "m_bruto", 1), ("Margen EBITDA", "m_ebitda", 1),
                                           ("Margen operacional", "m_operacional", 1),
                                           ("Margen antes de impuestos", "m_rai", 1), ("Margen neto", "m_neto", 1),
                                           ("ROE (anualizado)", "roe", 1), ("ROA operativo (anualizado)", "roa", 1)]),
                    ("Liquidez", "x", [("Razón corriente (veces)", "razon_corriente", 1),
                                       ("Prueba ácida (veces)", "prueba_acida", 1)]),
                    ("Capital de trabajo", "m", [("Capital de trabajo neto contable", "ktnc", 1),
                                                 ("Capital de trabajo operativo", "kto", 0),
                                                 ("Capital de trabajo neto operativo", "ktno", 0)]),
                    ("Endeudamiento", "p", [("Nivel de endeudamiento", "endeudamiento", -1),
                                            ("Concentración en el corto plazo", "corto_plazo", -1)]),
                    ("Cobertura y deuda", "x", [("Apalancamiento total (pasivo / patrimonio)", "apalancamiento", -1),
                                                ("Cobertura EBITDA / gastos financieros netos", "cobertura", 1),
                                                ("Deuda financiera neta / EBITDA 12 meses", "deuda_ebitda", -1)]),
                    ("Operación (días)", "x", [("Días de cartera", "dias_cartera", -1),
                                               ("Días de proveedores", "dias_proveedores", 1),
                                               ("Ciclo de efectivo", "ciclo_efectivo", -1)]),
                ]
                for titulo, fmt, items in secciones:
                    nodos = [dict(nombre=n, clase="linea", valores=vals(col), dif_col=3, dif_inv=(sig < 0))
                             for n, col, sig in items]
                    st.html(E.tabla_meses([dict(nombre=titulo, clase="seccion", valores=[None] * 4, hijos=nodos)],
                                          nombres + ["Variación vs año ant."], abierto=True,
                                          formatos=[fmt] * 4))
                st.caption("Verde = mejora frente al año anterior. ROE = utilidad neta acumulada anualizada / "
                           "patrimonio. ROA operativo = utilidad operacional × (1 − 35 %) anualizada / (capital de "
                           "trabajo operativo + propiedad, planta y equipo + intangibles). Capital de trabajo operativo "
                           "= clientes + deudores varios + inventarios; neto operativo = menos cuentas comerciales por "
                           "pagar. Días de cartera y de proveedores sobre ventas y egresos operacionales acumulados.")

    # ================= ESTADO DE RESULTADOS
    with t2:
        if t2.open:
            with st.container(key="card_ef_pyg"):
                titulo_tarjeta("Estado de resultados", f"{'Solo ' if SOLO else ''}{txt_per} frente al mismo periodo "
                               f"del año anterior · haz clic en un grupo para ver sus cuentas.")
                p = lin[lin["estado"] == "PYG"]

                def v_linea(a, m, idx):
                    x = p[(p["anio"] == a) & (p["mes"] == m) & (p["idx"] == idx)]["valor"].sum()
                    if SOLO and m > 1:
                        x -= p[(p["anio"] == a) & (p["mes"] == m - 1) & (p["idx"] == idx)]["valor"].sum()
                    return float(x)
                hay_ant = k_ant in r.index
                ing_act = fila["ing" + suf]

                def fila_v(act, prev, costo=False):
                    var = act - prev if hay_ant else None
                    return [act, prev if hay_ant else None, var,
                            (act / prev - 1) if hay_ant and prev else None, act / ing_act if ing_act else None]
                nodos = []
                lineas_pyg = pd.DataFrame(eeff.PYG, columns=["grupo", "linea", "c", "cod", "s", "sig"])
                subt = {"COSTOS": ("RESULTADO BRUTO", "bruto"), "GASTOS DE ADMINISTRACIÓN": ("EBITDA", "ebitda"),
                        "AMORTIZACIONES": ("RESULTADO OPERACIONAL", "operacional"),
                        "COSTOS FINANCIEROS": ("RESULTADO ANTES DE IMPUESTOS", "rai"),
                        "IMPUESTO DIFERIDO": ("UTILIDAD NETA", "neto")}
                col_g = {"INGRESOS": "ing", "COSTOS": "cos", "GASTOS DE ADMINISTRACIÓN": "adm", "DETERIORO": "det",
                         "DEPRECIACIONES": "dep", "AMORTIZACIONES": "amo", "OTROS INGRESOS": "oi",
                         "OTROS GASTOS": "og", "INGRESOS FINANCIEROS": "if", "COSTOS FINANCIEROS": "cf",
                         "IMPUESTO CORRIENTE": "impc", "IMPUESTO DIFERIDO": "impd"}
                for grupo in eeff.GRUPOS_PYG:
                    costo = eeff.SIGNO_GRUPO[grupo] < 0
                    hijos = []
                    for idx in lineas_pyg.index[lineas_pyg["grupo"] == grupo]:
                        a_ = v_linea(anio, mes, idx)
                        b_ = v_linea(anio - 1, mes, idx) if hay_ant else 0.0
                        if abs(a_) < 0.5 and abs(b_) < 0.5:
                            continue
                        hijos.append(dict(nombre=lineas_pyg.at[idx, "linea"].strip().capitalize(), clase="linea",
                                          valores=fila_v(a_, b_, costo), dif_col=2, dif_inv=costo))
                    act = fila[col_g[grupo] + suf]
                    prev = ant[col_g[grupo] + suf] if hay_ant else 0.0
                    if abs(act) < 0.5 and abs(prev) < 0.5:
                        continue
                    nodos.append(dict(nombre=grupo.capitalize(), clase="grupo", hijos=hijos,
                                      valores=fila_v(act, prev, costo), dif_col=2, dif_inv=costo))
                    if grupo in subt:
                        nombre, col = subt[grupo]
                        nodos.append(dict(nombre=nombre.capitalize(), clase="saldo", dif_col=2,
                                          valores=fila_v(fila[col + suf], ant[col + suf] if hay_ant else 0.0)))
                st.html(E.tabla_meses(nodos, [txt_per.capitalize() if SOLO else f"Acum. {MESES[mes - 1]} {anio}",
                                              f"{'' if SOLO else 'Acum. '}{MESES[mes - 1]} {anio - 1}", "Variación",
                                              "Var. %", "% de ingresos"],
                                      formatos=["m", "m", "m", "p", "p"]))
                st.caption("Ingresos y gastos en positivo. Variación en verde cuando favorece el resultado (más ingreso "
                           "o menos gasto).")

    # ================= BALANCE
    with t3:
        if t3.open:
            with st.container(key="card_ef_esf"):
                titulo_tarjeta("Estado de situación financiera",
                               f"Saldos al cierre de {MESES_L[mes - 1]} {anio} frente al cierre del año anterior y al "
                               "mismo mes del año anterior.")
                b = lin[lin["estado"] == "ESF"]
                piv = b.pivot_table(index=["grupo", "subgrupo", "idx", "linea"], columns=["anio", "mes"],
                                    values="valor", aggfunc="sum")

                def col(x):
                    return piv[x] if x in piv.columns else pd.Series(0.0, index=piv.index)
                cur, dic, mant = col(k), col(k_dic), col(k_ant)

                def vv(a_, d_, m_):
                    return [a_, d_, a_ - d_, m_, a_ - m_]
                nodos = []
                orden_g = eeff.GRUPOS_ESF
                for grupo in orden_g:
                    if grupo == "PASIVO CORRIENTE":
                        tot_a = [fila["ACTIVO"], r.loc[k_dic, "ACTIVO"] if k_dic in r.index else 0.0,
                                 r.loc[k_ant, "ACTIVO"] if k_ant in r.index else 0.0]
                        nodos.append(dict(nombre="Total activo", clase="saldo", valores=vv(*tot_a)))
                    hijos = []
                    sub_idx = [s for s in dict.fromkeys(x[1] for x in piv.index if x[0] == grupo)]
                    for sg in sub_idx:
                        m_ = [(g_, s_, i_, l_) for (g_, s_, i_, l_) in piv.index if g_ == grupo and s_ == sg]
                        nietos = []
                        for key in m_:
                            if abs(cur[key]) < 0.5 and abs(dic[key]) < 0.5 and abs(mant[key]) < 0.5:
                                continue
                            nietos.append(dict(nombre=key[3].capitalize(), clase="linea",
                                               valores=vv(cur[key], dic[key], mant[key])))
                        if not nietos:
                            continue
                        sa, sd, sm = (sum(s[key] for key in m_) for s in (cur, dic, mant))
                        hijos.append(dict(nombre=sg.strip().capitalize(), clase="linea", valores=vv(sa, sd, sm),
                                          hijos=nietos if len(nietos) > 1 else []))
                    ga, gd, gm = (sum(s[key] for key in piv.index if key[0] == grupo) for s in (cur, dic, mant))
                    nodos.append(dict(nombre=grupo.capitalize(), clase="grupo", valores=vv(ga, gd, gm), hijos=hijos))
                    if grupo == "PASIVO NO CORRIENTE":
                        tp = [fila["PASIVO"], r.loc[k_dic, "PASIVO"] if k_dic in r.index else 0.0,
                              r.loc[k_ant, "PASIVO"] if k_ant in r.index else 0.0]
                        nodos.append(dict(nombre="Total pasivo", clase="saldo", valores=vv(*tp)))
                pp = [fila["PASIVO"] + fila["PAT"]] + [
                    (r.loc[x, "PASIVO"] + r.loc[x, "PAT"]) if x in r.index else 0.0 for x in (k_dic, k_ant)]
                nodos.append(dict(nombre="Total pasivo + patrimonio", clase="saldo", valores=vv(*pp)))
                st.html(E.tabla_meses(nodos, [f"{MESES[mes - 1]} {anio}", f"Dic {anio - 1}", "Variación año",
                                              f"{MESES[mes - 1]} {anio - 1}", "Variación 12 meses"],
                                      formatos=["m"] * 5))
                st.caption("La porción no corriente de cuentas por cobrar, obligaciones financieras y demás no viene "
                           "en SAP: se mantiene la del mes anterior y Financiera la ajusta en Administración → "
                           "Estados financieros.")
                plano = b[(b["anio"] == anio) & (b["mes"] == mes)][["grupo", "subgrupo", "linea", "valor"]]
                st.download_button("Descargar balance del mes", lambda d=plano: a_excel(d),
                                   f"balance_{empresa}_{anio}_{mes:02d}.xlsx", icon=":material/download:")

    # ================= TENDENCIAS
    with t4:
        if t4.open:
            ult = r.loc[:k].tail(24)
            x = [f"{MESES[m - 1]} {str(a)[2:]}" for a, m in ult.index]
            c1, c2 = st.columns(2, gap="medium")
            with c1, st.container(key="card_ef_t1"):
                titulo_tarjeta("Ingresos y EBITDA del mes", "Últimos 24 meses · millones")
                fig = go.Figure()
                fig.add_bar(x=x, y=ult["ing_m"] / 1e6, name="Ingresos", marker_color="#c9d4e4")
                fig.add_bar(x=x, y=ult["ebitda_m"] / 1e6, name="EBITDA", marker_color=E.AZUL)
                fig.add_scatter(x=x, y=ult["ebitda_m"] / ult["ing_m"].where(ult["ing_m"] != 0) * 100,
                                name="Margen EBITDA %", yaxis="y2", mode="lines+markers",
                                line=dict(color="#e8892c", width=2))
                fig.update_layout(barmode="group", yaxis=dict(tickformat=",.0f"),
                                  yaxis2=dict(overlaying="y", side="right", ticksuffix="%", showgrid=False))
                st.plotly_chart(estilo_grafico(fig, 320), width="stretch", config={"displayModeBar": False})
            with c2, st.container(key="card_ef_t2"):
                titulo_tarjeta("Liquidez y endeudamiento", "Razón corriente (veces) y nivel de endeudamiento (%)")
                fig = go.Figure()
                fig.add_scatter(x=x, y=ult["razon_corriente"], name="Razón corriente", mode="lines+markers",
                                line=dict(color=E.AZUL, width=2.5))
                fig.add_scatter(x=x, y=ult["endeudamiento"] * 100, name="Endeudamiento %", yaxis="y2",
                                mode="lines+markers", line=dict(color="#d23b3b", width=2))
                fig.update_layout(yaxis=dict(tickformat=".2f"),
                                  yaxis2=dict(overlaying="y", side="right", ticksuffix="%", showgrid=False))
                st.plotly_chart(estilo_grafico(fig, 320), width="stretch", config={"displayModeBar": False})
            c3, c4 = st.columns(2, gap="medium")
            with c3, st.container(key="card_ef_t3"):
                titulo_tarjeta("Ciclo de efectivo", "Días de cartera, de proveedores y ciclo (acumulado del año)")
                fig = go.Figure()
                for col_, nom, colr in [("dias_cartera", "Días de cartera", E.AZUL),
                                        ("dias_proveedores", "Días de proveedores", "#2e9e5b"),
                                        ("ciclo_efectivo", "Ciclo de efectivo", "#e8892c")]:
                    fig.add_scatter(x=x, y=ult[col_], name=nom, mode="lines+markers", line=dict(color=colr))
                st.plotly_chart(estilo_grafico(fig, 300), width="stretch", config={"displayModeBar": False})
            with c4, st.container(key="card_ef_t4"):
                titulo_tarjeta("Composición del balance", f"Cierre de {MESES_L[mes - 1]} {anio} · millones")
                fig = go.Figure()
                fig.add_bar(x=["Activo"], y=[fila["AC"] / 1e6], name="Activo corriente", marker_color=E.AZUL)
                fig.add_bar(x=["Activo"], y=[fila["ANC"] / 1e6], name="Activo no corriente", marker_color="#8fb3e8")
                fig.add_bar(x=["Pasivo + patrimonio"], y=[fila["PC"] / 1e6], name="Pasivo corriente",
                            marker_color="#d23b3b")
                fig.add_bar(x=["Pasivo + patrimonio"], y=[fila["PNC"] / 1e6], name="Pasivo no corriente",
                            marker_color="#f0a0a0")
                fig.add_bar(x=["Pasivo + patrimonio"], y=[fila["PAT"] / 1e6], name="Patrimonio",
                            marker_color="#2e9e5b")
                fig.update_layout(barmode="stack", yaxis=dict(tickformat=",.0f"))
                st.plotly_chart(estilo_grafico(fig, 300), width="stretch", config={"displayModeBar": False})
            tabla = r.loc[:k].reset_index()
            st.download_button("Descargar indicadores (todos los meses)", lambda d=tabla: a_excel(d),
                               f"indicadores_{empresa}.xlsx", icon=":material/download:")
