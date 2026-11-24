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
    # vista: estados financieros o cartera de usuarios (como Combustible / Mantenimiento)
    if (st.segmented_control("Vista", ["Estados financieros", "Cartera"], default="Estados financieros",
                             key="ef_vista_mod", label_visibility="collapsed") or "Estados financieros") == "Cartera":
        import mod_cartera
        mod_cartera.mostrar(ctx, zona=zona)
        return
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
            hist = r.loc[:k].tail(12)

            def tend(col):
                return [float(x) if not pd.isna(x) else None for x in hist[col]]

            def margen_m(h):
                return [(e / i) if i else None for e, i in zip(h["operacional_m"], h["ing_m"])]

            def kpi2(nombre, act, prev, icono, fmt, bueno_si_sube, spark, sub=""):
                if prev is None or pd.isna(prev) or pd.isna(act):
                    dl = E.delta("sin año anterior", "na")
                elif fmt == "m":
                    dl = E.delta_valor(act - prev, prev, a_favor_si_sube=bueno_si_sube, sufijo="vs año ant.")
                else:
                    d = act - prev
                    txt = (f"{d * 100:+.1f} pp" if fmt == "p" else (f"{d:+.2f}x" if fmt == "x" else f"{d:+.0f} días"))
                    txt = txt.replace(".", ",")
                    tono = "na" if abs(d) < (0.0005 if fmt == "p" else 0.005 if fmt == "x" else 0.5) else \
                        ("ok" if (d > 0) == bueno_si_sube else "bad")
                    dl = E.delta(txt + " vs año ant.", tono, "up" if d > 0 else "down")
                val = (E.millones(act) if fmt == "m" else E.pct_txt(act) if fmt == "p" else
                       (f"{act:,.2f}x".replace(".", ",") if fmt == "x" else f"{act:,.0f} días".replace(",", ".")))
                color = {"ok": E.VERDE, "bad": E.ROJO}.get(dl.split('"')[1].split()[-1], E.AZUL)
                return dict(label=nombre, valor=val, delta=dl, sub=sub, icono=icono, color=color,
                            fondo={E.VERDE: E.VERDE_SUAVE, E.ROJO: E.ROJO_SUAVE}.get(color, E.AZUL_SUAVE),
                            spark=spark, spark_color=E.AZUL)
            # EBIT = resultado operacional (después de depreciaciones, amortizaciones y deterioro)
            m_e = fila["operacional" + suf] / fila["ing" + suf] if fila["ing" + suf] else float("nan")
            m_e_ant = (ant["operacional" + suf] / ant["ing" + suf]) if ant is not None and ant["ing" + suf] else None
            st.html(E.kpis([
                kpi2(f"Ingresos {'del mes' if SOLO else 'acumulados'}", fila["ing" + suf], g("ing" + suf), "wallet",
                     "m", True, tend("ing_m")),
                kpi2("EBIT", fila["operacional" + suf], g("operacional" + suf), "piggy", "m", True,
                     tend("operacional_m"), sub="Resultado operacional"),
                kpi2("Utilidad neta", fila["neto" + suf], g("neto" + suf), "check", "m", True, tend("neto_m"),
                     sub=f"Margen {E.pct_txt(fila['neto' + suf] / fila['ing' + suf] if fila['ing' + suf] else None)}"),
                kpi2("Margen EBIT", m_e, m_e_ant, "trend", "p", True, margen_m(hist)),
                kpi2("Razón corriente", fila["razon_corriente"], g("razon_corriente"), "calendar", "x", True,
                     tend("razon_corriente")),
            ], columnas=5))
            _vista_eeff(r, fila, ant, k, anio, mes, SOLO, suf, txt_per)
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
                    ("Rentabilidad", "p", [("Margen bruto", "m_bruto", 1),
                                           ("Margen EBIT (operacional)", "m_operacional", 1),
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
                                                ("Cobertura EBIT / gastos financieros netos", "cobertura_ebit", 1),
                                                ("Deuda financiera neta / EBIT 12 meses", "deuda_ebit", -1)]),
                    ("Operación (días)", "x", [("Días de cartera", "dias_cartera", -1),
                                               ("Días de proveedores", "dias_proveedores", 1)]),
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
                subt = {"COSTOS": ("RESULTADO BRUTO", "bruto"),
                        "AMORTIZACIONES": ("RESULTADO OPERACIONAL (EBIT)", "operacional"),
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
                        nodos.append(dict(nombre=nombre.capitalize().replace("(ebit)", "(EBIT)"), clase="saldo", dif_col=2,
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
                titulo_tarjeta("Ingresos y EBIT del mes", "Últimos 24 meses · millones")
                fig = go.Figure()
                fig.add_bar(x=x, y=ult["ing_m"] / 1e6, name="Ingresos", marker_color="#c9d4e4")
                fig.add_bar(x=x, y=ult["operacional_m"] / 1e6, name="EBIT", marker_color=E.AZUL)
                fig.add_scatter(x=x, y=ult["operacional_m"] / ult["ing_m"].where(ult["ing_m"] != 0) * 100,
                                name="Margen EBIT %", yaxis="y2", mode="lines+markers",
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
                titulo_tarjeta("Días de cartera y de proveedores", "Acumulado del año")
                fig = go.Figure()
                for col_, nom, colr in [("dias_cartera", "Días de cartera", E.AZUL),
                                        ("dias_proveedores", "Días de proveedores", "#2e9e5b")]:
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


def _vista_eeff(r, fila, ant, k, anio, mes, SOLO, suf, txt_per):
    """Lectura, estado de resultados resumido, flujo de efectivo indirecto, situación financiera y notas."""
    if SOLO:
        k_b = (anio, mes - 1) if mes > 1 else (anio - 1, 12)
        nom_b = f"{MESES[k_b[1] - 1]} {k_b[0]}"
    else:
        k_b = (anio - 1, 12)
        nom_b = f"Dic {anio - 1}"
    base = r.loc[k_b] if k_b in r.index else None
    hay_ant = ant is not None

    def v(col, fuente=fila):
        x = fuente[col] if fuente is not None else float("nan")
        return 0.0 if pd.isna(x) else float(x)

    # ---------------- flujo de efectivo (método indirecto) desde el balance
    flujo = None
    if base is not None:
        da = v("dep" + suf) + v("amo" + suf) + v("det" + suf)
        d_kt = -(v("ktno") - v("ktno", base))
        f_op = v("neto" + suf) + da + d_kt
        capex = -((v("activo_fijo") - v("activo_fijo", base)) + v("dep" + suf) + v("amo" + suf))
        d_deuda = v("deuda_fin") - v("deuda_fin", base)
        d_caja = v("efectivo") - v("efectivo", base)
        otros = d_caja - (f_op + capex + d_deuda)
        flujo = dict(neto=v("neto" + suf), da=da, kt=d_kt, op=f_op, capex=capex, deuda=d_deuda, otros=otros,
                     caja=d_caja, ini=v("efectivo", base), fin=v("efectivo"))

    # ---------------- lectura
    ing, ing_a = v("ing" + suf), v("ing" + suf, ant) if hay_ant else None
    gas, gas_a = v("cos" + suf) + v("adm" + suf), (v("cos" + suf, ant) + v("adm" + suf, ant)) if hay_ant else None
    ebi, ebi_a = v("operacional" + suf), v("operacional" + suf, ant) if hay_ant else None     # EBIT
    neto = v("neto" + suf)

    def pc(a_, b_):
        return (a_ / b_ - 1) if b_ else None
    if hay_ant:
        g_i, g_g, g_e = pc(ing, ing_a), pc(gas, gas_a), pc(ebi, ebi_a)
        titulo = (f"Ingresos {E.pct_txt(g_i) if g_i is None else ('+' if g_i >= 0 else '') + E.pct_txt(g_i)} "
                  f"frente al año anterior, EBIT {('+' if (g_e or 0) >= 0 else '') + E.pct_txt(g_e)}.")
        if g_i is not None and g_g is not None and g_g > g_i + 0.005:
            texto = (f"Costos y gastos crecen más rápido (<b>+{E.pct_txt(g_g)}</b>) y absorben parte del crecimiento "
                     f"de los ingresos.")
        elif g_i is not None and g_g is not None:
            texto = f"Costos y gastos crecen menos que los ingresos (<b>{('+' if g_g >= 0 else '') + E.pct_txt(g_g)}</b>)."
        else:
            texto = ""
        tono = "ok" if (g_e or 0) >= 0 else ("warn" if (g_e or 0) > -0.1 else "bad")
    else:
        titulo, texto, tono = f"EBIT {txt_per}: {E.millones(ebi)}.", "No hay datos del año anterior para comparar.", ""
    if flujo is not None and abs(neto) > 1:
        conv = flujo["op"] / neto
        texto += (f" La utilidad ({E.millones(neto)}) no es caja: el flujo operativo fue <b>{E.millones(flujo['op'])}</b>"
                  f" ({E.pct_txt(conv)} de conversión) y la caja {'subió' if flujo['caja'] >= 0 else 'bajó'} "
                  f"{E.millones(abs(flujo['caja']))} desde {nom_b.lower()}.")
    st.html(E.lectura(titulo, texto, tono))

    # ---------------- estado de resultados resumido + flujo de efectivo
    izq, der = st.columns([1.3, 1], gap="medium")
    with izq, st.container(key="card_ef_res"):

        titulo_tarjeta("Estado de resultados", f"{txt_per.capitalize()} vs mismo periodo del año anterior · verde = "
                       "favorable (más ingreso o menos gasto)")

        def fl(nombre, col, clase="linea", costo=False, cols=None):
            act = sum(v(c + suf) for c in (cols or [col]))
            prev = sum(v(c + suf, ant) for c in (cols or [col])) if hay_ant else None
            dif = (act - prev) if prev is not None else None
            return dict(nombre=nombre, clase=clase, dif_col=2, dif_inv=costo, valores=[act, prev, dif],
                        sub=(("+" if act >= prev else "") + E.pct_txt(act / prev - 1)) if prev else "")
        fin_neto_a = v("if" + suf) + v("oi" + suf) - v("cf" + suf) - v("og" + suf)
        nodos = [fl("Ingresos", "ing", "grupo"), fl("Costo de ventas", "cos", costo=True),
                 fl("Resultado bruto", "bruto", "saldo"), fl("Gastos de administración", "adm", costo=True),
                 fl("Depreciación, amortización y deterioro", None, costo=True, cols=["dep", "amo", "det"]),
                 fl("EBIT (resultado operacional)", "operacional", "saldo")]
        prev_fin = (v("if" + suf, ant) + v("oi" + suf, ant) - v("cf" + suf, ant) - v("og" + suf, ant)) if hay_ant \
            else None
        nodos.append(dict(nombre="Financieros y otros (neto)", clase="linea", dif_col=2,
                          valores=[fin_neto_a, prev_fin, (fin_neto_a - prev_fin) if prev_fin is not None else None]))
        nodos += [fl("Impuestos", None, costo=True, cols=["impc", "impd"]), fl("Utilidad neta", "neto", "saldo")]
        st.html(E.tabla_meses(nodos, [("Mes " if SOLO else "Acum. ") + f"{MESES[mes - 1]} {str(anio)[2:]}",
                                       f"{MESES[mes - 1]} {str(anio - 1)[2:]}", "Δ vs año ant."],
                               formatos=["m", "m", "m"], barras=[2], compacta=True))
    with der, st.container(key="card_ef_flujo"):
        titulo_tarjeta("Flujo de efectivo", f"Método indirecto · {txt_per} · desde el balance de {nom_b.lower()} · "
                       "verde = entra caja; rojo = sale")
        if flujo is None:
            st.caption(f"No hay balance de {nom_b} para calcular el flujo.")
        else:
            def ff(nombre, valor, clase="linea", nota=""):
                return dict(nombre=nombre, clase=clase, dif_col=0, valores=[valor], nota=nota)
            st.html(E.tabla_meses([
                ff("Utilidad neta", flujo["neto"]),
                ff("(+) Depreciación, amortización y deterioro", flujo["da"]),
                ff("(±) Capital de trabajo operativo", flujo["kt"], nota="Cartera + inventarios − cuentas por pagar"),
                ff("Flujo operativo", flujo["op"], "saldo"),
                ff("(−) Inversión en activos fijos", flujo["capex"],
                   nota="Variación de propiedad, planta, equipo e intangibles + depreciación y amortización"),
                ff("(±) Deuda financiera", flujo["deuda"]),
                ff("(±) Otros (patrimonio, impuestos, otras cuentas)", flujo["otros"],
                   nota="Diferencia para que cuadre con la variación del efectivo del balance"),
                ff("Variación neta de caja", flujo["caja"], "saldo"),
                dict(nombre=f"Caja inicial ({nom_b})", clase="linea", valores=[flujo["ini"]]),
                dict(nombre=f"Caja final ({MESES[mes - 1]} {anio})", clase="saldo", valores=[flujo["fin"]]),
            ], [txt_per.capitalize() if SOLO else f"Ene–{MESES[mes - 1]} {anio}"], barras=[0], compacta=True))

    # ---------------- situación financiera (activo | pasivo + patrimonio)
    with st.container(key="card_ef_sit"):
        titulo_tarjeta("Situación financiera", f"Cierre de {MESES_L[mes - 1]} {anio} vs {nom_b}")
        c1, c2 = st.columns(2, gap="medium")

        def fs(nombre, act, prev, clase="linea", inv=False):
            return dict(nombre=nombre, clase=clase, dif_col=2, dif_inv=inv,
                        valores=[act, prev, (act - prev) if prev is not None else None])

        def ambos(f_):
            return f_(fila), (f_(base) if base is not None else None)
        cxc = ambos(lambda x: v("clientes", x) + v("deudores", x))
        inv_ = ambos(lambda x: v("inventarios", x))
        caja = ambos(lambda x: v("efectivo", x))
        ac = ambos(lambda x: v("AC", x))
        af = ambos(lambda x: v("activo_fijo", x))
        anc = ambos(lambda x: v("ANC", x))
        cols3 = [f"{MESES[mes - 1]} {anio}", nom_b, "Δ"]
        with c1:
            st.html(E.tabla_meses([
                fs("Efectivo", *caja), fs("Cuentas por cobrar", *cxc), fs("Inventarios", *inv_),
                fs("Otros activos corrientes", ac[0] - caja[0] - cxc[0] - inv_[0],
                   (ac[1] - caja[1] - cxc[1] - inv_[1]) if base is not None else None),
                fs("Activo corriente", *ac, "grupo"),
                fs("Propiedad, planta, equipo e intangibles", *af),
                fs("Otros activos no corrientes", anc[0] - af[0], (anc[1] - af[1]) if base is not None else None),
                fs("Activo no corriente", *anc, "grupo"),
                fs("Total activo", *ambos(lambda x: v("ACTIVO", x)), "saldo"),
            ], cols3, formatos=["m"] * 3, compacta=True))
        cxp = ambos(lambda x: v("cxp", x))
        pcorr = ambos(lambda x: v("PC", x))
        with c2:
            st.html(E.tabla_meses([
                fs("Cuentas por pagar", *cxp, inv=True),
                fs("Otros pasivos corrientes", pcorr[0] - cxp[0], (pcorr[1] - cxp[1]) if base is not None else None,
                   inv=True),
                fs("Pasivo corriente", *pcorr, "grupo", inv=True),
                fs("Pasivo no corriente", *ambos(lambda x: v("PNC", x)), "grupo", inv=True),
                fs("Pasivo total", *ambos(lambda x: v("PASIVO", x)), "grupo", inv=True),
                fs("Patrimonio", *ambos(lambda x: v("PAT", x)), "grupo"),
                fs("Total pasivo + patrimonio", *ambos(lambda x: v("PASIVO", x) + v("PAT", x)), "saldo"),
                fs("Memo: deuda financiera (corto + largo plazo)", *ambos(lambda x: v("deuda_fin", x)), inv=True),
            ], cols3, formatos=["m"] * 3, compacta=True))
        st.caption("Δ en verde cuando mejora la posición (más activo o patrimonio, menos pasivo).")

    # ---------------- notas
    with st.container(key="card_ef_notas"):
        titulo_tarjeta("Notas del periodo", "Resumen automático de las cifras clave")
        gasto_dia = (v("cos" + suf) + v("adm" + suf)) / max(1, (1 if SOLO else mes) * 30)
        dias_caja = v("efectivo") / gasto_dia if gasto_dia else None
        dc, dc_a = v("dias_cartera"), (v("dias_cartera", ant) if hay_ant else None)
        notas = [f"<b>Efectivo {E.millones(v('efectivo'))}</b>: cubre unos "
                 + (f"{dias_caja:,.0f}".replace(",", ".") if dias_caja else "—")
                 + " días de costos y gastos al ritmo del periodo.",
                 f"<b>Cuentas por cobrar {E.millones(cxc[0])}</b>: " + f"{dc:,.0f}".replace(",", ".")
                 + " días de cartera"
                 + (f" ({'+' if dc - dc_a >= 0 else ''}" + f"{dc - dc_a:,.0f}".replace(",", ".")
                    + " días vs año anterior)" if dc_a is not None and not pd.isna(dc_a) else "") + ".",
                 f"<b>Deuda financiera {E.millones(v('deuda_fin'))}</b>: deuda neta {E.millones(v('deuda_neta'))}"
                 + (f" · {v('deuda_ebit'):,.2f}x EBIT 12 meses".replace(".", ",") if not pd.isna(fila["deuda_ebit"])
                    else "") + f" · endeudamiento {E.pct_txt(fila['endeudamiento'])}."]
        if flujo is not None and abs(neto) > 1:
            notas.append(f"<b>Utilidad vs caja</b>: utilidad {E.millones(neto)} y flujo operativo "
                         f"{E.millones(flujo['op'])} ({E.pct_txt(flujo['op'] / neto)} de conversión).")
        st.html("<ul style='margin:.2rem 0 .8rem 1.1rem;padding:0;font-size:.84rem;line-height:1.7'>"
                + "".join(f"<li>{n}</li>" for n in notas) + "</ul>")
        avisos = []
        descuadre = v("ACTIVO") - (v("PASIVO") + v("PAT"))
        avisos.append(("✓" if abs(descuadre) <= 1000 else E.millones(descuadre),
                       "Activo = pasivo + patrimonio" if abs(descuadre) <= 1000 else "Descuadre del balance",
                       "ok" if abs(descuadre) <= 1000 else "bad"))
        if dc_a is not None and not pd.isna(dc_a):
            avisos.append((f"{dc - dc_a:+,.0f}".replace(",", "."), "días de cartera vs año anterior",
                           "ok" if dc <= dc_a else "warn"))
        if not pd.isna(fila["razon_corriente"]):
            avisos.append((f"{fila['razon_corriente']:,.2f}x".replace(".", ","), "razón corriente",
                           "ok" if fila["razon_corriente"] >= 1 else "bad"))
        if flujo is not None:
            avisos.append((E.millones(flujo["capex"]), "inversión en activos fijos", "na"))
        st.html(E.avisos(avisos))
