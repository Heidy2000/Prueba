"""Módulo Flujo de caja: ejecutado mensual (calculado desde los movimientos de bancos de SAP), proyectado 4+8 y
detalle por tercero."""
from html import escape

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import estilo as E
from flujo import PENDIENTE
from ui import EMPRESAS, MESES, a_excel, estilo_grafico, selector_periodo, titulo_tarjeta

MAX_TERCEROS = 15
SALDOS_INI = {"saldo inicial", "saldo inicial otros componentes"}
SALDOS_FIN = {"flujo de caja operación", "flujo de caja financiero", "saldo final otros componentes",
              "saldo final caja"}


def _nombre(c: str) -> str:
    return c.rstrip("_").strip()


def _indicadores(serie, total, info, mov, per, meses, ult, operativa, c_ini, c_ing, c_egr, c_fin):
    """Indicadores de caja: generación, liquidez, estructura de egresos y concentración de pagos."""
    ms = [m for m in meses if m <= ult] or meses
    si, se, sf = serie(c_ing), serie(c_egr), serie(c_fin)
    ing, egr = float(si.reindex(ms).sum()), float(se.reindex(ms).sum())
    neto = ing + egr

    def egr_prom(m, n=3):           # egreso mensual promedio de los últimos n meses (hasta m)
        v = [-float(se.get(k, 0.0)) for k in range(max(1, m - n + 1), m + 1)]
        return sum(v) / len(v) if v else 0.0

    def dv(a, b):
        return a / b if b else None

    m_fin = ms[-1]
    saldo = float(sf.get(m_fin, 0.0))
    dias = dv(saldo, egr_prom(m_fin) / 30)
    cobertura = dv(ing, -egr)
    margen = dv(neto, ing)

    # concentración de pagos por tercero (sin traslados entre cuentas propias, anulados ni pendientes)
    sec = info[info["seccion"] == ("Egresos" if operativa else "Otros componentes")]["concepto"]
    de_egreso = {" ".join(str(c).lower().split()) for c in sec if "traslado" not in str(c).lower()}
    pagos = mov[(mov["valor"] < 0) & mov["concepto"].map(lambda c: " ".join(str(c).lower().split())).isin(de_egreso)]
    pagos = pagos[pagos["tercero"] != "(Sin tercero)"]
    por_ter = (-pagos.groupby("tercero")["valor"].sum()).sort_values(ascending=False)
    top10 = dv(float(por_ter.head(10).sum()), float(por_ter.sum()))

    with st.container(key="card_fc_ind"):
        titulo_tarjeta("Indicadores de caja", f"{per.texto.capitalize()} · las razones se calculan sobre el periodo; "
                       "los saldos, al cierre del último mes")
        st.html(E.kpis([
            dict(label="Cobertura de egresos", valor="—" if cobertura is None else f"{cobertura:,.2f} x"
                 .replace(".", ","), sub=(f"Entran ${cobertura:,.2f}".replace(".", ",") + " por cada $1 que sale") if cobertura
                 else "Sin egresos", icono="check", color=E.VERDE if (cobertura or 0) >= 1 else E.ROJO,
                 fondo=E.VERDE_SUAVE if (cobertura or 0) >= 1 else E.ROJO_SUAVE),
            dict(label="Margen de caja", valor=E.pct_txt(margen) if margen is not None else "—",
                 sub=f"Flujo neto {E.millones(neto)} sobre ingresos", icono="piggy",
                 color=E.VERDE if neto >= 0 else E.ROJO, fondo=E.VERDE_SUAVE if neto >= 0 else E.ROJO_SUAVE),
            dict(label="Días de caja", valor="—" if dias is None else f"{dias:,.0f} días".replace(",", "."),
                 sub=f"Saldo {E.millones(saldo)} ÷ egreso diario promedio (últimos 3 meses)", icono="calendar",
                 color=E.AZUL, fondo=E.AZUL_SUAVE),
            dict(label="Concentración de pagos", valor=E.pct_txt(top10) if top10 is not None else "—",
                 sub=(f"Top 10 terceros · el mayor: {escape(str(por_ter.index[0]))[:34]}" if len(por_ter)
                      else "Sin pagos con tercero"), icono="receipt", color=E.AZUL, fondo=E.AZUL_SUAVE),
        ]))

        # tabla mensual
        def fila(nombre, por_mes, del_periodo, fmt, nota=""):
            return dict(nombre=nombre, valores=[del_periodo] + [por_mes(m) for m in ms], clase="linea", fmt=fmt,
                        nota=nota)

        def seccion(nombre, hijos):
            return dict(nombre=nombre, valores=[None] * (len(ms) + 1), clase="seccion", hijos=hijos, fmt="x")

        ing_m = lambda m: float(si.get(m, 0.0))                     # noqa: E731
        egr_m = lambda m: float(se.get(m, 0.0))                     # noqa: E731
        generacion = [
            fila("Flujo neto (ingresos − egresos)", lambda m: ing_m(m) + egr_m(m), neto, "m"),
            fila("Margen de caja", lambda m: dv(ing_m(m) + egr_m(m), ing_m(m)), margen, "p",
                 "Flujo neto sobre ingresos"),
            fila("Cobertura de egresos (veces)", lambda m: dv(ing_m(m), -egr_m(m)), cobertura, "x",
                 "Ingresos ÷ egresos: más de 1 significa que la operación genera caja"),
        ]
        liquidez = [
            fila("Saldo final", lambda m: float(sf.get(m, 0.0)), saldo, "m"),
            fila("Días de caja", lambda m: dv(float(sf.get(m, 0.0)), egr_prom(m) / 30), dias, "n",
                 "Días que alcanza el saldo con el egreso promedio de los últimos 3 meses"),
            fila("Meses de cobertura", lambda m: dv(float(sf.get(m, 0.0)), egr_prom(m)), dv(saldo, egr_prom(m_fin)),
                 "x", "Saldo final ÷ egreso mensual promedio (3 meses)"),
        ]
        if operativa and "Saldo Final Caja" in set(info["concepto"]):
            st_tot = serie("Saldo Final Caja")
            if float(st_tot.reindex(ms).abs().sum()) != float(sf.reindex(ms).abs().sum()):
                liquidez.append(fila("Saldo final caja total (incl. otros componentes)",
                                     lambda m: float(st_tot.get(m, 0.0)), float(st_tot.get(m_fin, 0.0)), "m"))
        if operativa:
            top = info[(info["seccion"] == "Egresos") & (info["padre"] == "") & (info["tipo"] != "total")]
        else:
            top = info[info["padre"] == "Egresos otros componentes"]
        estructura = []
        for r in top.itertuples():
            s_c = serie(r.concepto)
            if "traslado" in r.concepto.lower() or abs(float(s_c.reindex(ms).sum())) < 0.001 * abs(ing or 1):
                continue                     # traslados entre cuentas propias y conceptos sin peso
            estructura.append((float(s_c.reindex(ms).sum()),
                               fila(_nombre(r.concepto), lambda m, s_c=s_c: dv(-float(s_c.get(m, 0.0)), ing_m(m)),
                                    dv(-float(s_c.reindex(ms).sum()), ing), "p")))
        estructura = [f for _, f in sorted(estructura, key=lambda x: x[0])]
        nodos = [seccion("Generación de caja", generacion), seccion("Liquidez", liquidez),
                 seccion("Egresos como % de los ingresos", estructura)]
        if operativa:
            fin = [c for c in ("Total Financiamiento",) if c in set(info["concepto"])]
            socios = [c for c in info[(info["seccion"] == "Socios") & (info["tipo"] == "linea")]["concepto"]
                      if "socio" in c.lower() or "dividend" in c.lower()]
            otros = []
            if fin:
                s_f = serie(fin[0])
                otros.append(fila("Financiamiento neto (préstamos recibidos − pagos)",
                                  lambda m: float(s_f.get(m, 0.0)), float(s_f.reindex(ms).sum()), "m"))
            s_s = sum((serie(c) for c in socios), pd.Series(0.0, index=range(1, 13)))
            if float(s_s.reindex(ms).abs().sum()) >= 1:
                otros.append(fila("Pagos a socios (dividendos y contrato)", lambda m: float(s_s.get(m, 0.0)),
                                  float(s_s.reindex(ms).sum()), "m"))
            flujo_ec = lambda m: ing_m(m) + egr_m(m)                 # noqa: E731
            if fin:
                otros.append(fila("Flujo neto después de financiamiento y socios",
                                  lambda m: flujo_ec(m) + float(s_f.get(m, 0.0)) + float(s_s.get(m, 0.0)),
                                  neto + float(s_f.reindex(ms).sum()) + float(s_s.reindex(ms).sum()), "m"))
            if otros:
                nodos.append(seccion("Financiamiento y socios", otros))
        st.html(E.tabla_meses(nodos, ["Periodo"] + [MESES[m - 1] for m in ms], abierto=True))
        if len(por_ter):
            with st.expander("Terceros a los que más se les pagó en el periodo"):
                t = por_ter.head(10).reset_index()
                t.columns = ["Tercero", "Pagado"]
                t["% de los pagos"] = t["Pagado"] / por_ter.sum()
                st.dataframe(t, hide_index=True, width="stretch", column_config={
                    "Pagado": st.column_config.NumberColumn(format="accounting"),
                    "% de los pagos": st.column_config.NumberColumn(format="percent")})


def mostrar(ctx):
    leer = ctx.leer
    zona_encabezado = st.container()
    hay = set(leer("SELECT DISTINCT empresa FROM flujo_lineas")["empresa"].fillna("ALDC"))
    empresas = list(EMPRESAS)              # siempre se muestran las dos compañías
    anios = leer("SELECT DISTINCT anio FROM flujo_lineas ORDER BY anio DESC")["anio"].tolist()
    if not anios:
        zona_encabezado.html(E.encabezado("Flujo de caja", "Flujo de caja ejecutado", []))
        st.info("Aún no hay flujo de caja cargado. Financiera lo carga en Administración → Flujo de caja.")
        return

    # ------------------------------------------------------------ filtros del informe
    with st.container(key="card_filtros"):
        st.html('<div class="nf-ftit">Filtros del informe</div>')
        if len(empresas) > 1:
            c0, *c = st.columns([0.8, 0.8, 1.55, 0.95, 1.8, 0.6], vertical_alignment="bottom")
            empresa = c0.selectbox("Compañía", empresas, key="fc_empresa",
                                   help=" · ".join(f"{e}: {EMPRESAS[e]}" for e in empresas))
        else:
            c = st.columns([0.8, 1.55, 0.95, 1.8, 1.4], vertical_alignment="bottom")
            empresa = empresas[0] if empresas else list(EMPRESAS)[0]
        if empresa not in hay:
            st.info(f"Aún no hay flujo de caja de **{empresa} · {EMPRESAS[empresa]}**. Financiera lo carga en "
                    "Administración → Flujo de caja, eligiendo esa compañía arriba.", icon=":material/info:")
            zona_encabezado.html(E.encabezado(f"Flujo de caja · {empresa}", EMPRESAS[empresa], []))
            return
        anios = leer("SELECT DISTINCT anio FROM flujo_lineas WHERE empresa=? ORDER BY anio DESC",
                     (empresa,))["anio"].tolist() or anios
        anio = c[0].selectbox("Año", anios, key="fc_anio")
        lin = leer("SELECT * FROM flujo_lineas WHERE anio=? AND empresa=? ORDER BY orden, mes", (anio, empresa))
        ult = lin.loc[(lin["tipo"] == "linea") & (lin["valor"] != 0), "mes"].max()
        ult = 1 if pd.isna(ult) else int(ult)
        per = selector_periodo(c[1], c[2], ult, "fc")
        vista = c[3].segmented_control("Vista", ["Caja operativa", "Otros componentes"], default="Caja operativa",
                                       key="fc_vista") or "Caja operativa"

    proy = leer("SELECT concepto, proyectado, observacion FROM flujo_proyectado WHERE anio=? AND empresa=?",
                (anio, empresa)) \
        .drop_duplicates("concepto").set_index("concepto")
    mov_all = leer("SELECT * FROM flujo_mov WHERE anio=? AND empresa=?", (anio, empresa))
    mov_all["concepto"] = mov_all["concepto"].fillna("").replace("", PENDIENTE)
    pendientes = mov_all[mov_all["concepto"] == PENDIENTE]
    mov = per.filtrar(mov_all)
    meses = [m for m in per.meses()]
    comparable = (not per.solo) and per.mes == ult and not proy.empty

    # tabla ancha concepto x mes (orden original del archivo)
    info = lin.drop_duplicates("orden").set_index("orden")[["seccion", "padre", "concepto", "tipo"]]
    ancho = lin.pivot_table(index="orden", columns="mes", values="valor", aggfunc="sum").reindex(columns=range(1, 13))
    por_nombre = {info.at[o, "concepto"]: o for o in info.index}

    def serie(concepto):
        o = por_nombre.get(concepto)
        return ancho.loc[o] if o is not None else pd.Series(0.0, index=range(1, 13))

    def total(concepto):
        s, k = serie(concepto), concepto.lower()
        if k in SALDOS_INI or k.startswith("saldo inicial"):
            return s.get(meses[0], 0.0)
        if k in SALDOS_FIN or k.startswith("saldo final"):        # ej. saldo de la reserva Santander - FINDETER
            return s.get(meses[-1], 0.0)
        return s.reindex(meses).sum()

    # ------------------------------------------------------------ encabezado
    chips = [(f"{'Mes' if per.solo else 'Corte'}: {MESES[per.mes-1]} {anio}", ""),
             (f"Ejecutado hasta {MESES[ult-1]} {anio}", "ok")]
    if per.mes > ult:
        chips.append((f"Sin datos después de {MESES[ult-1]}", "warn"))
    if not pendientes.empty:
        chips.append((f"{len(pendientes)} movimientos sin clasificar", "warn"))
    zona_encabezado.html(E.encabezado(f"Flujo de caja {anio}" + (f" · {empresa}" if len(empresas) > 1 else ""),
                                      f"Flujo de caja ejecutado · {vista.lower()} · {per.texto} · millones de pesos",
                                      chips))

    operativa = vista == "Caja operativa"
    if operativa:
        c_ini, c_ing, c_egr, c_fin = "Saldo inicial", "Total Ingresos", "Total Egresos", "Flujo de caja operación"
    else:
        c_ini, c_ing, c_egr, c_fin = ("Saldo inicial Otros componentes", "Ingresos otros componentes",
                                      "Egresos otros componentes", "Saldo Final otros componentes")
    ing, egr = total(c_ing), total(c_egr)
    st.html(E.kpis([
        dict(label="Saldo inicial", valor=E.millones(total(c_ini)), sub=f"Al inicio de {MESES[meses[0]-1].lower()}",
             icono="wallet", color=E.AZUL, fondo=E.AZUL_SUAVE),
        dict(label=f"Ingresos {per.suf}", valor=E.millones(ing), sub=per.texto.capitalize(),
             icono="piggy", color=E.VERDE, fondo=E.VERDE_SUAVE),
        dict(label=f"Egresos {per.suf}", valor=E.millones(-egr),
             sub=f"Flujo neto: <b>{E.millones(ing + egr)}</b>", icono="receipt", color=E.ROJO, fondo=E.ROJO_SUAVE),
        dict(label="Saldo final" + (" caja operativa" if operativa else " otros comp."),
             valor=E.millones(total(c_fin)),
             sub=f"Caja total (incl. otros componentes): {E.millones(total('Saldo Final Caja'))}",
             icono="calendar", color=E.AZUL, fondo=E.AZUL_SUAVE),
    ]))

    t1, t2, t3 = st.tabs(["Resumen", "Flujo por concepto", "Movimientos"], on_change="rerun", key="fc_tabs")

    # ============ RESUMEN
    with t1:
        if t1.open:   # solo se calcula la pestaña abierta
            izq, der = st.columns([2.1, 1], gap="medium")
            with izq:
                with st.container(key="card_fc_mes"):
                    titulo_tarjeta("Ingresos, egresos y saldo", "Barras: ingresos y egresos del mes · línea: saldo final · "
                                   "millones de pesos")
                    si, se, sf = serie(c_ing), serie(c_egr), serie(c_fin)
                    vacios = [m for m in range(1, 13) if m > ult]
                    for s in (si, se, sf):
                        s.loc[vacios] = None
                    fig = go.Figure()
                    etiqueta = dict(texttemplate="%{y:,.0f}", textposition="inside", insidetextanchor="end",
                                    textangle=-90, textfont=dict(color="white", size=11), cliponaxis=False)
                    fig.add_bar(x=MESES, y=si / 1e6, name="Ingresos", marker_color=E.VERDE, hovertemplate="%{y:,.1f} M",
                                **etiqueta)
                    fig.add_bar(x=MESES, y=se / 1e6, name="Egresos", marker_color="#e57373", hovertemplate="%{y:,.1f} M",
                                **etiqueta)
                    fig.add_scatter(x=MESES, y=sf / 1e6, name="Saldo final", mode="lines+markers+text",
                                    line=dict(color=E.NAVY, width=2.5), marker=dict(size=6), hovertemplate="%{y:,.1f} M",
                                    texttemplate="<b>%{y:,.0f}</b>", textposition="top center",
                                    textfont=dict(color=E.NAVY, size=11), cliponaxis=False)
                    if per.solo:
                        fig.add_vrect(x0=per.mes - 1.5, x1=per.mes - 0.5, fillcolor=E.AZUL, opacity=.07, line_width=0)
                    fig.update_layout(barmode="relative", bargap=.3, uniformtext=dict(minsize=9, mode="hide"),
                                      yaxis=dict(tickformat=",.0f", ticksuffix=" M"))
                    fig.add_hline(y=0, line_color=E.BORDE)
                    st.plotly_chart(estilo_grafico(fig, 400), width="stretch", config={"displayModeBar": False})
            with der:
                with st.container(key="card_fc_comp"):
                    titulo_tarjeta("Composición de egresos", per.texto.capitalize())
                    if operativa:
                        top = info[(info["seccion"] == "Egresos") & (info["padre"] == "") & (info["tipo"] != "total")]
                    else:
                        top = info[info["padre"] == "Egresos otros componentes"]
                    filas = [(_nombre(r.concepto), -total(r.concepto)) for r in top.itertuples()]
                    filas = sorted([f for f in filas if f[1] > 0], key=lambda x: -x[1])[:9]
                    st.html(E.lista_participacion(filas, color="#e57373", total=-egr))

            _indicadores(serie, total, info, mov, per, meses, ult, operativa, c_ini, c_ing, c_egr, c_fin)

            if operativa and not proy.empty:
                with st.container(key="card_fc_var"):
                    titulo_tarjeta("Ejecutado vs proyectado 4+8",
                                   f"Principales diferencias a {MESES[ult-1].lower()} · con las observaciones del archivo")
                    if not comparable:
                        st.caption(f"El proyectado 4+8 del archivo está calculado a {MESES[ult-1].lower()}. "
                                   f"Selecciona periodo Acumulado hasta {MESES[ult-1].lower()} para compararlo.")
                    else:
                        cand = info[(info["tipo"].isin(["linea", "grupo"])) &
                                    (info["seccion"].isin(["Ingresos", "Egresos", "Financiamiento", "Socios"]))]
                        filas = []
                        for r in cand.itertuples():
                            if r.concepto not in proy.index or r.padre:
                                continue
                            ej, pr = total(r.concepto), proy.at[r.concepto, "proyectado"]
                            obs = proy.at[r.concepto, "observacion"] or ""
                            # si la observación está en un hijo, súmala
                            if r.tipo == "grupo":
                                hijos = info[info["padre"] == r.concepto]["concepto"]
                                obs = " · ".join(filter(None, [obs] + [proy.at[h, "observacion"] for h in hijos
                                                                       if h in proy.index and proy.at[h, "observacion"]]))
                            filas.append((r.concepto, r.seccion, ej, pr, ej - pr, obs))
                        filas = sorted(filas, key=lambda x: -abs(x[4]))[:10]
                        cab = ('<div class="nf-dev nf-head" style="grid-template-columns:minmax(0,3fr) repeat(3,minmax(0,1fr))">'
                               '<div>Concepto</div><div>Ejecutado</div><div>Proyectado 4+8</div><div>Diferencia</div></div>')
                        cuerpo = "".join(
                            f'<div class="nf-dev" style="grid-template-columns:minmax(0,3fr) repeat(3,minmax(0,1fr))">'
                            f'<div><b>{escape(_nombre(n))}</b> <small style="color:{E.TEXTO_2}">· {escape(s)}</small>'
                            + (f'<br><small style="color:{E.TEXTO_2}">{escape(o)}</small>' if o else "")
                            + f'</div><div>{E.millones(ej)}</div><div>{E.millones(pr)}</div>'
                            f'<div class="{"nf-pos" if d > 0 else "nf-neg"}" style="font-weight:700">{E.millones(d)}</div></div>'
                            for n, s, ej, pr, d, o in filas)
                        st.html(f'<div class="nf-t notranslate" translate="no" style="border:1px solid {E.BORDE};'
                                f'border-radius:12px;overflow:hidden">{cab}{cuerpo}</div>')
                        st.caption("Diferencia = ejecutado − proyectado. En verde, a favor de la caja (más ingreso o menos "
                                   "egreso de lo proyectado); en rojo, en contra.")

    # ============ FLUJO POR CONCEPTO
    with t2:
        if t2.open:   # solo se calcula la pestaña abierta
            with st.container(key="card_fc_tabla"):
                titulo_tarjeta("Flujo por concepto",
                               "Haz clic en una sección o grupo para desplegarlo, y en un concepto para ver los terceros "
                               "(movimientos de bancos de SAP).")
                columnas = ["Total periodo"] + (["Proyectado 4+8", "Diferencia"] if comparable else []) + \
                    [MESES[m - 1] for m in meses]
                por_ter = {}
                if not mov.empty:
                    g = mov.groupby(["concepto", "tercero", "mes"])["valor"].sum().unstack(fill_value=0)
                    for concepto, sub in g.groupby(level=0):
                        por_ter[concepto] = sub.droplevel(0)

                def valores(concepto):
                    s = serie(concepto)
                    tot = float(total(concepto))
                    vals = [tot]
                    if comparable:
                        pr = proy.at[concepto, "proyectado"] if concepto in proy.index else None
                        vals += [pr, (tot - pr) if pr is not None else None]
                    return vals + [float(s.get(m, 0.0)) for m in meses]

                def nodo_terceros(concepto):
                    t = por_ter.get(concepto)
                    if t is None or t.empty:
                        return []
                    t = t.reindex(columns=range(1, 13), fill_value=0)
                    t["_tot"] = t[meses].sum(axis=1)
                    t = t.reindex(t["_tot"].abs().sort_values(ascending=False).index)
                    out = []
                    if len(t) > MAX_TERCEROS:
                        resto = t.iloc[MAX_TERCEROS:].sum()
                        t = t.iloc[:MAX_TERCEROS]
                        t.loc[f"Otros {len(por_ter[concepto]) - MAX_TERCEROS} terceros"] = resto
                    for ter, r in t.iterrows():
                        vals = [float(r["_tot"])] + ([None, None] if comparable else []) + [float(r[m]) for m in meses]
                        out.append(dict(nombre=str(ter), valores=vals, clase="tercero"))
                    return out

                def nodo(concepto, clase="linea", nombre=None, hijos=None):
                    de_terceros = hijos is None
                    h = nodo_terceros(concepto) if de_terceros else hijos
                    n_ter = len(por_ter.get(concepto, [])) if de_terceros else 0
                    return dict(nombre=nombre or _nombre(concepto), valores=valores(concepto), clase=clase, hijos=h,
                                dif_col=2 if comparable else None,
                                sub=f"{n_ter} tercero{'s' if n_ter != 1 else ''}" if n_ter else "")

                def lineas_de(seccion=None, padre=""):
                    q = info[(info["tipo"] != "total") & (info["padre"] == padre)]
                    if seccion:
                        q = q[q["seccion"] == seccion]
                    out = []
                    for r in q.itertuples():
                        if r.tipo == "grupo":
                            out.append(nodo(r.concepto, "grupo", hijos=lineas_de(padre=r.concepto)))
                        else:
                            out.append(nodo(r.concepto))
                    return out

                def seccion(nombre, concepto_total, hijos, clase="seccion"):
                    n = nodo(concepto_total, clase, nombre=nombre, hijos=hijos)
                    return n

                socios = info[info["seccion"] == "Socios"]
                socios_lineas = socios[socios["tipo"] == "linea"]["concepto"].tolist()
                s_socios = sum((serie(c) for c in socios_lineas), pd.Series(0.0, index=range(1, 13)))
                nodo_socios = dict(nombre="Socios (dividendos)", clase="seccion", hijos=lineas_de("Socios"),
                                   valores=[float(s_socios.reindex(meses).sum())] + ([None, None] if comparable else [])
                                   + [float(s_socios.get(m, 0)) for m in meses])
                otros = [nodo("Saldo inicial Otros componentes", "saldo", hijos=[])] + lineas_de("Otros componentes") + \
                        [nodo("Saldo Final otros componentes", "saldo", hijos=[])]
                nodos = [
                    nodo("Saldo inicial", "saldo", hijos=[]),
                    seccion("Ingresos", "Total Ingresos", lineas_de("Ingresos")),
                    seccion("Egresos", "Total Egresos", lineas_de("Egresos")),
                    nodo("Flujo de caja económico", "saldo", hijos=[]),
                    seccion("Financiamiento", "Total Financiamiento", lineas_de("Financiamiento")),
                    nodo_socios,
                    nodo("Flujo de caja operación", "saldo", nombre="Saldo final caja operativa", hijos=[]),
                    seccion("Otros componentes (fiducia)", "Saldo Final otros componentes", otros),
                    nodo("Saldo Final Caja", "saldo", nombre="Saldo final caja total", hijos=[]),
                ]
                abrir = st.toggle("Desplegar todo", key="fc_abrir")
                st.html(E.tabla_meses(nodos, columnas, abierto=abrir))
                st.caption("Saldos: el total del periodo muestra el saldo al inicio (saldos iniciales) o al cierre (saldos "
                           "finales). Algunas líneas incluyen ajustes manuales (Administración → Flujo de caja) que no "
                           "son movimientos de SAP; por eso la suma de terceros puede diferir del concepto.")

                plano = lin[lin["mes"].isin(meses)].pivot_table(index=["orden", "seccion", "padre", "concepto"],
                                                                columns="mes", values="valor", aggfunc="sum")
                plano.columns = [MESES[m - 1] for m in plano.columns]
                st.download_button("Descargar flujo", lambda d=plano.reset_index().drop(columns="orden"): a_excel(d),
                                   f"flujo_de_caja_{anio}.xlsx", icon=":material/download:")

    # ============ MOVIMIENTOS
    with t3:
        if t3.open:   # solo se calcula la pestaña abierta
            with st.container(key="card_fc_movs"):
                titulo_tarjeta("Movimientos de caja", f"Movimientos de bancos de SAP · {per.texto}")
                if mov.empty:
                    st.info("No hay movimientos de caja para este periodo.")
                else:
                    c1, c2 = st.columns([1.4, 1.6])
                    conceptos = c1.multiselect("Concepto", sorted(mov["concepto"].unique()), placeholder="Todos",
                                               key="fc_conc")
                    buscar = c2.text_input("Buscar", placeholder="Tercero, texto, documento…", key="fc_buscar")
                    m = mov[mov["concepto"].isin(conceptos)] if conceptos else mov
                    if buscar:
                        m = m[m[["tercero", "texto", "documento", "cuenta"]].astype(str)
                              .apply(lambda s: s.str.contains(buscar, case=False, regex=False)).any(axis=1)]
                    m = m.sort_values("fecha", ascending=False)[["fecha", "concepto", "tercero", "texto", "documento",
                                                                 "cuenta", "valor"]]
                    if not pendientes.empty:
                        st.warning(f"{len(pendientes)} movimientos de {anio} aún no tienen concepto de caja (por "
                                   f"{E.millones(pendientes['valor'].abs().sum())}); no suman en el flujo hasta que "
                                   "Financiera los clasifique en Administración → Flujo de caja.",
                                   icon=":material/pending_actions:")
                    st.html(f'<p class="nf-card-s">{len(m):,} movimientos · neto '.replace(",", ".")
                            + f'<b style="color:{E.TEXTO}">{E.pesos(m["valor"].sum())}</b></p>')
                    st.dataframe(m, hide_index=True, width="stretch", height=520, column_config={
                        "fecha": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"), "concepto": "Concepto",
                        "tercero": "Tercero", "texto": st.column_config.TextColumn("Texto", width="large"),
                        "documento": "Documento", "cuenta": "Cuenta",
                        "valor": st.column_config.NumberColumn("Valor", format="accounting")})
                    st.download_button("Descargar movimientos", lambda d=m: a_excel(d), f"movimientos_caja_{anio}.xlsx",
                                       icon=":material/download:")
