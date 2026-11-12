"""Módulo Nómina: presupuesto de personal por cargo vs ejecución de la nómina de SAP (costo y personas)."""
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import estilo as E
import nomina
from ui import MESES, MESES_L, a_excel, estilo_grafico, selector_periodo, titulo_tarjeta

SIN_AREA = "SIN ÁREA"
def _excel_hojas(hojas: dict) -> bytes:
    from io import BytesIO
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for nombre, df in hojas.items():
            df.to_excel(w, sheet_name=nombre[:31], index=False)
    return buf.getvalue()


def preparar(leer, anio: int):
    """Presupuesto y ejecución del año con área y cargo común. También el conteo de personas reales."""
    ppto = leer("SELECT * FROM nomina_ppto WHERE anio=?", (anio,))
    ejec = leer("SELECT * FROM nomina_ejec WHERE anio=?", (anio,))
    ceco = dict(leer("SELECT cc, area FROM nomina_ceco").values.tolist())
    ejec["area"] = ejec["cc"].map(ceco).fillna(SIN_AREA) if not ejec.empty else []
    ppto["area"] = ppto["area"].replace("", SIN_AREA) if not ppto.empty else []
    cp, ce = nomina.homologar(leer, ppto, ejec)
    ppto["comun"], ejec["comun"] = cp, ce
    # personas reales: cada empleado se cuenta una vez por mes, en el área y cargo donde tiene más costo
    if not ejec.empty:
        e = ejec[ejec["empleado_id"] != ""]
        base = e.groupby(["mes", "empleado_id", "area", "comun", "cargo"], as_index=False)["valor"].sum()
        base = base.sort_values("valor", ascending=False).drop_duplicates(["mes", "empleado_id"])
        pers = base.groupby(["mes", "area", "comun", "cargo"]).size().rename("personas").reset_index()
    else:
        pers = pd.DataFrame(columns=["mes", "area", "comun", "cargo", "personas"])
    return ppto, ejec, pers


def mostrar(ctx):
    leer = ctx.leer
    zona_encabezado = st.container()
    anios = sorted(set(leer("SELECT DISTINCT anio FROM nomina_ppto")["anio"])
                   | set(leer("SELECT DISTINCT anio FROM nomina_ejec")["anio"]), reverse=True)
    if not anios:
        zona_encabezado.html(E.encabezado("Nómina", "Presupuesto de personal vs nómina ejecutada", []))
        st.info("Aún no hay información de nómina. Financiera la carga en Administración → Nómina.")
        return

    # ------------------------------------------------------------ filtros del informe
    with st.container(key="card_filtros"):
        st.html('<div class="nf-ftit">Filtros del informe</div>')
        c = st.columns([0.7, 1.4, 0.95, 2.2], vertical_alignment="bottom")
        con_ejec = leer("SELECT DISTINCT anio FROM nomina_ejec")["anio"].tolist()
        anio = c[0].selectbox("Año", anios, index=anios.index(max(con_ejec)) if con_ejec else 0, key="nm_anio")
        ppto, ejec, pers = preparar(leer, int(anio))
        meses_ej = sorted(ejec["mes"].unique()) if not ejec.empty else []
        meses_pp = sorted(ppto["mes"].unique()) if not ppto.empty else [1]
        ult = int(max(meses_ej)) if meses_ej else int(max(meses_pp))
        per = selector_periodo(c[1], c[2], ult, "nm")
        areas = sorted(set(ppto["area"]) | set(ejec["area"]))
        areas_f = c[3].multiselect("Área", areas, placeholder="Todas las áreas", key="nm_area")
        c2 = st.columns([2.2, 1.3], vertical_alignment="bottom")
        if areas_f:
            ppto, ejec = ppto[ppto["area"].isin(areas_f)], ejec[ejec["area"].isin(areas_f)]
            pers = pers[pers["area"].isin(areas_f)]
        cargos = sorted(set(ppto["comun"]) | set(ejec["comun"]))
        cargos_f = c2[0].multiselect("Cargo", cargos, placeholder="Todos los cargos", key="nm_cargo")
        if cargos_f:
            ppto, ejec = ppto[ppto["comun"].isin(cargos_f)], ejec[ejec["comun"].isin(cargos_f)]
            pers = pers[pers["comun"].isin(cargos_f)]
        # meses comparables: en acumulado solo los meses que tienen ejecución (el presupuesto puede empezar luego)
        meses_per = [m for m in per.meses() if (m in meses_ej or not meses_ej)]
        hay_ppto_periodo = bool(set(meses_per) & set(meses_pp))
        if per.solo is False and meses_ej:
            sin_ppto_meses = [m for m in meses_per if m not in meses_pp]
            if sin_ppto_meses:
                c2[1].caption("Sin presupuesto cargado para: " + ", ".join(MESES[m - 1] for m in sin_ppto_meses)
                              + ". Se comparan solo los meses con presupuesto y ejecución.")
                meses_per = [m for m in meses_per if m in meses_pp]

    pp_per = ppto[ppto["mes"].isin(meses_per)]
    ej_per = ejec[ejec["mes"].isin(meses_per)]
    mes_pers = per.mes                           # las personas se miden en el mes de corte
    pp_mes = ppto[ppto["mes"] == mes_pers]
    re_mes = pers[pers["mes"] == mes_pers]
    periodo_txt = (f"solo {per.mes_txt}" if per.solo
                   else ("meses con datos hasta " + per.mes_txt if meses_per else per.texto))

    chips = [(f"{'Mes' if per.solo else 'Corte'}: {MESES[per.mes - 1]} {anio}", "")]
    if meses_ej:
        chips.append((f"Nómina SAP: {', '.join(MESES[m - 1] for m in meses_ej)}", "ok"))
    sin = ej_per.loc[ej_per["comun"] == nomina.SIN_PPTO, "valor"].sum()
    if sin:
        chips.append((f"Sin homologar: {E.millones(sin)}", "warn"))
    zona_encabezado.html(E.encabezado(f"Nómina {anio}", "Presupuesto de personal por cargo vs nómina de SAP · "
                                      "cifras en millones de pesos", chips))
    if not hay_ppto_periodo and not ej_per.empty:
        st.warning("No hay presupuesto de nómina para el periodo elegido.")

    t1, t2, t3, t4 = st.tabs(["Costo por cargo", "Personas por cargo", "Por concepto", "Sin homologar"],
                             on_change="rerun", key="nm_tabs")

    # ============ COSTO POR CARGO
    with t1:
        if t1.open:
            pa, ea = pp_per["valor"].sum(), ej_per["valor"].sum()
            var = ea - pa                      # negativo: la nómina no llega al presupuesto
            pct = ea / pa if pa else float("nan")
            color = E.color_barra(pct, ea, False)
            st.html(E.kpis([
                dict(label="Presupuesto", valor=E.millones(pa), sub=periodo_txt.capitalize(), icono="wallet",
                     color=E.AZUL, fondo=E.AZUL_SUAVE),
                dict(label="Nómina ejecutada", valor=E.millones(ea),
                     sub=f'{ej_per["empleado_id"].nunique():,} personas con pago'.replace(",", "."),
                     icono="receipt", color=E.AZUL, fondo=E.AZUL_SUAVE),
                dict(label="Variación", valor=f'<span style="color:{E.VERDE if var <= 0 else E.ROJO}">'
                     f'{"+" if var >= 0 else ""}{E.millones(var)}</span>',
                     sub="Ejecutado − presupuesto (negativo: por debajo)", icono="check" if var <= 0 else "alert",
                     color=E.VERDE if var <= 0 else E.ROJO, fondo=E.VERDE_SUAVE if var <= 0 else E.ROJO_SUAVE),
                dict(label="% de ejecución", valor=f'<span style="color:{color}">{E.pct_txt(pct)}</span>',
                     sub="Del presupuesto del periodo", icono="calendar", color=color, fondo=E.AZUL_SUAVE),
            ]))
            with st.container(key="card_nm_costo"):
                titulo_tarjeta("Variación por cargo", f"Presupuesto vs nómina ejecutada · {periodo_txt}. Haz clic "
                               "en un área para ver sus cargos y en un cargo para ver cómo viene en SAP.")
                pa_c = pp_per.groupby(["area", "comun"])["valor"].sum()
                ea_c = ej_per.groupby(["area", "comun"])["valor"].sum()
                ea_s = ej_per.groupby(["area", "comun", "cargo"])["valor"].sum()
                t = pd.DataFrame({"p": pa_c, "e": ea_c}).fillna(0)

                def v(p, e):
                    return [p, e, e - p, (e / p if p else None)]
                nodos = []
                for area in sorted(t.index.get_level_values(0).unique(), key=lambda a: -t.loc[a, "p"].sum()):
                    ta = t.loc[area].sort_values("p", ascending=False)
                    hijos = []
                    for comun, r in ta.iterrows():
                        nietos = []
                        if (area, comun) in ea_s.index.droplevel(2):
                            for cargo, val in ea_s.loc[(area, comun)].sort_values(ascending=False).items():
                                nietos.append(dict(nombre=f"SAP: {cargo}", clase="tercero",
                                                   valores=[None, val, None, None]))
                        hijos.append(dict(nombre=comun, clase="linea", dif_col=2, dif_inv=True, valores=v(r.p, r.e),
                                          hijos=nietos if len(nietos) > 1 or (nietos and nietos[0]["nombre"]
                                                                               != f"SAP: {comun}") else []))
                    nodos.append(dict(nombre=area, clase="grupo", dif_col=2, dif_inv=True, sub=f"{len(hijos)} cargos",
                                      valores=v(ta["p"].sum(), ta["e"].sum()), hijos=hijos))
                nodos.append(dict(nombre="Total", clase="saldo", dif_col=2, dif_inv=True,
                                  valores=v(t["p"].sum(), t["e"].sum())))
                abrir = st.toggle("Desplegar todo", key="nm_abrir1")
                st.html(E.tabla_meses(nodos, ["Presupuesto", "Ejecutado", "Variación", "% ejecución"],
                                      abierto=abrir, formatos=["m", "m", "m", "p"]))
                st.caption("Variación = ejecutado − presupuesto: negativa cuando la nómina no llega al presupuesto "
                           "(en verde, porque es un costo por debajo de lo previsto) y positiva cuando lo supera (en "
                           "rojo). Debajo de cada cargo se ve con qué nombre(s) viene en SAP.")
                plano = t.reset_index().rename(columns={"area": "Área", "comun": "Cargo", "p": "Presupuesto",
                                                        "e": "Ejecutado"})
                plano["Variación"] = plano["Ejecutado"] - plano["Presupuesto"]
                st.download_button("Descargar variación por cargo", lambda d=plano: a_excel(d),
                                   f"nomina_variacion_cargo_{anio}.xlsx", icon=":material/download:")

            with st.container(key="card_nm_mes"):
                titulo_tarjeta("Mes a mes", "Presupuesto vs nómina ejecutada (millones)")
                pm = ppto.groupby("mes")["valor"].sum().reindex(range(1, 13)) / 1e6
                em = ejec.groupby("mes")["valor"].sum().reindex(range(1, 13)) / 1e6
                fig = go.Figure()
                fig.add_bar(x=MESES, y=pm, name="Presupuesto", marker_color="#c9d4e4", hovertemplate="%{y:,.1f} M")
                fig.add_bar(x=MESES, y=em, name="Ejecutado", marker_color=E.AZUL, hovertemplate="%{y:,.1f} M")
                fig.update_layout(barmode="group", bargap=.25, yaxis=dict(tickformat=",.0f", ticksuffix=" M"))
                st.plotly_chart(estilo_grafico(fig, 320), width="stretch", config={"displayModeBar": False})

    # ============ PERSONAS POR CARGO
    with t2:
        if t2.open:
            pp_n, re_n = pp_mes["personas"].sum(), re_mes["personas"].sum()
            dif = re_n - pp_n                  # positivo: hay más personas que las presupuestadas
            st.html(E.kpis([
                dict(label=f"Personas presupuestadas · {MESES[mes_pers - 1]}", valor=f"{pp_n:,.0f}".replace(",", "."),
                     sub="Según el presupuesto por cargo", icono="wallet", color=E.AZUL, fondo=E.AZUL_SUAVE),
                dict(label=f"Personas en nómina · {MESES[mes_pers - 1]}", valor=f"{re_n:,.0f}".replace(",", "."),
                     sub="Empleados distintos con pago en SAP", icono="receipt", color=E.AZUL, fondo=E.AZUL_SUAVE),
                dict(label="Diferencia", valor=f'{"+" if dif > 0 else ""}{dif:,.0f}'.replace(",", "."),
                     sub="En nómina − presupuestadas", icono="alert" if dif else "check",
                     color=E.AZUL, fondo=E.AZUL_SUAVE),
                dict(label="Costo promedio por persona", valor=E.millones(
                    ej_per[ej_per["mes"] == mes_pers]["valor"].sum() / re_n if re_n else float("nan")),
                     sub=f"Ppto: {E.millones(pp_mes['valor'].sum() / pp_n if pp_n else float('nan'))}",
                     icono="calendar", color=E.AZUL, fondo=E.AZUL_SUAVE),
            ]))
            with st.container(key="card_nm_pers"):
                titulo_tarjeta("Variación en cantidad de personas",
                               f"Personas presupuestadas vs personas en la nómina de {MESES_L[mes_pers - 1]} "
                               f"{anio}, por área y cargo.")
                pn = pp_mes.groupby(["area", "comun"])["personas"].sum()
                rn = re_mes.groupby(["area", "comun"])["personas"].sum()
                rs = re_mes.groupby(["area", "comun", "cargo"])["personas"].sum()
                cpp = pp_mes.groupby(["area", "comun"])["valor"].sum()
                cre = ej_per[ej_per["mes"] == mes_pers].groupby(["area", "comun"])["valor"].sum()
                t = pd.DataFrame({"p": pn, "r": rn, "cp": cpp, "cr": cre}).fillna(0)
                t = t[(t["p"] != 0) | (t["r"] != 0)]

                def vp(p, r, cp, cr):
                    return [p, r, r - p, (cp / p if p else None), (cr / r if r else None)]
                nodos = []
                for area in sorted(t.index.get_level_values(0).unique(), key=lambda a: -t.loc[a, "p"].sum()):
                    ta = t.loc[area].sort_values("p", ascending=False)
                    hijos = []
                    for comun, r in ta.iterrows():
                        nietos = []
                        if (area, comun) in rs.index.droplevel(2):
                            for cargo, n in rs.loc[(area, comun)].sort_values(ascending=False).items():
                                nietos.append(dict(nombre=f"SAP: {cargo}", clase="tercero",
                                                   valores=[None, n, None, None, None]))
                        hijos.append(dict(nombre=comun, clase="linea", valores=vp(r.p, r.r, r.cp, r.cr),
                                          hijos=nietos if len(nietos) > 1 else []))
                    s = ta.sum()
                    nodos.append(dict(nombre=area, clase="grupo", sub=f"{len(hijos)} cargos",
                                      valores=vp(s.p, s.r, s.cp, s.cr), hijos=hijos))
                s = t.sum()
                nodos.append(dict(nombre="Total", clase="saldo", valores=vp(s.p, s.r, s.cp, s.cr)))
                abrir = st.toggle("Desplegar todo", key="nm_abrir2")
                st.html(E.tabla_meses(nodos, ["Personas ppto", "Personas nómina", "Diferencia",
                                              "Costo/persona ppto", "Costo/persona real"],
                                      abierto=abrir, formatos=["n", "n", "n", "m", "m"]))
                st.caption("Diferencia = en nómina − presupuestadas: positiva cuando hay más personas que las "
                           "presupuestadas; negativa cuando hay menos (vacantes). Cada empleado se cuenta una vez, en "
                           "el cargo y área donde tiene más costo en el mes.")
                plano = t.reset_index().rename(columns={"area": "Área", "comun": "Cargo", "p": "Personas ppto",
                                                        "r": "Personas nómina", "cp": "Costo ppto mes",
                                                        "cr": "Costo real mes"})
                plano["Diferencia"] = plano["Personas nómina"] - plano["Personas ppto"]
                st.download_button("Descargar personas por cargo", lambda d=plano: a_excel(d),
                                   f"nomina_personas_cargo_{anio}_{MESES[mes_pers - 1]}.xlsx",
                                   icon=":material/download:")

            with st.container(key="card_nm_pers_mes"):
                titulo_tarjeta("Personas mes a mes", "Presupuestadas vs en nómina")
                pmn = ppto.groupby("mes")["personas"].sum().reindex(range(1, 13))
                rmn = pers.groupby("mes")["personas"].sum().reindex(range(1, 13))
                fig = go.Figure()
                fig.add_scatter(x=MESES, y=pmn, name="Presupuestadas", line=dict(color="#8a9bb5", dash="dot"),
                                mode="lines+markers")
                fig.add_scatter(x=MESES, y=rmn, name="En nómina", line=dict(color=E.NAVY, width=2.5),
                                mode="lines+markers")
                fig.update_layout(yaxis=dict(tickformat=",.0f"))
                st.plotly_chart(estilo_grafico(fig, 300), width="stretch", config={"displayModeBar": False})

    # ============ POR CONCEPTO
    with t3:
        if t3.open:
            con_detalle = all(c in pp_per.columns for c in nomina.COLS_CONCEPTO) and \
                (pp_per.empty or pp_per[nomina.COLS_CONCEPTO].notna().any().any())
            nombres = [n for _, n in nomina.CONCEPTOS]
            if con_detalle:
                pc = pd.Series({n: pp_per[c].fillna(0).sum() for c, n in nomina.CONCEPTOS})
            else:   # presupuesto cargado con una versión anterior: no trae el reparto por concepto
                pc = pd.Series({n: 0.0 for n in nombres})
                pc["Otros conceptos"] = pp_per["valor"].sum()
            e = ej_per.assign(cc_=ej_per["concepto"].map(nomina.concepto_sap)) if not ej_per.empty else \
                ej_per.assign(cc_=[])
            ec = e.groupby("cc_")["valor"].sum().reindex(nombres).fillna(0)
            t = pd.DataFrame({"p": pc, "e": ec}).fillna(0)
            t = t[(t["p"].abs() > 0.5) | (t["e"].abs() > 0.5)]
            t["v"] = t["e"] - t["p"]
            pa, ea = t["p"].sum(), t["e"].sum()

            if not con_detalle and not pp_per.empty:
                st.warning("El presupuesto de nómina guardado no trae el reparto por concepto (se cargó con una "
                           "versión anterior), por eso todo aparece en 'Otros conceptos'. Financiera debe volver a "
                           "subir el archivo «Presupuesto detalle nomina» en Administración → Nómina → Cargar archivos "
                           "y pulsar «Cargar presupuesto de nómina».", icon=":material/warning:")
            # mayores variaciones (a favor y en contra)
            arriba = t[t["v"] > 0].sort_values("v", ascending=False)
            abajo = t[t["v"] < 0].sort_values("v")
            st.html(E.kpis([
                dict(label="Presupuesto", valor=E.millones(pa), sub=periodo_txt.capitalize(), icono="wallet",
                     color=E.AZUL, fondo=E.AZUL_SUAVE),
                dict(label="Nómina ejecutada", valor=E.millones(ea), sub=f"Variación {'+' if ea - pa >= 0 else ''}"
                     f"{E.millones(ea - pa)} ({E.pct_txt(ea / pa if pa else float('nan'))})", icono="receipt",
                     color=E.VERDE if ea <= pa else E.ROJO, fondo=E.VERDE_SUAVE if ea <= pa else E.ROJO_SUAVE),
                dict(label="Mayor sobre-ejecución", valor=(f'+{E.millones(arriba["v"].iloc[0])}'
                                                          if not arriba.empty else "—"),
                     sub=arriba.index[0] if not arriba.empty else "Ningún concepto supera el presupuesto",
                     icono="alert", color=E.ROJO, fondo=E.ROJO_SUAVE),
                dict(label="Mayor ahorro", valor=E.millones(abajo["v"].iloc[0]) if not abajo.empty else "—",
                     sub=abajo.index[0] if not abajo.empty else "Ningún concepto por debajo",
                     icono="check", color=E.VERDE, fondo=E.VERDE_SUAVE),
            ]))

            with st.container(key="card_nm_conc"):
                titulo_tarjeta("Presupuesto vs nómina por concepto",
                               f"{periodo_txt.capitalize()}. El presupuesto se reparte con la misma fórmula del "
                               "archivo (salario, horas extra, auxilios y porcentajes de seguridad social, vacaciones "
                               "y prestaciones). Haz clic en un concepto para ver cómo viene en SAP.")
                det = e.groupby(["cc_", "concepto"])["valor"].sum() if not e.empty else pd.Series(dtype=float)

                def v(p, x):
                    return [p, x, x - p, (x / p if p else None)]
                nodos = []
                for nombre, r in t.iterrows():
                    hijos = []
                    if nombre in det.index.get_level_values(0):
                        for conc, val in det.loc[nombre].sort_values(ascending=False).items():
                            if abs(val) > 0.5:
                                hijos.append(dict(nombre=f"SAP: {conc}", clase="tercero",
                                                  valores=[None, val, None, None]))
                    nodos.append(dict(nombre=nombre, clase="grupo", dif_col=2, dif_inv=True, valores=v(r.p, r.e),
                                      sub=f"{len(hijos)} conceptos SAP" if hijos else "", hijos=hijos))
                nodos.append(dict(nombre="Total", clase="saldo", dif_col=2, dif_inv=True, valores=v(pa, ea)))
                abrir = st.toggle("Desplegar todo", key="nm_abrir3")
                st.html(E.tabla_meses(nodos, ["Presupuesto", "Ejecutado", "Variación", "% ejecución"],
                                      abierto=abrir, formatos=["m", "m", "m", "p"]))
                st.caption("Variación = ejecutado − presupuesto: negativa (verde) cuando se pagó menos de lo previsto. "
                           "Sueldos y salarios incluye en SAP los días de vacaciones disfrutadas, incapacidades, "
                           "licencias y el apoyo de sostenimiento de aprendices, porque reemplazan días de salario; "
                           "Vacaciones es la provisión del mes.")

            with st.container(key="card_nm_conc_g"):
                titulo_tarjeta("Qué explica la variación", "Variación por concepto, de la mayor sobre-ejecución al mayor "
                               "ahorro (millones)")
                g = t.sort_values("v")
                fig = go.Figure(go.Bar(
                    y=g.index, x=g["v"] / 1e6, orientation="h",
                    marker_color=[E.ROJO if x > 0 else E.VERDE for x in g["v"]],
                    text=[f"{'+' if x > 0 else ''}{x / 1e6:,.1f}" for x in g["v"]], textposition="outside",
                    hovertemplate="%{y}: %{x:,.1f} M<extra></extra>"))
                fig.update_layout(xaxis=dict(tickformat=",.0f", ticksuffix=" M", zeroline=True,
                                             zerolinecolor="#9aa7b8"), yaxis=dict(automargin=True))
                st.plotly_chart(estilo_grafico(fig, 60 + 38 * len(g)), width="stretch",
                                config={"displayModeBar": False})

                plano = t.reset_index().rename(columns={"index": "Concepto", "p": "Presupuesto", "e": "Ejecutado",
                                                        "v": "Variación"})
                plano["% ejecución"] = plano["Ejecutado"] / plano["Presupuesto"].where(plano["Presupuesto"] != 0)
                det_x = det.reset_index().rename(columns={"cc_": "Concepto", "concepto": "Concepto SAP",
                                                          "valor": "Ejecutado"}) if not e.empty else pd.DataFrame()
                st.download_button("Descargar variación por concepto",
                                   lambda a=plano, b=det_x: a_excel(a) if b.empty else _excel_hojas(
                                       {"Por concepto": a, "Detalle SAP": b}),
                                   f"nomina_variacion_concepto_{anio}.xlsx", icon=":material/download:")

    # ============ SIN HOMOLOGAR
    with t4:
        if t4.open:
            with st.container(key="card_nm_sin"):
                titulo_tarjeta("Cargos sin cruce", "Cargos de SAP que no encontraron un cargo del presupuesto, y "
                               "cargos presupuestados sin nadie en la nómina. Financiera los homologa en "
                               "Administración → Nómina → Homologar cargos.")
                s1 = ej_per[ej_per["comun"] == nomina.SIN_PPTO].groupby(["area", "cargo"]).agg(
                    Personas=("empleado_id", "nunique"), Valor=("valor", "sum")).reset_index() \
                    .sort_values("Valor", ascending=False).rename(columns={"area": "Área", "cargo": "Cargo en SAP"})
                st.markdown(f"**En SAP sin presupuesto** · {len(s1)} cargos · {E.millones(s1['Valor'].sum())}")
                st.dataframe(s1, hide_index=True, width="stretch",
                             column_config={"Valor": st.column_config.NumberColumn(format="accounting")})
                con = set(zip(re_mes["area"], re_mes["comun"]))
                s2 = pp_mes[[(a, c) not in con for a, c in zip(pp_mes["area"], pp_mes["comun"])]] \
                    .groupby(["area", "comun"]).agg(Personas=("personas", "sum"), Valor=("valor", "sum")) \
                    .reset_index().query("Personas > 0").sort_values("Valor", ascending=False) \
                    .rename(columns={"area": "Área", "comun": "Cargo presupuestado"})
                st.markdown(f"**Presupuestados sin nadie en la nómina de {MESES_L[mes_pers - 1]}** · {len(s2)} cargos")
                st.dataframe(s2, hide_index=True, width="stretch",
                             column_config={"Valor": st.column_config.NumberColumn("Valor del mes", format="accounting"),
                                            "Personas": st.column_config.NumberColumn(format="%.0f")})
