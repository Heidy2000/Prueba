"""Módulo Nómina: presupuesto de personal por cargo vs ejecución de la nómina de SAP (costo y personas)."""
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import estilo as E
import nomina
import dotacion as D
import logica_m as L
import sispo_fuente
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


# ====================================================================== DOTACIÓN Y EPP's
def mostrar_dotacion(ctx):
    """Página «Dotación» de Talento humano (independiente de si hay nómina cargada)."""
    _dotacion(ctx)


def _ejecucion_dotacion(leer, empresa: str, anio: int):
    """Ejecución de dotación desde el SISPO guardado de la compañía (todos los meses del año) y errores de lectura."""
    partes, errores = [], []
    g = sispo_fuente.listar(leer)
    for per in sorted(g.loc[(g["empresa"] == empresa) & (g["periodo"] // 100 == anio), "periodo"]):
        try:
            partes.append(D.ejecucion_sispo(L.sispo_leer(sispo_fuente.abrir(leer, empresa, int(per)))))
        except Exception as err:
            errores.append(f"SISPO {empresa} {per}: {err}")
    vacio = pd.DataFrame(columns=["periodo", "mes", "codigo", "articulo", "cc_codigo", "cc_nombre", "cantidad", "valor"])
    return (pd.concat(partes, ignore_index=True) if partes else vacio), errores


def _datos_dotacion(leer, pp_all: pd.DataFrame, empresa: str, anio: int):
    """Todo lo que necesita el archivo de ejecución de UNA compañía y año: presupuesto, ejecución del SISPO, SAP, ejecución mensual de SAP,
    equivalencias y meses con ejecución. Es el mismo cálculo de la pantalla."""
    ppto = pp_all[(pp_all["empresa"] == empresa) & (pp_all["anio"] == anio)]
    ejec, _ = _ejecucion_dotacion(leer, empresa, anio)
    pool = leer("SELECT * FROM ejecucion WHERE empresa=? AND anio=?", (empresa, anio))
    sap_mes = D.sap_por_mes(pool, empresa, anio)
    meses_ej = sorted(set(int(m) for m in ejec["mes"].unique()) | set(int(m) for m in sap_mes))
    gu = leer("SELECT codigo, elemento FROM dot_map WHERE empresa=?", (empresa,))
    elementos = sorted(ppto.loc[ppto["seccion"].isin(D.SECCIONES_COMPARADAS), "elemento"].unique())
    eq = D.equivalencias(ejec, elementos, dict(zip(gu["codigo"], gu["elemento"])))
    return dict(ppto=ppto, ejec=ejec, pool=pool, sap_mes=sap_mes, meses_ej=meses_ej, eq=eq)


def _descarga_dotacion(leer, pp_all: pd.DataFrame, empresa_pag: str, anio: int, per):
    """Descarga del archivo de ejecución: compañía y mes se eligen aquí (por defecto, los de la página). Un mes = solo ese mes;
    «Acumulado» = todos los meses con ejecución del año."""
    empresas = sorted(pp_all["empresa"].unique())
    c1, c2, c3 = st.columns([1, 1.4, 2], vertical_alignment="bottom")
    emp = c1.selectbox("Compañía del informe", empresas, index=empresas.index(empresa_pag) if empresa_pag in empresas else 0, key="dot_emp_informe")
    if not (pp_all["empresa"] == emp).any() or not ((pp_all["empresa"] == emp) & (pp_all["anio"] == anio)).any():
        st.warning(f"{emp} no tiene presupuesto de dotación para {anio}.")
        return
    d = _datos_dotacion(leer, pp_all, emp, anio)
    meses_ej = d["meses_ej"]
    if not meses_ej:
        st.warning(f"No hay ejecución de {emp} en {anio}: el archivo saldría vacío.")
        return
    opciones = [0, *meses_ej]
    por_defecto = opciones.index(per.mes) if (per.solo and per.mes in meses_ej and emp == empresa_pag) else 0
    sel = c2.selectbox("Mes del informe", opciones, index=por_defecto, key=f"dot_mes_informe_{emp}",
                       format_func=lambda m: f"Acumulado a {MESES_L[meses_ej[-1] - 1]}" if m == 0 else MESES_L[m - 1].capitalize(),
                       help="Un mes = solo ese mes. Acumulado = todos los meses con ejecución del año.")
    meses = meses_ej if sel == 0 else [sel]
    txt = f"{MESES[meses[0] - 1]}–{MESES[meses[-1] - 1]}" if len(meses) > 1 else MESES[meses[0] - 1]
    sap_f = D.sap_filas(d["pool"], emp, anio)
    c3.download_button(f"Descargar ejecución de dotación · {emp} (Excel)",
                       lambda: D.libro_ejecucion(d["ppto"], d["ejec"], d["eq"], sap_f, d["sap_mes"], meses, f"Dotación y EPP's · {emp} {anio}"),
                       file_name=f"ejecucion_dotacion_{emp}_{anio}_{txt.replace(' ', '').replace('–', '-')}.xlsx", icon=":material/download:",
                       key=f"dot_desc_{emp}")


def _dotacion(ctx):
    leer = ctx.leer
    ctx.con.crear_tablas(D.TABLAS)
    st.html(E.encabezado("Dotación y EPP's", "Presupuesto por elemento vs salidas de inventario del SISPO · cifras en millones de pesos", []))
    pp_all = leer("SELECT empresa, anio, seccion, elemento, precio, mes, cantidad, valor FROM dot_ppto")
    if pp_all.empty:
        st.info("Aún no hay presupuesto de dotación. Financiera lo carga en Administración → Nómina → Dotación.")
        return
    with st.container(key="card_filtros_dot"):
        st.html('<div class="nf-ftit">Filtros del informe</div>')
        c = st.columns([0.9, 0.7, 1.4, 0.95], vertical_alignment="bottom")
        empresa = c[0].selectbox("Compañía", sorted(pp_all["empresa"].unique()), key="dot_emp")
        anios = sorted(pp_all.loc[pp_all["empresa"] == empresa, "anio"].unique(), reverse=True)
        anio = int(c[1].selectbox("Año", anios, key="dot_anio"))
        ppto = pp_all[(pp_all["empresa"] == empresa) & (pp_all["anio"] == anio)]
        ejec, errores = _ejecucion_dotacion(leer, empresa, anio)
        pool = leer("SELECT * FROM ejecucion WHERE empresa=? AND anio=?", (empresa, anio))
        sap_mes = D.sap_por_mes(pool, empresa, anio)
        meses_ej = sorted(set(ejec["mes"].unique()) | set(sap_mes))
        per = selector_periodo(c[2], c[3], int(max(meses_ej)) if meses_ej else 1, "dot")
    for e in errores:
        st.error(e)
    if not meses_ej:
        st.warning(f"No hay ejecución de SAP (3.1 · Dotación y EPP's) ni SISPO de {empresa} para {anio}.")
        return
    meses = [m for m in per.meses() if m in meses_ej]
    faltan = [m for m in per.meses() if m not in meses_ej]
    if faltan and not per.solo:
        st.caption("Sin ejecución en SAP para: " + ", ".join(MESES[m - 1] for m in faltan)
                   + ". Se compara solo con los meses que tienen ejecución (el presupuesto de esos meses).")
    if not meses:
        st.warning("No hay ejecución del mes elegido.")
        return
    gu = leer("SELECT codigo, elemento FROM dot_map WHERE empresa=?", (empresa,))
    elementos = sorted(ppto.loc[ppto["seccion"].isin(D.SECCIONES_COMPARADAS), "elemento"].unique())
    eq = D.equivalencias(ejec, elementos, dict(zip(gu["codigo"], gu["elemento"])))
    t = D.comparar(ppto, ejec, eq, meses, sap_mes)
    tot = D.totales(t)
    txt_per = ", ".join(MESES[m - 1] for m in meses) if len(meses) <= 3 else f"{MESES[meses[0] - 1]}–{MESES[meses[-1] - 1]}"

    def kpis(tt, sub):
        tv = D.totales(tt)
        ok = tv["variacion"] <= 0
        return E.kpis([
            dict(label="Presupuesto", valor=E.millones(tv["presupuesto"]), sub=sub, icono="wallet", color=E.AZUL, fondo=E.AZUL_SUAVE),
            dict(label="Ejecutado", valor=E.millones(tv["ejecutado"]), sub="SAP 3.1 · Dotación y EPP's, detallado con SISPO", icono="receipt", color=E.AZUL, fondo=E.AZUL_SUAVE),
            dict(label="Variación", valor=f'<span style="color:{E.VERDE if ok else E.ROJO}">{"+" if tv["variacion"] >= 0 else ""}{E.millones(tv["variacion"])}</span>',
                 sub="Ejecutado − presupuesto (negativo: por debajo)", icono="check" if ok else "alert",
                 color=E.VERDE if ok else E.ROJO, fondo=E.VERDE_SUAVE if ok else E.ROJO_SUAVE),
            dict(label="% de ejecución", valor=E.pct_txt(tv["pct"]), sub="Del presupuesto del periodo", icono="calendar",
                 color=E.color_barra(tv["pct"], tv["ejecutado"], False), fondo=E.AZUL_SUAVE)])

    if tot["otros"] > 0:
        n = int((eq["origen"] == "Pendiente").sum())
        st.warning(f"{E.millones(tot['otros'])} del ejecutado están en «Otros (sin clasificar)», dentro de Dotación: artículos del SISPO sin elemento "
                   f"({n} sin asignar) y movimientos de SAP que el SISPO no detalla. Asígnalos en Administración → Nómina → Dotación; "
                   "el detalle de cada línea va en el archivo de ejecución.")
    difs = []
    for m in [m for m in meses if m in set(ejec["mes"])]:
        r = D.cuadre_sap(ejec, pool, empresa, anio * 100 + m)
        if r["sap"] is None:
            difs.append(f"{MESES[m - 1]}: SAP no trae «Salida de inventario» de DOTACIÓN PERSONAL")
        elif abs(r["diferencia"]) > 1:
            difs.append(f"{MESES[m - 1]}: SAP {E.millones(r['sap'])} vs SISPO {E.millones(r['sispo'])}")
    if difs:
        st.warning("El SISPO no cuadra con SAP en la cuenta DOTACIÓN PERSONAL → " + " · ".join(difs))

    def filas_visual(g):
        """Filas para la tabla visual (solo elementos con presupuesto o ejecución), de mayor a menor ejecutado; «Otros» al final."""
        g = g[(g["presupuesto"] > 0) | (g["ejecutado"] > 0)]
        g = pd.concat([g[g["elemento"] != D.OTROS].sort_values(["ejecutado", "presupuesto"], ascending=False), g[g["elemento"] == D.OTROS]])
        return [dict(nombre=r.elemento, presupuesto=r.presupuesto, ejecutado=r.ejecutado,
                     sub="" if r.ppto_cant == 0 and r.ejec_cant == 0 else f"{r.ejec_cant:,.0f} de {r.ppto_cant:,.0f} und".replace(",", "."))
                for r in g.itertuples()]

    def participacion(g, titulo):
        """Barra al 100 %: cuánto pesa cada elemento en lo ejecutado (los 7 mayores y el resto en «Otros elementos»; «Otros (sin clasificar)» aparte)."""
        x = g[g["ejecutado"] > 0].sort_values("ejecutado", ascending=False)
        fijo = x[x["elemento"] == D.OTROS]
        x = x[x["elemento"] != D.OTROS]
        seg = [(r.elemento, r.ejecutado, None) for r in x.head(7).itertuples()]
        if len(x) > 7:
            seg.append(("Otros elementos", float(x.iloc[7:]["ejecutado"].sum()), None))
        seg += [(D.OTROS, float(r.ejecutado), "#8fa3bd") for r in fijo.itertuples()]
        if seg:
            st.markdown(f"**{titulo}**")
            st.html(E.barra_segmentada(seg))

    def tabla(g, clave):
        v = g[(g["presupuesto"] > 0) | (g["ejecutado"] > 0)].rename(columns={
            "elemento": "Elemento", "ppto_cant": "Cant. ppto", "presupuesto": "Presupuesto", "ejec_cant": "Cant. ejecutada", "ejecutado": "Ejecutado",
            "movimientos": "Salidas", "saldo": "Saldo", "pct": "% ejecución", "estado": "Estado"})
        st.dataframe(v[["Elemento", "Cant. ppto", "Cant. ejecutada", "Presupuesto", "Ejecutado", "Saldo", "% ejecución", "Estado"]], hide_index=True, width="stretch",
                     key=f"dot_tab_{clave}", column_config={
                         "Presupuesto": st.column_config.NumberColumn(format="$ %.0f"), "Ejecutado": st.column_config.NumberColumn(format="$ %.0f"),
                         "Saldo": st.column_config.NumberColumn(format="$ %.0f"), "% ejecución": st.column_config.NumberColumn(format="percent"),
                         "Cant. ppto": st.column_config.NumberColumn(format="%.0f"), "Cant. ejecutada": st.column_config.NumberColumn(format="%.0f")})

    tr, td_, te = st.tabs(["Resumen", "Dotación", "EPP's"], on_change="rerun", key="dot_tabs")
    with tr:
        if tr.open:
            st.html(kpis(t, f"Dotación y EPP's · {txt_per}"))
            with st.container(key="card_dot_res"):
                titulo_tarjeta("Por grupo", f"{empresa} {anio} · {txt_per}. Cada grupo se abre en sus elementos. Barra: relleno = ejecutado, marca = presupuesto.")
                grupos = []
                for nombre, sec in (("Dotación", D.DOTACION), ("EPP's", D.EPPS)):
                    x = t[t["seccion"] == sec]
                    grupos.append(dict(nombre=nombre, presupuesto=float(x["presupuesto"].sum()), ejecutado=float(x["ejecutado"].sum()),
                                       sub=f"{int(((x['presupuesto'] > 0) | (x['ejecutado'] > 0)).sum())} elementos", hijos=filas_visual(x)))
                st.html(E.tabla_visual(grupos, "Grupo", abierto=False))
                participacion(t, "Participación de cada elemento en lo ejecutado")
            with st.container(key="card_dot_mes"):
                titulo_tarjeta("Mes a mes", "Presupuesto de Dotación y EPP's vs ejecutado, en los meses con ejecución.")
                mm = []
                for m in meses_ej:
                    if m <= per.mes:
                        tm = D.totales(D.comparar(ppto, ejec, eq, [m], sap_mes))
                        mm.append((MESES[m - 1], tm["presupuesto"], tm["ejecutado"], tm["pct"]))
                fig2 = go.Figure()
                fig2.add_bar(x=[a for a, *_ in mm], y=[b / 1e6 for _, b, *_ in mm], name="Presupuesto", marker_color="#dbe3ee")
                fig2.add_bar(x=[a for a, *_ in mm], y=[c_ / 1e6 for _, _, c_, _ in mm], name="Ejecutado", marker_color=E.AZUL,
                             text=[E.pct_txt(p) for *_, p in mm], textposition="outside", cliponaxis=False)
                fig2.update_layout(barmode="group", yaxis_title="Millones de pesos")
                estilo_grafico(fig2, 300)
                st.plotly_chart(fig2, width="stretch", key="dot_graf_mes")
            with st.container(key="card_dot_desc"):
                titulo_tarjeta("Archivo de ejecución", "Excel con Resumen (presupuesto y ejecución mes a mes de Dotación y EPP's) y una hoja de movimientos "
                               "para Dotación y otra solo para EPPs. Elige la compañía y el mes del archivo.")
                _descarga_dotacion(leer, pp_all, empresa, anio, per)
    for pestana, sec, nombre in ((td_, D.DOTACION, "Dotación"), (te, D.EPPS, "EPP's")):
        with pestana:
            if pestana.open:
                x = t[t["seccion"] == sec]
                st.html(kpis(x, f"{nombre} · {txt_per}"))
                with st.container(key=f"card_dot_{sec}"):
                    titulo_tarjeta(f"{nombre} por elemento", f"{empresa} {anio} · {txt_per}. Presupuesto, ejecutado, saldo y participación en una sola línea. "
                                   "Barra: relleno = ejecutado, marca = presupuesto.")
                    fv = filas_visual(x)
                    if fv:
                        st.html(E.tabla_visual(fv, "Elemento"))
                        participacion(x, f"Participación de cada elemento en el ejecutado de {nombre}")
                    else:
                        st.info("No hay presupuesto ni ejecución en este grupo para el periodo elegido.")
                    with st.expander("Ver como tabla (cantidades y estado)", icon=":material/table:"):
                        tabla(x, sec)
