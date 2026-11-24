"""Módulo Nómina: presupuesto de personal por cargo vs ejecución de la nómina de SAP (costo y personas)."""
from html import escape

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import estilo as E
import nomina
from ui import EMPRESAS, MESES, MESES_L, a_excel, estilo_grafico, selector_periodo, titulo_tarjeta

SIN_AREA = "SIN ÁREA"
def _excel_hojas(hojas: dict) -> bytes:
    from io import BytesIO
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for nombre, df in hojas.items():
            df.to_excel(w, sheet_name=nombre[:31], index=False)
    return buf.getvalue()


def preparar(leer, anio: int, empresa: str = "ALDC"):
    """Presupuesto y ejecución del año y la compañía con área y cargo común. También el conteo de personas reales."""
    ppto = leer("SELECT * FROM nomina_ppto WHERE anio=? AND COALESCE(empresa, 'ALDC')=?", (anio, empresa))
    ejec = leer("SELECT * FROM nomina_ejec WHERE anio=? AND COALESCE(empresa, 'ALDC')=?", (anio, empresa))
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


# globos del módulo: la nómina y los demás reportes de personal (quién ve cada uno: ver app.puede_ver_reporte)
GLOBOS = [
    dict(nombre="Nómina", icono="badge", sub="Presupuesto de personal vs nómina ejecutada", base="Nómina"),
    dict(nombre="Dotación 1", icono="checkroom", sub="Dotación y EPP: kárdex de almacén, cruce y registros contables",
         base="Dotación y EPP", color="#2e9e5b", fondo="#e8f6ee"),
    dict(nombre="Dotación 2", icono="inventory_2", sub="Dotación y EPP por elemento: SAP detallado con el SISPO",
         base="Dotación y EPP", color="#2e7d32", fondo="#e6f4e7"),
    dict(nombre="Otros de personal", icono="groups", sub="Personal, dotación y otros gastos de RH · Excel para "
         "Gestión Humana", base="Nómina", color="#c47012", fondo="#fdf1e3"),
]


def mostrar(ctx):
    from ui import ADMIN, globos
    puede = getattr(ctx, "puede_ver", None)
    visibles = [g for g in GLOBOS if ctx.usuario == ADMIN or puede is None or puede(g["nombre"], g["base"])]
    if not visibles:
        st.warning("Tu usuario no tiene reportes en este módulo.")
        return
    vista = globos(visibles, key="nav_personal", titulo="Personal",
                   portada=("Personal", "Elige un reporte: nómina, dotación o el informe de otros gastos de personal"))
    if vista is None:                                   # menú: solo los globos
        return
    zona_encabezado = st.container()
    if vista == "Dotación 1":
        import mod_dotacion
        mod_dotacion.mostrar(ctx, zona_encabezado)
    elif vista == "Dotación 2":
        import mod_dotacion2
        mod_dotacion2.mostrar(ctx)
    elif vista == "Otros de personal":
        _informe_rh(ctx, zona_encabezado)
    else:
        _nomina(ctx, zona_encabezado)


def _nomina(ctx, zona_encabezado):
    leer = ctx.leer
    hay_emp = set(leer("SELECT DISTINCT COALESCE(empresa, 'ALDC') e FROM nomina_ppto")["e"]) | \
        set(leer("SELECT DISTINCT COALESCE(empresa, 'ALDC') e FROM nomina_ejec")["e"])
    empresas = [e for e in EMPRESAS if e in hay_emp]
    if not empresas:
        zona_encabezado.html(E.encabezado("Nómina", "Presupuesto de personal vs nómina ejecutada", []))
        st.info("Aún no hay información de nómina. Financiera la carga en Administración → Nómina.")
        return

    # ------------------------------------------------------------ filtros del informe
    with st.container(key="card_filtros"):
        st.html('<div class="nf-ftit">Filtros del informe</div>')
        c0, *c = st.columns([0.8, 0.7, 1.4, 0.95, 2.2], vertical_alignment="bottom")
        empresa = c0.selectbox("Compañía", empresas, key="nm_empresa",
                               help=" · ".join(f"{e}: {EMPRESAS[e]}" for e in EMPRESAS))
        anios = sorted(set(leer("SELECT DISTINCT anio FROM nomina_ppto WHERE COALESCE(empresa, 'ALDC')=?",
                                (empresa,))["anio"])
                       | set(leer("SELECT DISTINCT anio FROM nomina_ejec WHERE COALESCE(empresa, 'ALDC')=?",
                                  (empresa,))["anio"]), reverse=True)
        con_ejec = leer("SELECT DISTINCT anio FROM nomina_ejec WHERE COALESCE(empresa, 'ALDC')=?",
                        (empresa,))["anio"].tolist()
        anio = c[0].selectbox("Año", anios, index=anios.index(max(con_ejec)) if con_ejec else 0,
                              key=f"nm_anio_{empresa}")
        ppto, ejec, pers = preparar(leer, int(anio), empresa)
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
    zona_encabezado.html(E.encabezado(f"Nómina {anio} · {empresa}", "Presupuesto de personal por cargo vs nómina de SAP · "
                                      "cifras en millones de pesos", chips))
    if not hay_ppto_periodo and not ej_per.empty:
        st.warning("No hay presupuesto de nómina para el periodo elegido.")

    # «Sin homologar» está en Administración → Nómina (lo revisa y lo corrige Financiera)
    t1, t2, t3 = st.tabs(["Costo por cargo", "Personas por cargo", "Por concepto"], on_change="rerun", key="nm_tabs")

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
                     delta=E.delta_valor(var, pa, a_favor_si_sube=False),
                     sub=f'{ej_per["empleado_id"].nunique():,} personas con pago'.replace(",", "."),
                     icono="receipt", color=E.AZUL, fondo=E.AZUL_SUAVE,
                     spark=None if per.solo else [float(ejec.loc[ejec["mes"] == m, "valor"].sum())
                                                  if m in meses_ej else None for m in range(1, per.mes + 1)]),
                dict(label="Variación", valor=f'<span style="color:{E.VERDE if var <= 0 else E.ROJO}">'
                     f'{"+" if var >= 0 else ""}{E.millones(var)}</span>',
                     sub="Ejecutado − presupuesto (negativo: por debajo)", icono="check" if var <= 0 else "alert",
                     color=E.VERDE if var <= 0 else E.ROJO, fondo=E.VERDE_SUAVE if var <= 0 else E.ROJO_SUAVE),
                dict(label="% de ejecución", valor=f'<span style="color:{color}">{E.pct_txt(pct)}</span>',
                     sub="Del presupuesto del periodo", icono="calendar", color=color, fondo=E.AZUL_SUAVE),
            ]))
            if pa or ea:
                por_area = pd.DataFrame({"p": pp_per.groupby("area")["valor"].sum(),
                                         "e": ej_per.groupby("area")["valor"].sum()}).fillna(0)
                por_area["v"] = por_area["e"] - por_area["p"]
                por_cargo = pd.DataFrame({"p": pp_per.groupby("comun")["valor"].sum(),
                                          "e": ej_per.groupby("comun")["valor"].sum()}).fillna(0)
                por_cargo["v"] = por_cargo["e"] - por_cargo["p"]
                por_cargo = por_cargo[por_cargo.index != nomina.SIN_PPTO]
                texto = (f"Se pagaron {E.millones(ea)} frente a {E.millones(pa)} presupuestados "
                         f"(<b>{'+' if var >= 0 else ''}{E.millones(var)}</b>).")
                if not por_area.empty and por_area["v"].max() > 50_000:
                    a_ = por_area["v"].idxmax()
                    texto += f" El área con mayor sobrecosto es <b>{escape(str(a_))}</b> (+{E.millones(por_area.at[a_, 'v'])})"
                    if not por_cargo.empty and por_cargo["v"].max() > 50_000:
                        c_ = por_cargo["v"].idxmax()
                        texto += f" y el cargo, <b>{escape(str(c_))}</b> (+{E.millones(por_cargo.at[c_, 'v'])})"
                    texto += "."
                elif not por_area.empty and por_area["v"].min() < -50_000:
                    a_ = por_area["v"].idxmin()
                    texto += (f" Todas las áreas están dentro del presupuesto; la de mayor holgura es "
                              f"<b>{escape(str(a_))}</b> ({E.millones(por_area.at[a_, 'v'])}).")
                if sin:
                    texto += f" Hay {E.millones(sin)} de nómina sin homologar a un cargo del presupuesto."
                st.html(E.lectura(f"Nómina {per.suf}: {E.pct_txt(pct)} del presupuesto.", texto,
                                  "ok" if var <= 0 else ("warn" if pa and var / pa <= 0.03 else "bad")))
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
                                      abierto=abrir, formatos=["m", "m", "m", "p"], barras=[2]))
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

            _donde_bajo(t, e, pp_per, pp_mes, ej_per, re_mes, mes_pers, periodo_txt, anio, con_detalle)

    _informe_resumen(pp_per, ej_per, pp_mes, re_mes, ejec, mes_pers, per, anio, periodo_txt)



def _donde_bajo(t, e, pp_per, pp_mes, ej_per, re_mes, mes_pers, periodo_txt, anio, con_detalle):
    """Para un concepto (por defecto Sueldos y salarios): en qué áreas y cargos la nómina quedó por debajo del
    presupuesto, y si se explica por menos personas o por menor salario promedio (mes de corte)."""
    if t.empty or not con_detalle:
        return
    col_de = {n: c for c, n in nomina.CONCEPTOS}
    opciones = [n for n in t.index if n in col_de]
    if not opciones:
        return
    defecto = "Sueldos y salarios" if "Sueldos y salarios" in opciones else t["v"].idxmin()
    with st.container(key="card_nm_donde"):
        c1, c2 = st.columns([2.2, 1], vertical_alignment="bottom")
        with c1:
            titulo_tarjeta("¿Dónde bajó la nómina?",
                           f"{periodo_txt.capitalize()}. Por área (gerencia) y cargo: presupuesto vs ejecutado del "
                           "concepto elegido. La causa compara personas y valor promedio por persona en "
                           f"{MESES_L[mes_pers - 1]}.")
        conc = c2.selectbox("Concepto", opciones, index=opciones.index(defecto), key="nm_donde_conc")
        col = col_de[conc]
        ec = e[e["cc_"] == conc]
        t_ = pd.DataFrame({"p": pp_per.groupby(["area", "comun"])[col].sum(),
                           "e": ec.groupby(["area", "comun"])["valor"].sum()}).fillna(0.0)
        t_ = t_[(t_["p"].abs() > 0.5) | (t_["e"].abs() > 0.5)]
        if t_.empty:
            st.caption("Sin datos de este concepto en el periodo.")
            return
        t_["v"] = t_["e"] - t_["p"]
        # mes de corte: personas y valor del concepto por persona
        m = pd.DataFrame({"pp_pers": pp_mes.groupby(["area", "comun"])["personas"].sum(),
                          "re_pers": re_mes.groupby(["area", "comun"])["personas"].sum(),
                          "pp_mes": pp_mes.groupby(["area", "comun"])[col].sum(),
                          "e_mes": ec[ec["mes"] == mes_pers].groupby(["area", "comun"])["valor"].sum()}).fillna(0.0)
        t_ = t_.join(m, how="left").fillna(0.0)
        prom_pp = (t_["pp_mes"] / t_["pp_pers"]).where(t_["pp_pers"] > 0)
        t_["ef_pers"] = ((t_["re_pers"] - t_["pp_pers"]) * prom_pp).fillna(0.0)          # efecto de tener ± personas
        t_["ef_val"] = (t_["e_mes"] - t_["pp_mes"]) - t_["ef_pers"]                      # efecto del valor por persona

        def causa(r):
            if r.pp_pers == 0 and r.re_pers > 0:
                return "Sin presupuesto para el cargo"
            if r.re_pers == 0 and r.pp_pers > 0:
                return "Cargo sin personas en nómina"
            if abs(r.ef_pers) < 1 and abs(r.ef_val) < 1:
                return "—"
            if abs(r.ef_pers) >= abs(r.ef_val):
                return "Menos personas" if r.ef_pers < 0 else "Más personas"
            return "Menor valor por persona" if r.ef_val < 0 else "Mayor valor por persona"
        t_["causa"] = [causa(r) for r in t_.itertuples()]

        por_area = t_.groupby(level=0)[["p", "e", "v"]].sum().sort_values("v")
        bajan = t_[t_["v"] < -0.5].sort_values("v")
        suben = t_[t_["v"] > 0.5].sort_values("v", ascending=False)
        total_v = float(t_["v"].sum())

        # lectura
        if not bajan.empty:
            areas_b = por_area[por_area["v"] < 0]
            top_a = areas_b.head(3)
            causas = bajan.groupby("causa")["v"].sum().sort_values()
            top_c = bajan.head(3)
            txt = (f"<b>{escape(conc)}</b> quedó en {E.millones(t_['e'].sum())} frente a {E.millones(t_['p'].sum())} "
                   f"presupuestados ({'+' if total_v > 0 else ''}{E.millones(total_v)}). "
                   f"Disminuyó en <b>{len(bajan)}</b> cargos por {E.millones(bajan['v'].sum())}"
                   + (f" y aumentó en {len(suben)} cargos por +{E.millones(suben['v'].sum())}" if not suben.empty else "")
                   + ". ")
            if not top_a.empty:
                txt += ("Áreas con mayor disminución: " + ", ".join(
                    f"<b>{escape(str(a))}</b> ({E.millones(r.v)})" for a, r in top_a.iterrows()) + ". ")
            txt += ("Cargos: " + ", ".join(f"<b>{escape(str(c_))}</b> de {escape(str(a))} ({E.millones(r.v)})"
                                           for (a, c_), r in top_c.iterrows()) + ". ")
            txt += f"Causa que más pesa: <b>{escape(causas.index[0].lower())}</b> ({E.millones(causas.iloc[0])})."
            if causas.index[0] == "Cargo sin personas en nómina":
                txt += (" Revisa Administración → Nómina → Sin homologar: puede ser un cargo de SAP que no se cruzó "
                        "con el del presupuesto.")
            st.html(E.lectura(f"{conc}: dónde disminuyó.", txt, "ok" if total_v <= 0 else "warn"))

        # por área
        filas = [dict(nombre=str(a), sub=f"{E.millones(r.p)} → {E.millones(r.e)}",
                      celdas=[dict(v=r.v, pct=(r.e / r.p - 1) if r.p else None, fav=r.v <= 0)])
                 for a, r in por_area.iterrows()]
        tot = por_area.sum()
        filas.append(dict(nombre="Total", clase="total", sub=f"{E.millones(tot.p)} → {E.millones(tot.e)}",
                          celdas=[dict(v=tot.v, pct=(tot.e / tot.p - 1) if tot.p else None, fav=tot.v <= 0)]))
        st.html(E.barras_divergentes(filas, ["Variación (ejecutado − presupuesto)"], "Área"))

        # cargos
        def tabla(d):
            d = d.reset_index()
            return pd.DataFrame({
                "Área": d["area"], "Cargo": d["comun"], "Presupuesto": d["p"].round(0), "Ejecutado": d["e"].round(0),
                "Variación": d["v"].round(0), f"Personas ppto ({MESES[mes_pers - 1]})": d["pp_pers"].round(0),
                f"Personas nómina ({MESES[mes_pers - 1]})": d["re_pers"].round(0),
                "Valor/persona ppto": (d["pp_mes"] / d["pp_pers"]).where(d["pp_pers"] > 0).round(0),
                "Valor/persona real": (d["e_mes"] / d["re_pers"]).where(d["re_pers"] > 0).round(0),
                "Causa principal": d["causa"]})
        cfg = {c_: st.column_config.NumberColumn(format="localized")
               for c_ in ("Presupuesto", "Ejecutado", "Variación", "Valor/persona ppto", "Valor/persona real")}
        st.markdown(f"**Cargos donde disminuyó** · {len(bajan)} cargos")
        tb = tabla(bajan)
        st.dataframe(tb, hide_index=True, width="stretch", height=min(35 * (len(tb) + 1) + 3, 420), column_config=cfg)
        if not suben.empty:
            with st.expander(f"Cargos donde aumentó · {len(suben)} cargos ({E.millones(suben['v'].sum())})"):
                st.dataframe(tabla(suben), hide_index=True, width="stretch", column_config=cfg)
        resumen_causa = t_.groupby("causa").agg(Cargos=("v", "size"), Variación=("v", "sum")).reset_index() \
            .rename(columns={"causa": "Causa principal"}).sort_values("Variación")
        st.download_button(
            f"Descargar {conc.lower()} por área y cargo", icon=":material/download:", key="nm_donde_xlsx",
            data=lambda: _excel_hojas({"Disminuyó": tabla(bajan), "Aumentó": tabla(suben),
                                       "Por área": por_area.reset_index().rename(columns={
                                           "area": "Área", "p": "Presupuesto", "e": "Ejecutado", "v": "Variación"}),
                                       "Por causa": resumen_causa}),
            file_name=f"nomina_{col}_{anio}.xlsx")
        st.caption("Causa principal (mes de corte): «Menos personas» = la diferencia de personas × el valor "
                   "presupuestado por persona explica la mayor parte; «Menor valor por persona» = con las personas "
                   "que hay, se pagó menos por persona que lo presupuestado (salario menor, incapacidades, "
                   "ingresos a mitad de mes…).")


def _informe_resumen(pp_per, ej_per, pp_mes, re_mes, ejec, mes_pers, per, anio, periodo_txt):
    """Las principales diferencias de la nómina frente al presupuesto, en pocas frases (al final del módulo)."""
    pa, ea = float(pp_per["valor"].sum()), float(ej_per["valor"].sum())
    if not pa and not ea:
        return
    var = ea - pa
    general = [f"La nómina ejecutada fue <b>{E.millones(ea)}</b> frente a <b>{E.millones(pa)}</b> presupuestados: "
               f"{E.var_html(var)} ({E.pct_txt(ea / pa) if pa else '—'} del presupuesto)."]
    # personas y costo por persona en el mes de corte: separa la diferencia en "más gente" y "más costo por persona"
    pp_n, re_n = float(pp_mes["personas"].sum()), float(re_mes["personas"].sum())
    pa_m, ea_m = float(pp_mes["valor"].sum()), float(ejec.loc[ejec["mes"] == mes_pers, "valor"].sum())
    if pp_n and re_n and pa_m and ea_m:
        c_pp, c_re = pa_m / pp_n, ea_m / re_n
        ef_personas, ef_costo = (re_n - pp_n) * c_pp, (c_re - c_pp) * re_n
        mes_t = MESES_L[mes_pers - 1]
        general.append(f"En {mes_t} hubo <b>{re_n:,.0f}</b> personas con pago frente a <b>{pp_n:,.0f}</b> "
                       f"presupuestadas ({'+' if re_n >= pp_n else ''}{re_n - pp_n:,.0f}).".replace(",", "."))
        general.append(f"Costo promedio por persona en {mes_t}: <b>{E.pesos(c_re)}</b> frente a {E.pesos(c_pp)} "
                       f"presupuestado. De la diferencia del mes ({E.var_html(ea_m - pa_m)}), "
                       f"{E.var_html(ef_personas)} se explica por número de personas y {E.var_html(ef_costo)} por "
                       "costo por persona.")

    por_area = pd.DataFrame({"p": pp_per.groupby("area")["valor"].sum(),
                             "e": ej_per.groupby("area")["valor"].sum()}).fillna(0.0)
    por_area["v"] = por_area["e"] - por_area["p"]
    pers_pp = pp_mes.groupby("comun")["personas"].sum()
    pers_re = re_mes.groupby("comun")["personas"].sum()
    por_cargo = pd.DataFrame({"p": pp_per.groupby("comun")["valor"].sum(),
                              "e": ej_per.groupby("comun")["valor"].sum()}).fillna(0.0)
    por_cargo = por_cargo[por_cargo.index != nomina.SIN_PPTO]
    por_cargo["v"] = por_cargo["e"] - por_cargo["p"]

    def f_area(nombre, r):
        return (f"<b>{escape(str(nombre))}</b>: {E.var_html(r.v)} ({E.pct_txt(r.e / r.p) if r.p else 'sin presupuesto'}"
                " del presupuesto).")

    def f_cargo(nombre, r):
        n_p, n_r = float(pers_pp.get(nombre, 0)), float(pers_re.get(nombre, 0))
        gente = (f" — {n_r:,.0f} personas vs {n_p:,.0f} presupuestadas en {MESES_L[mes_pers - 1]}".replace(",", ".")
                 if n_p or n_r else "")
        return f"<b>{escape(str(nombre))}</b>: {E.var_html(r.v)}{gente}."
    areas_arriba = [f_area(n, r) for n, r in por_area[por_area["v"] > 50_000].nlargest(3, "v").iterrows()]
    areas_abajo = [f_area(n, r) for n, r in por_area[por_area["v"] < -50_000].nsmallest(3, "v").iterrows()]
    cargos_arriba = [f_cargo(n, r) for n, r in por_cargo[por_cargo["v"] > 50_000].nlargest(3, "v").iterrows()]
    cargos_abajo = [f_cargo(n, r) for n, r in por_cargo[por_cargo["v"] < -50_000].nsmallest(3, "v").iterrows()]

    pendientes = []
    sin = float(ej_per.loc[ej_per["comun"] == nomina.SIN_PPTO, "valor"].sum())
    if sin:
        pendientes.append(f"<b>{E.millones(sin)}</b> de nómina de cargos que no cruzan con el presupuesto (se "
                          "homologan en Administración → Nómina).")

    titulo = "Informe de personal (nómina)"
    sub = f"{periodo_txt.capitalize()} {anio} · presupuesto de personal por cargo vs nómina de SAP"
    secciones = [("En general", general), ("Áreas por encima del presupuesto", areas_arriba),
                 ("Áreas por debajo", areas_abajo), ("Cargos con mayor sobrecosto", cargos_arriba),
                 ("Cargos con mayor holgura", cargos_abajo), ("Pendiente", pendientes)]
    st.html(E.informe(titulo, sub, secciones))
    with st.expander("Copiar el informe en texto (para un correo)", icon=":material/content_copy:"):
        st.code(E.informe_texto(titulo, sub, secciones), language=None, wrap_lines=True)


def _informe_rh(ctx, zona_encabezado):
    """Descarga el Excel de personal, dotación y otros gastos de RH armado con lo que ya está cargado en el panel."""
    import informe_rh
    from ui import EMPRESAS
    leer = ctx.leer
    zona_encabezado.html(E.encabezado("Informe para Gestión Humana",
                                      "Personal, dotación y otros gastos de RH: presupuesto vs ejecución · Excel", []))
    with st.container(key="card_inf_rh"):
        titulo_tarjeta("Parámetros del informe",
                       "Sale de la ejecución de SAP y del presupuesto cargados en el panel, y de la dotación (kárdex y "
                       "registros contables) del módulo Dotación y EPP.")
        c1, c2, c3 = st.columns(3)
        empresa = c1.selectbox("Compañía", list(EMPRESAS), format_func=lambda e: f"{e} · {EMPRESAS[e]}",
                               key="inf_rh_emp")
        anios = sorted(leer("SELECT DISTINCT anio FROM ejecucion WHERE empresa=?", (empresa,))["anio"].tolist(),
                       reverse=True)
        if not anios:
            st.info("Aún no hay ejecución cargada para esta compañía.")
            return
        anio = c2.selectbox("Año", anios, key="inf_rh_anio")
        meses_e = leer("SELECT DISTINCT mes FROM ejecucion WHERE empresa=? AND anio=?", (empresa, int(anio)))["mes"]
        ult = int(meses_e.max()) if len(meses_e) else 12
        hasta = c3.selectbox("Ejecución hasta", list(range(1, ult + 1)), index=ult - 1,
                             format_func=lambda m: MESES_L[m - 1].capitalize(), key=f"inf_rh_hasta_{anio}")
        versiones = leer("SELECT DISTINCT version FROM presupuesto WHERE anio=? AND empresa=?",
                         (int(anio), empresa))["version"].tolist()
        if not versiones:
            st.info("Aún no hay presupuesto cargado para esta compañía y año.")
            return
        orden = sorted(versiones, key=lambda v: (v != "Inicial", v))
        c4, c5, c6 = st.columns(3)
        v1 = c4.selectbox("Presupuesto de los primeros meses", orden, index=0, key="inf_rh_v1")
        corte_v = c5.selectbox("Cambia de presupuesto desde", list(range(2, 13)), index=7,
                               format_func=lambda m: MESES_L[m - 1].capitalize(), key="inf_rh_corte") - 1
        v2_def = "8+4" if "8+4" in versiones else orden[-1]
        v2 = c6.selectbox("Presupuesto de los meses restantes (UPA)", orden, index=orden.index(v2_def),
                          key="inf_rh_v2")

        ppto, ejec = informe_rh.preparar(leer, empresa, anio, v1, v2, corte_v)
        r = informe_rh.resumen(informe_rh.tabla_mensual(ppto, ejec), hasta)
        vista = r.reset_index().rename(columns={"concepto_rh": "Concepto", "grupo": "Grupo", "ppto": "Presupuesto",
                                                "ejec": "Ejecutado", "dif": "Diferencia", "pct": "% ejecución"})
        vista = vista[["Grupo", "Concepto", "Presupuesto", "Ejecutado", "Diferencia", "% ejecución"]]
        tot_p, tot_e = float(r["ppto"].sum()), float(r["ejec"].sum())
        st.html(E.kpis([
            dict(label=f"Presupuesto a {MESES[hasta - 1].lower()}", valor=E.millones(tot_p),
                 sub=f"Ene–{MESES[corte_v - 1].lower()}: {v1} · {MESES[corte_v].lower()}–dic: {v2}"
                 if corte_v < 12 else v1, icono="wallet", color=E.AZUL, fondo=E.AZUL_SUAVE),
            dict(label=f"Ejecutado a {MESES[hasta - 1].lower()}", valor=E.millones(tot_e),
                 delta=E.delta_valor(tot_e - tot_p, tot_p, a_favor_si_sube=False), sub="Ejecución de SAP",
                 icono="receipt", color=E.AZUL, fondo=E.AZUL_SUAVE),
            dict(label="Diferencia", valor=E.millones(tot_e - tot_p), sub="Ejecutado − presupuesto",
                 icono="alert" if tot_e > tot_p else "check", color=E.ROJO if tot_e > tot_p else E.VERDE,
                 fondo=E.ROJO_SUAVE if tot_e > tot_p else E.VERDE_SUAVE),
        ], columnas=3))
        st.dataframe(vista, hide_index=True, width="stretch", column_config={
            c: st.column_config.NumberColumn(format="accounting") for c in ("Presupuesto", "Ejecutado", "Diferencia")}
            | {"% ejecución": st.column_config.NumberColumn(format="percent")})
        nombre = f"{empresa}_PERSONAL_DOTACION_OTROS_GASTOS_RH_{MESES[hasta - 1].upper()}_{anio}.xlsx"
        st.download_button("Descargar informe en Excel", type="primary", icon=":material/download:",
                           data=lambda: informe_rh.generar(leer, empresa, EMPRESAS[empresa], int(anio), hasta, v1, v2,
                                                           corte_v),
                           file_name=nombre, key="inf_rh_dl")
        st.caption("Hojas: Resumen, Mes a mes, Dotación comparativo, Detalle dotación y Detalle movimientos (sin "
                   "nómina). Personal = clase 3.0; beneficios en salud = 3.3 Servicios en salud; transporte = 3.6 "
                   "Transporte operarios; comedor = 3.9 Casino. El detalle por ítem de bienestar o formación "
                   "(cumpleaños, capacitaciones…) no está en el presupuesto cargado: se ve en los movimientos.")


def sin_homologar_admin(ctx):
    """Administración → Nómina → Sin homologar: cargos de SAP que no encontraron un cargo del presupuesto, y cargos
    presupuestados sin nadie en la nómina del mes. Se corrigen en «Homologar cargos»."""
    leer = ctx.leer
    with st.container(key="card_adm_nm_sin"):
        titulo_tarjeta("Sin homologar",
                       "Cargos de SAP que no encontraron un cargo del presupuesto, y cargos presupuestados sin nadie en "
                       "la nómina. Se corrigen en la pestaña «Homologar cargos».")
        hay = leer("SELECT DISTINCT COALESCE(empresa, 'ALDC') empresa, anio, mes FROM nomina_ejec")
        if hay.empty:
            st.info("Aún no hay nómina ejecutada cargada.")
            return
        c = st.columns([1.2, .8, 1, 1.4], vertical_alignment="bottom")
        empresas = [e for e in EMPRESAS if e in set(hay["empresa"])] or sorted(set(hay["empresa"]))
        empresa = c[0].selectbox("Compañía", empresas, key="adm_nm_sin_emp",
                                 format_func=lambda e: f"{e} · {EMPRESAS.get(e, e)}")
        h_e = hay[hay["empresa"] == empresa]
        anios = sorted(set(h_e["anio"].astype(int)), reverse=True)
        anio = int(c[1].selectbox("Año", anios, key=f"adm_nm_sin_anio_{empresa}"))
        meses = sorted(set(h_e.loc[h_e["anio"] == anio, "mes"].astype(int)))
        mes = int(c[2].selectbox("Mes", meses, index=len(meses) - 1, key=f"adm_nm_sin_mes_{empresa}_{anio}",
                                 format_func=lambda m: MESES_L[m - 1].capitalize()))
        modo = c[3].segmented_control("Periodo", ["Solo el mes", "Acumulado"], default="Solo el mes",
                                      key="adm_nm_sin_modo") or "Solo el mes"
        ppto, ejec, pers = preparar(leer, anio, empresa)
        ej_per = ejec[ejec["mes"] == mes] if modo == "Solo el mes" else ejec[ejec["mes"] <= mes]
        pp_mes, re_mes = ppto[ppto["mes"] == mes], pers[pers["mes"] == mes]
        periodo = MESES_L[mes - 1] if modo == "Solo el mes" else f"enero a {MESES_L[mes - 1]}"

        s1 = ej_per[ej_per["comun"] == nomina.SIN_PPTO].groupby(["area", "cargo"]).agg(
            Personas=("empleado_id", "nunique"), Valor=("valor", "sum")).reset_index() \
            .sort_values("Valor", ascending=False).rename(columns={"area": "Área", "cargo": "Cargo en SAP"})
        st.markdown(f"**En SAP sin presupuesto** ({periodo} {anio}) · {len(s1)} cargos · {E.millones(s1['Valor'].sum())}")
        st.dataframe(s1, hide_index=True, width="stretch",
                     column_config={"Valor": st.column_config.NumberColumn(format="accounting")})
        con = set(zip(re_mes["area"], re_mes["comun"]))
        s2 = pp_mes[[(a, c_) not in con for a, c_ in zip(pp_mes["area"], pp_mes["comun"])]] \
            .groupby(["area", "comun"]).agg(Personas=("personas", "sum"), Valor=("valor", "sum")) \
            .reset_index().query("Personas > 0").sort_values("Valor", ascending=False) \
            .rename(columns={"area": "Área", "comun": "Cargo presupuestado"})
        st.markdown(f"**Presupuestados sin nadie en la nómina de {MESES_L[mes - 1]} {anio}** · {len(s2)} cargos")
        st.dataframe(s2, hide_index=True, width="stretch",
                     column_config={"Valor": st.column_config.NumberColumn("Valor del mes", format="accounting"),
                                    "Personas": st.column_config.NumberColumn(format="%.0f")})
        if len(s1) or len(s2):
            st.download_button("Descargar sin homologar", lambda a=s1, b=s2: _excel_hojas(
                {"En SAP sin presupuesto": a, "Presupuestados sin nómina": b}),
                f"nomina_sin_homologar_{empresa}_{anio}{mes:02d}.xlsx", icon=":material/download:",
                key="adm_nm_sin_bajar")
