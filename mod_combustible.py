"""Módulo Combustible: relación de combustible por empresa (ALDC, ALSM), proveedor """
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import asignacion
import estilo as E
from ui import MESES, a_excel, estilo_grafico, selector_periodo, titulo_tarjeta

COLOR_EMPRESA = {"ALDC": E.AZUL, "ALSM": "#e8892c"}
COLOR_PROV = ["#1f5fc4", "#2e9e5b", "#e8892c", "#8a5cc4", "#9aa7b8"]


def _galones(v: float) -> str:
    return f"{v:,.0f} gal".replace(",", ".")


def _precio(v: float) -> str:
    return "—" if pd.isna(v) else "$" + f"{v:,.0f}".replace(",", ".") + "/gal"


def mostrar(ctx):
    leer = ctx.leer
    todo = leer("SELECT * FROM combustible")
    zona_encabezado = st.container()
    if todo.empty:
        zona_encabezado.html(E.encabezado("Combustible", "Relación de combustible por empresa, proveedor y vehículo", []))
        st.info("Aún no hay datos de combustible. Financiera los carga en Administración → Combustible.")
        return

    # ------------------------------------------------------------ filtros del informe
    with st.container(key="card_filtros"):
        st.html('<div class="nf-ftit">Filtros del informe</div>')
        c = st.columns([0.8, 1.1, 1.55, 0.95], vertical_alignment="bottom")
        anios = sorted(todo["anio"].unique(), reverse=True)
        anio = c[0].selectbox("Año", anios, key="cb_anio")
        base = todo[todo["anio"] == anio]
        empresas = sorted(base["empresa"].unique())
        emp_sel = c[1].selectbox("Empresa", ["Todas"] + empresas, key="cb_emp")
        if emp_sel != "Todas":
            base = base[base["empresa"] == emp_sel]
        per = selector_periodo(c[2], c[3], int(base["mes"].max()), "cb")

        c = st.columns([1.2, 1.2, 1.1, 1.9], vertical_alignment="bottom")
        prov = c[0].multiselect("Proveedor", sorted(base["fuente"].unique()), placeholder="Todos", key="cb_prov")
        tipos = c[1].multiselect("Tipo de combustible", sorted(base["tipo"].unique()), placeholder="Todos",
                                 key="cb_tipo")
        contratos = sorted(base["contrato"].unique())
        contrato = c[2].selectbox("Contrato", ["Todos"] + contratos, key="cb_contr") if len(contratos) > 1 else "Todos"
        if prov:
            base = base[base["fuente"].isin(prov)]
        if tipos:
            base = base[base["tipo"].isin(tipos)]
        if contrato != "Todos":
            base = base[base["contrato"] == contrato]
        placas = c[3].multiselect("Placa", sorted(base["placa"].unique()), placeholder="Todas las placas",
                                  key="cb_placa")
        if placas:
            base = base[base["placa"].isin(placas)]

    df = per.filtrar(base)
    ult = int(todo.loc[todo["anio"] == anio, "mes"].max())
    chips = [(f"{'Mes' if per.solo else 'Corte'}: {MESES[per.mes-1]} {anio}", ""),
             (f"Empresa: {emp_sel}", "filtro" if emp_sel != "Todas" else ""),
             (f"Datos hasta {MESES[ult-1]} {anio}", "ok")]
    if per.mes > ult:
        chips.append((f"Sin datos después de {MESES[ult-1]}", "warn"))
    zona_encabezado.html(E.encabezado(
        "Combustible", f"Relación de combustible · {per.texto} · cifras en millones de pesos", chips))

    if df.empty:
        st.info("No hay tanqueos para los filtros seleccionados.")
        return

    costo, gal = df["valor"].sum(), df["galones"].sum()
    st.html(E.kpis([
        dict(label=f"Costo {per.suf}", valor=E.millones(costo), sub=f"{len(df):,} tanqueos".replace(",", "."),
             icono="receipt", color=E.AZUL, fondo=E.AZUL_SUAVE),
        dict(label="Galones", valor=_galones(gal).replace(" gal", ""), sub="galones consumidos",
             icono="calendar", color=E.AZUL, fondo=E.AZUL_SUAVE),
        dict(label="Precio promedio", valor=_precio(costo / gal if gal else float("nan")),
             sub="costo total / galones", icono="wallet", color=E.AZUL, fondo=E.AZUL_SUAVE),
        dict(label="Vehículos", valor=str(df["placa"].nunique()),
             sub=f"promedio {E.millones(costo / df['placa'].nunique())} por vehículo",
             icono="check", color=E.VERDE, fondo=E.VERDE_SUAVE),
    ]))

    t1, t2, t3 = st.tabs(["Resumen", "Por vehículo", "Tanqueos"], on_change="rerun", key="cb_tabs")

    # ============ RESUMEN
    with t1:
        if t1.open:   # solo se calcula la pestaña abierta
            izq, der = st.columns([2.1, 1], gap="medium")
            with izq:
                with st.container(key="card_cb_mes"):
                    titulo_tarjeta("Mes a mes", "Barras: costo (millones de pesos) · línea: galones")
                    por = "empresa" if emp_sel == "Todas" and df["empresa"].nunique() > 1 else "fuente"
                    piv = base.pivot_table(index="mes", columns=por, values="valor", aggfunc="sum").reindex(range(1, 13))
                    fig = go.Figure()
                    for i, colname in enumerate(piv.columns):
                        fig.add_bar(x=MESES, y=piv[colname] / 1e6, name=str(colname),
                                    marker_color=COLOR_EMPRESA.get(colname, COLOR_PROV[i % len(COLOR_PROV)]),
                                    hovertemplate="%{y:,.1f} M")
                    galm = base.groupby("mes")["galones"].sum().reindex(range(1, 13))
                    fig.add_scatter(x=MESES, y=galm, name="Galones", yaxis="y2", mode="lines+markers",
                                    line=dict(color=E.NAVY, width=2.5), marker=dict(size=5),
                                    hovertemplate="%{y:,.0f} gal")
                    if per.solo:
                        fig.add_vrect(x0=per.mes - 1.5, x1=per.mes - 0.5, fillcolor=E.AZUL, opacity=.07, line_width=0)
                    else:
                        fig.add_vrect(x0=per.mes - 0.5, x1=11.5, fillcolor="#ffffff", opacity=.55, line_width=0)
                    fig.update_layout(barmode="stack", bargap=.35,
                                      yaxis=dict(tickformat=",.0f", ticksuffix=" M", rangemode="tozero"),
                                      yaxis2=dict(overlaying="y", side="right", showgrid=False, tickformat=",.0f",
                                                  rangemode="tozero"))
                    st.plotly_chart(estilo_grafico(fig, 360), width="stretch", config={"displayModeBar": False})
            with der:
                with st.container(key="card_cb_prov"):
                    titulo_tarjeta("Por proveedor", f"Participación en el costo · {per.texto}")
                    s = df.groupby("fuente")["valor"].sum().sort_values(ascending=False)
                    st.html(E.lista_participacion(list(s.items())))
                    titulo_tarjeta("Por tipo de combustible")
                    s = df.groupby("tipo")["valor"].sum().sort_values(ascending=False)
                    st.html(E.lista_participacion(list(s.items()), color="#8fb0e3"))

            izq, der = st.columns([2.1, 1], gap="medium")
            with izq:
                with st.container(key="card_cb_precio"):
                    titulo_tarjeta("Precio promedio por galón", "Por proveedor y mes · diésel")
                    d = base[base["tipo"] == "DIESEL"]
                    if d.empty:
                        d = base
                    pr = d.groupby(["mes", "fuente"]).agg(v=("valor", "sum"), g=("galones", "sum")).reset_index()
                    pr["precio"] = pr["v"] / pr["g"].where(pr["g"] > 0)
                    fig = go.Figure()
                    tabla_p = pr.pivot(index="mes", columns="fuente", values="precio")
                    for i, (fuente, g) in enumerate(pr.groupby("fuente")):
                        g = g.set_index("mes").reindex(range(1, 13))
                        color = COLOR_PROV[i % len(COLOR_PROV)]
                        # etiqueta del valor de cada mes: arriba si es el precio más alto del mes, abajo si no
                        pos = []
                        for m in range(1, 13):
                            fila = tabla_p.loc[m].dropna() if m in tabla_p.index else pd.Series(dtype=float)
                            pos.append("top center" if fila.empty or fila.idxmax() == fuente else "bottom center")
                        fig.add_scatter(x=MESES, y=g["precio"], name=fuente, mode="lines+markers+text",
                                        line=dict(color=color, width=2.5),
                                        text=["" if pd.isna(x) else f"${x:,.0f}".replace(",", ".") for x in g["precio"]],
                                        textposition=pos, textfont=dict(size=10, color=color, shadow="0px 0px 3px #ffffff"),
                                        cliponaxis=False, hovertemplate="$%{y:,.0f}/gal")
                    fig.update_layout(yaxis=dict(tickformat=",.0f", tickprefix="$"))
                    st.plotly_chart(estilo_grafico(fig, 320), width="stretch", config={"displayModeBar": False})
            with der:
                with st.container(key="card_cb_top"):
                    titulo_tarjeta("Vehículos con mayor costo", per.texto.capitalize())
                    s = df.groupby("placa")["valor"].sum().sort_values(ascending=False).head(8)
                    st.html(E.lista_participacion(list(s.items()), total=costo))

            # cruce con presupuesto y SAP de la compañía elegida (Todas = ALDC, la que tiene más flota)
            # solo lo ve quien ve toda la compañía o la gerencia dueña del sub-rubro Combustible
            ve_ppto = getattr(ctx, "ver_todo", True) or ctx.usuario in asignacion.dueno_de_sub(leer, "Combustible")
            emp_pp = "ALDC" if emp_sel == "Todas" else emp_sel
            if ve_ppto and emp_pp in ("ALDC", "ALSM") and not (prov or tipos or placas or contrato != "Todos"):
                vers = leer("SELECT DISTINCT version FROM presupuesto WHERE anio=? AND empresa=?",
                            (int(anio), emp_pp))["version"].tolist()
                v_ppto = "8+4" if "8+4" in vers else ("Inicial" if "Inicial" in vers else (vers[0] if vers else ""))
                pp = leer("SELECT mes, SUM(valor) v FROM presupuesto WHERE anio=? AND version=? AND empresa=? "
                          "AND sub_rubro='Combustible' GROUP BY mes", (int(anio), v_ppto, emp_pp)).set_index("mes")["v"]
                sap = leer("SELECT mes, SUM(valor) v FROM ejecucion WHERE anio=? AND empresa=? "
                           "AND sub_rubro='Combustible' GROUP BY mes", (int(anio), emp_pp)).set_index("mes")["v"]
                if not pp.empty:
                    rel = todo[(todo["anio"] == anio) & (todo["empresa"] == emp_pp)].groupby("mes")["valor"].sum()
                    with st.container(key="card_cb_ppto"):
                        titulo_tarjeta(f"Combustible {emp_pp}: presupuesto vs contabilizado vs relación",
                                       f"Presupuesto {v_ppto} del sub-rubro Combustible, ejecución contabilizada en SAP y "
                                       "costo según la relación de tanqueos · millones de pesos")
                        fig = go.Figure()
                        for nombre, serie, color in [("Presupuesto", pp, "#c9d4e4"), ("Contabilizado SAP", sap, E.AZUL),
                                                     ("Relación de combustible", rel, "#e8892c")]:
                            fig.add_bar(x=MESES, y=serie.reindex(range(1, 13)) / 1e6, name=nombre, marker_color=color,
                                        hovertemplate="%{y:,.1f} M")
                        fig.update_layout(barmode="group", bargap=.25,
                                          yaxis=dict(tickformat=",.0f", ticksuffix=" M"))
                        st.plotly_chart(estilo_grafico(fig, 300), width="stretch", config={"displayModeBar": False})
                        meses = per.meses()
                        a, b, r = pp.reindex(meses).sum(), sap.reindex(meses).sum(), rel.reindex(meses).sum()
                        st.html(f'<p class="nf-card-s">{per.texto.capitalize()}: presupuesto <b>{E.millones(a)}</b> · '
                                f'contabilizado <b>{E.millones(b)}</b> ({E.pct_txt(b / a if a else float("nan"))} del '
                                f'presupuesto) · relación de tanqueos <b>{E.millones(r)}</b>. La diferencia entre SAP y la '
                                f'relación suele ser por fechas de causación de las facturas.</p>')

    # ============ POR VEHÍCULO
    with t2:
        if t2.open:   # solo se calcula la pestaña abierta
            with st.container(key="card_cb_veh"):
                titulo_tarjeta("Consumo por vehículo", f"{per.texto.capitalize()} · haz clic en un encabezado para ordenar")
                v = df.groupby(["placa", "empresa", "contrato"]).agg(
                    tanqueos=("valor", "size"), galones=("galones", "sum"), costo=("valor", "sum"),
                    meses=("mes", "nunique"), ultimo=("fecha", "max"),
                    proveedor=("fuente", lambda s: s.value_counts().index[0])).reset_index()
                v["precio"] = v["costo"] / v["galones"].where(v["galones"] > 0)
                v["gal_tanqueo"] = v["galones"] / v["tanqueos"]
                v["part"] = v["costo"] / costo * 100
                v = v.sort_values("costo", ascending=False)
                st.dataframe(
                    v, hide_index=True, width="stretch", height=560,
                    column_order=["placa", "empresa", "contrato", "proveedor", "tanqueos", "galones", "gal_tanqueo",
                                  "costo", "precio", "part", "meses", "ultimo"],
                    column_config={
                        "placa": "Placa", "empresa": "Empresa", "contrato": "Contrato", "proveedor": "Proveedor principal",
                        "tanqueos": st.column_config.NumberColumn("Tanqueos", format="%d"),
                        "galones": st.column_config.NumberColumn("Galones", format="%.0f"),
                        "gal_tanqueo": st.column_config.NumberColumn("Gal / tanqueo", format="%.1f"),
                        "costo": st.column_config.NumberColumn("Costo", format="accounting"),
                        "precio": st.column_config.NumberColumn("$/galón", format="%.0f"),
                        "part": st.column_config.ProgressColumn("% del costo", format="%.1f%%", min_value=0,
                                                                max_value=float(max(v["part"].max(), 1))),
                        "meses": st.column_config.NumberColumn("Meses", format="%d"),
                        "ultimo": st.column_config.DateColumn("Último tanqueo", format="DD/MM/YYYY"),
                    })
                st.download_button("Descargar por vehículo", lambda d=v.drop(columns="part"): a_excel(d),
                                   f"combustible_por_vehiculo_{anio}.xlsx", icon=":material/download:")

    # ============ TANQUEOS
    with t3:
        if t3.open:   # solo se calcula la pestaña abierta
            with st.container(key="card_cb_det"):
                titulo_tarjeta("Detalle de tanqueos", f"{len(df):,} registros · {per.texto}".replace(",", "."))
                vista = df.sort_values("fecha", ascending=False)[
                    ["fecha", "empresa", "fuente", "placa", "tipo", "contrato", "galones", "valor"]]
                st.dataframe(vista, hide_index=True, width="stretch", height=560, column_config={
                    "fecha": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"), "empresa": "Empresa",
                    "fuente": "Proveedor", "placa": "Placa", "tipo": "Tipo", "contrato": "Contrato",
                    "galones": st.column_config.NumberColumn("Galones", format="%.2f"),
                    "valor": st.column_config.NumberColumn("Valor", format="accounting")})
                st.download_button("Descargar tanqueos", lambda d=vista: a_excel(d), f"tanqueos_{anio}.xlsx",
                                   icon=":material/download:")
