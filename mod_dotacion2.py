"""Dotación 2 (Personal): presupuesto de dotación y EPP's por elemento contra la ejecución de SAP (3.1 · Dotación y EPP's),
detallada con las salidas de inventario del SISPO. Desarrollo de Heidy, traído del panel de Streamlit.

La lógica está en dotacion_sispo.py; el presupuesto y la equivalencia de artículos se cargan en Administración → Dotación y EPP
→ Dotación 2. El SISPO es el mismo que alimenta Mantenimiento y Operaciones (Administración → Mantenimiento y Operaciones)."""
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import dotacion_sispo as D
import estilo as E
import logica_m as L
import sispo_fuente
from ui import EMPRESAS, MESES, MESES_L, estilo_grafico, selector_periodo, titulo_tarjeta


def mostrar(ctx):
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
    emp = c1.selectbox("Compañía del informe", empresas, index=empresas.index(empresa_pag) if empresa_pag in empresas else 0, key="dt2_emp_informe")
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
    sel = c2.selectbox("Mes del informe", opciones, index=por_defecto, key=f"dt2_mes_informe_{emp}",
                       format_func=lambda m: f"Acumulado a {MESES_L[meses_ej[-1] - 1]}" if m == 0 else MESES_L[m - 1].capitalize(),
                       help="Un mes = solo ese mes. Acumulado = todos los meses con ejecución del año.")
    meses = meses_ej if sel == 0 else [sel]
    txt = f"{MESES[meses[0] - 1]}–{MESES[meses[-1] - 1]}" if len(meses) > 1 else MESES[meses[0] - 1]
    sap_f = D.sap_filas(d["pool"], emp, anio)
    c3.download_button(f"Descargar ejecución de dotación · {emp} (Excel)",
                       lambda: D.libro_ejecucion(d["ppto"], d["ejec"], d["eq"], sap_f, d["sap_mes"], meses, f"Dotación y EPP's · {emp} {anio}"),
                       file_name=f"ejecucion_dotacion_{emp}_{anio}_{txt.replace(' ', '').replace('–', '-')}.xlsx", icon=":material/download:",
                       key=f"dt2_desc_{emp}")


def _dotacion(ctx):
    leer = ctx.leer
    ctx.con.crear_tablas(D.TABLAS)
    st.html(E.encabezado("Dotación 2 · Dotación y EPP's", "Presupuesto por elemento vs ejecución de SAP, detallada con las salidas de inventario del SISPO · cifras en millones de pesos", []))
    pp_all = leer("SELECT empresa, anio, seccion, elemento, precio, mes, cantidad, valor FROM dot_ppto")
    if pp_all.empty:
        st.info("Aún no hay presupuesto de dotación. Financiera lo carga en Administración → Dotación y EPP → Dotación 2.")
        return
    with st.container(key="card_filtros_dt2"):
        st.html('<div class="nf-ftit">Filtros del informe</div>')
        c = st.columns([0.9, 0.7, 1.4, 0.95], vertical_alignment="bottom")
        empresa = c[0].selectbox("Compañía", sorted(pp_all["empresa"].unique()), key="dt2_emp",
                                format_func=lambda e: f"{e} · {EMPRESAS.get(e, e)}")
        anios = sorted(pp_all.loc[pp_all["empresa"] == empresa, "anio"].unique(), reverse=True)
        anio = int(c[1].selectbox("Año", anios, key="dt2_anio"))
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
                   f"({n} sin asignar) y movimientos de SAP que el SISPO no detalla. Asígnalos en Administración → Dotación y EPP → Dotación 2; "
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
                     key=f"dt2_tab_{clave}", column_config={
                         "Presupuesto": st.column_config.NumberColumn(format="$ %.0f"), "Ejecutado": st.column_config.NumberColumn(format="$ %.0f"),
                         "Saldo": st.column_config.NumberColumn(format="$ %.0f"), "% ejecución": st.column_config.NumberColumn(format="percent"),
                         "Cant. ppto": st.column_config.NumberColumn(format="%.0f"), "Cant. ejecutada": st.column_config.NumberColumn(format="%.0f")})

    tr, td_, te = st.tabs(["Resumen", "Dotación", "EPP's"], on_change="rerun", key="dt2_tabs")
    with tr:
        if tr.open:
            st.html(kpis(t, f"Dotación y EPP's · {txt_per}"))
            with st.container(key="card_dt2_res"):
                titulo_tarjeta("Por grupo", f"{empresa} {anio} · {txt_per}. Cada grupo se abre en sus elementos. Barra: relleno = ejecutado, marca = presupuesto.")
                grupos = []
                for nombre, sec in (("Dotación", D.DOTACION), ("EPP's", D.EPPS)):
                    x = t[t["seccion"] == sec]
                    grupos.append(dict(nombre=nombre, presupuesto=float(x["presupuesto"].sum()), ejecutado=float(x["ejecutado"].sum()),
                                       sub=f"{int(((x['presupuesto'] > 0) | (x['ejecutado'] > 0)).sum())} elementos", hijos=filas_visual(x)))
                st.html(E.tabla_visual(grupos, "Grupo", abierto=False))
                participacion(t, "Participación de cada elemento en lo ejecutado")
            with st.container(key="card_dt2_mes"):
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
                st.plotly_chart(fig2, width="stretch", key="dt2_graf_mes")
            with st.container(key="card_dt2_desc"):
                titulo_tarjeta("Archivo de ejecución", "Excel con Resumen (presupuesto y ejecución mes a mes de Dotación y EPP's) y una hoja de movimientos "
                               "para Dotación y otra solo para EPPs. Elige la compañía y el mes del archivo.")
                _descarga_dotacion(leer, pp_all, empresa, anio, per)
    for pestana, sec, nombre in ((td_, D.DOTACION, "Dotación"), (te, D.EPPS, "EPP's")):
        with pestana:
            if pestana.open:
                x = t[t["seccion"] == sec]
                st.html(kpis(x, f"{nombre} · {txt_per}"))
                with st.container(key=f"card_dt2_{sec}"):
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
