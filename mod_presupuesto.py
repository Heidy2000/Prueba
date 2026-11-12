"""Módulo Presupuesto: presupuesto vs ejecución de todas las clases del PyG (1.0 a 7.0) con subtotales."""
from datetime import datetime
from io import BytesIO

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import asignacion
import cargas
import datos
import estilo as E
import lineas
from ui import EMPRESAS, MESES, MESES_L, SIN_DUENO, a_excel, estilo_grafico, titulo_tarjeta

COLS_RAW = ["Presupuesto anual", "Presupuesto al corte", "Ejecutado al corte"]


def _es_costo(*valores) -> bool:
    """Una fila en signo de resultado es costo si su primer valor distinto de cero es negativo."""
    for v in valores:
        if v is not None and not pd.isna(v) and abs(v) > 0.5:
            return v < 0
    return False


def _cuadre(cargado: pd.DataFrame, informe: pd.DataFrame, periodo_txt: str, filtrado: bool):
    """Cuadre entre la ejecución cargada (todo lo que hay en la base para el periodo) y lo que muestra el informe."""
    c = cargado.assign(codigo=cargado["codigo"].fillna(""))
    a = c.groupby("codigo").agg(n=("valor", "size"), cargado=("valor", "sum"))
    b = informe.groupby("codigo")["valor"].sum().rename("informe")
    t = a.join(b, how="outer").fillna(0)
    t["dif"] = t["cargado"] - t["informe"]
    dif_total = t["dif"].sum()
    titulo = (f"Cuadre con la ejecución cargada · {periodo_txt} · "
              + ("todo cuadra" if abs(dif_total) < 1 and not (t["dif"].abs() >= 1).any()
                 else f"diferencia {E.millones(dif_total)}"))
    with st.expander(titulo, icon=":material/rule:"):
        filas = []
        for cod, r in t.reindex(sorted(t.index, key=lambda x: (x == "", cargas.orden_clase(x)))).iterrows():
            if cod == "":
                motivo, nombre = "Sin clasificar: asígnales clase en Administración → Revisar y reclasificar", \
                    "Sin clasificar"
            elif cod not in cargas.RUBROS:
                motivo, nombre = "Código que no existe en la tabla de rubros: corrígelo al reclasificar", cod
            elif abs(r["dif"]) >= 1:
                motivo, nombre = ("Fuera del filtro de gerencia / rubro elegido" if filtrado
                                  else "Revisar"), cargas.RUBROS[cod]
            else:
                motivo, nombre = "", cargas.RUBROS[cod]
            filas.append({"Clase": nombre, "Movimientos": int(r["n"]), "Cargado (signo SAP)": r["cargado"],
                          "En el informe": r["informe"], "Diferencia": r["dif"], "Motivo": motivo})
        filas.append({"Clase": "Total", "Movimientos": int(t["n"].sum()), "Cargado (signo SAP)": t["cargado"].sum(),
                      "En el informe": t["informe"].sum(), "Diferencia": dif_total, "Motivo": ""})
        st.dataframe(pd.DataFrame(filas), hide_index=True, width="stretch", column_config={
            k: st.column_config.NumberColumn(format="accounting")
            for k in ["Cargado (signo SAP)", "En el informe", "Diferencia"]})
        st.caption("Cargado = todos los movimientos de la base de ejecución en el periodo, en el signo de SAP (ingresos "
                   "en negativo, costos en positivo), igual que la suma del archivo. La versión de presupuesto elegida "
                   "no cambia la ejecución.")


def _orden(rubro: str) -> tuple:
    return cargas.orden_clase(str(rubro).split(" ")[0])


def mostrar(ctx):
    leer, usuario, es_admin = ctx.leer, ctx.usuario, ctx.es_admin
    ver_todo = getattr(ctx, "ver_todo", es_admin)   # Financiera o perfiles con presupuesto completo
    zona_encabezado = st.container()   # se llena después de leer los filtros

    # ------------------------------------------------------------------ filtros del informe
    with st.container(key="card_filtros"):
        st.html('<div class="nf-ftit">Filtros del informe</div>')
        hay = set(leer("SELECT DISTINCT empresa FROM presupuesto UNION SELECT DISTINCT empresa FROM ejecucion")
                  ["empresa"].fillna("ALDC"))
        empresas = list(EMPRESAS)          # siempre se muestran las dos compañías
        anchos = [0.7, 1.0, 1.3, 1.5, 0.95] if ver_todo else [0.7, 1.0, 1.5, 0.95]
        if len(empresas) > 1:
            anchos = [0.8] + anchos
        cols = iter(st.columns(anchos, vertical_alignment="bottom"))
        empresa = next(cols).selectbox("Compañía", empresas, key="pp_empresa",
                                       help=" · ".join(f"{e}: {EMPRESAS[e]}" for e in empresas)) \
            if len(empresas) > 1 else empresas[0]

        if empresa not in hay:
            st.info(f"Aún no hay presupuesto ni ejecución de **{empresa} · {EMPRESAS[empresa]}**. Financiera los "
                    "carga en Administración → Presupuesto, eligiendo esa compañía arriba.", icon=":material/info:")
            zona_encabezado.html(E.encabezado(f"Presupuesto · {empresa}", EMPRESAS[empresa], []))
            return
        anios = leer("SELECT DISTINCT anio FROM presupuesto WHERE empresa=? ORDER BY anio DESC",
                     (empresa,))["anio"].tolist() or [datetime.now().year]
        con_ej = leer("SELECT MAX(anio) a FROM ejecucion WHERE empresa=?", (empresa,))["a"][0]
        anio = next(cols).selectbox("Año", anios, key=f"pp_anio_{empresa}",
                                    index=anios.index(con_ej) if con_ej in anios else 0)

        versiones = leer("SELECT DISTINCT version FROM presupuesto WHERE anio=? AND empresa=?",
                         (anio, empresa))["version"].tolist()
        orden_v = {"Inicial": 0, "4+8": 1, "8+4": 2}
        versiones = sorted(versiones, key=lambda v: (orden_v.get(v, 9), v)) or ["Inicial"]
        if st.session_state.get("pp_version") not in versiones:     # por defecto la versión más reciente (8+4)
            st.session_state["pp_version"] = "8+4" if "8+4" in versiones else versiones[-1]
        version = next(cols).selectbox("Comparar contra", versiones, key="pp_version",
                                       format_func=lambda v: f"Presupuesto {v}",
                                       help="Inicial: el aprobado. 4+8: 4 meses reales + 8 presupuestados. "
                                            "8+4: 8 meses reales + 4 presupuestados (cierre esperado del año).")

        asig = asignacion.mapa(leer)
        todas = sorted({g for g in asig.values() if g})
        if ver_todo:
            opcion = next(cols).selectbox("Gerencia", ["Todas las gerencias"] + todas + [SIN_DUENO])
            gerencias = todas + [SIN_DUENO] if opcion == "Todas las gerencias" else [opcion]
        else:
            gerencias = [usuario]

        modo = next(cols).segmented_control("Periodo", ["Acumulado", "Solo el mes"], default="Acumulado") \
            or "Acumulado"
        SOLO = modo == "Solo el mes"

        ult = leer("SELECT MAX(mes) m FROM ejecucion WHERE anio=? AND codigo<>'' AND empresa=?",
                   (anio, empresa))["m"][0]
        ult = int(ult) if pd.notna(ult) else 12
        corte = next(cols).selectbox("Mes" if SOLO else "Hasta", list(range(1, 13)), index=ult - 1,
                                     format_func=lambda m: MESES_L[m - 1].capitalize())

        # presupuesto de la versión elegida y ejecución con los nombres de esa versión
        ppto = leer("SELECT * FROM presupuesto WHERE anio=? AND version=? AND empresa=?", (anio, version, empresa))
        cols_e = ", ".join(sorted(ctx.con.columnas("ejecucion") - {"raw", "id"}))
        ejec_todo = leer(f"SELECT {cols_e} FROM ejecucion WHERE anio=? AND empresa=?", (anio, empresa))
        # líneas de negocio (ALSM): línea de cada movimiento y reparto de lo compartido
        CON_LINEAS = lineas.hay_lineas(leer, empresa)
        if CON_LINEAS:
            ejec_todo["linea"] = lineas.linea_ejecucion(ejec_todo, lineas.mapa_ceco(leer, empresa))
            PESOS_MES = lineas.pesos_ingreso(ejec_todo)
            ejec_todo = lineas.distribuir(ejec_todo, PESOS_MES)
            DIST = leer("SELECT anio, linea, pct FROM distribucion WHERE empresa=?", (empresa,))

            def repartir_ppto(df: pd.DataFrame, ver: str) -> pd.DataFrame:
                df = df.assign(linea=df["linea"].fillna("").replace("", lineas.SIN_LINEA))
                pa = lineas.pesos_anuales(DIST, set(df["linea"]) - set(lineas.COMPARTIDAS))
                return lineas.distribuir(df, PESOS_MES, pa, lineas.MESES_REALES.get(ver, 0))
            ppto = repartir_ppto(ppto, version)
        pendientes = ejec_todo[ejec_todo["codigo"].fillna("") == ""]
        ejec = ejec_todo[ejec_todo["codigo"].fillna("") != ""].copy()
        ejec["rubro"] = ejec["codigo"].map(cargas.rubro_de)
        homol = leer("SELECT * FROM homologacion")
        ejec = datos.homologar(ejec, ppto, homol, version)
        # cada gerencia solo ve las líneas (código + sub-rubro) asignadas a ella
        ppto["gerencia"] = asignacion.columna(ppto, asig)
        ejec["gerencia"] = asignacion.columna(ejec, asig)
        ppto = ppto[ppto["gerencia"].isin(gerencias)]
        ejec = ejec[ejec["gerencia"].isin(gerencias)]

        cols2 = st.columns([1, 1.4, 1.2] if CON_LINEAS else [1, 1.4], vertical_alignment="bottom")
        catalogo = pd.concat([ppto[["codigo", "rubro", "sub_rubro"]], ejec[["codigo", "rubro", "sub_rubro"]]]) \
            .drop_duplicates()
        rubros = sorted(catalogo["rubro"].unique(), key=_orden)
        rubros_f = cols2[0].multiselect("Rubro", rubros, placeholder="Todos los rubros")
        opciones_sub = catalogo[catalogo["rubro"].isin(rubros_f)] if rubros_f else catalogo
        subs_f = cols2[1].multiselect("Sub-rubro", sorted(opciones_sub["sub_rubro"].unique()),
                                      placeholder="Todos los sub-rubros")
        lineas_f = []
        if CON_LINEAS:
            ops_l = sorted(set(ppto["linea"]) | set(ejec["linea"]),
                           key=lambda x: (x in lineas.COMPARTIDAS or x == lineas.SIN_LINEA, x))
            lineas_f = cols2[2].multiselect("Línea de negocio", ops_l, placeholder="Todas las líneas",
                                            key="pp_lineas",
                                            help="Administrativo, operativo, mantenimiento y comercial se reparten "
                                                 "entre las líneas según su participación en el ingreso de cada mes "
                                                 "(presupuesto: % de la hoja 2 en los meses proyectados).")
        if rubros_f:
            ppto, ejec = ppto[ppto["rubro"].isin(rubros_f)], ejec[ejec["rubro"].isin(rubros_f)]
        if subs_f:
            ppto, ejec = ppto[ppto["sub_rubro"].isin(subs_f)], ejec[ejec["sub_rubro"].isin(subs_f)]
        if lineas_f:
            ppto, ejec = ppto[ppto["linea"].isin(lineas_f)], ejec[ejec["linea"].isin(lineas_f)]

    # ------------------------------------------------------------------ periodo y cálculos
    MES_TXT = MESES_L[corte - 1]
    SUF = f"de {MES_TXT}" if SOLO else "al corte"
    PERIODO_TXT = f"solo {MES_TXT}" if SOLO else f"enero a {MES_TXT}"
    ETIQUETAS = ["Presupuesto anual", f"Presupuesto {SUF}", f"Ejecutado {SUF}", "Variación"]
    # vista de resultados (EBITDA, resultado) solo para quien ve toda la compañía: con una parte del PyG no tiene sentido
    CON_PYG = ver_todo and "1.0" in set(ppto["codigo"]) | set(ejec["codigo"])

    def en_periodo(df: pd.DataFrame) -> pd.DataFrame:
        return df[df["mes"] == corte] if SOLO else df[df["mes"] <= corte]

    # ejecución que no cae en ninguna clase del PyG (sin clasificar o código que no existe): igual cuenta en el
    # resultado del ejercicio, para que la utilidad cuadre con la base de SAP. Solo con la compañía completa.
    sin_filtro = ver_todo and not (rubros_f or subs_f or lineas_f) and len(gerencias) == len(todas) + 1
    fuera = ejec_todo[~ejec_todo["codigo"].fillna("").isin(list(cargas.CLASES))] if sin_filtro \
        else ejec_todo.iloc[0:0]
    FUERA_EC = float(en_periodo(fuera)["valor"].sum())          # signo SAP
    FUERA_N = len(en_periodo(fuera))

    def resumen(por: list[str]) -> pd.DataFrame:
        """Presupuesto anual, presupuesto y ejecución del periodo en signo SAP, por nivel."""
        pa = ppto.groupby(por)["valor"].sum()
        pc = en_periodo(ppto).groupby(por)["valor"].sum()
        ec = en_periodo(ejec).groupby(por)["valor"].sum()
        ea = ejec[ejec["mes"] <= corte].groupby(por)["valor"].sum()
        return pd.DataFrame({"Presupuesto anual": pa, "Presupuesto al corte": pc,
                             "Ejecutado al corte": ec, "_acum": ea}).fillna(0)

    def con_metricas(t: pd.DataFrame) -> pd.DataFrame:
        """Agrega columnas para mostrar: valores en positivo, variación (ejecutado − presupuesto), a favor/en contra,
        % y estado."""
        m = t.apply(lambda r: pd.Series(E.metricas(r["Presupuesto anual"], r["Presupuesto al corte"],
                                                   r["Ejecutado al corte"], r["_acum"])), axis=1)
        return t.join(m)

    def total_clases(clases: list[str]) -> dict:
        """Totales en signo de resultado (ingreso +, gasto −) para un conjunto de clases."""
        p = en_periodo(ppto)
        e = en_periodo(ejec)
        return dict(pa=-ppto.loc[ppto["codigo"].isin(clases), "valor"].sum(),
                    pc=-p.loc[p["codigo"].isin(clases), "valor"].sum(),
                    ec=-e.loc[e["codigo"].isin(clases), "valor"].sum())

    def para_excel(df: pd.DataFrame) -> bytes:
        df = con_metricas(df.set_index([c for c in ["gerencia", "rubro", "sub_rubro"] if c in df.columns]))
        out = df[["pa", "pc", "ec", "var", "pct", "estado"]].reset_index().rename(columns={
            "gerencia": "Gerencia", "rubro": "Rubro", "sub_rubro": "Sub-rubro", "pa": ETIQUETAS[0],
            "pc": ETIQUETAS[1], "ec": ETIQUETAS[2], "var": "Variación (ejecutado − presupuesto)", "pct": "% ejecución",
            "estado": "Estado"})
        buf = BytesIO()
        out.to_excel(buf, index=False)
        return buf.getvalue()

    # ------------------------------------------------------------------ encabezado
    filtro_carga = "detalle LIKE 'ALSM%'" if empresa == "ALSM" else "detalle NOT LIKE 'ALSM%'"
    carga = leer(f"SELECT fecha FROM cargas WHERE tipo='Ejecución' AND {filtro_carga} ORDER BY rowid DESC LIMIT 1")
    chips = ([(f"{empresa} · {EMPRESAS[empresa]}", "filtro")] if len(empresas) > 1 else []) + \
        [(f"Presupuesto {version}", "filtro"),
             ((f"Mes: {MESES[corte - 1]} {anio}") if SOLO else (f"Corte: {MESES[corte - 1]} {anio}"), "")]
    if es_admin and not pendientes.empty:
        chips.append((f"{len(pendientes)} movimientos sin clasificar · {E.millones(pendientes['valor'].sum())}", "warn"))
    if rubros_f or subs_f:
        partes = ([f"{len(rubros_f)} rubro{'s' if len(rubros_f) > 1 else ''}"] if rubros_f else []) + \
                 ([f"{len(subs_f)} sub-rubro{'s' if len(subs_f) > 1 else ''}"] if subs_f else [])
        chips.append(("Filtro: " + " · ".join(partes), "filtro"))
    if not carga.empty:
        chips.append((f"SAP actualizado {pd.Timestamp(carga['fecha'][0]):%d/%m/%Y}", "ok"))
    if es_admin:
        lineas_v = leer("SELECT DISTINCT codigo, sub_rubro FROM presupuesto WHERE anio=? AND version=? AND empresa=?",
                      (anio, version, empresa))
        n_sin = int((asignacion.columna(lineas_v, asig) == SIN_DUENO).sum())
        if n_sin:
            chips.append((f"{n_sin} sub-rubros sin dueño", "warn"))
    zona_encabezado.html(E.encabezado(
        f"Presupuesto {anio}" + (f" · {empresa}" if len(empresas) > 1 else "") + ("" if ver_todo else f" · {usuario}"),
        (f"Presupuesto {version} vs ejecución de {MES_TXT} {anio}" if SOLO
         else f"Presupuesto {version} vs ejecución acumulada de enero a {MES_TXT}") + " · cifras en millones de pesos",
        chips))

    nombres_tabs = ["Resumen"] + (["Estado de resultados"] if CON_PYG else []) + \
        ["Presupuesto por rubro", "Movimientos", "Comparar presupuestos"]
    if st.session_state.get("pp_tabs") not in nombres_tabs:
        st.session_state.pop("pp_tabs", None)
    tabs = st.tabs(nombres_tabs, on_change="rerun", key="pp_tabs")
    t_res, t_rub, t_mov, t_cmp = tabs[0], tabs[-3], tabs[-2], tabs[-1]

    # ========== RESUMEN
    with t_res:
        if t_res.open:   # solo se calcula la pestaña abierta
            if ppto.empty and ejec.empty:
                st.info("Aún no hay presupuesto cargado"
                        + (" (pestaña Administración)." if es_admin
                           else " o tu gerencia aún no tiene sub-rubros asignados. Pide a Financiera que te los "
                                "asigne en Administración → Asignación por gerencia."))
            else:
                if CON_PYG:
                    ing = total_clases(["1.0"])
                    cyg = total_clases([c for c in cargas.EBITDA if c != "1.0"])
                    ebi = total_clases(cargas.SUBTOTALES["EBITDA"])
                    res = total_clases(cargas.SUBTOTALES["Resultado del ejercicio"])
                    res["ec"] -= FUERA_EC                 # incluye lo que aún no tiene clase (cuadra con SAP)

                    def tarjeta(nombre, d, icono, gasto=False):
                        # d en signo de resultado; para costos se muestra en positivo
                        ec, pc = (-d["ec"], -d["pc"]) if gasto else (d["ec"], d["pc"])
                        pct = ec / pc if pc else float("nan")
                        var = ec - pc                       # ejecutado − presupuesto
                        favor = -var if gasto else var      # en costos, ejecutar menos es a favor
                        color = E.VERDE if favor >= 0 else E.ROJO
                        return dict(label=f"{nombre} {SUF}", valor=E.millones(ec),
                                    sub=f'Ppto {E.millones(pc)} · <b style="color:{color}">'
                                        f'{"+" if var >= 0 else ""}{E.millones(var)}</b> ({E.pct_txt(pct)})',
                                    icono=icono, color=color, fondo=E.VERDE_SUAVE if favor >= 0 else E.ROJO_SUAVE)

                    k_ebi = tarjeta("EBITDA", ebi, "piggy")
                    if ing["ec"]:
                        k_ebi["sub"] += f' · margen {E.pct_txt(ebi["ec"] / ing["ec"])}'
                    k_res = tarjeta("Resultado del ejercicio", res, "check")
                    if FUERA_N:
                        k_res["sub"] += f' · incluye {E.millones(-FUERA_EC)} sin clasificar'
                    st.html(E.kpis([tarjeta("Ingresos", ing, "wallet"), tarjeta("Costos y gastos", cyg, "receipt", True),
                                    k_ebi, k_res]))
                    if FUERA_N:
                        st.warning(f"{FUERA_N} movimientos ({E.millones(FUERA_EC)} en signo SAP) aún no tienen clase: "
                                   "suman en el Resultado del ejercicio para que cuadre con SAP, pero no en ingresos, "
                                   "costos ni EBITDA. Clasifícalos en Administración → Revisar y reclasificar.",
                                   icon=":material/warning:")
                else:
                    tot = con_metricas(resumen(["anio"])).iloc[0]
                    tono = {"nf-bad": (E.ROJO, E.ROJO_SUAVE), "nf-warn": (E.AMBAR, E.AMBAR_SUAVE)}.get(
                        tot["clase"], (E.VERDE, E.VERDE_SUAVE))
                    st.html(E.kpis([
                        dict(label="Presupuesto anual", valor=E.millones(tot["pa"]),
                             sub=f"{ppto['sub_rubro'].nunique()} sub-rubros", icono="wallet", color=E.AZUL,
                             fondo=E.AZUL_SUAVE),
                        dict(label=ETIQUETAS[1], valor=E.millones(tot["pc"]), sub=PERIODO_TXT.capitalize(),
                             icono="calendar", color=E.AZUL, fondo=E.AZUL_SUAVE),
                        dict(label=ETIQUETAS[2], valor=E.millones(tot["ec"]),
                             sub=f'<b style="color:{tono[0]}">{E.pct_txt(tot["pct"])}</b> del presupuesto {SUF}',
                             icono="receipt", color=tono[0], fondo=tono[1]),
                        dict(label="Disponible del año", valor=E.millones(tot["disp"]),
                             sub=f"Variación {SUF}: {E.millones(tot['var'])}",
                             icono="piggy" if tot["disp"] >= 0 else "alert",
                             color=E.AZUL if tot["disp"] >= 0 else E.ROJO,
                             fondo=E.AZUL_SUAVE if tot["disp"] >= 0 else E.ROJO_SUAVE),
                    ]))

                izq, der = st.columns([2.1, 1], gap="medium")
                with izq:
                    with st.container(key="card_mes"):
                        if CON_PYG:
                            metrica = st.segmented_control(
                                "Indicador", ["Ingresos", "Costos y gastos", "EBITDA", "Resultado"], default="EBITDA",
                                key="pp_metrica", label_visibility="collapsed") or "EBITDA"
                            clases = {"Ingresos": ["1.0"], "Costos y gastos": [c for c in cargas.EBITDA if c != "1.0"],
                                      "EBITDA": cargas.SUBTOTALES["EBITDA"],
                                      "Resultado": cargas.SUBTOTALES["Resultado del ejercicio"]}[metrica]
                            sg = 1 if metrica == "Costos y gastos" else -1
                        else:
                            metrica, clases, sg = "Ejecución", None, 1
                        titulo_tarjeta(f"Mes a mes · {metrica}",
                                       "Barras: valor del mes · líneas: acumulado (millones de pesos)")
                        pp_ = ppto if clases is None else ppto[ppto["codigo"].isin(clases)]
                        ej_ = ejec if clases is None else ejec[ejec["codigo"].isin(clases)]
                        if clases is None:
                            sg = -1 if E.es_ingreso(pp_["valor"].sum()) else 1
                        pm = sg * pp_.groupby("mes")["valor"].sum().reindex(range(1, 13), fill_value=0) / 1e6
                        em = sg * ej_.groupby("mes")["valor"].sum().reindex(range(1, 13), fill_value=0) / 1e6
                        em[em.index > ult] = None
                        fig = go.Figure()
                        fig.add_bar(x=MESES, y=pm, name="Presupuesto", marker_color="#c9d4e4",
                                    marker_line_width=0, hovertemplate="%{y:,.1f} M")
                        fig.add_bar(x=MESES, y=em, name="Ejecutado", marker_color=E.AZUL,
                                    marker_line_width=0, hovertemplate="%{y:,.1f} M")
                        fig.add_scatter(x=MESES, y=pm.cumsum(), name="Presupuesto acumulado", yaxis="y2",
                                        line=dict(color="#8a9bb5", dash="dot", width=2), hovertemplate="%{y:,.0f} M")
                        fig.add_scatter(x=MESES, y=em.cumsum(), name="Ejecutado acumulado", yaxis="y2",
                                        line=dict(color=E.NAVY, width=2.5), mode="lines+markers",
                                        marker=dict(size=5), hovertemplate="%{y:,.0f} M")
                        if SOLO:
                            fig.add_vrect(x0=corte - 1.5, x1=corte - 0.5, fillcolor=E.AZUL, opacity=.07, line_width=0)
                        fig.update_layout(barmode="group", bargap=.25,
                                          yaxis=dict(tickformat=",.0f", ticksuffix=" M", rangemode="tozero"),
                                          yaxis2=dict(overlaying="y", side="right", showgrid=False,
                                                      tickformat=",.0f", ticksuffix=" M", rangemode="tozero"))
                        st.plotly_chart(estilo_grafico(fig, 380), width="stretch", config={"displayModeBar": False})
                with der:
                    with st.container(key="card_rubros"):
                        titulo_tarjeta("Ejecución por rubro", f"Ejecutado vs presupuesto · {PERIODO_TXT}")
                        r = resumen(["rubro"])
                        r = r.loc[sorted(r.index, key=_orden)]
                        st.html(E.lista_avance([(n, x["Presupuesto anual"], x["Presupuesto al corte"],
                                                 x["Ejecutado al corte"]) for n, x in r.iterrows()]))

                izq, der = st.columns([2.1, 1], gap="medium")
                with izq:
                    with st.container(key="card_top"):
                        titulo_tarjeta("Mayores desviaciones en contra",
                                       f"Gastos por encima o ingresos por debajo del presupuesto · {PERIODO_TXT}")
                        top = con_metricas(resumen(["rubro", "sub_rubro"])).reset_index()
                        top = top[top["fav"] < -50_000].nsmallest(8, "fav")
                        if top.empty:
                            st.success("No hay desviaciones en contra en el periodo.")
                        else:
                            st.html(E.lista_desviaciones(top, SUF))
                with der:
                    with st.container(key="card_estado"):
                        titulo_tarjeta("Estado de los sub-rubros", f"Según su % de ejecución · {PERIODO_TXT}")
                        sr = con_metricas(resumen(["rubro", "sub_rubro"]))
                        grupos = {"En rango": "En rango / cumplido", "Cumplido": "En rango / cumplido",
                                  "Alerta": "Alerta", "Sobreejecutado": "Desviación en contra",
                                  "Por debajo": "Desviación en contra", "Sin presupuesto": "Sin presupuesto",
                                  "Sin movimiento": "Sin movimiento"}
                        est = sr["estado"].map(grupos).value_counts()
                        etiquetas = ["En rango / cumplido", "Alerta", "Desviación en contra", "Sin presupuesto",
                                     "Sin movimiento"]
                        colores = [E.VERDE, "#e8892c", E.ROJO, "#8b1e1e", "#c3ccd8"]
                        vals = [int(est.get(k, 0)) for k in etiquetas]
                        dona = go.Figure(go.Pie(labels=etiquetas, values=vals, hole=.72, marker_colors=colores,
                                                sort=False, textinfo="none",
                                                hovertemplate="%{label}: %{value}<extra></extra>"))
                        dona.add_annotation(text=f"<b style='font-size:26px'>{sum(vals)}</b><br>sub-rubros",
                                            showarrow=False, font=dict(family=E.FUENTE, color=E.TEXTO))
                        dona.update_layout(showlegend=False)
                        st.plotly_chart(estilo_grafico(dona, 200), width="stretch", config={"displayModeBar": False})
                        st.html("".join(
                            f'<div style="display:flex;justify-content:space-between;font-size:.82rem;margin:.25rem 0">'
                            f'<span><i style="display:inline-block;width:8px;height:8px;border-radius:50%;background:{c};'
                            f'margin-right:.5rem"></i>{t}</span><b>{v}</b></div>'
                            for t, c, v in zip(etiquetas, colores, vals)))

                if ver_todo:
                    with st.container(key="card_ger"):
                        # solo costos y gastos: se excluyen los ingresos (1.0) y las clases que netean como ingreso
                        # (p. ej. suplidos facturados), porque mezclarlos con gastos distorsiona el %
                        gg = resumen(["gerencia", "codigo", "rubro"]).reset_index()
                        por_clase = gg.groupby("codigo")["Presupuesto anual"].sum()
                        excluidas = sorted({"1.0"} | set(por_clase[por_clase < 0].index), key=cargas.orden_clase)
                        gg = gg[~gg["codigo"].isin(excluidas)]
                        titulo_tarjeta("Ejecución por gerencia",
                                       f"Costos y gastos de los sub-rubros de cada gerencia · {PERIODO_TXT}. "
                                       "Haz clic en una gerencia para ver sus rubros.")

                        def vals_g(pa, pc, ec):
                            return [pa, pc, ec, ec - pc, (ec / pc if pc > 0 else None)]

                        tot = gg[COLS_RAW].sum()
                        pa_t, pc_t, ec_t = tot["Presupuesto anual"], tot["Presupuesto al corte"], tot["Ejecutado al corte"]
                        pct_t = ec_t / pc_t if pc_t > 0 else float("nan")
                        var_t = ec_t - pc_t                 # negativo: se gastó menos de lo presupuestado
                        color_t = E.color_barra(pct_t, ec_t, False)
                        st.html(E.kpis([
                            dict(label="Presupuesto anual", valor=E.millones(pa_t),
                                 sub=f"{gg['gerencia'].nunique()} gerencias", icono="wallet", color=E.AZUL,
                                 fondo=E.AZUL_SUAVE),
                            dict(label=f"Presupuesto {SUF}", valor=E.millones(pc_t), sub=PERIODO_TXT.capitalize(),
                                 icono="calendar", color=E.AZUL, fondo=E.AZUL_SUAVE),
                            dict(label=f"Ejecutado {SUF}", valor=E.millones(ec_t),
                                 sub=f'Variación <b style="color:{E.VERDE if var_t <= 0 else E.ROJO}">'
                                     f'{"+" if var_t >= 0 else ""}{E.millones(var_t)}</b>',
                                 icono="receipt", color=E.AZUL, fondo=E.AZUL_SUAVE),
                            dict(label="% de ejecución", valor=f'<span style="color:{color_t}">{E.pct_txt(pct_t)}</span>',
                                 sub=f"Del presupuesto {SUF}", icono="check" if var_t <= 0 else "alert",
                                 color=color_t, fondo=E.VERDE_SUAVE if var_t <= 0 else E.ROJO_SUAVE),
                        ]))

                        nodos_g = []
                        orden_g = gg.groupby("gerencia")["Presupuesto anual"].sum().sort_values(ascending=False).index
                        for g in sorted(orden_g, key=lambda x: x == SIN_DUENO):
                            t = gg[gg["gerencia"] == g]
                            t = t.reindex(t["codigo"].map(cargas.orden_clase).sort_values().index)
                            hijos = [dict(nombre=rub, clase="linea", dif_col=3, dif_inv=True, valores=vals_g(*v))
                                     for rub, v in zip(t["rubro"], t[COLS_RAW].values)]
                            s_ = t[COLS_RAW].sum()
                            nodos_g.append(dict(nombre=g, clase="grupo", dif_col=3, dif_inv=True, hijos=hijos,
                                                sub=f"{len(hijos)} rubros", valores=vals_g(*s_)))
                        nodos_g.append(dict(nombre="Total", clase="saldo", dif_col=3, dif_inv=True,
                                            valores=vals_g(pa_t, pc_t, ec_t)))
                        st.html(E.tabla_meses(nodos_g, ["Presupuesto anual", f"Presupuesto {SUF}", f"Ejecutado {SUF}",
                                                        "Variación", "% ejecución"], formatos=["m", "m", "m", "m", "p"]))
                        st.caption("Variación = ejecutado − presupuesto: negativa (en verde) cuando se gastó menos de lo "
                                   "presupuestado; positiva (en rojo) cuando se gastó más. No incluye: "
                                   + ", ".join(cargas.RUBROS.get(c, c) for c in excluidas)
                                   + " (son ingresos o recuperaciones; mezclarlos con gastos distorsiona el %).")

                if ver_todo and CON_LINEAS:
                    with st.container(key="card_lineas"):
                        titulo_tarjeta("Resultado por línea de negocio",
                                       f"{PERIODO_TXT.capitalize()} · con administrativo, operativo, mantenimiento y "
                                       "comercial repartidos según el ingreso de cada mes. Haz clic en una línea para "
                                       "ver ingresos y costos.")
                        p_l, e_l = en_periodo(ppto), en_periodo(ejec)

                        def suma(df, linea, ingreso):
                            d = df[(df["linea"] == linea) & ((df["codigo"] == "1.0") == ingreso)]
                            return -float(d["valor"].sum())          # signo de resultado

                        def fila(nombre, pc, ec, costo=False, hijos=None, clase="linea"):
                            var = ec - pc
                            if costo:            # costos: ejecutado − presupuesto en valor absoluto
                                var = -var
                            pct = (ec / pc) if pc else None
                            return dict(nombre=nombre, clase=clase, dif_col=2, dif_inv=costo,
                                        valores=[pc, ec, var, pct], hijos=hijos or [])
                        nodos_l, tot = [], [0.0, 0.0]
                        orden_l = sorted(set(p_l["linea"]) | set(e_l["linea"]),
                                         key=lambda x: (x in lineas.COMPARTIDAS or x == lineas.SIN_LINEA,
                                                        -(suma(e_l, x, True) + suma(e_l, x, False))))
                        for lin in orden_l:
                            pi, ei = suma(p_l, lin, True), suma(e_l, lin, True)
                            pg, eg = suma(p_l, lin, False), suma(e_l, lin, False)
                            if not any(abs(v) > 0.5 for v in (pi, ei, pg, eg)):
                                continue
                            tot[0] += pi + pg
                            tot[1] += ei + eg
                            nodos_l.append(fila(lin, pi + pg, ei + eg, clase="grupo", hijos=[
                                fila("Ingresos", pi, ei), fila("Costos y gastos", -pg, -eg, costo=True)]))
                        nodos_l.append(fila("Total", tot[0], tot[1], clase="saldo"))
                        st.html(E.tabla_meses(nodos_l, [f"Presupuesto {SUF}", f"Ejecutado {SUF}", "Variación",
                                                        "% ejecución"], formatos=["m", "m", "m", "p"]))
                        st.caption("Resultado = ingresos − costos y gastos de la línea (incluye todas las clases del "
                                   "PyG). Variación = ejecutado − presupuesto; en costos, gastar menos da negativo. "
                                   "Los % de reparto del presupuesto se ajustan en Administración.")

                if ver_todo:
                    _cuadre(en_periodo(ejec_todo), en_periodo(ejec), PERIODO_TXT, bool(rubros_f or subs_f or lineas_f)
                            or len(gerencias) < len(todas) + 1)

    # ========== ESTADO DE RESULTADOS
    if CON_PYG:
        with tabs[1]:
            if tabs[1].open:   # solo se calcula la pestaña abierta
                with st.container(key="card_pyg"):
                    titulo_tarjeta("Estado de resultados",
                                   f"Ingresos en positivo, costos y gastos en negativo · {PERIODO_TXT}. "
                                   "Haz clic en una clase para ver sus sub-rubros.")
                    columnas = ETIQUETAS[:3] + ["Variación", "% ejecución"]
                    formatos = ["m", "m", "m", "m", "p"]
                    t_sub = resumen(["codigo", "rubro", "sub_rubro"])

                    def valores(pa, pc, ec):
                        pa, pc, ec = -pa, -pc, -ec            # signo de resultado
                        return [pa, pc, ec, ec - pc, (ec / pc if pc else None)]

                    def fila_var(nodo, total=False):
                        """Variación = ejecutado − presupuesto en valor absoluto: en un costo, gastar menos da negativo
                        (verde). Los subtotales (EBITDA, resultado) quedan en signo de resultado."""
                        v = nodo["valores"]
                        if not total and _es_costo(v[1], v[2], v[0]):
                            v[3] = -v[3]
                            nodo["dif_inv"] = True
                        for h in nodo.get("hijos", []):
                            fila_var(h)
                        return nodo

                    def nodo_clase(cod):
                        t = t_sub.loc[[cod]] if cod in t_sub.index.get_level_values(0) else None
                        if t is None or t.empty:
                            return None
                        s = t[COLS_RAW].sum()
                        hijos = []
                        t = t.reindex(t["Presupuesto anual"].abs().sort_values(ascending=False).index)
                        for (_, rub, sub), r in t.iterrows():
                            hijos.append(dict(nombre=sub, clase="linea", dif_col=3,
                                              valores=valores(*r[COLS_RAW])))
                        return fila_var(dict(nombre=cargas.RUBROS.get(cod, cod), clase="grupo", hijos=hijos,
                                             dif_col=3, sub=f"{len(hijos)} sub-rubros", valores=valores(*s)))

                    def nodo_total(nombre, clases, extra=0.0):
                        sel = t_sub[t_sub.index.get_level_values(0).isin(clases)][COLS_RAW].sum()
                        sel["Ejecutado al corte"] += extra
                        return dict(nombre=nombre, clase="saldo", dif_col=3, valores=valores(*sel))

                    # las clases que no son del PyG van aparte (junto con lo sin clasificar), antes del resultado
                    t_sub = t_sub[t_sub.index.get_level_values(0).isin(list(cargas.CLASES))]
                    codigos = sorted(set(t_sub.index.get_level_values(0)), key=cargas.orden_clase)
                    nodos = []
                    for cod in codigos:
                        if cod >= "4" and not any(n["nombre"] == "EBITDA" for n in nodos):
                            nodos.append(nodo_total("EBITDA", cargas.SUBTOTALES["EBITDA"]))
                        if cod >= "7" and not any(n["nombre"] == "Resultado antes de impuestos" for n in nodos):
                            nodos.append(nodo_total("Resultado antes de impuestos",
                                                    cargas.SUBTOTALES["Resultado antes de impuestos"]))
                        n = nodo_clase(cod)
                        if n:
                            nodos.append(n)
                    if not any(n["nombre"] == "EBITDA" for n in nodos):
                        nodos.append(nodo_total("EBITDA", cargas.SUBTOTALES["EBITDA"]))
                    if not any(n["nombre"] == "Resultado antes de impuestos" for n in nodos):
                        nodos.append(nodo_total("Resultado antes de impuestos",
                                                cargas.SUBTOTALES["Resultado antes de impuestos"]))
                    if FUERA_N:
                        nodos.append(dict(nombre="Sin clasificar", clase="linea", dif_col=3,
                                          sub=f"{FUERA_N} movimientos", valores=valores(0.0, 0.0, FUERA_EC)))
                    nodos.append(nodo_total("Resultado del ejercicio", cargas.SUBTOTALES["Resultado del ejercicio"],
                                            FUERA_EC))
                    abrir = st.toggle("Desplegar todo", key="pyg_abrir")
                    st.html(E.tabla_meses(nodos, columnas, abierto=abrir, formatos=formatos))
                    st.caption("Variación = ejecutado − presupuesto, como en las demás vistas: negativa cuando se "
                               "ejecutó menos de lo presupuestado (en un costo, gastar menos da negativo). El color dice "
                               "si es a favor (verde) o en contra (rojo). En EBITDA y resultados, la variación es la "
                               "diferencia del resultado. Las clases sin línea en el presupuesto elegido muestran "
                               "presupuesto cero.")

    # ========== PRESUPUESTO POR RUBRO
    with t_rub:
        if t_rub.open:   # solo se calcula la pestaña abierta
            if ppto.empty and ejec.empty:
                st.info("Sin presupuesto cargado.")
            else:
                with st.container(key="card_tabla"):
                    titulo_tarjeta("Presupuesto por rubro",
                                   "Haz clic en un rubro para ver sus sub-rubros, y en un sub-rubro para ver los terceros "
                                   f"que lo ejecutan. Periodo: {PERIODO_TXT}.")
                    c1, c3, c4 = st.columns([2.4, 1, 1])
                    buscar = c1.text_input("Buscar", placeholder="Buscar sub-rubro o tercero: combustible, Terpel…",
                                           label_visibility="collapsed")
                    por_ger = c3.toggle("Agrupar por gerencia", value=False) if ver_todo else False
                    abrir = c4.toggle("Desplegar todo", value=bool(buscar or subs_f))

                    niveles = (["gerencia", "rubro"] if por_ger else ["rubro"])
                    det = resumen(niveles + ["sub_rubro"]).reset_index()
                    ters = en_periodo(ejec).groupby(["rubro", "sub_rubro", "tercero"]) \
                        .agg(valor=("valor", "sum"), docs=("documento", "nunique")).reset_index()
                    if buscar:
                        en_ter = ters["tercero"].str.contains(buscar, case=False, regex=False)
                        en_sub = ters["sub_rubro"].str.contains(buscar, case=False, regex=False)
                        subs_ok = set(det.loc[det["sub_rubro"].str.contains(buscar, case=False, regex=False), "sub_rubro"]) \
                            | set(ters.loc[en_ter, "sub_rubro"])
                        det = det[det["sub_rubro"].isin(subs_ok)]
                        ters = ters[en_ter | en_sub]
                    if det.empty:
                        st.caption("No hay sub-rubros ni terceros que coincidan con la búsqueda.")
                    else:
                        st.html(E.tabla_desplegable(det, niveles, abierto=abrir, terceros=ters, etiquetas=ETIQUETAS))
                        st.caption("Los valores se muestran en positivo y la variación es ejecutado − presupuesto: negativa "
                                   "cuando se ejecutó menos de lo presupuestado. El color dice si es a favor (verde) o en "
                                   "contra (rojo): en ingresos (1.0) lo favorable es ejecutar más; en gastos, ejecutar menos. "
                                   "Los terceros muestran solo su ejecución.")

                    b1, b2, _ = st.columns([1, 1, 2])
                    # los archivos se generan solo al hacer clic (generarlos en cada recarga hacía lenta la página)
                    b1.download_button("Descargar por sub-rubro", lambda d=det: para_excel(d),
                                       f"presupuesto_vs_ejecucion_{anio}.xlsx", icon=":material/download:")
                    t_xls = ters.rename(columns={"rubro": "Rubro", "sub_rubro": "Sub-rubro", "tercero": "Tercero",
                                                 "valor": f"{ETIQUETAS[2]} (signo SAP)", "docs": "Documentos"})
                    b2.download_button("Descargar por tercero", lambda d=t_xls: a_excel(d),
                                       f"ejecucion_por_tercero_{anio}.xlsx", icon=":material/download:")

    # ========== MOVIMIENTOS
    with t_mov:
        if t_mov.open:   # solo se calcula la pestaña abierta
            with st.container(key="card_movs"):
                titulo_tarjeta("Movimientos de ejecución", "Detalle de SAP (hoja BASE), en signo SAP: ingresos en negativo. "
                               f"Periodo: {PERIODO_TXT}. Usa los filtros del informe para acotar.")
                if ejec.empty:
                    st.info("No hay ejecución cargada.")
                else:
                    buscar = st.text_input("Buscar", placeholder="Proveedor, texto, documento, cuenta…",
                                           label_visibility="collapsed", key="buscar_movs")
                    mov = en_periodo(ejec)
                    if buscar:
                        m = mov[["tercero", "texto", "documento", "cc_nombre", "cuenta_nombre"]].astype(str) \
                            .apply(lambda c: c.str.contains(buscar, case=False, regex=False)).any(axis=1)
                        mov = mov[m]
                    vista = mov[["fecha", "rubro", "sub_rubro", "tercero", "texto", "documento", "cuenta_nombre",
                                 "cc_nombre", "valor"] + (["gerencia"] if ver_todo else [])
                               + (["linea", "linea_origen"] if "linea_origen" in mov else
                                  (["linea"] if "linea" in mov else []))] \
                        .sort_values("fecha", ascending=False)
                    st.html(f'<p class="nf-card-s">{len(vista):,} movimientos · total '.replace(",", ".")
                            + f'<b style="color:{E.TEXTO}">{E.pesos(vista["valor"].sum())}</b></p>')
                    st.dataframe(
                        vista, hide_index=True, width="stretch", height=520,
                        column_config={
                            "fecha": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"),
                            "rubro": "Rubro", "sub_rubro": "Sub-rubro", "tercero": "Proveedor / tercero",
                            "texto": st.column_config.TextColumn("Texto", width="large"),
                            "documento": "Documento", "cuenta_nombre": "Cuenta", "cc_nombre": "Centro de coste",
                            "valor": st.column_config.NumberColumn("Valor", format="accounting"),
                            "gerencia": "Gerencia dueña", "linea": "Línea de negocio",
                            "linea_origen": st.column_config.TextColumn(
                                "Repartido desde", help="Parte de un centro compartido (administrativo, operativo, "
                                                        "mantenimiento o comercial) asignada a esta línea."),
                        })
                    st.download_button("Descargar movimientos", lambda d=vista: a_excel(d), f"movimientos_{anio}.xlsx",
                                       icon=":material/download:")

    # ========== COMPARAR PRESUPUESTOS
    with t_cmp:
        if t_cmp.open:   # solo se calcula la pestaña abierta
            with st.container(key="card_cmp"):
                titulo_tarjeta("Comparar presupuestos",
                               "Presupuesto anual de cada versión por clase, el ajuste del 8+4 frente al inicial y lo "
                               f"ejecutado ({PERIODO_TXT}). Ingresos en positivo, costos y gastos en negativo.")
                todos = leer("SELECT * FROM presupuesto WHERE anio=? AND empresa=?", (anio, empresa))
                todos["gerencia"] = asignacion.columna(todos, asig)
                todos = todos[todos["gerencia"].isin(gerencias)]
                if rubros_f:
                    todos = todos[todos["rubro"].isin(rubros_f)]
                if subs_f:
                    todos = todos[todos["sub_rubro"].map(cargas.norm).isin({cargas.norm(x) for x in subs_f})]
                if lineas_f:
                    todos = pd.concat([repartir_ppto(g, v) for v, g in todos.groupby("version")], ignore_index=True)
                    todos = todos[todos["linea"].isin(lineas_f)]
                vs = [v for v in versiones if v in set(todos["version"])]
                if not vs:
                    st.info("No hay presupuestos cargados.")
                else:
                    todos["clave"] = todos["sub_rubro"].map(cargas.norm)
                    ejc = en_periodo(ejec).assign(clave=lambda d: d["sub_rubro"].map(cargas.norm))
                    anual = todos.pivot_table(index=["codigo", "clave"], columns="version", values="valor",
                                              aggfunc="sum").reindex(columns=vs).fillna(0)
                    nombres = todos.drop_duplicates(["codigo", "clave"]).set_index(["codigo", "clave"])["sub_rubro"]
                    ej = ejc.groupby(["codigo", "clave"])["valor"].sum()
                    base = anual.join(ej.rename("Ejecutado"), how="outer").fillna(0)
                    ref_ini, ref_fin = vs[0], vs[-1]
                    columnas = [f"Anual {v}" for v in vs] + [f"Ajuste {ref_fin} vs {ref_ini}", f"Ejecutado {SUF}"]

                    def vals(fila):
                        a = [-fila[v] for v in vs]
                        return a + [a[-1] - a[0], -fila["Ejecutado"]]

                    def ajuste_abs(nodo):
                        """Ajuste = versión final − inicial en valor absoluto: en un costo, bajar el presupuesto da negativo
                        (verde)."""
                        v = nodo["valores"]
                        if _es_costo(v[0], v[len(vs) - 1], v[len(vs) + 1]):
                            v[len(vs)] = -v[len(vs)]
                            nodo["dif_inv"] = True
                        for h in nodo.get("hijos", []):
                            ajuste_abs(h)
                        return nodo

                    def nodo_clase(cod):
                        t = base.loc[[cod]] if cod in base.index.get_level_values(0) else None
                        if t is None:
                            return None
                        hijos = []
                        for (_, cl), r in t.reindex(t[vs + ["Ejecutado"]].abs().max(axis=1)
                                                    .sort_values(ascending=False).index).iterrows():
                            nom = nombres.get((cod, cl), cl)
                            hijos.append(dict(nombre=nom, clase="linea", valores=vals(r), dif_col=len(vs)))
                        return ajuste_abs(dict(nombre=cargas.rubro_de(cod), clase="grupo", hijos=hijos,
                                               valores=vals(t.sum()), dif_col=len(vs), sub=f"{len(hijos)} sub-rubros"))

                    def nodo_total(nombre, clases):
                        sel = base[base.index.get_level_values(0).isin(clases)].sum()
                        return dict(nombre=nombre, clase="saldo", valores=vals(sel), dif_col=len(vs))

                    codigos = sorted(set(base.index.get_level_values(0)), key=cargas.orden_clase)
                    nodos, puestos = [], set()
                    for cod in codigos:
                        for umbral, nombre in [("4", "EBITDA"), ("7", "Resultado antes de impuestos")]:
                            if cod >= umbral and nombre not in puestos:
                                nodos.append(nodo_total(nombre, cargas.SUBTOTALES[nombre]))
                                puestos.add(nombre)
                        n = nodo_clase(cod)
                        if n:
                            nodos.append(n)
                    for nombre in ["EBITDA", "Resultado antes de impuestos"]:
                        if nombre not in puestos:
                            nodos.append(nodo_total(nombre, cargas.SUBTOTALES[nombre]))
                    nodos.append(nodo_total("Resultado del ejercicio", cargas.SUBTOTALES["Resultado del ejercicio"]))
                    abrir = st.toggle("Desplegar todo", key="cmp_abrir")
                    st.html(E.tabla_meses(nodos, columnas, abierto=abrir))
                    st.caption(f"Ajuste = {ref_fin} − {ref_ini}: negativo cuando la versión nueva presupuesta menos (en "
                               "un costo, bajar el gasto da negativo). Verde mejora el resultado (más ingreso o menos "
                               "gasto), rojo lo empeora; en EBITDA y resultados es la diferencia del resultado. Los "
                               "sub-rubros se cruzan entre versiones ignorando tildes, mayúsculas y espacios.")
