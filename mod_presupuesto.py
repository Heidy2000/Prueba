"""Módulo Presupuesto: presupuesto vs ejecución de todas las clases del PyG (1.0 a 7.0) con subtotales."""
from datetime import datetime
from html import escape
from io import BytesIO

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import asignacion
import cargas
import datos
import estilo as E
import lineas
from ui import EMPRESAS, MESES, MESES_L, SIN_DUENO, a_excel, cascada, dona_doble, estilo_grafico, titulo_tarjeta

COLS_RAW = ["Presupuesto anual", "Presupuesto al corte", "Ejecutado al corte"]


MAX_TER = 30          # terceros que se muestran por sub-rubro en Presupuesto por rubro (el resto se agrupa)


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


def _con_clases_fijas(t: pd.DataFrame, por: list[str], empresa: str, gerencias, todas, rubros_f, subs_f,
                      lineas_f) -> pd.DataFrame:
    """Agrega en cero las clases que la compañía debe ver aunque aún no tengan presupuesto ni ejecución
    (cargas.CLASES_FIJAS, p. ej. 4.2 Otros proyectos en ALDC). Solo en las vistas por clase / rubro y sub-rubro, y
    solo cuando no se está filtrando por gerencia, sub-rubro o línea (ahí una fila en cero no corresponde)."""
    fijas = cargas.CLASES_FIJAS.get(empresa, [])
    if not fijas or "sub_rubro" not in por or not ({"codigo", "rubro"} & set(por)) or set(por) - {
            "codigo", "rubro", "sub_rubro", "gerencia"} or subs_f or lineas_f or len(gerencias) < len(todas) + 1:
        return t
    codigos = set(t.index.get_level_values(por.index("codigo"))) if "codigo" in por else \
        {c for c in cargas.CLASES if cargas.RUBROS[c] in set(t.index.get_level_values(por.index("rubro")))}
    nuevas = []
    for cod, sub in fijas:
        if cod in codigos or (rubros_f and cargas.RUBROS[cod] not in rubros_f):
            continue
        valores = {"codigo": cod, "rubro": cargas.RUBROS[cod], "sub_rubro": sub, "gerencia": SIN_DUENO}
        nuevas.append(tuple(valores[n] for n in por))
    if not nuevas:
        return t
    extra = pd.DataFrame(0.0, columns=t.columns, index=pd.MultiIndex.from_tuples(nuevas, names=por)
                         if len(por) > 1 else pd.Index([n[0] for n in nuevas], name=por[0]))
    return pd.concat([t, extra])


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
        # las cascadas parten del presupuesto elegido en el filtro (no del inicial)
        ppto_ini = None
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
            if ppto_ini is not None:
                ppto_ini = repartir_ppto(ppto_ini, "Inicial")
            if st.session_state.get("pp_quindio", True):      # Quindío como una sola línea (se reparte igual)
                ppto["linea"] = lineas.unificar_quindio(ppto["linea"])
                ejec_todo["linea"] = lineas.unificar_quindio(ejec_todo["linea"])
        pendientes = ejec_todo[ejec_todo["codigo"].fillna("") == ""]
        ejec = ejec_todo[ejec_todo["codigo"].fillna("") != ""].copy()
        ejec["rubro"] = ejec["codigo"].map(cargas.rubro_de)
        homol = leer("SELECT * FROM homologacion")
        ejec = datos.homologar(ejec, ppto, homol, version)
        # cada gerencia solo ve las líneas (código + sub-rubro) asignadas a ella
        ppto["gerencia"] = asignacion.columna(ppto, asig)
        ejec["gerencia"] = asignacion.columna(ejec, asig)
        ppto = ppto[ppto["gerencia"].isin(gerencias)]
        if ppto_ini is not None:
            ppto_ini["gerencia"] = asignacion.columna(ppto_ini, asig)
            ppto_ini = ppto_ini[ppto_ini["gerencia"].isin(gerencias)]
        ejec = ejec[ejec["gerencia"].isin(gerencias)]

        cols2 = st.columns([1, 1.4, 1.2] if CON_LINEAS else [1, 1.4], vertical_alignment="bottom")
        fijas = pd.DataFrame([(c, cargas.RUBROS[c], sub) for c, sub in cargas.CLASES_FIJAS.get(empresa, [])],
                             columns=["codigo", "rubro", "sub_rubro"])          # p. ej. 4.2 Otros proyectos en ALDC
        catalogo = pd.concat([ppto[["codigo", "rubro", "sub_rubro"]], ejec[["codigo", "rubro", "sub_rubro"]], fijas]) \
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
            cols2[2].checkbox("Quindío unificado", value=True, key="pp_quindio",
                              help="Circasia, La Tebaida, Montenegro y Salento se ven como una sola línea: Quindío.")
        if rubros_f:
            ppto, ejec = ppto[ppto["rubro"].isin(rubros_f)], ejec[ejec["rubro"].isin(rubros_f)]
        if subs_f:
            ppto, ejec = ppto[ppto["sub_rubro"].isin(subs_f)], ejec[ejec["sub_rubro"].isin(subs_f)]
        if lineas_f:
            ppto, ejec = ppto[ppto["linea"].isin(lineas_f)], ejec[ejec["linea"].isin(lineas_f)]
        if ppto_ini is not None:
            # el inicial puede nombrar distinto los sub-rubros: se filtra por clase y, si hay, por sub-rubro elegido
            if rubros_f:
                ppto_ini = ppto_ini[ppto_ini["codigo"].isin({r.split(" ")[0] for r in rubros_f})]
            if subs_f:
                ppto_ini = ppto_ini[ppto_ini["sub_rubro"].map(cargas.norm).isin({cargas.norm(x) for x in subs_f})]
            if lineas_f:
                ppto_ini = ppto_ini[ppto_ini["linea"].isin(lineas_f)]
            if ppto_ini.empty:
                ppto_ini = None

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

    def resumen(por: list[str], fijas: bool = False) -> pd.DataFrame:
        """Presupuesto anual, presupuesto y ejecución del periodo en signo SAP, por nivel. fijas=True agrega en cero
        las clases que la compañía debe ver aunque no tengan datos (Estado de resultados y Presupuesto por rubro)."""
        pa = ppto.groupby(por)["valor"].sum()
        pc = en_periodo(ppto).groupby(por)["valor"].sum()
        ec = en_periodo(ejec).groupby(por)["valor"].sum()
        ea = ejec[ejec["mes"] <= corte].groupby(por)["valor"].sum()
        t = pd.DataFrame({"Presupuesto anual": pa, "Presupuesto al corte": pc,
                          "Ejecutado al corte": ec, "_acum": ea}).fillna(0)
        return _con_clases_fijas(t, por, empresa, gerencias, todas, rubros_f, subs_f, lineas_f) if fijas else t

    def con_metricas(t: pd.DataFrame) -> pd.DataFrame:
        """Agrega columnas para mostrar: valores en positivo, variación (ejecutado − presupuesto), a favor/en contra,
        % y estado."""
        m = t.apply(lambda r: pd.Series(E.metricas(r["Presupuesto anual"], r["Presupuesto al corte"],
                                                   r["Ejecutado al corte"], r["_acum"])), axis=1)
        return t.join(m)

    def total_clases(clases: list[str], base: pd.DataFrame | None = None) -> dict:
        """Totales en signo de resultado (ingreso +, gasto −) para un conjunto de clases. base: otro presupuesto
        (p. ej. el inicial) en lugar del de la versión elegida."""
        pp_ = ppto if base is None else base
        p = en_periodo(pp_)
        e = en_periodo(ejec)
        return dict(pa=-pp_.loc[pp_["codigo"].isin(clases), "valor"].sum(),
                    pc=-p.loc[p["codigo"].isin(clases), "valor"].sum(),
                    ec=-e.loc[e["codigo"].isin(clases), "valor"].sum())

    def serie_mes(df: pd.DataFrame, clases: list[str] | None, signo: int) -> list:
        """Valor de cada mes hasta el corte (para la mini tendencia de las tarjetas)."""
        d = df if clases is None else df[df["codigo"].isin(clases)]
        s = signo * d.groupby("mes")["valor"].sum().reindex(range(1, corte + 1), fill_value=0.0)
        return [None if m > ult else float(v) for m, v in s.items()]

    def para_excel(df: pd.DataFrame) -> bytes:
        df = con_metricas(df.set_index([c for c in ["gerencia", "rubro", "sub_rubro"] if c in df.columns]))
        out = df[["pa", "pc", "ec", "var", "pct", "estado"]].reset_index().rename(columns={
            "gerencia": "Gerencia", "rubro": "Rubro", "sub_rubro": "Sub-rubro", "pa": ETIQUETAS[0],
            "pc": ETIQUETAS[1], "ec": ETIQUETAS[2], "var": "Variación (ejecutado − presupuesto)", "pct": "% ejecución",
            "estado": "Estado"})
        buf = BytesIO()
        out.to_excel(buf, index=False)
        return buf.getvalue()

    # ------------------------------------------------------------------ lectura, cascadas y cumplimiento
    def _signo_pct(p) -> str:
        return "—" if p is None or pd.isna(p) else ("+" if p > 0 else "") + E.pct_txt(p)

    def puente(clases: list[str], gasto: bool) -> tuple:
        """(inicial, versión, ejecutado) del periodo, en positivo (costos en positivo)."""
        sg = -1 if gasto else 1
        d = total_clases(clases)
        ini = sg * total_clases(clases, ppto_ini)["pc"] if ppto_ini is not None else None
        return ini, sg * d["pc"], sg * d["ec"]

    def pasos_cascada(clases: list[str], ver, real, gasto: bool) -> list[dict]:
        """Presupuesto de la versión del filtro → diferencia de cada sub-rubro (ingresos) o de cada clase (costos)
        → ejecutado. Las barras del medio explican la desviación."""
        sg = -1 if gasto else 1
        pc_, ec_ = en_periodo(ppto), en_periodo(ejec)
        pc_, ec_ = pc_[pc_["codigo"].isin(clases)], ec_[ec_["codigo"].isin(clases)]
        por = "sub_rubro" if (clases == ["1.0"] or len(set(pc_["codigo"]) | set(ec_["codigo"])) <= 1) else "codigo"
        dif = (-sg * ec_.groupby(por)["valor"].sum()).sub(-sg * pc_.groupby(por)["valor"].sum(), fill_value=0)
        dif = dif[dif.abs() >= 1].sort_values(key=abs, ascending=False)
        filas = [(cargas.CLASES.get(k, k) if por == "codigo" else str(k), float(v)) for k, v in dif.items()]
        if len(filas) > 5:                   # tarjeta a media página: los 4 que más pesan + otros
            filas = filas[:4] + [("Otros", sum(v for _, v in filas[4:]))]
        pasos = [dict(nombre=f"Presupuesto {version}", valor=ver, tipo="total", color="#8fb0e3")]
        for n, v in filas:
            favorable = v <= 0 if gasto else v >= 0
            pasos.append(dict(nombre=n, valor=v, color=E.VERDE if favorable else E.ROJO))
        pasos.append(dict(nombre="Ejecutado", valor=real, tipo="total", color=E.AZUL))
        return pasos

    def anticipado(ini, ver, real) -> float | None:
        """Parte de la diferencia total (real − inicial) que la versión ya había previsto."""
        if ini is None:
            return None
        dif, aj = real - ini, ver - ini
        if abs(dif) < 1 or aj * dif <= 0:
            return 0.0
        return min(aj / dif, 1.0)

    def tarjeta_cascada(clave: str, titulo: str, clases: list[str], gasto: bool):
        ini, ver, real = puente(clases, gasto)
        with st.container(key=clave):
            fig, piso = cascada(pasos_cascada(clases, ver, real, gasto), 320)
            fig.update_xaxes(tickfont=dict(size=10))
            var = real - ver
            favorable = var <= 0 if gasto else var >= 0
            pill = (f'<span class="nf-pill {"nf-ok" if favorable else "nf-bad"}" style="float:right"><i></i>'
                    f'{"+" if var >= 0 else ""}{E.millones(var)} ({_signo_pct(var / ver if ver else None)})</span>')
            titulo_tarjeta(f"{pill}{titulo}",
                           f"Del presupuesto {version} a lo ejecutado: qué explica la diferencia · {PERIODO_TXT} · "
                           "millones" + (f" · eje desde {piso:,.0f} M".replace(",", ".") if piso else ""))
            st.plotly_chart(fig, width="stretch", config={"displayModeBar": False}, key=f"g_{clave}")
        return ini, ver, real

    def bloque_cascadas():
        """Lectura del periodo, cascada del presupuesto del filtro a lo ejecutado y cumplimiento."""
        sr = con_metricas(resumen(["rubro", "sub_rubro"])).reset_index()
        contra = sr[sr["fav"] < -50_000].nsmallest(1, "fav")
        favor = sr[sr["fav"] > 50_000].nlargest(1, "fav")
        frase_dev = ""
        if not contra.empty:
            c_ = contra.iloc[0]
            frase_dev += (f" La mayor desviación en contra es <b>{escape(str(c_['sub_rubro']))}</b> "
                          f"({escape(str(c_['rubro']))}): {E.millones(c_['var'])}.")
        if not favor.empty:
            f_ = favor.iloc[0]
            frase_dev += (f" La mayor a favor, <b>{escape(str(f_['sub_rubro']))}</b>: {E.millones(f_['var'])}.")

        if CON_PYG:
            cl_cyg = [c for c in cargas.EBITDA if c != "1.0"]
            i_ini, i_ver, i_real = puente(["1.0"], False)
            g_ini, g_ver, g_real = puente(cl_cyg, True)
            e_ver, e_real = i_ver - g_ver, i_real - g_real
            p_e = e_real / e_ver if e_ver else float("nan")
            tono = "ok" if e_real >= e_ver else ("warn" if not pd.isna(p_e) and p_e >= 0.95 else "bad")
            texto = (f"Va en {E.millones(e_real)} frente a {E.millones(e_ver)} del presupuesto {version}: ingresos "
                     f"<b>{_signo_pct((i_real - i_ver) / i_ver if i_ver else None)}</b> y costos y gastos "
                     f"<b>{_signo_pct((g_real - g_ver) / g_ver if g_ver else None)}</b>." + frase_dev)
            a_i, a_g = anticipado(i_ini, i_ver, i_real), anticipado(g_ini, g_ver, g_real)
            if a_i is not None:
                texto += (f" Frente al inicial, el {version} ya anticipaba el {E.pct_txt(a_i)} de la diferencia en "
                          f"ingresos y el {E.pct_txt(a_g)} en costos y gastos.")
            st.html(E.lectura(f"EBITDA {SUF}: {E.pct_txt(p_e)} del presupuesto.", texto, tono))

            c1, c2 = st.columns(2, gap="medium")
            with c1:
                tarjeta_cascada("card_cas_ing", "Ingresos", ["1.0"], False)
            with c2:
                tarjeta_cascada("card_cas_cyg", "Costos y gastos", cl_cyg, True)

            izq, der = st.columns([2.1, 1], gap="medium")
            with izq, st.container(key="card_cas_ebi"):
                titulo_tarjeta("Qué explica la diferencia del EBITDA",
                               f"Del EBITDA presupuestado ({version}) al ejecutado, clase por clase · {PERIODO_TXT}. "
                               "Verde suma al resultado; rojo resta.")
                pc_, ec_ = en_periodo(ppto), en_periodo(ejec)
                difs = []
                for cod in cargas.SUBTOTALES["EBITDA"]:
                    d = -(float(ec_.loc[ec_["codigo"] == cod, "valor"].sum())
                          - float(pc_.loc[pc_["codigo"] == cod, "valor"].sum()))     # signo de resultado
                    if abs(d) >= 1:
                        difs.append((cargas.CLASES.get(cod, cod), d))
                difs.sort(key=lambda x: -abs(x[1]))
                if len(difs) > 7:
                    difs = difs[:6] + [("Otras clases", sum(d for _, d in difs[6:]))]
                pasos = [dict(nombre=f"EBITDA {version}", valor=e_ver, tipo="total", color="#8fb0e3")]
                pasos += [dict(nombre=n, valor=d, color=E.VERDE if d >= 0 else E.ROJO) for n, d in difs]
                pasos += [dict(nombre="EBITDA ejecutado", valor=e_real, tipo="total", color=E.AZUL)]
                fig, piso = cascada(pasos, 330)
                fig.update_xaxes(tickfont=dict(size=10))
                st.plotly_chart(fig, width="stretch", config={"displayModeBar": False}, key="g_cas_ebi")
            with der, st.container(key="card_cumpl"):
                titulo_tarjeta("Cumplimiento de ingresos", PERIODO_TXT.capitalize())
                p1 = i_real / i_ver if i_ver else None
                p2 = i_real / i_ini if i_ini else None
                st.plotly_chart(dona_doble(p1, p2, "del presupuesto", "inicial", 200), width="stretch",
                                config={"displayModeBar": False}, key="g_cumpl")
                leyenda = (f'<div style="font-size:.8rem;margin:-.3rem 0 .6rem 0;line-height:1.7">'
                           f'<i style="display:inline-block;width:9px;height:9px;border-radius:50%;background:{E.AZUL};'
                           f'margin-right:.4rem"></i><b>{E.pct_txt(p1)}</b> vs presupuesto {version}')
                if p2 is not None:
                    leyenda += (f'<br><i style="display:inline-block;width:9px;height:9px;border-radius:50%;'
                                f'background:#7fa8e6;margin-right:.4rem"></i><b>{E.pct_txt(p2)}</b> vs inicial')
                falta = i_ver - i_real
                leyenda += ("<br>" + (f'<span class="nf-pill nf-bad"><i></i>Faltan {E.millones(falta)} para la meta'
                                      '</span>' if falta > 50_000 else
                                      f'<span class="nf-pill nf-ok"><i></i>Meta superada en {E.millones(-falta)}'
                                      '</span>')) + "</div>"
                st.html(leyenda)
                if i_ini is not None:
                    st.html(E.avisos([
                        (E.millones(i_ver - i_ini), f"Ingresos · ajuste ya previsto en el {version}", "na"),
                        (E.millones(i_real - i_ver), "Ingresos · no anticipado", "ok" if i_real >= i_ver else "bad"),
                        (E.millones(g_ver - g_ini), f"Costos · ajuste ya previsto en el {version}", "na"),
                        (E.millones(g_real - g_ver), "Costos · no anticipado", "ok" if g_real <= g_ver else "bad"),
                    ]))
        else:
            clases = sorted(set(ppto["codigo"]) | set(ejec["codigo"]), key=cargas.orden_clase)
            gasto = not E.es_ingreso(float(ppto["valor"].sum()))
            ini, ver, real = puente(clases, gasto)
            p_ = real / ver if ver else float("nan")
            fav = real <= ver if gasto else real >= ver
            tono = "ok" if fav else ("warn" if abs(p_ - 1) <= 0.05 else "bad")
            st.html(E.lectura(f"Ejecución {SUF}: {E.pct_txt(p_)} del presupuesto.",
                              f"Se ejecutaron {E.millones(real)} frente a {E.millones(ver)} del presupuesto {version} "
                              f"(<b>{'+' if real >= ver else ''}{E.millones(real - ver)}</b>)." + frase_dev, tono))
            izq, der = st.columns([2.1, 1], gap="medium")
            with izq:
                tarjeta_cascada("card_cas_tot", "Del presupuesto a lo ejecutado", clases, gasto)
            with der, st.container(key="card_cumpl"):
                titulo_tarjeta("Cumplimiento", PERIODO_TXT.capitalize())
                st.plotly_chart(dona_doble(p_, (real / ini) if ini else None, "del presupuesto", "inicial", 220),
                                width="stretch", config={"displayModeBar": False}, key="g_cumpl")
                st.html(f'<div style="font-size:.8rem;line-height:1.7"><b>{E.pct_txt(p_)}</b> vs presupuesto '
                        f'{version}' + (f"<br><b>{E.pct_txt(real / ini)}</b> vs inicial" if ini else "") + "</div>")

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

    nombres_tabs = ["Resumen", "Presupuesto por rubro", "Movimientos", "Comparar presupuestos"]
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

                    def tarjeta(nombre, d, icono, gasto=False, clases=None):
                        # d en signo de resultado; para costos se muestra en positivo
                        ec, pc = (-d["ec"], -d["pc"]) if gasto else (d["ec"], d["pc"])
                        pct = ec / pc if pc else float("nan")
                        var = ec - pc                       # ejecutado − presupuesto
                        favor = -var if gasto else var      # en costos, ejecutar menos es a favor
                        color = E.VERDE if favor >= 0 else E.ROJO
                        return dict(label=f"{nombre} {SUF}", valor=E.millones(ec),
                                    delta=E.delta_valor(var, pc, a_favor_si_sube=not gasto),
                                    sub=f'Ppto {E.millones(pc)} · <b style="color:{color}">'
                                        f'{"+" if var >= 0 else ""}{E.millones(var)}</b> ({E.pct_txt(pct)})',
                                    icono=icono, color=color, fondo=E.VERDE_SUAVE if favor >= 0 else E.ROJO_SUAVE,
                                    spark=serie_mes(ejec, clases, 1 if gasto else -1) if not SOLO else None,
                                    spark_color=E.AZUL)

                    cl_cyg = [c for c in cargas.EBITDA if c != "1.0"]
                    k_ebi = tarjeta("EBITDA", ebi, "piggy", clases=cargas.SUBTOTALES["EBITDA"])
                    if ing["ec"]:
                        k_ebi["sub"] += f' · margen {E.pct_txt(ebi["ec"] / ing["ec"])}'
                    k_res = tarjeta("Resultado del ejercicio", res, "check",
                                    clases=cargas.SUBTOTALES["Resultado del ejercicio"])
                    if FUERA_N:
                        k_res["sub"] += f' · incluye {E.millones(-FUERA_EC)} sin clasificar'
                    st.html(E.kpis([tarjeta("Ingresos", ing, "wallet", clases=["1.0"]),
                                    tarjeta("Costos y gastos", cyg, "receipt", True, clases=cl_cyg), k_ebi, k_res]))
                    if FUERA_N:
                        st.warning(f"{FUERA_N} movimientos ({E.millones(FUERA_EC)} en signo SAP) aún no tienen clase: "
                                   "suman en el Resultado del ejercicio para que cuadre con SAP, pero no en ingresos, "
                                   "costos ni EBITDA. Clasifícalos en Administración → Revisar y reclasificar.",
                                   icon=":material/warning:")
                else:
                    tot = con_metricas(resumen(["anio"])).iloc[0]
                    tono = {"nf-bad": (E.ROJO, E.ROJO_SUAVE), "nf-warn": (E.AMBAR, E.AMBAR_SUAVE)}.get(
                        tot["clase"], (E.VERDE, E.VERDE_SUAVE))
                    sg_t = -1 if tot["ingreso"] else 1
                    st.html(E.kpis([
                        dict(label="Presupuesto anual", valor=E.millones(tot["pa"]),
                             sub=f"{ppto['sub_rubro'].nunique()} sub-rubros", icono="wallet", color=E.AZUL,
                             fondo=E.AZUL_SUAVE),
                        dict(label=ETIQUETAS[1], valor=E.millones(tot["pc"]), sub=PERIODO_TXT.capitalize(),
                             icono="calendar", color=E.AZUL, fondo=E.AZUL_SUAVE),
                        dict(label=ETIQUETAS[2], valor=E.millones(tot["ec"]),
                             delta=E.delta_valor(tot["var"], tot["pc"], a_favor_si_sube=bool(tot["ingreso"])),
                             sub=f'<b style="color:{tono[0]}">{E.pct_txt(tot["pct"])}</b> del presupuesto {SUF}',
                             icono="receipt", color=tono[0], fondo=tono[1],
                             spark=serie_mes(ejec, None, sg_t) if not SOLO else None, spark_color=E.AZUL),
                        dict(label="Disponible del año", valor=E.millones(tot["disp"]),
                             sub=f"Variación {SUF}: {E.millones(tot['var'])}",
                             icono="piggy" if tot["disp"] >= 0 else "alert",
                             color=E.AZUL if tot["disp"] >= 0 else E.ROJO,
                             fondo=E.AZUL_SUAVE if tot["disp"] >= 0 else E.ROJO_SUAVE),
                    ]))

                bloque_cascadas()

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
                    etiqueta = dict(texttemplate="%{y:,.0f}", textposition="outside", textangle=-90,
                                    cliponaxis=False, textfont=dict(size=10))
                    fig.add_bar(x=MESES, y=pm, name="Presupuesto", marker_color="#c9d4e4",
                                marker_line_width=0, hovertemplate="%{y:,.1f} M",
                                **etiqueta | dict(textfont=dict(size=10, color=E.TEXTO_2)))
                    fig.add_bar(x=MESES, y=em, name="Ejecutado", marker_color=E.AZUL,
                                marker_line_width=0, hovertemplate="%{y:,.1f} M",
                                **etiqueta | dict(textfont=dict(size=10, color=E.AZUL)))
                    fig.add_scatter(x=MESES, y=pm.cumsum(), name="Presupuesto acumulado", yaxis="y2",
                                    line=dict(color="#8a9bb5", dash="dot", width=2), hovertemplate="%{y:,.0f} M")
                    fig.add_scatter(x=MESES, y=em.cumsum(), name="Ejecutado acumulado", yaxis="y2",
                                    line=dict(color=E.NAVY, width=2.5), mode="lines+markers",
                                    marker=dict(size=5), hovertemplate="%{y:,.0f} M")
                    if SOLO:
                        fig.add_vrect(x0=corte - 1.5, x1=corte - 0.5, fillcolor=E.AZUL, opacity=.07, line_width=0)
                    vals_ = [float(v) for v in list(pm.fillna(0)) + list(em.fillna(0))]
                    tope = max(vals_ + [1]) if min(vals_ + [0]) >= 0 else 0      # con negativos: escala automática
                    fig.update_layout(barmode="group", bargap=.25, uniformtext=dict(minsize=8, mode="show"),
                                      yaxis=dict(tickformat=",.0f", ticksuffix=" M", rangemode="tozero",
                                                 range=[0, tope * 1.22] if tope > 0 else None),
                                      yaxis2=dict(overlaying="y", side="right", showgrid=False,
                                                  tickformat=",.0f", ticksuffix=" M", rangemode="tozero"))
                    st.plotly_chart(estilo_grafico(fig, 380), width="stretch", config={"displayModeBar": False})

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
                                                        "Variación", "% ejecución"], formatos=["m", "m", "m", "m", "p"],
                                             barras=[3]))
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
                        filas_l = []
                        for lin in orden_l:
                            pi, ei = suma(p_l, lin, True), suma(e_l, lin, True)
                            pg, eg = -suma(p_l, lin, False), -suma(e_l, lin, False)
                            if not any(abs(v) > 0.5 for v in (pi, ei, pg, eg)):
                                continue
                            filas_l.append(dict(nombre=lin, celdas=[
                                dict(v=ei - pi, pct=(ei / pi - 1) if pi else None, fav=ei >= pi),
                                dict(v=eg - pg, pct=(eg / pg - 1) if pg else None, fav=eg <= pg),
                                dict(v=(ei - eg) - (pi - pg), pct=((ei - eg) / (pi - pg) - 1) if pi - pg else None,
                                     fav=(ei - eg) >= (pi - pg))]))
                        if filas_l:
                            ti = sum(f["celdas"][0]["v"] for f in filas_l)
                            tg = sum(f["celdas"][1]["v"] for f in filas_l)
                            peor_l = min(filas_l, key=lambda f: f["celdas"][2]["v"])
                            if peor_l["celdas"][2]["v"] < 0:
                                peor_l["clase"], peor_l["sub"] = "foco", "foco"
                            filas_l.insert(0, dict(nombre="Total", clase="total", celdas=[
                                dict(v=ti, fav=ti >= 0), dict(v=tg, fav=tg <= 0), dict(v=ti - tg, fav=ti - tg >= 0)]))
                            st.html(E.barras_divergentes(filas_l, ["Δ Ingresos", "Δ Costos y gastos", "Δ Resultado"],
                                                         "Línea"))
                            st.caption("Ejecutado vs presupuesto por línea. Rojo = desfavorable (en costos, subir es "
                                       "malo); verde = favorable.")
                        st.html(E.tabla_meses(nodos_l, [f"Presupuesto {SUF}", f"Ejecutado {SUF}", "Variación",
                                                        "% ejecución"], formatos=["m", "m", "m", "p"], barras=[2]))
                        st.caption("Resultado = ingresos − costos y gastos de la línea (incluye todas las clases del "
                                   "PyG). Variación = ejecutado − presupuesto; en costos, gastar menos da negativo. "
                                   "Los % de reparto del presupuesto se ajustan en Administración.")

                if ver_todo:
                    _cuadre(en_periodo(ejec_todo), en_periodo(ejec), PERIODO_TXT, bool(rubros_f or subs_f or lineas_f)
                            or len(gerencias) < len(todas) + 1)

    # ========== ESTADO DE RESULTADOS
    # ========== PRESUPUESTO POR RUBRO (formato de estado de resultados: clase → sub-rubro → terceros)
    with t_rub:
        if t_rub.open:   # solo se calcula la pestaña abierta
            if ppto.empty and ejec.empty:
                st.info("Sin presupuesto cargado.")
            else:
                with st.container(key="card_tabla"):
                    titulo_tarjeta("Presupuesto por rubro",
                                   f"Ingresos en positivo, costos y gastos en negativo · {PERIODO_TXT}. Haz clic en una "
                                   "clase para ver sus sub-rubros, y en un sub-rubro para ver los terceros que lo ejecutan.")
                    c1, c4 = st.columns([3, 1], vertical_alignment="center")
                    buscar = c1.text_input("Buscar", placeholder="Buscar sub-rubro o tercero: combustible, Terpel…",
                                           label_visibility="collapsed", key="pp_rub_buscar")
                    abrir = c4.toggle("Desplegar todo", value=bool(buscar or subs_f), key="pp_rub_abrir")
                    columnas = ETIQUETAS[:3] + ["Variación", "% ejecución"]
                    formatos = ["m", "m", "m", "m", "p"]
                    t_sub = resumen(["codigo", "rubro", "sub_rubro"], fijas=True)
                    ters = en_periodo(ejec).groupby(["codigo", "rubro", "sub_rubro", "tercero"]) \
                        .agg(valor=("valor", "sum"), docs=("documento", "nunique")).reset_index()
                    if buscar:                    # sub-rubros que coinciden, o que tienen un tercero que coincide
                        en_ter = ters["tercero"].str.contains(buscar, case=False, regex=False)
                        en_sub = ters["sub_rubro"].str.contains(buscar, case=False, regex=False)
                        llaves = {(c, s_) for c, _, s_ in t_sub.index if buscar.lower() in str(s_).lower()} \
                            | set(zip(ters.loc[en_ter, "codigo"], ters.loc[en_ter, "sub_rubro"]))
                        t_sub = t_sub[[(c, s_) in llaves for c, _, s_ in t_sub.index]]
                        ters = ters[en_ter | en_sub]
                    por_ter = {k: g for k, g in ters.groupby(["codigo", "sub_rubro"])}

                    def valores(pa, pc, ec):
                        pa, pc, ec = -pa, -pc, -ec            # signo de resultado
                        return [pa, pc, ec, ec - pc, (ec / pc if pc else None)]

                    def fila_var(nodo, total=False):
                        """Variación = ejecutado − presupuesto en valor absoluto: en un costo, gastar menos da negativo
                        (verde). Los subtotales (EBITDA, resultado) quedan en signo de resultado. Los terceros solo
                        traen ejecución."""
                        v = nodo["valores"]
                        if not total and v[3] is not None and _es_costo(v[1], v[2], v[0]):
                            v[3] = -v[3]
                            nodo["dif_inv"] = True
                        for h in nodo.get("hijos", []):
                            fila_var(h)
                        return nodo

                    def nodos_terceros(cod, sub):
                        g = por_ter.get((cod, sub))
                        if g is None or g.empty:
                            return []
                        g = g.reindex(g["valor"].abs().sort_values(ascending=False).index)
                        out = [dict(nombre=str(r.tercero) or "Sin tercero", clase="tercero",
                                    sub=f"{int(r.docs)} doc{'s' if r.docs != 1 else ''}",
                                    valores=[None, None, -float(r.valor), None, None])
                               for r in g.head(MAX_TER).itertuples()]
                        if len(g) > MAX_TER:
                            resto = g.iloc[MAX_TER:]
                            out.append(dict(nombre=f"Otros {len(resto)} terceros", clase="tercero",
                                            valores=[None, None, -float(resto["valor"].sum()), None, None]))
                        return out

                    def nodo_clase(cod):
                        t = t_sub.loc[[cod]] if cod in t_sub.index.get_level_values(0) else None
                        if t is None or t.empty:
                            return None
                        s_ = t[COLS_RAW].sum()
                        hijos = []
                        t = t.reindex(t["Presupuesto anual"].abs().sort_values(ascending=False).index)
                        for (_, rub, sub), r in t.iterrows():
                            ter = nodos_terceros(cod, sub)
                            n_ter = len(por_ter.get((cod, sub), []))
                            hijos.append(dict(nombre=sub, clase="linea", dif_col=3, hijos=ter,
                                              sub=f"{n_ter} tercero{'s' if n_ter != 1 else ''}" if ter else "",
                                              valores=valores(*r[COLS_RAW])))
                        return fila_var(dict(nombre=cargas.RUBROS.get(cod, cod), clase="grupo", hijos=hijos,
                                             dif_col=3, sub=f"{len(hijos)} sub-rubros", valores=valores(*s_)))

                    def nodo_total(nombre, clases, extra=0.0):
                        sel = t_sub[t_sub.index.get_level_values(0).isin(clases)][COLS_RAW].sum()
                        sel["Ejecutado al corte"] += extra
                        return dict(nombre=nombre, clase="saldo", dif_col=3, valores=valores(*sel))

                    # las clases que no son del PyG van aparte (junto con lo sin clasificar), antes del resultado
                    t_sub = t_sub[t_sub.index.get_level_values(0).isin(list(cargas.CLASES))]
                    con_totales = CON_PYG and not buscar     # subtotales solo con el PyG completo y sin búsqueda
                    codigos = sorted(set(t_sub.index.get_level_values(0)), key=cargas.orden_clase)
                    nodos = []
                    for cod in codigos:
                        if con_totales and cod >= "4" and not any(n["nombre"] == "EBITDA" for n in nodos):
                            nodos.append(nodo_total("EBITDA", cargas.SUBTOTALES["EBITDA"]))
                        if con_totales and cod >= "7" and not any(n["nombre"] == "Resultado antes de impuestos"
                                                                  for n in nodos):
                            nodos.append(nodo_total("Resultado antes de impuestos",
                                                    cargas.SUBTOTALES["Resultado antes de impuestos"]))
                        n = nodo_clase(cod)
                        if n:
                            nodos.append(n)
                    if con_totales:
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
                    if not nodos:
                        st.caption("No hay sub-rubros ni terceros que coincidan con la búsqueda.")
                    else:
                        st.html(E.tabla_meses(nodos, columnas, abierto=abrir, formatos=formatos, barras=[3]))
                        st.caption("Variación = ejecutado − presupuesto, como en las demás vistas: negativa cuando se "
                                   "ejecutó menos de lo presupuestado (en un costo, gastar menos da negativo). El color "
                                   "dice si es a favor (verde) o en contra (rojo). En EBITDA y resultados, la variación "
                                   "es la diferencia del resultado. Los terceros muestran solo su ejecución (los "
                                   f"{MAX_TER} mayores de cada sub-rubro; el resto se agrupa en «Otros»).")

                    b1, b2, _ = st.columns([1, 1, 2])
                    # los archivos se generan solo al hacer clic (generarlos en cada recarga hacía lenta la página)
                    det = t_sub.reset_index().drop(columns="codigo")
                    b1.download_button("Descargar por sub-rubro", lambda d=det: para_excel(d),
                                       f"presupuesto_vs_ejecucion_{anio}.xlsx", icon=":material/download:")
                    t_xls = ters.drop(columns="codigo").rename(columns={
                        "rubro": "Rubro", "sub_rubro": "Sub-rubro", "tercero": "Tercero",
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
                    f1, f2, f3 = st.columns([1.1, 1.3, 2.2], vertical_alignment="bottom")
                    rub_m = f1.multiselect("Categoría del presupuesto", sorted(ejec["rubro"].unique(), key=_orden),
                                           placeholder="Todas", key="mov_rubro")
                    base_m = ejec[ejec["rubro"].isin(rub_m)] if rub_m else ejec
                    sub_m = f2.multiselect("Sub-rubro", sorted(base_m["sub_rubro"].dropna().unique()),
                                           placeholder="Todos", key="mov_sub")
                    if sub_m:
                        base_m = base_m[base_m["sub_rubro"].isin(sub_m)]
                    peso_t = base_m.groupby(base_m["tercero"].fillna(""))["valor"].sum().abs() \
                        .sort_values(ascending=False)
                    ter_m = f3.multiselect("Terceros", [t for t in peso_t.index if t], placeholder="Todos los terceros",
                                           key="mov_ter", help="Puedes elegir varios. Al elegir terceros aparece su "
                                                               "histórico mes a mes.")
                    buscar = st.text_input("Buscar", placeholder="Proveedor, texto, documento, cuenta…",
                                           label_visibility="collapsed", key="buscar_movs")
                    if ter_m:
                        base_m = base_m[base_m["tercero"].isin(ter_m)]
                    mov = en_periodo(base_m)
                    if ter_m:
                        _historico_terceros(leer, empresa, anio, ult, ter_m, rub_m, sub_m, asig, gerencias, SOLO, corte)
                    if buscar:
                        m = mov[["tercero", "texto", "documento", "cc_nombre", "cuenta", "cuenta_nombre"]].astype(str) \
                            .apply(lambda c: c.str.contains(buscar, case=False, regex=False)).any(axis=1)
                        mov = mov[m]
                    vista = mov[["fecha", "rubro", "sub_rubro", "tercero", "cuenta", "cuenta_nombre", "texto",
                                 "documento", "cc_nombre", "valor"] + (["gerencia"] if ver_todo else [])
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
                            "documento": "Documento", "cuenta": "Cuenta",
                            "cuenta_nombre": st.column_config.TextColumn("Nombre de cuenta", width="medium"),
                            "cc_nombre": "Centro de coste",
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


def _historico_terceros(leer, empresa, anio, ult, terceros, rubros_f, subs_f, asig, gerencias, solo, corte):
    """Gráfico mes a mes (últimos 24 meses con datos) de la ejecución de los terceros elegidos."""
    marcas = ",".join("?" * len(terceros))
    h = leer(f"SELECT anio, mes, codigo, sub_rubro, tercero, valor FROM ejecucion WHERE empresa=? AND "
             f"tercero IN ({marcas}) AND anio<=?", (empresa, *terceros, int(anio)))
    h = h[h["codigo"].fillna("") != ""]
    h = h[(h["anio"] < anio) | (h["mes"] <= ult)]
    if rubros_f:
        h = h[h["codigo"].isin({r.split(" ")[0] for r in rubros_f})]
    if subs_f:
        h = h[h["sub_rubro"].map(cargas.norm).isin({cargas.norm(x) for x in subs_f})]
    if not h.empty:
        h = h[asignacion.columna(h, asig).isin(gerencias)]
    with st.container(key="card_mov_hist"):
        if h.empty:
            st.caption("Sin histórico para los terceros elegidos.")
            return
        sg = -1 if E.es_ingreso(float(h["valor"].sum())) else 1
        h = h.assign(valor=sg * h["valor"], per=h["anio"] * 100 + h["mes"])
        periodos = sorted(h["per"].unique())[-24:]
        h = h[h["per"].isin(periodos)]
        x = [f"{MESES[p % 100 - 1]} {str(p // 100)[2:]}" for p in periodos]
        tot_t = h.groupby("tercero")["valor"].sum().abs().sort_values(ascending=False)
        top = list(tot_t.index[:7])
        h["serie"] = h["tercero"].where(h["tercero"].isin(top), "Otros terceros")
        titulo_tarjeta("Histórico mes a mes de los terceros elegidos",
                       f"{x[0]} a {x[-1]} · millones de pesos · "
                       + ("ingresos en positivo" if sg < 0 else "gasto ejecutado") +
                       (" · respeta la categoría y el sub-rubro elegidos" if rubros_f or subs_f else ""))
        colores = [E.AZUL, "#7fa8e6", E.NAVY, "#2e9e5b", "#e8892c", "#8a9bb5", "#c47012", "#c3ccd8"]
        fig = go.Figure()
        series = top + (["Otros terceros"] if (h["serie"] == "Otros terceros").any() else [])
        for i, sname in enumerate(series):
            v = h[h["serie"] == sname].groupby("per")["valor"].sum().reindex(periodos, fill_value=0.0) / 1e6
            fig.add_bar(x=x, y=v, name=sname[:40], marker_color=colores[i % len(colores)], marker_line_width=0,
                        hovertemplate=f"{sname[:40]}<br>%{{x}}: %{{y:,.1f}} M<extra></extra>")
        total = h.groupby("per")["valor"].sum().reindex(periodos, fill_value=0.0) / 1e6
        fig.add_scatter(x=x, y=total, mode="text", text=[f"<b>{v:,.0f}</b>".replace(",", ".") for v in total],
                        textposition="top center", showlegend=False, hoverinfo="skip", cliponaxis=False,
                        textfont=dict(color=E.TEXTO, size=10))
        sel = f"{MESES[corte - 1]} {str(anio)[2:]}"
        if solo and sel in x:
            i = x.index(sel)
            fig.add_vrect(x0=i - 0.5, x1=i + 0.5, fillcolor=E.AZUL, opacity=.07, line_width=0)
        fig.update_layout(barmode="relative", bargap=.3, yaxis=dict(tickformat=",.0f", ticksuffix=" M"))
        st.plotly_chart(estilo_grafico(fig, 340), width="stretch", config={"displayModeBar": False})
        anio_t = h[h["anio"] == anio]["valor"].sum()
        meses_c = h[h["anio"] == anio]["mes"].nunique()
        st.html(f'<p class="nf-card-s">{anio}: <b>{E.millones(anio_t)}</b> en {meses_c} meses con movimiento · '
                f'promedio <b>{E.millones(anio_t / meses_c if meses_c else 0)}</b> por mes. El histórico usa la '
                'clasificación de cada movimiento en la ejecución cargada.</p>')
