"""Módulo Administración (solo Financiera): carga de archivos, dueños, historial y respaldo."""
import io
from datetime import datetime

import pandas as pd
import streamlit as st

import asignacion
import cargas
import cartera
import nomina
import clasificador
import datos
import eeff
import flujo
import estilo as E
import lineas
import logica_m as L
import mod_mantenimiento
import operaciones
import mod_corte_poda
import sispo_fuente
import sispo_oper
from ui import ADMIN, EMPRESAS, MESES, MESES_L, SIN_DUENO, a_excel, ahora, empresa_de_nombre, titulo_tarjeta


COLS_EJEC = ["anio", "mes", "fecha", "codigo", "rubro", "sub_rubro", "cc", "cc_nombre", "cuenta", "cuenta_nombre",
             "documento", "tercero", "texto", "valor", "clave", "origen", "regla", "confianza", "cc_sap",
             "contrapartida", "anulado", "mod_cod", "mod_sub", "mod_cc", "raw", "empresa"]
ORIGEN_TXT = {"validado": "Validado en archivo", "manual": "Corregido en la app", "automatica": "Automática",
              "ajuste": "Ajuste del archivo", "pendiente": "Pendiente"}


def _guardar_ejecucion(con, df: pd.DataFrame):
    df = df.copy()
    df["rubro"] = df["codigo"].map(lambda c: cargas.rubro_de(c) if c else "")
    con.anexar(df[COLS_EJEC], "ejecucion")


def _empresa_adm(ctx) -> str:
    """Compañía con la que se cargan y revisan presupuesto, ejecución y flujo de caja."""
    e = st.segmented_control("Compañía", list(EMPRESAS), default=list(EMPRESAS)[0], key="adm_empresa",
                             format_func=lambda x: f"{x} · {EMPRESAS[x]}",
                             help="Los archivos que cargues y lo que revises aquí son de esta compañía.") \
        or list(EMPRESAS)[0]
    ctx.empresa = e
    return e


def _aviso_empresa(ctx, nombre_archivo: str):
    sugerida = empresa_de_nombre(nombre_archivo)
    if sugerida != ctx.empresa:
        st.warning(f"El nombre del archivo parece de **{sugerida}**, pero arriba está elegida **{ctx.empresa}**. "
                   "Revisa la compañía antes de confirmar.", icon=":material/warning:")


# ------------------------------------------------------------------ 1. presupuestos
def _cargar_presupuestos(ctx):
    con, leer, usuario = ctx.con, ctx.leer, ctx.usuario
    with st.container(key="card_adm_ppto"):
        titulo_tarjeta(f"Presupuestos de {ctx.empresa} (Inicial, 4+8, 8+4)",
                       "Excel con una hoja por versión (ej. 'PPTO INICIAL', 'PPTO 4+8', 'PPTO 8+4'): código de "
                       "rubro, sub-rubro, línea de negocio (opcional) y los 12 meses en pesos. Cada versión que "
                       "cargues reemplaza la misma versión del mismo año de esta compañía.")
        f = st.file_uploader("Presupuestos", type=["xlsx", "xlsm"], key="up_ppto", label_visibility="collapsed")
        if f:
            try:
                p = cargas.leer_presupuestos(f)
            except Exception as err:
                st.error(f"No pude leer el archivo: {err}")
                return
            if p.empty:
                st.error("No encontré hojas con el formato esperado (fila de encabezado con los 12 meses).")
                return
            _aviso_empresa(ctx, f.name)
            p["empresa"] = ctx.empresa
            # hojas sin versión en el nombre (ej. 'Hoja1' con 2026 y 2027): se elige la versión de cada año
            if (p["version"] == "").any():
                anios_sv = sorted(p.loc[p["version"] == "", "anio"].unique())
                st.markdown("**Versión de cada año** (la hoja no la dice en su nombre):")
                cols_v = st.columns(min(len(anios_sv), 4))
                elegidas = {}
                for i, a in enumerate(anios_sv):
                    defecto = "8+4" if ("8+4" in f.name and i == 0) else ("4+8" if ("4+8" in f.name and i == 0)
                                                                         else "Inicial")
                    ops = ["Inicial", "4+8", "8+4"]
                    elegidas[a] = cols_v[i % len(cols_v)].selectbox(str(a), ops, index=ops.index(defecto),
                                                                    key=f"ppto_ver_{a}")
                sv = p["version"] == ""
                p.loc[sv, "version"] = p.loc[sv, "anio"].map(elegidas)
            # ingresos en signo SAP (negativo): si el archivo los trae en positivo se voltean
            for (a, v), g in p[p["codigo"] == "1.0"].groupby(["anio", "version"]):
                if g["valor"].sum() > 0:
                    p.loc[(p["codigo"] == "1.0") & (p["anio"] == a) & (p["version"] == v), "valor"] *= -1
                    st.info(f"{v} {a}: los ingresos (1.0) venían en positivo; se guardan en negativo como en SAP.")
            # % de reparto de lo administrativo por línea de negocio (hoja con años en el encabezado)
            dist = cargas.leer_distribucion(f) if (p["linea"] != "").any() else pd.DataFrame()
            if not dist.empty:
                validas = set(p["linea"]) - set(lineas.COMPARTIDAS) - {""}
                dist = dist[dist["linea"].isin(validas) & dist["anio"].isin(p["anio"].unique())]
                if not dist.empty:
                    st.markdown("**Reparto de lo administrativo por línea** (se guarda con el presupuesto):")
                    st.dataframe(dist.pivot_table(index="linea", columns="anio", values="pct").fillna(0)
                                 .rename(columns=str), width="stretch", column_config={
                                     str(a): st.column_config.NumberColumn(str(a), format="percent")
                                     for a in dist["anio"].unique()})
            res = p.groupby("version").agg(lineas=("sub_rubro", "nunique"), clases=("codigo", "nunique"))
            st.write("Versiones encontradas: " + " · ".join(
                f"**{v}** ({r.lineas} sub-rubros, {r.clases} clases)" for v, r in res.iterrows()))
            ej = leer("SELECT codigo, mes, valor FROM ejecucion WHERE codigo<>'' AND empresa=?", (ctx.empresa,))
            corte = int(ej["mes"].max()) if not ej.empty else 12
            tabla, alertas = datos.validar_versiones(p, ej, corte)
            for a in alertas:
                st.warning(a)
            st.caption(f"Totales por clase, enero a {MESES[corte-1].lower()} (millones, signo SAP: ingresos en "
                       "negativo). En las versiones 4+8 y 8+4 los meses reales deben coincidir con la ejecución.")
            st.dataframe(tabla, hide_index=True, width="stretch",
                         column_config={c: st.column_config.NumberColumn(c, format="%.1f")
                                        for c in tabla.columns[1:]})
            if st.button("Confirmar carga de presupuestos", type="primary"):
                for (a, v) in p[["anio", "version"]].drop_duplicates().itertuples(index=False):
                    con.execute("DELETE FROM presupuesto WHERE anio=? AND version=? AND empresa=?",
                                (int(a), v, ctx.empresa))
                con.anexar(p, "presupuesto")
                if not dist.empty:
                    for a in dist["anio"].unique():
                        con.execute("DELETE FROM distribucion WHERE empresa=? AND anio=?", (ctx.empresa, int(a)))
                    con.anexar(dist.assign(empresa=ctx.empresa)[["empresa", "anio", "linea", "pct"]], "distribucion")
                con.commit()
                ctx.registrar_carga(usuario, "Presupuesto", f.name,
                                    f"{ctx.empresa} · " + ", ".join(sorted(p["version"].unique())))
                st.success("Presupuestos cargados.")
                st.rerun()


def _editar_distribucion(ctx):
    """% con que se reparte lo administrativo del presupuesto entre las líneas de negocio, por año."""
    con, leer = ctx.con, ctx.leer
    lineas_ppto = leer("SELECT DISTINCT linea FROM presupuesto WHERE empresa=? AND linea<>''", (ctx.empresa,))
    if lineas_ppto.empty:
        return
    with st.container(key="card_adm_dist"):
        titulo_tarjeta(f"Reparto de lo administrativo por línea de negocio · {ctx.empresa}",
                       "Presupuesto: la línea ADMINISTRATIVO se reparte entre las demás con estos % de cada año (se "
                       "toman de la hoja 2 del archivo al cargarlo; aquí los puedes ajustar). En los meses reales de "
                       "la 8+4 (enero a agosto) y en la ejecución se reparte según el ingreso de cada mes.")
        anios = sorted(leer("SELECT DISTINCT anio FROM presupuesto WHERE empresa=?", (ctx.empresa,))["anio"])
        receptoras = sorted(set(lineas_ppto["linea"]) - set(lineas.COMPARTIDAS))
        d = leer("SELECT anio, linea, pct FROM distribucion WHERE empresa=?", (ctx.empresa,))
        t = d.pivot_table(index="linea", columns="anio", values="pct").reindex(index=receptoras, columns=anios) \
            .fillna(0.0)
        t.columns = [str(a) for a in t.columns]
        ed = st.data_editor(t * 100, width="stretch", key="ed_dist", column_config={
            c: st.column_config.NumberColumn(f"{c} (%)", min_value=0.0, max_value=100.0, format="%.2f")
            for c in t.columns})
        sumas = ed.sum()
        malos = [c for c in ed.columns if abs(sumas[c] - 100) > 0.05 and sumas[c] > 0]
        if malos:
            st.warning("Los % de " + ", ".join(malos) + " no suman 100 (se normalizan al repartir).")
        if st.button("Guardar reparto", key="dist_ok"):
            con.execute("DELETE FROM distribucion WHERE empresa=?", (ctx.empresa,))
            filas = [(ctx.empresa, int(a), lin, float(ed.at[lin, a]) / 100) for a in ed.columns for lin in ed.index
                     if float(ed.at[lin, a]) > 0]
            con.executemany("INSERT INTO distribucion(empresa, anio, linea, pct) VALUES(?,?,?,?)", filas)
            con.commit()
            st.success("Reparto guardado.")
            st.rerun()


# ------------------------------------------------------------------ 2. ejecución SAP
def _cargar_ejecucion(ctx):
    con, leer, usuario = ctx.con, ctx.leer, ctx.usuario
    with st.container(key="card_adm_ejec"):
        titulo_tarjeta(f"Ejecución SAP de {ctx.empresa}",
                       "Sube la base de SAP (todo el año, un mes o solo unos días). Si trae tus columnas de ajuste "
                       "(Modifica COD / Base Gasto / CC) o la clasificación validada (CODIGO / BASE GASTOS / CC), se "
                       "respetan; lo que no venga clasificado se clasifica solo con lo aprendido de los meses "
                       "anteriores.")
        f = st.file_uploader("Base de ejecución", type=["xlsx", "xlsm"], key="up_ejec", label_visibility="collapsed")
        modo = st.radio("Cómo cargar", ["Agregar y actualizar", "Reemplazar los periodos del archivo"],
                        horizontal=True, key="modo_ejec",
                        help="Agregar y actualizar: suma los movimientos nuevos y actualiza los que ya estaban; "
                             "sirve para cargar por días. Reemplazar: borra lo cargado de los periodos que trae el "
                             "archivo y deja solo lo del archivo (sirve si SAP reversó o cambió movimientos).")
        if not f:
            return
        try:
            nuevos, sap_cols = cargas.leer_ejecucion_sap(f)
        except Exception as err:
            st.error(f"No pude leer el archivo: {err}")
            return
        _aviso_empresa(ctx, f.name)
        nuevos["empresa"] = ctx.empresa
        ceco = cargas.leer_ceco_lineas(f)
        if not ceco.empty:
            st.info(f"Hoja de centros de coste: {len(ceco)} centros en {ceco['linea'].nunique()} líneas de negocio "
                    f"({', '.join(sorted(ceco['linea'].unique()))}). Se guarda con la carga y define la línea de cada "
                    "movimiento.")
        todos = leer("SELECT * FROM ejecucion")
        todos["empresa"] = todos["empresa"].fillna("ALDC")
        existentes = todos[todos["empresa"] == ctx.empresa]
        if "clave" in existentes and (existentes["clave"].fillna("") == "").any():
            st.warning("La ejecución cargada antes no tiene identificación por movimiento (formato anterior). "
                       "Usa 'Reemplazar los periodos del archivo' para la primera carga con el formato nuevo.")
        fijas = leer("SELECT * FROM reglas_fijas")
        # aprende de lo clasificado en cualquier compañía (mismo plan de cuentas)
        clasif = clasificador.preparar(nuevos, existentes[existentes["clave"].fillna("") != ""], fijas,
                                       otros=todos[todos["empresa"] != ctx.empresa])
        clasif["empresa"] = ctx.empresa
        periodos = sorted(set(zip(clasif["anio"], clasif["mes"])))
        ya = clasif["clave"].isin(set(existentes["clave"]))
        cuenta = clasif["origen"].map(ORIGEN_TXT).value_counts()
        st.write(f"**{len(clasif):,} movimientos** · periodos ".replace(",", ".")
                 + ", ".join(f"{MESES[m-1]} {a}" for a, m in periodos)
                 + f" · {int(ya.sum()):,} ya estaban cargados · {int((~ya).sum()):,} nuevos".replace(",", "."))
        c = st.columns(len(cuenta))
        for col, (k, v) in zip(c, cuenta.items()):
            col.metric(k, f"{v:,}".replace(",", "."))
        # cuadre por periodo: lo que trae el archivo vs lo que ya hay en la base para esos meses
        filas_c, sobran_n, sobran_v = [], 0, 0.0
        for a, m in periodos:
            arch = clasif[(clasif["anio"] == a) & (clasif["mes"] == m)]
            base = existentes[(existentes["anio"] == a) & (existentes["mes"] == m)]
            sobra = base[~base["clave"].fillna("").isin(set(arch["clave"]))]
            sobran_n += len(sobra)
            sobran_v += float(sobra["valor"].sum())
            filas_c.append({"Periodo": f"{MESES[m-1]} {a}", "Movs. archivo": len(arch),
                            "Total archivo": arch["valor"].sum(), "Movs. hoy en la base": len(base),
                            "Total hoy en la base": base["valor"].sum(),
                            "En la base y no en el archivo": sobra["valor"].sum()})
        st.dataframe(pd.DataFrame(filas_c), hide_index=True, width="stretch", column_config={
            k: st.column_config.NumberColumn(format="accounting")
            for k in ["Total archivo", "Total hoy en la base", "En la base y no en el archivo"]})
        if sobran_n:
            n_txt = f"{sobran_n:,}".replace(",", ".")
            st.warning(f"En la base hay **{n_txt} movimientos ({E.millones(sobran_v)}, signo SAP)** de estos "
                       "periodos que **no vienen en este archivo** (por ejemplo, de una carga parcial anterior o "
                       "movimientos que SAP reversó o cambió). Con «Agregar y actualizar» se quedan y el total no "
                       "cuadrará con el archivo. Si el archivo trae el mes completo, elige **«Reemplazar los periodos "
                       "del archivo»**.", icon=":material/warning:")
        pend = clasif[clasif["origen"] == "pendiente"]
        if not pend.empty:
            st.info(f"{len(pend)} movimientos ({E.millones(pend['valor'].sum())}) no coinciden con ninguna regla. "
                    "Quedan como pendientes; los clasificas en 'Revisar clasificación'.")
        if st.button("Confirmar carga de ejecución", type="primary"):
            if modo.startswith("Reemplazar"):
                for a, m in periodos:
                    con.execute("DELETE FROM ejecucion WHERE anio=? AND mes=? AND empresa=?",
                                (int(a), int(m), ctx.empresa))
            else:
                claves = list(clasif["clave"])
                for i in range(0, len(claves), 500):
                    lote = claves[i:i + 500]
                    con.execute(f"DELETE FROM ejecucion WHERE empresa=? AND clave IN ({','.join('?' * len(lote))})",
                                [ctx.empresa] + lote)
            _guardar_ejecucion(con, clasif)
            if not ceco.empty:
                con.execute("DELETE FROM ceco_linea WHERE empresa=?", (ctx.empresa,))
                con.anexar(ceco.assign(empresa=ctx.empresa)[["empresa", "cc", "gerencia", "linea"]], "ceco_linea")
            con.execute("DELETE FROM sap_columnas")
            con.executemany("INSERT INTO sap_columnas VALUES(?,?)", list(enumerate(sap_cols)))
            con.commit()
            ctx.registrar_carga(usuario, "Ejecución", f.name,
                                f"{ctx.empresa} · {len(clasif)} movs · " + ", ".join(f"{k}: {v}" for k, v in cuenta.items()))
            st.success("Ejecución cargada.")
            st.rerun()


def _herramientas_ejecucion(ctx):
    """Precisión del clasificador y descarga de la base (se muestran siempre, sin necesidad de subir archivo)."""
    leer = ctx.leer
    with st.container(key="card_adm_bt"):
        titulo_tarjeta("Qué tan bien clasifica",
                       "Simulación: para cada mes, la app aprende de los meses anteriores y clasifica ese mes como "
                       "si fuera nuevo; luego se compara con tu clasificación validada.")
        if st.button("Calcular precisión", key="bt"):
            hist = leer("SELECT * FROM ejecucion WHERE origen IN ('validado','manual')")
            if hist.empty:
                st.caption("Aún no hay movimientos validados.")
            else:
                bt = clasificador.backtest(hist)
                bt["Mes"] = bt["mes"].map(lambda m: MESES[m - 1])
                st.dataframe(bt[["Mes", "movimientos", "cubiertos", "acierto_rubro", "acierto_valor",
                                 "acierto_con_cc"]], hide_index=True, width="stretch", column_config={
                    "movimientos": "Movimientos",
                    "cubiertos": st.column_config.NumberColumn("Con regla", format="percent"),
                    "acierto_rubro": st.column_config.NumberColumn("Acierto rubro (movs)", format="percent"),
                    "acierto_valor": st.column_config.NumberColumn("Acierto rubro (valor)", format="percent"),
                    "acierto_con_cc": st.column_config.NumberColumn("Acierto incl. CC", format="percent")})

    with st.container(key="card_adm_exp"):
        titulo_tarjeta("Descargar base de ejecución",
                       "Mismas 29 columnas de tu archivo 'Base ejecucion' (A–AC): las de SAP, tus ajustes (Modifica "
                       "COD / Modifica Base Gasto / CC) y la clasificación final (CODIGO / BASE GASTOS EXPLOTACIO / "
                       "CC), con la información que hay hoy en la app (incluye las correcciones hechas aquí).")
        periodos = leer("SELECT DISTINCT anio, mes FROM ejecucion WHERE empresa=? ORDER BY anio, mes",
                        (ctx.empresa,))
        if periodos.empty:
            st.caption("Aún no hay ejecución cargada.")
        else:
            c1, c2, c3 = st.columns([0.7, 1.8, 1.2], vertical_alignment="bottom")
            anios = sorted(periodos["anio"].unique(), reverse=True)
            a_exp = c1.selectbox("Año", anios, key="exp_anio")
            meses_a = periodos.loc[periodos["anio"] == a_exp, "mes"].astype(int).tolist()
            m_exp = c2.multiselect("Meses", meses_a, format_func=lambda m: MESES[m - 1], placeholder="Todos",
                                   key="exp_meses")
            control = c3.checkbox("Agregar hoja con origen de cada clasificación", key="exp_ctrl",
                                  help="Una segunda hoja 'Control' con el origen (validado, automática, corregida "
                                       "en la app, pendiente), la regla y la confianza. La hoja principal no cambia.")
            sufijo = "".join(f"_{MESES[m - 1]}" for m in sorted(m_exp)) if m_exp else ""
            st.download_button("Descargar base de ejecución", lambda a=a_exp, m=tuple(m_exp), c=control, e=ctx.empresa: _exportar(leer, a, list(m), c, e),
                               f"Base ejecucion {ctx.empresa} {a_exp}{sufijo}.xlsx", icon=":material/download:", type="primary",
                               key="exp_dl")


# las 29 columnas del archivo 'Base ejecucion' (A–W SAP, X–Z ajustes, AA–AC clasificación final)
COLS_BASE_SAP = ["Referencia", "Denom.clase de coste", "Fecha de documento", "Nº docum.refer.", "Valor/mon.inf.",
                 "Texto de cabecera de documento", "Cta.contrapartida", "Denominacion cuenta contrapartida",
                 "Centro de coste", "Anulado", "De período", "Denom.cuenta contrapartida", "TEXTO",
                 "Denominación del objeto", "Documento anulación", "Ejercicio", "Fe.contabilización",
                 "Fecha entrada", "Nº ref.anulación", "Objeto", "Sociedad GL asociada", "Sociedad referencia",
                 "Usuario"]
FECHAS_BASE = {"Fecha de documento", "Fe.contabilización", "Fecha entrada"}


def _exportar(leer, anio=None, meses=None, control=False, empresa="ALDC") -> bytes:
    """Ejecución en la estructura del archivo 'Base ejecucion' (hoja Hoja1, columnas A–AC)."""
    import json
    from io import BytesIO
    sql, params = "SELECT * FROM ejecucion WHERE empresa=?", [empresa]
    if anio is not None:
        sql += " AND anio=?"
        params.append(int(anio))
        if meses:
            sql += f" AND mes IN ({','.join('?' * len(meses))})"
            params += [int(m) for m in meses]
    e = leer(sql + " ORDER BY anio, mes, fecha", tuple(params)).reset_index(drop=True)
    n = len(COLS_BASE_SAP)

    def fila_sap(r):
        vals = json.loads(r.raw) if r.raw else []
        if len(vals) >= n:
            return vals[:n]
        # sin datos originales: se arma con lo que guarda la app
        return [r.cuenta, r.cuenta_nombre, r.fecha, r.documento, r.valor, "", r.contrapartida, r.tercero, r.cc_sap,
                r.anulado, str(r.mes), r.tercero, r.texto, r.cc_nombre, "", str(r.anio), r.fecha, "", "", r.cc_sap,
                "", "", ""]
    sap = pd.DataFrame([fila_sap(r) for r in e.itertuples()], columns=COLS_BASE_SAP)
    for c in FECHAS_BASE:
        sap[c] = pd.to_datetime(sap[c].replace("", None), errors="coerce")
    sap["Valor/mon.inf."] = pd.to_numeric(sap["Valor/mon.inf."], errors="coerce")
    vacio = lambda s: s.fillna("").replace("", None)
    ajustes = pd.DataFrame({"Modifica COD": vacio(e["mod_cod"]), "Modifica Base Gasto": vacio(e["mod_sub"]),
                            "CC": vacio(e["mod_cc"])})
    final = pd.DataFrame({"CODIGO": vacio(e["codigo"]), "BASE GASTOS EXPLOTACIO": vacio(e["sub_rubro"]),
                          "CC": vacio(e["cc"])})
    # escritura directa con xlsxwriter (bastante más rápida que pandas para ~15.000 filas × 29 columnas)
    import xlsxwriter
    buf = BytesIO()
    libro = xlsxwriter.Workbook(buf, {"constant_memory": True, "in_memory": True})
    f_fecha = libro.add_format({"num_format": "dd/mm/yyyy"})
    f_valor = libro.add_format({"num_format": "#,##0.00;-#,##0.00"})
    f_tit = libro.add_format({"bold": True})
    encabezados = COLS_BASE_SAP + list(ajustes.columns) + list(final.columns)
    hoja = libro.add_worksheet("Hoja1")
    hoja.freeze_panes(1, 0)
    hoja.set_column(0, len(encabezados) - 1, 16)
    hoja.write_row(0, 0, encabezados, f_tit)
    i_fechas = [COLS_BASE_SAP.index(c) for c in FECHAS_BASE]
    i_valor = COLS_BASE_SAP.index("Valor/mon.inf.")
    datos = pd.concat([sap, ajustes.set_axis(["_m1", "_m2", "_m3"], axis=1),
                       final.set_axis(["_f1", "_f2", "_f3"], axis=1)], axis=1)
    for k, fila in enumerate(datos.itertuples(index=False, name=None), start=1):
        for c, v in enumerate(fila):
            if v is None or (isinstance(v, float) and v != v) or v is pd.NaT:
                continue
            if c in i_fechas:
                hoja.write_datetime(k, c, v.to_pydatetime(), f_fecha)
            elif c == i_valor:
                hoja.write_number(k, c, float(v), f_valor)
            else:
                hoja.write_string(k, c, str(v))
    hoja.autofilter(0, 0, len(datos), len(encabezados) - 1)
    if control:
        ctrl = pd.DataFrame({"Referencia": sap["Referencia"], "Nº docum.refer.": sap["Nº docum.refer."],
                             "Valor/mon.inf.": sap["Valor/mon.inf."], "CODIGO": e["codigo"],
                             "BASE GASTOS EXPLOTACIO": e["sub_rubro"],
                             "Origen clasificación": e["origen"].map(ORIGEN_TXT).fillna(e["origen"]),
                             "Regla": e["regla"], "Confianza": e["confianza"]})
        h2 = libro.add_worksheet("Control")
        h2.write_row(0, 0, list(ctrl.columns), f_tit)
        h2.set_column(0, 7, 18)
        for k, fila in enumerate(ctrl.itertuples(index=False, name=None), start=1):
            for c, v in enumerate(fila):
                if v is None or (isinstance(v, float) and v != v):
                    continue
                if c == 2:
                    h2.write_number(k, c, float(v), f_valor)
                else:
                    h2.write_string(k, c, str(v))
    libro.close()
    return buf.getvalue()


def _opciones_clasificacion(leer) -> list[str]:
    lineas = leer("SELECT DISTINCT codigo, sub_rubro FROM presupuesto UNION "
                  "SELECT DISTINCT codigo, sub_rubro FROM ejecucion WHERE codigo<>''")
    return sorted({f"{c} · {s}" for c, s in zip(lineas["codigo"], lineas["sub_rubro"]) if c},
                  key=lambda x: (cargas.orden_clase(x.split(" · ")[0]), x))


def _revisar_clasificacion(ctx):
    con, leer = ctx.con, ctx.leer
    ORIG = {"Pendientes": ["pendiente"], "Automáticas": ["automatica", "ajuste"],
            "Corregidas en la app": ["manual"], "Validadas en archivo": ["validado"]}
    with st.container(key="card_adm_rev"):
        titulo_tarjeta("Revisar y reclasificar",
                       "Filtra por tercero, cuenta contable, mes o clasificación y corrige en bloque o fila por fila. Lo que corrijas "
                       "queda como 'Corregido en la app' y la app aprende de ello para los próximos meses.")
        cols_e = ", ".join(sorted(con.columnas("ejecucion") - {"raw", "id"}))
        e = leer(f"SELECT rowid AS id, {cols_e} FROM ejecucion WHERE empresa=?", (ctx.empresa,))
        if e.empty:
            st.caption(f"Aún no hay ejecución cargada de {ctx.empresa}.")
            return
        e["Clasificación"] = [f"{c} · {s}" if c else "(sin clasificar)" for c, s in zip(e["codigo"], e["sub_rubro"])]

        f1, f2, f3 = st.columns([1.6, 1, 1.2])
        origenes = f1.multiselect("Origen", list(ORIG), default=["Pendientes", "Automáticas"], key="rev_origen")
        meses = f2.multiselect("Mes", sorted(e["mes"].unique()), format_func=lambda m: MESES[m - 1],
                               placeholder="Todos", key="rev_mes")
        buscar = f3.text_input("Buscar", placeholder="Texto, cuenta, documento…", key="rev_buscar")
        sel = e[e["origen"].isin(sum((ORIG[o] for o in origenes), []))] if origenes else e
        if meses:
            sel = sel[sel["mes"].isin(meses)]
        g1, g2, g3 = st.columns(3)
        por_ter = sel.groupby("tercero")["valor"].agg(lambda v: v.abs().sum()).sort_values(ascending=False)
        terceros = g1.multiselect("Tercero", list(por_ter.index), placeholder="Todos los terceros", key="rev_ter",
                                  format_func=lambda t: f"{t}  ·  {E.millones(por_ter[t])}")
        if terceros:
            sel = sel[sel["tercero"].isin(terceros)]
        por_cta = sel.groupby("cuenta")["valor"].agg(lambda v: v.abs().sum()).sort_values(ascending=False)
        nombre_cta = dict(zip(sel["cuenta"], sel["cuenta_nombre"]))
        cuentas = g2.multiselect("Cuenta contable", list(por_cta.index), placeholder="Todas las cuentas",
                                 key="rev_cta",
                                 format_func=lambda c: f"{c} · {nombre_cta.get(c, '')}  ·  {E.millones(por_cta[c])}")
        if cuentas:
            sel = sel[sel["cuenta"].isin(cuentas)]
        clases = g3.multiselect("Clasificación actual", sorted(sel["Clasificación"].unique()),
                                placeholder="Todas", key="rev_clas")
        if clases:
            sel = sel[sel["Clasificación"].isin(clases)]
        if buscar:
            sel = sel[sel[["texto", "cuenta_nombre", "cuenta", "documento"]].astype(str)
                      .apply(lambda s: s.str.contains(buscar, case=False, regex=False)).any(axis=1)]
        if sel.empty:
            st.success("No hay movimientos con estos filtros.")
            return

        res = sel.groupby("Clasificación")["valor"].agg(["size", "sum"]).sort_values("size", ascending=False)
        st.html(f'<p class="nf-card-s"><b style="color:{E.TEXTO}">{len(sel):,} movimientos</b> · '.replace(",", ".")
                + f'{E.pesos(sel["valor"].sum())} · clasificados hoy en: '
                + " · ".join(f"{k} ({int(r['size'])})" for k, r in res.head(6).iterrows())
                + (" …" if len(res) > 6 else "") + "</p>")

        opciones = _opciones_clasificacion(leer)
        # ---- reclasificación en bloque
        with st.container(key="card_adm_bloque"):
            st.html(f'<p class="nf-card-t" style="font-size:.95rem">Reclasificar en bloque los {len(sel)} '
                    'movimientos filtrados</p>')
            b1, b2, b3 = st.columns([2.2, 0.9, 1.3], vertical_alignment="bottom")
            nueva = b1.selectbox("Nueva clasificación", opciones, index=None, placeholder="Elige rubro · sub-rubro",
                                 key="rev_nueva")
            nuevo_cc = b2.text_input("CC (opcional)", placeholder="Igual", key="rev_cc",
                                     help="Déjalo vacío para conservar el centro de coste de cada movimiento.")
            if terceros:
                txt_fija = "Crear regla fija para los terceros filtrados"
            elif cuentas:
                txt_fija = "Crear regla fija para las cuentas filtradas (todos los terceros)"
            else:
                txt_fija = "Crear regla fija (filtra por tercero o por cuenta)"
            fija = b3.checkbox(txt_fija, key="rev_fija", disabled=not (terceros or cuentas),
                               help="En las próximas cargas de SAP, los movimientos de estos terceros o de estas "
                                    "cuentas irán siempre a esta clasificación.")
            por_cuenta = st.checkbox("La regla del tercero aplica solo a las cuentas contables de estos movimientos",
                                     value=True, key="rev_porcuenta", disabled=not (terceros and fija))
            if st.button(f"Aplicar a {len(sel)} movimientos", type="primary", disabled=nueva is None,
                         key="rev_aplicar"):
                cod, sub = nueva.split(" · ", 1)
                ids = [int(i) for i in sel["id"]]
                for i in range(0, len(ids), 500):
                    lote = ids[i:i + 500]
                    if nuevo_cc.strip():
                        con.execute(f"UPDATE ejecucion SET codigo=?, rubro=?, sub_rubro=?, cc=?, origen='manual', "
                                    f"regla='corrección en bloque', confianza='' WHERE rowid IN "
                                    f"({','.join('?' * len(lote))})",
                                    [cod, cargas.rubro_de(cod), sub, nuevo_cc.strip()] + lote)
                    else:
                        con.execute(f"UPDATE ejecucion SET codigo=?, rubro=?, sub_rubro=?, origen='manual', "
                                    f"regla='corrección en bloque', confianza='' WHERE rowid IN "
                                    f"({','.join('?' * len(lote))})", [cod, cargas.rubro_de(cod), sub] + lote)
                n_reglas = 0
                if fija and (terceros or cuentas):
                    if terceros:
                        pares = sel[["tercero", "cuenta"]].drop_duplicates() if por_cuenta else \
                            pd.DataFrame({"tercero": terceros, "cuenta": ""})
                    else:   # regla por cuenta, para cualquier tercero
                        pares = pd.DataFrame({"tercero": "", "cuenta": cuentas})
                    for t, cta in pares.itertuples(index=False):
                        con.execute("DELETE FROM reglas_fijas WHERE tercero=? AND cuenta=?", (t, cta))
                        con.execute("INSERT INTO reglas_fijas VALUES(?,?,?,?,?,?)",
                                    (t, cta, cod, sub, nuevo_cc.strip(), ahora().strftime("%Y-%m-%d %H:%M")))
                        n_reglas += 1
                con.commit()
                st.success(f"{len(ids)} movimientos reclasificados a {nueva}"
                           + (f" · {n_reglas} reglas fijas creadas" if n_reglas else "") + ".")
                st.rerun()

        # ---- edición fila por fila
        with st.expander("Corregir fila por fila", icon=":material/edit:"):
            vista = sel.sort_values(["mes", "valor"], ascending=[True, False]).head(500)
            vista = vista[["id", "fecha", "cuenta", "cuenta_nombre", "tercero", "texto", "cc_sap", "valor", "Clasificación",
                           "cc", "origen", "regla"]].reset_index(drop=True)
            vista["origen"] = vista["origen"].map(ORIGEN_TXT)
            vista.loc[vista["Clasificación"] == "(sin clasificar)", "Clasificación"] = None
            editado = st.data_editor(
                vista, hide_index=True, width="stretch", height=440, key="rev_editor",
                disabled=["id", "fecha", "cuenta", "cuenta_nombre", "tercero", "texto", "cc_sap", "valor", "origen", "regla"],
                column_config={
                    "id": None, "fecha": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"),
                    "cuenta": "Cuenta", "cuenta_nombre": "Nombre cuenta", "tercero": "Tercero",
                    "texto": st.column_config.TextColumn("Texto", width="medium"), "cc_sap": "CC SAP",
                    "valor": st.column_config.NumberColumn("Valor", format="accounting"),
                    "Clasificación": st.column_config.SelectboxColumn("Clasificación", options=opciones,
                                                                      width="large"),
                    "cc": "CC final", "origen": "Origen", "regla": "Regla usada"})
            if len(sel) > 500:
                st.caption("Se muestran los primeros 500; usa los filtros para acotar.")
            if st.button("Guardar correcciones", key="rev_guardar"):
                def _txt(v):   # celdas vacías llegan como None o NaN (número); se tratan como texto vacío
                    return v.strip() if isinstance(v, str) else ""

                cambios = 0
                for orig, nuevo in zip(vista.itertuples(), editado.itertuples()):
                    cla_n, cla_o = _txt(nuevo.Clasificación), _txt(orig.Clasificación)
                    cc_n, cc_o = _txt(nuevo.cc), _txt(orig.cc)
                    if cla_n and " · " in cla_n and (cla_n != cla_o or cc_n != cc_o):
                        cod, sub = cla_n.split(" · ", 1)
                        con.execute("UPDATE ejecucion SET codigo=?, rubro=?, sub_rubro=?, cc=?, origen='manual', "
                                    "regla='corrección manual', confianza='' WHERE rowid=?",
                                    (cod, cargas.rubro_de(cod), sub, cc_n, int(orig.id)))
                        cambios += 1
                con.commit()
                st.success(f"{cambios} movimientos guardados como corrección manual.")
                st.rerun()

    with st.container(key="card_adm_fijas"):
        titulo_tarjeta("Reglas fijas por tercero o por cuenta",
                       "Se aplican antes que lo aprendido en cada carga nueva de SAP. Marca 'Borrar' y guarda para "
                       "eliminar una regla.")
        fijas = leer("SELECT rowid AS id, * FROM reglas_fijas ORDER BY tercero")
        if fijas.empty:
            st.caption("Aún no hay reglas fijas. Se crean desde 'Reclasificar en bloque' filtrando por tercero o "
                       "por cuenta contable.")
        else:
            fijas["Clasificación"] = fijas["codigo"] + " · " + fijas["sub_rubro"]
            fijas["Borrar"] = False
            fijas["tercero"] = fijas["tercero"].replace("", "(todos los terceros)")
            nombres = dict(leer("SELECT DISTINCT cuenta, cuenta_nombre FROM ejecucion").values.tolist())
            fijas["Nombre cuenta"] = fijas["cuenta"].map(nombres).fillna("")
            ed = st.data_editor(fijas[["id", "tercero", "cuenta", "Nombre cuenta", "Clasificación", "cc", "creada",
                                       "Borrar"]],
                                hide_index=True, width="stretch", key="fijas_editor",
                                disabled=["id", "tercero", "cuenta", "Nombre cuenta", "Clasificación", "cc", "creada"],
                                column_config={"id": None, "tercero": "Tercero",
                                               "cuenta": st.column_config.TextColumn("Cuenta", help="Vacía = todas las cuentas"),
                                               "cc": "CC", "creada": "Creada"})
            if st.button("Guardar cambios en reglas", key="fijas_guardar"):
                borrar = [int(i) for i, b in zip(ed["id"], ed["Borrar"]) if b]
                for i in borrar:
                    con.execute("DELETE FROM reglas_fijas WHERE rowid=?", (i,))
                con.commit()
                st.success(f"{len(borrar)} reglas eliminadas.")
                st.rerun()


# ------------------------------------------------------------------ 4. homologar nombres
def _homologar(ctx):
    con, leer = ctx.con, ctx.leer
    with st.container(key="card_adm_hom"):
        titulo_tarjeta("Homologar nombres",
                       "Sub-rubros de la ejecución que no encuentran su línea en una versión del presupuesto (ya se "
                       "ignoran tildes, mayúsculas y espacios). Elige a qué línea corresponden; si no tienen "
                       "presupuesto, déjalos vacíos.")
        versiones = leer("SELECT DISTINCT version FROM presupuesto WHERE empresa=?", (ctx.empresa,))["version"].tolist()
        if not versiones:
            st.caption("Carga primero los presupuestos.")
            return
        version = st.segmented_control("Versión", versiones, default=versiones[0], key="hom_v") or versiones[0]
        ppto = leer("SELECT DISTINCT codigo, sub_rubro FROM presupuesto WHERE version=? AND empresa=?",
                    (version, ctx.empresa))
        ej = leer("SELECT codigo, sub_rubro, valor FROM ejecucion WHERE codigo<>'' AND empresa=?", (ctx.empresa,))
        homol = leer("SELECT * FROM homologacion")
        tabla = datos.sin_cruce(ej, ppto, homol, version)
        if tabla.empty:
            st.success(f"Todos los sub-rubros de la ejecución cruzan con el presupuesto {version}.")
            return
        opciones = [""] + sorted(f"{c} · {s}" for c, s in zip(ppto["codigo"], ppto["sub_rubro"]))
        tabla["Homologar a"] = [f"{c} · {s}" if s else "" for c, s in zip(tabla["Código"], tabla["Sugerencia"])]
        editado = st.data_editor(
            tabla.drop(columns=["Opciones", "Sugerencia"]), hide_index=True, width="stretch", key="hom_editor",
            disabled=["Código", "Sub-rubro en la ejecución", "Ejecutado (M)"],
            column_config={"Ejecutado (M)": st.column_config.NumberColumn(format="%.1f"),
                           "Homologar a": st.column_config.SelectboxColumn(options=opciones, width="large")})
        st.caption("Las sugerencias son por parecido de nombre; revísalas antes de guardar.")
        if st.button("Guardar homologaciones", type="primary"):
            n = 0
            for r in editado.itertuples():
                destino, sub_e = r._4, ("" if r._2 == "(vacío)" else r._2)
                con.execute("DELETE FROM homologacion WHERE version=? AND codigo=? AND sub_ejec=?",
                            (version, r.Código, sub_e))
                if destino:
                    cod, sub = destino.split(" · ", 1)
                    if cod == r.Código:
                        con.execute("INSERT INTO homologacion VALUES(?,?,?,?)", (version, r.Código, sub_e, sub))
                        n += 1
            con.commit()
            st.success(f"{n} homologaciones guardadas para el presupuesto {version}.")
            st.rerun()


# ------------------------------------------------------------------ asignación de sub-rubros a gerencias
NINGUNA = "(sin asignar)"


def _catalogo(leer, anio) -> pd.DataFrame:
    """Todas las líneas del año (código + sub-rubro) de todas las versiones y de la ejecución, con su asignación."""
    ppto = leer("SELECT version, codigo, rubro, sub_rubro, valor FROM presupuesto WHERE anio=?", (anio,))
    ejec = leer("SELECT codigo, sub_rubro, valor FROM ejecucion WHERE anio=? AND codigo<>''", (anio,))
    versiones = list(ppto["version"].unique())
    ref = "Inicial" if "Inicial" in versiones else (versiones[0] if versiones else "")
    p_ref = ppto[ppto["version"] == ref]
    if not ejec.empty:
        ejec = datos.homologar(ejec, p_ref, leer("SELECT * FROM homologacion"), ref)
    lineas = pd.concat([ppto[["codigo", "sub_rubro"]], ejec[["codigo", "sub_rubro"]]])
    lineas["sub_norm"] = lineas["sub_rubro"].map(cargas.norm)
    lineas = lineas.drop_duplicates(["codigo", "sub_norm"]).copy()
    pa = p_ref.assign(sub_norm=p_ref["sub_rubro"].map(cargas.norm)).groupby(["codigo", "sub_norm"])["valor"].sum()
    ea = ejec.assign(sub_norm=ejec["sub_rubro"].map(cargas.norm)).groupby(["codigo", "sub_norm"])["valor"].sum() \
        if not ejec.empty else pd.Series(dtype=float)
    llave = list(zip(lineas["codigo"], lineas["sub_norm"]))
    lineas["ppto"] = [abs(pa.get(k, 0)) / 1e6 for k in llave]
    lineas["ejec"] = [abs(ea.get(k, 0)) / 1e6 for k in llave]
    lineas["rubro"] = lineas["codigo"].map(cargas.rubro_de)
    g = asignacion.columna(lineas, asignacion.mapa(leer))
    lineas["gerencia"] = g.where(g != SIN_DUENO, NINGUNA)
    lineas["_orden"] = lineas["codigo"].map(cargas.orden_clase)
    return lineas.sort_values(["_orden", "sub_rubro"]).reset_index(drop=True), ref


def _asignar(ctx, anio):
    anio = int(anio)
    con, leer = ctx.con, ctx.leer
    gerencias = [u for u in ctx.usuarios if u != ADMIN]
    cat, ref = _catalogo(leer, anio)
    opciones = gerencias + sorted(set(cat["gerencia"]) - set(gerencias) - {NINGUNA}) + [NINGUNA]
    completo = list(getattr(ctx, "acceso", {}).get("Presupuesto completo", []))

    with st.container(key="card_asig"):
        titulo_tarjeta("Asignación de sub-rubros por gerencia",
                       "Cada gerencia, al ingresar, ve solo el presupuesto y la ejecución de los sub-rubros que le "
                       "asignes aquí. Financiera" + (f" y {', '.join(completo)}" if completo else "")
                       + " ven toda la compañía. La asignación vale para las tres versiones del presupuesto.")
        if cat.empty:
            st.info("Primero carga los presupuestos.")
            return
        n_sin = int((cat["gerencia"] == NINGUNA).sum())
        st.html(f'<p class="nf-card-s">{len(cat)} líneas del {anio} · <b>{len(cat) - n_sin}</b> asignadas · '
                + (f'<b style="color:{E.ROJO}">{n_sin} sin asignar</b>' if n_sin else "todas asignadas")
                + f' · cifras en millones, presupuesto {ref}</p>')

        # ---- asignación masiva por rubro
        with st.expander("Asignar un rubro completo a una gerencia", icon=":material/playlist_add:"):
            rubros = list(dict.fromkeys(cat["rubro"]))
            c1, c2, c3 = st.columns([1.6, 1.1, 0.9], vertical_alignment="bottom")
            r_sel = c1.multiselect("Rubros", rubros, key="asig_rub")
            g_sel = c2.selectbox("Gerencia", opciones, key="asig_ger")
            solo_libres = c3.checkbox("Solo los sin asignar", value=True, key="asig_libres")
            if r_sel:
                obj = cat[cat["rubro"].isin(r_sel)]
                if solo_libres:
                    obj = obj[obj["gerencia"] == NINGUNA]
                st.caption(f"Se asignarán {len(obj)} sub-rubros a {g_sel}.")
                if st.button("Asignar", type="primary", key="asig_btn", disabled=obj.empty):
                    asignacion.guardar(con, obj.assign(gerencia="" if g_sel == NINGUNA else g_sel))
                    ctx.registrar_carga(ctx.usuario, "Asignación", "", f"{len(obj)} sub-rubros → {g_sel}")
                    st.session_state["asig_v"] = st.session_state.get("asig_v", 0) + 1
                    st.rerun()

        # ---- editor línea por línea
        f1, f2, f3 = st.columns([1.3, 1, 1.3])
        fr = f1.multiselect("Rubro", list(dict.fromkeys(cat["rubro"])), placeholder="Todos los rubros",
                            key="asig_fr")
        fg = f2.selectbox("Gerencia", ["Todas"] + opciones, key="asig_fg")
        fb = f3.text_input("Buscar sub-rubro", key="asig_fb", placeholder="Combustible, nómina…")
        vista = cat
        if fr:
            vista = vista[vista["rubro"].isin(fr)]
        if fg != "Todas":
            vista = vista[vista["gerencia"] == fg]
        if fb:
            vista = vista[vista["sub_rubro"].str.contains(fb, case=False, regex=False)]

        editado = st.data_editor(
            vista[["codigo", "rubro", "sub_rubro", "ppto", "ejec", "gerencia"]], hide_index=True,
            width="stretch", height=480, key=f"asig_ed_{st.session_state.get('asig_v', 0)}",
            disabled=["codigo", "rubro", "sub_rubro", "ppto", "ejec"],
            column_config={
                "codigo": st.column_config.TextColumn("Código", width="small"),
                "rubro": "Rubro", "sub_rubro": st.column_config.TextColumn("Sub-rubro", width="large"),
                "ppto": st.column_config.NumberColumn(f"Ppto {ref} (M)", format="%.1f"),
                "ejec": st.column_config.NumberColumn("Ejecutado (M)", format="%.1f"),
                "gerencia": st.column_config.SelectboxColumn("Gerencia dueña", options=opciones, required=True),
            })
        cambios = editado[editado["gerencia"] != vista["gerencia"]]
        b1, b2, _ = st.columns([1.1, 1.1, 2])
        if b1.button(f"Guardar cambios ({len(cambios)})", type="primary", disabled=cambios.empty, key="asig_save"):
            asignacion.guardar(con, cambios.assign(gerencia=cambios["gerencia"].replace(NINGUNA, "")))
            ctx.registrar_carga(ctx.usuario, "Asignación", "", f"{len(cambios)} sub-rubros editados")
            st.session_state["asig_v"] = st.session_state.get("asig_v", 0) + 1
            st.rerun()
        b2.download_button("Descargar asignación", lambda d=cat[["codigo", "rubro", "sub_rubro", "gerencia"]].rename(
            columns={"codigo": "Código", "rubro": "Rubro", "sub_rubro": "Sub-rubro", "gerencia": "Gerencia"}): a_excel(d),
            f"asignacion_gerencias_{anio}.xlsx", icon=":material/download:", key="asig_dl")

    # ---- resumen por gerencia
    with st.container(key="card_asig_res"):
        titulo_tarjeta("Qué ve cada gerencia", f"Líneas asignadas y su presupuesto {ref} anual (millones)")
        res = cat.groupby("gerencia").agg(Líneas=("sub_rubro", "size"), Presupuesto=("ppto", "sum"),
                                          Ejecutado=("ejec", "sum"))
        res = res.reindex([g for g in opciones if g in res.index])
        sin_nada = [g for g in gerencias if g not in res.index]
        st.dataframe(res.reset_index().rename(columns={"gerencia": "Gerencia"}), hide_index=True, width="stretch",
                     column_config={"Presupuesto": st.column_config.NumberColumn(format="%.1f"),
                                    "Ejecutado": st.column_config.NumberColumn(format="%.1f")})
        if sin_nada:
            st.caption("Sin líneas asignadas (al ingresar no verán presupuesto): " + ", ".join(sin_nada) + ".")

    # ---- carga por archivo (opcional)
    with st.expander("Cargar la asignación desde un Excel", icon=":material/upload_file:"):
        st.caption("Excel con una columna de sub-rubro y otra con la gerencia. Si tiene columna de código, se usa; "
                   "si no, la asignación vale para ese sub-rubro en cualquier clase. Solo cambia las líneas del "
                   "archivo; lo demás queda como está.")
        fd = st.file_uploader("Archivo de asignación", type=["xlsx", "xlsm"], key="duenos",
                              label_visibility="collapsed")
        if fd:
            hs = cargas.hojas(fd)
            hoja_d = st.selectbox("Hoja", hs)
            df = pd.read_excel(fd, sheet_name=hoja_d)
            cols = [str(c) for c in df.columns]
            guess = lambda *ks: next((i for i, c in enumerate(cols) if any(k in c.lower() for k in ks)), 0)
            c1, c2, c3 = st.columns(3)
            col_sub = c1.selectbox("Columna del sub-rubro", cols, index=guess("sub", "concepto", "rubro"))
            col_ger = c2.selectbox("Columna de la gerencia", cols, index=guess("gerencia", "dueñ", "respons"))
            tiene_cod = any(k in c.lower() for c in cols for k in ("cod", "cód"))
            col_cod = c3.selectbox("Columna del código (opcional)", ["(no tiene)"] + cols,
                                   index=1 + guess("cód", "cod") if tiene_cod else 0)
            d_new = pd.DataFrame({
                "codigo": df[col_cod].astype(str).str.strip().replace("nan", "") if col_cod != "(no tiene)" else "",
                "sub_rubro": df[col_sub].astype(str).str.strip(),
                "gerencia": df[col_ger].astype(str).str.strip()})
            d_new = d_new[~d_new["sub_rubro"].isin(["", "nan"]) & ~d_new["gerencia"].isin(["", "nan"])]
            st.write(f"{len(d_new)} sub-rubros asignados a {d_new['gerencia'].nunique()} gerencias.")
            no_existen = sorted(set(d_new["sub_rubro"].map(cargas.norm)) - set(cat["sub_norm"]))
            if no_existen:
                st.warning(f"{len(no_existen)} sub-rubros no aparecen en el presupuesto {anio} (revisa la escritura): "
                           + ", ".join(no_existen[:15]) + ("…" if len(no_existen) > 15 else ""))
            sin_usuario = sorted(set(d_new["gerencia"]) - set(ctx.usuarios))
            if sin_usuario:
                st.warning("Gerencias sin usuario en secrets.toml (no podrán ingresar): " + ", ".join(sin_usuario))
            if st.button("Confirmar asignación", type="primary", key="asig_file"):
                asignacion.guardar(con, d_new)
                ctx.registrar_carga(ctx.usuario, "Asignación", fd.name, f"{len(d_new)} sub-rubros")
                st.session_state["asig_v"] = st.session_state.get("asig_v", 0) + 1
                st.success("Asignación actualizada.")
                st.rerun()

# ------------------------------------------------------------------ combustible
CAMPOS_COMB = ["fuente", "fecha", "anio", "mes", "placa", "tipo", "galones", "valor", "contrato", "documento",
               "estacion", "hora", "km"]


def _firma_comb(df: pd.DataFrame) -> pd.Series:
    """Texto comparable de los campos de un tanqueo (para saber si cambió)."""
    d = df[CAMPOS_COMB].copy()
    d["galones"] = pd.to_numeric(d["galones"], errors="coerce").round(3)
    d["valor"] = pd.to_numeric(d["valor"], errors="coerce").round(2)
    for c in ("anio", "mes"):
        d[c] = pd.to_numeric(d[c], errors="coerce").fillna(0).astype(int)
    return d.astype(str).fillna("").agg("|".join, axis=1)


@st.cache_data(show_spinner=False, max_entries=6)
def _leer_comb(contenido: bytes, nombre: str):
    """Lectura del archivo, guardada en memoria para no releerlo en cada clic."""
    from io import BytesIO
    return cargas.leer_relacion_combustible(BytesIO(contenido), nombre)


def _cargar_combustible(ctx):
    con, leer, usuario = ctx.con, ctx.leer, ctx.usuario
    with st.container(key="card_adm_comb"):
        titulo_tarjeta("Relación de combustible",
                       "Sube las relaciones tal como las entrega el área (ALDC, ALSM): la app reconoce sola las hojas "
                       "de cada proveedor (Petrobras, Terpel, Autogas), crea los tanqueos nuevos, actualiza los que "
                       "cambiaron y no duplica los que ya están.")
        fcs = st.file_uploader("Archivos de combustible", type=["xlsx", "xlsm"], key="comb",
                               accept_multiple_files=True, label_visibility="collapsed")
        if not fcs:
            return
        leidos, aguas = [], {}
        filas_info = []
        for fc in fcs:
            try:
                with st.spinner(f"Leyendo {fc.name}…"):
                    d, info = _leer_comb(fc.getvalue(), fc.name)
                d, info = d.copy(), {**info, "aguas": set(info["aguas"])}
            except Exception as err:
                st.error(f"No pude leer {fc.name}: {err}")
                continue
            leidos.append(d)
            for emp in d["empresa"].unique():
                aguas.setdefault(emp, set())
                if info["aguas"]:
                    aguas[emp] |= info["aguas"]
            for hoja, h in info["hojas"].items():
                sub = d[d["fuente"] == h["proveedor"]]
                filas_info.append({"Archivo": fc.name, "Empresa": ", ".join(sorted(sub["empresa"].unique())),
                                   "Hoja": hoja, "Tanqueos": len(sub), "Galones": sub["galones"].sum(),
                                   "Valor (M)": sub["valor"].sum() / 1e6,
                                   "Periodo": (f"{MESES[sub['mes'].min()-1]}–{MESES[sub['mes'].max()-1]} "
                                               f"{sub['anio'].max()}") if len(sub) else ""})
            for a in info["avisos"]:
                st.caption(f"{fc.name}: {a}")
            if info["ignoradas"]:
                st.caption(f"{fc.name}: hojas sin formato de proveedor (no se leen): {', '.join(info['ignoradas'])}.")
        if not leidos:
            return
        st.dataframe(pd.DataFrame(filas_info), hide_index=True, width="stretch", column_config={
            "Galones": st.column_config.NumberColumn(format="%.0f"),
            "Valor (M)": st.column_config.NumberColumn(format="%.1f")})

        nuevos = pd.concat(leidos, ignore_index=True).drop_duplicates("clave", keep="last")
        # placas del contrato Aguas de Bogotá: las del archivo (columna AGUAS) + las ya guardadas
        guardadas = leer("SELECT empresa, placa FROM combustible_aguas")
        lista = {e: set(guardadas.loc[guardadas["empresa"] == e, "placa"]) for e in nuevos["empresa"].unique()}
        for e, ps in aguas.items():
            if ps:
                lista[e] = ps               # el archivo trae la lista vigente: manda
        nuevos["contrato"] = [("Aguas de Bogotá" if p in lista.get(e, set()) else "Operación general")
                              for e, p in zip(nuevos["empresa"], nuevos["placa"])]

        # comparación con lo que ya está en la app
        emps = sorted(nuevos["empresa"].unique())
        marcas = ",".join("?" * len(emps))
        actual = leer(f"SELECT * FROM combustible WHERE empresa IN ({marcas})", tuple(emps))
        if "clave" not in actual:
            actual["clave"] = ""
        actual["clave"] = actual["clave"].fillna("")
        viejos = actual[actual["clave"] == ""]
        pares = set(zip(nuevos["empresa"], nuevos["anio"].astype(int)))
        viejos = viejos[pd.Series([(e, int(a)) in pares for e, a in zip(viejos["empresa"], viejos["anio"])],
                                  index=viejos.index, dtype=bool)]
        con_clave = actual[actual["clave"] != ""].set_index("clave")
        existe = nuevos["clave"].isin(con_clave.index)
        firma_n = _firma_comb(nuevos)
        firma_a = _firma_comb(con_clave.reindex(nuevos.loc[existe, "clave"]).reset_index(drop=True)) \
            if existe.any() else pd.Series(dtype=str)
        cambia = pd.Series(False, index=nuevos.index)
        if existe.any():
            cambia.loc[existe] = firma_n[existe].values != firma_a.values
        crear, actualizar = nuevos[~existe], nuevos[cambia]
        sin_cambio = int(existe.sum() - cambia.sum())
        solo_app = int((~con_clave.index.isin(nuevos["clave"])).sum())

        st.html(E.kpis([
            dict(label="Nuevos", valor=f"{len(crear):,}".replace(",", "."), sub=E.millones(crear["valor"].sum()),
                 icono="check", color=E.VERDE, fondo=E.VERDE_SUAVE),
            dict(label="Se actualizan", valor=f"{len(actualizar):,}".replace(",", "."),
                 sub="Ya estaban, con algún dato distinto", icono="receipt", color=E.AMBAR, fondo=E.AMBAR_SUAVE),
            dict(label="Sin cambios", valor=f"{sin_cambio:,}".replace(",", "."), sub="No se tocan",
                 icono="wallet", color=E.AZUL, fondo=E.AZUL_SUAVE),
            dict(label="Formato anterior", valor=f"{len(viejos):,}".replace(",", "."),
                 sub="Cargados antes sin identificación: se reemplazan", icono="alert", color=E.AZUL,
                 fondo=E.AZUL_SUAVE),
        ]))
        if solo_app:
            st.caption(f"{solo_app:,} tanqueos que ya están en la app no vienen en estos archivos; se conservan."
                       .replace(",", "."))
        if not actualizar.empty:
            with st.expander(f"Ver los {len(actualizar)} tanqueos que se actualizan", icon=":material/compare_arrows:"):
                antes = con_clave.reindex(actualizar["clave"]).reset_index()
                comp = actualizar[["clave", "fecha", "placa", "fuente", "galones", "valor", "mes", "contrato"]] \
                    .merge(antes[["clave", "galones", "valor", "mes", "contrato"]], on="clave", suffixes=("", " antes"))
                st.dataframe(comp.drop(columns="clave"), hide_index=True, width="stretch")
        hay_algo = len(crear) or len(actualizar) or len(viejos)
        if st.button("Cargar combustible", type="primary", disabled=not hay_algo, key="comb_ok"):
            try:
                for e, a in sorted({(e, int(a)) for e, a in zip(viejos["empresa"], viejos["anio"])}):
                    con.execute("DELETE FROM combustible WHERE empresa=? AND anio=? AND (clave IS NULL OR clave='')",
                                (e, a))
                if not actualizar.empty:
                    con.executemany("DELETE FROM combustible WHERE clave=?", [(c,) for c in actualizar["clave"]])
                con.anexar(pd.concat([crear, actualizar])[cargas.COLS_COMBUSTIBLE], "combustible")
                for e, ps in lista.items():
                    con.execute("DELETE FROM combustible_aguas WHERE empresa=?", (e,))
                    con.executemany("INSERT INTO combustible_aguas(empresa, placa) VALUES(?,?)",
                                    [(e, p) for p in sorted(ps)])
                con.commit()
            except Exception:
                con.rollback()
                raise
            ctx.registrar_carga(usuario, "Combustible", ", ".join(f.name for f in fcs),
                                f"{len(crear)} nuevos · {len(actualizar)} actualizados · {sin_cambio} sin cambios"
                                + (f" · {len(viejos)} de formato anterior reemplazados" if len(viejos) else ""))
            st.success(f"Combustible cargado: {len(crear):,} nuevos, {len(actualizar):,} actualizados."
                       .replace(",", "."))
            st.rerun()
        elif not hay_algo:
            st.success("Todo lo de estos archivos ya está cargado y sin cambios.")

# ------------------------------------------------------------------ estados financieros
@st.cache_data(show_spinner=False, max_entries=4)
def _leer_eeff(tipo: str, contenido: bytes):
    from io import BytesIO
    return eeff.leer_sap(BytesIO(contenido)) if tipo == "sap" else eeff.leer_historico(BytesIO(contenido))


def _cargar_eeff(ctx):
    con, leer, usuario, emp = ctx.con, ctx.leer, ctx.usuario, ctx.empresa
    cargados = leer("SELECT anio, mes, MAX(fuente) fuente FROM eeff WHERE empresa=? GROUP BY anio, mes "
                    "ORDER BY anio, mes", (emp,))
    with st.container(key="card_adm_ef1"):
        titulo_tarjeta(f"Balance de SAP del mes · {emp}",
                       "Sube el balance que descargas de SAP (transacción F.01) para un mes: columnas Sociedad, Pos. "
                       "balance/PyG, Texto, División, Tipo de moneda, Moneda, saldo inicial (cierre del mes anterior), "
                       "saldo final (cierre del mes), desviaciones. Con él se arman el balance, el estado de resultados "
                       "y los indicadores. Si el mes ya estaba, se reemplaza.")
        if not cargados.empty:
            u = cargados.iloc[-1]
            st.caption(f"Último mes cargado: {MESES[int(u['mes']) - 1]} {int(u['anio'])} "
                       f"({'SAP' if u['fuente'] == 'sap' else 'histórico'}) · {len(cargados)} meses en total.")
        f = st.file_uploader("Balance F.01 de SAP", type=["xlsx", "xlsm", "xls"], key="up_eeff_sap",
                             label_visibility="collapsed")
        if f:
            try:
                sap = _leer_eeff("sap", f.getvalue()).copy()
            except Exception as err:
                st.error(f"No pude leer el archivo: {err}")
                sap = None
            if sap is not None:
                _aviso_empresa(ctx, f.name)
                if cargados.empty:
                    sig_a, sig_m = ahora().year, ahora().month
                else:
                    a0, m0 = int(cargados.iloc[-1]["anio"]), int(cargados.iloc[-1]["mes"])
                    sig_a, sig_m = (a0 + 1, 1) if m0 == 12 else (a0, m0 + 1)
                c1, c2, _ = st.columns([0.7, 1, 2])
                anio = int(c1.number_input("Año", value=sig_a, step=1, format="%d", key="ef_up_anio"))
                mes = c2.selectbox("Mes del balance", list(range(1, 13)), index=sig_m - 1, key="ef_up_mes",
                                   format_func=lambda m: MESES[m - 1])
                pa, pm = (anio - 1, 12) if mes == 1 else (anio, mes - 1)
                ant = leer("SELECT idx, valor FROM eeff WHERE empresa=? AND anio=? AND mes=? AND estado='ESF'",
                           (emp, pa, pm))
                nc_prev = {int(i): float(v) for i, v in zip(ant["idx"], ant["valor"])
                           if eeff.ESF[int(i)][3] == "nc"}
                calc = eeff.calcular(sap, nc_prev)
                t = eeff.totales_esf(calc[calc["estado"] == "ESF"])
                pyg = calc[calc["estado"] == "PYG"]
                res = sum(eeff.SIGNO_GRUPO[g] * v for g, v in zip(pyg["grupo"], pyg["valor"]))
                # el saldo inicial debe ser el final del mes anterior: se compara el resultado acumulado
                ini = eeff.calcular(sap.assign(final=sap["inicial"]), nc_prev)
                pi = ini[ini["estado"] == "PYG"]
                res_ini = sum(eeff.SIGNO_GRUPO[g] * v for g, v in zip(pi["grupo"], pi["valor"]))
                prev_res = leer("SELECT grupo, SUM(valor) v FROM eeff WHERE empresa=? AND anio=? AND mes=? AND "
                                "estado='PYG' GROUP BY grupo", (emp, pa, pm))
                res_prev = sum(eeff.SIGNO_GRUPO.get(g, 0) * v for g, v in zip(prev_res["grupo"], prev_res["v"]))
                res_mes = res - (res_ini if mes > 1 else 0.0)
                st.html(E.kpis([
                    dict(label="Total activo", valor=E.millones(t["ACTIVO"]), sub=f"{len(sap):,} filas de SAP"
                         .replace(",", "."), icono="wallet", color=E.AZUL, fondo=E.AZUL_SUAVE),
                    dict(label="Pasivo + patrimonio", valor=E.millones(t["PASIVO + PATRIMONIO"]),
                         sub=f"Diferencia {E.millones(t['ACTIVO'] - t['PASIVO + PATRIMONIO'])}", icono="check",
                         color=E.VERDE if abs(t["ACTIVO"] - t["PASIVO + PATRIMONIO"]) < 1000 else E.ROJO,
                         fondo=E.VERDE_SUAVE if abs(t["ACTIVO"] - t["PASIVO + PATRIMONIO"]) < 1000 else E.ROJO_SUAVE),
                    dict(label="Utilidad acumulada", valor=E.millones(res), sub=f"{MESES[mes - 1]} {anio}",
                         icono="piggy", color=E.AZUL, fondo=E.AZUL_SUAVE),
                    dict(label="Utilidad del mes", valor=E.millones(res_mes), sub="Saldo final − saldo inicial",
                         icono="calendar", color=E.AZUL, fondo=E.AZUL_SUAVE),
                ]))
                if abs(t["ACTIVO"] - t["PASIVO + PATRIMONIO"]) >= 1000:
                    st.warning("El balance no cuadra (activo ≠ pasivo + patrimonio). Revisa que el archivo sea el F.01 "
                               "completo del mes y que la columna de la derecha sea el saldo final.")
                if mes > 1 and not prev_res.empty and abs(res_ini - res_prev) > 1000:
                    st.warning(f"El saldo inicial del archivo da una utilidad acumulada de {E.millones(res_ini)}, pero "
                               f"la guardada para {MESES[pm - 1]} {pa} es {E.millones(res_prev)}. Revisa el mes elegido "
                               "o el orden de las columnas de saldo (izquierda = inicial, derecha = final).")
                if not nc_prev:
                    st.info("No hay mes anterior cargado: la porción no corriente queda en cero (la puedes ajustar "
                            "abajo después de cargar).")
                if st.button(f"Cargar balance de {MESES[mes - 1]} {anio}", type="primary", key="ef_ok_sap"):
                    con.execute("DELETE FROM eeff WHERE empresa=? AND anio=? AND mes=?", (emp, anio, mes))
                    con.execute("DELETE FROM eeff_sap WHERE empresa=? AND anio=? AND mes=?", (emp, anio, mes))
                    con.anexar(calc.assign(empresa=emp, anio=anio, mes=mes, fuente="sap")
                               [["empresa", "anio", "mes", "estado", "idx", "grupo", "subgrupo", "linea", "valor",
                                 "fuente"]], "eeff")
                    con.anexar(sap.assign(empresa=emp, anio=anio, mes=mes)
                               [["empresa", "anio", "mes", "sociedad", "posicion", "texto", "inicial", "final"]],
                               "eeff_sap")
                    con.commit()
                    ctx.registrar_carga(usuario, "Estados financieros", f.name,
                                        f"{emp} · {MESES[mes - 1]} {anio} · utilidad acum. {E.millones(res)}")
                    st.success("Balance cargado.")
                    st.rerun()

    with st.container(key="card_adm_ef3"):
        titulo_tarjeta(f"Porción no corriente · {emp}",
                       "Lo que SAP no separa: la parte a más de un año de cuentas por cobrar, obligaciones financieras, "
                       "cuentas por pagar e impuestos. Al cargar un mes se copia la del mes anterior; aquí la ajustas "
                       "(la parte corriente se recalcula sola para que el total no cambie).")
        if cargados.empty:
            st.caption("Primero carga el histórico o un mes de SAP.")
        else:
            ops = [(int(a), int(m)) for a, m in zip(cargados["anio"], cargados["mes"])]
            sel = st.selectbox("Mes", ops[::-1], key="ef_nc_mes", format_func=lambda x: f"{MESES[x[1] - 1]} {x[0]}")
            cur = leer("SELECT idx, valor FROM eeff WHERE empresa=? AND anio=? AND mes=? AND estado='ESF'",
                       (emp, sel[0], sel[1]))
            vals = dict(zip(cur["idx"].astype(int), cur["valor"]))
            nc_idx = [i for i, rg in enumerate(eeff.ESF) if rg[3] == "nc"]
            t = pd.DataFrame({"Grupo": [eeff.ESF[i][1].strip().capitalize() for i in nc_idx],
                              "Línea": [eeff.ESF[i][2].capitalize() for i in nc_idx],
                              "Millones": [vals.get(i, 0.0) / 1e6 for i in nc_idx]}, index=nc_idx)
            ed = st.data_editor(t, hide_index=True, width="stretch", disabled=["Grupo", "Línea"], key="ed_nc",
                                column_config={"Millones": st.column_config.NumberColumn(format="%.1f")})
            if st.button("Guardar porción no corriente", key="ef_nc_ok"):
                for i in nc_idx:
                    nuevo = float(ed.at[i, "Millones"]) * 1e6
                    viejo = float(vals.get(i, 0.0))
                    if abs(nuevo - viejo) < 0.5:
                        continue
                    con.execute("UPDATE eeff SET valor=? WHERE empresa=? AND anio=? AND mes=? AND estado='ESF' "
                                "AND idx=?", (nuevo, emp, sel[0], sel[1], i))
                    for j, rg in enumerate(eeff.ESF):     # la parte corriente absorbe el cambio
                        if rg[3] in ("C", "N", "R") and rg[7] == i:
                            con.execute("UPDATE eeff SET valor=valor+? WHERE empresa=? AND anio=? AND mes=? AND "
                                        "estado='ESF' AND idx=?", (viejo - nuevo, emp, sel[0], sel[1], j))
                con.commit()
                st.success("Porción no corriente guardada.")
                st.rerun()

    with st.container(key="card_adm_ef2"):
        titulo_tarjeta(f"Histórico de balance y PyG · {emp}",
                       "Archivo con las hojas ESF y PYG (SUBGRUPO 1/2/3 y un mes por columna), como "
                       "'02_BALANCE Y PYG'. Carga todos los meses que trae (el PyG viene acumulado del año); los meses "
                       "cargados desde SAP se conservan.")
        fh = st.file_uploader("Histórico", type=["xlsx", "xlsm"], key="up_eeff_hist", label_visibility="collapsed")
        if fh:
            try:
                h = _leer_eeff("hist", fh.getvalue()).copy()
            except Exception as err:
                st.error(f"No pude leer el archivo: {err}")
                h = pd.DataFrame()
            if not h.empty:
                # solo meses con balance (el archivo puede traer columnas futuras en cero)
                con_datos = h[h["estado"] == "ESF"].groupby(["anio", "mes"])["valor"].apply(lambda s: s.abs().sum())
                con_datos = con_datos[con_datos > 0].index
                h = h.set_index(["anio", "mes"]).loc[lambda d: d.index.isin(con_datos)].reset_index()
                n_esf = (h["estado"] == "ESF").sum()
                ya_sap = set(zip(cargados.loc[cargados["fuente"] == "sap", "anio"],
                                 cargados.loc[cargados["fuente"] == "sap", "mes"]))
                h = h[[(a, m) not in ya_sap for a, m in zip(h["anio"], h["mes"])]]
                meses_h = sorted(set(zip(h["anio"], h["mes"])))
                if not meses_h:
                    st.info("Todos los meses del archivo ya están cargados desde SAP.")
                else:
                    st.write(f"**{len(meses_h)} meses** ({MESES[meses_h[0][1] - 1]} {meses_h[0][0]} a "
                             f"{MESES[meses_h[-1][1] - 1]} {meses_h[-1][0]}) · {n_esf:,} valores de balance."
                             .replace(",", "."))
                    if len(set(h.loc[h["estado"] == "ESF", "idx"])) != len(eeff.ESF) or \
                            len(set(h.loc[h["estado"] == "PYG", "idx"])) != len(eeff.PYG):
                        st.warning("El archivo no tiene exactamente las mismas líneas de balance y PyG que la app; "
                                   "revisa que sea el formato '02_BALANCE Y PYG'.")
                    if st.button("Cargar histórico", type="primary", key="ef_ok_hist"):
                        for a, m in meses_h:
                            con.execute("DELETE FROM eeff WHERE empresa=? AND anio=? AND mes=?", (emp, int(a), int(m)))
                        con.anexar(h.assign(empresa=emp, fuente="historico")
                                   [["empresa", "anio", "mes", "estado", "idx", "grupo", "subgrupo", "linea", "valor",
                                     "fuente"]], "eeff")
                        con.commit()
                        ctx.registrar_carga(usuario, "Estados financieros", fh.name,
                                            f"{emp} · histórico {len(meses_h)} meses")
                        st.success("Histórico cargado.")
                        st.rerun()


# ------------------------------------------------------------------ nómina
@st.cache_data(show_spinner=False, max_entries=6)
def _leer_nom(tipo: str, contenido: bytes, version: int = 0):
    # 'version' cambia cuando cambia la forma de leer los archivos: así no se reutiliza una lectura vieja en memoria
    from io import BytesIO
    f = {"ppto": nomina.leer_presupuesto, "ejec": nomina.leer_ejecucion, "ceco": nomina.leer_ceco}[tipo]
    return f(BytesIO(contenido))


def _cargar_nomina(ctx):
    con, leer, usuario = ctx.con, ctx.leer, ctx.usuario
    with st.container(key="card_adm_nom"):
        titulo_tarjeta("Nómina · archivos",
                       "Sube los archivos tal como los tienes. Presupuesto: reemplaza los años que trae (hojas "
                       "'3.0 Personal 2026', '… 2027'). Ejecución: reemplaza el mes que trae (no duplica si lo "
                       "vuelves a subir). Centros de coste: la relación centro de coste → área.")
        c1, c2, c3 = st.columns(3)
        fp = c1.file_uploader("Presupuesto detalle nómina", type=["xlsx", "xlsm"], key="nom_ppto")
        fe = c2.file_uploader("Ejecución nómina (SAP)", type=["xlsx", "xlsm"], key="nom_ejec")
        fc = c3.file_uploader("Relación centros de coste", type=["xlsx", "xlsm"], key="nom_ceco")
        for archivo, tipo in ((fp, "ppto"), (fe, "ejec"), (fc, "ceco")):
            if not archivo:
                continue
            try:
                with st.spinner(f"Leyendo {archivo.name}…"):
                    d = _leer_nom(tipo, archivo.getvalue(), nomina.VERSION_LECTOR).copy()
            except Exception as err:
                st.error(f"No pude leer {archivo.name}: {err}")
                continue
            if tipo == "ppto":
                res = d.groupby("anio").agg(meses=("mes", lambda m: f"{MESES[m.min()-1]}–{MESES[m.max()-1]}"),
                                            cargos=("cargo", "nunique"), valor=("valor", "sum"))
                st.markdown(f"**{archivo.name}:** " + " · ".join(
                    f"{a}: {r.meses}, {r.cargos} cargos, {E.millones(r.valor)}" for a, r in res.iterrows()))
                # validación: el costo de cada cargo recalculado con la fórmula del archivo, por concepto
                conc = d.groupby("anio")[nomina.COLS_CONCEPTO].sum()
                nombres = dict(nomina.CONCEPTOS)
                tabla = conc.T.rename(index=nombres)
                tabla.columns = [str(a) for a in tabla.columns]
                tabla.loc["Total recalculado"] = tabla.drop(index=nombres["c_otros"]).sum()
                tabla.loc["Total del archivo"] = res["valor"].set_axis([str(a) for a in res.index])
                st.dataframe((tabla / 1e6).round(1), width="stretch", column_config={
                    str(a): st.column_config.NumberColumn(f"{a} (millones)", format="%.1f") for a in tabla.columns})
                no_expl = conc["c_otros"].abs().sum()
                if no_expl > 1000:
                    st.warning(f"La fórmula no explica {E.millones(no_expl)} del costo del archivo: ese valor queda en "
                               "'Otros conceptos'. Revisa si cambiaron las columnas de salario, horas extra, "
                               "bonificaciones o porcentajes.")
                else:
                    st.caption("✓ El costo de cada cargo se recalculó con la fórmula del archivo y cuadra al peso: "
                               "(salario + bonificación de cumplimiento + horas extra del mes + subsidio de transporte "
                               "+ bonificación extralegal + % seg. × (salario + bonificación + horas) + % vacaciones × "
                               "(salario + bonificación) + % prestacional × (salario + bonificación + horas + "
                               "transporte)) × personas del mes.")
                if st.button("Cargar presupuesto de nómina", type="primary", key="nom_ok_ppto"):
                    for a in res.index:
                        con.execute("DELETE FROM nomina_ppto WHERE anio=?", (int(a),))
                    con.anexar(d, "nomina_ppto")
                    con.commit()
                    ctx.registrar_carga(usuario, "Nómina presupuesto", archivo.name,
                                        f"{len(res)} años · {E.millones(d['valor'].sum())}")
                    st.success("Presupuesto de nómina cargado.")
                    st.rerun()
            elif tipo == "ejec":
                per = sorted(set(zip(d["anio"], d["mes"])))
                txt = ", ".join(f"{MESES[m-1]} {a}" for a, m in per)
                ya = [p for p in per if not leer("SELECT 1 FROM nomina_ejec WHERE anio=? AND mes=? LIMIT 1",
                                                  (int(p[0]), int(p[1]))).empty]
                st.markdown(f"**{archivo.name}:** {txt} · {d['empleado_id'].nunique():,} personas · "
                            .replace(",", ".") + f"{d['cargo'].nunique()} cargos · {E.millones(d['valor'].sum())}"
                            + (" · **reemplaza lo ya cargado de ese mes**" if ya else ""))
                if st.button("Cargar ejecución de nómina", type="primary", key="nom_ok_ejec"):
                    for a, m in per:
                        con.execute("DELETE FROM nomina_ejec WHERE anio=? AND mes=?", (int(a), int(m)))
                    con.anexar(d, "nomina_ejec")
                    con.commit()
                    ctx.registrar_carga(usuario, "Nómina ejecución", archivo.name,
                                        f"{txt} · {E.millones(d['valor'].sum())}")
                    st.success("Ejecución de nómina cargada.")
                    st.rerun()
            else:
                st.markdown(f"**{archivo.name}:** {len(d)} centros de coste · {d['area'].nunique()} áreas")
                if st.button("Cargar relación de centros de coste", type="primary", key="nom_ok_ceco"):
                    con.execute("DELETE FROM nomina_ceco")
                    con.anexar(d, "nomina_ceco")
                    con.commit()
                    ctx.registrar_carga(usuario, "Nómina centros de coste", archivo.name, f"{len(d)} centros")
                    st.success("Relación de centros de coste cargada.")
                    st.rerun()
        # centros de coste de la nómina sin área
        sin = leer("SELECT DISTINCT e.cc FROM nomina_ejec e LEFT JOIN nomina_ceco c ON c.cc = e.cc "
                   "WHERE c.area IS NULL")
        if not sin.empty:
            st.warning("Centros de coste de la nómina sin área en la relación: " + ", ".join(sin["cc"])
                       + ". Agrégalos al archivo de centros de coste y vuelve a subirlo.")


AUTO = "(sugerencia automática)"


def _homologar_cargos(ctx):
    con, leer = ctx.con, ctx.leer
    with st.container(key="card_adm_nomh"):
        titulo_tarjeta("Homologar cargos",
                       "Cada cargo del presupuesto y de la nómina de SAP se lleva a un 'cargo común' para cruzarlos, "
                       "por centro de coste: el mismo cargo puede ir a un cargo común distinto según el centro de coste "
                       "donde esté. La app ya propone la mayoría (quita '(A)', el turno y las tildes; los conductores "
                       "van a 'Conductor'). Aquí asignas los que faltan o corriges los que no te gusten.")
        ppto = leer("SELECT cargo, area, anio, mes, personas, valor FROM nomina_ppto")
        ejec = leer("SELECT cargo, cc, anio, mes, empleado_id, valor FROM nomina_ejec")
        if ppto.empty and ejec.empty:
            st.caption("Primero carga el presupuesto y la ejecución de nómina.")
            return
        comunes = nomina.comunes_ppto(ppto) if not ppto.empty else {}
        manual = nomina.homologaciones(leer)
        partes = []
        # SAP: una fila por cargo y centro de coste (cifras del último mes cargado)
        if not ejec.empty:
            ceco = dict(leer("SELECT cc, area FROM nomina_ceco").values.tolist())
            ult = ejec[["anio", "mes"]].drop_duplicates().sort_values(["anio", "mes"]).iloc[-1]
            e = ejec[(ejec["anio"] == ult["anio"]) & (ejec["mes"] == ult["mes"])]
            d = e.groupby(["cargo", "cc"]).agg(Personas=("empleado_id", "nunique"), valor=("valor", "sum")) \
                .reset_index()
            viejos = ejec[~ejec.set_index(["cargo", "cc"]).index.isin(d.set_index(["cargo", "cc"]).index)]
            if not viejos.empty:                       # cargo/centro que no aparece en el último mes
                d = pd.concat([d, viejos.groupby(["cargo", "cc"]).agg(
                    Personas=("empleado_id", "nunique"), valor=("valor", "sum")).reset_index()])
            d["fuente"], d["Fuente"], d["llave"] = "ejec", "Nómina SAP", d["cc"]
            d["Área"] = d["cc"].map(ceco).fillna("SIN ÁREA")
            d["auto"] = d["cargo"].map(lambda c: nomina.sugerir(c, comunes) or nomina.SIN_PPTO)
            partes.append(d.rename(columns={"cc": "Centro de coste"}))
        # presupuesto: una fila por cargo y área (no trae centro de coste)
        if not ppto.empty:
            pp = ppto.groupby(["cargo", "area"]).apply(lambda g: pd.Series({
                "Personas": g.loc[g["mes"] == g["mes"].min(), "personas"].sum(), "valor": g["valor"].sum()}),
                include_groups=False).reset_index()
            pp["fuente"], pp["Fuente"], pp["Centro de coste"] = "ppto", "Presupuesto", ""
            pp["llave"] = "AREA:" + pp["area"].fillna("")
            pp["Área"] = pp["area"]
            pp["auto"] = pp["cargo"].map(lambda c: comunes.get(nomina.clave_cargo(c), nomina.nombre_limpio(c)))
            partes.append(pp.drop(columns="area"))
        det = pd.concat(partes, ignore_index=True).rename(columns={"cargo": "Cargo"})
        det["Valor (M)"] = det["valor"] / 1e6
        det["manual"] = [nomina.buscar_manual(manual, f, c, k) for f, c, k in zip(det["fuente"], det["Cargo"],
                                                                                det["llave"])]
        # una asignación guardada que coincide con la sugerencia no cambia nada: se trata como automática
        # (así un cargo que sigue sin presupuesto aparece como "Sin cruce", no como "Manual")
        det["manual"] = pd.Series([None if (not isinstance(m, str) or not m or nomina.norm(m) == nomina.norm(a))
                                   else m for m, a in zip(det["manual"], det["auto"])], index=det.index, dtype=object)
        det["Cargo común"] = det["manual"].where(det["manual"].notna(), det["auto"])
        det["Estado"] = ["Manual" if isinstance(m, str) else ("Sin cruce" if a == nomina.SIN_PPTO else "Automática")
                         for m, a in zip(det["manual"], det["auto"])]
        det["_orden"] = det["Estado"].map({"Sin cruce": 0, "Automática": 1, "Manual": 2})
        det["_tot"] = det.groupby(["fuente", "Cargo"])["Valor (M)"].transform("sum")
        det = det.sort_values(["_orden", "Fuente", "_tot", "Cargo", "Valor (M)"],
                              ascending=[True, True, False, True, False]).reset_index(drop=True)
        guardados = {v for v in manual.values()}
        opciones = sorted((set(comunes.values()) | guardados) - {nomina.SIN_PPTO})

        sin = det[det["Estado"] == "Sin cruce"]
        if not sin.empty:
            st.html(f'<p class="nf-card-s"><b style="color:{E.ROJO}">{sin["Cargo"].nunique()} cargos de SAP sin cruce '
                    f'en {len(sin)} centros de coste</b> ({E.millones(sin["valor"].sum())} en el último mes cargado). '
                    'Asígnales un cargo común abajo.</p>')
        else:
            st.html('<p class="nf-card-s">Todos los cargos de SAP tienen cargo común.</p>')

        # ---- 1. asignar
        v = st.session_state.get("nom_v", 0)
        with st.container(border=True):
            st.markdown("**1. Asignar cargo común**")
            g1, g2 = st.columns([1.3, 2.2])
            ccs_disp = sorted({c for c in det["Centro de coste"] if c})
            f_cc = g1.multiselect("Centro de coste", ccs_disp, placeholder="Todos", key=f"nom_cc_{v}",
                                  format_func=lambda c: f"{c} · {ceco.get(c, '') if not ejec.empty else ''}")
            sub = det[det["Centro de coste"].isin(f_cc)] if f_cc else det
            etiquetas = [f"{'⚠ ' if es == 'Sin cruce' else ''}{fu} · {ca} · "
                         + (f"{cc} ({ar})" if f == "ejec" else ar) + f" · {int(pe)} pers."
                         for es, fu, ca, cc, ar, f, pe in zip(sub["Estado"], sub["Fuente"], sub["Cargo"],
                                                              sub["Centro de coste"], sub["Área"], sub["fuente"],
                                                              sub["Personas"])]
            sel = g2.multiselect(
                "Cargos a asignar (cada uno en su centro de coste)", etiquetas, key=f"nom_sel_{v}",
                placeholder="Escribe para buscar: supervisor, gerente, aprendiz…",
                help="Filtra por centro de coste a la izquierda para ver solo los cargos de ese centro. Los marcados "
                     "con ⚠ son cargos de SAP que hoy no cruzan con el presupuesto.")
            b1, b2 = st.columns([2.2, 1], vertical_alignment="bottom")
            destino = b1.selectbox(
                "Cargo común", opciones, index=None, accept_new_options=True, key=f"nom_dest_{v}",
                placeholder="Elige uno de la lista o escribe uno nuevo (ej. Supervisor de operaciones)")
            elegidos = sub[[x in set(sel) for x in etiquetas]]
            if b2.button("Asignar", type="primary", key="nom_asignar", disabled=not (sel and destino),
                         width="stretch"):
                _guardar_homol(con, det, elegidos, str(destino).strip())
                st.rerun()
            c1, c2 = st.columns(2)
            if c1.button("Dejar fuera del cruce (sin presupuesto)", key="nom_fuera", disabled=not sel):
                _guardar_homol(con, det, elegidos, nomina.SIN_PPTO)
                st.rerun()
            if c2.button("Volver a la sugerencia de la app", key="nom_auto", disabled=not sel):
                _guardar_homol(con, det, elegidos, AUTO)
                st.rerun()

        # ---- 2. revisar
        st.markdown("**2. Cómo quedó cada cargo, por centro de coste**")
        cuenta = det["Estado"].value_counts()
        estados = ["Todos", "Sin cruce", "Automática", "Manual"]
        ver = st.segmented_control(
            "Ver", estados, default="Todos", key="nom_f_ver",
            format_func=lambda x: f"{x} ({len(det) if x == 'Todos' else int(cuenta.get(x, 0))})") or "Todos"
        f1, f3, f4 = st.columns([0.9, 1.3, 1.6])
        fu = f1.selectbox("Fuente", ["Todas", "Nómina SAP", "Presupuesto"], key="nom_f_fuente")
        es = [] if ver == "Todos" else [ver]
        ar = f3.multiselect("Área", sorted(set(det["Área"].dropna())), placeholder="Todas", key="nom_f_area")
        bu = f4.text_input("Buscar", key="nom_f_buscar", placeholder="Cargo, cargo común o centro de coste…")
        vista = det
        if fu != "Todas":
            vista = vista[vista["Fuente"] == fu]
        if es:
            vista = vista[vista["Estado"].isin(es)]
        if ar:
            vista = vista[vista["Área"].isin(ar)]
        if bu:
            vista = vista[vista["Cargo"].str.contains(bu, case=False, regex=False)
                          | vista["Cargo común"].str.contains(bu, case=False, regex=False)
                          | vista["Centro de coste"].str.contains(bu, case=False, regex=False)]
        st.dataframe(vista[["Estado", "Fuente", "Cargo", "Área", "Centro de coste", "Cargo común", "Personas",
                            "Valor (M)"]],
                     hide_index=True, width="stretch", height=420, column_config={
                         "Personas": st.column_config.NumberColumn(format="%.0f", width="small"),
                         "Valor (M)": st.column_config.NumberColumn(format="%.1f", width="small")})
        st.caption("Una fila por cargo y centro de coste (el presupuesto no trae centro de coste: una fila por cargo y "
                   "área). Automática = sugerencia de la app; Manual = asignada por ti, solo para ese centro de coste. "
                   "Para el presupuesto, Personas es la del primer mes y Valor el total del archivo; para SAP, las del "
                   "último mes cargado.")


def _guardar_homol(con, det: pd.DataFrame, elegidos: pd.DataFrame, comun: str):
    """Guarda el cargo común de cada (cargo, centro de coste) elegido. Si el cargo tenía una asignación general (de
    versiones anteriores, para todos sus centros), primero se reparte en sus centros para no perderla en los demás."""
    for (fuente, cargo), g in elegidos.groupby(["fuente", "Cargo"]):
        general = con.execute("SELECT cargo_comun FROM nomina_homol WHERE fuente=? AND cargo=? AND "
                              "(cc IS NULL OR cc='')", (fuente, cargo)).fetchall()
        if general:
            todas = det[(det["fuente"] == fuente) & (det["Cargo"] == cargo)]["llave"]
            for k in set(todas) - set(g["llave"]):
                con.execute("DELETE FROM nomina_homol WHERE fuente=? AND cargo=? AND cc=?", (fuente, cargo, k))
                con.execute("INSERT INTO nomina_homol(fuente, cargo, cargo_comun, cc) VALUES(?,?,?,?)",
                            (fuente, cargo, general[0][0], k))
            con.execute("DELETE FROM nomina_homol WHERE fuente=? AND cargo=? AND (cc IS NULL OR cc='')",
                        (fuente, cargo))
        for k in g["llave"]:
            con.execute("DELETE FROM nomina_homol WHERE fuente=? AND cargo=? AND cc=?", (fuente, cargo, k))
            if comun != AUTO:
                con.execute("INSERT INTO nomina_homol(fuente, cargo, cargo_comun, cc) VALUES(?,?,?,?)",
                            (fuente, cargo, comun, k))
    con.commit()
    st.session_state["nom_v"] = st.session_state.get("nom_v", 0) + 1


# ------------------------------------------------------------------ flujo de caja
@st.cache_data(show_spinner=False, max_entries=3)
def _leer_flujo(tipo: str, contenido: bytes):
    from io import BytesIO
    if tipo == "libro":
        return flujo.leer_libro(BytesIO(contenido))
    if tipo == "terceros":
        return flujo.leer_terceros(BytesIO(contenido))
    return flujo.leer_movimientos(BytesIO(contenido))


def _txt_periodos(periodos) -> str:
    return ", ".join(f"{MESES[m - 1]} {a}" for a, m in periodos)


def _por_tercero(leer, emp) -> dict:
    pt = leer("SELECT tercero, concepto FROM flujo_por_tercero WHERE empresa=?", (emp,))
    return {flujo._n(t): c for t, c in zip(pt["tercero"], pt["concepto"])}


def _fijas(leer, emp) -> pd.DataFrame:
    return leer("SELECT tercero, cuenta, texto, concepto FROM flujo_reglas WHERE empresa=?", (emp,))


def _flujo_mes_sap(ctx):
    con, leer, usuario, emp = ctx.con, ctx.leer, ctx.usuario, ctx.empresa
    est = leer("SELECT DISTINCT anio FROM flujo_estructura WHERE empresa=? ORDER BY anio", (emp,))
    ult = leer("SELECT anio, mes, COUNT(*) n FROM flujo_mov WHERE empresa=? GROUP BY anio, mes ORDER BY anio, mes",
               (emp,))
    with st.container(key="card_adm_fl_sap"):
        titulo_tarjeta(f"Movimientos de bancos del mes (SAP) · {emp}",
                       "Sube el reporte que descargas de SAP con las columnas de la A a la T (Clase de documento, Nº "
                       "documento, …, Anulado con), como lo pegas en la hoja Journal. La app lleva cada movimiento a su "
                       "concepto de caja con lo que aprendió de tus clasificaciones; lo que no entienda queda pendiente "
                       "para que lo decidas en 'Revisar y clasificar'. Si el mes ya estaba cargado se reemplaza, "
                       "conservando lo que ya clasificaste a mano.")
        if not ult.empty:
            u = ult.iloc[-1]
            st.caption(f"Último mes con movimientos: {MESES[int(u['mes']) - 1]} {int(u['anio'])} · "
                       f"{int(ult['n'].sum()):,} movimientos cargados.".replace(",", "."))
        if est.empty:
            st.info("Primero carga una vez el libro de flujo de caja (pestaña 'Libro inicial y terceros'): de ahí salen "
                    "la estructura del flujo, los ajustes y las clasificaciones con las que la app aprende.",
                    icon=":material/info:")
            return
        f = st.file_uploader("Reporte de SAP", type=["xlsx", "xlsm", "xls"], key="up_fl_sap",
                             label_visibility="collapsed")
        if not f:
            return
        try:
            nuevos = _leer_flujo("sap", f.getvalue()).copy()
        except Exception as err:
            st.error(f"No pude leer el archivo: {err}")
            return
        if nuevos.empty:
            st.warning("El archivo no tiene movimientos con fecha e importe.")
            return
        _aviso_empresa(ctx, f.name)
        nombres = flujo.nombres_tercero(leer, emp)
        nuevos["tercero"] = nuevos["ter_cod"].map(nombres).fillna(nuevos["ref3"])
        existentes = flujo.movimientos(leer, emp)
        a_max = int(nuevos["anio"].max())
        a_est = est["anio"].astype(int)
        a_ref = int(a_est[a_est <= a_max].max()) if (a_est <= a_max).any() else int(a_est.min())
        with st.spinner("Clasificando movimientos…"):
            m = flujo.preparar(nuevos, existentes, _fijas(leer, emp), _por_tercero(leer, emp),
                               validos=flujo.conceptos(leer, emp, a_ref))
        periodos = sorted(set(zip(m["anio"].astype(int), m["mes"].astype(int))))
        ya = [p for p in periodos if ((ult["anio"] == p[0]) & (ult["mes"] == p[1])).any()]
        alta = m["origen"].isin(["validado", "manual"]) | (m["confianza"] == "Alta")
        media = (m["origen"] == "automatica") & (m["confianza"] == "Media")
        pend = m["origen"] == "pendiente"
        st.html(E.kpis([
            dict(label="Movimientos", valor=f"{len(m):,}".replace(",", "."), sub=_txt_periodos(periodos),
                 icono="receipt", color=E.AZUL, fondo=E.AZUL_SUAVE),
            dict(label="Clasificados con confianza", valor=f"{alta.sum():,}".replace(",", "."),
                 sub=f"{alta.mean():.0%} de los movimientos", icono="check", color=E.VERDE, fondo=E.VERDE_SUAVE),
            dict(label="Para revisar", valor=f"{media.sum():,}".replace(",", "."),
                 sub=f"Confianza media · {E.millones(m.loc[media, 'valor'].abs().sum())}", icono="calendar",
                 color=E.AZUL, fondo=E.AZUL_SUAVE),
            dict(label="Pendientes", valor=f"{pend.sum():,}".replace(",", "."),
                 sub=f"Sin clasificar · {E.millones(m.loc[pend, 'valor'].abs().sum())}", icono="wallet",
                 color=E.ROJO if pend.any() else E.VERDE, fondo=E.ROJO_SUAVE if pend.any() else E.VERDE_SUAVE),
        ]))
        if ya:
            st.warning(f"Ya hay movimientos de {_txt_periodos(ya)}: se reemplazan por los de este archivo (lo que "
                       "clasificaste a mano se conserva).", icon=":material/sync:")
        res = m[m["concepto"] != ""].groupby("concepto")["valor"].agg(["size", "sum"]) \
            .sort_values("sum", key=abs, ascending=False)
        with st.expander("Ver resultado por concepto"):
            st.dataframe(res.rename(columns={"size": "Movimientos", "sum": "Valor"}), width="stretch",
                         column_config={"Valor": st.column_config.NumberColumn(format="accounting")})
        if st.button(f"Cargar movimientos de {_txt_periodos(periodos)}", type="primary", key="fl_ok_sap"):
            cuentas = leer("SELECT cuenta, nombre FROM flujo_cuentas WHERE empresa=?", (emp,))
            anios = sorted({a for a, _ in periodos})
            for a in anios:
                if not flujo.asegurar_estructura(con, emp, a):
                    st.error(f"No hay estructura de flujo para {a}.")
                    return
            for a, mes in periodos:
                con.execute("DELETE FROM flujo_mov WHERE empresa=? AND anio=? AND mes=?", (emp, a, mes))
            con.anexar(flujo.a_tabla(m, emp, nombres, dict(zip(cuentas["cuenta"], cuentas["nombre"]))), "flujo_mov")
            for a in anios:
                flujo.recalcular(con, emp, a)
            con.commit()
            ctx.registrar_carga(usuario, "Flujo de caja (SAP)", f.name,
                                f"{emp} · {_txt_periodos(periodos)} · {len(m)} movimientos · {pend.sum()} pendientes")
            st.session_state["fl_msg"] = (f"Movimientos cargados. {pend.sum()} quedaron pendientes y {media.sum()} "
                                          "para revisar: hazlo en 'Revisar y clasificar'.")
            st.rerun()
    if st.session_state.get("fl_msg"):
        st.success(st.session_state.pop("fl_msg"))


def _flujo_revisar(ctx):
    con, leer, emp = ctx.con, ctx.leer, ctx.empresa
    with st.container(key="card_adm_fl_rev"):
        titulo_tarjeta(f"Revisar y clasificar movimientos de caja · {emp}",
                       "Elige el concepto de los movimientos pendientes (o corrige los sugeridos) y guarda. Marca "
                       "'Crear regla' para que los próximos movimientos de ese tercero en esa cuenta se clasifiquen "
                       "solos. Lo que guardes aquí también le enseña a la app.")
        mov = flujo.movimientos(leer, emp)
        if mov.empty:
            st.info("Aún no hay movimientos de caja cargados de esta compañía.")
            return
        mov["concepto"] = mov["concepto"].fillna("")
        pend_tot = mov[mov["origen"] == "pendiente"]
        if not pend_tot.empty:
            g = pend_tot.groupby(["anio", "mes"]).size()
            st.warning("Pendientes por clasificar: " + ", ".join(f"{MESES[int(m) - 1]} {int(a)} ({n})"
                                                                   for (a, m), n in g.items()),
                       icon=":material/pending_actions:")
        periodos = sorted(set(zip(mov["anio"].astype(int), mov["mes"].astype(int))), reverse=True)
        con_pend = [p for p in periodos if ((pend_tot["anio"] == p[0]) & (pend_tot["mes"] == p[1])).any()]
        c1, c2, c3 = st.columns([1, 1.6, 1.4], vertical_alignment="bottom")
        per = c1.selectbox("Mes", periodos, index=periodos.index(con_pend[0]) if con_pend else 0, key="fl_rev_per",
                           format_func=lambda p: f"{MESES[p[1] - 1]} {p[0]}")
        filtro = c2.segmented_control("Ver", ["Pendientes", "Para revisar", "Todos"], default="Pendientes",
                                      key="fl_rev_filtro") or "Pendientes"
        buscar = c3.text_input("Buscar", placeholder="Tercero, texto, documento…", key="fl_rev_buscar")
        sel = mov[(mov["anio"] == per[0]) & (mov["mes"] == per[1])]
        if filtro == "Pendientes":
            sel = sel[sel["origen"] == "pendiente"]
        elif filtro == "Para revisar":
            sel = sel[(sel["origen"] == "automatica") & (sel["confianza"] != "Alta")]
        if buscar:
            sel = sel[sel[["tercero", "texto", "documento", "ter_cod"]].astype(str)
                      .apply(lambda s: s.str.contains(buscar, case=False, regex=False)).any(axis=1)]
        if sel.empty:
            st.success("No hay movimientos en esta vista.", icon=":material/task_alt:")
        else:
            opciones = flujo.conceptos(leer, emp, per[0])
            opciones += sorted(set(sel["concepto"]) - set(opciones) - {""})
            base = sel.sort_values("valor", key=abs, ascending=False).head(800)
            tabla = pd.DataFrame({
                "fecha": pd.to_datetime(base["fecha"]), "valor": base["valor"], "tercero": base["tercero"],
                "texto": base["texto"], "cuenta": base["cuenta"],
                "como": base["regla"].where(base["origen"] != "manual", "manual").fillna(""),
                "concepto": base["concepto"].replace("", None), "regla": False}).set_index(base["clave"])
            if len(sel) > len(base):
                st.caption(f"Se muestran los 800 de mayor valor de {len(sel):,}.".replace(",", "."))
            ed = st.data_editor(
                tabla, hide_index=True, width="stretch", height=min(560, 38 + 35 * len(tabla)),
                key=f"fl_ed_{per[0]}_{per[1]}_{filtro}_{buscar}",
                disabled=["fecha", "valor", "tercero", "texto", "cuenta", "como"],
                column_order=["fecha", "valor", "concepto", "regla", "tercero", "texto", "como", "cuenta"],
                column_config={
                    "fecha": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY", width="small"),
                    "valor": st.column_config.NumberColumn("Valor", format="accounting"),
                    "tercero": "Tercero", "texto": st.column_config.TextColumn("Texto", width="medium"),
                    "cuenta": "Cuenta", "como": st.column_config.TextColumn("Sugerido por", width="small"),
                    "concepto": st.column_config.SelectboxColumn("Concepto de caja", options=opciones, width="medium"),
                    "regla": st.column_config.CheckboxColumn("Crear regla", help="Tercero + cuenta → este concepto, "
                                                             "para los próximos movimientos", width="small")})
            b1, b2, _ = st.columns([1, 1.3, 2])
            guardar = b1.button("Guardar cambios", type="primary", key="fl_rev_ok")
            confirmar = b2.button("Confirmar todos los que ves", key="fl_rev_conf",
                                  help="Marca como revisados (manual) los movimientos de la tabla que tienen concepto.")
            if guardar or confirmar:
                ed["concepto"] = ed["concepto"].fillna("")
                orig = tabla["concepto"].fillna("")
                cambio = (ed["concepto"] != orig) | ed["regla"]
                if confirmar:
                    cambio = cambio | (ed["concepto"] != "")
                cambio &= ed["concepto"] != ""
                filas = [(c, emp, k) for k, c in ed.loc[cambio, "concepto"].items()]
                con.executemany("UPDATE flujo_mov SET concepto=?, origen='manual', regla='manual', confianza='' "
                                "WHERE empresa=? AND clave=?", filas)
                nuevas = base.set_index("clave").loc[ed.index[ed["regla"] & (ed["concepto"] != "")]]
                for k, r in nuevas.iterrows():
                    txt = "" if r["ter_cod"] else " ".join(str(r["texto"]).upper().split()[:3])
                    con.execute("INSERT INTO flujo_reglas(empresa, tercero, cuenta, texto, concepto, creada) "
                                "VALUES(?,?,?,?,?,?)", (emp, r["ter_cod"], r["cuenta_cod"] or r["cuenta"], txt,
                                                         ed.at[k, "concepto"], ahora().strftime("%Y-%m-%d %H:%M")))
                n_regla = 0
                if len(nuevas):            # las reglas nuevas se aplican de una vez a los demás pendientes
                    pend = flujo.movimientos(leer, emp)
                    pend = pend[pend["origen"] == "pendiente"]
                    if not pend.empty:
                        r = flujo.aplicar_fijas(pend, _fijas(leer, emp))
                        ok = r != ""
                        n_regla = int(ok.sum())
                        con.executemany("UPDATE flujo_mov SET concepto=?, origen='automatica', regla='regla fija', "
                                        "confianza='Alta' WHERE empresa=? AND clave=?",
                                        [(c, emp, k) for c, k in zip(r[ok], pend.loc[ok, "clave"])])
                for a in sorted(set(mov["anio"].astype(int))):
                    flujo.recalcular(con, emp, a)
                con.commit()
                st.session_state["fl_rev_msg"] = (f"{len(filas)} movimientos guardados"
                                                  + (f" · {len(nuevas)} reglas nuevas (clasificaron {n_regla} "
                                                     "pendientes más)" if len(nuevas) else "") + ".")
                st.rerun()
        if st.session_state.get("fl_rev_msg"):
            st.success(st.session_state.pop("fl_rev_msg"))

    with st.container(key="card_adm_fl_reglas"):
        titulo_tarjeta("Reglas fijas de clasificación",
                       "Mandan sobre lo aprendido. Tercero = código de 6 cifras; Cuenta = cuenta de bancos; Texto = "
                       "palabras que contiene el texto del movimiento. Deja vacío lo que no aplique. Agrega o borra "
                       "filas y guarda.")
        reglas = leer("SELECT tercero, cuenta, texto, concepto FROM flujo_reglas WHERE empresa=?", (emp,))
        opciones = flujo.conceptos(leer, emp)
        er = st.data_editor(reglas.fillna(""), num_rows="dynamic", hide_index=True, width="stretch", key="fl_reglas_ed",
                            column_config={"tercero": "Tercero (código)", "cuenta": "Cuenta", "texto": "Texto contiene",
                                           "concepto": st.column_config.SelectboxColumn(
                                               "Concepto de caja", options=sorted(set(opciones) |
                                                                                  set(reglas["concepto"].dropna())))})
        if st.button("Guardar reglas", key="fl_reglas_ok"):
            er = er.fillna("").astype(str)
            er = er[(er["concepto"] != "") & ((er["tercero"] != "") | (er["cuenta"] != "") | (er["texto"] != ""))]
            con.execute("DELETE FROM flujo_reglas WHERE empresa=?", (emp,))
            con.anexar(er.assign(empresa=emp, creada=ahora().strftime("%Y-%m-%d %H:%M")), "flujo_reglas")
            con.commit()
            st.success(f"{len(er)} reglas guardadas. Se aplican a los próximos movimientos que cargues.")


def _flujo_ajustes(ctx):
    con, leer, emp = ctx.con, ctx.leer, ctx.empresa
    with st.container(key="card_adm_fl_aj"):
        titulo_tarjeta(f"Ajustes manuales del flujo · {emp}",
                       "Valores que no vienen de los movimientos de SAP y se suman al concepto en el mes, como en el "
                       "libro: facturación conjunta y GMF de Enel, restringido, saldo inicial de enero… Los del libro se "
                       "cargaron solos; aquí registras los de los meses nuevos. Cifras en pesos.")
        anios = leer("SELECT DISTINCT anio FROM flujo_estructura WHERE empresa=? ORDER BY anio DESC", (emp,))["anio"]
        if anios.empty:
            st.info("Primero carga el libro de flujo de caja.")
            return
        c1, c2, c3 = st.columns([0.7, 1, 2], vertical_alignment="bottom")
        anio = int(c1.selectbox("Año", anios.tolist(), key="fl_aj_anio"))
        um = leer("SELECT MAX(mes) m FROM flujo_mov WHERE empresa=? AND anio=?", (emp, anio))["m"]
        um = int(um.iloc[0]) if um.notna().all() else 1
        mes = c2.selectbox("Mes", list(range(1, 13)), index=um - 1, key="fl_aj_mes", format_func=lambda m: MESES[m - 1])
        todos = c3.toggle("Ver todos los conceptos", key="fl_aj_todos")
        est = leer("SELECT * FROM flujo_estructura WHERE empresa=? AND anio=? ORDER BY orden", (emp, anio))
        aj = leer("SELECT orden, mes, COALESCE(tipo, '') tipo, SUM(valor) valor FROM flujo_ajustes "
                  "WHERE empresa=? AND anio=? GROUP BY orden, mes, COALESCE(tipo, '')", (emp, anio))
        val = leer("SELECT orden, valor FROM flujo_lineas WHERE empresa=? AND anio=? AND mes=?", (emp, anio, mes))
        val = val.set_index("orden")["valor"]
        nombre = dict(zip(est["orden"].astype(int), est["concepto"].str.rstrip("_")))
        vinc = flujo.vinculos(est)

        # ---------------- partidas de Enel (facturador conjunto)
        fc = est[est["concepto"].map(flujo._n).str.contains("facturacion conjunta") & (est["tipo"] == "linea")]
        rem = est[est["remun"].fillna(0) != 0] if "remun" in est else est.iloc[0:0]
        if not fc.empty or not rem.empty:
            st.html('<div class="nf-ftit" style="margin-top:.4rem">Enel · facturador conjunto</div>')
            st.caption("Son manuales: Enel recauda, descuenta su costo y consigna el neto. Para que el ingreso quede "
                       "bruto, como en el libro, el costo de la facturación conjunta se muestra como egreso y se suma de "
                       "vuelta en Prestación del servicio; el GMF de la remuneración suma al ingreso y sale en Gastos "
                       "Financieros. El flujo neto no cambia: solo se presenta completo.")
            e1, e2, e3 = st.columns([1, 1, 1.3], vertical_alignment="bottom")
            fc_orden = int(fc["orden"].iloc[0]) if not fc.empty else None
            fc_act = -float(aj[(aj["orden"] == fc_orden) & (aj["mes"] == mes)]["valor"].sum()) if fc_orden is not None \
                else 0.0
            g_act = aj[(aj["mes"] == mes) & (aj["tipo"] == flujo.GMF_ENEL)]
            g_act = float(g_act[g_act["valor"] > 0]["valor"].sum()) if not g_act.empty else 0.0
            v_fc = e1.number_input(f"Costo facturación conjunta {MESES[mes - 1]}", value=fc_act, step=1_000_000.0,
                                   format="%.0f", disabled=fc_orden is None, key=f"fl_fc_{anio}_{mes}",
                                   help="En positivo; se registra como egreso")
            v_g = e2.number_input(f"GMF remuneración Enel {MESES[mes - 1]}", value=g_act, step=1_000_000.0,
                                  format="%.0f", disabled=rem.empty, key=f"fl_gmf_{anio}_{mes}",
                                  help="Columna gmf de la hoja Remuneraciones del libro")
            if fc_orden is not None and fc_orden in vinc:
                quien = " y ".join(f"{n} (se {'resta' if s < 0 else 'suma'})" for n, s in vinc[fc_orden])
                e3.caption(f"La facturación conjunta también se refleja en: {quien}.")
            if e3.button(f"Guardar partidas de Enel de {MESES[mes - 1]}", key="fl_enel_ok"):
                if fc_orden is not None:
                    con.execute("DELETE FROM flujo_ajustes WHERE empresa=? AND anio=? AND orden=? AND mes=?",
                                (emp, anio, fc_orden, mes))
                    if v_fc:
                        con.execute("INSERT INTO flujo_ajustes(empresa, anio, orden, mes, valor, tipo) "
                                    "VALUES(?,?,?,?,?,'')", (emp, anio, fc_orden, mes, -abs(float(v_fc))))
                con.execute("DELETE FROM flujo_ajustes WHERE empresa=? AND anio=? AND mes=? AND tipo=?",
                            (emp, anio, mes, flujo.GMF_ENEL))
                for r in rem.itertuples():
                    if v_g:
                        con.execute("INSERT INTO flujo_ajustes(empresa, anio, orden, mes, valor, tipo) "
                                    "VALUES(?,?,?,?,?,?)", (emp, anio, int(r.orden), mes, int(r.remun) * abs(float(v_g)),
                                                            flujo.GMF_ENEL))
                flujo.recalcular(con, emp, anio)
                con.commit()
                st.session_state["fl_aj_msg"] = f"Partidas de Enel de {MESES[mes - 1]} guardadas y flujo recalculado."
                st.rerun()
            st.html('<div class="nf-ftit" style="margin-top:.8rem">Otros ajustes</div>')

        # ---------------- demás ajustes (sin el GMF de Enel, que va arriba)
        otros = aj[aj["tipo"] != flujo.GMF_ENEL]
        con_aj = set(otros.loc[otros["valor"].abs() > 0.5, "orden"].astype(int))
        visibles = est if todos else est[est["orden"].isin(con_aj) | (est["clase"] == "manual")]
        visibles = visibles[~visibles["orden"].isin(fc["orden"])]        # la facturación conjunta va arriba
        visibles = visibles[~visibles["clase"].eq("formula") | visibles["orden"].isin(con_aj) |
                            visibles["terminos"].str.contains(", -1,", regex=False)]
        del_mes = otros[otros["mes"] == mes].groupby("orden")["valor"].sum()
        ordenes = visibles["orden"].astype(int)
        refleja = ordenes.map(lambda o: ", ".join(f"{n} (se {'resta' if s < 0 else 'suma'})" for n, s in vinc.get(o, [])))
        tabla = pd.DataFrame({"concepto": ordenes.map(nombre).values, "seccion": visibles["seccion"].values,
                              "ajuste": ordenes.map(del_mes).fillna(0.0).values,
                              "valor": ordenes.map(val).fillna(0.0).values, "refleja": refleja.values},
                             index=ordenes.values)
        ed = st.data_editor(tabla, hide_index=True, width="stretch", key=f"fl_aj_ed_{anio}_{mes}_{todos}",
                            disabled=["concepto", "seccion", "valor", "refleja"],
                            column_config={"concepto": "Concepto", "seccion": "Sección",
                                           "ajuste": st.column_config.NumberColumn(f"Ajuste {MESES[mes - 1]}",
                                                                                   format="accounting"),
                                           "valor": st.column_config.NumberColumn("Valor del concepto (con ajuste)",
                                                                                  format="accounting"),
                                           "refleja": st.column_config.TextColumn(
                                               "También mueve", help="Conceptos cuya fórmula suma o resta este valor")})
        st.caption("El saldo inicial de enero va como ajuste (no viene de SAP); los demás meses toman el saldo final del "
                   "mes anterior y su ajuste es cero.")
        if st.button(f"Guardar ajustes de {MESES[mes - 1]}", type="primary", key="fl_aj_ok"):
            cambia = ed.index[(ed["ajuste"].fillna(0) - tabla["ajuste"]).abs() > 0.004]
            for o in cambia:
                con.execute("DELETE FROM flujo_ajustes WHERE empresa=? AND anio=? AND orden=? AND mes=? "
                            "AND COALESCE(tipo, '')<>?", (emp, anio, int(o), mes, flujo.GMF_ENEL))
                v = float(ed.at[o, "ajuste"] or 0)
                if v:
                    con.execute("INSERT INTO flujo_ajustes(empresa, anio, orden, mes, valor, tipo) VALUES(?,?,?,?,?,'')",
                                (emp, anio, int(o), mes, v))
            flujo.recalcular(con, emp, anio)
            con.commit()
            st.session_state["fl_aj_msg"] = f"{len(cambia)} ajustes guardados y flujo recalculado."
            st.rerun()
        if st.session_state.get("fl_aj_msg"):
            st.success(st.session_state.pop("fl_aj_msg"))


def _flujo_libro(ctx):
    con, leer, usuario, emp = ctx.con, ctx.leer, ctx.usuario, ctx.empresa
    with st.container(key="card_adm_fl_libro"):
        titulo_tarjeta(f"Libro de flujo de caja (carga inicial) · {emp}",
                       "Una sola vez por año: el libro completo (ej. ALDC_Flujo de Caja 2026.xlsx). Se toman la "
                       "estructura y fórmulas de la hoja 'Flujo de caja', los movimientos ya clasificados de la hoja "
                       "Journal / SAP (con ellos aprende la app), los ajustes manuales, el proyectado 4+8, los terceros y "
                       "el subrubro por tercero de la hoja Param. Reemplaza todo lo del mismo año de esta compañía.")
        f = st.file_uploader("Libro de flujo de caja", type=["xlsx", "xlsm"], key="up_fl_libro",
                             label_visibility="collapsed")
        if f:
            try:
                with st.spinner("Leyendo el libro (puede tardar unos segundos)…"):
                    L = _leer_flujo("libro", f.getvalue())
            except Exception as err:
                st.error(f"No pude leer el libro: {err}")
                L = None
            if L is not None:
                _aviso_empresa(ctx, f.name)
                est, vals, mov = L["estructura"], L["valores"], L["movimientos"].copy()
                mov["concepto"] = mov["sub"]
                anio = int(mov["anio"].mode().iloc[0]) if not mov.empty else ahora().year
                anio = int(st.number_input("Año del flujo", value=anio, step=1, format="%d", key="fl_lib_anio"))
                mov = mov[mov["anio"] == anio]
                aj = flujo.ajustes_del_archivo(est, vals, mov)
                calc = flujo.calcular(est, flujo.sumas_de(mov), {(o, m): v for o, m, v in aj.itertuples(index=False)})
                aj = flujo.partir_ajustes(est, aj, L.get("gmf", {}))       # el GMF de Enel queda aparte
                dif = calc.merge(vals, on=["orden", "mes"], suffixes=("", "_a"))
                dif = float((dif["valor"] - dif["valor_a"]).abs().max())
                sin = (mov["sub"] == "").sum()
                meses_m = sorted(mov["mes"].unique())
                st.html(E.kpis([
                    dict(label="Conceptos", valor=str(int((est["clase"] == "sap").sum())),
                         sub=f"{len(est)} filas en la hoja Flujo de caja", icono="receipt", color=E.AZUL,
                         fondo=E.AZUL_SUAVE),
                    dict(label="Movimientos clasificados", valor=f"{len(mov) - sin:,}".replace(",", "."),
                         sub=(f"{MESES[meses_m[0] - 1]}–{MESES[meses_m[-1] - 1]} {anio}" if meses_m else "")
                         + (f" · {sin} sin subrubro" if sin else ""), icono="check", color=E.VERDE, fondo=E.VERDE_SUAVE),
                    dict(label="Ajustes manuales", valor=str(aj["orden"].nunique()),
                         sub=f"{len(aj)} valores por mes que no vienen de SAP", icono="calendar", color=E.AZUL,
                         fondo=E.AZUL_SUAVE),
                    dict(label="Cuadre con el libro", valor="OK" if dif < 1 else E.millones(dif),
                         sub=f"{len(L['terceros']):,} terceros en el maestro".replace(",", "."), icono="wallet",
                         color=E.VERDE if dif < 1 else E.ROJO, fondo=E.VERDE_SUAVE if dif < 1 else E.ROJO_SUAVE),
                ]))
                prep = flujo.preparar(mov, None, validos=est.loc[est["sap"] != 0, "concepto"].tolist() + [flujo.ANULADO])
                bt = flujo.backtest(prep)
                if not bt.empty:
                    peso = bt["movimientos"]
                    st.caption(f"Prueba con tus datos: aprendiendo de los meses anteriores, la app habría clasificado "
                               f"sola el {(bt['cubiertos'] * peso).sum() / peso.sum():.0%} de los movimientos de cada "
                               f"mes, acertando el {(bt['acierto_valor'] * peso).sum() / peso.sum():.0%} del valor.")
                if st.button(f"Cargar libro de flujo {anio}", type="primary", key="fl_lib_ok"):
                    for t in ("flujo_estructura", "flujo_ajustes", "flujo_mov", "flujo_proyectado"):
                        con.execute(f"DELETE FROM {t} WHERE empresa=? AND anio=?", (emp, anio))
                    con.anexar(est.assign(empresa=emp, anio=anio), "flujo_estructura")
                    con.anexar(aj.assign(empresa=emp, anio=anio), "flujo_ajustes")
                    proy = L["proyectado"].assign(empresa=emp, anio=anio)
                    con.anexar(proy[["anio", "concepto", "proyectado", "observacion", "empresa"]], "flujo_proyectado")
                    ter, cu, pt = flujo.terceros_de(L["terceros"], emp), L["cuentas"], L["por_tercero"]
                    # un libro de un año anterior al último cargado no pisa terceros, cuentas ni el subrubro por
                    # tercero (los del libro más reciente mandan); solo agrega lo que falte
                    ultimo = leer("SELECT MAX(anio) a FROM flujo_estructura WHERE empresa=?", (emp,))["a"]
                    viejo = ultimo.notna().all() and anio < int(ultimo.iloc[0])
                    if viejo:
                        hay_t = set(leer("SELECT codigo FROM flujo_terceros WHERE empresa=?", (emp,))["codigo"])
                        hay_c = set(leer("SELECT cuenta FROM flujo_cuentas WHERE empresa=?", (emp,))["cuenta"])
                        ter, cu = ter[~ter["codigo"].isin(hay_t)], cu[~cu["cuenta"].isin(hay_c)]
                        pt = pt.iloc[0:0]
                    if not ter.empty:
                        _guardar_terceros(con, emp, ter)
                    if not cu.empty:
                        if not viejo:
                            con.execute("DELETE FROM flujo_cuentas WHERE empresa=?", (emp,))
                        con.anexar(cu.assign(empresa=emp), "flujo_cuentas")
                    if not pt.empty:
                        con.execute("DELETE FROM flujo_por_tercero WHERE empresa=?", (emp,))
                        con.anexar(pt.assign(empresa=emp), "flujo_por_tercero")
                    cu = leer("SELECT cuenta, nombre FROM flujo_cuentas WHERE empresa=?", (emp,))
                    nombres = flujo.nombres_tercero(leer, emp)
                    mov["tercero"] = mov["ter_cod"].map(nombres).fillna(mov["ref3"])
                    m = prep.assign(origen=prep["origen"].where(prep["concepto"] != "", "pendiente"))
                    con.anexar(flujo.a_tabla(m, emp, nombres, dict(zip(cu["cuenta"], cu["nombre"]))), "flujo_mov")
                    flujo.recalcular(con, emp, anio)
                    con.commit()
                    ctx.registrar_carga(usuario, "Flujo de caja (libro)", f.name,
                                        f"{emp} · {anio} · {len(est)} filas · {len(mov)} movimientos · {len(aj)} ajustes")
                    st.success("Libro de flujo cargado.")
                    st.rerun()


def _guardar_terceros(con, emp, ter: pd.DataFrame):
    con.executemany("DELETE FROM flujo_terceros WHERE empresa=? AND codigo=?", [(emp, c) for c in ter["codigo"]])
    con.anexar(ter[["codigo", "nombre", "nit"]].assign(empresa=emp), "flujo_terceros")
    con.executemany("UPDATE flujo_mov SET tercero=? WHERE empresa=? AND ter_cod=?",
                    [(n, emp, c) for c, n in zip(ter["codigo"], ter["nombre"])])


def _flujo_terceros(ctx):
    con, leer, usuario, emp = ctx.con, ctx.leer, ctx.usuario, ctx.empresa
    with st.container(key="card_adm_fl_ter"):
        titulo_tarjeta(f"Maestro de terceros · {emp}",
                       "Código de 6 cifras (columna K del reporte, 'Texto cab.documento') → nombre del tercero. Sube la "
                       "base de terceros de SAP (columnas Acreedor, Nombre 1 y Soc.): los de DB01 quedan para ALDC y los "
                       "de DB02 para ALSM, las dos compañías de una vez. Los códigos que ya estaban se actualizan y el "
                       "nombre se corrige también en los movimientos cargados.")
        n = leer("SELECT COUNT(*) n FROM flujo_terceros WHERE empresa=?", (emp,))["n"].iloc[0]
        sin = leer("SELECT ter_cod, MAX(tercero) tercero, COUNT(*) n, SUM(valor) valor FROM flujo_mov "
                   "WHERE empresa=? AND ter_cod<>'' AND ter_cod NOT IN (SELECT codigo FROM flujo_terceros "
                   "WHERE empresa=?) GROUP BY ter_cod ORDER BY n DESC", (emp, emp))
        st.caption(f"{int(n):,} terceros en el maestro · {len(sin)} códigos de los movimientos sin nombre en el maestro"
                   .replace(",", ".") + (" (se muestra el nombre abreviado de SAP)." if len(sin) else "."))
        if len(sin):
            with st.expander("Ver códigos sin nombre"):
                st.dataframe(sin, hide_index=True, width="stretch",
                             column_config={"ter_cod": "Código", "tercero": "Nombre en SAP (abreviado)",
                                            "n": "Movimientos",
                                            "valor": st.column_config.NumberColumn("Valor", format="accounting")})
        f = st.file_uploader("Maestro de terceros", type=["xlsx", "xlsm", "xls"], key="up_fl_ter",
                             label_visibility="collapsed")
        if f:
            try:
                ter = _leer_flujo("terceros", f.getvalue())
            except Exception as err:
                st.error(f"No pude leer el archivo: {err}")
                return
            if ter.empty:
                st.warning("No encontré una hoja con las columnas de código (Acreedor) y nombre (Nombre 1).")
                return
            por_soc = (ter["sociedad"] != "").any()
            destinos = {e: flujo.terceros_de(ter, e) for e in (flujo.SOCIEDADES.values() if por_soc else [emp])}
            destinos = {e: t for e, t in destinos.items() if not t.empty}
            st.write(" · ".join(f"**{e}**: {len(t):,} terceros".replace(",", ".") for e, t in destinos.items())
                     + ("" if por_soc else " (el archivo no trae la columna Soc.: se guardan en la compañía elegida)"))
            otras = sorted(set(ter["sociedad"]) - set(flujo.SOCIEDADES) - {""})
            if otras:
                st.caption("Sociedades que no son ALDC ni ALSM (no se cargan): " + ", ".join(otras))
            if st.button("Guardar terceros", type="primary", key="fl_ter_ok"):
                for e, t in destinos.items():
                    _guardar_terceros(con, e, t)
                con.commit()
                ctx.registrar_carga(usuario, "Terceros flujo", f.name,
                                    " · ".join(f"{e} {len(t)}" for e, t in destinos.items()))
                st.session_state["fl_ter_msg"] = "Terceros guardados: " + ", ".join(
                    f"{e} ({len(t):,})".replace(",", ".") for e, t in destinos.items()) + "."
                st.rerun()
        if st.session_state.get("fl_ter_msg"):
            st.success(st.session_state.pop("fl_ter_msg"))


def _cargar_flujo(ctx):
    s1, s2, s3, s4 = st.tabs(["Movimientos del mes (SAP)", "Revisar y clasificar", "Ajustes manuales",
                              "Libro inicial y terceros"], on_change="rerun", key="adm_fl")
    with s1:
        if s1.open:
            _flujo_mes_sap(ctx)
    with s2:
        if s2.open:
            _flujo_revisar(ctx)
    with s3:
        if s3.open:
            _flujo_ajustes(ctx)
    with s4:
        if s4.open:
            _flujo_libro(ctx)
            _flujo_terceros(ctx)


# ------------------------------------------------------------------ cartera de usuarios
def _cargar_cartera(ctx):
    con, leer, usuario, emp = ctx.con, ctx.leer, ctx.usuario, ctx.empresa
    meses = leer("SELECT periodo, archivo, filas, usuarios, usuarios_propio, propio, terceros, ajustes, total, fecha, "
                 "usuario FROM cartera_meses WHERE empresa=? ORDER BY periodo", (emp,))
    with st.container(key="card_adm_cart1"):
        titulo_tarjeta(f"Base de cartera del mes · {emp}",
                       "Sube la base «Reporte Cartera AñoMesTarifa» del sistema comercial: el .txt (separado por |), "
                       "el Excel (hoja Base, como CARTERA ACTUAL.xlsx) o el archivo comprimido (.bz2 o .zip, que pesa "
                       "mucho menos). El mes sale de la columna FECHA CORTE. Si el mes ya estaba, se reemplaza. "
                       "Se guarda el saldo por cuenta contrato y tipo de cartera, y los saldos por edad y componente.")
        if not meses.empty:
            u = meses.iloc[-1]
            st.caption(f"Último mes cargado: {cartera.nombre_mes(int(u['periodo']), True)} · {len(meses)} meses en "
                       "total.")
        f = st.file_uploader("Base de cartera", type=["txt", "csv", "xlsx", "xlsm", "zip", "bz2", "gz"],
                             key="up_cartera", label_visibility="collapsed")
        st.caption("Si el archivo sale en rojo (muy grande para el servidor), comprímelo: clic derecho sobre el .txt → "
                   "Enviar a → Carpeta comprimida (en zip). El .zip pesa unas 7 veces menos y se carga igual.")
        if f:
            clave = f"cart_{f.name}_{f.size}"
            if st.session_state.get("cart_leida", ("",))[0] != clave:
                excel = f.name.lower().endswith((".xlsx", ".xlsm"))
                if st.button("Leer la base", type="primary", icon=":material/upload_file:", key="cart_leer"):
                    barra = st.progress(0.0, text="Leyendo la base…")
                    comprimido = f.name.lower().endswith((".zip", ".bz2", ".gz"))
                    estimado = max(f.size * (11 if comprimido else 1) / 250, 1)   # filas aproximadas según el peso
                    def avance(n):
                        barra.progress(min(n / estimado, 0.99), text=f"Leyendo la base… {n:,} filas".replace(",", "."))
                    try:
                        res = cartera.comprimir(cartera.leer(f.getvalue(), f.name, avance=avance))
                    except Exception as err:
                        barra.empty()
                        st.error(f"No pude leer el archivo: {err}")
                        return
                    barra.empty()
                    st.session_state["cart_leida"] = (clave, res)
                    st.rerun()
                if excel:
                    st.caption("Un Excel de un millón de filas tarda unos 4 minutos en leerse; el .txt o el .bz2 "
                               "tardan menos de un minuto.")
                return
            res = st.session_state["cart_leida"][1]
            r = cartera.resumen(res)
            per = res["periodo"]
            for a in res["avisos"]:
                st.warning(a, icon=":material/warning:")
            st.html(E.kpis([
                dict(label="Mes de la cartera", valor=cartera.nombre_mes(per, True).capitalize(),
                     sub=f"{res['filas']:,} filas de la base".replace(",", "."), icono="calendar", color=E.AZUL,
                     fondo=E.AZUL_SUAVE),
                dict(label="Cartera propia", valor=E.millones(r["propio"]),
                     sub=f"{r['usuarios_propio']:,} usuarios con saldo".replace(",", "."), icono="wallet",
                     color=E.AZUL, fondo=E.AZUL_SUAVE),
                dict(label="Componentes de terceros", valor=E.millones(r["terceros"]),
                     sub="Aprovechamiento, disposición, lixiviados, VIAT, CCS aprov.", icono="receipt",
                     color=E.AMBAR, fondo=E.AMBAR_SUAVE),
                dict(label="Total cartera (todos)", valor=E.millones(r["total"]),
                     sub=f"Ajustes y saldos a favor: {E.millones(r['ajustes'])}", icono="check", color=E.VERDE,
                     fondo=E.VERDE_SUAVE),
            ]))
            ya = per in set(meses["periodo"].astype(int))
            if ya:
                st.warning(f"{cartera.nombre_mes(per, True).capitalize()} ya estaba cargado: se reemplaza.",
                           icon=":material/warning:")
            c1, c2, _ = st.columns([1.2, 0.8, 2])
            if c1.button(f"Guardar cartera de {cartera.nombre_mes(per, True)}", type="primary", key="cart_ok"):
                with st.spinner("Guardando…"):
                    cartera.guardar(con, emp, res, archivo=f.name, usuario=usuario,
                                    fecha=ahora().strftime("%Y-%m-%d %H:%M"))
                    ctx.registrar_carga(usuario, "Cartera", f.name,
                                        f"{emp} · {cartera.nombre_mes(per)} · {r['usuarios']:,} usuarios · "
                                        f"propia {E.millones(r['propio'])}".replace(",", "."))
                st.session_state.pop("cart_leida", None)
                st.cache_data.clear()
                st.success(f"Cartera de {cartera.nombre_mes(per, True)} guardada.", icon=":material/check:")
            if c2.button("Descartar", key="cart_no"):
                st.session_state.pop("cart_leida", None)
                st.rerun()

    if not meses.empty:
        with st.container(key="card_adm_cart2"):
            titulo_tarjeta("Meses cargados")
            t = pd.DataFrame({"Mes": [cartera.nombre_mes(int(p), True).capitalize() for p in meses["periodo"]],
                              "Usuarios con saldo propio": meses["usuarios_propio"].astype(int),
                              "Cartera propia": meses["propio"], "Terceros": meses["terceros"],
                              "Ajustes y saldos a favor": meses["ajustes"], "Total": meses["total"],
                              "Archivo": meses["archivo"], "Cargado": meses["fecha"] + " · " + meses["usuario"]})
            num = ["Usuarios con saldo propio", "Cartera propia", "Terceros", "Ajustes y saldos a favor", "Total"]
            t[num] = t[num].round(0)
            st.dataframe(t, hide_index=True, width="stretch", column_config={
                c: st.column_config.NumberColumn(format="localized") for c in num})
            c1, c2, _ = st.columns([1, 0.8, 2], vertical_alignment="bottom")
            quitar = c1.selectbox("Borrar un mes", [None] + [int(p) for p in meses["periodo"]], key="cart_borrar",
                                  format_func=lambda p: "—" if p is None else cartera.nombre_mes(p, True))
            if quitar and c2.button("Borrar", key="cart_borrar_ok", icon=":material/delete:"):
                cartera.borrar(con, emp, quitar)
                ctx.registrar_carga(usuario, "Cartera", "", f"{emp} · borrado {cartera.nombre_mes(quitar)}")
                st.cache_data.clear()
                st.rerun()


def _sispo_validar(contenido: bytes):
    """Lee el SISPO con el lector único y devuelve (análisis, error). Exige lo que necesitan Mantenimiento y Operaciones."""
    try:
        d = L.sispo_leer(io.BytesIO(contenido))
        sispo_oper.leer(io.BytesIO(contenido))                      # Operaciones exige además Grupo, IdKardex y Codigo Articulo
    except Exception as e:
        return None, str(e)
    lin, info = L.sispo_filtrar(d, L.SISPO_CUENTAS)
    return dict(d=d, info=info, periodos=L.sispo_periodos(lin), empresa=L.sispo_empresa_probable(lin) if len(lin) else None), None


def _cargar_sispo(ctx):
    """SISPO del mes: se carga y se guarda UNA vez aquí; Mantenimiento y Operaciones lo toman de lo guardado."""
    with st.container(key="card_adm_sispo"):
        titulo_tarjeta("SISPO del mes", "Un solo archivo alimenta Mantenimiento y Operaciones. Cargar otro del mismo mes y compañía lo reemplaza.")
        g = sispo_fuente.listar(ctx.leer)
        if g.empty:
            st.caption("Todavía no hay SISPO guardado.")
        else:
            st.dataframe(g.assign(tamano=(g["tamano"] / 1024 / 1024).round(2)), hide_index=True, width="stretch", column_config={
                "empresa": "Compañía", "periodo": st.column_config.NumberColumn("Periodo", format="%d"), "archivo": "Archivo",
                "usuario": "Usuario", "fecha": "Cargado", "tamano": st.column_config.NumberColumn("MB", format="%.2f")})
        f = st.file_uploader("SISPO del mes (.xlsx)", type=["xlsx"], key="adm_sispo_up")
        if f:
            contenido = f.getvalue()
            llave = (f.name, len(contenido))
            memo = st.session_state.get("_adm_sispo_memo")
            if memo and memo[0] == llave:
                a, error = memo[1]
            else:
                with st.spinner("Leyendo el SISPO…"):
                    a, error = _sispo_validar(contenido)
                st.session_state["_adm_sispo_memo"] = (llave, (a, error))
            if error:
                st.error(f"No pude leer el SISPO: {error}")
            elif a["info"]["en_cuenta"] == 0:
                st.error("El SISPO no trae movimientos de las cuentas que usan Mantenimiento y Operaciones.")
            else:
                c1, c2, c3 = st.columns(3)
                inferida = a["empresa"]
                empresa = c1.selectbox("Compañía", list(EMPRESAS), index=list(EMPRESAS).index(inferida) if inferida else 0, key="adm_sispo_emp",
                                       help="Se propone según «EmpresaAgencia» del archivo.")
                bloqueos = []
                if inferida and inferida != empresa:
                    bloqueos.append(f"El archivo parece de {inferida}, pero elegiste {empresa}.")
                periodo = None
                if len(a["periodos"]) != 1:
                    bloqueos.append("El SISPO debe traer un solo periodo de cierre; trae: " + (", ".join(map(str, a["periodos"])) or "ninguno") + ".")
                else:
                    periodo = int(a["periodos"][0])
                    c2.metric("Periodo (cierre)", f"{MESES_L[periodo % 100 - 1].capitalize()} {periodo // 100}")
                c3.metric("Registros", f"{a['info']['total']:,}".replace(",", "."))
                for b in bloqueos:
                    st.error(b, icon=":material/block:")
                existe = periodo is not None and not g[(g["empresa"] == empresa) & (g["periodo"] == periodo)].empty
                if st.button("Reemplazar el guardado" if existe else "Guardar SISPO", type="primary", icon=":material/upload:",
                             disabled=bool(bloqueos), key="adm_sispo_guardar"):
                    sispo_fuente.guardar(ctx.con, empresa, periodo, f.name, ctx.usuario, ahora().strftime("%Y-%m-%d %H:%M"), contenido)
                    ctx.registrar_carga(ctx.usuario, "SISPO", f.name, f"{empresa} {periodo} · {a['info']['total']} registros")
                    st.session_state.pop("_adm_sispo_memo", None)
                    st.success("SISPO guardado. Mantenimiento y Operaciones lo toman de aquí.")
                    st.rerun()
        if not g.empty:
            opciones = [f"{r.empresa} · {r.periodo}" for r in g.itertuples()]
            c_sel, c_ok, c_btn = st.columns([2, 2, 1], vertical_alignment="bottom")
            elegido = c_sel.selectbox("Archivo a quitar", opciones, key="adm_sispo_quitar")
            ok = c_ok.checkbox("Confirmo quitar este archivo guardado", key="adm_sispo_quitar_ok",
                               help="Solo borra el archivo guardado; lo ya aplicado en Mantenimiento y Operaciones no se toca.")
            if c_btn.button("Quitar", icon=":material/delete:", disabled=not ok, key="adm_sispo_quitar_btn"):
                emp, per_ = elegido.split(" · ")
                sispo_fuente.borrar(ctx.con, emp, int(per_))
                st.rerun()


def _cargar_mantenimiento_operaciones(ctx):
    """Todas las cargas de Mantenimiento y Operaciones en un solo lugar."""
    s1, s2, s3 = st.tabs(["SISPO del mes", "Base maestra", "Presupuestos y datos base"], on_change="rerun", key="adm_mo")
    with s1:
        if s1.open:
            _cargar_sispo(ctx)
    with s2:
        if s2.open:
            with st.container(key="card_adm_maestra"):
                titulo_tarjeta("Base maestra de referencias y conceptos", "Mantenimiento la usa para asignar concepto a cada referencia del SISPO.")
                mod_mantenimiento.admin_base_maestra(ctx)
    with s3:
        if s3.open:
            with st.container(key="card_adm_mo_datos"):
                titulo_tarjeta("Mantenimiento · placas, presupuesto de flota, histórico y Tecnomecánica/SOAT")
                mod_mantenimiento.admin_datos_base(ctx)
            with st.container(key="card_adm_op_ppto"):
                titulo_tarjeta("Operaciones · libro «3.1 operaciones»", "Se carga por compañía.")
                for tab, e in zip(st.tabs([f"{e} · {EMPRESAS[e]}" for e in EMPRESAS]), EMPRESAS):
                    with tab:
                        operaciones.admin_presupuesto(ctx, e)


def mostrar(ctx):
    con, leer, usuario, registrar_carga = ctx.con, ctx.leer, ctx.usuario, ctx.registrar_carga
    anio = (leer("SELECT MAX(anio) a FROM presupuesto")["a"][0] or datetime.now().year)
    st.html(E.encabezado("Administración", "Carga de archivos, asignación por gerencia, historial y respaldo", []))
    t1, t2, t6, t8, t3, t9, t10, t5, t4 = st.tabs(["Presupuesto", "Flujo de caja", "Estados financieros", "Cartera",
                                              "Combustible", "Mantenimiento y Operaciones", "Corte y poda", "Nómina",
                                              "Historial y respaldo"], on_change="rerun", key="adm_tabs")
    with t6:
        if t6.open:   # solo se calcula la pestaña abierta
            _empresa_adm(ctx)
            _cargar_eeff(ctx)
    with t8:
        if t8.open:   # solo se calcula la pestaña abierta
            _empresa_adm(ctx)
            _cargar_cartera(ctx)
    with t5:
        if t5.open:   # solo se calcula la pestaña abierta
            n1, n2 = st.tabs(["Cargar archivos", "Homologar cargos"], on_change="rerun", key="adm_nom")
            with n1:
                if n1.open:
                    _cargar_nomina(ctx)
            with n2:
                if n2.open:
                    _homologar_cargos(ctx)


    with t1:
        if t1.open:   # solo se calcula la pestaña abierta
            _empresa_adm(ctx)
            sub1, sub2, sub3, sub4, sub5 = st.tabs(["Presupuestos", "Ejecución SAP", "Revisar y reclasificar",
                                                   "Homologar nombres", "Asignación por gerencia"], on_change="rerun", key="adm_sub")
            with sub1:
                if sub1.open:   # solo se calcula la pestaña abierta
                    _cargar_presupuestos(ctx)
                    _editar_distribucion(ctx)
            with sub2:
                if sub2.open:   # solo se calcula la pestaña abierta
                    _cargar_ejecucion(ctx)
                    _herramientas_ejecucion(ctx)
            with sub3:
                if sub3.open:   # solo se calcula la pestaña abierta
                    _revisar_clasificacion(ctx)
            with sub4:
                if sub4.open:   # solo se calcula la pestaña abierta
                    _homologar(ctx)
            with sub5:
                if sub5.open:   # solo se calcula la pestaña abierta
                    _asignar(ctx, anio)

    with t2:
        if t2.open:   # solo se calcula la pestaña abierta
            _empresa_adm(ctx)
            _cargar_flujo(ctx)

    with t3:
        if t3.open:   # solo se calcula la pestaña abierta
            _cargar_combustible(ctx)

    with t9:
        if t9.open:   # solo se calcula la pestaña abierta
            _cargar_mantenimiento_operaciones(ctx)

    with t10:
        if t10.open:   # solo se calcula la pestaña abierta
            mod_corte_poda.admin_cargas(ctx)

    with t4:
        if t4.open:   # solo se calcula la pestaña abierta
            with st.container(key="card_adm3"):
                titulo_tarjeta("3. Historial de cargas y respaldo")
                st.dataframe(leer("SELECT fecha Fecha, usuario Usuario, tipo Tipo, archivo Archivo, detalle Detalle "
                                  "FROM cargas ORDER BY rowid DESC LIMIT 20"), hide_index=True, width="stretch")
                donde = "la nube (PostgreSQL)" if con.motor == "postgres" else "el archivo presupuesto.db de este equipo"
                st.caption(f"Los datos están guardados en {donde}.")
                if st.button("Preparar respaldo", icon=":material/backup:", key="resp_prep"):
                    st.session_state["respaldo"] = con.exportar_sqlite()
                if st.session_state.get("respaldo"):
                    st.download_button("Descargar respaldo (.db)", st.session_state["respaldo"],
                                       f"respaldo_presupuesto_{ahora():%Y%m%d_%H%M}.db", icon=":material/download:",
                                       key="resp_dl")

            with st.container(key="card_adm_rest"):
                titulo_tarjeta("Restaurar un respaldo",
                               "Sube un archivo .db (un respaldo descargado aquí, o el presupuesto.db de tu equipo) para "
                               "reemplazar TODA la información actual. Sirve para pasar los datos del equipo a la nube.")
                fr = st.file_uploader("Respaldo .db", type=["db"], key="resp_up", label_visibility="collapsed")
                if fr:
                    ok = st.checkbox("Entiendo que se reemplazan todos los datos actuales por los del respaldo",
                                     key="resp_ok")
                    if st.button("Restaurar", type="primary", disabled=not ok, key="resp_btn"):
                        with st.spinner("Restaurando…"):
                            res = con.importar_sqlite(fr.getvalue())
                        registrar_carga(usuario, "Restauración", fr.name,
                                        " · ".join(f"{t} {n}" for t, n in res.items() if n))
                        st.success("Respaldo restaurado: " + ", ".join(f"{t} ({n:,})".replace(",", ".")
                                                                         for t, n in res.items() if n))
