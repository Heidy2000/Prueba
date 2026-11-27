"""Administración de Mantenimiento, Operaciones y Dotación 2 (desarrollos de Heidy, traídos del panel de Streamlit).
"""
import io

import pandas as pd
import streamlit as st

import estilo as E
import logica_m as L
import mod_mantenimiento
import operaciones
import sispo_fuente
import sispo_oper
from ui import EMPRESAS, MESES_L, ahora, titulo_tarjeta


def cargas(ctx):
    """Pestaña «Mantenimiento y Operaciones» de Administración."""
    s0, s4 = st.tabs(["Cargas", "Estado de ejecuciones"], on_change="rerun", key="adm_mo_vista")
    with s0:
        if s0.open:
            _cargar_mantenimiento_operaciones(ctx)
    with s4:
        if s4.open:
            _ejecuciones(ctx)


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
        titulo_tarjeta("SISPO del mes", "Un solo archivo alimenta Mantenimiento y Operaciones, y se aplica en ambos al guardarlo. Cargar otro del mismo mes y compañía lo reemplaza.")
        r = st.session_state.pop("_adm_sispo_aplic", None)          # resultado de la aplicación automática del último guardado
        if r is not None:
            st.success("SISPO guardado." + (" Aplicado en: " + "; ".join(r["ok"]) + "." if r["ok"] else ""), icon=":material/check_circle:")
            for m in r["avisos"]:
                st.warning(m, icon=":material/warning:")
            for m in r["errores"]:
                st.error(m, icon=":material/error:")
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
                    with st.spinner("Aplicando en Mantenimiento y Operaciones…"):          # sin pasos manuales: se concilia y se aplica al guardar
                        st.session_state["_adm_sispo_aplic"] = mod_mantenimiento.aplicar_sispo_guardado(ctx, empresa, periodo)
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
    s1, s2, s3, s4, s5 = st.tabs(["SISPO del mes", "Base maestra", "Presupuestos y datos base", "Mantenimiento: reclasificación", "Operaciones: ítems y hojas"], on_change="rerun", key="adm_mo")
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
    with s4:
        if s4.open:
            st.caption("Reclasificación de movimientos pendientes, reglas y descargas de control de Mantenimiento (se calcula al abrir esta pestaña).")
            mod_mantenimiento.admin_revision_tecnica(ctx)
    with s5:
        if s5.open:
            st.caption("Equivalencia referencia → ítem del presupuesto y asignación de sub-rubros a hojas, con el resumen técnico de Operaciones.")
            operaciones.admin_revision_tecnica(ctx)


def _ejecuciones(ctx):
    """Vista central (solo lectura) de las ejecuciones que ya están cargadas y procesadas: SAP por rubro y mes, y el estado de cada periodo
    (SISPO guardado, lo aplicado en Mantenimiento y en Operaciones). Todo sale de las tablas existentes con consultas agregadas; no carga
    ni recalcula nada, así que no puede diferir de lo que muestran los módulos."""
    leer = ctx.leer
    mod_mantenimiento._asegurar_tablas(ctx)
    operaciones._asegurar_tablas(ctx)
    anios = leer("SELECT DISTINCT anio FROM ejecucion WHERE anio IS NOT NULL ORDER BY anio DESC")["anio"].astype(int).tolist()
    if not anios:
        st.info("Aún no hay ejecución cargada. Se carga en Administración → Presupuesto → Ejecución SAP.", icon=":material/info:")
        return
    c1, c2 = st.columns([1, 3], vertical_alignment="bottom")
    anio = c1.selectbox("Año", anios, key="ej_anio")
    emp_sel = c2.segmented_control("Compañía", ["Todas", *EMPRESAS], default="Todas", key="ej_empresa",
                                   format_func=lambda x: x if x == "Todas" else f"{x} · {EMPRESAS[x]}") or "Todas"
    empresas = list(EMPRESAS) if emp_sel == "Todas" else [emp_sel]
    ph = ",".join("?" * len(empresas))
    sap = leer(f"SELECT empresa, mes, codigo, rubro, COUNT(*) AS lineas, SUM(valor) AS valor FROM ejecucion "
               f"WHERE anio=? AND empresa IN ({ph}) GROUP BY empresa, mes, codigo, rubro", (anio, *empresas))
    if sap.empty:
        st.info(f"No hay ejecución de SAP de {emp_sel} en {anio}.", icon=":material/info:")
        return
    sap["mes"] = pd.to_numeric(sap["mes"], errors="coerce").fillna(0).astype(int)
    ult = int(sap.loc[sap["valor"].abs() > 0, "mes"].max()) if (sap["valor"].abs() > 0).any() else int(sap["mes"].max())
    lo, hi = anio * 100, anio * 100 + 99
    try:
        archivos = sispo_fuente.listar(leer)
    except Exception:
        archivos = pd.DataFrame(columns=["empresa", "periodo", "archivo", "usuario", "fecha", "tamano"])
    archivos = archivos[archivos["empresa"].isin(empresas) & archivos["periodo"].between(lo, hi)]
    mt = leer(f"SELECT empresa, periodo, SUM(n_lineas) AS lineas, SUM(valor_sispo) AS valor, MAX(estado) AS estado FROM mtto_sispo_periodo "
              f"WHERE periodo BETWEEN ? AND ? AND empresa IN ({ph}) GROUP BY empresa, periodo", (lo, hi, *empresas))
    op = leer(f"SELECT p.empresa, p.periodo, p.n_lineas AS lineas, COALESCE(v.valor, 0) AS valor FROM oper_sispo_periodo p LEFT JOIN "
              f"(SELECT empresa, periodo, SUM(valor) AS valor FROM oper_sispo WHERE hoja <> ? GROUP BY empresa, periodo) v "
              f"ON v.empresa = p.empresa AND v.periodo = p.periodo WHERE p.periodo BETWEEN ? AND ? AND p.empresa IN ({ph})",
              (sispo_oper.HOJA_PAPELERIA, lo, hi, *empresas))

    st.html(E.kpis([
        dict(label="Líneas de SAP cargadas", valor=f"{int(sap['lineas'].sum()):,}".replace(",", "."), sub=f"{emp_sel} · {anio}", icono="receipt",
             color=E.AZUL, fondo=E.AZUL_SUAVE),
        dict(label="Ejecución SAP (neta)", valor=E.millones(sap["valor"].sum()), sub="Todos los rubros cargados", icono="wallet", color=E.AZUL, fondo=E.AZUL_SUAVE),
        dict(label="Último mes con ejecución", valor=MESES_L[ult - 1].capitalize() if 1 <= ult <= 12 else "—", sub=f"{sap['mes'].nunique()} meses con datos",
             icono="calendar", color=E.AZUL, fondo=E.AZUL_SUAVE),
        dict(label="SISPO guardados", valor=str(len(archivos)), sub="Archivos del año (compañía × mes)", icono="check", color=E.AZUL, fondo=E.AZUL_SUAVE)]))

    with st.container(key="card_ej_rubro"):
        titulo_tarjeta("Ejecución de SAP por rubro y mes", f"{emp_sel} · {anio} · valores netos; el detalle de cada rubro está en su módulo")
        cod, rub = sap["codigo"].fillna("").astype(str).str.strip(), sap["rubro"].fillna("").astype(str).str.strip()
        sap["Rubro"] = [r if (not c or r.startswith(c)) else f"{c} · {r}".strip(" ·") for c, r in zip(cod, rub)]
        piv = sap.pivot_table(index="Rubro", columns="mes", values="valor", aggfunc="sum", fill_value=0.0)
        piv.columns = [MESES_L[int(m) - 1].capitalize() if 1 <= int(m) <= 12 else f"Mes {int(m)}" for m in piv.columns]
        piv["Total"] = piv.sum(axis=1)
        piv = piv.sort_values("Total", ascending=False)
        piv.loc["TOTAL"] = piv.sum()
        seg = E.segmentos_top([(r, v) for r, v in piv["Total"].drop("TOTAL").items()])
        if seg:
            st.markdown("**Participación de cada rubro en la ejecución**")
            st.html(E.barra_segmentada(seg))
        st.dataframe(piv.reset_index(), hide_index=True, width="stretch", height=min(40 + 35 * len(piv), 460),
                     column_config={c: st.column_config.NumberColumn(c, format="accounting") for c in piv.columns})

    with st.container(key="card_ej_estado"):
        titulo_tarjeta("Estado de cada periodo", "Qué se cargó y qué se aplicó por compañía y mes. Si falta algo, se carga en «Mantenimiento y Operaciones» → «SISPO del mes».")
        filas = []
        for e in empresas:
            for m in sorted(set(sap.loc[sap["empresa"] == e, "mes"])):
                per = anio * 100 + int(m)
                s_ = sap[(sap["empresa"] == e) & (sap["mes"] == m)]
                a = archivos[(archivos["empresa"] == e) & (archivos["periodo"] == per)]
                x = mt[(mt["empresa"] == e) & (mt["periodo"] == per)]
                o = op[(op["empresa"] == e) & (op["periodo"] == per)]
                falta = [n for n, ok in (("SISPO guardado", len(a)), ("Mantenimiento", len(x)), ("Operaciones", len(o))) if not ok]
                filas.append({"Compañía": e, "Mes": MESES_L[int(m) - 1].capitalize() if 1 <= int(m) <= 12 else f"Mes {int(m)}",
                              "SAP · líneas": int(s_["lineas"].sum()), "SAP · $": float(s_["valor"].sum()),
                              "SISPO guardado": (str(a["archivo"].iloc[0]) if len(a) else "—"),
                              "Mantenimiento · líneas": int(x["lineas"].sum()) if len(x) else 0,
                              "Mantenimiento · $ SISPO": float(x["valor"].sum()) if len(x) else 0.0,
                              "Operaciones · líneas": int(o["lineas"].sum()) if len(o) else 0,
                              "Operaciones · $ (pesos)": float(o["valor"].sum()) if len(o) else 0.0,
                              "Estado": "Completo" if not falta else "Falta: " + ", ".join(falta)})
        t = pd.DataFrame(filas)
        if t.empty:
            st.caption("No hay periodos para mostrar.")
        else:
            nc = st.column_config.NumberColumn
            st.dataframe(t, hide_index=True, width="stretch", height=min(40 + 35 * len(t), 420), column_config={
                "SAP · líneas": nc(format="localized"), "SAP · $": nc(format="accounting"),
                "Mantenimiento · líneas": nc(format="localized"), "Mantenimiento · $ SISPO": nc(format="accounting"),
                "Operaciones · líneas": nc(format="localized"), "Operaciones · $ (pesos)": nc(format="accounting"),
                "Estado": st.column_config.TextColumn("Estado", width="medium")})
            pend = t[t["Estado"] != "Completo"]
            if len(pend):
                st.warning(f"{len(pend)} de {len(t)} periodos tienen algo pendiente. Un SISPO guardado alimenta Mantenimiento y Operaciones; "
                           "falta «aplicarlo» cuando la columna de líneas está en 0.", icon=":material/rule:")
            st.caption("«Mantenimiento · $ SISPO» y «Operaciones · $» son lo que cada módulo guardó del SISPO (Operaciones, solo Consumibles, Herramientas y Bolsas). "
                       "La ejecución final de cada módulo se ve en su página.")

    with st.container(key="card_ej_cargas"):
        titulo_tarjeta("Últimas cargas de ejecución", "Del historial de cargas (SAP y SISPO)")
        h = leer("SELECT fecha AS Fecha, usuario AS Usuario, tipo AS Tipo, archivo AS Archivo, detalle AS Detalle FROM cargas "
                 "WHERE tipo IN ('Ejecución', 'SISPO', 'Reclasificación SISPO') ORDER BY rowid DESC LIMIT 15")
        if h.empty:
            st.caption("Aún no hay cargas registradas.")
        else:
            st.dataframe(h, hide_index=True, width="stretch", height=min(40 + 35 * len(h), 420))


def cargar_dotacion2(ctx):
    """Dotación y EPP's: carga del presupuesto por elemento y equivalencia artículo SISPO -> elemento."""
    import dotacion_sispo as D
    con, leer, usuario = ctx.con, ctx.leer, ctx.usuario
    con.crear_tablas(D.TABLAS)
    guardados = sispo_fuente.listar(leer)
    ult_sispo = int(guardados["periodo"].max()) // 100 if not guardados.empty else None
    with st.container(key="card_adm_dt2"):
        titulo_tarjeta("Dotación 2 · presupuesto",
                       "Sube el libro «3.1 DOTACION Y EPPS». Reemplaza el presupuesto de esa compañía y año (si el libro trae más de un año, de todos ellos). "
                       "La ejecución es la de SAP (código 3.1 · «Dotación y epp´s»); el SISPO (cuenta DOTACIÓN PERSONAL) la detalla por elemento.")
        c1, c2 = st.columns([2, 1])
        f = c1.file_uploader("Presupuesto de dotación", type=["xlsx", "xlsm"], key="dt2_ppto_up")
        empresa = c2.selectbox("Compañía", ["ALDC", "ALSM"], key="dt2_ppto_emp")
        if f:
            try:
                d, info = D.leer_ppto(f)
            except Exception as err:
                st.error(f"No pude leer {f.name}: {err}")
            else:
                anio_hoja = info["anio"]
                por_defecto = ult_sispo or anio_hoja or 2026
                anio = int(st.number_input("Año del presupuesto", 2020, 2100, int(por_defecto), key="dt2_ppto_anio"))
                if anio_hoja and anio != anio_hoja:
                    st.warning(f"La hoja se llama «{info['hoja']}» (año {anio_hoja}) pero la vas a guardar como {anio}. "
                               "Confirma que es el año correcto: la comparación usa este año para cruzar con el SISPO.")
                if info["bloques"] > 1:       # el libro trae varios años: el primero es el «Año del presupuesto» y los siguientes van +1, +2…
                    st.markdown(f"**{f.name}:** {info['elementos']} elementos · {info['bloques']} años en el libro (el primero es el año indicado arriba)")
                    for k, sec in info["por_bloque"].items():
                        st.markdown(f"- **{anio + k}:** " + " · ".join(f"{n}: {E.millones(v)}" for n, v in sec.items()) + f" · total {E.millones(sum(sec.values()))}")
                else:
                    st.markdown(f"**{f.name}:** {info['elementos']} elementos · " + " · ".join(
                        f"{k}: {E.millones(v)}" for k, v in info["por_seccion"].items()) + f" · total {E.millones(info['valor_total'])}")
                st.caption("«Actividades» (exámenes, capacitaciones…) se guarda pero no se compara: no pasa por el inventario del SISPO.")
                if st.button("Cargar presupuesto de dotación", type="primary", key="dt2_ppto_ok"):
                    d = d.assign(empresa=empresa, anio=anio + d["anio_rel"])
                    for a_ in sorted(d["anio"].unique()):
                        con.execute("DELETE FROM dot_ppto WHERE empresa=? AND anio=?", (empresa, int(a_)))
                    con.anexar(d[["empresa", "anio", "seccion", "elemento", "precio", "mes", "cantidad", "valor"]], "dot_ppto")
                    con.commit()
                    ctx.registrar_carga(usuario, "Dotación presupuesto", f.name, f"{empresa} {anio} · {E.millones(d['valor'].sum())}")
                    st.success("Presupuesto de dotación cargado.")
                    st.rerun()
    # ---- equivalencias
    with st.container(key="card_adm_dt2_map"):
        titulo_tarjeta("Dotación 2 · equivalencia de artículos",
                       "Cada artículo del SISPO se lleva a un elemento del presupuesto. Las sugerencias son solo coincidencias evidentes; "
                       "lo que dejes en «Sin elemento asignado» aparece aparte en la comparación (no se pierde del total).")
        emp = st.selectbox("Compañía", ["ALDC", "ALSM"], key="dt2_map_emp")
        pp = leer("SELECT DISTINCT seccion, elemento FROM dot_ppto WHERE empresa=?", (emp,))
        elementos = sorted(pp.loc[pp["seccion"].isin(D.SECCIONES_COMPARADAS), "elemento"])
        arts = []
        for per in guardados.loc[guardados["empresa"] == emp, "periodo"]:
            try:
                arts.append(D.ejecucion_sispo(L.sispo_leer(sispo_fuente.abrir(leer, emp, int(per)))))
            except Exception as err:
                st.error(f"No pude leer el SISPO {emp} {per}: {err}")
        if not elementos:
            st.info("Primero carga el presupuesto de dotación de esa compañía.")
        elif not arts:
            st.info("No hay SISPO guardado de esa compañía (se carga en Administración → Mantenimiento y Operaciones → SISPO del mes).")
        else:
            ej = pd.concat(arts, ignore_index=True)
            gu = leer("SELECT codigo, elemento FROM dot_map WHERE empresa=?", (emp,))
            eq = D.equivalencias(ej, elementos, dict(zip(gu["codigo"], gu["elemento"])))
            val = ej.groupby("codigo")["valor"].sum()
            eq["ejecutado"] = eq["codigo"].map(val)
            opciones = [D.SIN_ELEMENTO] + elementos
            ed = st.data_editor(eq[["codigo", "articulo", "elemento", "origen", "ejecutado"]].sort_values("ejecutado", ascending=False),
                                hide_index=True, width="stretch", key=f"dt2_map_ed_{emp}",
                                disabled=["codigo", "articulo", "origen", "ejecutado"],
                                column_config={"elemento": st.column_config.SelectboxColumn("Elemento del presupuesto", options=opciones, required=True),
                                               "ejecutado": st.column_config.NumberColumn("Ejecutado SISPO", format="$ %.0f")})
            st.caption(f"{(eq['origen'] == 'Pendiente').sum()} artículos sin elemento · {(eq['origen'] == 'Sugerido').sum()} con sugerencia sin confirmar.")
            if st.button("Guardar equivalencias", type="primary", key="dt2_map_ok"):
                filas = ed[ed["elemento"] != D.SIN_ELEMENTO][["codigo", "articulo", "elemento"]].assign(
                    empresa=emp, usuario=usuario, fecha=pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"))
                con.execute("DELETE FROM dot_map WHERE empresa=?", (emp,))
                if not filas.empty:
                    con.anexar(filas[["empresa", "codigo", "articulo", "elemento", "usuario", "fecha"]], "dot_map")
                con.commit()
                ctx.registrar_carga(usuario, "Dotación equivalencias", "pantalla", f"{emp}: {len(filas)} artículos")
                st.success("Equivalencias guardadas.")
                st.rerun()


AUTO = "(sugerencia automática)"
