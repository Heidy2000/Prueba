"""Módulo Dotación y EPP: indicadores, lo ejecutado (salidas de almacén), el comparativo contra el presupuesto y el
presupuesto por artículo."""
from html import escape

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import dotacion
import estilo as E
from ui import EMPRESAS, MESES, MESES_L, a_excel, ahora, estilo_grafico, selector_periodo, titulo_tarjeta

ORDEN_CATEGORIAS = ["DOTACIÓN", "EPP", "ACTIVIDADES", "SIN HOMOLOGAR"]
NOMBRE = {"DOTACIÓN": "Dotación", "EPP": "EPP", "ACTIVIDADES": "Actividades",
          "SIN HOMOLOGAR": "Sin cruce / sin clasificar"}
ESTADO_TXT = {"sin_cruce": "sin cruce", "sin_concepto": "sin concepto en la tabla",
              "sin_presupuesto": "sin presupuesto", "sin_clasificar": "contable sin clasificar"}


def _cats(presentes) -> list[str]:
    presentes = set(presentes)
    return [c for c in ORDEN_CATEGORIAS if c in presentes] + sorted(presentes - set(ORDEN_CATEGORIAS))


def _unidades(v: float) -> str:
    return f"{v:,.0f}".replace(",", ".")


def mostrar(ctx, zona_encabezado=None):
    leer = ctx.leer
    zona_encabezado = zona_encabezado or st.container()
    hay = pd.concat([leer("SELECT DISTINCT empresa, anio FROM dotacion_ppto"),
                     leer("SELECT DISTINCT empresa, anio FROM dotacion_ejec")]).drop_duplicates()
    if hay.empty:
        zona_encabezado.html(E.encabezado("Dotación y EPP", "Dotación y elementos de protección personal", []))
        st.info("Aún no hay información de dotación. Financiera la carga en Administración → Dotación y EPP.")
        return

    # ------------------------------------------------------------ filtros
    with st.container(key="card_filtros"):
        st.html('<div class="nf-ftit">Filtros del informe</div>')
        c0, c1, c2, c3, c4 = st.columns([1.1, 0.6, 1.3, 0.9, 1.2], vertical_alignment="bottom")
        empresa = c0.selectbox("Compañía", list(EMPRESAS), key="dt_empresa",
                               format_func=lambda e: f"{e} · {EMPRESAS[e]}")
        anios = sorted(hay.loc[hay["empresa"] == empresa, "anio"].astype(int).unique(), reverse=True)
        if not anios:
            zona_encabezado.html(E.encabezado(f"Dotación y EPP · {empresa}", EMPRESAS[empresa], []))
            st.info(f"Aún no hay información de dotación de {empresa}.", icon=":material/info:")
            return
        anio = c1.selectbox("Año", anios, key="dt_anio",
                            index=anios.index(ahora().year) if ahora().year in anios else 0)
        ppto = leer("SELECT * FROM dotacion_ppto WHERE empresa=? AND anio=?", (empresa, int(anio)))
        ejec = leer("SELECT * FROM dotacion_ejec WHERE empresa=? AND anio=?", (empresa, int(anio)))
        cruce = leer("SELECT articulo, concepto, origen FROM dotacion_cruce WHERE empresa=?", (empresa,)).fillna("")
        # ejecución contable de los rubros 3.1 (dotación y SST) que no pasa por el almacén
        sap = leer("SELECT anio, mes, fecha, codigo, sub_rubro, cc_nombre, documento, tercero, texto, valor, clave "
                   "FROM ejecucion WHERE empresa=? AND anio=? AND codigo='3.1'", (empresa, int(anio)))
        clasif = leer("SELECT tipo, llave, concepto FROM dotacion_contable WHERE empresa=?", (empresa,)).fillna("")
        meses_ej = list(ejec["mes"]) + list(sap["mes"]) if not sap.empty else list(ejec["mes"])
        ult = int(max(meses_ej)) if meses_ej else 12
        per = selector_periodo(c2, c3, ult, "dt")
        medida = c4.segmented_control("Ver", ["Valor", "Cantidad"], default="Valor", key="dt_medida") or "Valor"

    meses = per.meses()
    if ejec.empty:
        asig = ejec.assign(estado="", concepto="", categoria="")
    else:
        asig = dotacion.asignar(ejec, cruce, ppto)
    asig = asig.assign(origen=dotacion.ORIGEN_INV, detalle=asig["articulo"])
    directos, inv_sap = dotacion.contables(sap, clasif, ppto)
    if not directos.empty:
        asig = pd.concat([asig, directos.drop(columns=["patron", "sugerido", "sub_rubro", "tercero"])],
                         ignore_index=True)
    cmp = dotacion.comparativo(ppto, asig, meses)
    dec = dotacion.decisiones(ejec, cruce, ppto) if not ejec.empty else pd.DataFrame(
        columns=["tipo", "articulo", "detalle", "valor", "accion"])
    pend = directos[directos["estado"].isin(["sin_clasificar", "sin_presupuesto"])]
    if not pend.empty:
        extra = []
        for (_, estado), g in pend.groupby(["patron", "estado"]):
            extra.append(("Registro contable sin clasificar" if estado == "sin_clasificar" else "Concepto sin presupuesto",
                          f"{g['tercero'].iloc[0]} · {g['detalle'].iloc[0]}",
                          f"{len(g)} registros contables · {g['grupo'].iloc[0]}", float(g["valor"].sum()),
                          "Clasifícalo en Administración → Dotación y EPP → Ejecución contable"))
        dec = pd.concat([dec, pd.DataFrame(extra, columns=["tipo", "articulo", "detalle", "valor", "accion"])],
                        ignore_index=True).sort_values("valor", ascending=False, key=abs).reset_index(drop=True)

    chips = [(f"{'Mes' if per.solo else 'Corte'}: {MESES[per.mes - 1]} {anio}", "")]
    if not ejec.empty:
        chips.append((f"Ejecutado hasta {MESES[ult - 1]} {anio}", "ok"))
    if not dec.empty:
        chips.append((f"{dec['articulo'].nunique()} artículos por decidir", "warn"))
    zona_encabezado.html(E.encabezado(f"Dotación y EPP · {empresa}",
                                      f"Presupuesto vs ejecutado · {per.texto} · millones de pesos con IVA", chips))

    t1, t2, t3, t5, t4 = st.tabs(["Indicadores", "Ejecutado", "Comparativo", "Mensual", "Presupuesto"],
                                 on_change="rerun", key="dt_tabs")
    with t1:
        if t1.open:
            _indicadores(ppto, asig, cmp, per, ult)
    with t2:
        if t2.open:
            _ejecutado(asig, per, medida, empresa, anio, inv_sap, ejec)
    with t5:
        if t5.open:
            _mensual(ppto, asig, per, ult, empresa, anio)
    with t3:
        if t3.open:
            _comparativo(cmp, dec, per, empresa, anio)
    with t4:
        if t4.open:
            _presupuesto(ppto, medida, empresa, anio)

    _informe_resumen(ppto, asig, cmp, per, ult, empresa, anio)


def _semaforo(malo: bool) -> dict:
    return dict(color=E.ROJO if malo else E.VERDE, fondo=E.ROJO_SUAVE if malo else E.VERDE_SUAVE)


# ------------------------------------------------------------------ indicadores
def _indicadores(ppto, asig, cmp, per, ult):
    """Vista corta: cuatro cifras, una frase que explica la desviación, el mes a mes, las categorías y las cinco
    desviaciones que más pesan (el detalle completo está en Comparativo y Mensual)."""
    if asig.empty:
        st.info("Aún no hay ejecución de dotación de este año. Financiera la carga en Administración → Dotación y EPP.")
        return
    act = asig[asig["estado"] != "excluido"]
    e_per, p_per = per.filtrar(act), per.filtrar(ppto)
    ej, pp = float(e_per["valor"].sum()), float(p_per["valor"].sum())
    anual = float(ppto["valor"].sum())
    desv = ej - pp
    restante = float(ppto.loc[ppto["mes"] > per.mes, "valor"].sum())
    proy = float(act.loc[act["mes"] <= per.mes, "valor"].sum()) + restante
    ef = dotacion.efectos(cmp) if not cmp.empty else pd.DataFrame()
    e_precio = float(ef["efecto_precio"].sum()) if not ef.empty else 0.0
    e_cant = float(ef["efecto_cantidad"].sum()) if not ef.empty else 0.0
    sin = e_per[e_per["estado"].isin(["sin_cruce", "sin_concepto", "sin_presupuesto", "sin_clasificar"])]
    cumpl = ej / pp if pp else float("nan")

    st.html(E.kpis([
        dict(label=f"Ejecutado {per.suf}", valor=E.millones(ej),
             delta=E.delta_valor(desv, pp, a_favor_si_sube=False), sub=f"Presupuesto {E.millones(pp)}",
             icono="receipt", color=E.AZUL, fondo=E.AZUL_SUAVE,
             spark=None if per.solo else [float(act.loc[act["mes"] == m, "valor"].sum()) if m <= ult else None
                                          for m in range(1, per.mes + 1)]),
        dict(label="Cumplimiento", valor=E.pct_txt(cumpl),
             sub=f"Desviación {'+' if desv >= 0 else ''}{E.millones(desv)}", icono="check",
             **_semaforo(desv > 0)),
        dict(label="Cierre proyectado", valor=E.millones(proy),
             delta=E.delta_valor(proy - anual, anual, a_favor_si_sube=False, sufijo="vs anual"),
             sub=f"Presupuesto anual {E.millones(anual)}", icono="piggy", **_semaforo(proy > anual)),
        dict(label="Por resolver", valor=E.millones(sin["valor"].sum()),
             sub=f"{sin['articulo'].nunique()} partidas sin cruce o sin clasificar", icono="alert",
             **_semaforo(len(sin) > 0)),
    ]))

    # una sola frase: cuánto, por qué (cantidad o precio) y qué falta resolver
    causa = ""
    if abs(e_cant) >= 1 or abs(e_precio) >= 1:
        fuerte, debil = (("cantidad", e_cant), ("precio", e_precio))[::1 if abs(e_cant) >= abs(e_precio) else -1]
        txt_ = {"cantidad": ("se entregaron más unidades", "se entregaron menos unidades"),
                "precio": ("se pagó más caro", "se compró más barato")}
        causa = (f" Pesa más la <b>{fuerte[0]}</b>: {txt_[fuerte[0]][0 if fuerte[1] > 0 else 1]} "
                 f"({'+' if fuerte[1] >= 0 else ''}{E.millones(fuerte[1])}); por {debil[0]}, "
                 f"{'+' if debil[1] >= 0 else ''}{E.millones(debil[1])}.")
    pend = f" Quedan {E.millones(sin['valor'].sum())} por resolver en Administración." if len(sin) else ""
    st.html(E.lectura(f"Dotación {per.suf}: {E.pct_txt(cumpl)} del presupuesto "
                      f"({'+' if desv >= 0 else ''}{E.millones(desv)}).", causa.strip() + pend,
                      "ok" if desv <= 0 else ("warn" if pp and desv / pp <= 0.05 else "bad")))

    izq, der = st.columns([2.1, 1], gap="medium")
    with izq, st.container(key="card_dt_mes"):
        titulo_tarjeta("Ejecutado vs presupuesto por mes", "Barras: ejecutado · línea: presupuesto · millones")
        e_m = act.groupby("mes")["valor"].sum().reindex(range(1, 13))
        p_m = ppto.groupby("mes")["valor"].sum().reindex(range(1, 13)).fillna(0)
        e_m.loc[[m for m in range(1, 13) if m > ult]] = None
        fig = go.Figure()
        fig.add_bar(x=MESES, y=e_m / 1e6, name="Ejecutado", marker_color=E.AZUL, marker_line_width=0,
                    texttemplate="%{y:,.0f}", textposition="outside", cliponaxis=False, hovertemplate="%{y:,.1f} M")
        fig.add_scatter(x=MESES, y=p_m / 1e6, name="Presupuesto", mode="lines+markers",
                        line=dict(color="#e8892c", width=2.5), marker=dict(size=6), hovertemplate="%{y:,.1f} M")
        if per.solo:
            fig.add_vrect(x0=per.mes - 1.5, x1=per.mes - 0.5, fillcolor=E.AZUL, opacity=.07, line_width=0)
        fig.update_layout(bargap=.35, yaxis=dict(tickformat=",.0f", ticksuffix=" M"))
        st.plotly_chart(estilo_grafico(fig, 320), width="stretch", config={"displayModeBar": False})
    with der, st.container(key="card_dt_cat"):
        titulo_tarjeta("Por categoría", "% del presupuesto del periodo")
        filas = [(NOMBRE.get(cat, cat.capitalize()), float(e_per.loc[e_per["categoria"] == cat, "valor"].sum()),
                  float(p_per.loc[p_per["categoria"] == cat, "valor"].sum()))
                 for cat in _cats(set(e_per["categoria"]) | set(p_per["categoria"]))]
        st.html(E.lista_avance(filas))

    if not ef.empty:
        with st.container(key="card_dt_desv"):
            titulo_tarjeta("Las 5 desviaciones que más pesan",
                           f"Ejecutado − presupuesto {per.suf} · la marca dice si la explica la cantidad o el precio")
            ef = ef.assign(desv=ef["ejec_valor"] - ef["ppto_valor"]).sort_values("desv", key=abs, ascending=False)
            filas = []
            for r in ef.head(5).itertuples():
                motivo = "Cantidad" if abs(r.efecto_cantidad) >= abs(r.efecto_precio) else "Precio"
                filas.append((motivo, r.concepto, f"{'+' if r.desv >= 0 else ''}{E.millones(r.desv)}",
                              "bad" if r.desv > 0 else "ok"))
            st.html(E.lista_marcada(filas))
            st.caption("El detalle por concepto, con unidades y precios, está en Comparativo; el mes a mes, en Mensual.")


# ------------------------------------------------------------------ ejecutado
def _ejecutado(asig, per, medida, empresa, anio, inv_sap=None, kardex=None):
    if asig.empty:
        st.info("Aún no hay ejecución de dotación de este año. Financiera la carga en Administración → Dotación y EPP.")
        return
    with st.container(key="card_dt_ej"):
        titulo_tarjeta("Dotación ejecutada", "Salidas de almacén y registros contables directos al costo por "
                                             "artículo o concepto · haz clic en una categoría para desplegarla")
        c0, c1, c2 = st.columns([1.5, 1.4, 1.0])
        origenes = [o for o in (dotacion.ORIGEN_INV, dotacion.ORIGEN_DIRECTO, dotacion.ORIGEN_MANUAL,
                                dotacion.ORIGEN_PEND) if o in set(asig["origen"])]
        or_f = c0.multiselect("Origen", origenes, placeholder="Todos los orígenes", key="dt_ej_or")
        cc_f = c1.multiselect("Centro de costo", sorted(asig["cc"].astype(str).unique()), placeholder="Todos",
                              key="dt_ej_cc")
        bod_f = c2.multiselect("Bodega", sorted(asig["bodega"].astype(str).unique()), placeholder="Todas",
                               key="dt_ej_bod")
        e = asig[asig["mes"].isin(per.meses())]
        if or_f:
            e = e[e["origen"].isin(or_f)]
        if cc_f:
            e = e[e["cc"].isin(cc_f)]
        if bod_f:
            e = e[e["bodega"].isin(bod_f)]
        campo, fmt = ("valor", "m") if medida == "Valor" else ("cantidad", "n")
        st.html(E.kpis([
            dict(label="Ejecutado", valor=E.millones(e["valor"].sum()),
                 sub=f"Sin IVA {E.millones(e['valor_sin_iva'].sum())}", icono="receipt", color=E.AZUL,
                 fondo=E.AZUL_SUAVE),
            dict(label="Unidades entregadas", valor=_unidades(e["cantidad"].sum()),
                 sub=f"{e['articulo'].nunique()} artículos distintos", icono="check", color=E.VERDE,
                 fondo=E.VERDE_SUAVE),
            dict(label="Salidas de almacén", valor=_unidades(int((e["origen"] == dotacion.ORIGEN_INV).sum())),
                 sub=per.texto.capitalize(), icono="calendar", color=E.AZUL, fondo=E.AZUL_SUAVE),
            dict(label="Registros contables", valor=E.millones(e.loc[e["origen"] != dotacion.ORIGEN_INV, "valor"].sum()),
                 sub=f"{int((e['origen'] != dotacion.ORIGEN_INV).sum())} registros directos al costo",
                 icono="wallet", color=E.AZUL, fondo=E.AZUL_SUAVE),
        ]))
        por_o = e.groupby("origen")["valor"].sum()
        st.html(E.lista_participacion([(o, float(por_o[o])) for o in origenes if o in por_o.index],
                                      total=float(por_o.sum())))
        meses = per.meses()
        ancho = e.pivot_table(index=["categoria", "articulo"], columns="mes", values=campo, aggfunc="sum") \
            .reindex(columns=meses).fillna(0.0)
        nodos = []
        for cat in (_cats(ancho.index.get_level_values(0)) if not ancho.empty else []):
            sub = ancho.loc[cat]
            sub = sub.loc[sub.sum(axis=1).sort_values(ascending=False).index]
            hijos = [dict(nombre=a, valores=[float(r.sum())] + [float(r[m]) for m in meses], clase="linea", fmt=fmt)
                     for a, r in sub.iterrows()]
            tot = sub.sum()
            nodos.append(dict(nombre=NOMBRE.get(cat, cat.capitalize()), clase="seccion", hijos=hijos, fmt=fmt,
                              valores=[float(tot.sum())] + [float(tot[m]) for m in meses], sub=f"{len(sub)} artículos"))
        if nodos:
            st.html(E.tabla_meses(nodos, ["Total"] + [MESES[m - 1] for m in meses]))
        st.caption("Salidas de almacén con IVA; registros contables por el valor llevado al costo. Los artículos "
                   "'Sin cruce' y los registros contables sin clasificar los decide Financiera en Administración → "
                   "Dotación y EPP. Los registros contables no traen cantidades.")

    with st.container(key="card_dt_movs"):
        titulo_tarjeta("Detalle de la ejecución", f"Salidas de almacén y registros contables · {per.texto}")
        buscar = st.text_input("Buscar", placeholder="Artículo, texto, centro de costo, código…", key="dt_ej_buscar")
        d = e.assign(fecha=pd.to_datetime(e["fecha"], errors="coerce"))
        if buscar:
            d = d[d[["articulo", "detalle", "cc", "codigo", "concepto"]].astype(str)
                  .apply(lambda s: s.str.contains(buscar, case=False, regex=False)).any(axis=1)]
        d = d.sort_values("fecha", ascending=False)[["fecha", "origen", "detalle", "concepto", "cc", "bodega",
                                                     "cantidad", "costo_unitario", "valor"]]
        st.dataframe(d, hide_index=True, width="stretch", height=420, column_config={
            "fecha": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"), "origen": "Origen",
            "detalle": st.column_config.TextColumn("Artículo / texto contable", width="large"),
            "concepto": "Concepto del presupuesto", "cc": "Centro de costo", "bodega": "Bodega",
            "cantidad": st.column_config.NumberColumn("Cantidad", format="%.0f"),
            "costo_unitario": st.column_config.NumberColumn("Costo unitario sin IVA", format="accounting"),
            "valor": st.column_config.NumberColumn("Valor", format="accounting")})
        st.download_button("Descargar ejecución", lambda x=d: a_excel(x), f"dotacion_ejecutada_{empresa}_{anio}.xlsx",
                           icon=":material/download:", key="dt_ej_dl")

    if inv_sap is not None and not inv_sap.empty and kardex is not None and not kardex.empty:
        with st.expander("Cuadre: salidas del kárdex vs registro contable de las salidas de almacén",
                         icon=":material/rule:"):
            k_m = kardex.groupby("mes")["valor"].sum()
            c_m = inv_sap.groupby("mes")["valor"].sum()
            t = pd.DataFrame({"Kárdex (con IVA)": k_m, "Contabilidad (salidas de inventario)": c_m}).fillna(0.0)
            t["Diferencia"] = t["Contabilidad (salidas de inventario)"] - t["Kárdex (con IVA)"]
            t.index = [MESES[int(m) - 1] for m in t.index]
            st.dataframe(t, width="stretch", column_config={
                c: st.column_config.NumberColumn(format="accounting") for c in t.columns})
            st.caption("Los registros contables de 'Salida de inventario' no se suman (ya están en el kárdex); esta "
                       "tabla solo sirve para verificar que el almacén y la contabilidad cuadran mes a mes.")


# ------------------------------------------------------------------ comparativo
def _comparativo(cmp, dec, per, empresa, anio):
    if not dec.empty:
        with st.container(key="card_dt_dec"):
            titulo_tarjeta("Decisiones pendientes",
                           f"{dec['articulo'].nunique()} artículos o registros ({E.millones(dec['valor'].sum())}) no "
                           "cruzan bien con el presupuesto. Igual suman en el comparativo; Financiera los resuelve en "
                           "Administración → Dotación y EPP (Cruce de artículos o Ejecución contable).")
            plantilla = "minmax(0,2.2fr) minmax(0,1.2fr) minmax(0,2.4fr) minmax(0,1fr)"
            cab = (f'<div class="nf-dev nf-head" style="grid-template-columns:{plantilla}"><div>Artículo de la '
                   'ejecución</div><div>Qué pasa</div><div>Decisión que debes tomar</div><div>Ejecutado</div></div>')
            filas = "".join(
                f'<div class="nf-dev" style="grid-template-columns:{plantilla}"><div><b>{escape(r.articulo)}</b><br>'
                f'<small style="color:{E.TEXTO_2}">{escape(r.detalle)}</small></div><div>{escape(r.tipo)}</div>'
                f'<div><small>{escape(r.accion)}</small></div><div style="font-weight:700">{E.millones(r.valor)}</div>'
                '</div>' for r in dec.head(12).itertuples())
            st.html(f'<div class="nf-t notranslate" translate="no" style="border:1px solid {E.BORDE};'
                    f'border-radius:12px;overflow:hidden">{cab}{filas}</div>')
            if len(dec) > 12:
                st.caption(f"Y {len(dec) - 12} más en Administración.")

    with st.container(key="card_dt_cmp"):
        titulo_tarjeta("Ejecutado vs presupuesto por concepto",
                       f"{per.texto.capitalize()} · haz clic en una categoría para ver sus conceptos")
        if cmp.empty:
            st.info("No hay presupuesto ni ejecución para este periodo.")
            return
        columnas = ["Presupuesto", "Ejecutado", "Diferencia", "% ejecución", "Unid. ppto", "Unid. entregadas"]
        formatos = ["m", "m", "m", "p", "n", "n"]

        def valores(p, e, up, ue):
            return [p, e, e - p, (e / p) if p else None, up, ue]

        nodos = []
        for cat in _cats(cmp["categoria"]):
            g = cmp[cmp["categoria"] == cat]
            g = g.loc[(g["ppto_valor"] + g["ejec_valor"]).sort_values(ascending=False).index]
            g = g[(g["ppto_valor"] != 0) | (g["ejec_valor"] != 0)]
            if g.empty:
                continue
            hijos = [dict(nombre=r.concepto, valores=valores(r.ppto_valor, r.ejec_valor, r.ppto_cant, r.ejec_cant),
                          clase="linea", dif_col=2, dif_inv=True, sub=ESTADO_TXT.get(r.estado, ""))
                     for r in g.itertuples(index=False)]
            nodos.append(dict(nombre=NOMBRE.get(cat, cat.capitalize()), clase="seccion", hijos=hijos, dif_col=2,
                              dif_inv=True, valores=valores(g["ppto_valor"].sum(), g["ejec_valor"].sum(),
                                                            g["ppto_cant"].sum(), g["ejec_cant"].sum())))
        nodos.append(dict(nombre="Total", clase="saldo", dif_col=2, dif_inv=True,
                          valores=valores(cmp["ppto_valor"].sum(), cmp["ejec_valor"].sum(), cmp["ppto_cant"].sum(),
                                          cmp["ejec_cant"].sum())))
        st.html(E.tabla_meses(nodos, columnas, formatos=formatos))
        st.caption("Diferencia = ejecutado − presupuesto: en rojo, por encima del presupuesto; en verde, por debajo. "
                   "Los conceptos marcados 'sin cruce' o 'sin presupuesto' suman en el total mientras se decide su cruce.")
        plano = cmp.rename(columns={"categoria": "Categoría", "concepto": "Concepto", "ppto_valor": "Presupuesto",
                                    "ejec_valor": "Ejecutado", "ppto_cant": "Unid. ppto",
                                    "ejec_cant": "Unid. entregadas", "estado": "Estado"})
        st.download_button("Descargar comparativo", lambda d=plano: a_excel(d),
                           f"dotacion_comparativo_{empresa}_{anio}.xlsx", icon=":material/download:", key="dt_cmp_dl")


# ------------------------------------------------------------------ presupuesto
def _presupuesto(p, medida, empresa, anio):
    if p.empty:
        st.info("Aún no hay presupuesto de dotación de este año.")
        return
    por_cat = p.groupby("categoria")["valor"].sum()
    total = float(por_cat.sum())
    tarjetas = [dict(label=f"Presupuesto {anio}", valor=E.millones(total),
                     sub=f"{p.loc[p['valor'] != 0, 'articulo'].nunique()} artículos", icono="receipt", color=E.AZUL,
                     fondo=E.AZUL_SUAVE)]
    for cat in ["DOTACIÓN", "EPP", "ACTIVIDADES"]:
        v = float(por_cat.get(cat, 0.0))
        tarjetas.append(dict(label=NOMBRE[cat], valor=E.millones(v), sub=f"{v / total:.0%} del total" if total else "",
                             icono="wallet", color=E.AZUL, fondo=E.AZUL_SUAVE))
    st.html(E.kpis(tarjetas))
    with st.container(key="card_dt_tabla"):
        titulo_tarjeta("Presupuesto por artículo", "Haz clic en una categoría para ver sus artículos.")
        campo, fmt = ("valor", "m") if medida == "Valor" else ("cantidad", "n")
        p = p[(p["valor"] != 0) | (p["cantidad"] != 0)]
        ancho = p.pivot_table(index=["categoria", "articulo"], columns="mes", values=campo, aggfunc="sum") \
            .reindex(columns=range(1, 13)).fillna(0.0)
        orden_art = p.groupby(["categoria", "articulo"])["orden"].min()
        nodos = []
        for cat in _cats(ancho.index.get_level_values(0)):
            sub = ancho.loc[cat]
            sub = sub.loc[sorted(sub.index, key=lambda a: orden_art.get((cat, a), 0))]
            hijos = [dict(nombre=a, valores=[float(r.sum())] + [float(r[m]) for m in range(1, 13)], clase="linea",
                          fmt=fmt) for a, r in sub.iterrows()]
            tot = sub.sum()
            nodos.append(dict(nombre=NOMBRE.get(cat, cat.capitalize()), clase="seccion", hijos=hijos, fmt=fmt,
                              valores=[float(tot.sum())] + [float(tot[m]) for m in range(1, 13)]))
        st.html(E.tabla_meses(nodos, ["Total año"] + MESES))
        plano = ancho.reset_index()
        plano.columns = ["Categoría", "Artículo"] + MESES
        st.download_button("Descargar presupuesto", lambda d=plano: a_excel(d),
                           f"dotacion_{empresa}_{anio}_{campo}.xlsx", icon=":material/download:", key="dt_pp_dl")


# ------------------------------------------------------------------ comparativo mensual
def _mensual(ppto, asig, per, ult, empresa, anio):
    """Presupuesto vs ejecutado mes a mes, en valor y en cantidades: variación, % de cumplimiento, acumulado del año
    y los meses con desviación (con los conceptos que más la explican)."""
    act = asig[asig["estado"] != "excluido"] if not asig.empty else asig
    if ppto.empty and act.empty:
        st.info("No hay presupuesto ni ejecución de dotación de este año.")
        return
    with st.container(key="card_dt_men_f"):
        c1, c2 = st.columns([2.2, 1], vertical_alignment="bottom")
        cats = _cats(set(ppto["categoria"]) | set(act["categoria"]) - {"SIN HOMOLOGAR", "EXCLUIDO"})
        cat_f = c1.multiselect("Categoría", cats, format_func=lambda c: NOMBRE.get(c, c.capitalize()),
                               placeholder="Todas las categorías", key="dt_men_cat")
        umbral = c2.number_input("Desviación a resaltar (%)", 1, 100, 10, step=1, key="dt_men_umbral") / 100
    p = ppto[ppto["categoria"].isin(cat_f)] if cat_f else ppto
    e = act[act["categoria"].isin(cat_f)] if cat_f else act
    corte = min(per.mes, ult)
    meses = list(range(1, ult + 1))

    def serie(df, col, m_):
        return df.groupby("mes")[col].sum().reindex(m_, fill_value=0.0) if not df.empty else \
            pd.Series(0.0, index=m_)
    # cantidades: solo dotación y EPP que pasan por el almacén (los registros contables no traen unidades)
    p_q = p[p["categoria"].isin(["DOTACIÓN", "EPP"])]
    e_q = e[(e["origen"] == dotacion.ORIGEN_INV) & e["categoria"].isin(["DOTACIÓN", "EPP", "SIN HOMOLOGAR"])]
    pv, ev = serie(p, "valor", meses), serie(e, "valor", meses)
    pq, eq = serie(p_q, "cantidad", meses), serie(e_q, "cantidad", meses)

    def pct(a, b):
        return (a / b - 1) if b else None

    # ---------------- tarjetas
    pv_a, ev_a = float(pv.loc[:corte].sum()), float(ev.loc[:corte].sum())
    pq_a, eq_a = float(pq.loc[:corte].sum()), float(eq.loc[:corte].sum())
    fuera_v = [m for m in meses if pv[m] and abs(ev[m] / pv[m] - 1) > umbral]
    fuera_q = [m for m in meses if pq[m] and abs(eq[m] / pq[m] - 1) > umbral]
    st.html(E.kpis([
        dict(label=f"Costo ejecutado YTD a {MESES[corte - 1].lower()}", valor=E.millones(ev_a),
             delta=E.delta_valor(ev_a - pv_a, pv_a, a_favor_si_sube=False), sub=f"Presupuesto {E.millones(pv_a)}",
             icono="receipt", color=E.AZUL, fondo=E.AZUL_SUAVE, spark=[float(ev[m]) for m in meses]),
        dict(label="Cumplimiento en costo YTD", valor=E.pct_txt(ev_a / pv_a if pv_a else None),
             sub="Ejecutado ÷ presupuesto acumulado", icono="check", color=E.color_barra(
                 ev_a / pv_a if pv_a else float("nan"), ev_a, False), fondo=E.AZUL_SUAVE),
        dict(label="Unidades entregadas YTD", valor=_unidades(eq_a),
             delta=E.delta_valor(eq_a - pq_a, pq_a, a_favor_si_sube=False, sufijo="vs ppto"),
             sub=f"Presupuestadas {_unidades(pq_a)} (dotación y EPP)", icono="calendar", color=E.AZUL,
             fondo=E.AZUL_SUAVE, spark=[float(eq[m]) for m in meses]),
        dict(label="Meses con desviación", valor=f"{len(set(fuera_v) | set(fuera_q))} de {len(meses)}",
             sub=f"Más de ±{umbral:.0%} en costo ({len(fuera_v)}) o en cantidad ({len(fuera_q)})".replace(".", ","),
             icono="alert" if fuera_v or fuera_q else "check", color=E.ROJO if fuera_v or fuera_q else E.VERDE,
             fondo=E.ROJO_SUAVE if fuera_v or fuera_q else E.VERDE_SUAVE),
    ]))

    # ---------------- tendencia de las variaciones
    with st.container(key="card_dt_men_g"):
        titulo_tarjeta("Tendencia mensual de las variaciones",
                       f"Variación % del mes frente al presupuesto · franja: ±{umbral:.0%} · línea: cumplimiento "
                       "acumulado".replace(".", ","))
        vv = [pct(ev[m], pv[m]) for m in meses]
        vq = [pct(eq[m], pq[m]) for m in meses]
        cum = [(float(ev.loc[:m].sum()) / float(pv.loc[:m].sum())) if float(pv.loc[:m].sum()) else None for m in meses]
        x = [MESES[m - 1] for m in meses]
        fig = go.Figure()
        fig.add_hrect(y0=-umbral * 100, y1=umbral * 100, fillcolor=E.VERDE, opacity=.07, line_width=0)
        fig.add_bar(x=x, y=[None if v is None else v * 100 for v in vv], name="Costo",
                    marker_color=E.AZUL, marker_line_color=[E.ROJO if v is not None and abs(v) > umbral
                                                            else E.AZUL for v in vv], marker_line_width=2.5,
                    texttemplate="%{y:+.0f}%", textposition="outside", cliponaxis=False,
                    hovertemplate="Costo %{y:+.1f}%<extra></extra>")
        fig.add_bar(x=x, y=[None if v is None else v * 100 for v in vq], name="Cantidad",
                    marker_color="#9dbbea", marker_line_color=[E.ROJO if v is not None and abs(v) > umbral
                                                                else "#9dbbea" for v in vq], marker_line_width=2.5,
                    texttemplate="%{y:+.0f}%", textposition="outside", cliponaxis=False,
                    hovertemplate="Cantidad %{y:+.1f}%<extra></extra>")
        fig.add_scatter(x=x, y=[None if v is None else v * 100 for v in cum], name="Cumplimiento acumulado (costo)",
                        yaxis="y2", mode="lines+markers", line=dict(color=E.NAVY, width=2),
                        hovertemplate="%{y:.1f}%<extra></extra>")
        fig.add_hline(y=0, line_color=E.BORDE)
        fig.update_layout(barmode="group", bargap=.3, yaxis=dict(ticksuffix="%", title=None),
                          yaxis2=dict(overlaying="y", side="right", ticksuffix="%", tickformat=".0f",
                                      showgrid=False))
        st.plotly_chart(estilo_grafico(fig, 340), width="stretch", config={"displayModeBar": False})
        st.caption("Barras: (ejecutado ÷ presupuesto del mes) − 1. Costo en azul y cantidad en azul claro; con borde "
                   "rojo los meses fuera de la franja. Cantidades solo de dotación y EPP entregadas por el almacén.")

    # ---------------- tablas mes a mes
    def tabla(pres, ejec, fmt, clave_t):
        nodos = []
        for m in meses:
            v_ = pct(ejec[m], pres[m])
            alerta = v_ is not None and abs(v_) > umbral
            nodos.append(dict(nombre=MESES_L[m - 1].capitalize(), clase="linea", dif_col=[2, 6], dif_inv=True, fmt=fmt,
                              sub="fuera de rango" if alerta else "",
                              valores=[pres[m], ejec[m], ejec[m] - pres[m], v_,
                                       (ejec[m] / pres[m]) if pres[m] else None, float(ejec.loc[:m].sum()),
                                       float(ejec.loc[:m].sum() - pres.loc[:m].sum()),
                                       (float(ejec.loc[:m].sum()) / float(pres.loc[:m].sum()))
                                       if float(pres.loc[:m].sum()) else None]))
        pa, ea = float(pres.loc[:corte].sum()), float(ejec.loc[:corte].sum())
        nodos.append(dict(nombre=f"Acumulado a {MESES_L[corte - 1]} (YTD)", clase="saldo", dif_col=[2, 6], dif_inv=True,
                          fmt=fmt, valores=[pa, ea, ea - pa, pct(ea, pa), (ea / pa) if pa else None, ea, ea - pa,
                                            (ea / pa) if pa else None]))
        formatos = [fmt, fmt, fmt, "p", "p", fmt, fmt, "p"]
        for nd in nodos:
            nd.pop("fmt")
        st.html(E.tabla_meses(nodos, ["Presupuesto", "Ejecutado", "Variación", "Var. %", "% cumplimiento",
                                      "Ejec. acumulado", "Var. acumulada", "% cumpl. acumulado"], formatos=formatos,
                              barras=[2] if fmt == "m" else None))
    with st.container(key="card_dt_men_v"):
        titulo_tarjeta("Costo: presupuesto vs ejecutado mes a mes",
                       "Millones con IVA · incluye salidas de almacén y registros contables directos al costo")
        tabla(pv, ev, "m", "v")
    with st.container(key="card_dt_men_q"):
        titulo_tarjeta("Cantidades: presupuesto vs ejecutado mes a mes",
                       "Unidades de dotación y EPP · actividades y registros contables no tienen unidades")
        tabla(pq, eq, "n", "q")

    # ---------------- meses con desviación y sus causas
    with st.container(key="card_dt_men_c"):
        titulo_tarjeta("Meses con desviación y qué la explica",
                       f"Meses fuera de ±{umbral:.0%} en costo · los conceptos con mayor diferencia ese mes"
                       .replace(".", ","))
        if not fuera_v:
            st.success("Ningún mes se sale de la franja en costo.")
        cols_m = st.columns(2, gap="medium") if fuera_v else []
        for i_m, m in enumerate(fuera_v):
            pc_ = p[p["mes"] == m].groupby("articulo")["valor"].sum()
            ec_ = e[e["mes"] == m].groupby("concepto")["valor"].sum()
            dif = ec_.reindex(ec_.index.union(pc_.index), fill_value=0) - pc_.reindex(ec_.index.union(pc_.index),
                                                                                   fill_value=0)
            dif = dif[dif.abs() >= 1].sort_values(key=abs, ascending=False).head(4)
            v_ = ev[m] - pv[m]
            cols_m[i_m % 2].html(f'<div style="font-size:.85rem;font-weight:700;margin:.5rem 0 .3rem 0">'
                                 f'{MESES_L[m - 1].capitalize()}'
                    f' · <span style="color:{E.ROJO if v_ > 0 else E.VERDE}">{"+" if v_ > 0 else ""}'
                    f'{E.millones(v_)} ({"+" if v_ > 0 else ""}{E.pct_txt(pct(ev[m], pv[m]))})</span></div>'
                    + E.lista_marcada([("+" if d > 0 else "−", c, E.millones(d), "bad" if d > 0 else "ok")
                                       for c, d in dif.items()]))
        plano = pd.DataFrame({"Mes": [MESES[m - 1] for m in meses], "Ppto costo": pv.values, "Ejec. costo": ev.values,
                              "Ppto unidades": pq.values, "Ejec. unidades": eq.values})
        plano["Var. costo"] = plano["Ejec. costo"] - plano["Ppto costo"]
        plano["Var. unidades"] = plano["Ejec. unidades"] - plano["Ppto unidades"]
        st.download_button("Descargar comparativo mensual", lambda d=plano: a_excel(d),
                           f"dotacion_mensual_{empresa}_{anio}.xlsx", icon=":material/download:", key="dt_men_dl")


# ------------------------------------------------------------------ informe de resumen (al final del módulo)
def _informe_resumen(ppto, asig, cmp, per, ult, empresa, anio):
    """Las principales diferencias frente al presupuesto, en pocas frases, para leer o copiar en un correo."""
    if asig.empty and ppto.empty:
        return
    act = asig[asig["estado"] != "excluido"] if not asig.empty else asig
    e_per, p_per = per.filtrar(act), per.filtrar(ppto)
    ej, pp = float(e_per["valor"].sum()), float(p_per["valor"].sum())
    desv = ej - pp
    anual = float(ppto["valor"].sum())
    proy = float(act.loc[act["mes"] <= per.mes, "valor"].sum()) + float(ppto.loc[ppto["mes"] > per.mes, "valor"].sum())
    ef = dotacion.efectos(cmp) if not cmp.empty else pd.DataFrame()
    e_precio = float(ef["efecto_precio"].sum()) if not ef.empty else 0.0
    e_cant = float(ef["efecto_cantidad"].sum()) if not ef.empty else 0.0
    sin = e_per[e_per["estado"].isin(["sin_cruce", "sin_concepto", "sin_presupuesto", "sin_clasificar"])]

    general = [f"Se ejecutaron <b>{E.millones(ej)}</b> frente a <b>{E.millones(pp)}</b> presupuestados: "
               f"{E.var_html(desv)} ({E.pct_txt(ej / pp) if pp else '—'} del presupuesto)."]
    if abs(e_cant) >= 1 or abs(e_precio) >= 1:
        general.append(f"Por cantidad: {E.var_html(e_cant)} "
                       f"({'se entregaron más' if e_cant > 0 else 'se entregaron menos'} unidades de las presupuestadas). "
                       f"Por precio: {E.var_html(e_precio)} "
                       f"({'se pagó más caro' if e_precio > 0 else 'se compró más barato'} que el precio presupuestado).")
    if not per.solo and anual:
        general.append(f"Si el resto del año se ejecuta como está presupuestado, el año cerraría en "
                       f"<b>{E.millones(proy)}</b>: {E.var_html(proy - anual)} frente al presupuesto anual.")

    categorias = []
    for cat in _cats(set(e_per["categoria"]) | set(p_per["categoria"])):
        e_c = float(e_per.loc[e_per["categoria"] == cat, "valor"].sum())
        p_c = float(p_per.loc[p_per["categoria"] == cat, "valor"].sum())
        if abs(e_c) < 1 and abs(p_c) < 1:
            continue
        if cat == "SIN HOMOLOGAR":
            categorias.append(f"<b>{escape(NOMBRE[cat])}</b>: {E.millones(e_c)} ejecutados aún sin concepto del "
                              "presupuesto.")
            continue
        categorias.append(f"<b>{escape(NOMBRE.get(cat, cat.capitalize()))}</b>: {E.millones(e_c)} de {E.millones(p_c)} "
                          f"({E.pct_txt(e_c / p_c) if p_c else 'sin presupuesto'}) · {E.var_html(e_c - p_c)}.")

    def frase(r):
        motivo = ""
        if r.ppto_cant > 0 and r.ejec_cant > 0:
            pr_p, pr_e = r.ppto_valor / r.ppto_cant, r.ejec_valor / r.ejec_cant
            motivo = (f" — {_unidades(r.ejec_cant)} unidades vs {_unidades(r.ppto_cant)} presupuestadas; precio "
                      f"{E.pesos(pr_e)} vs {E.pesos(pr_p)}")
        elif r.ejec_valor and not r.ejec_cant and r.ppto_valor:
            motivo = " — registros contables (sin unidades)"
        elif r.ppto_valor and not r.ejec_valor:
            motivo = " — presupuestado y aún sin ejecutar"
        elif r.ejec_valor and not r.ppto_valor:
            motivo = " — ejecutado sin presupuesto"
        return f"<b>{escape(str(r.concepto))}</b>: {E.var_html(r.ejec_valor - r.ppto_valor)}{motivo}."
    c = cmp.assign(dif=cmp["ejec_valor"] - cmp["ppto_valor"]) if not cmp.empty else pd.DataFrame(columns=["dif"])
    arriba = [frase(r) for r in c[c["dif"] > 50_000].nlargest(3, "dif").itertuples()]
    abajo = [frase(r) for r in c[c["dif"] < -50_000].nsmallest(3, "dif").itertuples()]

    pendientes = []
    if len(sin):
        pendientes.append(f"{sin['articulo'].nunique()} partidas por <b>{E.millones(sin['valor'].sum())}</b> no cruzan "
                          "con un concepto del presupuesto (se resuelven en Administración → Dotación y EPP).")

    titulo = f"Informe de dotación · {EMPRESAS.get(empresa, empresa)}"
    sub = f"{per.texto.capitalize()} {anio} · valores con IVA · salidas de almacén y registros contables"
    secciones = [("En general", general), ("Por categoría", categorias),
                 ("Lo que más se pasó del presupuesto", arriba), ("Lo que quedó por debajo", abajo),
                 ("Pendiente", pendientes)]
    st.html(E.informe(titulo, sub, secciones))
    with st.expander("Copiar el informe en texto (para un correo)", icon=":material/content_copy:"):
        st.code(E.informe_texto(titulo, sub, secciones), language=None, wrap_lines=True)
