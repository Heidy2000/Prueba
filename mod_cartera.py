"""Pestaña Cartera (dentro de Estados financieros): compara la cartera de usuarios de un mes contra otro.

Replica el informe de cartera de Power BI: saldo por cuenta contrato del mes actual vs el anterior, usuarios que
pagaron o abonaron (bajó su saldo), usuarios que aumentaron, edades y tipos de cartera. Por defecto solo cuenta los
componentes propios de la compañía (ver cartera.py)."""
import re
from html import escape

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import cartera as C
import estilo as E
from ui import EMPRESAS, _partir, cascada, estilo_grafico, titulo_tarjeta

PROPIO, TODOS = "Propios", "Todos"
TOL = C.TOL


def _n(x) -> str:
    return f"{int(round(x)):,}".replace(",", ".")


def _pct(a, b):
    return (a / b - 1) if b and abs(b) > 1e-9 else float("nan")


# ------------------------------------------------------------------------------------------------ cálculos
@st.cache_data(show_spinner="Cargando la cartera del mes…", max_entries=12)
def _mes(_leer, emp: str, per: int, ver: str, cols: tuple = ()) -> pd.DataFrame:
    return C.cuentas_de(_leer, emp, per, cols or None)


def _filtrar(df: pd.DataFrame, tipos: tuple = (), cuentas: tuple = ()) -> pd.DataFrame:
    if tipos:
        df = df[df["tipo"].isin(tipos)]
    if cuentas:
        df = df[df["cuenta"].isin(cuentas)]
    return df


def _cuentas_txt(texto: str) -> tuple:
    """'12415890, 10918191 ...' -> (12415890, 10918191): acepta coma, punto y coma, espacios o saltos de línea."""
    return tuple(sorted({int(x) for x in re.split(r"[\s,;]+", texto or "") if x.strip().isdigit()}))


def _por_cuenta(df: pd.DataFrame, cols: list[str], tipos: tuple, cuentas: tuple = ()) -> pd.DataFrame:
    df = _filtrar(df, tipos, cuentas)
    d = df.assign(v=df[cols].sum(axis=1))
    agg = dict(v=("v", "sum"), clase=("clase", "first"), estrato=("estrato", "first"))
    return d.groupby("cuenta", sort=False).agg(**agg)


@st.cache_data(show_spinner="Comparando las cuentas de los dos meses…", max_entries=16)
def _base(_leer, emp: str, act: int, ant: int, propio: bool, edades, tipos: tuple, ver: str,
          cuentas: tuple = ()) -> pd.DataFrame:
    """Saldo por cuenta en los dos meses (a = actual, b = anterior) con la edad y los tipos elegidos."""
    cols = C.columnas_saldo(propio, edades)
    usar = ("cuenta", "tipo", "clase", "estrato", *cols)
    A = _por_cuenta(_mes(_leer, emp, act, ver, usar), cols, tipos, cuentas)
    B = _por_cuenta(_mes(_leer, emp, ant, ver, usar), cols, tipos, cuentas)
    m = A[["v"]].rename(columns={"v": "a"}).join(B[["v"]].rename(columns={"v": "b"}), how="outer")
    m = m.apply(pd.to_numeric, errors="coerce").fillna(0.0)
    info = pd.concat([A[["clase", "estrato"]], B[["clase", "estrato"]]])
    info = info[~info.index.duplicated()]
    m = m.join(info)
    m.index.name = "cuenta"
    return m.reset_index()


def _categoria(a, b):
    d = a - b
    return np.select([(d < -TOL) & (a.abs() < TOL), d < -TOL, (d > TOL) & (b.abs() < TOL), d > TOL],
                     ["salio", "bajo", "nuevo", "subio"], "igual")


@st.cache_data(show_spinner=False, max_entries=16)
def _comparar(_leer, emp: str, act: int, ant: int, propio: bool, edades, tipos: tuple, ver: str,
              cuentas: tuple = ()) -> dict:
    m = _base(_leer, emp, act, ant, propio, edades, tipos, ver, cuentas)
    m = m.assign(cat=_categoria(m["a"], m["b"]), d=m["a"] - m["b"],
                 na=(m["a"].abs() >= TOL).astype(int), nb=(m["b"].abs() >= TOL).astype(int))
    cat = m.groupby("cat").agg(n=("a", "size"), a=("a", "sum"), b=("b", "sum"), na=("na", "sum"), nb=("nb", "sum"))
    cat = cat.reindex(["bajo", "salio", "subio", "nuevo", "igual"]).fillna(0)
    baja = m[m["d"] < -TOL].nsmallest(15, "d")[["cuenta", "a", "b", "d", "clase", "estrato"]]
    sube = m[m["d"] > TOL].nlargest(15, "d")[["cuenta", "a", "b", "d", "clase", "estrato"]]
    cols = C.columnas_saldo(propio, edades)
    tablas = []
    for per in (act, ant):
        df = _filtrar(_mes(_leer, emp, per, ver, ("cuenta", "tipo", "clase", "estrato", *cols)), tipos, cuentas)
        v = df[cols].sum(axis=1)
        tablas.append(df.assign(v=v, n=(v.abs() >= TOL).astype(int), periodo=per))
    t = pd.concat(tablas)
    por_tipo = t.groupby(["periodo", "tipo"], observed=True)[["v", "n"]].sum().reset_index()
    por_clase = t.groupby(["periodo", "clase"], observed=True)[["v", "n"]].sum().reset_index()
    por_tipo["tipo"] = por_tipo["tipo"].astype(str)
    por_clase["clase"] = por_clase["clase"].astype(str)
    for x in (baja, sube):
        x["clase"] = x["clase"].astype(str)
    return dict(cat=cat, baja=baja, sube=sube, por_tipo=por_tipo, por_clase=por_clase)


@st.cache_data(show_spinner=False, max_entries=24)
def _por_edad_cuentas(_leer, emp: str, per: int, propio: bool, tipos: tuple, cuentas: tuple, ver: str) -> pd.Series:
    """Saldo por edad (1-9) de un mes, sumando el saldo por edad de cada cuenta (sirve con filtro de cuentas)."""
    letra = "p" if propio else "t"
    cols = [f"{letra}{i}" for i in range(1, 10)]
    df = _filtrar(_mes(_leer, emp, per, ver, ("cuenta", "tipo", *cols)), tipos, cuentas)
    return pd.Series(df[cols].sum().to_numpy(), index=range(1, 10), dtype=float)


def _excel_edades(leer, emp, act, ant, aa, propio, tipos, cuentas, ver, ed) -> bytes:
    """Cartera por edades: resumen por edad (meses comparados y mismo mes del año anterior), por edad y tipo de
    cartera, por edad y componente y, si hay filtro de cuentas, cada cuenta por edad."""
    import io
    m_act, m_ant = C.nombre_mes(act), C.nombre_mes(ant)
    letra = "p" if propio else "t"
    cols = [f"{letra}{i}" for i in range(1, 10)]
    nombres = [C.EDADES[i][0] for i in range(1, 10)]
    pers_ = [p_ for p_ in (aa, ant, act) if p_]
    serie = {p_: _por_edad_cuentas(leer, emp, p_, propio, tipos, cuentas, ver) for p_ in pers_}
    rg = _rangos(leer, emp, act, ant, propio, tipos, ver, cuentas)
    rg_e = rg[rg["rango"].isin(nombres)].set_index("rango").reindex(nombres)
    r1 = pd.DataFrame({"Edad": nombres, "Rango (tabla de edades PBI)": [C.EDADES[i][1] for i in range(1, 10)]})
    if aa:
        r1[f"Saldo {C.nombre_mes(aa)}"] = serie[aa].round(0).to_numpy()
    r1[f"Saldo {m_ant}"] = serie[ant].round(0).to_numpy()
    r1[f"Saldo {m_act}"] = serie[act].round(0).to_numpy()
    r1["Variación"] = r1[f"Saldo {m_act}"] - r1[f"Saldo {m_ant}"]
    r1["% var."] = [(a / b - 1) if b else None for a, b in zip(r1[f"Saldo {m_act}"], r1[f"Saldo {m_ant}"])]
    r1["% de la cartera"] = r1[f"Saldo {m_act}"] / r1[f"Saldo {m_act}"].sum() if r1[f"Saldo {m_act}"].sum() else None
    r1[f"Usuarios {m_ant}"] = rg_e["nb"].to_numpy()
    r1[f"Usuarios {m_act}"] = rg_e["na"].to_numpy()
    r1["Usuarios que aumentaron"] = rg_e["ns"].to_numpy()
    r1["Valor del aumento"] = rg_e["vs"].round(0).to_numpy()
    r1["Usuarios que disminuyeron"] = rg_e["nd"].to_numpy()
    r1["Valor de la disminución"] = rg_e["vd"].round(0).to_numpy()
    tot = r1.select_dtypes("number").sum()
    tot["% var."] = (tot[f"Saldo {m_act}"] / tot[f"Saldo {m_ant}"] - 1) if tot[f"Saldo {m_ant}"] else None
    r1 = pd.concat([r1, pd.DataFrame([{"Edad": "Total", "Rango (tabla de edades PBI)": "", **tot.to_dict()}])],
                   ignore_index=True)

    hojas = {"Por edad": r1}
    for p_ in (act, ant):
        df = _filtrar(_mes(leer, emp, p_, ver, ("cuenta", "tipo", *cols)), tipos, cuentas)
        t = df.groupby("tipo", observed=True)[cols].sum()
        t.columns = nombres
        t.index = [C.TIPOS.get(str(x), str(x)) for x in t.index]
        t["Total"] = t.sum(axis=1)
        t.loc["Total"] = t.sum()
        hojas[f"Edad y tipo {C.nombre_mes(p_)}"] = t.round(0).reset_index().rename(columns={"index": "Tipo de cartera"})
    if not cuentas:
        clases = {"propio"} if propio else {"propio", "terceros", "ajustes"}
        e = ed[(ed["periodo"] == act) & ed["clase"].isin(clases)]
        if tipos:
            e = e[e["tipo"].isin(tipos)]
        t = e.pivot_table(index="componente", columns="edad", values="valor", aggfunc="sum").reindex(
            columns=range(1, 10)).fillna(0.0)
        t.columns = nombres
        t.index = [f"{C.COMPONENTES.get(k, (k,))[0]} ({k})" for k in t.index]
        t["Total"] = t.sum(axis=1)
        t.loc["Total"] = t.sum()
        hojas[f"Edad y componente {m_act}"] = t.round(0).reset_index().rename(columns={"index": "Componente"})
    else:
        a = _filtrar(_mes(leer, emp, act, ver, ("cuenta", "tipo", "clase", "estrato", *cols)), tipos, cuentas)
        b = _filtrar(_mes(leer, emp, ant, ver, ("cuenta", "tipo", *cols)), tipos, cuentas)
        a = a.assign(tipo=a["tipo"].astype(str), clase=a["clase"].astype(str))
        b = b.assign(tipo=b["tipo"].astype(str))
        det = a.rename(columns=dict(zip(cols, [f"{m_act}: {n}" for n in nombres])))
        det[f"Total {m_act}"] = a[cols].sum(axis=1)
        det = det.merge(b.assign(**{f"Total {m_ant}": b[cols].sum(axis=1)})[["cuenta", "tipo", f"Total {m_ant}"]],
                        on=["cuenta", "tipo"], how="outer")
        det = det.rename(columns={"cuenta": "Cuenta contrato", "tipo": "Tipo", "clase": "Clase de uso",
                                  "estrato": "Estrato"}).sort_values(["Cuenta contrato", "Tipo"])
        hojas["Cuentas por edad"] = det
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
        wb = w.book
        cab = wb.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#305496", "border": 1,
                             "text_wrap": True, "valign": "vcenter", "font_name": "Calibri", "font_size": 10})
        num = wb.add_format({"num_format": "#,##0;[Red]-#,##0", "font_name": "Calibri", "font_size": 10})
        pct = wb.add_format({"num_format": "0.0%;[Red]-0.0%", "font_name": "Calibri", "font_size": 10})
        tit = wb.add_format({"bold": True, "font_size": 13, "font_color": "#002060", "font_name": "Calibri"})
        txt = wb.add_format({"font_name": "Calibri", "font_size": 9, "font_color": "#595959"})
        filtro = (f"{emp} · {'componentes propios' if propio else 'todos los componentes'} · tipos: "
                  f"{', '.join(tipos) if tipos else 'todos'}"
                  + (f" · cuentas: {', '.join(map(str, cuentas[:20]))}{'…' if len(cuentas) > 20 else ''}"
                     if cuentas else ""))
        for nombre, d in hojas.items():
            d.to_excel(w, sheet_name=nombre[:31], index=False, startrow=3)
            ws = w.sheets[nombre[:31]]
            ws.hide_gridlines(2)
            ws.write(0, 0, f"Cartera por edades · {nombre}", tit)
            ws.write(1, 0, filtro, txt)
            for j, c_ in enumerate(d.columns):
                ws.write(3, j, c_, cab)
                es_pct = str(c_).startswith("%")
                ancho = 26 if j == 0 else (11 if es_pct else 15)
                ws.set_column(j, j, ancho, pct if es_pct else (num if pd.api.types.is_numeric_dtype(d[c_])
                                                                and c_ not in ("Cuenta contrato", "Estrato")
                                                                else None))
            ws.set_row(3, 32)
            ws.freeze_panes(4, 1)
    return buf.getvalue()


@st.cache_data(show_spinner=False, max_entries=8)
def _edades(_leer, emp: str, ver: str) -> pd.DataFrame:
    e = _leer("SELECT periodo, tipo, edad, componente, servicio, subsidio, intereses FROM cartera_edades "
              "WHERE empresa=?", (emp,))
    e["valor"] = e["servicio"] + e["subsidio"] + e["intereses"]
    e["clase"] = e["componente"].map(lambda c: {True: "propio", False: "terceros", None: "ajustes"}
                                     [C.COMPONENTES.get(c, ("", None))[1]])
    return e


def _estado(a, b, v) -> str:
    if v < -TOL:
        return "Quedó en cero" if abs(a) < TOL else "Pagó o abonó"
    if v > TOL:
        return "Nuevo en el rango" if abs(b) < TOL else "Aumentó"
    return "Sin cambio"


RANGOS_TODOS = {**C.RANGOS, **{f"{C.EDADES[i][0]}": (i,) for i in range(1, 10)}}


@st.cache_data(show_spinner="Calculando usuarios por edad…", max_entries=16)
def _rangos(_leer, emp: str, act: int, ant: int, propio: bool, tipos: tuple, ver: str,
            cuentas: tuple = ()) -> pd.DataFrame:
    """Para cada rango de edad (más de 1, 2, 3 años y cada edad): usuarios y saldos de los dos meses, cuántos
    aumentaron (y cuántos entraron nuevos al rango) y cuántos bajaron."""
    letra = "p" if propio else "t"
    cols = [f"{letra}{i}" for i in range(1, 10)]

    def por_cuenta(per):
        df = _filtrar(_mes(_leer, emp, per, ver, ("cuenta", "tipo", *cols)), tipos, cuentas)
        return df.groupby("cuenta")[cols].sum()
    A, B = por_cuenta(act), por_cuenta(ant)
    A, B = A.align(B, join="outer", fill_value=0.0)
    filas = []
    for n, eds in RANGOS_TODOS.items():
        c = [f"{letra}{i}" for i in eds]
        a, b = A[c].sum(axis=1), B[c].sum(axis=1)
        d = a - b
        filas.append(dict(rango=n, nb=int((b.abs() >= TOL).sum()), na=int((a.abs() >= TOL).sum()),
                          sb=float(b.sum()), sa=float(a.sum()), ns=int((d > TOL).sum()), vs=float(d[d > TOL].sum()),
                          nn=int(((d > TOL) & (b.abs() < TOL)).sum()), nd=int((d < -TOL).sum()),
                          vd=float(d[d < -TOL].sum())))
    return pd.DataFrame(filas)


def _tarjetas_rangos(t3: pd.DataFrame, dias: dict) -> str:
    """Una tarjeta por rango (más de 1, 2 y 3 años): usuarios que aumentaron y que disminuyeron, con su valor."""
    acentos = [E.AMBAR, E.ROJO, "#8b1e1e"]
    tarjetas = []
    for k, f in enumerate(t3.itertuples()):
        neto = f.vs + f.vd
        lado = lambda flecha, titulo, n, v, color, fondo: (  # noqa: E731
            f'<div style="flex:1 1 100px;min-width:0;background:{fondo};border-radius:10px;padding:9px 10px">'
            f'<div style="font-size:12px;font-weight:600;color:{color};letter-spacing:.02em">{flecha} {titulo}</div>'
            f'<div style="font-size:22px;font-weight:700;color:{E.TEXTO};line-height:1.25;'
            f'font-variant-numeric:tabular-nums">{_n(n)}</div>'
            f'<div style="font-size:12px;color:{E.TEXTO_2};margin-bottom:2px">usuarios</div>'
            f'<div style="font-size:17px;font-weight:700;color:{color};font-variant-numeric:tabular-nums">'
            f'{E.millones(v)}</div></div>')
        tarjetas.append(
            f'<div style="min-width:0;background:#fff;border:1px solid {E.BORDE};'
            f'border-left:4px solid {acentos[k % 3]};border-radius:14px;padding:14px 16px">'
            f'<div style="font-size:14px;font-weight:700;color:{E.TEXTO};margin-bottom:10px">{f.rango} '
            f'<span style="font-weight:500;color:{E.TEXTO_2}">· más de {dias[f.rango]} días</span></div>'
            f'<div style="display:flex;flex-wrap:wrap;gap:8px">'
            + lado("▲", "Aumentaron", f.ns, f.vs, E.ROJO, E.ROJO_SUAVE)
            + lado("▼", "Disminuyeron", f.nd, -f.vd, E.VERDE, E.VERDE_SUAVE)
            + f'</div><div style="font-size:12.5px;color:{E.TEXTO_2};margin-top:10px;line-height:1.5">'
            f'Neto <b style="color:{E.ROJO if neto > 0 else E.VERDE}">{"+" if neto > 0 else "−"}'
            f'{E.millones(abs(neto))}</b> · {_n(f.nn)} nuevos en el rango<br>'
            f'Saldo del rango: <b style="color:{E.TEXTO}">{E.millones(f.sa)}</b> ({_n(f.na)} usuarios)</div></div>')
    return (f'<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:14px">'
            f'{"".join(tarjetas)}</div>')


def _tabla_rangos(rg: pd.DataFrame, m_ant: str, m_act: str) -> pd.DataFrame:
    return pd.DataFrame({
        "Rango de edad": rg["rango"], f"Usuarios {m_ant}": rg["nb"], f"Usuarios {m_act}": rg["na"],
        f"Saldo {m_ant}": rg["sb"].round(0), f"Saldo {m_act}": rg["sa"].round(0),
        "Variación": (rg["sa"] - rg["sb"]).round(0), "Usuarios que aumentaron": rg["ns"],
        "De ellos nuevos en el rango": rg["nn"], "Valor del aumento": rg["vs"].round(0),
        "Usuarios que bajaron": rg["nd"], "Valor de la disminución": rg["vd"].round(0)})


def _excel(leer, emp, act, ant, propio, tipos, edades, etiqueta_edad, ver, componentes: pd.DataFrame,
           cuentas: tuple = ()) -> bytes:
    """Excel para el contador: resumen por rangos de edad, detalle por cuenta (con la edad elegida) y componentes."""
    import io
    import xlsxwriter
    m_act, m_ant = C.nombre_mes(act), C.nombre_mes(ant)
    letra = "p" if propio else "t"
    base = _base(leer, emp, act, ant, propio, edades, tipos, ver, cuentas)
    tot = "propio" if propio else "total"
    df = _filtrar(_mes(leer, emp, act, ver, ("cuenta", "tipo", tot, "edad_propio" if propio else "edad",
                                             *[f"{letra}{i}" for i in range(1, 10)])), tipos, cuentas)
    y = df.groupby("cuenta").agg(ta=(tot, "sum"), edad=("edad_propio" if propio else "edad", "max"),
                                 **{f"e{i}": (f"{letra}{i}", "sum") for i in range(1, 10)})
    dfb = _filtrar(_mes(leer, emp, ant, ver, ("cuenta", "tipo", tot)), tipos, cuentas)
    yb = dfb.groupby("cuenta")[tot].sum().rename("tb")
    det = base.set_index("cuenta").join(y).join(yb).fillna({"ta": 0.0, "tb": 0.0, "edad": 0,
                                                             **{f"e{i}": 0.0 for i in range(1, 10)}})
    det = det[(det["a"].abs() >= TOL) | (det["b"].abs() >= TOL)].reset_index()
    det = det.assign(_d=det["a"] - det["b"]).sort_values("_d", ascending=False)
    rg = _tabla_rangos(_rangos(leer, emp, act, ant, propio, tipos, ver, cuentas), m_ant, m_act)

    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"constant_memory": True, "in_memory": True})
    f_tit = wb.add_format({"bold": True, "font_size": 14, "font_color": "#002060", "font_name": "Calibri"})
    f_txt = wb.add_format({"font_name": "Calibri", "font_size": 10})
    f_cab = wb.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#305496", "font_name": "Calibri",
                           "font_size": 10, "text_wrap": True, "valign": "vcenter", "border": 1})
    f_num = wb.add_format({"num_format": "#,##0;[Red]-#,##0", "font_name": "Calibri", "font_size": 10})
    f_ent = wb.add_format({"num_format": "#,##0", "font_name": "Calibri", "font_size": 10})
    f_neg = wb.add_format({"bold": True, "num_format": "#,##0;[Red]-#,##0", "font_name": "Calibri",
                           "font_size": 10, "bg_color": "#B7DEE8"})

    def hoja(nombre, titulo, lineas, cab, filas, anchos, formatos):
        ws = wb.add_worksheet(nombre)
        ws.hide_gridlines(2)
        ws.write(0, 0, titulo, f_tit)
        for i, t in enumerate(lineas, start=1):
            ws.write(i, 0, t, f_txt)
        r0 = len(lineas) + 2
        for j, (c, w) in enumerate(zip(cab, anchos)):
            ws.set_column(j, j, w)
            ws.write(r0, j, c, f_cab)
        ws.set_row(r0, 30)
        for i, fila in enumerate(filas, start=r0 + 1):
            for j, v in enumerate(fila):
                if v is None or (isinstance(v, float) and pd.isna(v)):
                    continue
                fm = formatos[j](fila) if callable(formatos[j]) else formatos[j]
                ws.write(i, j, v, fm)
        ws.freeze_panes(r0 + 1, 1)
        if filas:
            ws.autofilter(r0, 0, r0 + len(filas), len(cab) - 1)
        return ws

    filtros = [f"Compañía: {emp} · {m_act} comparado con {m_ant}",
               f"Componentes: {'propios de la compañía' if propio else 'todos (incluye terceros y ajustes)'}",
               f"Tipos de cartera: {', '.join(tipos) if tipos else 'todos'} · Edad del detalle: {etiqueta_edad}"
               + (f" · Cuentas contrato: {', '.join(map(str, cuentas[:20]))}{'…' if len(cuentas) > 20 else ''}"
                  if cuentas else ""),
               "Aumentaron = su saldo en ese rango de edad subió frente al mes anterior (incluye a quienes entran al "
               "rango porque su deuda envejeció). Bajaron = pagaron, abonaron o quedaron con saldo negativo."]
    filas = rg.astype(object).values.tolist()
    hoja("Resumen por edad", "Usuarios y saldos por rango de edad (deterioro y castigo)", filtros, list(rg.columns),
         filas, [24] + [14] * (len(rg.columns) - 1),
         [lambda f: f_neg if str(f[0]).startswith("Más de") and "año" in str(f[0]) else f_txt]
         + [lambda f, k=k: f_neg if str(f[0]).startswith("Más de") and "año" in str(f[0]) else
            (f_ent if k in (1, 2, 6, 7, 9) else f_num) for k in range(1, len(rg.columns))])

    cab = ["Cuenta contrato", "Clase de uso", "Estrato", f"Saldo {m_ant} ({etiqueta_edad})",
           f"Saldo {m_act} ({etiqueta_edad})", "Variación", "Estado", f"Deuda más antigua ({m_act})",
           *[f"{m_act}: {C.EDADES[i][0]}" for i in range(1, 10)], f"Saldo total {m_ant}", f"Saldo total {m_act}"]
    filas = ([int(r.cuenta), r.clase, None if pd.isna(r.estrato) else int(r.estrato), round(r.b, 2), round(r.a, 2),
              round(r.a - r.b, 2), _estado(r.a, r.b, r.a - r.b),
              C.EDADES[int(r.edad)][0] if r.edad and int(r.edad) in C.EDADES else "Sin saldo",
              *[round(getattr(r, f"e{i}"), 2) for i in range(1, 10)], round(r.tb, 2), round(r.ta, 2)]
             for r in det.itertuples(index=False))
    filas = list(filas)
    hoja("Cuentas", f"Comparativo por cuenta contrato · {etiqueta_edad}",
         filtros[:3] + [f"{len(filas):,} cuentas con saldo en alguno de los dos meses.".replace(",", ".")],
         cab, filas, [14, 30, 8, 16, 16, 16, 18, 18] + [14] * 11,
         [f_ent, f_txt, f_ent, f_num, f_num, f_num, f_txt, f_txt] + [f_num] * 11)

    if componentes is not None and not componentes.empty:
        filas = [[None if (isinstance(v, float) and pd.isna(v)) else v for v in r]
                 for r in componentes.itertuples(index=False)]
        hoja("Componentes", "Componentes de la tarifa (propios y de terceros)", filtros[:3], list(componentes.columns),
             filas, [26, 16, 14] + [16] * (len(componentes.columns) - 3),
             [f_txt, f_txt, f_txt] + [f_num] * (len(componentes.columns) - 3))
    wb.close()
    return buf.getvalue()


# ------------------------------------------------------------------------------------------------ vista
def mostrar(ctx, empresa: str | None = None, zona=None):
    leer = ctx.leer
    zona = zona or st.container()
    with st.container(key="card_cart_filtros"):
        c = st.columns([0.9, 1, 1, 1.1, 1.6], vertical_alignment="bottom")
        if empresa is None:
            empresa = c[0].selectbox("Compañía", list(EMPRESAS), key="ca_empresa",
                                     help=" · ".join(f"{e}: {n}" for e, n in EMPRESAS.items()))
        else:
            c[0].markdown(f"**{empresa}**")
        filtros = st.container()
    meses = leer("SELECT periodo, archivo, filas, usuarios, usuarios_propio, propio, terceros, ajustes, total, fecha "
                 "FROM cartera_meses WHERE empresa=? AND periodo IN (SELECT periodo FROM cartera_archivo WHERE empresa=?) "
                 "ORDER BY periodo", (empresa, empresa))
    if meses.empty:
        zona.html(E.encabezado(f"Cartera de usuarios · {empresa}", EMPRESAS.get(empresa, empresa), []))
        st.info(f"Aún no hay bases de cartera de **{empresa}**. Se cargan mes a mes en Administración → Cartera "
                "(el .txt del sistema comercial, el Excel o el archivo comprimido).",
                icon=":material/info:")
        return
    pers = [int(p) for p in meses["periodo"]]
    ver = f"{len(pers)}-{'-'.join(str(f) for f in meses['fecha'])}"

    zona.html(E.encabezado(f"Cartera de usuarios · {empresa}",
                           f"{EMPRESAS.get(empresa, empresa)} · comparativo de la cartera de un mes contra otro",
                           [(f"Último mes cargado: {C.nombre_mes(pers[-1])}", ""), (f"{len(pers)} meses", "ok")]))
    with filtros:
        c = c[1:]
        act = c[0].selectbox("Mes actual", pers[::-1], key=f"ca_act_{empresa}", format_func=C.nombre_mes)
        previos = [p for p in pers if p < act][::-1] or [p for p in pers if p != act][::-1]
        if not previos:
            st.info("Solo hay un mes cargado: carga otro mes para comparar.", icon=":material/info:")
            return
        ant = c[1].selectbox("Comparar con", previos, key=f"ca_ant_{empresa}_{act}", format_func=C.nombre_mes)
        modo = c[2].segmented_control("Componentes", [PROPIO, TODOS], default=PROPIO, key="ca_modo",
                                      help="Propios: solo los componentes de la tarifa que son ingreso de la compañía. "
                                           "Todos: el total de la factura (como la hoja «Cartera todos componentes» "
                                           "del Power BI).") or PROPIO
        tipos_disp = sorted(set(leer("SELECT DISTINCT tipo FROM cartera_usuarios WHERE empresa=? AND periodo IN (?, ?)",
                                     (empresa, act, ant))["tipo"]))
        tipos = tuple(c[3].multiselect("Tipo de cartera", tipos_disp, key="ca_tipos", placeholder="Todos",
                                       format_func=lambda t: C.TIPOS.get(t, t)))
        c = st.columns([3.9, 1.7], vertical_alignment="bottom")
        cuentas = _cuentas_txt(c[1].text_input(
            "Cuenta contrato", key="ca_cuentas", placeholder="Ej.: 12415890, 10918191",
            help="Una o varias cuentas contrato separadas por coma o espacio. Filtra toda la vista y los Excel."))
        opc_edad = ["Toda la cartera", *C.RANGOS, "Elegir edades"]
        f_edad = c[0].segmented_control("Edad de la cartera", opc_edad, default=opc_edad[0], key="ca_edad",
                                        format_func=lambda o: {"Toda la cartera": "Toda", "Elegir edades": "Elegir…"}
                                        .get(o, o.replace("Más de ", "Más de\u00a0")),
                                        help="Filtra los usuarios que pagaron o aumentaron, el puente, las cuentas y "
                                             "el Excel por la edad de la deuda (como el filtro de edad del Power BI). "
                                             "Más de 1 año = más de 360 días; 2 años = más de 720; 3 años = más de "
                                             "1.080.") or opc_edad[0]
        if f_edad == "Elegir edades":
            elegidas = st.multiselect("Edades", list(C.EDADES), default=[7, 8, 9], key="ca_edades",
                                        format_func=lambda i: C.EDADES[i][0])
            edades = tuple(sorted(elegidas)) or None
            etiqueta_edad = ", ".join(C.EDADES[i][0] for i in edades) if edades else "toda la cartera"
        elif f_edad in C.RANGOS:
            edades, etiqueta_edad = C.RANGOS[f_edad], f_edad.lower()
        else:
            edades, etiqueta_edad = None, "toda la cartera"
    m_act, m_ant = C.nombre_mes(act), C.nombre_mes(ant)
    st.caption(("Solo componentes de la compañía: barrido y limpieza, limpieza urbana, recolección y transporte, "
                "comercialización, CCS no aprovechable y servicios especiales (servicio + subsidio/contribución + "
                "intereses). No incluye aprovechamiento, disposición final, lixiviados, VIAT, CCS aprovechable, "
                "ajuste a la decena ni saldos a favor.") if modo == PROPIO else
               ("Todos los componentes de la factura (TOTAL_CARTERA): incluye los de terceros (aprovechamiento, "
                "disposición final, lixiviados, VIAT, CCS aprovechable), el ajuste a la decena y los saldos a favor."))

    if cuentas:
        st.info(f"Filtrado por {len(cuentas)} cuenta{'s' if len(cuentas) > 1 else ''} contrato: "
                + ", ".join(map(str, cuentas[:12])) + ("…" if len(cuentas) > 12 else ""), icon=":material/filter_alt:")
    r = _comparar(leer, empresa, act, ant, modo == PROPIO, edades, tipos, ver, cuentas)
    cat = r["cat"]
    sal_a, sal_b = float(cat["a"].sum()), float(cat["b"].sum())
    usu_a, usu_b = int(cat["na"].sum()), int(cat["nb"].sum())
    n_baja, n_cero = int(cat.loc["bajo", "n"]), int(cat.loc["salio", "n"])
    v_baja = float((cat.loc["bajo", "a"] - cat.loc["bajo", "b"]) + (cat.loc["salio", "a"] - cat.loc["salio", "b"]))
    n_sube, n_nuevo = int(cat.loc["subio", "n"]), int(cat.loc["nuevo", "n"])
    v_sube = float(cat.loc["subio", "a"] - cat.loc["subio", "b"])
    v_nuevo = float(cat.loc["nuevo", "a"] - cat.loc["nuevo", "b"])
    var = sal_a - sal_b
    sin_previo = int(cat.loc["bajo", "n"] - cat.loc["bajo", "nb"])   # bajaron sin tener saldo: quedaron en negativo

    ed = _edades(leer, empresa, ver)
    clases = {"propio"} if modo == PROPIO else {"propio", "terceros", "ajustes"}
    e2 = ed[ed["periodo"].isin([act, ant]) & ed["clase"].isin(clases)]
    if tipos:
        e2 = e2[e2["tipo"].isin(tipos)]
    rg = _rangos(leer, empresa, act, ant, modo == PROPIO, tipos, ver, cuentas)
    if cuentas:     # por cuenta no hay detalle por componente: la edad sale del saldo de cada cuenta por edad
        por_edad = pd.DataFrame({p_: _por_edad_cuentas(leer, empresa, p_, modo == PROPIO, tipos, cuentas, ver)
                                 for p_ in (ant, act)})
    else:
        por_edad = e2.pivot_table(index="edad", columns="periodo", values="valor", aggfunc="sum").reindex(
            index=range(1, 10), columns=[ant, act]).fillna(0.0)
    vencida = lambda p: float(por_edad.loc[7:9, p].sum())      # más de 360 días

    # ---------- lectura y KPIs
    tono = "ok" if var < 0 else "bad"
    texto = (("" if not edades else f"<i>Edad: {escape(etiqueta_edad)}.</i> ") +
             f"La cartera {'propia ' if modo == PROPIO else ''}pasó de <b>{E.millones(sal_b)}</b> en {m_ant} a "
             f"<b>{E.millones(sal_a)}</b> en {m_act} ({'+' if var > 0 else ''}{E.millones(var)}, "
             f"{E.pct_txt(_pct(sal_a, sal_b))}). <b>{_n(n_baja + n_cero)}</b> usuarios pagaron o abonaron "
             f"{E.millones(-v_baja)} y <b>{_n(n_sube + n_nuevo)}</b> aumentaron su deuda en "
             f"{E.millones(v_sube + v_nuevo)}. La de más de 360 días va de {E.millones(vencida(ant))} a "
             f"{E.millones(vencida(act))}.")
    st.html(E.lectura("Bajó la cartera." if var < 0 else "Subió la cartera.", texto, tono))
    st.html(E.kpis([
        dict(label=f"Cartera {m_act}", valor=E.millones(sal_a), icono="wallet", color=E.AZUL, fondo=E.AZUL_SUAVE,
             delta=E.delta_valor(var, sal_b, a_favor_si_sube=False, sufijo=f"vs {m_ant}"),
             sub=f"{m_ant}: {E.millones(sal_b)}",
             spark=[float(x) for x in meses["propio" if modo == PROPIO else "total"]]
             if not tipos and not edades else None),
        dict(label="Usuarios con saldo" + (f" ({etiqueta_edad})" if edades else ""), valor=_n(usu_a), icono="users", color=E.AZUL, fondo=E.AZUL_SUAVE,
             delta=E.delta_valor(usu_a - usu_b, usu_b, a_favor_si_sube=False, sufijo=f"vs {m_ant}"),
             sub=f"{m_ant}: {_n(usu_b)}"),
        dict(label="Pagaron o abonaron", valor=_n(n_baja + n_cero), icono="check", color=E.VERDE,
             fondo=E.VERDE_SUAVE, acento=E.VERDE,
             sub=f"{E.millones(-v_baja)} menos de saldo · {_n(n_cero)} quedaron en cero"
                 + (f" · {_n(sin_previo)} sin saldo en {m_ant}" if sin_previo else "")),
        dict(label="Aumentaron su deuda", valor=_n(n_sube + n_nuevo), icono="trend", color=E.ROJO,
             fondo=E.ROJO_SUAVE, acento=E.ROJO,
             sub=f"{E.millones(v_sube + v_nuevo)} más · {_n(n_nuevo)} nuevos en cartera"),
        dict(label="Más de 360 días", valor=E.millones(vencida(act)), icono="alert", color=E.AMBAR,
             fondo=E.AMBAR_SUAVE, acento=E.AMBAR,
             delta=E.delta_valor(vencida(act) - vencida(ant), vencida(ant), a_favor_si_sube=False,
                                 sufijo=f"vs {m_ant}"),
             sub=f"{E.pct_txt(vencida(act) / sal_a if sal_a else float('nan'))} de la cartera"),
    ], columnas=5))

    # ---------- deterioro y castigo: usuarios que aumentaron por rango de edad
    with st.container(key="card_cart_det"):
        titulo_tarjeta("Usuarios que aumentaron y disminuyeron en cartera mayor a 1, 2 y 3 años",
                       "Para el cálculo de deterioro y castigo: un usuario «aumentó» si su saldo en ese rango de edad "
                       "subió frente al mes anterior, incluidos los que entran al rango porque su deuda envejeció.")
        t3 = rg[rg["rango"].isin(list(C.RANGOS))]
        dias = {"Más de 1 año": "360", "Más de 2 años": "720", "Más de 3 años": "1.080"}
        st.html(_tarjetas_rangos(t3, dias))
        tr = _tabla_rangos(rg, m_ant, m_act)
        with st.expander("Ver la tabla por rango y por cada edad"):
            st.dataframe(tr, hide_index=True, width="stretch", height=35 * (len(tr) + 1) + 3,
                         column_config={c_: st.column_config.NumberColumn(format="localized") for c_ in tr.columns[1:]})
        excel_det = st.empty()

    # ---------- puente y edades
    c1, c2 = st.columns([1, 1.15])
    with c1, st.container(key="card_cart_puente"):
        titulo_tarjeta(f"De {m_ant} a {m_act}", "Cifras en millones. Pagos = usuarios cuyo saldo bajó; aumentos = "
                                               "usuarios cuyo saldo subió (incluye los que entran a cartera).")
        fig, _ = cascada([
            dict(nombre=f"Cartera {m_ant}", valor=sal_b, tipo="total", color=E.NAVY),
            dict(nombre="Pagos y abonos", valor=v_baja, color=E.VERDE),
            dict(nombre="Aumentos", valor=v_sube, color=E.ROJO),
            dict(nombre="Nuevos en cartera", valor=v_nuevo, color=E.AMBAR),
            dict(nombre=f"Cartera {m_act}", valor=sal_a, tipo="total", color=E.AZUL)], alto=600)
        st.plotly_chart(fig, width="stretch", config={"displayModeBar": False}, key="ca_puente")
    with c2, st.container(key="card_cart_vedad"):
        titulo_tarjeta("Variación por edad", "Negativa (verde) = la cartera de esa edad bajó.")
        filas = [dict(nombre=C.EDADES[i][0], sub=f"{E.millones(por_edad.loc[i, ant])} → "
                                                f"{E.millones(por_edad.loc[i, act])}",
                      celdas=[dict(v=por_edad.loc[i, act] - por_edad.loc[i, ant],
                                   pct=_pct(por_edad.loc[i, act], por_edad.loc[i, ant]),
                                   fav=(por_edad.loc[i, act] - por_edad.loc[i, ant]) < 0)])
                 for i in por_edad.index if abs(por_edad.loc[i, act]) + abs(por_edad.loc[i, ant]) >= 1]
        tot = por_edad.sum()
        filas.append(dict(nombre="Total", clase="total", sub=f"{E.millones(tot[ant])} → {E.millones(tot[act])}",
                          celdas=[dict(v=tot[act] - tot[ant], pct=_pct(tot[act], tot[ant]),
                                       fav=tot[act] < tot[ant])]))
        st.html(E.barras_divergentes(filas, ["Variación"], "Edad"))

    with st.container(key="card_cart_edad"):
        titulo_tarjeta("Cartera por edad", "Meses entre el mes de corte y el mes de la tarifa adeudada (igual que "
                                           "la tabla de edades del Power BI). Cifras en millones; incluye el mismo "
                                           "mes del año anterior si está cargado.")
        nombres = [_partir(C.EDADES[i][0], 12) for i in por_edad.index]
        # mismo mes del año anterior, si está cargado (y no es ya el mes con el que se compara)
        hace_un_anio = act - 100
        series = [(ant, m_ant, "#b4c0d3")]
        if hace_un_anio in pers and hace_un_anio != ant:
            if cuentas:
                por_edad_aa = _por_edad_cuentas(leer, empresa, hace_un_anio, modo == PROPIO, tipos, cuentas, ver)
            else:
                e_aa = ed[(ed["periodo"] == hace_un_anio) & ed["clase"].isin(clases)]
                if tipos:
                    e_aa = e_aa[e_aa["tipo"].isin(tipos)]
                por_edad_aa = e_aa.groupby("edad")["valor"].sum().reindex(range(1, 10)).fillna(0.0)
            series.insert(0, (hace_un_anio, C.nombre_mes(hace_un_anio), "#dfe4ec"))
        else:
            por_edad_aa = None
        series.append((act, m_act, E.AZUL))
        fig = go.Figure()
        for per_s, nombre_s, color_s in series:
            y = por_edad_aa if per_s == hace_un_anio and por_edad_aa is not None else por_edad[per_s]
            fig.add_bar(x=nombres, y=y / 1e6, name=nombre_s, marker_color=color_s,
                        text=[f"{v / 1e6:,.0f}".replace(",", ".") for v in y], textposition="outside",
                        textfont=dict(size=10 if len(series) > 2 else 11,
                                      color=E.AZUL if per_s == act else "#5b6b82"),
                        cliponaxis=False, hovertemplate="%{x}: %{y:,.1f} M<extra>" + nombre_s + "</extra>")
        fig.update_layout(barmode="group", bargap=.14, bargroupgap=.03, uniformtext=dict(minsize=9, mode="show"),
                          yaxis=dict(ticksuffix=" M", tickformat=",.0f"))
        fig = estilo_grafico(fig, 380)
        fig.update_xaxes(tickangle=0, tickfont=dict(size=10))
        st.plotly_chart(fig, width="stretch", config={"displayModeBar": False}, key="ca_graf_edad")
        aa_ = hace_un_anio if (hace_un_anio in pers and hace_un_anio != ant) else None
        st.download_button(
            "Descargar cartera por edades (Excel)", icon=":material/download:", key="ca_xlsx_edades",
            data=lambda: _excel_edades(leer, empresa, act, ant, aa_, modo == PROPIO, tipos, cuentas, ver, ed),
            file_name=f"cartera_edades_{empresa}_{act}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Saldo y usuarios por edad de los meses comparados (y del mismo mes del año anterior si está "
                 "cargado), edad por tipo de cartera, edad por componente y, si filtras cuentas, cada cuenta por "
                 "edad. Respeta los filtros de componentes, tipo de cartera y cuentas.")

    # ---------- por tipo de cartera
    with st.container(key="card_cart_tipo"):
        titulo_tarjeta("Usuarios y saldo por tipo de cartera", "Usuarios = cuentas con saldo en ese tipo.")
        pt = r["por_tipo"]
        t = pt.pivot_table(index="tipo", columns="periodo", values=["v", "n"], aggfunc="sum").fillna(0)
        filas = []
        for tipo in t.index:
            va, vb = t.loc[tipo].get(("v", act), 0.0), t.loc[tipo].get(("v", ant), 0.0)
            na, nb = t.loc[tipo].get(("n", act), 0), t.loc[tipo].get(("n", ant), 0)
            if not (na or nb):
                continue
            filas.append(dict(nombre=C.TIPOS.get(tipo, tipo), sub=f"{_n(nb)} → {_n(na)} usuarios · "
                                                                  f"{E.millones(va)}",
                              celdas=[dict(v=va - vb, pct=_pct(va, vb), fav=va < vb)]))
        filas.sort(key=lambda f: -abs(f["celdas"][0]["v"]))
        st.html(E.barras_divergentes(filas, ["Variación del saldo"], "Tipo"))

    # ---------- componentes
    with st.container(key="card_cart_comp"):
        titulo_tarjeta("Componentes de la tarifa",
                       "Solo los marcados como propios suman en «Componentes propios». Los de terceros se recaudan "
                       "en la factura pero no son ingreso de la compañía.")
        e3 = ed[ed["periodo"].isin([act, ant])]
        if edades:
            e3 = e3[e3["edad"].isin(edades)]
        if tipos:
            e3 = e3[e3["tipo"].isin(tipos)]
        comp = e3.pivot_table(index="componente", columns="periodo", values="valor", aggfunc="sum").reindex(
            columns=[ant, act]).fillna(0.0)
        intereses = e3[e3["periodo"] == act].groupby("componente")["intereses"].sum()
        filas = []
        if cuentas:     # por cuenta solo se guarda propio / terceros / ajustes (no cada componente)
            st.caption("Con filtro de cuentas se muestra el saldo propio, de terceros y los ajustes de esas cuentas "
                       "(el detalle por componente se guarda solo para toda la cartera).")
            sumas = {p_: _filtrar(_mes(leer, empresa, p_, ver, ("cuenta", "tipo", "propio", "terceros", "ajustes")),
                                  tipos, cuentas)[["propio", "terceros", "ajustes"]].sum() for p_ in (ant, act)}
            comp = pd.DataFrame(sumas).reindex(columns=[ant, act])
            intereses = pd.Series(dtype=float)
        for k, (nombre, prop) in ({"propio": ("Componentes propios", True), "terceros": ("Componentes de terceros",
                                   False), "ajustes": ("Ajustes y saldos a favor", None)}.items() if cuentas
                                  else C.COMPONENTES.items()):
            if k not in comp.index:
                continue
            filas.append({"Componente": nombre, "Código": k,
                          "Es de": "La compañía" if prop else ("Terceros" if prop is False else "Ajuste"),
                          f"Saldo {m_ant}": comp.loc[k, ant], f"Saldo {m_act}": comp.loc[k, act],
                          "Variación": comp.loc[k, act] - comp.loc[k, ant],
                          f"De ello intereses ({m_act})": float(intereses.get(k, 0.0))})
        tc = pd.DataFrame(filas, columns=["Componente", "Código", "Es de", f"Saldo {m_ant}", f"Saldo {m_act}",
                                          "Variación", f"De ello intereses ({m_act})"])
        propios = tc["Es de"] == "La compañía"
        totales = [{"Componente": n_, "Código": "", "Es de": "", **tc[m_].sum(numeric_only=True).to_dict()}
                   for n_, m_ in (("Total propios", propios), ("Total terceros y ajustes", ~propios))]
        tc = pd.concat([tc, pd.DataFrame(totales)], ignore_index=True)
        tc[tc.columns[3:]] = tc[tc.columns[3:]].round(0)
        st.dataframe(tc, hide_index=True, width="stretch", height=35 * (len(tc) + 1) + 3, column_config={
            c_: st.column_config.NumberColumn(format="localized") for c_ in tc.columns[3:]})

    # ---------- usuarios por clase de uso
    with st.container(key="card_cart_uso"):
        titulo_tarjeta("Por clase de uso", "Residencial por estrato (incluye multiusuario) y no residencial.")
        pc = r["por_clase"].copy()
        pc["grupo"] = pc["clase"].map(C.grupo_uso)
        g = pc.pivot_table(index="grupo", columns="periodo", values=["v", "n"], aggfunc="sum").fillna(0)
        tg = pd.DataFrame({
            "Clase de uso": g.index,
            f"Usuarios {m_ant}": [int(g.loc[i].get(("n", ant), 0)) for i in g.index],
            f"Usuarios {m_act}": [int(g.loc[i].get(("n", act), 0)) for i in g.index],
            f"Saldo {m_ant}": [g.loc[i].get(("v", ant), 0.0) for i in g.index],
            f"Saldo {m_act}": [g.loc[i].get(("v", act), 0.0) for i in g.index]})
        tg["Variación"] = tg[f"Saldo {m_act}"] - tg[f"Saldo {m_ant}"]
        tg["% var."] = [_pct(a, b) for a, b in zip(tg[f"Saldo {m_act}"], tg[f"Saldo {m_ant}"])]
        tg = tg.sort_values(f"Saldo {m_act}", ascending=False)
        tg[tg.columns[3:6]] = tg[tg.columns[3:6]].round(0)
        st.dataframe(tg, hide_index=True, width="stretch", height=35 * (len(tg) + 1) + 3, column_config={
            **{c_: st.column_config.NumberColumn(format="localized") for c_ in tg.columns[3:6]},
            **{c_: st.column_config.NumberColumn(format="localized") for c_ in tg.columns[1:3]},
            "% var.": st.column_config.NumberColumn(format="percent")})

    # ---------- cuentas
    with st.container(key="card_cart_cuentas"):
        titulo_tarjeta("Cuentas con mayor movimiento", f"Saldo {m_ant} → {m_act}.")

        def tabla(d):
            return pd.DataFrame({"Cuenta": d["cuenta"].astype(str), f"{m_ant}": d["b"].round(0),
                                 f"{m_act}": d["a"].round(0), "Variación": d["d"].round(0),
                                 "Clase de uso": d["clase"]})
        cfg = {c_: st.column_config.NumberColumn(format="localized") for c_ in (m_ant, m_act, "Variación")}
        cfg["Cuenta"] = st.column_config.TextColumn(width="small")
        c1, c2 = st.columns(2)
        c1.markdown("**Mayores pagos o abonos**")
        c1.dataframe(tabla(r["baja"]), hide_index=True, width="stretch", column_config=cfg, height=563)
        c2.markdown("**Mayores aumentos**")
        c2.dataframe(tabla(r["sube"]), hide_index=True, width="stretch", column_config=cfg, height=563)

        b1, b2 = st.columns([2, 1], vertical_alignment="bottom")
        b1.caption("Para ver el detalle de una o varias cuentas, escríbelas en «Cuenta contrato» en los filtros.")
        b2.download_button("Descargar comparativo en Excel", icon=":material/download:",
                           data=lambda: _excel(leer, empresa, act, ant, modo == PROPIO, tipos, edades, etiqueta_edad,
                                               ver, tc, cuentas),
                           file_name=f"cartera_{empresa}_{ant}_vs_{act}.xlsx", key="ca_xlsx2", width="stretch",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        ids = list(cuentas[:300])
        if ids:
            st.markdown("**Detalle de las cuentas filtradas**")
            cols_b = ["cuenta", "tipo", "clase", "estrato", "edad", "propio", "terceros", "ajustes", "total"]
            d = pd.concat([_mes(leer, empresa, p_, ver, tuple(cols_b)).loc[lambda x: x["cuenta"].isin(ids), cols_b]
                           .assign(periodo=p_) for p_ in (ant, act)])
            d = d.sort_values(["cuenta", "periodo", "tipo"])[["periodo", *cols_b]]
            d["tipo"] = d["tipo"].astype(str)
            d["clase"] = d["clase"].astype(str)
            if d.empty:
                st.caption("Esas cuentas no tienen saldo en ninguno de los dos meses.")
            else:
                d["periodo"] = d["periodo"].map(C.nombre_mes)
                d["edad"] = d["edad"].map(lambda i: C.EDADES.get(int(i), ("—",))[0] if i else "Sin saldo")
                d = d.rename(columns={"periodo": "Mes", "cuenta": "Cuenta", "tipo": "Tipo", "clase": "Clase de uso",
                                      "estrato": "Estrato", "edad": "Deuda más antigua", "propio": "Propio",
                                      "terceros": "Terceros", "ajustes": "Ajustes", "total": "Total"})
                d["Cuenta"] = d["Cuenta"].astype(str)
                st.dataframe(d.round({"Propio": 0, "Terceros": 0, "Ajustes": 0, "Total": 0}), hide_index=True,
                             width="stretch", column_config={
                                 c_: st.column_config.NumberColumn(format="localized")
                                 for c_ in ("Propio", "Terceros", "Ajustes", "Total")})

    # ---------- evolución
    if len(pers) > 2:
        with st.container(key="card_cart_evol"):
            titulo_tarjeta("Evolución de la cartera", "Todos los meses cargados · componentes propios y de terceros.")
            x = [C.nombre_mes(p) for p in meses["periodo"]]
            ev = meses
            if cuentas or tipos:      # con filtros: se suma cada mes desde el detalle por cuenta
                filas_ev = []
                for p_ in pers:
                    d_ = _filtrar(_mes(leer, empresa, p_, ver, ("cuenta", "tipo", "propio", "terceros")),
                                  tipos, cuentas)
                    g_ = d_.groupby("cuenta")["propio"].sum()
                    filas_ev.append(dict(propio=d_["propio"].sum(), terceros=d_["terceros"].sum(),
                                         usuarios_propio=int((g_.abs() >= TOL).sum())))
                ev = pd.DataFrame(filas_ev)
            fig = go.Figure()
            fig.add_bar(x=x, y=ev["propio"] / 1e6, name="Propios", marker_color=E.AZUL)
            fig.add_bar(x=x, y=ev["terceros"] / 1e6, name="Terceros", marker_color="#b4c0d3")
            fig.add_scatter(x=x, y=ev["usuarios_propio"], name="Usuarios con saldo propio", yaxis="y2",
                            mode="lines+markers", line=dict(color=E.AMBAR, width=2))
            fig.update_layout(barmode="stack", yaxis=dict(ticksuffix=" M", tickformat=",.0f"),
                              yaxis2=dict(overlaying="y", side="right", showgrid=False, tickformat=",.0f"))
            st.plotly_chart(estilo_grafico(fig, 300), width="stretch", config={"displayModeBar": False},
                            key="ca_evol")

    excel_det.download_button(
        f"Descargar Excel para deterioro y castigo · detalle {etiqueta_edad}", icon=":material/download:",
        data=lambda: _excel(leer, empresa, act, ant, modo == PROPIO, tipos, edades, etiqueta_edad, ver, tc, cuentas),
        file_name=f"cartera_{empresa}_{ant}_vs_{act}_{etiqueta_edad.replace(' ', '_')[:30]}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="ca_xlsx", type="primary",
        help="Hojas: resumen por rangos de edad (más de 1, 2 y 3 años y cada edad), detalle por cuenta contrato con la "
             "edad elegida arriba y su saldo en cada edad, y componentes propios y de terceros.")
    C.liberar_memoria()
    _informe(empresa, m_act, m_ant, modo, sal_a, sal_b, usu_a, usu_b, n_baja, n_cero, v_baja, n_sube, n_nuevo,
             v_sube, v_nuevo, por_edad, act, ant, r["por_tipo"], ed, tipos, sin_previo, cuentas)


def _informe(empresa, m_act, m_ant, modo, sal_a, sal_b, usu_a, usu_b, n_baja, n_cero, v_baja, n_sube, n_nuevo,
             v_sube, v_nuevo, por_edad, act, ant, por_tipo, ed, tipos, sin_previo=0, cuentas=()):
    var = sal_a - sal_b
    general = [
        f"Cartera {'propia ' if modo == PROPIO else ''}de <b>{E.millones(sal_a)}</b>: "
        f"{E.var_html(var)} frente a {m_ant} ({E.pct_txt(_pct(sal_a, sal_b))}).",
        f"Usuarios con saldo: <b>{_n(usu_a)}</b> ({'+' if usu_a >= usu_b else ''}{_n(usu_a - usu_b)} vs {m_ant}).",
        f"<b>{_n(n_baja + n_cero)}</b> usuarios pagaron o abonaron <b>{E.millones(-v_baja)}</b>; {_n(n_cero)} de "
        "ellos quedaron en cero.",
        f"<b>{_n(n_sube)}</b> usuarios aumentaron su deuda en {E.millones(v_sube)} y <b>{_n(n_nuevo)}</b> entraron a "
        f"cartera con {E.millones(v_nuevo)}."]
    d = (por_edad[act] - por_edad[ant])
    edades = [f"<b>{C.EDADES[i][0]}</b>: {E.var_html(v)} (queda en {E.millones(por_edad.loc[i, act])})."
              for i, v in d.reindex(d.abs().sort_values(ascending=False).index).items() if abs(v) >= 1e6][:4]
    pt = por_tipo.pivot_table(index="tipo", columns="periodo", values="v", aggfunc="sum").reindex(
        columns=[ant, act]).fillna(0)
    pt["d"] = pt[act] - pt[ant]
    tipos_txt = [f"<b>{escape(C.TIPOS.get(t, t))}</b>: {E.var_html(v)} (queda en {E.millones(pt.loc[t, act])})."
                 for t, v in pt["d"].reindex(pt["d"].abs().sort_values(ascending=False).index).items()
                 if abs(v) >= 1e6][:3]
    notas = []
    if sin_previo:
        notas.append(f"<b>{_n(sin_previo)}</b> de los usuarios que «pagaron» no tenían saldo en {m_ant}: su saldo "
                     f"quedó negativo en {m_act} (abonos, descuentos o notas a favor). El Power BI también los cuenta "
                     "como pagos.")
    if usu_b and abs(usu_a / usu_b - 1) > 0.3:
        notas.append(f"Los usuarios con saldo cambian {E.pct_txt(usu_a / usu_b - 1)}: revisa el tipo de cartera, "
                     "porque un mismo ajuste aplicado a muchas cuentas mueve el conteo aunque el saldo cambie poco.")
    if modo == PROPIO and not cuentas:
        e = ed[(ed["periodo"] == act) & (ed["clase"] != "propio")]
        if tipos:
            e = e[e["tipo"].isin(tipos)]
        notas.append(f"No se incluyen <b>{E.millones(float(e['valor'].sum()))}</b> de componentes de terceros, "
                     "ajuste a la decena y saldos a favor que también están en la base.")
    titulo = f"Informe de cartera · {empresa}"
    sub = (f"{m_act} vs {m_ant} · {modo.lower()}" + (f" · tipos {', '.join(tipos)}" if tipos else "")
           + (f" · cuentas {', '.join(map(str, cuentas[:10]))}{'…' if len(cuentas) > 10 else ''}" if cuentas else ""))
    secciones = [("En general", general), ("Por edad", edades), ("Por tipo de cartera", tipos_txt),
                 ("Para tener en cuenta", notas)]
    st.html(E.informe(titulo, sub, secciones))
    with st.expander("Copiar el informe en texto (para un correo)", icon=":material/content_copy:"):
        st.code(E.informe_texto(titulo, sub, secciones), language=None, wrap_lines=True)
