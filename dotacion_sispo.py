"""Dotación y EPP's (Talento humano): presupuesto por elemento vs ejecución tomada del SISPO.

Lógica pura (sin Streamlit). Presupuesto: el libro «3.1 DOTACION Y EPPS» (tres bloques: DOTACION, EPP'S y ACTIVIDADES, con cantidad y valor por
mes). Ejecución: las salidas de inventario del SISPO de la cuenta «DOTACIÓN PERSONAL» (el mismo filtro de Mantenimiento: costo negativo). Esa
cuenta cuadra en centavos, centro de costo por centro de costo, con el genérico «Salida de inventario» que SAP registra en la misma cuenta.
Cada artículo del SISPO se lleva al elemento presupuestado por una equivalencia (código -> elemento) que se guarda en `dot_map`; mientras
un artículo no tenga elemento queda como «Sin elemento asignado» (cuenta en el total, no se inventa su elemento).
ACTIVIDADES (exámenes médicos, capacitaciones, etc.) son servicios que no pasan por inventario: se cargan pero no se comparan con SISPO.
"""
import re
import unicodedata

import pandas as pd

import logica_m as L

CUENTA_SISPO = "DOTACIÓN PERSONAL"
DOTACION, EPPS, ACTIVIDADES = "Dotación", "EPP's", "Actividades"
SECCIONES_COMPARADAS = (DOTACION, EPPS)
SIN_ELEMENTO = "Sin elemento asignado"
OTROS = "Otros (sin clasificar)"                    # en Dotación: artículos sin elemento y movimientos de SAP que el SISPO no detalla
CODIGO_SAP, SUBRUBRO_SAP = "3.1", "Dotación y epp´s"      # SAP: código 3.1 y BASE GASTOS EXPLOTACIO (sub_rubro)
MESES_NOMBRE = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]

TABLAS = """
CREATE TABLE IF NOT EXISTS dot_ppto(empresa TEXT, anio INT, seccion TEXT, elemento TEXT, precio DOUBLE PRECISION, mes INT,
    cantidad DOUBLE PRECISION, valor DOUBLE PRECISION);
CREATE TABLE IF NOT EXISTS dot_map(empresa TEXT, codigo TEXT, articulo TEXT, elemento TEXT, usuario TEXT, fecha TEXT)
"""


def _norm(x) -> str:
    t = unicodedata.normalize("NFKD", str(x if x is not None else "")).encode("ascii", "ignore").decode().upper()
    return re.sub(r"\s+", " ", re.sub(r"[^A-Z0-9 ]", " ", t)).strip()


def _num(x):
    try:
        v = float(x)
    except (TypeError, ValueError):
        return None
    return None if v != v else v


# ================================================================== presupuesto
def leer_ppto(archivo) -> tuple[pd.DataFrame, dict]:
    """Lee el libro de presupuesto de dotación. Devuelve (filas largas: seccion, elemento, precio, mes, cantidad, valor, anio_rel; info).
    `anio_rel` es 0 para el primer bloque de meses del libro, 1 para el siguiente año, etc.
    Las celdas con fórmula se leen por su valor guardado; si no lo traen, el valor se calcula como cantidad × precio unitario."""
    import openpyxl
    wb = openpyxl.load_workbook(archivo, data_only=True)
    ws = next((h for h in wb.worksheets if "DOTAC" in _norm(h.title)), wb.worksheets[0])
    meses_cols, cab_fila = [], None                              # [(columna de «Cant.», mes)] de la fila de encabezados; la de «Valor» es la siguiente
    for fila in ws.iter_rows(min_row=1, max_row=12):
        hallados = [(c.column, MESES_NOMBRE.index(_norm(c.value)) + 1) for c in fila if c.value is not None and _norm(c.value) in MESES_NOMBRE]
        if len(hallados) >= 3:                                 # el UPA 2026 solo trae septiembre-diciembre
            meses_cols, cab_fila = hallados, fila
            break
    if not meses_cols:
        raise ValueError("No encontré los meses (ENERO … DICIEMBRE) en los encabezados del libro de dotación.")
    # Un libro puede traer varios años uno al lado del otro (p. ej. sep-dic 2026 y luego ene-dic 2027): el bloque nuevo empieza
    # cuando el mes no sube. Cada bloque tiene su propia columna «PRECIO UNITARIO» (la más cercana a su izquierda; por defecto la D).
    # Una columna «PROYECTO» no estorba: cada fila es un elemento de un proyecto y sus cantidades y valores se suman por elemento.
    precios = [c.column for c in cab_fila if c.value is not None and _norm(c.value) == "PRECIO UNITARIO"]
    bloques, previo = [], None
    for col, mes in meses_cols:
        if previo is None or mes <= previo:
            bloques.append(dict(meses={}, precio=max([x for x in precios if x < col], default=4)))
        bloques[-1]["meses"][mes] = col
        previo = mes
    seccion, filas, sin_valor = None, [], 0
    for r in range(1, ws.max_row + 1):
        b, c = ws.cell(r, 2).value, ws.cell(r, 3).value
        nb, nc = _norm(b), _norm(c)
        if nb.startswith("TOTAL") or nc.startswith("TOTAL"):
            continue
        if nb.startswith("DOTACION") and nc not in ("ELEMENTO", "CONCEPTO"):
            seccion = DOTACION
        elif nb.startswith("EPP"):
            seccion = EPPS
            continue                                             # esa fila es el encabezado del bloque
        elif nb.startswith("ACTIVIDADES"):
            seccion = ACTIVIDADES
            continue
        if seccion is None or not nc or nc in ("ELEMENTO", "CONCEPTO"):
            continue
        elemento = " ".join(str(c).split())
        for k, bl in enumerate(bloques):
            precio = _num(ws.cell(r, bl["precio"]).value)
            if precio is None:
                continue
            for mes, cc in bl["meses"].items():
                cant, val = _num(ws.cell(r, cc).value), _num(ws.cell(r, cc + 1).value)
                if val is None:
                    val = (cant or 0.0) * precio
                    sin_valor += cant is not None
                filas.append(dict(seccion=seccion, elemento=elemento, precio=precio, mes=mes, cantidad=cant or 0.0, valor=val, anio_rel=k))
    if not filas:
        raise ValueError("No encontré elementos con precio unitario en el libro de dotación.")
    d = pd.DataFrame(filas)
    anio = next((int(y) for y in re.findall(r"20\d{2}", ws.title)), None)
    info = dict(hoja=ws.title, anio=anio, elementos=int(d["elemento"].nunique()), valor_total=float(d["valor"].sum()), sin_valor=int(sin_valor),
                por_seccion=d.groupby("seccion")["valor"].sum().to_dict(), bloques=len(bloques),
                por_bloque={int(k): g.groupby("seccion")["valor"].sum().to_dict() for k, g in d.groupby("anio_rel")})
    return d, info


# ================================================================== ejecución desde el SISPO
def ejecucion_sispo(d: pd.DataFrame) -> pd.DataFrame:
    """Salidas de inventario de la cuenta DOTACIÓN PERSONAL: periodo (AAAAMM), mes, codigo, articulo, centro de costo, cantidad y valor
    (positivo = costo). Usa el mismo filtro de Mantenimiento (`L.sispo_filtrar`): costo negativo, una línea por IdKardex."""
    neg, _ = L.sispo_filtrar(d, (CUENTA_SISPO,))
    if neg.empty:
        return pd.DataFrame(columns=["periodo", "mes", "codigo", "articulo", "cc_codigo", "cc_nombre", "cantidad", "valor"])
    f = pd.to_datetime(neg["fecha"], errors="coerce")
    cierre = neg["cierre"].where(neg["cierre"].astype(str).str.fullmatch(r"\d{6}"), f.dt.strftime("%Y%m").fillna("")) if "cierre" in neg else f.dt.strftime("%Y%m")
    out = pd.DataFrame({"periodo": pd.to_numeric(cierre, errors="coerce"), "codigo": neg["codigo"].astype(str).str.strip(),
                        "articulo": neg["articulo"].astype(str).str.strip(), "cc_codigo": neg["cc_codigo"], "cc_nombre": neg["cc_nombre"],
                        "cantidad": pd.to_numeric(neg["cantidad"], errors="coerce").abs(), "valor": -pd.to_numeric(neg["costo_con_iva"], errors="coerce")})
    out = out.dropna(subset=["periodo"])
    out["periodo"] = out["periodo"].astype(int)
    out["mes"] = out["periodo"] % 100
    return out.reset_index(drop=True)


# ================================================================== equivalencia artículo SISPO -> elemento del presupuesto
def sugerir(articulo: str, elementos: list[str]) -> str | None:
    """Elemento presupuestado que corresponde a un artículo del SISPO SOLO si es evidente: mismo nombre, o todas las palabras de uno están en
    el otro (y es el único que cumple). Si hay duda o varios candidatos, None: lo asigna una persona."""
    a = set(_norm(articulo).split())
    if not a:
        return None
    exactos = [e for e in elementos if _norm(e) == _norm(articulo)]
    if len(exactos) == 1:
        return exactos[0]
    cont = [e for e in elementos if len(a) >= 2 and len(set(_norm(e).split())) >= 2
            and (a <= set(_norm(e).split()) or set(_norm(e).split()) <= a)]
    return cont[0] if len(cont) == 1 else None


def equivalencias(articulos: pd.DataFrame, elementos: list[str], guardadas: dict) -> pd.DataFrame:
    """Por cada artículo (codigo, articulo) su elemento: el guardado manda; si no hay, el sugerido evidente; si no, SIN_ELEMENTO.
    Devuelve codigo, articulo, elemento, origen («Guardado» | «Sugerido» | «Pendiente»)."""
    filas = []
    for r in articulos.drop_duplicates("codigo").itertuples():
        g = guardadas.get(r.codigo)
        if g:
            filas.append((r.codigo, r.articulo, g, "Guardado"))
            continue
        s = sugerir(r.articulo, elementos)
        filas.append((r.codigo, r.articulo, s or SIN_ELEMENTO, "Sugerido" if s else "Pendiente"))
    return pd.DataFrame(filas, columns=["codigo", "articulo", "elemento", "origen"])


# ================================================================== comparación presupuesto vs ejecución
def sap_filas(pool: pd.DataFrame, empresa: str, anio: int) -> pd.DataFrame:
    """Líneas de SAP de dotación: código 3.1 y BASE GASTOS EXPLOTACIO «Dotación y epp´s» (incluye anulados y reversos)."""
    if pool is None or pool.empty:
        return pd.DataFrame(columns=["mes", "valor"])
    k = lambda x: re.sub(r"[^A-Z0-9]", "", _norm(x))
    es = ((pool["empresa"] == empresa) & (pd.to_numeric(pool["anio"], errors="coerce") == anio)
          & (pool["codigo"].astype(str).str.strip() == CODIGO_SAP) & (pool["sub_rubro"].map(k) == k(SUBRUBRO_SAP)))
    return pool[es].assign(mes=lambda d: pd.to_numeric(d["mes"], errors="coerce"))


def sap_por_mes(pool: pd.DataFrame, empresa: str, anio: int) -> dict[int, float]:
    """Ejecución OFICIAL de dotación por mes: SAP código 3.1 y BASE GASTOS EXPLOTACIO «Dotación y epp´s», neta."""
    f = sap_filas(pool, empresa, anio)
    return {int(m): round(float(v), 2) for m, v in f.groupby("mes")["valor"].sum().items()} if not f.empty else {}


def comparar(ppto: pd.DataFrame, ejec: pd.DataFrame, equiv: pd.DataFrame, meses: list[int], sap_mes: dict | None = None) -> pd.DataFrame:
    """Presupuesto vs ejecutado por elemento en los `meses` pedidos (solo DOTACIÓN y EPP's). Las salidas sin elemento aparecen como
    «Sin elemento asignado» (una fila por artículo) para que el total cuadre con el SISPO. Columnas: seccion, elemento, ppto_cant, presupuesto,
    ejec_cant, ejecutado, movimientos, más saldo, pct y estado."""
    p = ppto[ppto["seccion"].isin(SECCIONES_COMPARADAS) & ppto["mes"].isin(meses)]
    pp = p.groupby(["seccion", "elemento"], as_index=False).agg(ppto_cant=("cantidad", "sum"), presupuesto=("valor", "sum"))
    seccion_de = p.drop_duplicates("elemento").set_index("elemento")["seccion"].to_dict()
    e = ejec[ejec["mes"].isin(meses)].merge(equiv[["codigo", "elemento"]], on="codigo", how="left")
    e["elemento"] = e["elemento"].fillna(SIN_ELEMENTO).replace(SIN_ELEMENTO, OTROS)
    ee = e.groupby("elemento", as_index=False).agg(ejec_cant=("cantidad", "sum"), ejecutado=("valor", "sum"), movimientos=("valor", "size"))
    ee["seccion"] = ee["elemento"].map(lambda k: DOTACION if k == OTROS else seccion_de.get(k, DOTACION))
    if sap_mes is not None:                                       # el total es el de SAP: lo que el SISPO no detalla también va en «Otros»
        sispo_mes = ejec.groupby("mes")["valor"].sum().round(2)
        resto = round(sum(sap_mes.get(m, 0.0) - float(sispo_mes.get(m, 0.0)) for m in meses), 2)
        if abs(resto) >= 0.005:
            ee = pd.concat([ee, pd.DataFrame([dict(elemento=OTROS, ejec_cant=0.0, ejecutado=resto, movimientos=0, seccion=DOTACION)])], ignore_index=True)
        ee = ee.groupby(["seccion", "elemento"], as_index=False).agg(ejec_cant=("ejec_cant", "sum"), ejecutado=("ejecutado", "sum"), movimientos=("movimientos", "sum"))
    t = pp.merge(ee, on=["seccion", "elemento"], how="outer")
    for c in ("ppto_cant", "presupuesto", "ejec_cant", "ejecutado", "movimientos"):
        t[c] = t[c].fillna(0.0)
    t["movimientos"] = t["movimientos"].astype(int)
    t = L.con_metricas(t.assign(concepto=t["elemento"]))
    t["_o"] = t["seccion"].map({DOTACION: 0, EPPS: 1}).fillna(2)
    t["_otros"] = (t["elemento"] == OTROS).astype(int)               # «Otros» siempre al final de su sección
    return t.sort_values(["_o", "_otros", "ejecutado", "presupuesto"], ascending=[True, True, False, False]).drop(columns=["_o", "_otros", "concepto"]).reset_index(drop=True)


def totales(t: pd.DataFrame) -> dict:
    """Presupuesto, ejecutado y % de la tabla comparada. `otros` = ejecutado en «Otros (sin clasificar)» (ya incluido en el total)."""
    p, e = float(t["presupuesto"].sum()), float(t["ejecutado"].sum())
    return dict(presupuesto=p, ejecutado=e, variacion=e - p, pct=(e / p) if p > 0 else float("nan"),
                otros=float(t.loc[t["elemento"] == OTROS, "ejecutado"].sum()))


def cuadre_sap(ejec: pd.DataFrame, pool_sap: pd.DataFrame, empresa: str, periodo: int) -> dict:
    """Compara en centavos el SISPO de DOTACIÓN PERSONAL de un mes con el genérico «Salida de inventario» de SAP en la misma cuenta
    (por centro de costo). Devuelve sap, sispo y diferencia en pesos; sap=None si SAP no trae el genérico de ese mes."""
    sap = L.sispo_sap_generico(pool_sap, empresa, periodo, CUENTA_SISPO) if len(pool_sap) else pd.DataFrame()
    sispo = float(ejec.loc[ejec["periodo"] == periodo, "valor"].sum())
    if sap is None or sap.empty:
        return dict(sap=None, sispo=sispo, diferencia=None)
    s = float(pd.to_numeric(sap["valor"], errors="coerce").sum())
    return dict(sap=s, sispo=sispo, diferencia=round(s - sispo, 2))


def libro_ejecucion(ppto: pd.DataFrame, ejec: pd.DataFrame, equiv: pd.DataFrame, sap: pd.DataFrame, sap_mes: dict, meses: list[int],
                    titulo: str = "", rotulo: str = "PRESUPUESTO") -> bytes:
    """Excel de la ejecución de dotación con tres hojas:
    - «Resumen»: presupuesto (`rotulo`) de todos los meses del libro y EJECUCION mes a mes (solo los `meses` pedidos), una fila por grupo
      (Dotación y EPP's), sin detalle por elemento. «Otros (sin clasificar)» va dentro de Dotación. Totales con fórmulas.
    - «Dotación» y «EPPs»: los movimientos que componen la ejecución de cada grupo: las salidas del SISPO (con su elemento) y, en Dotación,
      las líneas de SAP que el SISPO no detalla. Si SAP y SISPO difieren en la «Salida de inventario», la diferencia aparece como una fila
      «Ajuste» para que los movimientos sumen la ejecución.
    Debajo de los TOTALES hay un control: lo que suman las dos hojas de movimientos y su diferencia con la ejecución."""
    from io import BytesIO
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter as col

    p = ppto[ppto["seccion"].isin(SECCIONES_COMPARADAS)]
    meses_col = sorted(set(int(m) for m in p["mes"].unique()) | set(int(m) for m in meses))
    p_mes = p.groupby(["seccion", "elemento", "mes"])["valor"].sum()
    e_mes = {}                                                    # (sección, elemento, mes) -> ejecutado, igual que la pantalla (`comparar`)
    for m in meses:
        tm = comparar(ppto, ejec, equiv, [m], sap_mes)
        for r in tm.itertuples():
            e_mes[(r.seccion, r.elemento, int(m))] = float(r.ejecutado)
    filas = []
    for sec, nombre in ((DOTACION, "Dotación"), (EPPS, "EPP's")):
        elems = {k[1] for k in p_mes.index if k[0] == sec} | {k[1] for k in e_mes if k[0] == sec}
        hijos = []
        for el in elems:
            pv = [float(p_mes.get((sec, el, m), 0.0)) for m in meses_col]
            ev = [float(e_mes.get((sec, el, m), 0.0)) if m in meses else 0.0 for m in meses_col]
            if any(abs(v) >= 0.005 for v in pv + ev):
                hijos.append(dict(nombre=el, p=pv, e=ev, nivel=1, _o=(el == OTROS, -sum(ev), -sum(pv), el)))
        filas.append(dict(nombre=nombre, p=[sum(h["p"][i] for h in hijos) for i in range(len(meses_col))],
                          e=[sum(h["e"][i] for h in hijos) for i in range(len(meses_col))], nivel=0))   # una fila por grupo: Dotación y EPP's

    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen"
    ws.cell(row=1, column=1, value=titulo or "Dotación y EPP's: presupuesto vs ejecución").font = Font(bold=True, size=13)
    ws.cell(row=2, column=1, value="Ejecución: " + ", ".join(MESES_NOMBRE[m - 1].capitalize() for m in meses)).font = Font(italic=False)
    info = L.hoja_resumen_mensual(ws, filas, [MESES_NOMBRE[m - 1] for m in meses_col], rotulo, fila_enc=4)

    sis = ejec[ejec["mes"].isin(meses)].merge(equiv[["codigo", "elemento"]], on="codigo", how="left")
    sis["elemento"] = sis["elemento"].fillna(SIN_ELEMENTO).replace(SIN_ELEMENTO, OTROS)
    seccion_de = p.drop_duplicates("elemento").set_index("elemento")["seccion"].to_dict()
    sis["seccion"] = sis["elemento"].map(lambda k: DOTACION if k == OTROS else seccion_de.get(k, DOTACION))
    cols = ["Mes", "Fuente", "Sección", "Elemento", "Código", "Artículo / texto SAP", "Centro de costo", "Nombre centro de costo", "Documento SAP",
            "Fecha", "Cantidad", "Valor"]
    mov = []
    for r in sis.sort_values(["mes", "seccion", "elemento", "articulo"]).itertuples():
        mov.append([f"{MESES_NOMBRE[int(r.mes) - 1].capitalize()}", "SISPO", r.seccion, r.elemento, r.codigo, r.articulo, r.cc_codigo, r.cc_nombre,
                    None, None, float(r.cantidad), float(r.valor)])
    s_ = sap[sap["mes"].isin(meses)] if len(sap) else sap
    gen = s_["texto"].map(L.norm_texto).str.contains(L.SISPO_GENERICO, regex=False) if len(s_) and "texto" in s_ else pd.Series(False, index=s_.index)
    for r in s_[~gen].sort_values(["mes", "fecha"] if "fecha" in s_ else ["mes"]).to_dict("records"):
        f = pd.to_datetime(r.get("fecha"), errors="coerce")
        mov.append([MESES_NOMBRE[int(r["mes"]) - 1].capitalize(), "SAP (sin detalle SISPO)", DOTACION, OTROS, None, r.get("texto"), r.get("cc_sap"), None,
                    r.get("documento"), None if pd.isna(f) else f.to_pydatetime(), None, float(r["valor"])])
    for m in meses:                                               # diferencia entre la «Salida de inventario» de SAP y el SISPO: se muestra, no se oculta
        total_m = sum(v for (sc, el, mm), v in e_mes.items() if mm == m)
        suma_m = sum(x[-1] for x in mov if x[0] == MESES_NOMBRE[m - 1].capitalize())
        if abs(total_m - suma_m) >= 0.005:
            mov.append([MESES_NOMBRE[m - 1].capitalize(), "Ajuste SAP vs SISPO", DOTACION, OTROS, None,
                        "Diferencia entre la «Salida de inventario» de SAP y el SISPO de DOTACIÓN PERSONAL", None, None, None, None, None, round(total_m - suma_m, 2)])
    cols.remove("Sección")                                        # el nombre de la hoja ya dice la sección
    anchos = {"Mes": 14, "Fuente": 22, "Elemento": 36, "Código": 14, "Artículo / texto SAP": 56, "Centro de costo": 14,
              "Nombre centro de costo": 26, "Documento SAP": 16, "Fecha": 12, "Cantidad": 10, "Valor": 16}
    refs = []
    for sec, hoja in ((DOTACION, "Dotación"), (EPPS, "EPPs")):    # los movimientos, separados: Dotación por un lado y solo EPP's por otro
        filas_h = [x[:2] + x[3:] for x in mov if x[2] == sec]
        L.escribir_hoja_movimientos(wb, hoja, cols, filas_h, {"Fecha": "dd/mm/yyyy", "Cantidad": "#,##0", "Valor": '#,##0.00'}, anchos)
        refs.append(f"'{hoja}'!{col(len(cols))}2:{col(len(cols))}{max(len(filas_h) + 1, 2)}")
    rc = info["fila_total"] + 2
    ref = ",".join(refs)
    ce = col(info["col_ejec_total"])
    ws.cell(row=rc, column=1, value="Control · suma de las hojas de movimientos")
    ws.cell(row=rc, column=info["col_ejec_total"], value=f"=SUM({ref})").number_format = L.FORMATO_DINERO
    ws.cell(row=rc + 1, column=1, value="Control · diferencia con la ejecución (debe ser 0)")
    ws.cell(row=rc + 1, column=info["col_ejec_total"], value=f"=ROUND({ce}{rc}-{ce}{info['fila_total']},2)").number_format = L.FORMATO_DINERO
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()
