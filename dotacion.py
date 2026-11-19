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
from difflib import SequenceMatcher

import pandas as pd

import logica_m as L

CUENTA_SISPO = "DOTACIÓN PERSONAL"
DOTACION, EPPS, ACTIVIDADES = "Dotación", "EPP's", "Actividades"
SECCIONES_COMPARADAS = (DOTACION, EPPS)
SIN_ELEMENTO = "Sin elemento asignado"
SIN_DETALLE = "Sin detalle de SISPO"
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
    """Lee el libro de presupuesto de dotación. Devuelve (filas largas: seccion, elemento, precio, mes, cantidad, valor; info).
    Las celdas con fórmula se leen por su valor guardado; si no lo traen, el valor se calcula como cantidad × precio unitario."""
    import openpyxl
    wb = openpyxl.load_workbook(archivo, data_only=True)
    ws = next((h for h in wb.worksheets if "DOTAC" in _norm(h.title)), wb.worksheets[0])
    col_mes = {}                                                 # mes -> columna de «Cant.»; la de «Valor» es la siguiente
    for fila in ws.iter_rows(min_row=1, max_row=12):
        hallados = {MESES_NOMBRE.index(_norm(c.value)) + 1: c.column for c in fila if c.value is not None and _norm(c.value) in MESES_NOMBRE}
        if len(hallados) >= 3:                                 # el UPA 2026 solo trae septiembre-diciembre
            col_mes = hallados
            break
    if not col_mes:
        raise ValueError("No encontré los meses (ENERO … DICIEMBRE) en los encabezados del libro de dotación.")
    seccion, filas, sin_valor = None, [], 0
    for r in range(1, ws.max_row + 1):
        b, c, d = ws.cell(r, 2).value, ws.cell(r, 3).value, ws.cell(r, 4).value
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
        if seccion is None or not nc or nc in ("ELEMENTO", "CONCEPTO") or _num(d) is None:
            continue
        precio = _num(d)
        elemento = " ".join(str(c).split())
        for mes, cc in col_mes.items():
            cant, val = _num(ws.cell(r, cc).value), _num(ws.cell(r, cc + 1).value)
            if val is None:
                val = (cant or 0.0) * precio
                sin_valor += cant is not None
            filas.append(dict(seccion=seccion, elemento=elemento, precio=precio, mes=mes, cantidad=cant or 0.0, valor=val))
    if not filas:
        raise ValueError("No encontré elementos con precio unitario en el libro de dotación.")
    d = pd.DataFrame(filas)
    anio = next((int(y) for y in re.findall(r"20\d{2}", ws.title)), None)
    info = dict(hoja=ws.title, anio=anio, elementos=int(d["elemento"].nunique()), valor_total=float(d["valor"].sum()), sin_valor=int(sin_valor),
                por_seccion=d.groupby("seccion")["valor"].sum().to_dict())
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
    e["elemento"] = e["elemento"].fillna(SIN_ELEMENTO)
    e["clave"] = e["elemento"].where(e["elemento"] != SIN_ELEMENTO, SIN_ELEMENTO + " · " + e["articulo"])
    ee = e.groupby("clave", as_index=False).agg(ejec_cant=("cantidad", "sum"), ejecutado=("valor", "sum"), movimientos=("valor", "size"))
    ee["seccion"] = ee["clave"].map(lambda k: seccion_de.get(k, SIN_ELEMENTO))
    ee = ee.rename(columns={"clave": "elemento"})
    if sap_mes is not None:                                       # el total es el de SAP: lo que el SISPO no detalla va en una fila aparte
        sispo_mes = ejec.groupby("mes")["valor"].sum().round(2)
        resto = round(sum(sap_mes.get(m, 0.0) - float(sispo_mes.get(m, 0.0)) for m in meses), 2)
        if abs(resto) >= 0.005:
            ee = pd.concat([ee, pd.DataFrame([dict(elemento="Otros movimientos de SAP (SISPO no los detalla)", ejec_cant=0.0, ejecutado=resto,
                                                   movimientos=0, seccion=SIN_DETALLE)])], ignore_index=True)
    t = pp.merge(ee, on=["seccion", "elemento"], how="outer")
    for c in ("ppto_cant", "presupuesto", "ejec_cant", "ejecutado", "movimientos"):
        t[c] = t[c].fillna(0.0)
    t["movimientos"] = t["movimientos"].astype(int)
    t = L.con_metricas(t.assign(concepto=t["elemento"]))
    orden = {DOTACION: 0, EPPS: 1, SIN_ELEMENTO: 2, SIN_DETALLE: 3}
    t["_o"] = t["seccion"].map(orden).fillna(4)
    return t.sort_values(["_o", "ejecutado", "presupuesto"], ascending=[True, False, False]).drop(columns=["_o", "concepto"]).reset_index(drop=True)


def totales(t: pd.DataFrame) -> dict:
    """Presupuesto, ejecutado y % de la tabla comparada (todas las filas, incluidas las «Sin elemento asignado»)."""
    p, e = float(t["presupuesto"].sum()), float(t["ejecutado"].sum())
    return dict(presupuesto=p, ejecutado=e, variacion=e - p, pct=(e / p) if p > 0 else float("nan"),
                sin_elemento=float(t.loc[t["seccion"] == SIN_ELEMENTO, "ejecutado"].sum()),
                sin_detalle=float(t.loc[t["seccion"] == SIN_DETALLE, "ejecutado"].sum()))


def cuadre_sap(ejec: pd.DataFrame, pool_sap: pd.DataFrame, empresa: str, periodo: int) -> dict:
    """Compara en centavos el SISPO de DOTACIÓN PERSONAL de un mes con el genérico «Salida de inventario» de SAP en la misma cuenta
    (por centro de costo). Devuelve sap, sispo y diferencia en pesos; sap=None si SAP no trae el genérico de ese mes."""
    sap = L.sispo_sap_generico(pool_sap, empresa, periodo, CUENTA_SISPO) if len(pool_sap) else pd.DataFrame()
    sispo = float(ejec.loc[ejec["periodo"] == periodo, "valor"].sum())
    if sap is None or sap.empty:
        return dict(sap=None, sispo=sispo, diferencia=None)
    s = float(pd.to_numeric(sap["valor"], errors="coerce").sum())
    return dict(sap=s, sispo=sispo, diferencia=round(s - sispo, 2))


def libro_ejecucion(t: pd.DataFrame, ejec: pd.DataFrame, equiv: pd.DataFrame, sap: pd.DataFrame, meses: list[int]) -> bytes:
    """Excel de la ejecución de dotación: comparación por elemento, detalle del SISPO (con su elemento) y líneas de SAP que el SISPO
    no detalla (movimientos que no son «Salida de inventario»). Solo los meses pedidos."""
    from io import BytesIO
    c = t.rename(columns={"seccion": "Sección", "elemento": "Elemento", "ppto_cant": "Cant. presupuesto", "presupuesto": "Presupuesto",
                          "ejec_cant": "Cant. ejecutada", "ejecutado": "Ejecutado", "movimientos": "Salidas", "saldo": "Saldo",
                          "pct": "% ejecución", "estado": "Estado"})
    det = ejec[ejec["mes"].isin(meses)].merge(equiv[["codigo", "elemento"]], on="codigo", how="left")
    det["elemento"] = det["elemento"].fillna(SIN_ELEMENTO)
    det = det.rename(columns={"periodo": "Periodo", "mes": "Mes", "codigo": "Código", "articulo": "Artículo", "cc_codigo": "Centro de costo",
                              "cc_nombre": "Nombre centro de costo", "cantidad": "Cantidad", "valor": "Valor", "elemento": "Elemento del presupuesto"})
    s = sap[sap["mes"].isin(meses)]
    gen = s["texto"].map(L.norm_texto).str.contains(L.SISPO_GENERICO, regex=False) if "texto" in s and len(s) else pd.Series(False, index=s.index)
    otros = s[~gen]
    cols = [x for x in ("mes", "fecha", "documento", "cc_sap", "cuenta_nombre", "texto", "valor") if x in otros]
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        c.to_excel(w, sheet_name="Presupuesto vs ejecutado", index=False)
        det.to_excel(w, sheet_name="Detalle SISPO", index=False)
        otros[cols].to_excel(w, sheet_name="SAP sin detalle SISPO", index=False)
    return buf.getvalue()
