"""
Dotación y EPP: presupuesto por artículo, ejecución (salidas de almacén) y su cruce.

Archivo de presupuesto (ej. "PRESUPUESTO 2026 - 2027 DOTACION.xlsx"): una hoja por año o por parte del año
("PPTO 2026 ENE - AGO", "PPTO 2026 SEP - DIC", "PPTO 2027"). En cada hoja:
  - columna A: categoría (DOTACION, EPP´S, ACTIVIDADES), con el encabezado SUBRUBRO;
  - columna B: artículo; C: precio unitario; D: cantidad total;
  - desde la E, por cada mes (nombre del mes en la fila de encabezado) dos columnas: Cant. y Valor.
El año sale del nombre de la hoja y los meses del encabezado, así que da igual cómo se reparta el año entre hojas.
"""
import re
import unicodedata

import pandas as pd

MESES_TXT = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
             "septiembre": 9, "setiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}
COLS_PPTO = ["anio", "mes", "categoria", "articulo", "orden", "precio", "cantidad", "valor"]


def _n(x) -> str:
    s = unicodedata.normalize("NFKD", str(x)).encode("ascii", "ignore").decode().lower()
    return " ".join(re.sub(r"[^a-z0-9 ]", " ", s).split())


def categoria_norm(x) -> str:
    """'EPP´S', 'EPPS', 'epp' -> 'EPP'; 'DOTACION' -> 'DOTACIÓN'; 'ACTIVIDADES' igual."""
    t = _n(x)
    if t.startswith("epp"):
        return "EPP"
    if t.startswith("dotac"):
        return "DOTACIÓN"
    if t.startswith("activ"):
        return "ACTIVIDADES"
    return " ".join(str(x).split()).upper() if pd.notna(x) else ""


def _texto(x) -> str:
    return " ".join(str(x).split()) if pd.notna(x) else ""


def leer_presupuesto(archivo) -> tuple[pd.DataFrame, list[str]]:
    """Devuelve (presupuesto, avisos). presupuesto: anio, mes, categoria, articulo, orden, precio, cantidad, valor
    (una fila por artículo y mes con cantidad o valor; los artículos sin ningún valor quedan con una fila en cero)."""
    xls = archivo if isinstance(archivo, pd.ExcelFile) else pd.ExcelFile(archivo)
    filas, avisos = [], []
    for hoja in xls.sheet_names:
        m = re.search(r"(20\d\d)", hoja)
        if not m:
            avisos.append(f"Hoja '{hoja}' omitida: no tiene el año en el nombre.")
            continue
        anio = int(m.group(1))
        df = xls.parse(hoja, header=None)
        f_enc = next((i for i in range(min(10, len(df)))
                      if any(_n(v) == "articulo" for v in df.iloc[i] if pd.notna(v))), None)
        if f_enc is None:
            avisos.append(f"Hoja '{hoja}' omitida: no encontré el encabezado 'Articulo'.")
            continue
        enc = [_n(v) if pd.notna(v) else "" for v in df.iloc[f_enc]]
        c_art = enc.index("articulo")
        c_cat = c_art - 1 if c_art > 0 else None
        c_pre = next((j for j, e in enumerate(enc) if e.startswith("precio")), None)
        meses = {MESES_TXT[e]: j for j, e in enumerate(enc) if e in MESES_TXT}
        if not meses:
            avisos.append(f"Hoja '{hoja}' omitida: no encontré los meses en el encabezado.")
            continue
        cuerpo = df.iloc[f_enc + 1:]
        # fila de subtítulos (Cant. / Valor) debajo del encabezado
        if cuerpo.shape[0] and _n(cuerpo.iloc[0, min(meses.values())]).startswith("cant"):
            cuerpo = cuerpo.iloc[1:]
        for orden, (_, r) in enumerate(cuerpo.iterrows()):
            art = _texto(r.iloc[c_art])
            if not art or _n(art).startswith("total"):
                continue
            cat = categoria_norm(r.iloc[c_cat]) if c_cat is not None else ""
            precio = pd.to_numeric(r.iloc[c_pre], errors="coerce") if c_pre is not None else float("nan")
            precio = None if pd.isna(precio) else float(precio)
            con_valor = False
            for mes, j in meses.items():
                cant = pd.to_numeric(r.iloc[j], errors="coerce")
                val = pd.to_numeric(r.iloc[j + 1], errors="coerce") if j + 1 < len(r) else float("nan")
                cant = 0.0 if pd.isna(cant) else float(cant)
                val = 0.0 if pd.isna(val) else float(val)
                if cant or val:
                    con_valor = True
                    filas.append((anio, mes, cat, art, orden, precio, cant, val))
            if not con_valor:     # el artículo existe en el presupuesto aunque sin valor: queda en cero (para el cruce)
                filas.append((anio, min(meses), cat, art, orden, precio, 0.0, 0.0))
    return pd.DataFrame(filas, columns=COLS_PPTO), avisos


# ------------------------------------------------------------------ ejecución: salidas de almacén
COLS_EJEC = ["anio", "mes", "fecha", "clave", "bodega", "grupo", "cc", "codigo", "articulo", "unidad", "cantidad",
             "costo_unitario", "valor_sin_iva", "valor"]
EXCLUIR = "(No comparar)"          # concepto para artículos que no van contra el presupuesto de dotación


def clave(x) -> str:
    """Nombre para cruzar: sin tildes, mayúsculas, signos ni espacios dobles."""
    return _n(x)


def _numero(s: pd.Series) -> pd.Series:
    """'$ 129.494,40' / '-1,00' / 0.19 -> número."""
    def conv(x):
        if pd.isna(x):
            return 0.0
        if isinstance(x, (int, float)):
            return float(x)
        t = str(x).replace("$", "").replace(" ", "").strip()
        if "," in t:
            t = t.replace(".", "").replace(",", ".")
        try:
            return float(t)
        except ValueError:
            return 0.0
    return s.map(conv)


def _col(df: pd.DataFrame, *nombres):
    norm = {clave(c): c for c in df.columns}
    for n in nombres:
        if clave(n) in norm:
            return norm[clave(n)]
    return None


def leer_ejecucion(archivo) -> tuple[pd.DataFrame, list[str]]:
    """Reporte de movimientos de almacén (kárdex): Cierre, Fecha Movimiento, Articulo, Cantidad Salida, Costo Total
    Con Iva… Se toman las salidas (entregas). Devuelve (ejecucion, avisos)."""
    xls = archivo if isinstance(archivo, pd.ExcelFile) else pd.ExcelFile(archivo)
    avisos = []
    for hoja in xls.sheet_names:
        df = xls.parse(hoja, dtype=object)
        if _col(df, "Articulo") is None or _col(df, "Cantidad Salida") is None:
            continue
        c = lambda *n: _col(df, *n)                               # noqa: E731
        fecha = pd.to_datetime(df[c("Fecha Movimiento", "Fecha")], errors="coerce")
        cierre = df[c("Cierre")].astype(str).str.extract(r"(20\d\d)(\d\d)") if c("Cierre") else None
        anio = pd.to_numeric(cierre[0], errors="coerce") if cierre is not None else fecha.dt.year
        mes = pd.to_numeric(cierre[1], errors="coerce") if cierre is not None else fecha.dt.month
        anio, mes = anio.fillna(fecha.dt.year), mes.fillna(fecha.dt.month)
        txt = lambda col: df[col].map(_texto) if col else ""      # noqa: E731
        sal = _numero(df[c("Cantidad Salida")])
        ent = _numero(df[c("Cantidad Entrada")]) if c("Cantidad Entrada") else 0.0
        out = pd.DataFrame({
            "anio": anio, "mes": mes, "fecha": fecha.dt.strftime("%Y-%m-%d"),
            "clave": txt(c("IdKardex", "Id Kardex")), "bodega": txt(c("Bodega")), "grupo": txt(c("Grupo")),
            "cc": txt(c("Centro Costo", "Centro de Costo")), "codigo": txt(c("Codigo Articulo")),
            "articulo": txt(c("Articulo")), "unidad": txt(c("Unidad Medida")),
            # salidas en positivo (una devolución, entrada en el mismo reporte, resta)
            "cantidad": -sal - ent if (sal <= 0).all() else sal - ent,
            "costo_unitario": _numero(df[c("Costo Unidad Sin Iva")]) if c("Costo Unidad Sin Iva") else 0.0,
            "valor_sin_iva": -_numero(df[c("Costo Total Sin Iva")]) if c("Costo Total Sin Iva") else 0.0,
            "valor": -_numero(df[c("Costo Total Con Iva")]) if c("Costo Total Con Iva") else 0.0,
        })
        out = out[out["anio"].notna() & (out["articulo"] != "")].copy()
        out[["anio", "mes"]] = out[["anio", "mes"]].astype(int)
        if (out["valor"] < 0).mean() > 0.5:          # el reporte trae las salidas en positivo
            out[["valor", "valor_sin_iva"]] = -out[["valor", "valor_sin_iva"]]
        if out["clave"].eq("").all():
            base = out["fecha"] + "|" + out["articulo"] + "|" + out["valor"].round(2).astype(str)
            out["clave"] = base + "#" + out.groupby(base).cumcount().astype(str)
        return out[COLS_EJEC].reset_index(drop=True), avisos
    raise ValueError("No encontré el reporte de almacén: faltan las columnas 'Articulo' y 'Cantidad Salida'.")


def leer_cruce(archivo) -> pd.DataFrame:
    """Tabla de homologación: ARTICULO (como sale en la ejecución) -> CONCEPTO PRESUPUESTO. Devuelve articulo,
    concepto ('' si la tabla lo deja vacío), con todas las filas (también las repetidas, para avisar)."""
    xls = archivo if isinstance(archivo, pd.ExcelFile) else pd.ExcelFile(archivo)
    for hoja in xls.sheet_names:
        df = xls.parse(hoja, dtype=object)
        ca, cc = _col(df, "Articulo"), _col(df, "Concepto Presupuesto", "Concepto", "Presupuesto")
        if ca is None or cc is None:
            continue
        out = pd.DataFrame({"articulo": df[ca].map(_texto), "concepto": df[cc].map(_texto)})
        return out[out["articulo"] != ""].reset_index(drop=True)
    raise ValueError("No encontré la tabla de cruce: faltan las columnas 'ARTICULO' y 'CONCEPTO PRESUPUESTO'.")


# ------------------------------------------------------------------ cruce y comparativo
def asignar(ejec: pd.DataFrame, cruce: pd.DataFrame, ppto: pd.DataFrame) -> pd.DataFrame:
    """Le pone a cada movimiento su concepto del presupuesto y su categoría.
    estado: 'ok' | 'sin_cruce' (el artículo no está en la tabla) | 'sin_concepto' (está pero sin concepto) |
            'sin_presupuesto' (el concepto no existe en el presupuesto) | 'excluido'."""
    e = ejec.copy()
    mapa = {}
    for r in cruce.itertuples():
        mapa.setdefault(clave(r.articulo), r.concepto)          # si hay repetidos, manda el primero (se avisa)
    nombres = ppto.drop_duplicates("articulo")
    nombre_de = dict(zip(nombres["articulo"].map(clave), nombres["articulo"]))
    cat_de = dict(zip(nombres["articulo"].map(clave), nombres["categoria"]))
    k = e["articulo"].map(clave)
    conc = k.map(mapa)
    e["estado"] = "ok"
    e.loc[conc.isna(), "estado"] = "sin_cruce"
    e.loc[conc.notna() & conc.fillna("").eq(""), "estado"] = "sin_concepto"
    e.loc[conc == EXCLUIR, "estado"] = "excluido"
    kc = conc.fillna("").map(clave)
    e.loc[(e["estado"] == "ok") & ~kc.isin(nombre_de), "estado"] = "sin_presupuesto"
    e["concepto"] = kc.map(nombre_de).fillna(conc.fillna(""))
    e["categoria"] = kc.map(cat_de).fillna("SIN HOMOLOGAR")
    sin = e["estado"].isin(["sin_cruce", "sin_concepto"])
    e.loc[sin, "concepto"] = "Sin cruce · " + e.loc[sin, "articulo"]
    e.loc[e["estado"] == "excluido", "categoria"] = "EXCLUIDO"
    return e


def comparativo(ppto: pd.DataFrame, ejec_asig: pd.DataFrame, meses: list[int]) -> pd.DataFrame:
    """Por concepto: categoria, concepto, ppto_valor, ppto_cant, ejec_valor, ejec_cant, estado (en los meses dados).
    Incluye lo ejecutado sin cruce o sin presupuesto, para que el comparativo sume todo lo ejecutado."""
    p = ppto[ppto["mes"].isin(meses)]
    pp = p.groupby(["categoria", "articulo"]).agg(ppto_valor=("valor", "sum"), ppto_cant=("cantidad", "sum")) \
        .reset_index().rename(columns={"articulo": "concepto"})
    e = ejec_asig[ejec_asig["mes"].isin(meses) & (ejec_asig["estado"] != "excluido")]
    ee = e.groupby(["categoria", "concepto"]).agg(ejec_valor=("valor", "sum"), ejec_cant=("cantidad", "sum"),
                                                 estado=("estado", "first")).reset_index()
    out = pp.merge(ee, on=["categoria", "concepto"], how="outer")
    for c in ("ppto_valor", "ppto_cant", "ejec_valor", "ejec_cant"):
        out[c] = out[c].fillna(0.0)
    out["estado"] = out["estado"].fillna("ok")
    out.loc[(out["ejec_valor"] != 0) & (out["ppto_valor"] == 0) & (out["estado"] == "ok"), "estado"] = "sin_presupuesto"
    return out


def decisiones(ejec: pd.DataFrame, cruce: pd.DataFrame, ppto: pd.DataFrame) -> pd.DataFrame:
    """Lo que Financiera debe decidir para que el comparativo quede bien. Columnas: tipo, articulo, detalle, valor,
    accion. Ordenado por valor ejecutado."""
    filas = []
    a = asignar(ejec, cruce, ppto)
    act = a[a["estado"] != "excluido"]
    for (estado, art), g in act.groupby(["estado", "articulo"]):
        v = float(g["valor"].sum())
        if estado == "sin_cruce":
            filas.append(("Artículo sin cruce", art, f"{g['grupo'].iloc[0]} · {len(g)} salidas", v,
                          "Asígnale un concepto del presupuesto o márcalo como 'No comparar'"))
        elif estado == "sin_concepto":
            filas.append(("Cruce sin concepto", art, "Está en la tabla de cruce pero sin concepto", v,
                          "Escribe el concepto del presupuesto o márcalo como 'No comparar'"))
        elif estado == "sin_presupuesto":
            filas.append(("Concepto sin presupuesto", art, f"Va a '{g['concepto'].iloc[0]}', que no tiene presupuesto "
                          "este año", v, "Confirma el concepto o cámbialo por uno que sí tenga presupuesto"))
    rep = cruce.assign(k=cruce["articulo"].map(clave), kc=cruce["concepto"].map(clave))
    for k, g in rep.groupby("k"):
        if g["kc"].nunique() > 1:
            art = g["articulo"].iloc[0]
            v = float(a.loc[a["articulo"].map(clave) == k, "valor"].sum())
            filas.append(("Artículo con dos conceptos", art, "En la tabla va a: " + " / ".join(g["concepto"].unique()),
                          v, f"Se está usando '{g['concepto'].iloc[0]}'. Deja uno solo"))
    ya = {f[1] for f in filas}
    # lo que Financiera ya confirmó a mano en la app no se vuelve a preguntar
    confirmados = set(cruce.loc[cruce["origen"] == "manual", "articulo"].map(clave)) if "origen" in cruce else set()
    otros = act[~act["grupo"].map(clave).str.contains("dotacion|insumo") & ~act["articulo"].isin(ya)
                & ~act["articulo"].map(clave).isin(confirmados)]
    for art, g in otros.groupby("articulo"):
        filas.append(("Grupo distinto a dotación", art, f"Grupo de almacén: {g['grupo'].iloc[0]}",
                      float(g["valor"].sum()), "Revisa si corresponde a dotación o márcalo como 'No comparar'"))
    out = pd.DataFrame(filas, columns=["tipo", "articulo", "detalle", "valor", "accion"])
    return out.sort_values("valor", ascending=False, key=abs).reset_index(drop=True)


def efectos(cmp: pd.DataFrame) -> pd.DataFrame:
    """Desviación por concepto partida en efecto precio y efecto cantidad (solo conceptos con presupuesto y
    ejecución con cantidades): precio = (precio real − precio ppto) × cant. real; cantidad = (cant. real − cant.
    ppto) × precio ppto."""
    c = cmp[(cmp["ppto_cant"] > 0) & (cmp["ejec_cant"] > 0)].copy()
    c["precio_ppto"] = c["ppto_valor"] / c["ppto_cant"]
    c["precio_real"] = c["ejec_valor"] / c["ejec_cant"]
    c["efecto_precio"] = (c["precio_real"] - c["precio_ppto"]) * c["ejec_cant"]
    c["efecto_cantidad"] = (c["ejec_cant"] - c["ppto_cant"]) * c["precio_ppto"]
    return c


# ------------------------------------------------------------------ ejecución contable (rubros 3.1 del presupuesto)
# Parte de la ejecución de dotación y de seguridad en el trabajo no pasa por el almacén: se registra en contabilidad
# directo al costo (exámenes médicos, purificadores, extintores, compras directas…). Esos registros se toman de la
# ejecución del presupuesto (hoja BASE de SAP, ya cargada en Presupuesto) y se suman al análisis de dotación.
ORIGEN_INV = "Movimiento de inventario"
ORIGEN_DIRECTO = "Registro contable directo al costo"
ORIGEN_MANUAL = "Clasificado manualmente"
ORIGEN_PEND = "Registro contable sin clasificar"
SIN_CLASIFICAR = "Registros contables sin clasificar"
_MESES_TXT = r"\b(ene|feb|mar|abr|may|jun|jul|ago|sep|sept|oct|nov|dic|enero|febrero|marzo|abril|mayo|junio|julio|" \
             r"agosto|septiembre|setiembre|octubre|noviembre|diciembre)\b"
_VACIAS = {"serv", "comp", "prov", "arre", "hono", "para", "con", "del", "las", "los", "por", "sas", "ltda", "s a s",
           "provision", "servicio", "servicios", "compra", "factura", "ajuste", "nota", "credito"}


def es_rubro_dotacion(sub_rubro) -> bool:
    """Sub-rubros 3.1 que van al análisis de dotación: 'Dotación y EPP' y 'Actividades de Seguridad (y Salud) en el
    Trabajo'."""
    n = _n(sub_rubro)
    return "dotacion" in n or ("seguridad" in n and ("salud" in n or "trabajo" in n))


def es_inventario(texto, tercero="") -> bool:
    """El registro contable es la salida de almacén (ya está en el kárdex): no se suma dos veces."""
    t, r = _n(texto), _n(tercero)
    # la contrapartida es la cuenta de inventarios ("REPUESTOS CAMIONES") o el texto dice salida de inventario
    return bool(re.search(r"\binv(entario)?\b", t)) or ("repuestos" in r and "camiones" in r)


def patron(tercero, texto) -> str:
    """Llave de una regla: tercero + texto sin fechas, meses ni números (así sirve para los meses siguientes)."""
    t = re.sub(_MESES_TXT, " ", _n(texto))
    t = re.sub(r"\d+", " ", t)
    return f"{_n(tercero)}|{' '.join(t.split()[:4])}"      # las primeras palabras: el resto suele ser la sede o el mes


def _raices(x) -> set:
    return {w[:6] for w in _n(x).split() if len(w) >= 4 and w not in _VACIAS}


def sugerir(texto, conceptos: list[str]) -> str:
    """Concepto del presupuesto que más palabras comparte con el texto del registro ('' si ninguno)."""
    r = _raices(texto)
    mejor, puntos = "", 0
    for c in conceptos:
        p = len(r & _raices(c))
        if p > puntos:
            mejor, puntos = c, p
    return mejor


def contables(sap: pd.DataFrame, clasif: pd.DataFrame, ppto: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    sap: ejecución del presupuesto (anio, mes, fecha, codigo, sub_rubro, cc_nombre, documento, tercero, texto, valor,
         clave). clasif: clasificaciones guardadas (tipo 'registro'|'patron', llave, concepto).
    Devuelve (directos, inventario):
      directos: registros que no pasan por el almacén, con las columnas de la ejecución de dotación + estado,
                concepto, categoria, origen, detalle, patron;
      inventario: los registros contables de las salidas de almacén (para cuadrar contra el kárdex).
    """
    cols = COLS_EJEC + ["estado", "concepto", "categoria", "origen", "detalle", "patron", "sugerido", "sub_rubro",
                        "tercero"]
    if sap.empty:
        return pd.DataFrame(columns=cols), sap
    sap = sap.copy()
    for c in ("texto", "tercero", "sub_rubro", "cc_nombre", "documento", "clave", "fecha"):
        sap[c] = sap[c].fillna("").astype(str)
    s = sap[sap["codigo"].astype(str).eq("3.1") & sap["sub_rubro"].map(es_rubro_dotacion)].copy()
    inv = s[[es_inventario(t, r) for t, r in zip(s["texto"], s["tercero"])]]
    d = s.drop(inv.index).copy()
    if d.empty:
        return pd.DataFrame(columns=cols), inv
    nombres = ppto.drop_duplicates("articulo")
    conceptos = list(nombres["articulo"])
    nombre_de = dict(zip(nombres["articulo"].map(clave), nombres["articulo"]))
    cat_de = dict(zip(nombres["articulo"].map(clave), nombres["categoria"]))
    por_reg = dict(zip(clasif.loc[clasif["tipo"] == "registro", "llave"],
                       clasif.loc[clasif["tipo"] == "registro", "concepto"])) if not clasif.empty else {}
    por_pat = dict(zip(clasif.loc[clasif["tipo"] == "patron", "llave"],
                       clasif.loc[clasif["tipo"] == "patron", "concepto"])) if not clasif.empty else {}
    d["patron"] = [patron(r, t) for r, t in zip(d["tercero"], d["texto"])]
    d["sugerido"] = [sugerir(t, conceptos) for t in d["texto"]]
    manual = d["clave"].map(por_reg)
    regla = d["patron"].map(por_pat)
    conc = manual.fillna(regla)
    d["origen"] = ORIGEN_PEND
    d.loc[conc.notna(), "origen"] = ORIGEN_MANUAL
    auto = conc.isna() & d["sugerido"].ne("")
    d.loc[auto, "origen"] = ORIGEN_DIRECTO
    conc = conc.where(conc.notna(), d["sugerido"].where(d["sugerido"].ne(""), None))
    kc = conc.fillna("").map(clave)
    d["estado"] = "ok"
    d.loc[conc.isna(), "estado"] = "sin_clasificar"
    d.loc[conc == EXCLUIR, "estado"] = "excluido"
    d.loc[(d["estado"] == "ok") & ~kc.isin(nombre_de), "estado"] = "sin_presupuesto"
    d["concepto"] = kc.map(nombre_de).fillna(conc.fillna(""))
    d["categoria"] = kc.map(cat_de).fillna("SIN HOMOLOGAR")
    d.loc[d["estado"] == "sin_clasificar", "concepto"] = SIN_CLASIFICAR
    d.loc[d["estado"] == "excluido", "categoria"] = "EXCLUIDO"
    d["detalle"] = d["texto"]
    d["articulo"] = d["concepto"].where(d["estado"] != "sin_clasificar", d["texto"]).astype(str) + " · contable"
    d["bodega"] = "(Contable)"
    d["grupo"] = d["sub_rubro"]
    d["cc"] = d["cc_nombre"]
    d["codigo"] = d["documento"]
    d["unidad"] = ""
    d["cantidad"] = 0.0                    # los registros contables no traen cantidades
    d["costo_unitario"] = 0.0
    d["valor_sin_iva"] = d["valor"]
    for c in cols:
        if c not in d:
            d[c] = ""
    return d[cols].reset_index(drop=True), inv
