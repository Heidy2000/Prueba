"""
Flujo de caja calculado desde los movimientos de bancos de SAP

- Movimientos: el reporte de SAP (columnas A a T de la hoja Journal / SAP del libro de flujo). Cada movimiento se
  lleva a un concepto de la caja ("Subrubro caja"). El tercero sale del código de 6 cifras de 'Texto cab.documento'.
- Estructura: las filas de la hoja 'Flujo de caja' del libro, con sus fórmulas traducidas a términos
  (fila, mes anterior/mismo mes, signo). Las filas de concepto suman los movimientos de ese concepto en el mes;
  las de totales y saldos suman otras filas, como en Excel.
- Ajustes: lo que el libro suma a mano y no viene de SAP (remuneración y GMF de Enel, facturación conjunta,
  restringido, saldo inicial de enero…). Se toman del libro (valor del archivo − valor calculado) y se pueden
  editar mes a mes en la app.
- Clasificación: aprende de los movimientos ya clasificados (del libro o decididos en la app) y la aplica a los
  nuevos, de la regla más específica a la más general; lo que no resuelve con confianza queda pendiente.
"""
import json
import re
import unicodedata

import numpy as np
import pandas as pd

import cargas

ANULADO = "Anulado"
PENDIENTE = "Pendiente por clasificar"
ORIGENES_ENTRENAMIENTO = ("validado", "manual")


def _n(x) -> str:
    s = unicodedata.normalize("NFKD", str(x)).encode("ascii", "ignore").decode().lower()
    return " ".join(re.sub(r"[^a-z0-9 ]", " ", s).split())


def clave_concepto(x) -> str:
    """Los SUMIFS de Excel no distinguen mayúsculas ni espacios de los extremos."""
    return " ".join(str(x).lower().split()) if pd.notna(x) else ""


# ------------------------------------------------------------------ movimientos de SAP
# nombre interno: encabezados posibles (normalizados, sin tildes ni signos)
COLUMNAS_SAP = {
    "clase": ["clase de documento"], "documento": ["no documento", "n documento", "numero documento"],
    "referencia": ["referencia"], "fecha": ["fe contabilizacion", "fecha contabilizacion"],
    "cuenta": ["cuenta", "cuenta de mayor"], "valor": ["importe en moneda local"], "texto": ["texto"],
    "cc": ["centro de coste"], "cab": ["texto cab documento"], "acreedor": ["acreedor"], "deudor": ["deudor"],
    "ref1": ["clave referencia 1"], "ref3": ["clave referencia 3"], "compensacion": ["doc compensacion"],
    "anulado": ["anulado con"], "manual": ["modificacion manual"], "sub": ["subrubro caja", "tipo 3"],
}
OBLIGATORIAS = ["clase", "documento", "fecha", "cuenta", "valor"]


def abrir(archivo) -> pd.ExcelFile:
    """Abre el libro una sola vez (abrir un libro grande toma segundos; las hojas se leen de este objeto)."""
    return archivo if isinstance(archivo, pd.ExcelFile) else pd.ExcelFile(archivo)


def _hoja_movimientos(xls: pd.ExcelFile):
    """(hoja, fila de encabezado) de la primera hoja con el reporte de SAP."""
    for h in xls.sheet_names:
        df = xls.parse(h, header=None, nrows=12, dtype=str)
        for i in range(len(df)):
            fila = {_n(v) for v in df.iloc[i] if pd.notna(v)}
            if "clase de documento" in fila and "importe en moneda local" in fila:
                return h, i
    raise ValueError("No encontré el reporte de SAP: falta una hoja con los encabezados 'Clase de documento' e "
                     "'Importe en moneda local'.")


def _texto(s: pd.Series) -> pd.Series:
    return s.map(lambda x: "" if pd.isna(x) else (str(int(x)) if isinstance(x, float) and x.is_integer()
                                                  else " ".join(str(x).split())))


def leer_movimientos(archivo) -> pd.DataFrame:
    """Reporte de SAP (columnas A a T; si trae 'Modificacion manual' / 'Subrubro caja' también se leen)."""
    xls = abrir(archivo)
    hoja, fila = _hoja_movimientos(xls)
    todo = xls.parse(hoja, header=None, dtype=object)
    enc = [_n(v) if pd.notna(v) else "" for v in todo.iloc[fila]]
    crudo = todo.iloc[fila + 1:].reset_index(drop=True)
    pos = {}
    for campo, alias in COLUMNAS_SAP.items():
        j = next((j for a in alias for j, e in enumerate(enc) if e == a and j not in pos.values()), None)
        if j is not None:
            pos[campo] = j
    falta = [c for c in OBLIGATORIAS if c not in pos]
    if falta:
        raise ValueError("Faltan columnas en el reporte de SAP: " + ", ".join(falta))
    col = lambda c: crudo.iloc[:, pos[c]] if c in pos else pd.Series("", index=crudo.index)  # noqa: E731
    fecha = pd.to_datetime(col("fecha"), errors="coerce")
    valor = pd.to_numeric(col("valor"), errors="coerce")
    out = pd.DataFrame({c: _texto(col(c)) for c in COLUMNAS_SAP if c not in ("fecha", "valor")})
    out["fecha"], out["valor"] = fecha.dt.strftime("%Y-%m-%d"), valor
    out = out[fecha.notna() & valor.notna() & (out["documento"] != "")].copy()
    fecha = fecha[out.index]
    out["anio"], out["mes"] = fecha.dt.year.astype(int), fecha.dt.month.astype(int)
    # tercero: código de 6 cifras del texto de cabecera (si no, el acreedor / deudor)
    ter = out["cab"].where(out["cab"].str.fullmatch(r"\d{6}"), "")
    for alt in ("acreedor", "deudor"):
        ter = ter.where(ter != "", out[alt].where(out[alt].str.fullmatch(r"\d{4,10}"), ""))
    out["ter_cod"] = ter
    out.loc[out["sub"].str.startswith("#"), "sub"] = ""          # #N/A del libro
    base = out["documento"] + "|" + out["cuenta"] + "|" + out["fecha"] + "|" + out["valor"].round(2).astype(str)
    out["clave"] = base + "#" + out.groupby(base).cumcount().astype(str)
    return out.reset_index(drop=True)


def leer_terceros(archivo) -> pd.DataFrame:
    """Maestro de terceros de SAP (hoja con 'Acreedor' y 'Nombre 1'): codigo, nombre, nit, sociedad (columna
    'Soc.': DB01 = ALDC, DB02 = ALSM). Vacío si no hay."""
    xls = abrir(archivo)
    for h in xls.sheet_names:
        df = xls.parse(h, header=None, nrows=6, dtype=str)
        for i in range(len(df)):
            enc = [_n(v) if pd.notna(v) else "" for v in df.iloc[i]]
            c_cod = next((j for j, e in enumerate(enc) if e in ("acreedor", "deudor", "codigo", "tercero")), None)
            c_nom = next((j for j, e in enumerate(enc) if e in ("nombre 1", "nombre", "razon social")), None)
            if c_cod is None or c_nom is None:
                continue
            t = xls.parse(h, header=None, skiprows=i + 1, dtype=object)
            c_nit = next((j for j, e in enumerate(enc) if e.startswith("n ident fis") or e in ("nit", "nif")), None)
            c_soc = next((j for j, e in enumerate(enc) if e in ("soc", "sociedad")), None)
            out = pd.DataFrame({"codigo": _texto(t.iloc[:, c_cod]), "nombre": _texto(t.iloc[:, c_nom]),
                                "nit": _texto(t.iloc[:, c_nit]) if c_nit is not None else "",
                                "sociedad": _texto(t.iloc[:, c_soc]).str.upper() if c_soc is not None else ""})
            out = out[(out["codigo"] != "") & (out["nombre"] != "")]
            return out.drop_duplicates(["codigo", "sociedad"]).reset_index(drop=True)
    return pd.DataFrame(columns=["codigo", "nombre", "nit", "sociedad"])


SOCIEDADES = {"DB01": "ALDC", "DB02": "ALSM"}     # columna 'Soc.' del maestro de SAP


def terceros_de(ter: pd.DataFrame, empresa: str) -> pd.DataFrame:
    """Terceros de una compañía según la sociedad (DB01 = ALDC, DB02 = ALSM); si el archivo no trae sociedad,
    todos."""
    if "sociedad" in ter and (ter["sociedad"] != "").any():
        ter = ter[ter["sociedad"].map(SOCIEDADES) == empresa]
    return ter.drop_duplicates("codigo")


def leer_param(archivo) -> tuple[pd.DataFrame, pd.DataFrame]:
    """De la hoja Param: (cuentas: cuenta, nombre) y (por_tercero: tercero, concepto) — el subrubro que el libro
    pone por tercero cuando no hay modificación manual."""
    vacio = (pd.DataFrame(columns=["cuenta", "nombre"]), pd.DataFrame(columns=["tercero", "concepto"]))
    xls = abrir(archivo)
    if "Param" not in xls.sheet_names:
        return vacio
    p = xls.parse("Param", header=None, dtype=object)
    cuentas = []
    for j in range(min(12, p.shape[1])):        # columnas con "1105100000 CAJAS MENORES"
        s = p.iloc[1:, j].dropna().astype(str)
        m = s.str.extract(r"^(\d{10})\s+(.+)$").dropna()
        if len(m) >= 1:
            cuentas.append(m.rename(columns={0: "cuenta", 1: "nombre"}))
    cuentas = pd.concat(cuentas).drop_duplicates("cuenta") if cuentas else vacio[0]
    por_ter = vacio[1]
    for i in range(3):
        for j in range(p.shape[1] - 1):
            if _n(p.iat[i, j]) == "tercero" and _n(p.iat[i, j + 1]) == "concepto":
                t = p.iloc[i + 1:, [j, j + 1]].dropna()
                t.columns = ["tercero", "concepto"]
                t = t.astype(str).map(lambda x: " ".join(x.split()))
                por_ter = t.drop_duplicates("tercero", keep="first")
    return cuentas.reset_index(drop=True), por_ter.reset_index(drop=True)


# ------------------------------------------------------------------ estructura y fórmulas de la hoja de flujo
_REF = r"\$?([A-Z]{1,3})\$?(\d+)"


def _col_num(letras: str) -> int:
    n = 0
    for ch in letras:
        n = n * 26 + ord(ch) - 64
    return n


def _partir(expr: str, seps: str) -> list[tuple[str, str]]:
    """Parte en el nivel superior (fuera de paréntesis y corchetes). Devuelve [(separador previo, término)]."""
    out, nivel, actual, previo = [], 0, "", "+"
    for ch in expr:
        if ch in "([":
            nivel += 1
        elif ch in ")]":
            nivel -= 1
        if ch in seps and nivel == 0:
            if actual.strip():
                out.append((previo, actual.strip()))
            elif ch == "-":                 # "--" o "=-": cambia el signo
                previo = "-" if previo == "+" else "+"
                continue
            previo, actual = ch, ""
            continue
        actual += ch
    if actual.strip():
        out.append((previo, actual.strip()))
    return out


def _referencias(arg: str, col: int, tablas: dict) -> list[tuple[int, int]] | None:
    """Celdas (fila, desfase de mes) de un argumento: A1, A1:A9, Tabla[Columna], Tabla[[#Totals],[Columna]]."""
    arg = arg.strip()
    if "!" in arg:
        return None
    m = re.fullmatch(_REF + r"(?::" + _REF + r")?", arg)
    if m:
        c1, f1 = _col_num(m.group(1)), int(m.group(2))
        f2 = int(m.group(4)) if m.group(4) else f1
        return [(f, c1 - col) for f in range(min(f1, f2), max(f1, f2) + 1)]
    m = re.fullmatch(r"(\w+)\[(.+)\]", arg)
    if m and m.group(1) in tablas:
        t = tablas[m.group(1)]
        dentro = m.group(2)
        nombre_col = re.findall(r"\[?([^\[\],#]+)\]?$", dentro)
        nombre_col = nombre_col[0].strip() if nombre_col else ""
        if nombre_col not in t["columnas"]:
            return None
        c = t["c1"] + t["columnas"].index(nombre_col)
        if "#Totals" in dentro:
            return [(t["f2"], c - col)] if t["totales"] else []
        return [(f, c - col) for f in range(t["f1"] + (1 if t["encabezado"] else 0),
                                            t["f2"] + (0 if t["totales"] else 1))]
    return None


def traducir(formula, col: int, tablas: dict) -> dict:
    """Fórmula de una celda -> {'sap': signo de SUMIFS(Journal/SAP), 'remun': signo de SUMIFS(Remuneraciones)
    (GMF de la remuneración de Enel), 'refs': [(fila, desfase, signo)], 'otro': bool}."""
    res = {"sap": 0, "remun": 0, "refs": [], "otro": False}
    if formula is None:
        return res
    if not isinstance(formula, str) or not formula.startswith("="):
        res["otro"] = bool(pd.to_numeric(formula, errors="coerce") or 0)
        return res
    for sep, termino in _partir(formula[1:], "+-"):
        signo = -1 if sep == "-" else 1
        t = termino.lstrip("+")
        if re.match(r"SUMIFS\(\s*'?(Journal|SAP)'?!", t, re.I):
            res["sap"] += signo
            continue
        if re.match(r"SUMIFS\(\s*'?Remuneraciones'?!", t, re.I):
            res["remun"] += signo
            res["otro"] = True
            continue
        m = re.fullmatch(r"SUM\((.*)\)", t, re.I)
        args = [a for _, a in _partir(m.group(1), ",")] if m else [t]
        for a in args:
            refs = _referencias(a, col, tablas)
            if refs is None:
                res["otro"] = True
            else:
                res["refs"] += [(f, d, signo) for f, d in refs]
    return res


def _tablas(ws) -> dict:
    out = {}
    for tabla in ws.tables.values():
        m = re.fullmatch(_REF + ":" + _REF, str(tabla.ref))
        if not m:
            continue
        out[tabla.name] = {"c1": _col_num(m.group(1)), "f1": int(m.group(2)), "f2": int(m.group(4)),
                           "columnas": [c.name for c in tabla.tableColumns], "totales": bool(tabla.totalsRowCount),
                           "encabezado": tabla.headerRowCount != 0}
    return out


def hoja_flujo(xls: pd.ExcelFile) -> str:
    """Hoja del flujo (en libros viejos puede llamarse distinto, ej. 'Flujo de caja 2023'): la que tiene 'ENE' en la
    columna C de sus primeras filas, prefiriendo las que se llaman 'flujo…'."""
    nombres = sorted(xls.sheet_names, key=lambda h: (not _n(h).startswith("flujo"), h != "Flujo de caja"))
    for h in nombres:
        d = xls.parse(h, header=None, nrows=15)
        if d.shape[1] > 2 and d.iloc[:, 2].astype(str).str.strip().str.upper().eq("ENE").any():
            return h
    raise ValueError("No encontré la hoja del flujo de caja (con los meses ENE, FEB… en la columna C).")


def leer_estructura(archivo, hoja: str = "Flujo de caja") -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    (estructura, valores_archivo, proyectado)
      estructura: orden, fila, concepto, seccion, padre, tipo, clase ('sap'|'formula'|'manual'), sap, terminos(json)
      valores_archivo: orden, mes, valor (lo que muestra el libro)
      proyectado: concepto, proyectado, observacion
    """
    import io
    from openpyxl import load_workbook
    datos = archivo.getvalue() if hasattr(archivo, "getvalue") else open(archivo, "rb").read()
    lin, proy = cargas.leer_flujo(io.BytesIO(datos), hoja)
    ws = load_workbook(io.BytesIO(datos), read_only=False, data_only=False)[hoja]
    tablas = _tablas(ws)
    info = lin.drop_duplicates("orden").sort_values("orden")
    filas = []
    for r in info.itertuples():
        trad = {m: traducir(ws.cell(int(r.fila), 2 + m).value, 2 + m, tablas) for m in range(1, 13)}
        sap = next((t["sap"] for t in trad.values() if t["sap"]), 0)
        remun = next((t["remun"] for t in trad.values() if t["remun"]), 0)
        con_refs = [m for m in range(1, 13) if trad[m]["refs"]]
        clase = "sap" if sap else ("formula" if con_refs else "manual")
        # términos generales: los que más se repiten en los meses (un mes con un ajuste a mano no cambia la regla;
        # en enero el saldo inicial es un valor fijo y de febrero en adelante viene del mes anterior)
        patrones = pd.Series([json.dumps(trad[m]["refs"]) for m in range(1, 13)
                              if ws.cell(int(r.fila), 2 + m).value is not None])
        base = patrones.mode().iloc[-1] if not patrones.empty else "[]"
        if base == "[]" and con_refs and clase == "formula":
            base = json.dumps(trad[max(con_refs)]["refs"])
        filas.append(dict(orden=r.orden, fila=int(r.fila), concepto=r.concepto, seccion=r.seccion, padre=r.padre,
                          tipo=r.tipo, clase=clase, sap=sap, terminos=base, remun=remun))
    est = pd.DataFrame(filas)
    vals = lin[["orden", "mes", "valor"]].copy()
    return est, vals, proy


# ------------------------------------------------------------------ cálculo del flujo
def calcular(est: pd.DataFrame, sumas: dict, ajustes: dict) -> pd.DataFrame:
    """
    est: estructura; sumas: {(clave_concepto, mes): valor de movimientos}; ajustes: {(orden, mes): valor}.
    Devuelve orden, mes, valor. Las referencias a filas que no son conceptos (filas vacías) valen 0, y las del mes
    anterior en enero también (el saldo inicial de enero va como ajuste).
    """
    por_fila = {int(r.fila): r for r in est.itertuples()}
    terminos = {int(r.fila): json.loads(r.terminos) for r in est.itertuples()}
    memo = {}

    def valor(fila, mes):
        if mes < 1 or fila not in por_fila:
            return 0.0
        k = (fila, mes)
        if k in memo:
            return memo[k]
        memo[k] = 0.0               # protege de referencias circulares
        r = por_fila[fila]
        v = ajustes.get((int(r.orden), mes), 0.0)
        if r.sap:
            v += r.sap * sumas.get((clave_concepto(r.concepto), mes), 0.0)
        for f, d, s in terminos[fila]:
            v += s * valor(f, mes + d)
        memo[k] = v
        return v

    return pd.DataFrame([(int(r.orden), m, valor(int(r.fila), m)) for r in est.itertuples() for m in range(1, 13)],
                        columns=["orden", "mes", "valor"])


def sumas_de(mov: pd.DataFrame) -> dict:
    m = mov[(mov["concepto"] != "") & (mov["concepto"] != ANULADO)]
    g = m.groupby([m["concepto"].map(clave_concepto), "mes"])["valor"].sum()
    return {(c, int(mes)): float(v) for (c, mes), v in g.items()}


def ajustes_del_archivo(est: pd.DataFrame, vals: pd.DataFrame, mov: pd.DataFrame) -> pd.DataFrame:
    """Ajuste = valor del libro − valor calculado con los movimientos y las fórmulas (por fila y mes)."""
    archivo = {(int(o), int(m)): float(v) for o, m, v in vals[["orden", "mes", "valor"]].itertuples(index=False)}
    orden_de = {int(r.fila): int(r.orden) for r in est.itertuples()}
    sumas = sumas_de(mov)
    filas = []
    for r in est.itertuples():
        for m in range(1, 13):
            calc = r.sap * sumas.get((clave_concepto(r.concepto), m), 0.0)
            for f, d, s in json.loads(r.terminos):
                if f in orden_de and m + d >= 1:
                    calc += s * archivo.get((orden_de[f], m + d), 0.0)
            aj = archivo.get((int(r.orden), m), 0.0) - calc
            if abs(aj) > 0.5:
                filas.append((int(r.orden), m, aj))
    return pd.DataFrame(filas, columns=["orden", "mes", "valor"])


def lineas(est: pd.DataFrame, valores: pd.DataFrame, anio: int, empresa: str) -> pd.DataFrame:
    """Tabla flujo_lineas (la que lee la vista)."""
    v = valores.merge(est[["orden", "seccion", "padre", "concepto", "tipo"]], on="orden")
    v["anio"], v["empresa"] = int(anio), empresa
    return v[["anio", "orden", "seccion", "padre", "concepto", "tipo", "mes", "valor", "empresa"]]


# ------------------------------------------------------------------ clasificación automática
REGLAS = [
    ("an", "cuenta", "ter_cod", "txd"),
    ("an", "cuenta", "ter_cod", "tx"),
    ("an", "cuenta", "ter_cod"),
    ("an", "ter_cod"),
    ("an", "cuenta", "txd"),
    ("an", "cuenta", "tx", "sg"),
    ("an", "tx", "sg"),
    ("an", "cuenta", "sg"),
]
NOMBRE_REGLA = {REGLAS[0]: "cuenta+tercero+texto", REGLAS[1]: "cuenta+tercero+texto", REGLAS[2]: "cuenta+tercero",
                REGLAS[3]: "tercero", REGLAS[4]: "cuenta+texto", REGLAS[5]: "cuenta+texto",
                REGLAS[6]: "texto", REGLAS[7]: "cuenta"}
PUREZA_MIN = 0.8
# palabras del texto que no dicen nada del concepto
_VACIAS = {"PAGO", "PG", "PAGOS", "FACTURA", "FACTURAS", "FRA", "FRAS", "FR", "FE", "MES", "DE", "DEL", "LA", "EL", "LOS",
           "LAS", "Y", "A", "EN", "POR", "CTA", "CTAS", "COB", "NO", "N", "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO",
           "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE", "ENE", "FEB", "MAR", "ABR",
           "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC", "Q", "QNA", "SAS", "S", "SA", "LTDA", "COP", "TRAS",
           "VARIAS", "VARIOS", "ABONO", "SALDO", "TOTAL", "CON", "PARA", "AL"}


def _palabras(m: pd.DataFrame) -> pd.DataFrame:
    """Una fila por (movimiento, palabra significativa del texto)."""
    w = m["texto"].fillna("").str.upper().str.replace(r"[^A-ZÑ ]", " ", regex=True).str.split()
    w = w.map(lambda x: sorted({p for p in x if len(p) >= 3 and p not in _VACIAS}))
    return m.assign(palabra=w).explode("palabra").dropna(subset=["palabra"])


def _rasgos(m: pd.DataFrame) -> pd.DataFrame:
    m = m.copy()
    m["an"] = np.where(m["anulado"].fillna("") != "", "1", "0")
    up = m["texto"].fillna("").str.upper()
    m["tx"] = up.str.replace(r"[^A-ZÑ ]", " ", regex=True).str.split().str[:3].str.join(" ")
    m["txd"] = up.str.replace(r"[^A-ZÑ0-9 ]", " ", regex=True).str.split().str[:4].str.join(" ")
    m["sg"] = np.where(m["valor"] >= 0, "+", "-")
    m["ter_cod"] = m["ter_cod"].fillna("")
    m["cuenta"] = m["cuenta"].fillna("").astype(str)
    return m


def entrenar(hist: pd.DataFrame) -> list:
    h = _rasgos(hist[hist["concepto"] != ""])
    modelo = []
    for claves in REGLAS:
        g = h.groupby(list(claves))["concepto"]
        conteo = g.value_counts()
        top = conteo.groupby(level=list(range(len(claves)))).head(1)
        tabla = top.reset_index(name="n_top").set_index(list(claves))
        tabla["n"] = g.size()
        tabla["pureza"] = tabla["n_top"] / tabla["n"]
        modelo.append((claves, tabla[["concepto", "pureza", "n"]]))
    # última regla aprendida: una palabra del texto que casi siempre va al mismo concepto (RETEFUENTE, NOMINA…)
    p = _palabras(h)
    g = p.groupby(["an", "sg", "palabra"])["concepto"]
    top = g.value_counts().groupby(level=[0, 1, 2]).head(1)
    tabla = top.reset_index(name="n_top").set_index(["an", "sg", "palabra"])
    tabla["n"] = g.size()
    tabla["pureza"] = tabla["n_top"] / tabla["n"]
    modelo.append((("palabra",), tabla[(tabla["pureza"] >= 0.9) & (tabla["n"] >= 3)][["concepto", "pureza", "n"]]))
    return modelo


def aplicar_fijas(m: pd.DataFrame, fijas: pd.DataFrame | None) -> pd.Series:
    """Reglas creadas en la app: tercero (código), cuenta y/o texto que contiene -> concepto.
    De la más general a la más específica: la última que aplica manda."""
    res = pd.Series("", index=m.index, dtype=object)
    if fijas is None or fijas.empty:
        return res
    texto = m["texto"].fillna("").str.upper()
    f = fijas.fillna("").copy()
    f["_n"] = (f["tercero"] != "").astype(int) + (f["cuenta"] != "").astype(int) + (f["texto"] != "").astype(int)
    for r in f.sort_values("_n").itertuples():
        ok = pd.Series(True, index=m.index)
        if r.tercero:
            ok &= m["ter_cod"] == str(r.tercero)
        if r.cuenta:
            ok &= m["cuenta"] == str(r.cuenta)
        if r.texto:
            ok &= texto.str.contains(str(r.texto).upper(), regex=False)
        res[ok] = r.concepto
    return res


def clasificar(movs: pd.DataFrame, modelo: list, fijas: pd.DataFrame | None = None,
               por_tercero: dict | None = None) -> pd.DataFrame:
    """Devuelve concepto, origen ('automatica'|'pendiente'), regla y confianza ('Alta'|'Media')."""
    m = _rasgos(movs)
    pred = pd.Series(pd.NA, index=m.index, dtype=object)
    regla = pd.Series("", index=m.index, dtype=object)
    conf = pd.Series("", index=m.index, dtype=object)
    fija = aplicar_fijas(m, fijas)
    hay = fija != ""
    pred[hay], regla[hay], conf[hay] = fija[hay], "regla fija", "Alta"
    for claves, tabla in modelo:
        falta = pred.isna()
        if not falta.any():
            break
        if claves == ("palabra",):
            j = _palabras(m.loc[falta, ["texto", "an", "sg"]]).join(tabla, on=["an", "sg", "palabra"]).dropna(subset=["concepto"])
            j = j.sort_values(["pureza", "n"], ascending=False)
            j = j[~j.index.duplicated()]
            pred[j.index], regla[j.index], conf[j.index] = j["concepto"], "palabra del texto", "Media"
            continue
        j = m.loc[falta, list(claves)].join(tabla, on=list(claves))
        ok = j["concepto"].notna() & (j["pureza"] >= PUREZA_MIN)
        idx = j.index[ok]
        pred[idx], regla[idx] = j.loc[ok, "concepto"], NOMBRE_REGLA[claves]
        conf[idx] = ["Alta" if (p >= 0.95 and n >= 3) else "Media" for p, n in zip(j.loc[ok, "pureza"], j.loc[ok, "n"])]
    if por_tercero:                              # subrubro por tercero de la hoja Param
        falta = pred.isna()
        t = m.loc[falta, "tercero"].map(lambda x: por_tercero.get(_n(x)))
        idx = t.index[t.notna()]
        pred[idx], regla[idx], conf[idx] = t[idx], "tercero (Param)", "Media"
    out = pd.DataFrame({"concepto": pred.fillna(""), "origen": "automatica", "regla": regla, "confianza": conf},
                       index=m.index)
    out.loc[pred.isna(), "origen"] = "pendiente"
    return out


def preparar(nuevos: pd.DataFrame, existentes: pd.DataFrame | None, fijas: pd.DataFrame | None = None,
             por_tercero: dict | None = None, validos: list | None = None, meses_historia: int = 24) -> pd.DataFrame:
    """
    Clasifica un lote de movimientos:
    - si el archivo trae el subrubro ya calculado (libro de flujo), se usa -> 'validado';
    - si no, la 'Modificacion manual' del archivo (si la trae) -> 'validado';
    - si el movimiento ya estaba cargado y clasificado a mano o validado, se conserva;
    - lo demás, con las reglas aprendidas de lo ya clasificado de la compañía: solo los últimos
      'meses_historia' meses (con años viejos cargados, la forma de clasificar de hace años no manda) y solo
      conceptos que existen en la estructura actual ('validos'), para no llevar movimientos a conceptos que ya no
      están en el flujo.
    """
    m = nuevos.copy()
    for c in ("concepto", "origen", "regla", "confianza"):
        m[c] = ""
    # el subrubro del libro (columna 'Subrubro caja') es el que suma el Excel: normalmente es igual a la
    # modificación manual, pero a veces se escribe encima a mano; la modificación manual solo cuenta si no hay
    # subrubro (reporte de SAP con la columna de modificación diligenciada)
    for col in ("sub", "manual"):
        if col in m:
            ok = (m["concepto"] == "") & (m[col].fillna("") != "")
            m.loc[ok, "concepto"], m.loc[ok, "origen"] = m.loc[ok, col], "validado"
    val = m["origen"] == "validado"
    hist = [m[val]]
    conservar = pd.Series(False, index=m.index)
    if existentes is not None and not existentes.empty:
        previos = existentes[existentes["origen"].isin(ORIGENES_ENTRENAMIENTO)].drop_duplicates("clave") \
            .set_index("clave")
        conservar = ~val & m["clave"].isin(previos.index)
        for c in ("concepto", "origen", "regla", "confianza"):
            m.loc[conservar, c] = m.loc[conservar, "clave"].map(previos[c])
        hist.append(existentes[existentes["origen"].isin(ORIGENES_ENTRENAMIENTO)])
    resto = ~val & ~conservar
    if resto.any():
        historial = pd.concat(hist, ignore_index=True)
        if not historial.empty and meses_historia:
            per = historial["anio"].astype(int) * 12 + historial["mes"].astype(int)
            tope = int((m["anio"].astype(int) * 12 + m["mes"].astype(int)).max())
            historial = historial[(per > tope - meses_historia) & (per <= tope)]
        if validos is not None and not historial.empty:
            ok = {clave_concepto(c) for c in validos}
            historial = historial[historial["concepto"].map(clave_concepto).isin(ok)]
        modelo = entrenar(historial) if not historial.empty else []
        r = clasificar(m[resto], modelo, fijas, por_tercero)
        if validos is not None:          # una regla fija o de Param hacia un concepto que ya no existe: pendiente
            ok = {clave_concepto(c) for c in validos}
            malo = (r["concepto"] != "") & ~r["concepto"].map(clave_concepto).isin(ok)
            r.loc[malo, ["concepto", "regla", "confianza"]] = ""
            r.loc[malo, "origen"] = "pendiente"
        for c in r.columns:
            m.loc[resto, c] = r[c]
    return m


def backtest(movs: pd.DataFrame) -> pd.DataFrame:
    """Precisión simulada: para cada mes, aprende de los meses anteriores y clasifica ese mes."""
    v = movs[movs["origen"].isin(ORIGENES_ENTRENAMIENTO) & (movs["concepto"] != "")]
    filas = []
    for (anio, mes), prueba in v.groupby(["anio", "mes"]):
        previo = v[(v["anio"] < anio) | ((v["anio"] == anio) & (v["mes"] < mes))]
        if previo.empty:
            continue
        r = clasificar(prueba, entrenar(previo))
        cub = r["origen"] != "pendiente"
        bien = cub & (r["concepto"] == prueba["concepto"])
        peso = prueba["valor"].abs()
        filas.append({"anio": anio, "mes": mes, "movimientos": len(prueba), "cubiertos": cub.mean(),
                      "acierto": bien.sum() / max(cub.sum(), 1),
                      "acierto_valor": peso[bien].sum() / peso[cub].sum() if peso[cub].sum() else 1.0})
    return pd.DataFrame(filas)


def leer_libro(archivo) -> dict:
    """Libro de flujo de caja completo (carga inicial): estructura, valores, proyectado, movimientos ya
    clasificados, terceros, cuentas y subrubro por tercero. Abre el archivo una vez por lector."""
    import io
    datos = archivo.getvalue() if hasattr(archivo, "getvalue") else open(archivo, "rb").read()
    xls = pd.ExcelFile(io.BytesIO(datos))
    est, vals, proy = leer_estructura(io.BytesIO(datos), hoja_flujo(xls))
    mov = leer_movimientos(xls)
    cuentas, por_ter = leer_param(xls)
    return dict(estructura=est, valores=vals, proyectado=proy, movimientos=mov, terceros=leer_terceros(xls),
                cuentas=cuentas, por_tercero=por_ter, gmf=leer_gmf_enel(xls))


GMF_ENEL = "gmf_enel"            # tipo de ajuste: GMF de la remuneración de Enel (hoja Remuneraciones del libro)


def leer_gmf_enel(xls) -> dict:
    """{mes: GMF} de la hoja Remuneraciones (primer bloque: MES APLICACIÓN y gmf). Vacío si no hay."""
    if "Remuneraciones" not in xls.sheet_names:
        return {}
    r = xls.parse("Remuneraciones", header=None)
    for i in range(min(6, len(r))):
        enc = [_n(v) if pd.notna(v) else "" for v in r.iloc[i]]
        if "mes aplicacion" in enc and "gmf" in enc:
            c_mes, c_gmf = enc.index("mes aplicacion"), enc.index("gmf")
            c_fec = c_mes - 1 if c_mes > 0 else c_mes
            t = r.iloc[i + 1:]
            mes = pd.to_datetime(t.iloc[:, c_fec], errors="coerce").dt.month
            mes = mes.fillna(pd.to_numeric(t.iloc[:, c_mes], errors="coerce"))
            v = pd.to_numeric(t.iloc[:, c_gmf], errors="coerce")
            g = v[mes.notna() & v.notna()].groupby(mes[mes.notna() & v.notna()].astype(int)).sum()
            return {int(m): float(x) for m, x in g.items() if x}
    return {}


def partir_ajustes(est: pd.DataFrame, aj: pd.DataFrame, gmf: dict) -> pd.DataFrame:
    """Separa del ajuste de cada fila el GMF de Enel (las filas que en el libro suman la hoja Remuneraciones), para
    poder registrarlo aparte mes a mes. Devuelve orden, mes, valor, tipo."""
    aj = aj.assign(tipo="")
    if not gmf or "remun" not in est:
        return aj
    filas = []
    for r in est[est["remun"] != 0].itertuples():
        for m, g in gmf.items():
            filas.append((int(r.orden), int(m), float(r.remun * g), GMF_ENEL))
    g = pd.DataFrame(filas, columns=["orden", "mes", "valor", "tipo"])
    resto = aj.merge(g[["orden", "mes", "valor"]], on=["orden", "mes"], how="outer", suffixes=("", "_g")).fillna(0)
    resto["valor"] = resto["valor"] - resto["valor_g"]
    resto = resto[resto["valor"].abs() > 0.5][["orden", "mes", "valor"]].assign(tipo="")
    return pd.concat([resto, g], ignore_index=True)


def vinculos(est: pd.DataFrame) -> dict:
    """{orden: [(concepto que la usa, signo)]}: filas que suman o restan a otra (ej. la facturación conjunta se
    resta en Prestación del servicio), sin contar totales ni grupos."""
    orden_de = dict(zip(est["fila"].astype(int), est["orden"].astype(int)))
    out = {}
    for r in est[est["tipo"] == "linea"].itertuples():
        for f, d, s in json.loads(r.terminos):
            if d == 0 and f in orden_de:
                out.setdefault(orden_de[f], []).append((str(r.concepto).rstrip("_"), s))
    return out


# ------------------------------------------------------------------ base de datos
COLS_MOV = ["anio", "mes", "fecha", "concepto", "tercero", "texto", "documento", "cuenta", "valor", "empresa",
            "clave", "clase", "cuenta_cod", "cc", "cab", "ter_cod", "ref1", "ref3", "referencia", "compensacion",
            "anulado", "origen", "regla", "confianza"]


def nombres_tercero(leer, empresa: str) -> dict:
    t = leer("SELECT codigo, nombre FROM flujo_terceros WHERE empresa=?", (empresa,))
    return dict(zip(t["codigo"].astype(str), t["nombre"]))


def a_tabla(m: pd.DataFrame, empresa: str, nombres: dict, cuentas: dict) -> pd.DataFrame:
    """Movimientos leídos + clasificación -> filas de flujo_mov."""
    t = m["ter_cod"].map(nombres)
    tercero = t.where(t.notna(), m["ref3"]).fillna("")
    tercero = tercero.where(tercero.str.strip() != "", "(Sin tercero)")
    out = pd.DataFrame({
        "anio": m["anio"], "mes": m["mes"], "fecha": m["fecha"], "concepto": m["concepto"], "tercero": tercero,
        "texto": m["texto"], "documento": m["documento"],
        "cuenta": (m["cuenta"] + " " + m["cuenta"].map(cuentas).fillna("")).str.strip(),
        "valor": m["valor"], "empresa": empresa, "clave": m["clave"], "clase": m["clase"], "cuenta_cod": m["cuenta"],
        "cc": m["cc"], "cab": m["cab"], "ter_cod": m["ter_cod"], "ref1": m["ref1"], "ref3": m["ref3"],
        "referencia": m["referencia"], "compensacion": m["compensacion"], "anulado": m["anulado"],
        "origen": m["origen"], "regla": m["regla"], "confianza": m["confianza"]})
    return out[COLS_MOV]


def movimientos(leer, empresa: str, anio: int | None = None) -> pd.DataFrame:
    q = "SELECT * FROM flujo_mov WHERE empresa=?" + (" AND anio=?" if anio else "")
    m = leer(q, (empresa, int(anio)) if anio else (empresa,))
    for c in ("concepto", "origen", "ter_cod", "cuenta_cod", "anulado", "clave", "texto"):
        if c in m:
            m[c] = m[c].fillna("")
    if "cuenta_cod" in m:
        m["cuenta"] = m["cuenta_cod"].where(m["cuenta_cod"] != "", m["cuenta"])
    return m


def conceptos(leer, empresa: str, anio: int | None = None) -> list[str]:
    """Conceptos de la caja a los que se puede llevar un movimiento (filas que suman movimientos de SAP)."""
    q = "SELECT concepto, MIN(orden) o FROM flujo_estructura WHERE empresa=? AND sap<>0" + \
        (" AND anio=?" if anio else "") + " GROUP BY concepto ORDER BY o"
    t = leer(q, (empresa, int(anio)) if anio else (empresa,))
    return t["concepto"].tolist() + [ANULADO]


def asegurar_estructura(con, empresa: str, anio: int) -> bool:
    """Si el año no tiene estructura, la copia del último año anterior; el saldo inicial de enero queda como
    ajuste con el saldo de diciembre de ese año. Devuelve False si no hay ninguna estructura de la compañía."""
    leer = con.leer
    if leer("SELECT 1 FROM flujo_estructura WHERE empresa=? AND anio=? LIMIT 1", (empresa, int(anio))).shape[0]:
        return True
    prev = leer("SELECT MAX(anio) a FROM flujo_estructura WHERE empresa=? AND anio<?", (empresa, int(anio)))["a"]
    if prev.isna().all():
        return False
    a0 = int(prev.iloc[0])
    est = leer("SELECT * FROM flujo_estructura WHERE empresa=? AND anio=?", (empresa, a0))
    dic = leer("SELECT orden, valor FROM flujo_lineas WHERE empresa=? AND anio=? AND mes=12", (empresa, a0))
    dic = dict(zip(dic["orden"].astype(int), dic["valor"]))
    orden_de = dict(zip(est["fila"].astype(int), est["orden"].astype(int)))
    ajustes = []
    for r in est.itertuples():
        v = sum(s * dic.get(orden_de.get(f), 0.0) for f, d, s in json.loads(r.terminos) if d == -1)
        if v:
            ajustes.append((empresa, int(anio), int(r.orden), 1, float(v)))
    con.anexar(est.drop(columns=[c for c in est.columns if c in ("id",)]).assign(anio=int(anio)), "flujo_estructura")
    con.executemany("INSERT INTO flujo_ajustes(empresa, anio, orden, mes, valor) VALUES(?,?,?,?,?)", ajustes)
    return True


def recalcular(con, empresa: str, anio: int) -> pd.DataFrame:
    """Recalcula y guarda flujo_lineas de la compañía y año (estructura + movimientos + ajustes)."""
    leer = con.leer
    est = leer("SELECT * FROM flujo_estructura WHERE empresa=? AND anio=? ORDER BY orden", (empresa, int(anio)))
    if est.empty:
        return pd.DataFrame()
    aj = leer("SELECT orden, mes, SUM(valor) valor FROM flujo_ajustes WHERE empresa=? AND anio=? GROUP BY orden, mes",
              (empresa, int(anio)))
    mov = leer("SELECT concepto, mes, valor FROM flujo_mov WHERE empresa=? AND anio=?", (empresa, int(anio)))
    mov["concepto"] = mov["concepto"].fillna("")
    vals = calcular(est, sumas_de(mov), {(int(o), int(m)): float(v) for o, m, v in aj.itertuples(index=False)})
    lin = lineas(est, vals, anio, empresa)
    con.execute("DELETE FROM flujo_lineas WHERE empresa=? AND anio=?", (empresa, int(anio)))
    con.anexar(lin, "flujo_lineas")
    return lin
