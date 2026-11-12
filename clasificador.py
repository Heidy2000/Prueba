"""
Clasificación automática de la ejecución SAP

Aprende de los movimientos ya clasificados (validados por Financiera) qué código, sub-rubro y centro de coste
corresponde según los datos de SAP, y lo aplica a los movimientos nuevos. Prueba las reglas de la más específica
a la más general y se queda con la primera que sea confiable:

  1. cuenta + contrapartida (tercero) + centro de coste + inicio del texto
  2. cuenta + contrapartida + centro de coste
  3. cuenta + contrapartida
  4. cuenta + nombre del tercero
  5. cuenta + centro de coste
  6. cuenta

Antes de todo eso se aplican las reglas fijas que Financiera crea en la app (por tercero, por cuenta contable o
por ambos -> clasificación). Los ajustes que vengan en el archivo (Modifica COD / Base Gasto / CC) siempre mandan.
Lo que ninguna regla resuelve queda "Pendiente" para revisión manual.
"""
import pandas as pd

REGLAS = [
    ("cuenta", "contrapartida", "cc_sap", "texto3"),
    ("cuenta", "contrapartida", "cc_sap"),
    ("cuenta", "contrapartida"),
    ("cuenta", "tercero"),
    ("cuenta", "cc_sap"),
    ("cuenta",),
]
NOMBRE_REGLA = {
    REGLAS[0]: "cuenta+tercero+CC+texto", REGLAS[1]: "cuenta+tercero+CC", REGLAS[2]: "cuenta+tercero",
    REGLAS[3]: "cuenta+nombre tercero", REGLAS[4]: "cuenta+CC", REGLAS[5]: "cuenta",
}
PUREZA_MIN = 0.6     # la regla debe acertar al menos en el 60 % de su historial
ORIGENES_ENTRENAMIENTO = ("validado", "manual")


def _texto3(s: pd.Series) -> pd.Series:
    return (s.fillna("").str.upper().str.replace(r"[^A-ZÑ ]", " ", regex=True)
            .str.split().str[:3].str.join(" "))


def _objetivo(df: pd.DataFrame) -> pd.Series:
    return df["codigo"] + "\x1f" + df["sub_rubro"] + "\x1f" + df["cc"]


def entrenar(historial: pd.DataFrame) -> list:
    """historial: movimientos con codigo, sub_rubro, cc ya definidos (validados o corregidos a mano)."""
    h = historial.copy()
    h["texto3"] = _texto3(h["texto"])
    h["_obj"] = _objetivo(h)
    modelo = []
    for claves in REGLAS:
        g = h.groupby(list(claves))["_obj"]
        conteo = g.value_counts()
        top = conteo.groupby(level=list(range(len(claves)))).head(1)
        tabla = top.reset_index(name="n_top").set_index(list(claves))
        tabla["n"] = g.size()
        tabla["pureza"] = tabla["n_top"] / tabla["n"]
        modelo.append((claves, tabla[["_obj", "pureza", "n"]]))
    return modelo


def _norm(s: pd.Series) -> pd.Series:
    import unicodedata
    return s.fillna("").map(lambda x: " ".join(unicodedata.normalize("NFKD", str(x)).encode("ascii", "ignore")
                                                .decode().upper().split()))


def aplicar_fijas(movs: pd.DataFrame, fijas: pd.DataFrame | None) -> pd.Series:
    """Clasificación de reglas fijas por movimiento ('' si no aplica). Gana la regla con cuenta sobre la general."""
    res = pd.Series("", index=movs.index, dtype=object)
    if fijas is None or fijas.empty:
        return res
    t = _norm(movs["tercero"])
    cta = movs["cuenta"].fillna("").astype(str)
    f = fijas.copy()
    f["_t"], f["_c"] = f["tercero"].fillna("") != "", f["cuenta"].fillna("").astype(str) != ""
    # de la más general a la más específica (la última que aplica manda):
    # 1) solo cuenta (cualquier tercero)  2) solo tercero  3) tercero + cuenta
    for con_t, con_c in ((False, True), (True, False), (True, True)):
        for r in f[(f["_t"] == con_t) & (f["_c"] == con_c)].itertuples():
            ok = pd.Series(True, index=movs.index)
            if con_t:
                ok &= t == _norm(pd.Series([r.tercero])).iloc[0]
            if con_c:
                ok &= cta == str(r.cuenta)
            res[ok] = f"{r.codigo}\x1f{r.sub_rubro}\x1f{r.cc or ''}"
    return res


def clasificar(movs: pd.DataFrame, modelo: list, fijas: pd.DataFrame | None = None) -> pd.DataFrame:
    """
    movs: movimientos nuevos (con columnas de SAP y mod_cod/mod_sub/mod_cc).
    Devuelve las columnas codigo, sub_rubro, cc, origen, regla, confianza.
    """
    m = movs.copy()
    m["texto3"] = _texto3(m["texto"])
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
        j = m.loc[falta, list(claves)].join(tabla, on=list(claves))
        ok = j["_obj"].notna() & (j["pureza"] >= PUREZA_MIN)
        idx = j.index[ok]
        pred[idx] = j.loc[ok, "_obj"]
        regla[idx] = NOMBRE_REGLA[claves]
        conf[idx] = ["Alta" if (p >= 0.95 and n >= 3) else "Media"
                     for p, n in zip(j.loc[ok, "pureza"], j.loc[ok, "n"])]

    partes = pred.fillna("\x1f\x1f").str.split("\x1f", expand=True).reindex(columns=[0, 1, 2]).fillna("")
    out = pd.DataFrame({"codigo": partes[0], "sub_rubro": partes[1], "cc": partes[2],
                        "origen": "automatica", "regla": regla, "confianza": conf}, index=m.index)
    sin = pred.isna()
    out.loc[sin, "origen"] = "pendiente"
    out.loc[sin, "cc"] = m.loc[sin, "cc_sap"]
    out.loc[out["cc"] == "", "cc"] = m["cc_sap"]

    # ajustes del archivo (Modifica COD / Base Gasto / CC) siempre mandan
    for campo, ajuste in [("codigo", "mod_cod"), ("sub_rubro", "mod_sub"), ("cc", "mod_cc")]:
        if ajuste in m:
            hay = m[ajuste].fillna("") != ""
            out.loc[hay, campo] = m.loc[hay, ajuste]
            out.loc[hay & (out["origen"] == "pendiente") & (out["codigo"] != ""), "origen"] = "ajuste"
    out.loc[out["cc"] == "", "cc"] = m["cc_sap"]
    return out


def preparar(movs: pd.DataFrame, existentes: pd.DataFrame | None = None,
             fijas: pd.DataFrame | None = None, otros: pd.DataFrame | None = None) -> pd.DataFrame:
    """
    Clasifica un lote de movimientos de SAP:
    - si el archivo trae la clasificación validada (CODIGO...), se usa tal cual (origen 'validado');
    - si el movimiento ya estaba cargado y clasificado (validado o corregido a mano), se conserva;
    - lo demás se clasifica con las reglas aprendidas del historial.
    'existentes' son los movimientos ya cargados de la misma compañía; 'otros', los de las demás compañías
    (solo se usan para aprender).
    """
    m = movs.copy()
    for c in ["codigo", "sub_rubro", "cc", "origen", "regla", "confianza"]:
        m[c] = ""
    val = m["val_cod"].fillna("") != ""
    m.loc[val, "codigo"] = m.loc[val, "val_cod"]
    m.loc[val, "sub_rubro"] = m.loc[val, "val_sub"]
    m.loc[val, "cc"] = m.loc[val, "val_cc"].where(m.loc[val, "val_cc"] != "", m.loc[val, "cc_sap"])
    m.loc[val, "origen"] = "validado"

    hist = [m[val]]
    if existentes is not None and not existentes.empty:
        previos = existentes.set_index("clave")
        conservar = (~val) & m["clave"].isin(previos.index[previos["origen"].isin(ORIGENES_ENTRENAMIENTO)])
        for c in ["codigo", "sub_rubro", "cc", "origen", "regla", "confianza"]:
            m.loc[conservar, c] = m.loc[conservar, "clave"].map(previos[c])
        hist.append(existentes[existentes["origen"].isin(ORIGENES_ENTRENAMIENTO)])
    else:
        conservar = pd.Series(False, index=m.index)
    if otros is not None and not otros.empty:
        hist.append(otros[otros["origen"].isin(ORIGENES_ENTRENAMIENTO)])

    nuevos = ~val & ~conservar
    historial = pd.concat(hist, ignore_index=True)
    if nuevos.any():
        modelo = entrenar(historial) if not historial.empty else []
        r = clasificar(m[nuevos], modelo, fijas)
        for c in r.columns:
            m.loc[nuevos, c] = r[c]
    return m


def backtest(movs: pd.DataFrame) -> pd.DataFrame:
    """Precisión simulada: para cada mes, entrena con los meses anteriores y clasifica ese mes."""
    v = movs[movs["origen"].isin(ORIGENES_ENTRENAMIENTO)]
    filas = []
    for (anio, mes), prueba in v.groupby(["anio", "mes"]):
        previo = v[(v["anio"] < anio) | ((v["anio"] == anio) & (v["mes"] < mes))]
        if previo.empty:
            continue
        r = clasificar(prueba.assign(mod_cod="", mod_sub="", mod_cc=""), entrenar(previo))
        bien = (r["codigo"] == prueba["codigo"]) & (r["sub_rubro"] == prueba["sub_rubro"])
        bien_cc = bien & (r["cc"] == prueba["cc"])
        peso = prueba["valor"].abs()
        filas.append({"anio": anio, "mes": mes, "movimientos": len(prueba),
                      "cubiertos": (r["origen"] != "pendiente").mean(),
                      "acierto_rubro": bien.mean(), "acierto_valor": peso[bien].sum() / peso.sum() if peso.sum() else 1,
                      "acierto_con_cc": bien_cc.mean()})
    return pd.DataFrame(filas)
