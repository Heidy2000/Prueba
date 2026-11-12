"""
Líneas de negocio (ALSM): de qué línea es cada movimiento y cómo se reparten los costos compartidos.

- Ejecución: la línea sale del centro de coste (hoja 'CECO ALSM', columna PROYECTO, guardada en la tabla
  ceco_linea). Los centros compartidos (ADMINISTRATIVO, OPERATIVO, MANTENIMIENTO, COMERCIAL) se reparten cada mes
  entre las líneas que facturan ese mes, según su participación en el ingreso del mes (como la hoja
  'PYG Linea de negocio' del cierre).
- Presupuesto: la línea viene en el archivo. Lo compartido (ADMINISTRATIVO) se reparte con los % por año de la
  tabla 'distribucion' (hoja 2 del archivo de presupuesto). En los meses reales de la versión (1–8 en la 8+4,
  1–4 en la 4+8) se reparte igual que la ejecución, para que esos meses coincidan.
"""
import pandas as pd

COMPARTIDAS = ["ADMINISTRATIVO", "OPERATIVO", "MANTENIMIENTO", "COMERCIAL"]
SIN_LINEA = "SIN LÍNEA"
MESES_REALES = {"8+4": 8, "4+8": 4}


def mapa_ceco(leer, empresa: str) -> dict:
    t = leer("SELECT cc, linea FROM ceco_linea WHERE empresa=?", (empresa,))
    return dict(zip(t["cc"].astype(str).str.strip(), t["linea"]))


def hay_lineas(leer, empresa: str) -> bool:
    return bool(leer("SELECT 1 FROM ceco_linea WHERE empresa=? LIMIT 1", (empresa,)).shape[0]) or \
        bool(leer("SELECT 1 FROM presupuesto WHERE empresa=? AND linea<>'' LIMIT 1", (empresa,)).shape[0])


def linea_ejecucion(ejec: pd.DataFrame, mapa: dict) -> pd.Series:
    """Línea de cada movimiento según su centro de coste (el clasificado; si no hay, el de SAP)."""
    cc = ejec["cc"].fillna("").astype(str).str.strip()
    cc = cc.where(cc != "", ejec["cc_sap"].fillna("").astype(str).str.strip())
    return cc.map(mapa).fillna(SIN_LINEA)


def pesos_ingreso(ejec: pd.DataFrame) -> dict:
    """{(anio, mes): {linea: peso}}: participación de cada línea (no compartida) en el ingreso del mes."""
    ing = ejec[(ejec["codigo"] == "1.0") & ~ejec["linea"].isin(COMPARTIDAS + [SIN_LINEA])]
    g = (-ing.groupby(["anio", "mes", "linea"])["valor"].sum())          # ingreso en positivo
    g = g[g > 0]
    out = {}
    for (a, m), s in g.groupby(level=[0, 1]):
        s = s.droplevel([0, 1])
        out[(int(a), int(m))] = (s / s.sum()).to_dict()
    return out


def pesos_anuales(dist: pd.DataFrame, lineas_validas=None) -> dict:
    """{anio: {linea: peso}} con los % de la tabla de distribución (solo líneas con %>0 y, si se da, existentes)."""
    out = {}
    for a, g in dist.groupby("anio"):
        g = g[g["pct"] > 0]
        if lineas_validas is not None:
            g = g[g["linea"].isin(lineas_validas)]
        if g["pct"].sum() > 0:
            out[int(a)] = dict(zip(g["linea"], g["pct"] / g["pct"].sum()))
    return out


def distribuir(df: pd.DataFrame, pesos_mes: dict, pesos_anio: dict | None = None,
               meses_reales: int = 0) -> pd.DataFrame:
    """Reparte las filas de líneas compartidas entre las líneas receptoras. El total no cambia.
    pesos_mes manda en la ejecución y en los meses reales del presupuesto; pesos_anio en los demás meses."""
    if df.empty or "linea" not in df:
        return df
    comp = df["linea"].isin(COMPARTIDAS)
    if not comp.any():
        return df
    partes = [df[~comp]]
    sin = []
    for (a, m), g in df[comp].groupby(["anio", "mes"]):
        a, m = int(a), int(m)
        if pesos_anio is None or m <= meses_reales:
            pesos = pesos_mes.get((a, m)) or (pesos_anio or {}).get(a)
        else:
            pesos = pesos_anio.get(a) or pesos_mes.get((a, m))
        if not pesos:
            sin.append(g)              # sin base para repartir: queda en su línea compartida
            continue
        for linea, p in pesos.items():
            partes.append(g.assign(valor=g["valor"] * p, linea=linea, linea_origen=g["linea"]))
    out = pd.concat(partes + sin, ignore_index=True)
    if "linea_origen" in out:
        out["linea_origen"] = out["linea_origen"].fillna("")
    return out
