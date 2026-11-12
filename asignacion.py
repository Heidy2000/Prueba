"""
Asignación de sub-rubros a gerencias (no necesitas modificar este archivo).

Cada línea del presupuesto (código de clase + sub-rubro) tiene una gerencia dueña. La llave ignora tildes,
mayúsculas y espacios, así la asignación sirve para las tres versiones (Inicial, 4+8, 8+4) aunque el nombre
cambie levemente. Se guarda en la tabla 'asignacion':
  - codigo = '3.1', sub_norm = 'COMBUSTIBLE'  -> asignación hecha en la app para esa clase;
  - codigo = ''                              -> asignación general por nombre (archivo de dueños), vale para
                                                cualquier clase con ese sub-rubro.
La asignación específica manda sobre la general. Una específica con gerencia vacía = "sin asignar" a propósito.
"""
import pandas as pd

import cargas
from ui import SIN_DUENO


def tabla(leer) -> pd.DataFrame:
    return leer("SELECT codigo, sub_norm, sub_rubro, gerencia FROM asignacion")


def mapa(leer) -> dict:
    a = tabla(leer)
    return {(c or "", n): (g or "") for c, n, g in zip(a["codigo"], a["sub_norm"], a["gerencia"])}


def gerencia(codigo: str, sub: str, m: dict) -> str:
    n = cargas.norm(sub)
    g = m[(codigo, n)] if (codigo, n) in m else m.get(("", n), "")
    return g or SIN_DUENO


def columna(df: pd.DataFrame, m: dict) -> pd.Series:
    """Gerencia dueña de cada fila de un DataFrame con columnas codigo y sub_rubro."""
    if df.empty:
        return pd.Series([], index=df.index, dtype=object)
    pares = df[["codigo", "sub_rubro"]].fillna("").drop_duplicates()
    cache = {(c, s): gerencia(c, s, m) for c, s in zip(pares["codigo"], pares["sub_rubro"])}
    return pd.Series([cache[(c, s)] for c, s in zip(df["codigo"].fillna(""), df["sub_rubro"].fillna(""))],
                     index=df.index, dtype=object)


def dueno_de_sub(leer, sub: str) -> set:
    """Gerencias dueñas de un sub-rubro en cualquier clase (para cruces de otros módulos)."""
    n = cargas.norm(sub)
    a = tabla(leer)
    return set(a.loc[(a["sub_norm"] == n) & (a["gerencia"].fillna("") != ""), "gerencia"])


def guardar(con, filas: pd.DataFrame):
    """filas: codigo, sub_rubro, gerencia ('' = sin asignar). Reemplaza la asignación de esas líneas."""
    for r in filas.itertuples():
        cod, n = r.codigo or "", cargas.norm(r.sub_rubro)
        con.reemplazar_por("asignacion", {"codigo": cod, "sub_norm": n},
                           {"codigo": cod, "sub_norm": n, "sub_rubro": r.sub_rubro, "gerencia": r.gerencia or ""})
    con.commit()


def migrar(con):
    """Pasa la tabla antigua 'duenos' (solo por nombre) a la nueva, una sola vez."""
    if con.execute("SELECT COUNT(*) FROM asignacion").fetchone()[0]:
        return
    viejos = con.execute("SELECT sub_rubro, gerencia FROM duenos").fetchall()
    for sub, ger in viejos:
        n = cargas.norm(sub)
        con.reemplazar_por("asignacion", {"codigo": "", "sub_norm": n},
                           {"codigo": "", "sub_norm": n, "sub_rubro": sub, "gerencia": ger or ""})
    con.commit()
