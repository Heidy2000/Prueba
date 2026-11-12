"""Cruce entre ejecución y presupuesto según la versión elegida (no necesitas modificar este archivo)."""
import difflib

import pandas as pd

import cargas


def homologar(ejec: pd.DataFrame, ppto_version: pd.DataFrame, homol: pd.DataFrame, version: str) -> pd.DataFrame:
    """
    Devuelve la ejecución con 'sub_rubro' expresado con los nombres de esa versión de presupuesto:
    1) movimientos sin sub-rubro: van a la única línea de su clase, o a '<clase> · sin concepto';
    2) mismo nombre ignorando tildes, mayúsculas y espacios -> nombre del presupuesto;
    3) homologación guardada por Financiera (nombre de ejecución -> nombre del presupuesto);
    4) si no, queda con su nombre (se verá como 'Sin presupuesto').
    """
    e = ejec.copy()
    e["sub_original"] = e["sub_rubro"]
    lineas = ppto_version.drop_duplicates(["codigo", "sub_rubro"])
    por_norm = {(c, cargas.norm(s)): s for c, s in zip(lineas["codigo"], lineas["sub_rubro"])}
    unicas = lineas.groupby("codigo")["sub_rubro"].agg(list)
    h = homol[homol["version"] == version] if not homol.empty else homol
    por_homol = {(c, cargas.norm(s)): p for c, s, p in zip(h.get("codigo", []), h.get("sub_ejec", []),
                                                           h.get("sub_ppto", []))}

    def destino(cod, sub):
        if (cod, cargas.norm(sub)) in por_homol:
            return por_homol[(cod, cargas.norm(sub))]
        if not sub:
            subs = unicas.get(cod, [])
            return subs[0] if len(subs) == 1 else f"{cargas.CLASES.get(cod, cod)} · sin concepto"
        clave = (cod, cargas.norm(sub))
        return por_homol.get(clave) or por_norm.get(clave) or sub

    pares = e[["codigo", "sub_rubro"]].drop_duplicates()
    mapa = {(c, s): destino(c, s) for c, s in zip(pares["codigo"], pares["sub_rubro"])}
    e["sub_rubro"] = [mapa[(c, s)] for c, s in zip(e["codigo"], e["sub_rubro"])]
    return e


def sin_cruce(ejec: pd.DataFrame, ppto_version: pd.DataFrame, homol: pd.DataFrame, version: str) -> pd.DataFrame:
    """Sub-rubros de la ejecución que no encuentran línea en esa versión, con una sugerencia de homologación."""
    e = homologar(ejec, ppto_version, homol, version)
    lineas = set(zip(ppto_version["codigo"], ppto_version["sub_rubro"]))
    g = e.groupby(["codigo", "sub_original", "sub_rubro"])["valor"].sum().reset_index()
    g = g[[(c, s) not in lineas for c, s in zip(g["codigo"], g["sub_rubro"])]]
    filas = []
    for r in g.itertuples():
        candidatas = sorted(ppto_version.loc[ppto_version["codigo"] == r.codigo, "sub_rubro"].unique())
        sug = difflib.get_close_matches(r.sub_original, candidatas, n=1, cutoff=0.55)
        filas.append({"Código": r.codigo, "Sub-rubro en la ejecución": r.sub_original or "(vacío)",
                      "Ejecutado (M)": r.valor / 1e6, "Sugerencia": sug[0] if sug else "",
                      "Opciones": candidatas})
    return pd.DataFrame(filas)


def validar_versiones(ppto: pd.DataFrame, ejec: pd.DataFrame, corte: int) -> tuple[pd.DataFrame, list[str]]:
    """Tabla por clase (presupuesto de cada versión y ejecución, enero al corte) y lista de alertas."""
    alertas = []
    pc = ppto[ppto["mes"] <= corte].pivot_table(index="codigo", columns="version", values="valor", aggfunc="sum")
    ec = ejec[ejec["mes"] <= corte].groupby("codigo")["valor"].sum().rename("Ejecución")
    t = pc.join(ec, how="outer")
    for version in pc.columns:
        p = ppto[ppto["version"] == version]
        codigos = set(p["codigo"])
        malos = sorted(c for c in codigos if c not in cargas.RUBROS)
        if malos:
            alertas.append(f"**{version}**: códigos que no existen en la tabla de rubros: {', '.join(malos)}.")
        faltan = sorted((c for c in set(ejec["codigo"]) - codigos if c in cargas.RUBROS), key=cargas.orden_clase)
        if faltan:
            alertas.append(f"**{version}**: clases con ejecución pero sin presupuesto: "
                           + ", ".join(cargas.RUBROS[c] for c in faltan) + ".")
        ing = p.loc[p["codigo"] == "1.0", "valor"].sum()
        if ing > 0:
            alertas.append(f"**{version}**: los ingresos (1.0) están en positivo; en SAP van en negativo. "
                           "Revisa el signo en el Excel.")
    t = t.reindex(sorted(t.index, key=cargas.orden_clase))
    t.insert(0, "Clase", [cargas.RUBROS.get(c, f"{c} (no reconocido)") for c in t.index])
    for col in t.columns[1:]:
        t[col] = t[col] / 1e6
    return t.reset_index(drop=True), alertas
