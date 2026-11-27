"""Hallazgos de mantenimiento de Flota: recurrencias, repeticiones en periodos cortos, eventos atípicos, concentración del gasto y
alertas de seguimiento por placa.
"""
from __future__ import annotations

import numpy as np
import pandas as pd

# parámetros por defecto (la pantalla deja cambiar los dos primeros)
MIN_REINCIDENCIA = 3        # intervenciones del mismo concepto en la misma placa para llamarla reincidente
DIAS_REPETICION = 30        # dos intervenciones del mismo concepto a menos de estos días = repetición en periodo corto
Z_ATIPICO = 3.0             # z robusto (mediana y MAD de log(valor)) desde el cual un valor es atípico para su concepto
VECES_MEDIANA = 3.0         # además debe costar al menos tantas veces la mediana del concepto
MIN_INTERV_CONCEPTO = 6     # mínimo de intervenciones de un concepto (en 3 placas) para poder hablar de lo habitual
RAZON_TIPO = 3.0            # gasto de una placa / mediana de su tipo desde el cual es desproporcionado
Z_TIPO = 2.0                # además debe quedar así de lejos de sus pares en escala logarítmica (z robusto)
RAZON_FRECUENCIA = 2.0      # intervenciones de un concepto en una placa / lo típico de su tipo desde el cual llama la atención
MIN_PLACAS_TIPO = 3         # placas con gasto que debe tener el tipo para compararse
MIN_PLACAS_RUTINA = 4       # placas que deben tener el concepto para poder llamarlo rutinario
PARETO = 0.8                # los conceptos que juntos explican esta parte del gasto

SEGUIMIENTO = {1: "Seguimiento", 2: "Media", 3: "Alta"}


def _pesos(x) -> str:
    return "$" + f"{float(x):,.0f}".replace(",", ".")


def _fecha(x) -> str:
    return "" if pd.isna(x) else pd.Timestamp(x).strftime("%d/%m/%Y")


def _veces(n: int) -> str:
    return {2: "dos", 3: "tres", 4: "cuatro", 5: "cinco", 6: "seis"}.get(int(n), str(int(n)))


def _vacio(cols) -> pd.DataFrame:
    return pd.DataFrame(columns=list(cols))


def preparar(mov: pd.DataFrame, pendiente: str, otros: str, sin_placa: str) -> tuple[pd.DataFrame, dict]:
    """Movimientos analizables -> (intervenciones, cobertura). `mov` = movimientos del ámbito de placas del periodo elegido."""
    m = mov.copy()
    m["valor"] = pd.to_numeric(m["valor"], errors="coerce").fillna(0.0)
    concepto = m["concepto"].fillna("").astype(str)
    es_pend = concepto.eq(pendiente) | concepto.eq("")
    es_otros = concepto.eq(otros)
    es_sinp = m["placa"].fillna(sin_placa).eq(sin_placa)
    ok = ~es_pend & ~es_otros & ~es_sinp
    cob = dict(total=float(m["valor"].sum()), movimientos=len(m),
               pendiente=float(m.loc[es_pend, "valor"].sum()), otros=float(m.loc[~es_pend & es_otros, "valor"].sum()),
               sin_placa=float(m.loc[~es_pend & ~es_otros & es_sinp, "valor"].sum()),
               analizado=float(m.loc[ok, "valor"].sum()), movimientos_analizados=int(ok.sum()))
    cob["concilia"] = abs(cob["total"] - (cob["pendiente"] + cob["otros"] + cob["sin_placa"] + cob["analizado"])) < 1.0
    a = m[ok].copy()
    if a.empty:
        return _vacio(["placa", "concepto", "fecha", "valor", "movimientos", "documentos", "tercero", "texto"]), cob
    a["fecha"] = pd.to_datetime(a["fecha"], errors="coerce").dt.normalize()
    a["documento"] = a["documento"].fillna("").astype(str)
    a["tercero"] = a["tercero"].fillna("").astype(str) if "tercero" in a else ""
    a["texto"] = a["texto"].fillna("").astype(str) if "texto" in a else ""
    primero = lambda s: next((x for x in s if x and x != "nan"), "")
    iv = (a.groupby(["placa", "concepto", "fecha"], dropna=False)
          .agg(valor=("valor", "sum"), movimientos=("valor", "size"), documentos=("documento", lambda s: len({x for x in s if x})),
               tercero=("tercero", primero), texto=("texto", primero)).reset_index())
    return iv, cob


def _por_placa_concepto(iv: pd.DataFrame) -> pd.DataFrame:
    p = iv[iv["valor"] > 0]
    if p.empty:
        return _vacio(["placa", "concepto", "intervenciones", "costo", "primera", "ultima", "dias_min", "repeticiones_cortas"])
    return p.groupby(["placa", "concepto"]).agg(intervenciones=("valor", "size"), costo=("valor", "sum"),
                                                primera=("fecha", "min"), ultima=("fecha", "max")).reset_index()


def rutinarios(iv: pd.DataFrame, parte: float = 0.5) -> list[str]:
    """Conceptos que en la mayoría de las placas que los tienen se repiten (consumibles, lubricación…): repetirse es lo normal en ellos y
    taparía a los realmente llamativos. Se excluyen de reincidencia y repeticiones, y la pantalla lo informa."""
    p = iv[iv["valor"] > 0]
    if p.empty:
        return []
    n = p.groupby(["concepto", "placa"]).size().reset_index(name="n")
    r = n.groupby("concepto")["n"].agg(lambda s: (s >= 2).mean())
    k = n.groupby("concepto")["placa"].size()                       # con pocas placas no se puede hablar de «lo normal» del concepto
    return sorted(r[(r >= parte) & (k.reindex(r.index) >= MIN_PLACAS_RUTINA)].index)


def reincidencia(iv: pd.DataFrame, tipos: dict, minimo: int = MIN_REINCIDENCIA) -> pd.DataFrame:
    g = _por_placa_concepto(iv)
    g = g[g["intervenciones"] >= minimo].copy()
    if g.empty:
        return _vacio(["placa", "tipo", "concepto", "intervenciones", "costo", "primera", "ultima", "mensaje"])
    g["tipo"] = g["placa"].map(tipos).fillna("")
    g["mensaje"] = [
        f"{r.placa} tuvo {_veces(r.intervenciones)} intervenciones de «{r.concepto}» entre {_fecha(r.primera)} y {_fecha(r.ultima)}, "
        f"con un costo acumulado de {_pesos(r.costo)}. Es una recurrencia observada; conviene que mantenimiento valide si responde a una "
        "condición persistente o a trabajos programados." for r in g.itertuples()]
    return g.sort_values(["intervenciones", "costo"], ascending=False)[
        ["placa", "tipo", "concepto", "intervenciones", "costo", "primera", "ultima", "mensaje"]].reset_index(drop=True)


def repeticiones_cortas(iv: pd.DataFrame, tipos: dict, dias: int = DIAS_REPETICION) -> pd.DataFrame:
    """Pares de intervenciones consecutivas del mismo concepto en la misma placa separadas por `dias` o menos (y más de 0)."""
    p = iv[(iv["valor"] > 0) & iv["fecha"].notna()].sort_values(["placa", "concepto", "fecha"])
    filas = []
    for (pl, co), g in p.groupby(["placa", "concepto"]):
        f, v = g["fecha"].tolist(), g["valor"].tolist()
        for i in range(1, len(f)):
            d = (f[i] - f[i - 1]).days
            if 0 < d <= dias:
                filas.append(dict(placa=pl, tipo=tipos.get(pl, ""), concepto=co, primera=f[i - 1], segunda=f[i], dias=d,
                                  valor_primera=v[i - 1], valor_segunda=v[i], costo=v[i - 1] + v[i]))
    if not filas:
        return _vacio(["placa", "tipo", "concepto", "primera", "segunda", "dias", "valor_primera", "valor_segunda", "costo", "mensaje"])
    t = pd.DataFrame(filas)
    t["mensaje"] = [f"En {r.placa}, «{r.concepto}» se intervino el {_fecha(r.primera)} ({_pesos(r.valor_primera)}) y de nuevo el "
                    f"{_fecha(r.segunda)} ({_pesos(r.valor_segunda)}), {r.dias} días después. Repetición en periodo corto observada: "
                    "puede ser un trabajo en etapas o una reparación que no resolvió; requiere validación técnica." for r in t.itertuples()]
    return t.sort_values(["dias", "costo"], ascending=[True, False]).reset_index(drop=True)


def atipicos(iv: pd.DataFrame, tipos: dict, z_min: float = Z_ATIPICO, veces: float = VECES_MEDIANA) -> pd.DataFrame:
    """Intervenciones cuyo valor se sale de lo habitual del MISMO concepto en la flota.
    Los costos de mantenimiento son muy asimétricos (muchos arreglos pequeños y pocos grandes), así que se miden en ESCALA LOGARÍTMICA
    (z robusto con mediana y MAD de log(valor)): un valor es atípico si queda muy por encima de la mayoría Y cuesta al menos `veces` la
    mediana del concepto. Medido sobre el valor directo, casi cualquier arreglo grande de una distribución asimétrica saldría marcado."""
    p = iv[iv["valor"] > 0]
    filas = []
    for co, g in p.groupby("concepto"):
        if len(g) < MIN_INTERV_CONCEPTO or g["placa"].nunique() < 3:
            continue
        v = g["valor"].to_numpy(float)
        med = float(np.median(v))
        lv = np.log(v)
        lmed = float(np.median(lv))
        mad = float(np.median(np.abs(lv - lmed)))
        if mad > 0:
            z = 0.6745 * (lv - lmed) / mad
        else:                                           # sin dispersión: se usa el rango intercuartílico
            q1, q3 = np.percentile(lv, [25, 75])
            z = np.where(lv > q3 + 1.5 * max(q3 - q1, 1e-9), z_min, 0.0) if q3 > q1 else np.zeros(len(v))
        for (_, r), zi in zip(g.iterrows(), z):
            if zi >= z_min and r["valor"] >= veces * med:
                filas.append(dict(placa=r["placa"], tipo=tipos.get(r["placa"], ""), concepto=co, fecha=r["fecha"], valor=float(r["valor"]),
                                  mediana_concepto=med, veces_mediana=float(r["valor"]) / med if med else np.nan,
                                  tercero=r["tercero"], detalle=r["texto"]))
    if not filas:
        return _vacio(["placa", "tipo", "concepto", "fecha", "valor", "mediana_concepto", "veces_mediana", "tercero", "detalle", "mensaje"])
    t = pd.DataFrame(filas)
    t["mensaje"] = [f"{r.placa} registra en «{r.concepto}» un gasto de {_pesos(r.valor)} el {_fecha(r.fecha)}, "
                    f"{r.veces_mediana:.1f} veces lo habitual del concepto en la flota (mediana {_pesos(r.mediana_concepto)}). "
                    "Es un valor atípico frente al resto; no implica por sí mismo una falla: puede ser un trabajo mayor."
                    for r in t.itertuples()]
    return t.sort_values("valor", ascending=False).reset_index(drop=True)


def concentracion(iv: pd.DataFrame, pareto: float = PARETO) -> tuple[pd.DataFrame, str]:
    """Gasto por concepto (solo conceptos clasificados) y qué conceptos explican `pareto` del total."""
    if iv.empty:
        return _vacio(["concepto", "valor", "participacion", "acumulado", "placas", "intervenciones", "en_pareto"]), ""
    g = (iv.groupby("concepto").agg(valor=("valor", "sum"), placas=("placa", "nunique"), intervenciones=("valor", "size"))
         .reset_index().sort_values("valor", ascending=False))
    tot = g["valor"].sum()
    if tot <= 0:
        return g.assign(participacion=np.nan, acumulado=np.nan, en_pareto=False), ""
    g["participacion"] = g["valor"] / tot
    g["acumulado"] = g["participacion"].cumsum()
    g["en_pareto"] = (g["acumulado"] - g["participacion"]) < pareto
    n = int(g["en_pareto"].sum())
    top = g[g["en_pareto"]]
    nombres = ", ".join(f"«{c}»" for c in top["concepto"].head(4)) + (f" y {n - 4} más" if n > 4 else "")
    msg = (f"{n} de {len(g)} conceptos explican el {top['participacion'].sum() * 100:.0f}% del gasto clasificado ({_pesos(top['valor'].sum())}): "
           f"{nombres}. El de mayor peso es «{g.iloc[0]['concepto']}» con {g.iloc[0]['participacion'] * 100:.0f}% "
           f"({_pesos(g.iloc[0]['valor'])}), repartido en {int(g.iloc[0]['placas'])} placas.")
    return g.reset_index(drop=True), msg


def desproporcion_tipo(iv: pd.DataFrame, tipos: dict, razon: float = RAZON_TIPO) -> pd.DataFrame:
    """Placas cuyo gasto clasificado es `razon` veces o más la mediana de las placas de su mismo tipo."""
    if iv.empty:
        return _vacio(["placa", "tipo", "gasto", "mediana_tipo", "veces_tipo", "placas_tipo", "mensaje"])
    g = iv.groupby("placa")["valor"].sum().rename("gasto").reset_index()
    g["tipo"] = g["placa"].map(tipos).fillna("")
    g = g[(g["tipo"] != "") & (g["gasto"] > 0)]
    filas = []
    for tp, x in g.groupby("tipo"):
        if len(x) < MIN_PLACAS_TIPO:
            continue
        med = float(x["gasto"].median())
        if med <= 0:
            continue
        lg = np.log(x["gasto"].to_numpy(float))
        lmed = float(np.median(lg))
        mad = float(np.median(np.abs(lg - lmed)))
        zs = 0.6745 * (lg - lmed) / mad if mad > 0 else np.zeros(len(lg))          # el gasto es asimétrico: se mide en logaritmo
        for r, z in zip(x.itertuples(), zs):
            if r.gasto >= razon * med and (z >= Z_TIPO or mad == 0):
                filas.append(dict(placa=r.placa, tipo=tp, gasto=float(r.gasto), mediana_tipo=med, veces_tipo=float(r.gasto) / med,
                                  placas_tipo=len(x)))
    if not filas:
        return _vacio(["placa", "tipo", "gasto", "mediana_tipo", "veces_tipo", "placas_tipo", "mensaje"])
    t = pd.DataFrame(filas)
    t["mensaje"] = [f"{r.placa} concentra {_pesos(r.gasto)} en mantenimiento clasificado, {r.veces_tipo:.1f} veces la mediana de las "
                    f"{r.placas_tipo} placas de su tipo «{r.tipo}» ({_pesos(r.mediana_tipo)}). Es un gasto desproporcionado frente a sus "
                    "pares; falta validar si se explica por edad, uso o una reparación mayor." for r in t.itertuples()]
    return t.sort_values("veces_tipo", ascending=False).reset_index(drop=True)


def frecuencia_pares(iv: pd.DataFrame, tipos: dict, minimo: int = MIN_REINCIDENCIA, razon: float = RAZON_FRECUENCIA) -> pd.DataFrame:
    """Placas que intervienen un concepto MUCHO más veces que las de su mismo tipo. Lo típico de un concepto en un tipo es la mediana de
    intervenciones entre las placas de ese tipo que lo tuvieron (así, si todo un tipo repite mangueras, repetirlas no es llamativo).
    Solo se compara cuando al menos MIN_PLACAS_TIPO placas del tipo tienen el concepto. La placa debe tener `minimo` o más intervenciones
    y al menos `razon` veces lo típico."""
    g = _por_placa_concepto(iv)
    cols = ["placa", "tipo", "concepto", "intervenciones", "tipico", "pares", "veces", "costo", "primera", "ultima", "mensaje"]
    if g.empty:
        return _vacio(cols)
    g["tipo"] = g["placa"].map(tipos).fillna("")
    filas = []
    for (tp, co), x in g[g["tipo"] != ""].groupby(["tipo", "concepto"]):
        if len(x) < MIN_PLACAS_TIPO:
            continue
        tip = float(x["intervenciones"].median())
        for r in x.itertuples():
            if r.intervenciones >= minimo and r.intervenciones >= razon * max(tip, 1.0):
                filas.append(dict(placa=r.placa, tipo=tp, concepto=co, intervenciones=int(r.intervenciones), tipico=tip, pares=len(x),
                                  veces=r.intervenciones / max(tip, 1.0), costo=float(r.costo), primera=r.primera, ultima=r.ultima))
    if not filas:
        return _vacio(cols)
    t = pd.DataFrame(filas)
    t["mensaje"] = [f"{r.placa} intervino «{r.concepto}» {r.intervenciones} veces entre {_fecha(r.primera)} y {_fecha(r.ultima)} ({_pesos(r.costo)}); lo típico "
                    f"en «{r.tipo}» es {r.tipico:g} (mediana de {r.pares} placas que lo intervinieron). Es más de lo que hacen sus pares; "
                    "falta validar con mantenimiento si responde a una condición de la placa." for r in t.itertuples()]
    return t.sort_values(["veces", "costo"], ascending=False).reset_index(drop=True)


def comparativo_placa(iv: pd.DataFrame, tipos: dict, placa: str) -> pd.DataFrame:
    """Para una placa: por concepto, sus intervenciones frente a lo típico de su tipo (misma definición que `frecuencia_pares`)."""
    g = _por_placa_concepto(iv)
    cols = ["concepto", "intervenciones", "costo", "tipico", "pares", "veces"]
    if g.empty or placa not in set(g["placa"]):
        return _vacio(cols)
    g["tipo"] = g["placa"].map(tipos).fillna("")
    tp = tipos.get(placa, "")
    mia = g[g["placa"] == placa].copy()
    pares = g[(g["tipo"] == tp) & (tp != "")]
    tip = pares.groupby("concepto")["intervenciones"].agg(tipico="median", pares="size")
    mia = mia.merge(tip, left_on="concepto", right_index=True, how="left")
    mia.loc[mia["pares"].fillna(0) < MIN_PLACAS_TIPO, ["tipico", "pares"]] = np.nan          # con pocos pares no hay «lo típico»
    mia["veces"] = mia["intervenciones"] / mia["tipico"].clip(lower=1.0)
    return mia.sort_values("costo", ascending=False)[cols].reset_index(drop=True)


def gasto_relativo(iv: pd.DataFrame, tipos: dict) -> pd.DataFrame:
    """Gasto clasificado de cada placa frente a la mediana de las de su tipo (solo tipos con MIN_PLACAS_TIPO placas con gasto)."""
    g = iv.groupby("placa")["valor"].sum().rename("gasto").reset_index() if len(iv) else _vacio(["placa", "gasto"])
    g["tipo"] = g["placa"].map(tipos).fillna("")
    g = g[(g["tipo"] != "") & (g["gasto"] > 0)]
    med = g.groupby("tipo")["gasto"].agg(mediana="median", n="size")
    g = g.merge(med, left_on="tipo", right_index=True)
    g = g[g["n"] >= MIN_PLACAS_TIPO]
    g["veces_tipo"] = g["gasto"] / g["mediana"]
    return g[["placa", "veces_tipo"]].reset_index(drop=True)


def pareto_sobrecosto(iv: pd.DataFrame, tipos: dict) -> pd.DataFrame:
    """Sobrecosto de cada placa = su gasto clasificado menos la mediana de las placas de su tipo (solo la parte positiva), ordenado de mayor a
    menor con su participación y su acumulado. Es una medida de «cuánto gasta por encima de lo típico de su tipo»: NO es ejecución menos
    presupuesto. Solo entran tipos con MIN_PLACAS_TIPO placas con gasto (con menos no hay «lo típico»)."""
    cols = ["placa", "tipo", "gasto", "mediana_tipo", "sobrecosto", "participacion", "acumulado"]
    if iv.empty:
        return _vacio(cols)
    g = iv.groupby("placa")["valor"].sum().rename("gasto").reset_index()
    g["tipo"] = g["placa"].map(tipos).fillna("")
    g = g[(g["tipo"] != "") & (g["gasto"] > 0)]
    med = g.groupby("tipo")["gasto"].agg(mediana_tipo="median", n="size")
    g = g.merge(med, left_on="tipo", right_index=True)
    g = g[g["n"] >= MIN_PLACAS_TIPO]
    g["sobrecosto"] = (g["gasto"] - g["mediana_tipo"]).clip(lower=0.0)
    g = g[g["sobrecosto"] > 0].sort_values("sobrecosto", ascending=False)
    if g.empty:
        return _vacio(cols)
    tot = g["sobrecosto"].sum()
    g["participacion"] = g["sobrecosto"] / tot
    g["acumulado"] = g["participacion"].cumsum()
    return g[cols].reset_index(drop=True)


def clase_pareto(par: pd.DataFrame, corte: float = PARETO, corte_b: float = 0.95) -> pd.Series:
    """A: placas hasta cubrir `corte` del sobrecosto (la que lo cruza entra); B: hasta `corte_b`; C: el resto."""
    previo = par["acumulado"] - par["participacion"]
    return pd.Series(np.where(previo < corte, "A", np.where(previo < corte_b, "B", "C")), index=par.index)


def alertas(iv: pd.DataFrame, frec: pd.DataFrame, atip: pd.DataFrame, desp: pd.DataFrame, tipos: dict) -> pd.DataFrame:
    """Una fila por placa con las señales que presenta frente a sus pares. Más señales de tipos distintos = más prioridad de revisión.
    Tipos de señal: Frecuencia (repite un concepto más que las de su tipo), Gasto atípico (un valor fuera de lo habitual del concepto) y
    Gasto vs su tipo (gasta mucho más que la mediana de su tipo)."""
    señales: dict[str, dict[str, list[str]]] = {}
    for r in frec.itertuples():
        señales.setdefault(r.placa, {}).setdefault("Frecuencia", []).append(f"{r.concepto} ×{r.intervenciones} (típico {r.tipico:g})")
    for r in atip.itertuples():
        señales.setdefault(r.placa, {}).setdefault("Gasto atípico", []).append(f"{r.concepto} {_pesos(r.valor)}")
    for r in desp.itertuples():
        señales.setdefault(r.placa, {}).setdefault("Gasto vs su tipo", []).append(f"{r.veces_tipo:.1f}× la mediana")
    if not señales:
        return _vacio(["placa", "tipo", "gasto", "señales", "detalle", "prioridad", "mensaje"])
    gasto = iv.groupby("placa")["valor"].sum()
    filas = []
    for pl, s in señales.items():
        n = len(s)
        filas.append(dict(placa=pl, tipo=tipos.get(pl, ""), gasto=float(gasto.get(pl, 0.0)), n_señales=n, señales=" · ".join(s),
                          detalle=" | ".join(f"{k}: {', '.join(v)}" for k, v in s.items()), prioridad=SEGUIMIENTO[min(n, 3)]))
    t = pd.DataFrame(filas)
    t["mensaje"] = [f"{r.placa} presenta {r.n_señales} tipo{'s' if r.n_señales > 1 else ''} de señal frente a sus pares ({r.señales.lower()}) sobre un "
                    f"gasto clasificado de {_pesos(r.gasto)}. Amerita revisión técnica para confirmar si hay una condición de fondo; por ahora son "
                    "patrones observados, no un diagnóstico." for r in t.itertuples()]
    return t.sort_values(["n_señales", "gasto"], ascending=False).drop(columns="n_señales").reset_index(drop=True)


def analizar(mov: pd.DataFrame, tipos: dict, pendiente: str, otros: str, sin_placa: str, *, minimo: int = MIN_REINCIDENCIA,
             dias: int = DIAS_REPETICION, concepto: str | None = None, incluir_rutinarios: bool = False) -> dict:
    """Todo el análisis. `concepto`: si se da, los hallazgos por placa/evento se limitan a ese concepto (la concentración y la cobertura no).
    Reincidencia y repeticiones dejan fuera los conceptos rutinarios, salvo que se pida lo contrario o se elija ese concepto."""
    iv, cob = preparar(mov, pendiente, otros, sin_placa)
    conc, msg_conc = concentracion(iv)
    rut = rutinarios(iv)
    base = iv if not concepto else iv[iv["concepto"] == concepto]
    if not (incluir_rutinarios or concepto):
        base = base[~base["concepto"].isin(rut)]
    rein, rep = reincidencia(base, tipos, minimo), repeticiones_cortas(base, tipos, dias)        # absolutas: informativas (no mandan en la prioridad)
    frec = frecuencia_pares(iv if not concepto else iv[iv["concepto"] == concepto], tipos, minimo)
    atip = atipicos(iv, tipos)
    atip = atip if not concepto else atip[atip["concepto"] == concepto]
    desp = desproporcion_tipo(iv, tipos)
    return dict(intervenciones=iv, cobertura=cob, concentracion=conc, mensaje_concentracion=msg_conc, reincidencia=rein,
                repeticiones=rep, frecuencia=frec, rutinarios=rut, atipicos=atip, desproporcion=desp, gasto_relativo=gasto_relativo(iv, tipos), pareto=pareto_sobrecosto(iv, tipos),
                alertas=alertas(iv, frec, atip, desp, tipos))
