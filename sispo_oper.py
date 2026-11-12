"""
SISPO para el módulo Operaciones (lógica pura, sin Streamlit ni base de datos, para poder probarla sola)
=========================================================================================================
Filtros previos (se aplican a los datos ORIGINALES del SISPO, antes de clasificar, agrupar o consolidar; el archivo nunca se modifica):
  1. Centro de costo (código) en DB010500 Staf Operaciones, DB010504 Barrido Manual, DB010515 Proyecto Aguas de Bogotá,
     DB010501 Recolección Domiciliaria, DB010509 Planeación.
  2. Grupo [5] - INSUMOS o [7] - HERRAMIENTAS, salvo las referencias que CRUZAN con un ítem del presupuesto (por la columna «Concepto» del
     SISPO o por la tabla de equivalencias): esas entran sea cual sea su grupo (ej. Aro 20 aluminio es grupo 1 Repuestos). Toda la cuenta
     OTROS de esos centros también entra, de cualquier grupo: SAP la carga completa en 3.1. Lo que no es grupo 5/7 ni cruza va a Herramientas
     (donde SAP carga el genérico) con la observación de `observacion()`.
  3. «Nombre Cta» = ÚTILES, PAPELERÍA Y OTROS (Papelería) u OTROS (Consumibles, Herramientas y Otros).
  4. Se excluye «Staf Mantenimiento» aunque pase lo anterior (por nombre, también con coincidencias parciales).
Reglas por hoja, sobre lo que pasó los filtros:
- Papelería    <- cuenta ÚTILES, PAPELERÍA Y OTROS. Solo salidas (cantidad negativa). Se compara en CANTIDADES: costo cero.
- Consumibles  <- Grupo [5] - INSUMOS y cuenta OTROS, salvo los dos códigos de bolsas. Solo salidas (costo negativo).
- Herramientas <- Grupo [7] - HERRAMIENTAS y cuenta OTROS. Solo salidas (costo negativo).
- Otros        <- exclusivamente los códigos 5064610001 y 5064610002. Solo salidas (costo negativo).
Una línea cae en UNA sola hoja. Lo que no cumple ninguna regla no se muestra en Operaciones.

Relación con SAP: SAP carga todas las salidas de inventario de la cuenta OTROS como UN genérico («Salida de inventario <mes>») en
el sub-rubro Herramientas del rubro 3.1. El detalle del SISPO reemplaza ese genérico SOLO si el total de la cuenta OTROS del SISPO
coincide al centavo con el de SAP (misma regla que Mantenimiento); si no cuadra, SAP queda como está y se avisa.

Este módulo NO reemplaza ni modifica la reclasificación de Mantenimiento: lee el archivo por su cuenta y guarda en tablas propias.
"""
import pandas as pd

import logica_m as L

HOJA_PAPELERIA, HOJA_CONSUMIBLES, HOJA_HERRAMIENTAS, HOJA_OTROS = "Papelería", "Consumibles", "Herramientas", "Otros"
HOJAS_PESOS = (HOJA_CONSUMIBLES, HOJA_HERRAMIENTAS, HOJA_OTROS)       # se comparan en pesos; Papelería, en cantidades
CUENTA_PAPELERIA = "ÚTILES, PAPELERÍA Y OTROS"                       # «Nombre Cta» completo; el archivo lo trae recortado a 20 caracteres
CUENTA_OTROS = "OTROS"
CENTROS_AUTORIZADOS = {"DB010500": "STAF OPERACIONES", "DB010504": "BARRIDO MANUAL", "DB010515": "PROYECTO AGUAS DE BOGOTÁ",
                       "DB010501": "RECOLECCIÓN DOMICILIARIA", "DB010509": "PLANEACIÓN"}
GRUPOS_AUTORIZADOS = ("5INSUMOS", "7HERRAMIENTAS")                   # «[5] - INSUMOS» / «[7] - HERRAMIENTAS» en forma canónica
PATRON_EXCLUIDO = r"STAFF?MANT"                                       # «Staf Mantenimiento» (también recortado o con una f de más)
CODIGOS_OTROS = ("5064610001", "5064610002")        # BOLSA CESTA 66x76 y BOLSA DESARENE 40x49
GRUPO_INSUMOS, GRUPO_HERRAMIENTAS = GRUPOS_AUTORIZADOS
SIN_HOJA = "Sin hoja asignada"

# Equivalencias de partida referencia SISPO -> ítem del presupuesto (validadas con el usuario). Editables en pantalla (admin).
MAPEO_INICIAL = {
    # Papelería
    "5023610006": "Café (libra)", "52336010042": "Instacream (tarro 450gr)", "52336010039": "Azucar (kilo)",
    "5033810063": "Resma papel carta", "5033810028": "Esferos", "5033810020": "Corrector", "5033810045": "Pagastic",
    "5013510026": "Mezcaldores (Paquete)", "5033810018": "Clips Mariposa", "5033810058": "Reslatadores", "5033810060": "Reslatadores",
    "5033810077": "Marcadores permanente", "5033810090": "Marcadores permanente",
    # Consumibles
    "5043910017": "Thiner gl", "5054510055": "Corazas", "5064810006": "Neumático", "5064810003": "Kit despinche",
    "52748010034": "Caneca carro papeleo", "52235010073": "Detergente", "52748010043": "Nylon (rollo 400 m)",
    "52951010070": "Llanta carro papelero", "5013510034": "Toalla de manos",
    # presupuestados en Herramientas
    "5064710001": "Cepíllos", "52850010053": "Rastrillos", "5043910033": "Espatulas 4\"",
    "7015019003": "Pala No. 4", "7015019007": "Palin", "7015019013": "Canastillas", "7015019046": "Desmoñador",
    # Otros (las dos bolsas)
    "5064610001": "Cesta", "5064610002": "Desarene",
    # tomadas de la columna «Concepto» del SISPO de septiembre y de la lista de ítems
    "1801012631": "Aro 20 aluminio", "1941011215": "Toalla de retazo (kilo)", "5064710002": "Cabos cepillo", "52748010135": "Cepíllos",
    "7015019018": "Cabo pala", "7015019022": "Cabos rastrillos",
}

COLUMNAS = ["cuenta", "grupo", "fecha", "articulo", "cantidad", "costo_con_iva", "codigo", "id_kardex", "cierre", "empresa_txt", "tipo_mov",
            "bodega", "cc_codigo", "cc_nombre", "concepto"]


# ================================================================== lectura y clasificación
def leer(archivo) -> pd.DataFrame:
    """Lee el SISPO con el lector único (`logica_m.sispo_leer`, el mismo de Mantenimiento) y deja las columnas que usa Operaciones. A diferencia
    de Mantenimiento, aquí «Grupo», «IdKardex» y «Codigo Articulo» son obligatorios. El centro de costo se pasa a mayúsculas («db010500»
    de C.C. MANUAL es el mismo centro que «DB010500»)."""
    d = L.sispo_leer(archivo)
    faltan = ["Grupo"] if d["grupo"].eq("").all() else []
    marca = L.sispo_sin_columnas_clave(d)
    faltan += (["IdKardex"] if marca["sin_id"] else []) + (["Codigo Articulo"] if marca["sin_codigo"] else [])
    if faltan:
        raise ValueError("Al SISPO le faltan columnas: " + ", ".join(f"«{n}»" for n in faltan) + ".")
    d = d[COLUMNAS].copy()
    d["cc_codigo"] = d["cc_codigo"].str.upper()
    return d.reset_index(drop=True)


def _codigo_cc(x: pd.Series) -> pd.Series:
    return x.astype(str).str.replace(r"\s+", "", regex=True).str.upper()


def _mascaras(d: pd.DataFrame) -> dict:
    """Un filtro por entrada (True = pasa). Sirve igual para el archivo original y para las líneas guardadas. Si `d` trae la columna `cruza`
    (la referencia coincide con un ítem del presupuesto), el filtro de grupo la deja pasar."""
    cuentas = {L.clase20(CUENTA_PAPELERIA), L.clase20(CUENTA_OTROS)}
    cruza = d["cruza"].astype(bool) if "cruza" in d else pd.Series(False, index=d.index)
    if "tipo_mov" in d and d["tipo_mov"].fillna("").astype(str).str.strip().ne("").any():        # solo salidas (el archivo trae Entradas y Salidas)
        salida = d["tipo_mov"].map(L._canon).str.startswith("SALIDA")
    else:
        salida = pd.Series(True, index=d.index)
    es_otros = d["cuenta"].map(L.clase20) == L.clase20(CUENTA_OTROS)       # SAP carga TODA la cuenta OTROS del centro en 3.1: sin filtro de grupo
    return {"salida": salida, "centro": _codigo_cc(d["cc_codigo"]).isin(CENTROS_AUTORIZADOS),
            "grupo": d["grupo"].map(L._canon).isin(GRUPOS_AUTORIZADOS) | cruza | es_otros,
            "cuenta": d["cuenta"].map(L.clase20).isin(cuentas),
            "staf_mtto": ~d["cc_nombre"].map(L._canon).str.contains(PATRON_EXCLUIDO, regex=True)}


def items_presupuestados(ppto: pd.DataFrame) -> set:
    """Nombres canónicos de TODOS los ítems del presupuesto cargado (cualquier hoja): contra esto se decide si una referencia «cruza»."""
    if ppto is None or ppto.empty:
        return set()
    return {k for k in ppto["item"].map(L._canon) if k}


def resolver_items(lin: pd.DataFrame, mapa: dict, presup: set | None = None) -> pd.DataFrame:
    """Agrega `item_res` (ítem del presupuesto al que corresponde la línea, o vacío) y `cruza` (True si ese ítem está presupuestado).
    Prioridad: la columna «Concepto» del SISPO si coincide con un ítem presupuestado; si no, la tabla de equivalencias por código."""
    if lin is None or lin.empty:
        return lin.assign(item_res=pd.Series(dtype=str), cruza=pd.Series(dtype=bool)) if lin is not None else lin
    conc = lin["concepto"].fillna("").astype(str) if "concepto" in lin else pd.Series("", index=lin.index)
    por_codigo = lin["codigo"].map(lambda c: mapa.get(str(c), ""))
    if presup is None:
        item = conc.where(conc != "", por_codigo)
    else:
        item = conc.where(conc.map(L._canon).isin(presup), por_codigo)
    cruza = item.map(L._canon).isin(presup) if presup else pd.Series(False, index=lin.index)
    return lin.assign(item_res=item, cruza=cruza)


def depurar(d: pd.DataFrame) -> tuple[pd.DataFrame, list]:
    """Filtros previos sobre el SISPO original, en este orden. Devuelve (líneas que pasan todos, embudo [(paso, filas que quedan)])."""
    m = _mascaras(d)
    embudo, ok = [("Filas del archivo", len(d))], pd.Series(True, index=d.index)
    for clave, etiqueta in (("salida", "Solo salidas (Tipo Movimiento = Salida)"), ("centro", "Centro de costo autorizado (5 códigos)"), ("grupo", "Grupo 5 INSUMOS o 7 HERRAMIENTAS (o ítem del presupuesto o cuenta OTROS)"),
                            ("cuenta", "Cuenta ÚTILES, PAPELERÍA Y OTROS u OTROS"), ("staf_mtto", "Sin Staf Mantenimiento")):
        ok &= m[clave]
        embudo.append((etiqueta, int(ok.sum())))
    return d[ok], embudo


def filtrar_guardado(lin: pd.DataFrame, mapa: dict, presup: set | None) -> pd.DataFrame:
    """Los mismos filtros aplicados a las líneas ya guardadas (no hace falta recargar el SISPO para que el recorte aplique). Deja las
    columnas `item_res` y `cruza` para el resto del cálculo."""
    if lin is None or lin.empty:
        return lin
    lin = resolver_items(lin, mapa, presup)
    m = _mascaras(lin)
    return lin[m["salida"] & m["centro"] & m["grupo"] & m["cuenta"] & m["staf_mtto"]].reset_index(drop=True)


def clasificar(d: pd.DataFrame, mapa: dict | None = None, presup: set | None = None) -> tuple[pd.DataFrame, int]:
    """Agrega `hoja` (vacía si la línea no es de Operaciones). Primero se depura el archivo original (`depurar`) y solo sobre lo que pasa se
    aplican las reglas de cada hoja. Devuelve (líneas, solapes): los solapes son líneas que cumplirían más de una regla; se asignan a una sola
    (Otros > Consumibles > Herramientas > Papelería) y se informan."""
    d = resolver_items(d, dict(MAPEO_INICIAL) if mapa is None else mapa, presup)
    dep, _ = depurar(d)
    cuenta, grupo = dep["cuenta"].map(L.clase20), dep["grupo"].map(L._canon)
    sale = dep["costo_con_iva"] < 0                                                        # salida = número negativo
    c_otros, bolsa = cuenta == L.clase20(CUENTA_OTROS), dep["codigo"].isin(CODIGOS_OTROS)
    reglas = {HOJA_PAPELERIA: (cuenta == L.clase20(CUENTA_PAPELERIA)) & (dep["cantidad"] < 0),
              HOJA_HERRAMIENTAS: (grupo == GRUPO_HERRAMIENTAS) & c_otros & sale,
              HOJA_CONSUMIBLES: ((grupo == GRUPO_INSUMOS) | (dep["cruza"] & (grupo != GRUPO_HERRAMIENTAS))) & c_otros & sale & ~bolsa,
              HOJA_OTROS: bolsa & sale}
    solapes = int((sum(m.astype(int) for m in reglas.values()) > 1).sum())
    resto = c_otros & sale & ~(sum(m.astype(int) for m in reglas.values()) > 0)              # otros grupos con cuenta OTROS (ver `observacion`)
    reglas[HOJA_HERRAMIENTAS] = reglas[HOJA_HERRAMIENTAS] | resto
    hoja = pd.Series("", index=d.index)
    for h, m in reglas.items():                                                           # el último que se escribe manda
        hoja.loc[m[m].index] = h
    return d.assign(hoja=hoja), solapes


def observacion(lin: pd.DataFrame) -> pd.Series:
    """Texto para las líneas que entraron solo por la cuenta OTROS: su grupo no es 5 ni 7 y no cruzan con un ítem del presupuesto. Va a
    Herramientas porque ahí SAP carga el genérico. Se calcula al mostrar (grupo y cuenta ya están guardados)."""
    vacio = pd.Series("", index=lin.index)
    if lin is None or lin.empty or "grupo" not in lin or "cuenta" not in lin:
        return vacio
    cruza = lin["cruza"].astype(bool) if "cruza" in lin else pd.Series(False, index=lin.index)
    fuera = ~lin["grupo"].map(L._canon).isin(GRUPOS_AUTORIZADOS) & ~cruza & (lin["cuenta"].map(L.clase20) == L.clase20(CUENTA_OTROS))
    return vacio.where(~fuera, "Grupo " + lin["grupo"].astype(str) + ": entra por la cuenta OTROS")


def total_otros_c(d: pd.DataFrame) -> int:
    """Centavos de las salidas de TODA la cuenta OTROS (cualquier grupo): es lo que SAP carga como un solo genérico."""
    s = d[(d["cuenta"].map(L.clase20) == L.clase20(CUENTA_OTROS)) & (d["costo_con_iva"] < 0)]
    return int(sum(L.centavos(-v) for v in s["costo_con_iva"]))


def para_guardar(d: pd.DataFrame, empresa: str, periodo: int) -> pd.DataFrame:
    """Solo las líneas que son de alguna hoja. `valor` con el signo de SAP (costo positivo) y `cantidad` en positivo."""
    x = d[d["hoja"] != ""]
    return pd.DataFrame({
        "empresa": empresa, "periodo": int(periodo), "hoja": x["hoja"], "id_kardex": x["id_kardex"],
        "fecha": x["fecha"].dt.strftime("%Y-%m-%d %H:%M:%S"), "bodega": x["bodega"], "codigo": x["codigo"], "articulo": x["articulo"],
        "grupo": x["grupo"], "cuenta": x["cuenta"], "cc_codigo": x["cc_codigo"], "cc_nombre": x["cc_nombre"],
        "cantidad": x["cantidad"].abs(), "valor": (-x["costo_con_iva"]).where(x["hoja"] != HOJA_PAPELERIA, 0.0),
        "concepto": x["concepto"]}).reset_index(drop=True)


def resumen(d: pd.DataFrame) -> pd.DataFrame:
    """Por hoja: líneas, referencias, cantidad y valor (con signo SAP). Para validar a ojo antes de guardar."""
    x = d[d["hoja"] != ""]
    g = x.groupby("hoja").agg(lineas=("id_kardex", "size"), referencias=("codigo", "nunique"), cantidad=("cantidad", lambda s: s.abs().sum()),
                              valor=("costo_con_iva", lambda s: -s.sum())).reindex(
        [HOJA_PAPELERIA, HOJA_CONSUMIBLES, HOJA_HERRAMIENTAS, HOJA_OTROS]).fillna(0).reset_index()
    g.loc[g["hoja"] == HOJA_PAPELERIA, "valor"] = 0.0
    return g


# ================================================================== equivalencias referencia -> ítem
def mapa_items(filas) -> dict:
    """Equivalencias vigentes: las de partida, sobre las que mandan las guardadas (un ítem vacío quita la equivalencia)."""
    m = dict(MAPEO_INICIAL)
    for cod, item in filas:
        cod = str(cod).strip()
        if not cod:
            continue
        if str(item or "").strip():
            m[cod] = " ".join(str(item).split())
        else:
            m.pop(cod, None)
    return m


def hojas_de_items(ppto: pd.DataFrame) -> dict:
    """{ítem canónico: hoja} según las listas del presupuesto cargado, solo para Consumibles y Herramientas. Un ítem que aparezca en
    las dos hojas es ambiguo y se omite (esa referencia queda con la regla por Grupo)."""
    if ppto is None or ppto.empty:
        return {}
    p = ppto[ppto["hoja"].isin([HOJA_CONSUMIBLES, HOJA_HERRAMIENTAS])]
    por = p.assign(k=p["item"].map(L._canon)).groupby("k")["hoja"].agg(lambda s: set(s))
    return {k: next(iter(h)) for k, h in por.items() if k and len(h) == 1}


def hoja_por_item(lin: pd.DataFrame, mapa: dict, item_hoja: dict) -> pd.DataFrame:
    """La lista de ítems del presupuesto manda sobre el Grupo del SISPO: una referencia con equivalencia a un ítem que está presupuestado
    en Consumibles o en Herramientas va a esa hoja. Solo mueve líneas que ya están en Consumibles/Herramientas (cuenta OTROS: son las que
    SAP carga como genérico); no toca Papelería ni Otros. Se aplica al mostrar, sin volver a cargar el SISPO."""
    if lin is None or lin.empty or not item_hoja:
        return lin
    item = lin["item_res"] if "item_res" in lin else lin["codigo"].map(lambda c: mapa.get(str(c), ""))
    nueva = item.map(lambda i: item_hoja.get(L._canon(i), "") if i else "")
    mueve = lin["hoja"].isin([HOJA_CONSUMIBLES, HOJA_HERRAMIENTAS]) & (nueva != "") & (nueva != lin["hoja"])
    return lin.assign(hoja=lin["hoja"].where(~mueve, nueva))


# ================================================================== aplicar al libro de ejecución
def _columnas_vacias() -> pd.DataFrame:
    return pd.DataFrame(columns=["empresa", "anio", "mes", "fecha", "documento", "tercero", "texto", "cuenta_nombre", "cc", "cc_nombre", "valor",
                                 "cantidad", "hoja", "gerencia", "sub_rubro_sap", "sispo", "codigo_articulo", "item", "item_vista", "grupo"])


def _vista_lin(l: pd.DataFrame, mapa: dict) -> tuple[pd.Series, pd.Series]:
    """Ítem de cada línea: el ya resuelto (`item_res`: Concepto del SISPO o equivalencia) y, si no hay, «<artículo> · sin ítem»."""
    if "item_res" in l:
        item = l["item_res"].fillna("")
        return item, item.where(item != "", l["articulo"].astype(str) + " · sin ítem")
    return _vista_item(l["articulo"], l["codigo"], mapa)


def _vista_item(articulo: pd.Series, codigo: pd.Series, mapa: dict) -> tuple[pd.Series, pd.Series]:
    item = codigo.map(lambda c: mapa.get(str(c), ""))
    return item, item.where(item != "", articulo.astype(str) + " · sin ítem")


def papeleria(lin: pd.DataFrame, cfg: dict, empresa: str, mapa: dict) -> pd.DataFrame:
    """Salidas de Papelería (cantidades) de la compañía y el escenario, con la forma de un movimiento: `valor` = cantidad, para que
    se calcule igual que lo demás. Sin gerencia: solo la ven quienes ven toda la compañía."""
    if lin.empty:
        return _columnas_vacias()
    m0, m1 = cfg["meses"]
    l = lin[(lin["empresa"] == empresa) & (lin["hoja"] == HOJA_PAPELERIA)]
    l = l[(l["periodo"] // 100 == cfg["anio"]) & (l["periodo"] % 100 >= m0) & (l["periodo"] % 100 <= m1)].reset_index(drop=True)
    if l.empty:
        return _columnas_vacias()
    item, vista = _vista_lin(l, mapa)
    return pd.DataFrame({
        "empresa": empresa, "anio": (l["periodo"] // 100).astype(int), "mes": (l["periodo"] % 100).astype(int),
        "fecha": pd.to_datetime(l["fecha"], errors="coerce"),
        "documento": "Kardex " + l["id_kardex"].astype(str), "tercero": l["bodega"], "texto": l["articulo"],
        "cuenta_nombre": "SISPO · " + l["cuenta"], "cc": l["cc_codigo"], "cc_nombre": l["cc_nombre"], "valor": l["cantidad"],
        "cantidad": l["cantidad"], "hoja": HOJA_PAPELERIA, "gerencia": "", "sub_rubro_sap": "", "sispo": True,
        "codigo_articulo": l["codigo"], "item": item, "item_vista": vista, "grupo": l["grupo"]})
