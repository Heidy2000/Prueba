"""
Módulo Operaciones (rubro 3.1): presupuesto vs ejecución por hoja del libro «3.1 operaciones»
==============================================================================================
Una pantalla con cinco pestañas:
- Resumen       <- las cuatro hojas lado a lado (más la ejecución del rubro 3.1 que no cae en ninguna hoja).
- Papelería · Consumibles · Herramientas · Otros
                <- una pestaña por hoja del libro: presupuesto mes a mes (y por ítem) contra la ejecución, con variación,
                   % de ejecución, acumulados, movimientos y Excel. La hoja «Bolsas» del libro pasó a llamarse «Otros».

Qué compara y a qué nivel
- Presupuesto: el del libro (mes a mes y por ítem: cantidad × costo unitario = total). Se carga por compañía (ALDC / ALSM).
- Ejecución: SOLO el SISPO del mes (ver sispo_oper.py; se carga una sola vez en Mantenimiento y se guarda también aquí), sin valores de SAP.
  Se reparte así: Consumibles (grupo 5, cuenta OTROS, sin las 2 bolsas), Herramientas (grupo 7, cuenta OTROS) y Otros (las 2 bolsas), en pesos.
  Papelería sale del SISPO (centros Staf Operaciones, Aguas de Bogotá y Planeación) y se compara en CANTIDADES (costo cero).
  Con el SISPO, el detalle por ítem también tiene ejecución: la referencia se liga a su ítem con la tabla de equivalencias
  (editable); lo ejecutado sin ítem presupuestado aparece en la tabla con presupuesto 0.
- ¿Qué línea va a qué hoja? La que le da el SISPO (reglas de sispo_oper.py y equivalencias). El sub-rubro de SAP solo asigna hoja al PRESUPUESTO del
  módulo Presupuesto (si el nombre contiene PAPELERIA, CONSUMIBLE, HERRAMIENTA o BOLSA se asigna solo; si no, el administrador lo asigna en
  «Asignación de sub-rubros»).
- Ejecución = SOLO SISPO: Consumibles, Herramientas y Otros (en pesos) y Papelería (en cantidades) salen del detalle del SISPO de cada mes, cuadre
  o no con SAP; SAP no aporta ningún valor. Solo presta la gerencia (cada línea hereda la de SAP en el mismo centro de costo). La comparación de
  la cuenta OTROS contra el genérico de SAP es informativa (avisos y «Control» del Resumen). Un mes sin SISPO aplicado queda en cero, y se avisa si
  SAP trae «Salida de inventario» ese mes. El SAP con clase OTROS y objeto Staf Mantenimiento no interviene (el SISPO ya excluye Staf Mantenimiento;
  esa ejecución se cuenta en Mantenimiento, Flota y Equipo de transporte).
- Escenarios (los mismos de Mantenimiento): UPA 2026 = sep–dic 2026; Normal 2027 = ene–dic 2027. Nunca se mezclan.
- «Presupuesto al corte» = presupuesto de los meses hasta «Ejecución hasta». El % de ejecución se calcula contra ese valor
  (comparar lo ejecutado a octubre contra el presupuesto de todo el año subestimaría la ejecución). El saldo sí es contra
  el presupuesto del escenario completo.
- Permisos: quien no ve toda la compañía solo ve las hojas cuyos sub-rubros tiene asignados su gerencia, y solo sus movimientos.
  La carga del libro y la asignación de sub-rubros son solo del administrador (candado en la lógica, no solo en pantalla).

Tablas propias: oper_ppto (presupuesto del libro, formato largo), oper_mapeo (sub-rubro -> hoja), oper_sispo / oper_sispo_periodo
(SISPO de Operaciones) y oper_item_map (referencia -> ítem).
app.py lo llama como: mod.operaciones.mostrar(ctx)
"""
import re
import traceback
from datetime import date, datetime
from io import BytesIO
from types import SimpleNamespace

import openpyxl
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import asignacion
import cargas
import datos
import estilo as E
import logica_m as L
import sispo_oper as OS
from ui import EMPRESAS, MESES, MESES_L, ahora, estilo_grafico, titulo_tarjeta

VERSION = "v3 · 9-nov-2026"          # se muestra en el encabezado: sirve para confirmar qué versión está corriendo
RUBRO_COD = "3.1"
HOJAS = [OS.HOJA_PAPELERIA, OS.HOJA_CONSUMIBLES, OS.HOJA_HERRAMIENTAS, OS.HOJA_OTROS]
RAIZ_HOJA = {"PAPELERIA": "Papelería", "CONSUMIBLE": "Consumibles", "HERRAMIENTA": "Herramientas", "BOLSA": "Otros"}   # sub-rubros de SAP
RAIZ_LIBRO = {**RAIZ_HOJA, "OTROS": "Otros"}                    # nombres de hoja del libro (acepta «Bolsas» u «Otros»)
SIN_HOJA = OS.SIN_HOJA
TODAS = "Consolidado"

TABLAS = """
CREATE TABLE IF NOT EXISTS oper_ppto(empresa TEXT, hoja TEXT, item TEXT, orden INT, anio INT, mes INT,
    cantidad DOUBLE PRECISION, costo_unit DOUBLE PRECISION, valor DOUBLE PRECISION);
CREATE TABLE IF NOT EXISTS oper_mapeo(sub_norm TEXT, sub_rubro TEXT, hoja TEXT, usuario TEXT, fecha TEXT);
CREATE TABLE IF NOT EXISTS oper_sispo(empresa TEXT, periodo INT, hoja TEXT, id_kardex TEXT, fecha TEXT, bodega TEXT, codigo TEXT,
    articulo TEXT, grupo TEXT, cuenta TEXT, cc_codigo TEXT, cc_nombre TEXT, cantidad DOUBLE PRECISION, valor DOUBLE PRECISION, concepto TEXT);
CREATE TABLE IF NOT EXISTS oper_sispo_periodo(empresa TEXT, periodo INT, archivo TEXT, usuario TEXT, fecha TEXT, n_lineas INT,
    total_otros DOUBLE PRECISION);
CREATE TABLE IF NOT EXISTS oper_item_map(codigo TEXT, item TEXT, usuario TEXT, fecha TEXT);
"""


def _asegurar_tablas(ctx):
    if not st.session_state.get("_oper_tablas2"):
        ctx.con.crear_tablas(TABLAS)
        ctx.con.commit()
        try:                                              # bases creadas antes de existir la columna «concepto»
            ctx.con.execute("SELECT concepto FROM oper_sispo LIMIT 1")
        except Exception:
            try:
                ctx.con.rollback()
            except Exception:
                pass
            ctx.con.execute("ALTER TABLE oper_sispo ADD COLUMN concepto TEXT")
            ctx.con.commit()
        st.session_state["_oper_tablas2"] = True


def _exigir_admin(ctx):
    """Candado de permisos en la LÓGICA: toda escritura en la base exige administrador."""
    if not getattr(ctx, "es_admin", False):
        raise PermissionError("Acción reservada a administradores.")


# ================================================================== lectura del libro de presupuesto
def hoja_canonica(nombre, raices=None) -> str | None:
    """«Papeleria», «Papelería y útiles», «CONSUMIBLES»… -> nombre canónico de la hoja; None si no se reconoce."""
    k = re.sub(r"[^A-Z0-9]", "", L.sin_tildes(nombre).upper())
    return next((h for raiz, h in (raices or RAIZ_HOJA).items() if raiz in k), None)


def _num(v) -> float:
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else float("nan")


def _tol(t: float) -> float:
    return max(1.0, abs(t) * 1e-6)


def _puntaje(grupos, filas) -> float:
    """Fracción de celdas donde cantidad × costo unitario = total (con tolerancia de un peso)."""
    ok = tot = 0
    for f in filas:
        for q, u, t in grupos:
            vq, vu, vt = (_num(f[c]) if c < len(f) else float("nan") for c in (q, u, t))
            if pd.isna(vt):
                continue
            tot += 1
            if (not pd.isna(vq) and not pd.isna(vu) and abs(vq * vu - vt) <= _tol(vt)) or \
                    (abs(vt) < 0.5 and (pd.isna(vq) or pd.isna(vu) or vq == 0 or vu == 0)):
                ok += 1
    return ok / tot if tot else 0.0


def _leer_hoja(ws, hoja: str):
    """Una hoja del libro -> (filas largas, info). Estructura esperada: encabezado con «Mes» en la primera columna, una fila por
    mes (fecha), una fila «Total», y por cada ítem tres columnas (cantidad, costo unitario, total) —o, en Bolsas, primero los
    precios, luego las cantidades y luego los totales—, más una columna final con el total de la hoja. El total de cada ítem es el
    valor del libro; se verifica que cantidad × costo unitario lo reproduzca."""
    filas = list(ws.iter_rows(values_only=True))
    i_enc = next((i for i, f in enumerate(filas[:8]) if f and f[0] is not None and L.norm_texto(f[0]) == "MES"), None)
    if i_enc is None:
        raise ValueError("no encuentro la fila de encabezados (la primera columna debe decir «Mes»)")
    enc = [("" if v is None else " ".join(str(v).split())) for v in filas[i_enc]]
    while enc and not enc[-1]:
        enc.pop()
    meses = [f for f in filas[i_enc + 1:] if f and isinstance(f[0], (datetime, date))]
    fila_total = next((f for f in filas[i_enc + 1:] if f and isinstance(f[0], str) and L.norm_texto(f[0]) == "TOTAL"), None)
    if not meses:
        raise ValueError("no hay filas de meses (fechas en la primera columna)")
    n = len(enc)
    col_total = n - 1 if n > 2 and "TOTAL" in L.norm_texto(enc[-1]) else None
    cuerpo = list(range(1, n - 1 if col_total is not None else n))
    if not cuerpo or len(cuerpo) % 3:
        raise ValueError(f"las columnas de ítems ({len(cuerpo)}) no son múltiplo de 3 (cantidad, costo unitario, total)")
    k = len(cuerpo) // 3
    a = [(cuerpo[3 * j], cuerpo[3 * j + 1], cuerpo[3 * j + 2]) for j in range(k)]            # cantidad, costo, total, consecutivos
    b = [(cuerpo[k + j], cuerpo[j], cuerpo[2 * k + j]) for j in range(k)]                    # precios | cantidades | totales
    pa, pb = _puntaje(a, meses), _puntaje(b, meses)
    if max(pa, pb) < 0.9:
        raise ValueError("la estructura no se reconoce: cantidad × costo unitario no reproduce el total en la mayoría de las celdas")
    grupos, nombre_col = (a, 0) if pa >= pb else (b, 2)
    layout = "A" if pa >= pb else "B"

    nombres, vistos = [], {}
    for q, u, t in grupos:
        nom = enc[(q, u, t)[nombre_col]] or f"Ítem {len(nombres) + 1}"
        vistos[nom] = vistos.get(nom, 0) + 1
        nombres.append(nom if vistos[nom] == 1 else f"{nom} ({vistos[nom]})")

    largo, dif_qxu, dif_mes = [], [], []
    for f in meses:
        suma = 0.0
        for orden, ((q, u, t), nom) in enumerate(zip(grupos, nombres), start=1):
            vq, vu, vt = (_num(f[c]) if c < len(f) else float("nan") for c in (q, u, t))
            if pd.isna(vt):
                if pd.isna(vq) or pd.isna(vu):
                    continue
                vt = vq * vu
            elif not pd.isna(vq) and not pd.isna(vu) and abs(vq * vu - vt) > _tol(vt):
                dif_qxu.append((nom, f[0].strftime("%Y-%m"), vq * vu - vt))
            suma += vt
            largo.append(dict(hoja=hoja, item=nom, orden=orden, anio=f[0].year, mes=f[0].month,
                              cantidad=None if pd.isna(vq) else vq, costo_unit=None if pd.isna(vu) else vu, valor=vt))
        if col_total is not None and col_total < len(f):
            tt = _num(f[col_total])
            if not pd.isna(tt) and abs(tt - suma) > 1:
                dif_mes.append((f[0].strftime("%Y-%m"), tt - suma))
    df = pd.DataFrame(largo)
    por_esc = {esc: float(df.loc[(df["anio"] == c["anio"]) & df["mes"].between(*c["meses"]), "valor"].sum()) for esc, c in L.ESCENARIOS.items()}
    fuera = int((~df.apply(lambda r: any(r["anio"] == c["anio"] and c["meses"][0] <= r["mes"] <= c["meses"][1]
                                          for c in L.ESCENARIOS.values()), axis=1)).sum()) if len(df) else 0
    dif_fila = None
    cant_mal = []
    if fila_total is not None:
        if col_total is not None and col_total < len(fila_total) and not pd.isna(_num(fila_total[col_total])):
            dif_fila = _num(fila_total[col_total]) - float(df["valor"].sum())
        for (q, _, _), nom in zip(grupos, nombres):
            ft = _num(fila_total[q]) if q < len(fila_total) else float("nan")
            suma_q = float(df.loc[df["item"] == nom, "cantidad"].sum())
            if not pd.isna(ft) and abs(ft - suma_q) > 0.5:
                cant_mal.append(nom)
    info = dict(layout=layout, n_items=k, n_meses=len(meses), total=float(df["valor"].sum()), por_escenario=por_esc,
                dif_qxu=dif_qxu, dif_mes=dif_mes, dif_fila_total=dif_fila, cant_total_no_cuadra=cant_mal,
                filas_fuera_escenario=fuera)
    return df, info


def leer_libro(archivo):
    """Libro de presupuesto de Operaciones -> (filas largas de todas las hojas reconocidas, info por hoja)."""
    wb = openpyxl.load_workbook(archivo, data_only=True)
    partes, infos, ignoradas = [], {}, []
    for ws in wb.worksheets:
        hoja = hoja_canonica(ws.title, RAIZ_LIBRO)
        if hoja is None or hoja in infos:
            ignoradas.append(ws.title)
            continue
        try:
            df, inf = _leer_hoja(ws, hoja)
            inf["hoja_excel"] = ws.title
            infos[hoja] = inf
            partes.append(df)
        except ValueError as e:
            infos[hoja] = dict(error=str(e), hoja_excel=ws.title)
    largo = pd.concat(partes, ignore_index=True) if partes else pd.DataFrame(
        columns=["hoja", "item", "orden", "anio", "mes", "cantidad", "costo_unit", "valor"])
    return largo, dict(hojas=infos, ignoradas=ignoradas, faltan=[h for h in HOJAS if h not in infos])


# ================================================================== datos de la base
def _version_para(ctx, empresa, anio):
    vs = ctx.leer("SELECT DISTINCT version FROM presupuesto WHERE anio=? AND empresa=?", (int(anio), empresa))["version"].tolist()
    orden = {"Inicial": 0, "4+8": 1, "8+4": 2}
    vs = sorted(vs, key=lambda v: (orden.get(v, 9), v))
    return "8+4" if "8+4" in vs else (vs[-1] if vs else "Inicial")


def _base(ctx, empresa, anio, version):
    """Presupuesto (tabla presupuesto) y ejecución del año con el mismo tratamiento que el módulo Presupuesto (homologación de
    sub-rubros), para que los totales coincidan."""
    leer = ctx.leer
    ppto = leer("SELECT * FROM presupuesto WHERE anio=? AND version=? AND empresa=?", (int(anio), version, empresa))
    cols_e = ", ".join(sorted(ctx.con.columnas("ejecucion") - {"raw"}))
    ejec = leer(f"SELECT {cols_e} FROM ejecucion WHERE anio=? AND empresa=?", (int(anio), empresa))
    ejec = ejec[ejec["codigo"].fillna("") != ""].copy()
    ejec["rubro"] = ejec["codigo"].map(cargas.rubro_de)
    ejec["sub_rubro_sap"] = ejec["sub_rubro"]
    if not ejec.empty:
        ejec = datos.homologar(ejec, ppto, leer("SELECT * FROM homologacion"), version)
    ppto["_sub"] = ppto["sub_rubro"].fillna("").map(cargas.norm)
    ejec["_sub"] = ejec["sub_rubro"].fillna("").map(cargas.norm)
    return ppto, ejec


def _es_rubro(codigos: pd.Series) -> pd.Series:
    return codigos.fillna("").astype(str).str.strip().str.replace(",", ".", regex=False) == RUBRO_COD


def _hojas_de(df: pd.DataFrame, mapeo: dict) -> list:
    """Hoja de cada fila: la asignación guardada manda; si no hay, por el nombre del sub-rubro; si no, «Sin hoja asignada»."""
    out = []
    for sub_norm, nombre in zip(df["_sub"], df["sub_rubro"]):
        out.append((mapeo.get(sub_norm) or SIN_HOJA) if sub_norm in mapeo else (hoja_canonica(nombre) or SIN_HOJA))
    return out


def _solo_sispo(ejec_sap: pd.DataFrame, ejec_all: pd.DataFrame, lin: pd.DataFrame, per: pd.DataFrame, empresa: str, items: dict, cfg: dict,
                mtto_per: set = frozenset()):
    """La ejecución de Operaciones sale SOLO del SISPO (Consumibles, Herramientas y Otros en pesos): no se usa ningún valor de SAP.
    Se toma todo periodo con SISPO guardado, cuadre o no con SAP. SAP solo presta la GERENCIA: cada línea hereda la de los movimientos
    de SAP del mismo centro de costo (o la predominante). La comparación con la cuenta OTROS de SAP es informativa y va en `avisos`.
    Devuelve (ejecución, avisos); la ejecución trae las columnas que usa el resto del módulo."""
    avisos, nuevos = [], []
    m0, m1 = cfg["meses"]
    ejec_all = ejec_all.assign(empresa=empresa)
    for p in per.itertuples():
        pn = int(p.periodo)
        sap_all = _generico_otros_sap(ejec_all, empresa, pn)
        sis_c = L.centavos(p.total_otros)
        if sap_all.empty:
            avisos.append(f"{empresa} {pn}: se toma el SISPO; SAP todavía no trae la «Salida de inventario» de la cuenta OTROS para compararlo.")
        else:
            sap_c = int(sum(L.centavos(v) for v in pd.to_numeric(sap_all["valor"], errors="coerce").fillna(0.0)))
            if sap_c != sis_c:
                avisos.append(f"{empresa} {pn}: se toma el SISPO, pero la cuenta OTROS difiere de SAP: SAP {sap_c / 100:,.2f}, SISPO {sis_c / 100:,.2f} "
                              f"(diferencia {(sap_c - sis_c) / 100:,.2f}).")
        l = lin[(lin["periodo"] == pn) & lin["hoja"].isin(OS.HOJAS_PESOS)].reset_index(drop=True)
        if l.empty:
            continue
        item, vista = OS._vista_lin(l, items)
        ger = L.sispo_destinos(l, ejec_sap)["gerencia"]
        nuevos.append(pd.DataFrame({
            "empresa": empresa, "anio": pn // 100, "mes": pn % 100, "fecha": pd.to_datetime(l["fecha"], errors="coerce"),
            "documento": "Kardex " + l["id_kardex"].astype(str), "tercero": l["bodega"], "texto": l["articulo"],
            "cuenta_nombre": "SISPO · " + l["cuenta"], "cc": l["cc_codigo"], "cc_nombre": l["cc_nombre"], "codigo": RUBRO_COD,
            "sub_rubro": l["hoja"], "sub_rubro_sap": "", "valor": l["valor"], "cantidad": l["cantidad"], "hoja": l["hoja"], "gerencia": ger,
            "sispo": True, "codigo_articulo": l["codigo"], "item": item, "item_vista": vista, "grupo": l["grupo"], "pend_sispo": False}))
    con_sispo = {int(x) for x in per["periodo"]}
    for m in range(m0, m1 + 1):                                  # meses con genérico en SAP y sin SISPO: Operaciones queda en cero ahí
        pn = cfg["anio"] * 100 + m
        if pn not in con_sispo and not _generico_otros_sap(ejec_all, empresa, pn).empty:
            if (empresa, pn) in mtto_per:
                avisos.append(f"{empresa} {pn}: Mantenimiento ya aplicó este SISPO, pero Operaciones no lo tiene (Mantenimiento no guarda el grupo del "
                              "artículo ni Papelería, que Operaciones necesita). Sube ese mismo archivo en Mantenimiento («Cargar SISPO del mes») y "
                              "pulsa «Aplicar»: solo guarda la parte de Operaciones, no cambia ni duplica nada en Mantenimiento.")
            else:
                avisos.append(f"{empresa} {pn}: SAP registra «Salida de inventario» de la cuenta OTROS, pero no hay SISPO aplicado; Operaciones no muestra "
                              "Consumibles, Herramientas ni Otros de ese mes. Aplícalo en Mantenimiento («Cargar SISPO del mes»).")
    cols = ["empresa", "anio", "mes", "fecha", "documento", "tercero", "texto", "cuenta_nombre", "cc", "cc_nombre", "codigo", "sub_rubro",
            "sub_rubro_sap", "valor", "cantidad", "hoja", "gerencia", "sispo", "codigo_articulo", "item", "item_vista", "grupo", "pend_sispo"]
    ejec = pd.concat(nuevos, ignore_index=True) if nuevos else pd.DataFrame(columns=cols)
    for c in ("valor", "mes", "anio"):
        ejec[c] = pd.to_numeric(ejec[c], errors="coerce")
    ejec["valor"] = ejec["valor"].fillna(0.0)
    ejec["sispo"] = ejec["sispo"].fillna(False).astype(bool)
    ejec["pend_sispo"] = ejec["pend_sispo"].fillna(False).astype(bool)
    return ejec, avisos


def _generico_otros_sap(ejec_all: pd.DataFrame, empresa: str, periodo: int) -> pd.DataFrame:
    """Movimientos de SAP «Salida de inventario» de la cuenta OTROS de un periodo (todos los rubros)."""
    return L.sispo_sap_generico(ejec_all, empresa, periodo, OS.CUENTA_OTROS)


def _datos_empresa(ctx, empresa: str, cfg: dict, asig, mapeo: dict, sispo=None) -> SimpleNamespace:
    version = _version_para(ctx, empresa, cfg["anio"])
    ppto_base, ejec = _base(ctx, empresa, cfg["anio"], version)
    ejec_all = ejec.assign(empresa=empresa)              # todos los rubros: con esto se concilia la cuenta OTROS con el SISPO
    ppto_base = ppto_base[_es_rubro(ppto_base["codigo"])].copy() if len(ppto_base) else ppto_base
    ejec = ejec[_es_rubro(ejec["codigo"])].copy()
    m0, m1 = cfg["meses"]
    ejec = ejec[(ejec["mes"] >= m0) & (ejec["mes"] <= m1)].copy()
    ejec["empresa"] = empresa
    ejec["hoja"] = _hojas_de(ejec, mapeo) if len(ejec) else pd.Series(dtype=str)
    ejec["gerencia"] = asignacion.columna(ejec, asig) if len(ejec) else pd.Series(dtype=str)
    if len(ppto_base):
        ppto_base["hoja"] = _hojas_de(ppto_base, mapeo)
        ppto_base["gerencia"] = asignacion.columna(ppto_base, asig)
    # Ejecución = SOLO SISPO (Consumibles, Herramientas y Otros en pesos); Papelería en cantidades. SAP no aporta valores.
    per = sispo.per[(sispo.per["empresa"] == empresa) & (sispo.per["periodo"] >= cfg["anio"] * 100 + m0)
                    & (sispo.per["periodo"] <= cfg["anio"] * 100 + m1)] if sispo is not None else pd.DataFrame(columns=["empresa", "periodo", "total_otros"])
    lin = sispo.lin[sispo.lin["empresa"] == empresa] if sispo is not None else pd.DataFrame()
    items = sispo.items if sispo is not None else {}
    ejec_sap = ejec.assign(_sub=ejec["sub_rubro"].fillna("").map(cargas.norm))        # SAP rubro 3.1: ya no es ejecución; solo para la asignación de sub-rubros
    try:                                                           # periodos que Mantenimiento ya aplicó (la tabla no existe si nunca se usó)
        mtto_per = {(r.empresa, int(r.periodo)) for r in ctx.leer("SELECT DISTINCT empresa, periodo FROM mtto_sispo_periodo").itertuples()}
    except Exception:
        mtto_per = set()
        try:
            ctx.con.rollback()
        except Exception:
            pass
    ejec, avisos = _solo_sispo(ejec, ejec_all, lin, per, empresa, items, cfg, mtto_per)
    ejec["_sub"] = ejec["sub_rubro"].fillna("").map(cargas.norm)
    return SimpleNamespace(empresa=empresa, ppto_base=ppto_base, ejec=ejec, ejec_sap=ejec_sap, ejec_all=ejec_all, avisos=avisos,
                           lin_pap=OS.papeleria(lin, cfg, empresa, items))


# ================================================================== cálculo
def _tono(est: str):
    return {"Sobre ejecutado": (E.ROJO, E.ROJO_SUAVE), "Sin presupuesto asignado": (E.ROJO, E.ROJO_SUAVE),
            "Presupuesto agotado": (E.AMBAR, E.AMBAR_SUAVE)}.get(est, (E.VERDE, E.VERDE_SUAVE) if est == "Dentro del presupuesto"
                                                                    else (E.AZUL, E.AZUL_SUAVE))


def _pesos(x) -> str:
    return "—" if pd.isna(x) else f"{x:,.0f}".replace(",", ".")


def _ppto_hoja(S, hoja: str) -> pd.DataFrame:
    return S.ppto[S.ppto["hoja"] == hoja]


def _ejec_hoja(S, hoja: str) -> pd.DataFrame:
    return S.ejec[S.ejec["hoja"] == hoja]


def _cifras(S, hoja: str) -> dict:
    p, e = _ppto_hoja(S, hoja), _ejec_hoja(S, hoja)
    p_esc = float(p["valor"].sum())
    p_corte = float(p.loc[p["mes"] <= S.hasta_mes, "valor"].sum())
    ec = e[e["mes"] <= S.hasta_mes]
    ejecutado, n = float(ec["valor"].sum()), len(ec)
    t = L.con_metricas(pd.DataFrame({"presupuesto": [p_corte], "ejecutado": [ejecutado], "movimientos": [n]})).iloc[0]
    return dict(hoja=hoja, presupuesto_esc=p_esc, presupuesto=p_corte, ejecutado=ejecutado, variacion=ejecutado - p_corte,
                saldo=p_esc - ejecutado, pct=t["pct"], estado=t["estado"], movimientos=n)


def _mensual(S, hoja: str) -> pd.DataFrame:
    """Mes a mes: presupuesto, ejecutado (solo hasta el corte), variación, % y acumulados. Última fila: total al corte."""
    p = _ppto_hoja(S, hoja).groupby("mes")["valor"].sum()
    e = _ejec_hoja(S, hoja).groupby("mes")["valor"].sum()
    filas = []
    for m in range(S.m0, S.m1 + 1):
        hasta = m <= S.hasta_mes
        filas.append(dict(mes=m, presupuesto=float(p.get(m, 0.0)), ejecutado=float(e.get(m, 0.0)) if hasta else float("nan")))
    t = pd.DataFrame(filas)
    t["variacion"] = t["ejecutado"] - t["presupuesto"]
    t["pct"] = (t["ejecutado"] / t["presupuesto"]).where(t["presupuesto"] > 0)
    t["ppto_acum"] = t["presupuesto"].cumsum().where(t["ejecutado"].notna())
    t["ejec_acum"] = t["ejecutado"].fillna(0).cumsum().where(t["ejecutado"].notna())
    t["pct_acum"] = (t["ejec_acum"] / t["ppto_acum"]).where(t["ppto_acum"] > 0)
    al_corte = t[t["ejecutado"].notna()]
    pt, et = float(al_corte["presupuesto"].sum()), float(al_corte["ejecutado"].sum())
    t["Mes"] = t["mes"].map(lambda m: f"{MESES_L[m - 1].capitalize()} {S.cfg['anio']}")
    total = dict(mes=0, Mes="Total al corte", presupuesto=pt, ejecutado=et, variacion=et - pt, pct=(et / pt) if pt > 0 else float("nan"),
                 ppto_acum=float("nan"), ejec_acum=float("nan"), pct_acum=float("nan"))
    return pd.concat([t, pd.DataFrame([total])], ignore_index=True)


def _items(S, hoja: str) -> pd.DataFrame:
    """Presupuesto del libro por ítem y, al lado, la ejecución del SISPO ligada a su ítem (tabla de equivalencias). Lo ejecutado cuyo
    ítem no está presupuestado en esta hoja (o que no tiene ítem) aparece igual, con presupuesto 0. Lo que viene de SAP sin detalle
    de SISPO va en una fila aparte, para que la suma de la tabla coincida con la de la hoja."""
    p = _ppto_hoja(S, hoja)
    if p.empty:
        g = pd.DataFrame({"orden": pd.Series(dtype=float), "item": pd.Series(dtype=str), "cantidad": pd.Series(dtype=float),
                          "presupuesto_esc": pd.Series(dtype=float), "presupuesto": pd.Series(dtype=float)})
    else:
        g = p.groupby(["orden", "item"], as_index=False).agg(cantidad=("cantidad", "sum"), presupuesto_esc=("valor", "sum"))
        corte = p[p["mes"] <= S.hasta_mes].groupby("item")["valor"].sum()
        g["presupuesto"] = g["item"].map(corte).fillna(0.0)
    g["_k"] = g["item"].map(L._canon)
    e = _ejec_hoja(S, hoja)
    e = e[e["mes"] <= S.hasta_mes]
    es_s = e["sispo"].fillna(False).astype(bool) if "sispo" in e else pd.Series(False, index=e.index)
    s = e[es_s]
    ex = (s.assign(_k=s["item_vista"].map(L._canon)).groupby("_k").agg(item_e=("item_vista", "first"), ejecutado=("valor", "sum")).reset_index()
          if len(s) else pd.DataFrame({"_k": pd.Series(dtype=str), "item_e": pd.Series(dtype=str), "ejecutado": pd.Series(dtype=float)}))
    g = g.merge(ex, on="_k", how="outer")
    g["item"] = g["item"].fillna(g["item_e"])
    for c in ("cantidad", "presupuesto_esc", "presupuesto", "ejecutado"):       # float de verdad: con dtype object, dividir por 0 lanza error
        g[c] = pd.to_numeric(g[c], errors="coerce").fillna(0.0).astype(float)
    sap = float(e.loc[~es_s, "valor"].sum())
    if abs(sap) >= 0.5:
        g = pd.concat([g, pd.DataFrame([dict(item="Otros movimientos de SAP (sin ítem)", cantidad=0.0, presupuesto_esc=0.0,
                                             presupuesto=0.0, ejecutado=sap)])], ignore_index=True)
    g["costo_unit"] = g["presupuesto_esc"] / g["cantidad"].where(g["cantidad"] > 0)
    tot = g["presupuesto_esc"].sum()
    g["peso"] = g["presupuesto_esc"] / tot if tot > 0 else float("nan")
    g["variacion"] = g["ejecutado"] - g["presupuesto"]
    g["pct"] = g["ejecutado"] / g["presupuesto"].where(g["presupuesto"] > 0)
    return g.sort_values(["presupuesto_esc", "ejecutado"], ascending=False)[
        ["item", "cantidad", "costo_unit", "presupuesto_esc", "presupuesto", "ejecutado", "variacion", "pct", "peso"]]


# ================================================================== piezas de pantalla
def _fmt(S, x) -> str:
    """Cifra para tarjetas: millones de pesos, o unidades en la vista de cantidades (Papelería)."""
    if getattr(S, "cant", False):
        return "—" if pd.isna(x) else f"{x:,.0f}".replace(",", ".") + " und"
    return E.millones(x)


def _vista_cantidades(S):
    """Misma pantalla vista en CANTIDADES (Papelería): el presupuesto es la cantidad del libro y la ejecución, las salidas del SISPO.
    Se reutilizan los mismos cálculos; solo cambia qué se suma (`valor` = cantidad) y cómo se muestra."""
    q = SimpleNamespace(**vars(S))
    q.ppto = S.ppto.assign(valor=S.ppto["cantidad"].fillna(0.0))
    q.ejec = S.lin_pap
    q.cant = True
    return q


def _kpis(c: dict, S):
    color, fondo = _tono(c["estado"])
    st.html(E.kpis([
        dict(label="Presupuesto del escenario", valor=_fmt(S, c["presupuesto_esc"]), sub=S.escenario, icono="wallet",
             color=E.AZUL, fondo=E.AZUL_SUAVE),
        dict(label="Presupuesto al corte", valor=_fmt(S, c["presupuesto"]), sub=f"meses hasta {MESES_L[S.hasta_mes - 1]}",
             icono="wallet", color=E.AZUL, fondo=E.AZUL_SUAVE),
        dict(label=f"Ejecutado {S.hasta}", valor=_fmt(S, c["ejecutado"]),
             sub=f"{c['movimientos']:,} movimientos".replace(",", "."), icono="receipt", color=E.AZUL, fondo=E.AZUL_SUAVE),
        dict(label="% ejecución vs presupuesto al corte", valor=f'<span style="color:{color}">{E.pct_txt(c["pct"])}</span>',
             sub=f"{c['estado']} · saldo del escenario {_fmt(S, c['saldo'])}",
             icono="alert" if c["estado"] in ("Sobre ejecutado", "Sin presupuesto asignado") else "check", color=color, fondo=fondo),
    ]))


def _color_var(v):
    """Variación = ejecutado − presupuesto: positiva (gastó de más) en rojo, negativa (por debajo) en verde."""
    return "" if pd.isna(v) or abs(v) < 0.5 else (f"color:{E.ROJO};font-weight:600" if v > 0 else f"color:{E.VERDE};font-weight:600")


def _estilo(df: pd.DataFrame, col: str):
    sty = df.style
    return (getattr(sty, "map", None) or sty.applymap)(_color_var, subset=[col])


def _columnas_mes(S=None):
    nc = st.column_config.NumberColumn
    fm = "localized" if getattr(S, "cant", False) else "accounting"
    return {"Mes": "Mes", "presupuesto": nc("Presupuesto", format=fm), "ejecutado": nc("Ejecutado", format=fm),
            "variacion": nc("Variación (ejecutado − presupuesto)", format=fm), "pct": nc("% ejecución", format="percent"),
            "ppto_acum": nc("Presupuesto acumulado", format=fm), "ejec_acum": nc("Ejecutado acumulado", format=fm),
            "pct_acum": nc("% acumulado", format="percent")}


def _graficos_mes(S, t: pd.DataFrame, hoja: str):
    m = t[t["mes"] > 0]
    x = [f"{MESES[mm - 1]} {str(S.cfg['anio'])[2:]}" for mm in m["mes"]]
    div, suf, que = (1.0, " und", "unidades") if S.cant else (1e6, " M", "millones de pesos")
    c1, c2 = st.columns(2)
    with c1:
        with st.container(key=f"op_card_graf_mes_{hoja}"):
            titulo_tarjeta("Mes a mes", f"Presupuesto vs ejecutado ({que})")
            fig = go.Figure()
            fig.add_bar(x=x, y=m["presupuesto"] / div, name="Presupuesto", marker_color="#c9d4e4", marker_line_width=0,
                        hovertemplate="%{y:,.1f}" + suf)
            fig.add_bar(x=x, y=m["ejecutado"] / div, name="Ejecutado", marker_color=E.AZUL, marker_line_width=0,
                        hovertemplate="%{y:,.1f}" + suf)
            fig.update_layout(barmode="group", bargap=.25, yaxis=dict(tickformat=",.0f", ticksuffix=suf, rangemode="tozero"))
            st.plotly_chart(estilo_grafico(fig, 300), width="stretch", config={"displayModeBar": False})
    with c2:
        with st.container(key=f"op_card_graf_acum_{hoja}"):
            titulo_tarjeta("Acumulado", f"Presupuesto acumulado vs ejecución acumulada ({que})")
            fig = go.Figure()
            fig.add_scatter(x=x, y=(m["presupuesto"].cumsum()) / div, name="Presupuesto acumulado", mode="lines+markers",
                            line=dict(color="#9fb0c9", width=2, dash="dot"), hovertemplate="%{y:,.1f}" + suf)
            fig.add_scatter(x=x, y=m["ejec_acum"] / div, name="Ejecutado acumulado", mode="lines+markers",
                            line=dict(color=E.AZUL, width=3), hovertemplate="%{y:,.1f}" + suf)
            fig.update_layout(yaxis=dict(tickformat=",.0f", ticksuffix=suf, rangemode="tozero"))
            st.plotly_chart(estilo_grafico(fig, 300), width="stretch", config={"displayModeBar": False})


def _graf_items(it: pd.DataFrame, S):
    top = it[it["presupuesto_esc"] > 0].head(12).iloc[::-1]
    if top.empty:
        return
    div, suf = (1.0, " und") if S.cant else (1e6, " M")
    fig = go.Figure()
    fig.add_bar(y=[s[:48] for s in top["item"]], x=top["presupuesto_esc"] / div, orientation="h", marker_color=E.AZUL,
                marker_line_width=0, hovertemplate="%{y}: %{x:,.2f}" + suf + "<extra>Presupuesto</extra>", name="Presupuesto del escenario")
    fig.update_layout(xaxis=dict(tickformat=",.1f", ticksuffix=suf), yaxis=dict(type="category"), showlegend=False)
    st.plotly_chart(estilo_grafico(fig, 420), width="stretch", config={"displayModeBar": False})


def _excel(hojas: dict) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for nombre, df in hojas.items():
            df.to_excel(w, sheet_name=nombre[:31], index=False)
    return buf.getvalue()


def _tab_hoja(S, hoja: str):
    if hoja == OS.HOJA_PAPELERIA:                       # Papelería se compara en cantidades (costo cero): mismos cálculos, otra unidad
        S = _vista_cantidades(S)
    p = _ppto_hoja(S, hoja)
    if p.empty and _ejec_hoja(S, hoja).empty:
        st.info(f"No hay presupuesto de «{hoja}» cargado para {S.escenario}"
                + ("." if S.ctx.es_admin else ". El administrador lo carga en «Datos base»."), icon=":material/info:")
        return
    if p.empty:
        st.info(f"No hay presupuesto de «{hoja}» cargado para {S.escenario}: lo ejecutado se muestra contra presupuesto 0.", icon=":material/info:")
    c = _cifras(S, hoja)
    _kpis(c, S)
    t = _mensual(S, hoja)
    with st.container(key=f"op_card_tabla_{hoja}"):
        titulo_tarjeta(f"{hoja} · presupuesto vs ejecución", f"{S.escenario} · ejecutado {S.hasta} · variación = ejecutado − presupuesto"
                       + (" · EN CANTIDADES (unidades): estas referencias tienen costo cero" if S.cant else ""))
        vista = t[["Mes", "presupuesto", "ejecutado", "variacion", "pct", "ppto_acum", "ejec_acum", "pct_acum"]]
        st.dataframe(_estilo(vista, "variacion"), hide_index=True, width="stretch", height=40 + 35 * len(vista), column_config=_columnas_mes(S))
        st.caption("«% ejecución» compara cada mes con su presupuesto. El total y los KPI se calculan contra el presupuesto al corte; "
                   "el saldo, contra el del escenario completo.")
    _graficos_mes(S, t, hoja)
    it = _items(S, hoja)
    with st.container(key=f"op_card_items_{hoja}"):
        titulo_tarjeta("Presupuesto y ejecución por ítem", f"{S.escenario} · "
                       + ("cantidad presupuestada del libro contra salidas del SISPO" if S.cant else
                          "presupuesto del libro (cantidad × costo unitario) contra salidas del SISPO"))
        c_t, c_g = st.columns([1.35, 1])
        with c_t:
            nc = st.column_config.NumberColumn
            fm = "localized" if S.cant else "accounting"
            vis = it.drop(columns=["cantidad", "costo_unit"]) if S.cant else it
            conf = {"item": "Ítem", "cantidad": nc("Cantidad presupuestada", format="localized"),
                    "costo_unit": nc("Costo unit. promedio", format="accounting"),
                    "presupuesto_esc": nc("Presupuesto escenario", format=fm), "presupuesto": nc("Presupuesto al corte", format=fm),
                    "ejecutado": nc("Ejecutado", format=fm), "variacion": nc("Variación", format=fm),
                    "pct": nc("% ejecución", format="percent"), "peso": nc("% del total", format="percent")}
            st.dataframe(_estilo(vis, "variacion"), hide_index=True, width="stretch", height=min(40 + 35 * len(vis), 460),
                         column_config={k: v for k, v in conf.items() if k in vis.columns})
            st.caption("Lo ejecutado que no tiene ítem presupuestado en esta hoja aparece con presupuesto 0 (marcado «sin ítem» si la referencia "
                       "no tiene equivalencia). Las equivalencias referencia → ítem se editan en «Datos base».")
        with c_g:
            _graf_items(it, S)
    e = _ejec_hoja(S, hoja)
    e = e[e["mes"] <= S.hasta_mes]
    cols_mov = ["empresa", "fecha", "mes", "documento", "tercero", "texto", "codigo_articulo", "cuenta_nombre", "sub_rubro_sap", "cc_nombre"]
    if not S.cant:
        cols_mov.append("cantidad")
    mov = e.assign(fecha=pd.to_datetime(e["fecha"], errors="coerce"))[
        [c for c in [*cols_mov, "valor"] if c in e.columns]].sort_values("fecha", ascending=False)
    mov_todos = mov                                       # lo que se descarga es todo, aunque se filtre la vista con la búsqueda
    with st.expander(f"Movimientos que componen la ejecución ({len(mov):,})".replace(",", "."), icon=":material/receipt_long:"):
        if mov.empty:
            st.info("No hay movimientos asignados a esta hoja hasta el corte.")
        else:
            buscar = st.text_input("Buscar", placeholder="Proveedor, texto, documento…", label_visibility="collapsed", key=f"op_buscar_{hoja}")
            if buscar:
                mov = mov[mov.astype(str).apply(lambda col: col.str.contains(buscar, case=False, regex=False)).any(axis=1)]
            total_txt = f"{mov['valor'].sum():,.0f} unidades".replace(",", ".") if S.cant else E.pesos(mov["valor"].sum())
            st.html(f'<p class="nf-card-s">{len(mov):,} movimientos · total <b style="color:{E.TEXTO}">{total_txt}</b></p>'.replace(",", "."))
            st.dataframe(mov, hide_index=True, width="stretch", height=380, column_config={
                "empresa": "Compañía", "fecha": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"), "mes": "Mes",
                "documento": "Documento", "tercero": "Proveedor / tercero", "texto": st.column_config.TextColumn("Texto", width="large"),
                "codigo_articulo": "Referencia (SISPO)", "cantidad": st.column_config.NumberColumn("Cantidad", format="localized"),
                "cuenta_nombre": "Cuenta", "sub_rubro_sap": "Sub-rubro (SAP)", "cc_nombre": "Centro de coste",
                "valor": st.column_config.NumberColumn("Cantidad" if S.cant else "Valor", format="localized" if S.cant else "accounting")})
    st.download_button(f"Descargar {hoja} (Excel)", lambda: _excel({
        "Mensual": t[["Mes", "presupuesto", "ejecutado", "variacion", "pct", "ppto_acum", "ejec_acum", "pct_acum"]].rename(columns={
            "presupuesto": "Presupuesto", "ejecutado": "Ejecutado", "variacion": "Variación", "pct": "% ejecución",
            "ppto_acum": "Presupuesto acumulado", "ejec_acum": "Ejecutado acumulado", "pct_acum": "% acumulado"}),
        "Ítems": it.drop(columns=["costo_unit"]).rename(columns={"item": "Ítem", "cantidad": "Cantidad presupuestada",
                                    "presupuesto_esc": "Presupuesto escenario", "presupuesto": "Presupuesto al corte",
                                    "ejecutado": "Ejecutado", "variacion": "Variación", "pct": "% ejecución", "peso": "% del total"}),
        "Movimientos": mov_todos}), f"operaciones_{hoja}_{S.escenario.replace(' ', '_')}_{S.hasta_mes:02d}.xlsx".replace("í", "i"),
        icon=":material/download:", key=f"op_dl_{hoja}")


def _tabla_resumen(S):
    """Tabla del resumen (una fila por hoja en pesos + «Sin hoja asignada» + Total). La usan la pantalla y el informe descargable."""
    hojas_pesos = [h for h in S.hojas if h != OS.HOJA_PAPELERIA]                   # Papelería va en cantidades: no se suma con pesos
    filas = [_cifras(S, h) for h in hojas_pesos]
    sin = S.ejec[(S.ejec["hoja"] == SIN_HOJA) & (S.ejec["mes"] <= S.hasta_mes)]
    if len(sin):
        e = float(sin["valor"].sum())
        t = L.con_metricas(pd.DataFrame({"presupuesto": [0.0], "ejecutado": [e], "movimientos": [len(sin)]})).iloc[0]
        filas.append(dict(hoja=SIN_HOJA, presupuesto_esc=0.0, presupuesto=0.0, ejecutado=e, variacion=e, saldo=-e, pct=t["pct"],
                          estado=t["estado"], movimientos=len(sin)))
    t = pd.DataFrame(filas)
    if t.empty:
        return None
    tot = dict(presupuesto_esc=t["presupuesto_esc"].sum(), presupuesto=t["presupuesto"].sum(), ejecutado=t["ejecutado"].sum(),
               movimientos=int(t["movimientos"].sum()))
    m = L.con_metricas(pd.DataFrame({"presupuesto": [tot["presupuesto"]], "ejecutado": [tot["ejecutado"]], "movimientos": [tot["movimientos"]]})).iloc[0]
    t = pd.concat([t, pd.DataFrame([dict(hoja="Total", presupuesto_esc=tot["presupuesto_esc"], presupuesto=tot["presupuesto"],
                                         ejecutado=tot["ejecutado"], variacion=tot["ejecutado"] - tot["presupuesto"],
                                         saldo=tot["presupuesto_esc"] - tot["ejecutado"], pct=m["pct"], estado=m["estado"],
                                         movimientos=tot["movimientos"])])], ignore_index=True)
    return t, tot, m, sin


def _control_meses(S):
    """Solo administrador. Por compañía y mes (hasta el corte): lo que hay en SAP (rubro 3.1), lo que muestra Operaciones (solo SISPO) y
    la comparación de la cuenta OTROS entre SAP y SISPO. La comparación es informativa: no condiciona lo que se muestra."""
    filas = []
    val = lambda df: float(pd.to_numeric(df["valor"], errors="coerce").fillna(0.0).sum()) if len(df) else 0.0
    for e in S.sel:
        d = S.datos_e[e]
        a = d.ejec_all[_es_rubro(d.ejec_all["codigo"])] if len(d.ejec_all) else d.ejec_all
        for m in range(S.m0, S.hasta_mes + 1):
            pn = S.cfg["anio"] * 100 + m
            sap = a[pd.to_numeric(a["mes"], errors="coerce") == m] if len(a) else a
            gen = _generico_otros_sap(d.ejec_all, e, pn)
            op = d.ejec[d.ejec["mes"] == m]
            per = S.per_sispo[(S.per_sispo["empresa"] == e) & (S.per_sispo["periodo"] == pn)]
            sis_otros = float(per["total_otros"].iloc[0]) if len(per) else 0.0
            dif = round(val(gen) - sis_otros, 2)
            if per.empty and gen.empty and not len(sap):
                dx = "Sin SAP ni SISPO de este mes."
            elif per.empty:
                dx = "Falta el SISPO: aplícalo en Mantenimiento («Cargar SISPO del mes»); se guarda también en Operaciones."
            elif gen.empty:
                dx = "SISPO tomado. SAP no trae la «Salida de inventario» de OTROS: no hay con qué compararlo."
            elif L.centavos(val(gen)) != L.centavos(sis_otros):
                dx = "SISPO tomado, pero la cuenta OTROS difiere de SAP: revisa el SISPO o el SAP de este mes."
            else:
                dx = "Completo: SISPO tomado y cuadra con SAP."
            filas.append({"Compañía": e, "Mes": MESES_L[m - 1].capitalize(), "SAP rubro 3.1 ($)": val(sap), "OTROS en SAP ($)": val(gen),
                          "OTROS en SISPO ($)": sis_otros, "Diferencia ($)": dif, "En Operaciones (SISPO) ($)": val(op),
                          "Líneas SISPO": 0 if per.empty else int(((S.lin_sispo["empresa"] == e) & (S.lin_sispo["periodo"] == pn)).sum()),
                          "Diagnóstico": dx})
    if not filas:
        return
    with st.expander("Control: SAP vs SISPO por mes (la ejecución de Operaciones sale solo del SISPO)", expanded=False):
        nc = st.column_config.NumberColumn
        st.dataframe(pd.DataFrame(filas), hide_index=True, width="stretch", column_config={
            c: nc(c, format="accounting") for c in ("SAP rubro 3.1 ($)", "OTROS en SAP ($)", "OTROS en SISPO ($)", "Diferencia ($)", "En Operaciones (SISPO) ($)")})
        st.caption("«En Operaciones» es solo SISPO (Consumibles, Herramientas y Otros). El SAP del rubro 3.1 se muestra únicamente para comparar. "
                   "El SISPO se aplica desde Mantenimiento (la carga es única).")


def _tab_resumen(S):
    res = _tabla_resumen(S)
    if res is None:
        st.info("No hay presupuesto ni ejecución del rubro 3.1 para mostrar.")
        return
    t, tot, m, sin = res
    _kpis(dict(presupuesto_esc=tot["presupuesto_esc"], presupuesto=tot["presupuesto"], ejecutado=tot["ejecutado"], movimientos=tot["movimientos"],
               pct=m["pct"], estado=m["estado"], saldo=tot["presupuesto_esc"] - tot["ejecutado"]), S)
    with st.container(key="op_card_resumen"):
        titulo_tarjeta("Operaciones · presupuesto vs ejecución por hoja", f"{S.escenario} · ejecutado {S.hasta} · variación = ejecutado − presupuesto al corte")
        nc = st.column_config.NumberColumn
        vista = t[["hoja", "presupuesto_esc", "presupuesto", "ejecutado", "variacion", "pct", "saldo", "estado", "movimientos"]]
        st.dataframe(_estilo(vista, "variacion"), hide_index=True, width="stretch", height=40 + 35 * len(vista), column_config={
            "hoja": "Hoja", "presupuesto_esc": nc("Presupuesto escenario", format="accounting"), "presupuesto": nc("Presupuesto al corte", format="accounting"),
            "ejecutado": nc("Ejecutado", format="accounting"), "variacion": nc("Variación", format="accounting"),
            "pct": nc("% ejecución", format="percent"), "saldo": nc("Saldo del escenario", format="accounting"), "estado": "Estado",
            "movimientos": nc("Movimientos", format="%d")})
        st.caption("La ejecución sale solo del SISPO (Consumibles, Herramientas y Otros); Papelería, en cantidades. SAP no aporta valores a Operaciones.")
        if "pend_sispo" in sin:
            pend, resto = sin[sin["pend_sispo"].fillna(False).astype(bool)], sin[~sin["pend_sispo"].fillna(False).astype(bool)]
        else:
            pend, resto = sin.iloc[0:0], sin
        if len(pend):
            st.warning(f"{len(pend)} movimientos por {E.pesos(pend['valor'].sum())} son la «Salida de inventario» de la cuenta OTROS de meses sin SISPO "
                       "aplicado (o que no cuadra con SAP). Al aplicarlo en Mantenimiento («Cargar SISPO del mes»), se reparte entre Consumibles, "
                       "Herramientas y Otros." + ("" if S.ctx.es_admin else " El administrador carga el SISPO."), icon=":material/inventory_2:")
        if len(resto):
            st.warning(f"{len(resto)} movimientos por {E.pesos(resto['valor'].sum())} son del rubro 3.1 pero su sub-rubro no está asignado a ninguna "
                       "hoja" + (": asígnalo en «Asignación de sub-rubros a hojas»." if S.ctx.es_admin else "; el administrador lo asigna."),
                       icon=":material/rule:")
    if S.ctx.es_admin:
        _control_meses(S)
    graf = t[t["hoja"] != "Total"]
    graf = graf[(graf["presupuesto"] > 0) | (graf["ejecutado"] != 0)]
    if not graf.empty:
        with st.container(key="op_card_resumen_graf"):
            titulo_tarjeta("Por hoja", "Presupuesto al corte vs ejecutado (millones de pesos)")
            fig = go.Figure()
            fig.add_bar(x=graf["hoja"], y=graf["presupuesto"] / 1e6, name="Presupuesto al corte", marker_color="#c9d4e4", marker_line_width=0,
                        hovertemplate="%{y:,.1f} M")
            fig.add_bar(x=graf["hoja"], y=graf["ejecutado"] / 1e6, name="Ejecutado", marker_color=E.AZUL, marker_line_width=0,
                        hovertemplate="%{y:,.1f} M")
            fig.update_layout(barmode="group", bargap=.3, yaxis=dict(tickformat=",.0f", ticksuffix=" M", rangemode="tozero"))
            st.plotly_chart(estilo_grafico(fig, 320), width="stretch", config={"displayModeBar": False})
    if OS.HOJA_PAPELERIA in S.hojas:                    # Papelería: en cantidades, aparte de los pesos
        q = _vista_cantidades(S)
        if len(_ppto_hoja(q, OS.HOJA_PAPELERIA)) or len(q.ejec):
            c = _cifras(q, OS.HOJA_PAPELERIA)
            with st.container(key="op_card_resumen_pap"):
                titulo_tarjeta("Papelería · cantidades", f"{S.escenario} · ejecutado {S.hasta} · unidades (costo cero: no hay pesos que comparar)")
                nc = st.column_config.NumberColumn
                vp = pd.DataFrame([dict(presupuesto_esc=c["presupuesto_esc"], presupuesto=c["presupuesto"], ejecutado=c["ejecutado"],
                                        variacion=c["variacion"], pct=c["pct"], saldo=c["saldo"], estado=c["estado"], movimientos=c["movimientos"])])
                st.dataframe(_estilo(vp, "variacion"), hide_index=True, width="stretch", column_config={
                    "presupuesto_esc": nc("Cantidad presupuestada (escenario)", format="localized"),
                    "presupuesto": nc("Cantidad presupuestada al corte", format="localized"), "ejecutado": nc("Cantidad ejecutada", format="localized"),
                    "variacion": nc("Diferencia", format="localized"), "pct": nc("% ejecución", format="percent"),
                    "saldo": nc("Saldo del escenario", format="localized"), "estado": "Estado", "movimientos": nc("Salidas", format="%d")})


# ================================================================== carga y asignación (administrador)
def _insertar_por_lotes(ctx, sql_cols: str, filas: list, lote=100):
    n = len(filas[0])
    for i in range(0, len(filas), lote):
        parte = filas[i:i + lote]
        ctx.con.execute(f"INSERT INTO {sql_cols} VALUES " + ",".join(["(" + ",".join("?" * n) + ")"] * len(parte)),
                        tuple(v for f in parte for v in f))


def _a_db(v):
    return None if (not isinstance(v, str) and pd.isna(v)) else (v.item() if hasattr(v, "item") else v)


def _guardar_presupuesto(ctx, empresa: str, largo: pd.DataFrame, archivo: str) -> int:
    """Reemplaza, para la compañía, SOLO las hojas que trae el libro (las demás quedan como estaban). Idempotente."""
    _exigir_admin(ctx)
    n = 0
    for hoja, g in largo.groupby("hoja"):
        ctx.con.execute("DELETE FROM oper_ppto WHERE empresa=? AND hoja=?", (empresa, hoja))
        _insertar_por_lotes(ctx, "oper_ppto(empresa, hoja, item, orden, anio, mes, cantidad, costo_unit, valor)",
                            [(empresa, hoja, r.item, int(r.orden), int(r.anio), int(r.mes), _a_db(r.cantidad), _a_db(r.costo_unit), float(r.valor))
                             for r in g.itertuples(index=False)], lote=80)
        n += len(g)
    ctx.registrar_carga(ctx.usuario, "Presupuesto operaciones", archivo,
                        f"{empresa}: {', '.join(sorted(largo['hoja'].unique()))} · {n} filas · {E.pesos(largo['valor'].sum())}")
    ctx.con.commit()
    return n


def _guardar_mapeo(ctx, filas: list[tuple]) -> int:
    """filas: (sub_norm, sub_rubro, hoja). Guarda la decisión explícita para cada sub-rubro recibido."""
    _exigir_admin(ctx)
    ts = ahora().strftime("%Y-%m-%d %H:%M")
    for sub_norm, sub, hoja in filas:
        ctx.con.execute("DELETE FROM oper_mapeo WHERE sub_norm=?", (sub_norm,))
        ctx.con.execute("INSERT INTO oper_mapeo(sub_norm, sub_rubro, hoja, usuario, fecha) VALUES(?,?,?,?,?)",
                        (sub_norm, sub, hoja, ctx.usuario, ts))
    ctx.registrar_carga(ctx.usuario, "Asignación sub-rubros operaciones", "pantalla", f"{len(filas)} sub-rubros")
    ctx.con.commit()
    return len(filas)


def _resumen_libro(info: dict) -> pd.DataFrame:
    filas = []
    for hoja in HOJAS:
        i = info["hojas"].get(hoja)
        if i is None:
            filas.append(dict(Hoja=hoja, Estado="No viene en el libro"))
            continue
        if "error" in i:
            filas.append(dict(Hoja=hoja, Estado="Error: " + i["error"]))
            continue
        obs = []
        if i["dif_qxu"]:
            obs.append(f"{len(i['dif_qxu'])} celdas donde cantidad × costo ≠ total")
        if i["dif_mes"]:
            obs.append(f"{len(i['dif_mes'])} meses donde la suma de ítems ≠ total de la hoja")
        if i["dif_fila_total"] is not None and abs(i["dif_fila_total"]) > 1:
            obs.append(f"fila «Total» del libro difiere {_pesos(i['dif_fila_total'])}")
        if i["cant_total_no_cuadra"]:
            obs.append(f"en la fila «Total» la cantidad de {len(i['cant_total_no_cuadra'])} ítems no suma todos los meses")
        if i["filas_fuera_escenario"]:
            obs.append(f"{i['filas_fuera_escenario']} filas fuera de los escenarios (no se usan)")
        filas.append(dict(Hoja=hoja, Estado="Listo", Ítems=i["n_items"], Meses=i["n_meses"],
                          **{f"Total {esc}": v for esc, v in i["por_escenario"].items()}, Observaciones=" · ".join(obs) or "—"))
    return pd.DataFrame(filas)


def _admin_presupuesto(ctx, empresa: str):
    _exigir_admin(ctx)
    cargado = ctx.leer("SELECT hoja, COUNT(*) filas, SUM(valor) total FROM oper_ppto WHERE empresa=? GROUP BY hoja", (empresa,))
    if cargado.empty:
        st.caption(f"{empresa}: todavía no hay presupuesto de Operaciones cargado.")
    else:
        st.caption(f"{empresa} tiene cargado: " + " · ".join(f"{r.hoja} {E.pesos(r.total)}" for r in cargado.itertuples()))
    g = st.file_uploader(f"Libro «3.1 operaciones» (.xlsx) · {empresa}", type=["xlsx"], key=f"op_up_{empresa}",
                         help="Una hoja por Papelería, Consumibles, Herramientas y Bolsas: una fila por mes y, por ítem, cantidad, costo "
                              "unitario y total. Cargar el libro reemplaza las hojas que trae para esta compañía; no toca la otra.")
    if not g:
        return
    try:
        largo, info = leer_libro(g)
    except Exception as e:
        st.error(f"No pude leer el libro: {e}")
        return
    st.dataframe(_resumen_libro(info), hide_index=True, width="stretch", height=40 + 35 * len(HOJAS),
                 column_config={c: st.column_config.NumberColumn(c, format="accounting") for c in (f"Total {x}" for x in L.ESCENARIOS)})
    if info["ignoradas"]:
        st.caption("Hojas del libro que no se usan: " + ", ".join(info["ignoradas"]))
    errores = {h: i["error"] for h, i in info["hojas"].items() if "error" in i}
    if errores:
        st.error("Hay hojas que no se pueden cargar; corrígelas en el libro: " + " · ".join(f"{h}: {e}" for h, e in errores.items()))
    if info["faltan"]:
        st.warning("El libro no trae: " + ", ".join(info["faltan"]) + ". Esas hojas quedan como estaban en la base.")
    if largo.empty or errores:
        return
    if st.button(f"Guardar presupuesto de Operaciones · {empresa}", key=f"op_guardar_{empresa}", icon=":material/save:"):
        n = _guardar_presupuesto(ctx, empresa, largo, g.name)
        st.success(f"Presupuesto de {empresa} guardado: {n} filas de {', '.join(sorted(largo['hoja'].unique()))}.")
        st.rerun()


def _admin_mapeo(S):
    _exigir_admin(S.ctx)
    partes = []
    for d in S.datos_e.values():
        if len(d.ppto_base):
            partes.append(d.ppto_base.groupby(["_sub", "sub_rubro"], as_index=False).agg(presupuesto=("valor", "sum")))
        if len(d.ejec_sap):
            partes.append(d.ejec_sap.groupby(["_sub", "sub_rubro"], as_index=False).agg(ejecutado=("valor", "sum")))
    if not partes:
        st.info("No hay sub-rubros del rubro 3.1 en la base para este escenario.")
        return
    t = pd.concat(partes, ignore_index=True).groupby("_sub", as_index=False).agg(
        sub_rubro=("sub_rubro", "first"), presupuesto=("presupuesto", "sum"), ejecutado=("ejecutado", "sum"))
    t["hoja"] = _hojas_de(t, S.mapeo)
    t = t.sort_values("sub_rubro")
    ed = st.data_editor(t[["sub_rubro", "presupuesto", "ejecutado", "hoja"]], hide_index=True, width="stretch", key="op_mapeo_ed",
                        disabled=["sub_rubro", "presupuesto", "ejecutado"], column_config={
                            "sub_rubro": "Sub-rubro del rubro 3.1",
                            "presupuesto": st.column_config.NumberColumn("Presupuesto (módulo Presupuesto)", format="accounting"),
                            "ejecutado": st.column_config.NumberColumn("Ejecutado (escenario)", format="accounting"),
                            "hoja": st.column_config.SelectboxColumn("Hoja", options=[SIN_HOJA, *HOJAS], required=True)})
    st.caption("Los movimientos de SAP de cada sub-rubro se comparan con la hoja elegida. Se propone una hoja cuando el nombre del sub-rubro "
               "la contiene (Papelería, Consumibles, Herramientas, Bolsas); lo demás queda en «Sin hoja asignada».")
    if st.button("Guardar asignación", key="op_mapeo_guardar", icon=":material/save:"):
        n = _guardar_mapeo(S.ctx, [(s, nom, h) for s, nom, h in zip(t["_sub"], ed["sub_rubro"], ed["hoja"])])
        st.success(f"Asignación guardada ({n} sub-rubros).")
        st.rerun()


def _guardar_sispo(ctx, empresa: str, periodo: int, d: pd.DataFrame, archivo: str) -> int:
    """Guarda las líneas de Operaciones del SISPO de una compañía y un periodo. Idempotente: reemplaza lo cargado antes de ese periodo."""
    _exigir_admin(ctx)
    g = OS.para_guardar(d, empresa, periodo)
    ctx.con.execute("DELETE FROM oper_sispo WHERE empresa=? AND periodo=?", (empresa, int(periodo)))
    ctx.con.execute("DELETE FROM oper_sispo_periodo WHERE empresa=? AND periodo=?", (empresa, int(periodo)))
    if len(g):
        _insertar_por_lotes(ctx, "oper_sispo(empresa, periodo, hoja, id_kardex, fecha, bodega, codigo, articulo, grupo, cuenta, cc_codigo, "
                                 "cc_nombre, cantidad, valor, concepto)",
                            [tuple(_a_db(v) for v in r) for r in g.itertuples(index=False, name=None)], lote=40)
    ctx.con.execute("INSERT INTO oper_sispo_periodo(empresa, periodo, archivo, usuario, fecha, n_lineas, total_otros) VALUES(?,?,?,?,?,?,?)",
                    (empresa, int(periodo), archivo, ctx.usuario, ahora().strftime("%Y-%m-%d %H:%M"), len(g), OS.total_otros_c(d) / 100))
    ctx.registrar_carga(ctx.usuario, "SISPO operaciones", archivo, f"{empresa} {periodo}: {len(g)} líneas")
    ctx.con.commit()
    return len(g)


def sispo_analizar(ctx, up) -> dict:
    """Lo que Operaciones toma de un SISPO (la carga es única y se hace en Mantenimiento: ese módulo llama a esta función y a
    `sispo_guardar`). No guarda nada. Devuelve `d` (archivo clasificado), `bloqueos` (solo de Operaciones), `embudo`, `resumen`,
    `huerfanos` (valores de «Concepto» que no son ítems del presupuesto) y `total_otros_c` (centavos de la cuenta OTROS)."""
    _asegurar_tablas(ctx)
    items = OS.mapa_items(ctx.leer("SELECT codigo, item FROM oper_item_map").itertuples(index=False, name=None))
    presup = OS.items_presupuestados(ctx.leer("SELECT item FROM oper_ppto"))
    d, solapes = OS.clasificar(OS.leer(BytesIO(up.getvalue())), items, presup)
    bloqueos = []
    if d["id_kardex"].eq("").any() or d["id_kardex"].duplicated().any():
        bloqueos.append("Hay IdKardex vacíos o repetidos: revisa el archivo para no duplicar movimientos.")
    if d[d["hoja"] != ""]["codigo"].eq("").any():
        bloqueos.append("Hay líneas de Operaciones sin «Codigo Articulo»: no se pueden ligar a un ítem.")
    if solapes:
        bloqueos.append(f"{solapes} líneas cumplen más de una regla de hoja: revisa la clasificación antes de cargar.")
    huerfanos = []
    if "concepto" in d and presup:
        conc = d["concepto"].fillna("").astype(str)
        huerfanos = sorted({c for c in conc[(conc != "") & ~conc.map(L._canon).isin(presup)]})
    return dict(d=d, bloqueos=bloqueos, embudo=OS.depurar(d)[1], resumen=OS.resumen(d), huerfanos=huerfanos, total_otros_c=OS.total_otros_c(d))


def sispo_estado(ctx, empresa: str, periodo: int, a: dict) -> str:
    """«nuevo» (el periodo no está guardado), «igual» (ya guardado con el mismo número de líneas y el mismo total OTROS) o «distinto»."""
    _asegurar_tablas(ctx)
    p = ctx.leer("SELECT n_lineas, total_otros FROM oper_sispo_periodo WHERE empresa=? AND periodo=?", (empresa, int(periodo)))
    if p.empty:
        return "nuevo"
    n = len(OS.para_guardar(a["d"], empresa, periodo))
    igual = int(p.iloc[0]["n_lineas"]) == n and abs(float(p.iloc[0]["total_otros"]) * 100 - a["total_otros_c"]) < 0.5
    return "igual" if igual else "distinto"


def sispo_guardar(ctx, empresa: str, periodo: int, a: dict, archivo: str) -> int:
    """Guarda en Operaciones las líneas del análisis hecho con `sispo_analizar` (idempotente por compañía y periodo)."""
    _asegurar_tablas(ctx)
    return _guardar_sispo(ctx, empresa, periodo, a["d"], archivo)


def _revertir_sispo(ctx, empresa: str, periodo: int):
    _exigir_admin(ctx)
    for t in ("oper_sispo", "oper_sispo_periodo"):
        ctx.con.execute(f"DELETE FROM {t} WHERE empresa=? AND periodo=?", (empresa, int(periodo)))
    ctx.registrar_carga(ctx.usuario, "Reversión SISPO operaciones", "pantalla", f"{empresa} {periodo}")
    ctx.con.commit()


def _admin_sispo(S):
    _exigir_admin(S.ctx)
    ctx = S.ctx
    cargado = ctx.leer("SELECT empresa, periodo, archivo, usuario, fecha, n_lineas, total_otros FROM oper_sispo_periodo ORDER BY periodo DESC, empresa")
    if cargado.empty:
        st.caption("Todavía no hay SISPO cargado en Operaciones.")
    else:
        st.dataframe(cargado, hide_index=True, width="stretch", column_config={
            "empresa": "Compañía", "periodo": st.column_config.NumberColumn("Periodo", format="%d"), "archivo": "Archivo", "usuario": "Usuario",
            "fecha": "Cargado", "n_lineas": st.column_config.NumberColumn("Líneas", format="%d"),
            "total_otros": st.column_config.NumberColumn("Total cuenta OTROS (SISPO)", format="accounting")})
    st.info("El SISPO se carga UNA sola vez, en Mantenimiento («Cargar SISPO del mes y reemplazar las salidas de inventario»). Al aplicarlo, "
            "ese mismo archivo también se guarda aquí: Operaciones toma sus 5 centros de costo, grupos 5 y 7 y cuentas ÚTILES, PAPELERÍA Y OTROS u "
            "OTROS, sin Staf Mantenimiento. Volver a aplicar el mismo periodo lo reemplaza, no lo duplica.", icon=":material/sync_alt:")
    if not cargado.empty:
        opciones = [f"{r.empresa} · {r.periodo}" for r in cargado.itertuples()]
        c_sel, c_ok, c_btn = st.columns([2, 2, 1], vertical_alignment="bottom")
        elegido = c_sel.selectbox("Periodo cargado a quitar", opciones, key="op_sispo_quitar")
        ok = c_ok.checkbox("Confirmo quitar este SISPO de Operaciones", key="op_sispo_quitar_ok")
        if c_btn.button("Quitar", key="op_sispo_quitar_btn", icon=":material/undo:", disabled=not ok):
            emp, per_ = elegido.split(" · ")
            _revertir_sispo(ctx, emp, int(per_))
            st.success("SISPO quitado: vuelven los movimientos de SAP tal como estaban.")
            st.rerun()


def _guardar_items(ctx, filas: list[tuple]):
    _exigir_admin(ctx)
    ts = ahora().strftime("%Y-%m-%d %H:%M")
    ctx.con.execute("DELETE FROM oper_item_map")
    _insertar_por_lotes(ctx, "oper_item_map(codigo, item, usuario, fecha)", [(c, i, ctx.usuario, ts) for c, i in filas], lote=80)
    ctx.registrar_carga(ctx.usuario, "Equivalencias referencia-ítem operaciones", "pantalla", f"{len(filas)} equivalencias")
    ctx.con.commit()


def _admin_items(S):
    _exigir_admin(S.ctx)
    nombres = {}
    if len(S.lin_sispo):
        nombres = dict(zip(S.lin_sispo["codigo"], S.lin_sispo["articulo"]))
    t = pd.DataFrame({"codigo": sorted(S.items), "item": [S.items[c] for c in sorted(S.items)]})
    t["articulo"] = t["codigo"].map(nombres).fillna("")
    ed = st.data_editor(t[["codigo", "articulo", "item"]], hide_index=True, width="stretch", num_rows="dynamic", key="op_items_ed",
                        disabled=["articulo"], column_config={"codigo": "Código del artículo (SISPO)", "articulo": "Artículo (último SISPO cargado)",
                                                              "item": "Ítem del presupuesto"})
    st.caption("El ítem se escribe como está en el libro (la unión ignora mayúsculas y tildes). Una referencia sin equivalencia aparece en las "
               "tablas por ítem como «<artículo> · sin ítem», con presupuesto 0.")
    conocidos = {L._canon(i) for i in S.ppto["item"].dropna().unique()}
    sin_ppto = sorted({i for i in S.items.values() if L._canon(i) not in conocidos})
    if sin_ppto and conocidos:
        st.warning("Ítems de las equivalencias que no existen en el presupuesto cargado (¿error de escritura?): " + " · ".join(sin_ppto),
                   icon=":material/spellcheck:")
    if st.button("Guardar equivalencias", key="op_items_guardar", icon=":material/save:"):
        filas = [(str(c).strip(), " ".join(str(i).split()) if pd.notna(i) else "") for c, i in zip(ed["codigo"], ed["item"])
                 if pd.notna(c) and str(c).strip()]
        borradas = [(c, "") for c in OS.MAPEO_INICIAL if c not in {f[0] for f in filas}]      # quitar una de partida se guarda como «sin ítem»
        _guardar_items(S.ctx, filas + borradas)
        st.success(f"Equivalencias guardadas ({len(filas)}).")
        st.rerun()


def _admin(S):
    _exigir_admin(S.ctx)
    with st.expander("Datos base · presupuesto, SISPO, equivalencias y asignación de sub-rubros (administrador)", icon=":material/settings:"):
        st.markdown("**1. Presupuesto del libro «3.1 operaciones».** Se carga por compañía.")
        for tab, e in zip(st.tabs([f"{e} · {EMPRESAS[e]}" for e in EMPRESAS]), EMPRESAS):
            with tab:
                _admin_presupuesto(S.ctx, e)
        st.divider()
        st.markdown("**2. SISPO del mes (carga única en Mantenimiento).** Reparte Consumibles, Herramientas y Otros (en pesos) y trae Papelería (en cantidades).")
        _admin_sispo(S)
        st.divider()
        st.markdown("**3. Equivalencias referencia → ítem del presupuesto.**")
        _admin_items(S)
        st.divider()
        st.markdown("**4. Asignación de sub-rubros a hojas.**")
        _admin_mapeo(S)


# ================================================================== pantalla
# ================================================================== informe Excel (solo da formato; las cifras vienen de las funciones de arriba)
_X_NAVY, _X_AZUL, _X_CLARO, _X_GRIS, _X_BLANCO = "1F3864", "2F5597", "DCE6F1", "F2F2F2", "FFFFFF"
_X_VERDE, _X_AMBAR, _X_ROJO = "E2EFDA", "FFF2CC", "F8CBAD"
_X_FUENTE = "Calibri"
_X_F_PESOS, _X_F_CANT, _X_F_PCT = '#,##0;[Red]-#,##0;"-"', '#,##0;[Red]-#,##0;"-"', '0.0%'
_X_fino = Side(style="thin", color="BFBFBF")
_X_BORDE = Border(left=_X_fino, right=_X_fino, top=_X_fino, bottom=_X_fino)


def _X_f(bold=False, size=10, color="000000", italic=False):
    return Font(name=_X_FUENTE, bold=bold, size=size, color=color, italic=italic)


def _X_fill(c):
    return PatternFill("solid", fgColor=c)


def _X_estado_color(est: str) -> str:
    return {"Dentro del presupuesto": _X_VERDE, "Presupuesto agotado": _X_AMBAR, "Sobre ejecutado": _X_ROJO,
            "Sin presupuesto asignado": _X_ROJO}.get(est, _X_CLARO)


def _X_num(v):
    """None si no hay número (celda vacía), para no escribir NaN."""
    try:
        return None if pd.isna(v) else float(v)
    except (TypeError, ValueError):
        return None


def _X_encabezado(ws, titulo, subtitulo, ancho):
    ws.sheet_view.showGridLines = False
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ancho)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ancho)
    c = ws.cell(1, 1, titulo)
    c.font, c.fill, c.alignment = _X_f(True, 15, _X_BLANCO), _X_fill(_X_NAVY), Alignment(vertical="center", indent=1)
    s = ws.cell(2, 1, subtitulo)
    s.font, s.fill, s.alignment = _X_f(False, 10, _X_BLANCO), _X_fill(_X_AZUL), Alignment(vertical="center", indent=1)
    for col in range(2, ancho + 1):
        ws.cell(1, col).fill, ws.cell(2, col).fill = _X_fill(_X_NAVY), _X_fill(_X_AZUL)
    ws.row_dimensions[1].height, ws.row_dimensions[2].height = 30, 20
    ws.page_setup.orientation, ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = "landscape", 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def _X_seccion(ws, fila, texto, ancho):
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=ancho)
    c = ws.cell(fila, 1, texto)
    c.font, c.fill, c.alignment = _X_f(True, 11, _X_NAVY), _X_fill(_X_CLARO), Alignment(vertical="center", indent=1)
    for col in range(2, ancho + 1):
        ws.cell(fila, col).fill = _X_fill(_X_CLARO)
    ws.row_dimensions[fila].height = 20
    return fila + 1


def _X_cabecera_tabla(ws, fila, titulos):
    for j, t in enumerate(titulos, 1):
        c = ws.cell(fila, j, t)
        c.font, c.fill, c.border = _X_f(True, 10, _X_BLANCO), _X_fill(_X_AZUL), _X_BORDE
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[fila].height = 32
    return fila + 1


def _X_escribir(ws, fila, valores, formatos, estilo=None, rellenos=None):
    """estilo: None (dato), 'sub' (subtotal) o 'tot' (total)."""
    for j, (v, fm) in enumerate(zip(valores, formatos), 1):
        c = ws.cell(fila, j, v)
        c.border = _X_BORDE
        c.font = _X_f(estilo is not None, 10, _X_BLANCO if estilo == "tot" else "000000")
        if estilo == "tot":
            c.fill = _X_fill(_X_NAVY)
        elif estilo == "sub":
            c.fill = _X_fill(_X_GRIS)
        if rellenos and j in rellenos:
            c.fill = _X_fill(rellenos[j])
        if fm:
            c.number_format = fm
            c.alignment = Alignment(horizontal="right")
        else:
            c.alignment = Alignment(vertical="top", wrap_text=False)
    return fila + 1


def _X_kpis(ws, fila, items, fmt_val):
    """items = [(rótulo, valor, formato, color_fondo|None)] en columnas consecutivas desde la B."""
    for j, (rot, val, fm, color) in enumerate(items, 2):
        a, b = ws.cell(fila, j, rot), ws.cell(fila + 1, j, val)
        a.font, a.alignment, a.border = _X_f(False, 9, "595959"), Alignment(horizontal="center", vertical="center", wrap_text=True), _X_BORDE
        b.font, b.border = _X_f(True, 13, _X_NAVY), _X_BORDE
        b.alignment = Alignment(horizontal="center", vertical="center")
        a.fill = b.fill = _X_fill(color or _X_GRIS)
        if fm:
            b.number_format = fm
    ws.row_dimensions[fila].height, ws.row_dimensions[fila + 1].height = 28, 26
    return fila + 3


def _X_anchos(ws, anchos):
    for j, w in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(j)].width = w


# ---------------------------------------------------------------------------------------------------------------- resumen
def _X_matriz(ws, fila, meses, filas, fmt, titulo_col, n_corte):
    """Tabla por mes: para cada mes, Presupuesto · Ejecutado · % de cumplimiento, y al final el total al corte.
    `filas` = [(etiqueta, [presupuesto por mes], [ejecutado por mes o None después del corte])]. Última fila: Total."""
    n = len(meses)
    c = ws.cell(fila, 1, titulo_col)
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila + 1, end_column=1)
    c.font, c.fill, c.border = _X_f(True, 10, _X_BLANCO), _X_fill(_X_AZUL), _X_BORDE
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for j, m in enumerate(meses + ["Total al corte"]):
        col = 2 + 3 * j
        ws.merge_cells(start_row=fila, start_column=col, end_row=fila, end_column=col + 2)
        h = ws.cell(fila, col, m)
        h.font, h.alignment = _X_f(True, 10, _X_BLANCO), Alignment(horizontal="center", vertical="center")
        for k, t in enumerate(("Presupuesto", "Ejecutado", "% cumplimiento")):
            x = ws.cell(fila + 1, col + k, t)
            x.font, x.alignment = _X_f(True, 9, _X_BLANCO), Alignment(horizontal="center", vertical="center", wrap_text=True)
        for k in range(3):
            for rr in (fila, fila + 1):
                cc = ws.cell(rr, col + k)
                cc.fill, cc.border = _X_fill(_X_NAVY if j == n else _X_AZUL), _X_BORDE
        for k in range(3):
            if ws.column_dimensions[get_column_letter(col + k)].width in (None, 13.0) or col + k > 12:
                ws.column_dimensions[get_column_letter(col + k)].width = max(ws.column_dimensions[get_column_letter(col + k)].width or 0, 16)
    ws.row_dimensions[fila + 1].height = 28
    fila += 2

    def pct(p, e):
        return None if e is None or not p or p <= 0 else e / p

    def escribir(etq, pl, el, estilo):
        nonlocal fila
        vals = [etq]
        for j in range(n):
            vals += [pl[j], el[j], pct(pl[j], el[j])]
        pt = sum(pl[:n_corte])
        et = sum(v for v in el[:n_corte] if v is not None)
        vals += [pt, et, pct(pt, et)]
        fm = [None] + [fmt, fmt, _X_F_PCT] * (n + 1)
        fila = _X_escribir(ws, fila, vals, fm, estilo)

    for etq, pl, el in filas:
        escribir(etq, pl, el, None)
    tp = [sum(f[1][j] for f in filas) for j in range(n)]
    te = [None if all(f[2][j] is None for f in filas) else sum(f[2][j] or 0.0 for f in filas) for j in range(n)]
    escribir("Total", tp, te, "tot")
    return fila + 1


def _X_hoja_resumen(wb, meta, r):
    ws = wb.active
    ws.title = "1 Resumen"
    _X_anchos(ws, [30, 18, 18, 18, 18, 14, 18, 24, 13])
    _X_encabezado(ws, "1. RESUMEN GENERAL · PRESUPUESTO VS. EJECUCIÓN", meta["subtitulo"], 9)
    fila = 4
    fila = _X_seccion(ws, fila, "Indicadores del rubro 3.1 · valores en pesos (Papelería va aparte, en cantidades)", 9)
    k = r["kpis"]
    fila = _X_kpis(ws, fila, [
        ("Presupuesto del escenario", k["presupuesto_esc"], _X_F_PESOS, None), ("Presupuesto al corte", k["presupuesto"], _X_F_PESOS, None),
        ("Valor ejecutado", k["ejecutado"], _X_F_PESOS, None), ("Saldo disponible del escenario", k["saldo"], _X_F_PESOS, None),
        ("% de ejecución", _X_num(k["pct"]), _X_F_PCT, None), ("Variación (ejecutado − presupuesto al corte)", k["variacion"], _X_F_PESOS, None),
        ("Estado", k["estado"], None, _X_estado_color(k["estado"]))], _X_F_PESOS)
    if r.get("mensual_hojas"):
        mh = r["mensual_hojas"]
        fila = _X_seccion(ws, fila, "Presupuesto mensual, ejecución mensual y % de cumplimiento por hoja · pesos", 9)
        fila = _X_matriz(ws, fila, mh["meses"], mh["filas"], _X_F_PESOS, "Hoja", mh["n_corte"])
    if r.get("papeleria"):
        c = r["papeleria"]
        fila = _X_seccion(ws, fila, "Papelería · cantidades (costo cero: no hay pesos que comparar)", 9)
        if r.get("mensual_pap"):
            mp = r["mensual_pap"]
            fila = _X_matriz(ws, fila, mp["meses"], mp["filas"], _X_F_CANT, "Papelería (unidades)", mp["n_corte"])
        fila = _X_cabecera_tabla(ws, fila, ["Concepto", "Cantidad presupuestada (escenario)", "Cantidad presupuestada al corte", "Cantidad ejecutada",
                                          "Diferencia", "% ejecución", "Saldo del escenario", "Estado", "Salidas"])
        fila = _X_escribir(ws, fila, ["Papelería (unidades)", c["presupuesto_esc"], c["presupuesto"], c["ejecutado"], c["variacion"], _X_num(c["pct"]),
                                    c["saldo"], c["estado"], int(c["movimientos"])],
                           [None, _X_F_CANT, _X_F_CANT, _X_F_CANT, _X_F_CANT, _X_F_PCT, _X_F_CANT, None, "#,##0"], None, {8: _X_estado_color(c["estado"])})
        fila += 1
    fallas = [f"{n}: {d}" for n, res, d in r.get("validaciones", []) if res == "NO CUMPLE"] + list(r.get("avisos", []))   # solo se muestra si hay algo
    if fallas:
        fila = _X_seccion(ws, fila, "ADVERTENCIA · validaciones que no cumplen y meses sin repartir", 9)
        for texto in fallas:
            ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=9)
            c = ws.cell(fila, 1, texto)
            c.font, c.fill, c.alignment = _X_f(True, 10), _X_fill(_X_ROJO), Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[fila].height = 30
            fila += 1
    ws.freeze_panes = "A4"


# ---------------------------------------------------------------------------------------------------------------- hojas
def _X_hoja_detalle(wb, meta, s):
    ws = wb.create_sheet(s["hoja_excel"])
    cant = s["unidad"] == "unidades"
    fv = _X_F_CANT if cant else _X_F_PESOS
    u = "unidades" if cant else "pesos"
    ancho = 12
    _X_anchos(ws, [46, 16, 18, 18, 18, 16, 22, 34, 24, 24, 18, 18])
    _X_encabezado(ws, s["titulo"], meta["subtitulo"] + f" · valores en {u}", ancho)
    fila = 4
    c = s["cifras"]
    fila = _X_seccion(ws, fila, f"Indicadores · {u}", ancho)
    fila = _X_kpis(ws, fila, [
        ("Presupuesto del escenario", c["presupuesto_esc"], fv, None), ("Presupuesto al corte", c["presupuesto"], fv, None),
        ("Ejecutado", c["ejecutado"], fv, None), ("Saldo disponible del escenario", c["saldo"], fv, None),
        ("% de ejecución", _X_num(c["pct"]), _X_F_PCT, None), ("Variación", c["variacion"], fv, None),
        ("Estado", c["estado"], None, _X_estado_color(c["estado"]))], fv)

    # --- por ítem (presupuesto del libro ↔ ejecución)
    fila = _X_seccion(ws, fila, "Presupuesto vs ejecución por ítem", ancho)
    it = s["items"]
    if cant:
        fila = _X_cabecera_tabla(ws, fila, ["Ítem", "Cantidad presupuestada (escenario)", "Cantidad presupuestada al corte", "Cantidad ejecutada",
                                          "Diferencia", "% ejecución", "% de la cantidad presupuestada"])
        fi = [None, "#,##0", "#,##0", fv, fv, _X_F_PCT, _X_F_PCT]
        for row in it.itertuples(index=False):
            fila = _X_escribir(ws, fila, [row.item, row.presupuesto_esc, row.presupuesto, row.ejecutado, row.variacion, _X_num(row.pct), _X_num(row.peso)], fi)
        pt = it["presupuesto"].sum()
        fila = _X_escribir(ws, fila, ["Total", it["presupuesto_esc"].sum(), pt, it["ejecutado"].sum(), it["variacion"].sum(),
                                    _X_num(it["ejecutado"].sum() / pt) if pt > 0 else None, None], fi, "tot")
    else:
        fila = _X_cabecera_tabla(ws, fila, ["Ítem", "Cantidad presupuestada", "Presupuesto escenario", "Presupuesto al corte",
                                          "% ejecución", "Ejecutado", "Variación", "% del presupuesto"])
        fi = [None, "#,##0", fv, fv, _X_F_PCT, fv, fv, _X_F_PCT]
        for row in it.itertuples(index=False):
            fila = _X_escribir(ws, fila, [row.item, _X_num(row.cantidad), row.presupuesto_esc, row.presupuesto,
                                        _X_num(row.pct), row.ejecutado, row.variacion, _X_num(row.peso)], fi)
        pt = it["presupuesto"].sum()
        fila = _X_escribir(ws, fila, ["Total", it["cantidad"].sum(), it["presupuesto_esc"].sum(), pt, _X_num(it["ejecutado"].sum() / pt) if pt > 0 else None,
                                    it["ejecutado"].sum(), it["variacion"].sum(), None], fi, "tot")
    fila += 1

    # --- por ítem y por mes
    im = s["items_mes"]
    fila = _X_seccion(ws, fila, f"Presupuesto mensual, ejecución mensual y % de cumplimiento por ítem · {u}", ancho)
    fila = _X_matriz(ws, fila, im["meses"], im["filas"], fv, "Ítem", im["n_corte"])

    # --- mensual
    fila = _X_seccion(ws, fila, "Evolución mensual", ancho)
    fila = _X_cabecera_tabla(ws, fila, ["Mes", "Presupuesto", "Ejecutado", "Variación", "% ejecución", "Presupuesto acumulado", "Ejecutado acumulado",
                                      "% acumulado"])
    fm = [None, fv, fv, fv, _X_F_PCT, fv, fv, _X_F_PCT]
    for row in s["mensual"].itertuples(index=False):
        es_tot = row.Mes == "Total al corte"
        fila = _X_escribir(ws, fila, [row.Mes, row.presupuesto, _X_num(row.ejecutado), _X_num(row.variacion), _X_num(row.pct), _X_num(row.ppto_acum),
                                    _X_num(row.ejec_acum), _X_num(row.pct_acum)], fm, "tot" if es_tot else None)
    fila += 1

    # --- movimientos con subtotal por mes y total
    mv = s["movs"]
    fila = _X_seccion(ws, fila, f"Movimientos que componen la ejecución ({len(mv):,})".replace(",", "."), ancho)
    cab = ["Mes", "Fecha", "Documento", "Proveedor / tercero", "Texto", "Referencia (SISPO)", "Ítem del presupuesto", "Cuenta", "Sub-rubro (SAP)",
           "Centro de coste", "Cantidad"] + ([] if cant else ["Valor"])
    fmm = [None, "dd/mm/yyyy", None, None, None, None, None, None, None, None, "#,##0"] + ([] if cant else [fv])
    sufijo = lambda q, v: [_X_num(q)] + ([] if cant else [v])
    fila = _X_cabecera_tabla(ws, fila, cab)
    if mv.empty:
        ws.cell(fila, 1, "No hay movimientos asignados a esta hoja hasta el corte.").font = _X_f(False, 10, "595959", True)
        fila += 1
    for mes, g in mv.groupby("mes", sort=True):
        for row in g.itertuples(index=False):
            f = row.fecha.to_pydatetime() if pd.notna(row.fecha) else None
            fila = _X_escribir(ws, fila, [row.mes_txt, f, row.documento, row.tercero, row.texto, row.codigo_articulo, row.item, row.cuenta_nombre,
                                        row.sub_rubro_sap, row.cc_nombre] + sufijo(row.cantidad, _X_num(row.valor)), fmm)
        fila = _X_escribir(ws, fila, [f"Subtotal {g['mes_txt'].iloc[0]}"] + [None] * 9 + sufijo(g["cantidad"].sum(), float(g["valor"].sum())), fmm, "sub")
    fila = _X_escribir(ws, fila, ["TOTAL EJECUTADO"] + [None] * 9 + sufijo(mv["cantidad"].sum(), float(mv["valor"].sum())), fmm, "tot")
    ws.freeze_panes = "A4"


def _construir_informe(meta: dict, resumen: dict, secciones: list) -> bytes:
    wb = Workbook()
    _X_hoja_resumen(wb, meta, resumen)
    for s in secciones:
        _X_hoja_detalle(wb, meta, s)
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _sin_staf_mtto(df: pd.DataFrame) -> pd.Series:
    """True en los movimientos del centro de coste «Staf Mantenimiento» (Denominación del objeto)."""
    if df.empty or "cc_nombre" not in df:
        return pd.Series(False, index=df.index)
    return df["cc_nombre"].fillna("").map(L._canon).str.contains(r"STAFF?MANTENIMIENTO", regex=True)


_SECCIONES = [(OS.HOJA_PAPELERIA, "2 Papelería", "2. PAPELERÍA · MOVIMIENTOS DE PAPELERÍA"),
              (OS.HOJA_CONSUMIBLES, "3 Consumibles", "3. CONSUMIBLES"),
              (OS.HOJA_HERRAMIENTAS, "4 Herramientas", "4. HERRAMIENTAS"),
              (OS.HOJA_OTROS, "5 Bolsas", "5. BOLSAS")]


def _movimientos_informe(Q, hoja: str) -> pd.DataFrame:
    e = _ejec_hoja(Q, hoja)
    e = e[e["mes"] <= Q.hasta_mes].copy()
    for c in ("fecha", "documento", "tercero", "texto", "codigo_articulo", "item_vista", "cuenta_nombre", "sub_rubro_sap", "cc_nombre", "cantidad"):
        if c not in e:
            e[c] = None
    e["fecha"] = pd.to_datetime(e["fecha"], errors="coerce")
    e["valor"] = pd.to_numeric(e["valor"], errors="coerce").fillna(0.0)
    e["cantidad"] = pd.to_numeric(e["cantidad"], errors="coerce")
    e["item"] = e["item_vista"].where(e["item_vista"].notna(), "Sin ítem (SAP, sin detalle de SISPO)")
    e["mes"] = pd.to_numeric(e["mes"], errors="coerce").astype(int)
    e["mes_txt"] = e["mes"].map(lambda m: f"{MESES_L[m - 1].capitalize()} {Q.cfg['anio']}")
    for c in ("documento", "tercero", "texto", "codigo_articulo", "cuenta_nombre", "sub_rubro_sap", "cc_nombre"):
        e[c] = e[c].where(e[c].notna(), None)
    return e.sort_values(["mes", "fecha"])[["mes", "mes_txt", "fecha", "documento", "tercero", "texto", "codigo_articulo", "item",
                                             "cuenta_nombre", "sub_rubro_sap", "cc_nombre", "cantidad", "valor"]].reset_index(drop=True)


def _validaciones_informe(S2, secciones: list, resumen: dict, n_staf_sap: int) -> list:
    """Verificaciones sobre lo que realmente entra al informe: [(verificación, 'CUMPLE'|'NO CUMPLE'|'INFORMATIVO', detalle)]."""
    es_s = S2.ejec["sispo"].fillna(False).astype(bool) if "sispo" in S2.ejec else pd.Series(False, index=S2.ejec.index)
    sis = pd.concat([S2.ejec[es_s], S2.lin_pap], ignore_index=True) if len(S2.lin_pap) else S2.ejec[es_s]
    n_sap = int((~es_s).sum())
    presup = OS.items_presupuestados(S2.ppto_todo) if hasattr(S2, "ppto_todo") else OS.items_presupuestados(S2.ppto)
    cruza = sis["item"].fillna("").map(L._canon).isin(presup) if len(sis) and "item" in sis else False
    m = OS._mascaras(sis.assign(cruza=cruza, cc_codigo=sis["cc"] if "cc" in sis else "", grupo=sis["grupo"] if "grupo" in sis else "",
                                cuenta=sis["cuenta_nombre"].fillna("").str.replace("SISPO · ", "", regex=False) if "cuenta_nombre" in sis else "",
                                cc_nombre=sis["cc_nombre"] if "cc_nombre" in sis else "")) if len(sis) else {k: pd.Series(dtype=bool) for k in
                                                                                                      ("centro", "grupo", "cuenta", "staf_mtto")}
    v = [("Origen de los movimientos", "INFORMATIVO",
          f"{len(sis)} movimientos del SISPO y {n_sap} de SAP (compras y otros del rubro 3.1 que el módulo mantiene). Los filtros del SISPO se aplican "
          "al archivo original, antes de clasificar.")]

    def regla(nombre, clave, ok_txt):
        malos = int((~m[clave]).sum()) if len(sis) else 0
        v.append((nombre, "CUMPLE" if malos == 0 else "NO CUMPLE", ok_txt if malos == 0 else f"{malos} movimientos del SISPO no cumplen."))
    regla("Solo los 5 centros de costo autorizados (DB010500, 504, 515, 501, 509)", "centro", f"{len(sis)} movimientos del SISPO revisados.")
    regla("Solo grupos 5 (Insumos) y 7 (Herramientas), salvo ítems del presupuesto", "grupo", "Sin otros grupos que no crucen con el presupuesto.")
    regla("Solo cuentas ÚTILES, PAPELERÍA Y OTROS y OTROS", "cuenta", "Sin otras cuentas.")
    regla("Sin Staf Mantenimiento en el SISPO", "staf_mtto", f"Además se excluyeron {n_staf_sap} movimientos de SAP de ese centro de coste.")
    no_sal = int((pd.to_numeric(sis["cantidad"], errors="coerce").fillna(0) <= 0).sum()) if len(sis) and "cantidad" in sis else 0
    v.append(("Solo salidas (cantidades positivas = salidas de almacén)", "CUMPLE" if no_sal == 0 else "NO CUMPLE",
              "Sin entradas ni movimientos en cero." if no_sal == 0 else f"{no_sal} movimientos con cantidad cero o negativa."))
    dup = int(sis["documento"].duplicated().sum()) if len(sis) else 0
    v.append(("Sin registros duplicados (un Kardex por movimiento)", "CUMPLE" if dup == 0 else "NO CUMPLE",
              "Sin repetidos." if dup == 0 else f"{dup} movimientos repiten el número de Kardex."))
    malas = [s["hoja_excel"] for s in secciones if abs(float(s["movs"]["valor"].sum()) - float(s["cifras"]["ejecutado"])) >= 0.5]
    v.append(("El detalle de movimientos suma el ejecutado de cada hoja", "CUMPLE" if not malas else "NO CUMPLE",
              f"{len(secciones)} hojas revisadas." if not malas else "No cuadran: " + ", ".join(malas)))
    malas_mes = []
    for s in secciones:
        im = s["items_mes"]
        e_mes = sum(sum(v for v in f[2] if v is not None) for f in im["filas"])
        p_mes = sum(sum(f[1][:im["n_corte"]]) for f in im["filas"])
        if abs(e_mes - float(s["cifras"]["ejecutado"])) >= 0.5 or abs(p_mes - float(s["cifras"]["presupuesto"])) >= 0.5:
            malas_mes.append(s["hoja_excel"])
    v.append(("La tabla por ítem y por mes suma el presupuesto al corte y el ejecutado de cada hoja", "CUMPLE" if not malas_mes else "NO CUMPLE",
              "Presupuesto y ejecución mensual cuadran con los indicadores." if not malas_mes else "No cuadran: " + ", ".join(malas_mes)))
    suma = sum(float(s["cifras"]["ejecutado"]) for s in secciones if s["unidad"] == "pesos")
    tot = float(resumen["kpis"]["ejecutado"])
    v.append(("El resumen general coincide con la suma de las hojas en pesos", "CUMPLE" if abs(suma - tot) < 0.5 else "NO CUMPLE",
              f"Resumen {tot:,.2f} · hojas {suma:,.2f}".replace(",", " ")))
    return v


def _items_mes(S, hoja: str) -> dict:
    """Presupuesto y ejecución mes a mes por ítem, con el mismo criterio de `_items` (mismos ítems y mismo orden; lo de SAP sin detalle de
    SISPO va en su fila aparte). Después del mes de corte la ejecución queda en None."""
    it = _items(S, hoja)
    meses = list(range(S.m0, S.m1 + 1))
    p = _ppto_hoja(S, hoja)
    p = p.assign(_k=p["item"].map(L._canon)).groupby(["_k", "mes"])["valor"].sum() if len(p) else pd.Series(dtype=float)
    e = _ejec_hoja(S, hoja)
    e = e[e["mes"] <= S.hasta_mes]
    if len(e):
        es_s = e["sispo"].fillna(False).astype(bool) if "sispo" in e else pd.Series(False, index=e.index)
        iv = e["item_vista"] if "item_vista" in e else pd.Series("", index=e.index)
        k = iv.fillna("").map(L._canon).where(es_s, "__SAP__")
        e = e.assign(_k=k, valor=pd.to_numeric(e["valor"], errors="coerce").fillna(0.0)).groupby(["_k", "mes"])["valor"].sum()
    else:
        e = pd.Series(dtype=float)
    filas = []
    for item in it["item"]:
        k = "__SAP__" if item == "Otros movimientos de SAP (sin ítem)" else L._canon(item)
        filas.append((item, [float(p.get((k, m), 0.0)) for m in meses],
                      [float(e.get((k, m), 0.0)) if m <= S.hasta_mes else None for m in meses]))
    return dict(meses=[f"{MESES_L[m - 1].capitalize()} {S.cfg['anio']}" for m in meses], filas=filas, n_corte=S.hasta_mes - S.m0 + 1)


def _mensual_hojas(S2, hojas: list, con_sin_hoja: bool) -> dict:
    """Presupuesto y ejecución mes a mes por hoja (para el resumen). Sale de `_mensual`, la misma tabla de la pantalla."""
    meses = list(range(S2.m0, S2.m1 + 1))
    filas = []
    for h in hojas + ([SIN_HOJA] if con_sin_hoja else []):
        t = _mensual(S2, h)
        t = t[t["mes"] > 0].sort_values("mes")
        filas.append((h, [float(v) for v in t["presupuesto"]], [None if pd.isna(v) else float(v) for v in t["ejecutado"]]))
    return dict(meses=[f"{MESES_L[m - 1].capitalize()} {S2.cfg['anio']}" for m in meses], filas=filas, n_corte=S2.hasta_mes - S2.m0 + 1)


def _informe(S) -> bytes:
    """Informe Excel: resumen general + una hoja por Papelería, Consumibles, Herramientas y Bolsas. Todo sale de las mismas funciones que la
    pantalla (`_tabla_resumen`, `_cifras`, `_items`, `_mensual`). Excluye «Staf Mantenimiento» de la ejecución, también en el resumen, para
    que las cifras del informe cuadren entre sí; la conciliación contra la pantalla va en el resumen."""
    S2 = SimpleNamespace(**vars(S))
    S2.ejec = S.ejec[~_sin_staf_mtto(S.ejec)]
    S2.lin_pap = S.lin_pap[~_sin_staf_mtto(S.lin_pap)] if len(S.lin_pap) else S.lin_pap
    notas = ["Fuente: presupuesto cargado en el módulo y ejecución tomada solo del SISPO (SAP no aporta valores a Operaciones). "
             "El informe usa los mismos cálculos de la pantalla.",
             "Variación = ejecutado − presupuesto al corte. Saldo disponible = presupuesto del escenario − ejecutado. "
             "Los movimientos se muestran con el signo de SAP (costo positivo)."]
    res = _tabla_resumen(S2)
    base = _tabla_resumen(S)
    kp = None
    if res is not None:
        t, tot, m, _ = res
        kp = dict(presupuesto_esc=tot["presupuesto_esc"], presupuesto=tot["presupuesto"], ejecutado=tot["ejecutado"],
                  saldo=tot["presupuesto_esc"] - tot["ejecutado"], pct=m["pct"], variacion=tot["ejecutado"] - tot["presupuesto"], estado=m["estado"])
    resumen = dict(kpis=kp or dict(presupuesto_esc=0.0, presupuesto=0.0, ejecutado=0.0, saldo=0.0, pct=float("nan"), variacion=0.0, estado="—"),
                   tabla=res[0] if res is not None else pd.DataFrame(columns=["hoja", "presupuesto_esc", "presupuesto", "ejecutado", "variacion",
                                                                            "pct", "saldo", "estado", "movimientos"]))
    if base is not None and res is not None:
        pantalla, informe = float(base[1]["ejecutado"]), float(res[1]["ejecutado"])
        if abs(pantalla - informe) >= 0.5:
            resumen["concilia"] = [("Ejecutado en pesos en la pantalla del módulo", pantalla),
                                   ("(−) Centro de coste «Staf Mantenimiento» (excluido del informe)", informe - pantalla),
                                   ("Ejecutado en pesos en este informe", informe)]
            notas.append("El presupuesto del libro no está desglosado por centro de coste, por lo que solo se excluye la ejecución de Staf Mantenimiento.")
    if OS.HOJA_PAPELERIA in S.hojas:
        q = _vista_cantidades(S2)
        if len(_ppto_hoja(q, OS.HOJA_PAPELERIA)) or len(q.ejec):
            resumen["papeleria"] = _cifras(q, OS.HOJA_PAPELERIA)
    hay_sin = bool(len(S2.ejec[(S2.ejec["hoja"] == SIN_HOJA) & (S2.ejec["mes"] <= S2.hasta_mes)]))
    resumen["mensual_hojas"] = _mensual_hojas(S2, [h for h in S.hojas if h != OS.HOJA_PAPELERIA], hay_sin)
    if resumen.get("papeleria"):
        qp = _vista_cantidades(S2)
        resumen["mensual_pap"] = _mensual_hojas(qp, [OS.HOJA_PAPELERIA], False)
    secciones = []
    for hoja, hoja_xl, titulo in _SECCIONES:
        if hoja not in S.hojas:
            continue
        Q = _vista_cantidades(S2) if hoja == OS.HOJA_PAPELERIA else S2
        secciones.append(dict(hoja_excel=hoja_xl, titulo=titulo, unidad="unidades" if hoja == OS.HOJA_PAPELERIA else "pesos",
                              cifras=_cifras(Q, hoja), mensual=_mensual(Q, hoja), items=_items(Q, hoja), items_mes=_items_mes(Q, hoja),
                              movs=_movimientos_informe(Q, hoja)))
    if len(S2.ejec[(S2.ejec["hoja"] == SIN_HOJA) & (S2.ejec["mes"] <= S2.hasta_mes)]):      # para que el total del resumen cuadre con las hojas
        secciones.append(dict(hoja_excel="6 Sin hoja asignada", titulo="6. EJECUCIÓN SIN HOJA ASIGNADA", unidad="pesos", cifras=_cifras(S2, SIN_HOJA),
                              mensual=_mensual(S2, SIN_HOJA), items=_items(S2, SIN_HOJA), items_mes=_items_mes(S2, SIN_HOJA),
                              movs=_movimientos_informe(S2, SIN_HOJA)))
        notas.append("La hoja «6 Sin hoja asignada» reúne ejecución del rubro 3.1 que aún no se reparte (SISPO sin cargar o sin cuadrar, o sub-rubro "
                     "sin asignar). Se incluye para que el total del resumen coincida con la suma de las hojas.")
    avisos = list(S.avisos)
    if "pend_sispo" in S2.ejec:
        pend = S2.ejec[S2.ejec["pend_sispo"].fillna(False).astype(bool) & (S2.ejec["mes"] <= S2.hasta_mes)]
        for mes, g in pend.groupby("mes"):
            avisos.append(f"{MESES_L[int(mes) - 1].capitalize()} {S.cfg['anio']}: {len(g)} movimientos por {g['valor'].sum():,.0f} de «Salida de inventario» "
                          "(cuenta OTROS) no se repartieron entre Consumibles, Herramientas y Bolsas: falta el SISPO completo de ese mes o no cuadra con SAP.")
    resumen["avisos"] = avisos
    resumen["validaciones"] = _validaciones_informe(S2, secciones, resumen, int(_sin_staf_mtto(S.ejec).sum()))
    for s in secciones:                                      # control: lo listado debe sumar lo que dice el indicador
        if abs(float(s["movs"]["valor"].sum()) - float(s["cifras"]["ejecutado"])) >= 0.5:
            notas.append(f"ADVERTENCIA: en «{s['hoja_excel']}» los movimientos no suman el ejecutado del indicador.")
    sel_txt = " + ".join(S.sel)
    meta = dict(subtitulo=f"Compañía: {sel_txt} · {S.escenario} · ejecutado {S.hasta} · generado {ahora().strftime('%d/%m/%Y %H:%M')}", notas=notas)
    return _construir_informe(meta, resumen, secciones)


def _descarga_ejecucion(S):
    st.download_button("Descargar informe (Excel)", lambda: _informe(S), f"informe_operaciones_{S.escenario.replace(' ', '_')}_{S.hasta_mes:02d}.xlsx",
                       icon=":material/download:", key="op_dl_informe",
                       help="Resumen general + Papelería, Consumibles, Herramientas y Bolsas. Sin los movimientos de Staf Mantenimiento.")


def _mostrar(ctx):
    leer, usuario, es_admin = ctx.leer, ctx.usuario, ctx.es_admin
    ver_todo = getattr(ctx, "ver_todo", es_admin)
    _asegurar_tablas(ctx)
    zona_encabezado = st.container()
    with st.container(key="op_barra"):
        c_emp, c_esc, c_hasta = st.columns([1, 1, 1], vertical_alignment="bottom")
        vista_emp = c_emp.selectbox("Compañía", [TODAS, *EMPRESAS], key="op_empresa",
                                    help="Consolidado = ALDC + ALSM. " + " · ".join(f"{e}: {EMPRESAS[e]}" for e in EMPRESAS))
        escenario = c_esc.selectbox("Escenario presupuestal", list(L.ESCENARIOS), key="op_escenario",
                                    help="UPA 2026: sep–dic 2026. Normal 2027: ene–dic 2027. Sus cifras nunca se mezclan.")
    cfg = dict(L.ESCENARIOS[escenario], escenario=escenario)
    sel = list(EMPRESAS) if vista_emp == TODAS else [vista_emp]
    m0, m1 = cfg["meses"]

    mapeo = {r.sub_norm: r.hoja for r in leer("SELECT sub_norm, hoja FROM oper_mapeo").itertuples()}
    asig = asignacion.mapa(leer)
    lin_sispo = leer("SELECT empresa, periodo, hoja, id_kardex, fecha, bodega, codigo, articulo, grupo, cuenta, cc_codigo, cc_nombre, cantidad, valor, concepto "
                     "FROM oper_sispo")
    per_sispo = leer("SELECT empresa, periodo, total_otros FROM oper_sispo_periodo")
    for c in ("periodo", "cantidad", "valor"):
        if c in lin_sispo:
            lin_sispo[c] = pd.to_numeric(lin_sispo[c], errors="coerce")
    per_sispo["periodo"] = pd.to_numeric(per_sispo["periodo"], errors="coerce")
    per_sispo["total_otros"] = pd.to_numeric(per_sispo["total_otros"], errors="coerce")
    items_eq = OS.mapa_items(leer("SELECT codigo, item FROM oper_item_map").itertuples(index=False, name=None))
    ppto_op = leer("SELECT empresa, hoja, item, orden, anio, mes, cantidad, costo_unit, valor FROM oper_ppto")
    for c in ("anio", "mes", "orden", "cantidad", "costo_unit", "valor"):
        ppto_op[c] = pd.to_numeric(ppto_op[c], errors="coerce")
    ppto_op["hoja"] = ppto_op["hoja"].replace({"Bolsas": OS.HOJA_OTROS})              # la hoja «Bolsas» pasó a llamarse «Otros»
    lin_sispo = OS.filtrar_guardado(lin_sispo, items_eq, OS.items_presupuestados(ppto_op))   # mismos filtros del SISPO sobre lo ya guardado
    lin_sispo = OS.hoja_por_item(lin_sispo, items_eq, OS.hojas_de_items(ppto_op))      # la lista presupuestada manda sobre el Grupo
    datos_e = {e: _datos_empresa(ctx, e, cfg, asig, mapeo, SimpleNamespace(lin=lin_sispo, per=per_sispo, items=items_eq)) for e in EMPRESAS}
    ejec_todo = pd.concat([d.ejec for d in datos_e.values()], ignore_index=True)
    ult = int(ejec_todo["mes"].max()) if len(ejec_todo) else m1
    hasta_mes = c_hasta.selectbox("Ejecución hasta", list(range(m0, m1 + 1)), index=min(max(ult, m0), m1) - m0, key=f"op_hasta_{escenario}",
                                  format_func=lambda m: MESES_L[m - 1].capitalize())
    hasta = f"{MESES_L[hasta_mes - 1]} {cfg['anio']}" if m0 == hasta_mes else (
        f"{MESES_L[m0 - 1]}–{MESES_L[hasta_mes - 1]} {cfg['anio']}" if m0 != 1 else f"a {MESES_L[hasta_mes - 1]} {cfg['anio']}")

    ppto_v = ppto_op[(ppto_op["anio"] == cfg["anio"]) & ppto_op["mes"].between(m0, m1) & ppto_op["empresa"].isin(sel)]
    ejec = pd.concat([datos_e[e].ejec for e in sel], ignore_index=True)

    S = SimpleNamespace(ctx=ctx, cfg=cfg, escenario=escenario, sel=sel, multi=len(sel) > 1, m0=m0, m1=m1, hasta_mes=hasta_mes, hasta=hasta,
                        ppto=ppto_v, ppto_todo=ppto_op, ejec=ejec, mapeo=mapeo, datos_e=datos_e, ver_todo=ver_todo, cant=False, items=items_eq, lin_sispo=lin_sispo, per_sispo=per_sispo,
                        lin_pap=(pd.concat([datos_e[e].lin_pap for e in sel], ignore_index=True) if ver_todo else OS.papeleria(pd.DataFrame(), cfg, "", {})),
                        avisos=[a for e in sel for a in datos_e[e].avisos])
    zona_encabezado.html(E.encabezado(
        "Operaciones" + ("" if ver_todo else f" · {usuario}"),
        "Operaciones · Papelería · Consumibles · Herramientas · Bolsas: presupuesto vs ejecución · valores en pesos · " + VERSION,
        [(vista_emp if not S.multi else f"Consolidado · {' + '.join(sel)}", "filtro")]))

    if not ver_todo:                                    # solo ve lo que su gerencia tiene asignado
        propias = set()
        for d in datos_e.values():
            if len(d.ejec):
                propias |= set(d.ejec.loc[d.ejec["gerencia"] == usuario, "hoja"])
            if len(d.ppto_base):
                propias |= set(d.ppto_base.loc[d.ppto_base["gerencia"] == usuario, "hoja"])
        propias -= {SIN_HOJA}
        S.ejec = S.ejec[S.ejec["gerencia"] == usuario]
        S.ppto = S.ppto[S.ppto["hoja"].isin(propias)]
        S.hojas = [h for h in HOJAS if h in propias]
        if not S.hojas:
            st.info("Tu gerencia no tiene asignado ningún sub-rubro del rubro 3.1 de Operaciones.")
            return
    else:
        S.hojas = list(HOJAS)

    if ppto_op.empty:
        st.warning("Aún no hay presupuesto de Operaciones cargado." + (" Cárgalo abajo, en «Datos base»." if es_admin
                                                                      else " El administrador lo carga en «Datos base»."), icon=":material/warning:")
        if es_admin:
            _admin(S)
        return
    if ejec.empty:
        st.info(f"No hay ejecución del rubro {RUBRO_COD} en {escenario} para {vista_emp}: se muestra el presupuesto y la ejecución sale en cero "
                "hasta que se cargue en el módulo de ejecución.", icon=":material/info:")

    for aviso in S.avisos:                                  # SISPO que difiere de SAP, o mes con SAP y sin SISPO
        st.warning(aviso, icon=":material/rule:")
    pestanas = st.tabs(["Resumen", *S.hojas])
    with pestanas[0]:
        _tab_resumen(S)
    for tab, hoja in zip(pestanas[1:], S.hojas):
        with tab:
            _tab_hoja(S, hoja)
    _descarga_ejecucion(S)
    if es_admin:
        _admin(S)


def mostrar(ctx):
    """Punto de entrada. Si algo falla, muestra el error completo (Streamlit lo oculta en la nube)."""
    try:
        _mostrar(ctx)
    except Exception:
        st.error("El módulo de Operaciones tuvo un error. Copia este detalle y compártelo para corregirlo.")
        st.code(traceback.format_exc())
