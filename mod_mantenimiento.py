"""
Módulo Mantenimiento: presupuesto vs ejecución · Flota y Equipo de transporte, Tecnomecánica/SOAT y Contenedores
================================================================================================================
Una sola pantalla con cuatro pestañas (se abren hasta el MOVIMIENTO, solo para el administrador):
- Mantenimiento                <- resumen del rubro por sub-rubro (sin movimientos).
- Flota y Equipo de Transporte <- concepto presupuestal (descripción TIPO VEHICULO del archivo de presupuesto, incluidos los
                                  generales sin placa: In House, grúa, lavado…) -> placa -> concepto de gasto -> movimientos.
- Tecnomecánica y SOAT         <- dos conceptos independientes; presupuesto mensual por placa del archivo «Soat y Tecno».
- Contenedores                 <- vista simple: presupuesto, ejecutado, saldo y % mes a mes.

Compañías: ALDC y ALSM se cargan por separado y cada movimiento conserva la suya. Arriba se elige «Consolidado» (suma de
las dos) o una compañía. El listado oficial de placas es uno solo, compartido.

Un movimiento tiene UN concepto, y el concepto define el ámbito donde aparece (no se cuenta dos veces).
Identificador único de movimiento: «compañía|clave de SAP» (con sufijo #2, #3… si dos fueran idénticos).
Orden de clasificación: reasignación manual > histórico aprendido > reglas editables > pendiente.
La placa sale del texto (solo contra el listado oficial) y se puede corregir a mano.
Reclasificar o eliminar un movimiento queda en mtto_clasif_log (usuario, fecha, antes, después, motivo). Eliminar NO borra la
ejecución de la base: guarda el movimiento en mtto_excluidos, así no reaparece en cargas posteriores y se puede restaurar.

El selector de PLACA es global: filtra los DATOS (presupuesto, ejecución, saldo, %, gráficos, tablas, detalle y Excel) de
Flota y Tecnomecánica/SOAT. Los conceptos generales no se atribuyen a placas: con una placa elegida no se muestran.
Escenario UPA 2026 = sep–dic 2026; Normal 2027 = ene–dic 2027. Sus cifras nunca se mezclan.

La lógica de cálculo y el Excel (VEHÍCULOS, TECNOMECÁNICAS, SOAT) están en logica_m.py (sin Streamlit).
Tablas propias: mtto_placas, mtto_ppto, mtto_ppto_ts, mtto_reglas, mtto_clasif_manual, mtto_clasif_log y mtto_excluidos.

app.py lo llama como: mod_mantenimiento.mostrar(ctx)
"""
import math
import traceback
from html import escape
from types import SimpleNamespace

import openpyxl
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import asignacion
import cargas
import datos
import estilo as E
import logica_m as L
from ui import EMPRESAS, MESES, MESES_L, a_excel, ahora, estilo_grafico, titulo_tarjeta

# Estado que dejó la prueba de octubre 2026 en `mtto_sispo_periodo` (aplicado sin conciliar con SAP). Ya no se puede generar: solo
# sirve para reconocer esos registros heredados, excluirlos de los cálculos y avisar que hay que revertirlos y volver a aplicar.
ESTADO_APLICACION_SIN_SAP_HEREDADA = "Prueba sin SAP"

TABLAS = """
CREATE TABLE IF NOT EXISTS mtto_placas(placa TEXT, numero_interno TEXT, tipo TEXT);
CREATE TABLE IF NOT EXISTS mtto_ppto(empresa TEXT, escenario TEXT, placa TEXT, numero_interno TEXT, tipo TEXT,
    concepto TEXT, orden INT, valor DOUBLE PRECISION, es_global INT, texto_original TEXT);
CREATE TABLE IF NOT EXISTS mtto_ppto_ts(empresa TEXT, placa TEXT, concepto TEXT, anio INT, mes INT,
    valor DOUBLE PRECISION, texto_original TEXT);
CREATE TABLE IF NOT EXISTS mtto_clasif_manual(empresa TEXT, llave TEXT, concepto TEXT, texto_tipo TEXT, tercero TEXT,
    origen TEXT, usuario TEXT, fecha TEXT, placa TEXT);
CREATE TABLE IF NOT EXISTS mtto_clasif_log(empresa TEXT, fecha TEXT, usuario TEXT, llave TEXT, texto TEXT,
    valor DOUBLE PRECISION, concepto_antes TEXT, concepto_despues TEXT, placa_antes TEXT, placa_despues TEXT, accion TEXT,
    motivo TEXT);
CREATE TABLE IF NOT EXISTS mtto_excluidos(empresa TEXT, llave TEXT, fecha_mov TEXT, documento TEXT, tercero TEXT, texto TEXT,
    valor DOUBLE PRECISION, concepto TEXT, placa TEXT, motivo TEXT, usuario TEXT, fecha TEXT);
CREATE TABLE IF NOT EXISTS mtto_reglas(orden DOUBLE PRECISION, patron TEXT, proveedor TEXT, concepto TEXT,
    confianza TEXT, observacion TEXT, activa INT);
CREATE TABLE IF NOT EXISTS mtto_sispo_ref(codigo TEXT, articulo TEXT, concepto TEXT, usuario TEXT, fecha TEXT);
CREATE TABLE IF NOT EXISTS mtto_sispo_periodo(empresa TEXT, periodo INT, archivo TEXT, huella TEXT, huella_maestro TEXT,
    usuario TEXT, fecha TEXT, n_lineas INT, n_pendientes INT, valor_sispo DOUBLE PRECISION, valor_sap DOUBLE PRECISION,
    diferencia DOUBLE PRECISION, estado TEXT, n_reemplazados INT, generico TEXT, cuenta TEXT DEFAULT 'FLOTA Y EQUIPO DE TR');
CREATE TABLE IF NOT EXISTS mtto_sispo_lineas(empresa TEXT, periodo INT, id_kardex TEXT, bodega TEXT, fecha_mov TEXT, codigo TEXT,
    articulo TEXT, cantidad DOUBLE PRECISION, costo_unidad DOUBLE PRECISION, costo_sin_iva DOUBLE PRECISION,
    costo_con_iva DOUBLE PRECISION, centro_costo TEXT, cc_codigo TEXT, cc_nombre TEXT, cuenta_nombre TEXT, placa TEXT,
    placa_estado TEXT, concepto TEXT, concepto_estado TEXT, valor_sap DOUBLE PRECISION);
CREATE TABLE IF NOT EXISTS mtto_sispo_reemplazados(empresa TEXT, periodo INT, llave TEXT, fecha_mov TEXT, documento TEXT, cc TEXT,
    texto TEXT, valor DOUBLE PRECISION, gerencia TEXT, codigo TEXT, cuenta TEXT DEFAULT 'FLOTA Y EQUIPO DE TR',
    sub_rubro TEXT DEFAULT '', sub_rubro_sap TEXT DEFAULT '');
"""
# columnas agregadas a tablas que ya existían en versiones anteriores del módulo
MIGRACIONES = {"mtto_ppto": {"texto_original": "TEXT DEFAULT ''"}, "mtto_clasif_manual": {"placa": "TEXT DEFAULT ''"},
               "mtto_clasif_log": {"motivo": "TEXT DEFAULT ''"},
               # la reclasificación SISPO pasó de una cuenta (flota) a siete: lo ya aplicado queda como de la flota
               "mtto_sispo_periodo": {"cuenta": "TEXT DEFAULT 'FLOTA Y EQUIPO DE TR'"},
               "mtto_sispo_reemplazados": {"cuenta": "TEXT DEFAULT 'FLOTA Y EQUIPO DE TR'", "sub_rubro": "TEXT DEFAULT ''",
                                           "sub_rubro_sap": "TEXT DEFAULT ''"}}
ESTADOS = ["Sobre ejecutado", "Presupuesto agotado", "Dentro del presupuesto", "Sin presupuesto asignado",
           "Sin ejecución", "Pendiente de clasificación"]
AUTO = "(clasificación automática)"
AUTO_PLACA = "(placa automática)"
TODAS = "Consolidado"                  # opción del selector de compañía
TODAS_PLACAS = "Todas las placas"      # opción del selector global de placa
MAX_MOV_ARBOL = 2500                   # por encima, el árbol no incrusta los movimientos (se ven en la tabla de auditoría)
PAGINA = 200                           # filas por página en la reclasificación


# ================================================================== datos
def _asegurar_tablas(ctx):
    if not st.session_state.get("_mtto_tablas7"):
        ctx.con.crear_tablas(TABLAS)
        for tabla, cols in MIGRACIONES.items():
            existentes = ctx.con.columnas(tabla)
            for c, tipo in cols.items():
                if c not in existentes:
                    ctx.con.execute(f"ALTER TABLE {tabla} ADD COLUMN {c} {tipo}")
        ctx.con.commit()
        st.session_state["_mtto_tablas7"] = True


def _version_para(ctx, empresa, anio, preferida=None):
    vs = ctx.leer("SELECT DISTINCT version FROM presupuesto WHERE anio=? AND empresa=?", (int(anio), empresa))["version"].tolist()
    orden = {"Inicial": 0, "4+8": 1, "8+4": 2}
    vs = sorted(vs, key=lambda v: (orden.get(v, 9), v))
    if preferida in vs:
        return preferida
    return "8+4" if "8+4" in vs else (vs[-1] if vs else "Inicial")


def _base(ctx, empresa, anio, version):
    """Presupuesto (tabla presupuesto) y ejecución del año, con el mismo tratamiento que el módulo Presupuesto
    (homologación de sub-rubros) para que los totales coincidan."""
    leer = ctx.leer
    ppto = leer("SELECT * FROM presupuesto WHERE anio=? AND version=? AND empresa=?", (int(anio), version, empresa))
    cols_e = ", ".join(sorted(ctx.con.columnas("ejecucion") - {"raw"}))      # raw (fila original de SAP) pesa: se pide aparte
    ejec = leer(f"SELECT {cols_e} FROM ejecucion WHERE anio=? AND empresa=?", (int(anio), empresa))
    # Clase de coste OTROS con objeto «Staf Mantenimiento»: gasto del área que el presupuesto no trae en la flota. Se pasa al sub-rubro
    # «Flota y Equipo de transporte» (solo en esta pantalla; la base no se toca) y se marca para que `_rubro` los traiga aunque su
    # código de presupuesto sea otro o no tengan ninguno. Se detecta por las columnas de la base y, si así no aparece ninguno, por la
    # fila original de SAP (`raw`, filtrada en la consulta para no leerla completa).
    ids_staf = L.ids_staf_otros_ejec(leer, ejec, anio, empresa)
    ejec["_staf_otros"] = ejec["id"].isin(ids_staf).to_numpy() if len(ejec) and "id" in ejec else False
    sin_codigo = ejec["codigo"].fillna("") == ""
    ejec = ejec[~sin_codigo | ejec["_staf_otros"]].copy()               # sin código solo se conservan las OTROS/Staf (se pasan a Flota)
    ejec["rubro"] = ejec["codigo"].map(lambda c: cargas.rubro_de(c) if isinstance(c, str) and c else "")
    ejec["sub_rubro_sap"] = ejec["sub_rubro"]                       # nombre original, para rastrear
    if not ejec.empty:
        conc = ejec["codigo"].fillna("") == ""                       # estas no pasan por la homologación (no tienen código que homologar)
        h = datos.homologar(ejec[~conc].copy(), ppto, leer("SELECT * FROM homologacion"), version) if (~conc).any() else ejec[~conc]
        ejec = pd.concat([h, ejec[conc]], ignore_index=True) if conc.any() else h
        ejec["_staf_otros"] = ejec["id"].isin(ids_staf).to_numpy() if "id" in ejec else False
        if ejec["_staf_otros"].any():
            ejec.loc[ejec["_staf_otros"], "sub_rubro"] = L.SUBRUBRO_FLOTA
    else:
        ejec["_staf_otros"] = pd.Series(dtype=bool)
    ppto["_sub"] = ppto["sub_rubro"].fillna("").map(cargas.norm)
    ejec["_sub"] = ejec["sub_rubro"].fillna("").map(cargas.norm)
    return ppto, ejec


def _con_gerencia(df: pd.DataFrame, asig) -> pd.DataFrame:
    df = df.copy()
    df["gerencia"] = asignacion.columna(df, asig) if not df.empty else pd.Series(dtype=str)
    return df


def _solo_sub(df: pd.DataFrame, sub: str) -> pd.DataFrame:
    return df[df["_sub"] == cargas.norm(sub)]


def _universo(ejec: pd.DataFrame, asig) -> pd.DataFrame:
    """Movimientos que alimentan la pantalla: todo el sub-rubro «Flota y Equipo de transporte» MÁS los movimientos de
    SOAT o tecnomecánica (por su texto o cuenta) que estén en otro sub-rubro, porque el archivo de presupuesto de SOAT
    corresponde a otro rubro (3,7). Los traídos de fuera se marcan en `fuera_flota`. Ningún movimiento queda dos veces."""
    flota = _con_gerencia(_solo_sub(ejec, L.SUBRUBRO_FLOTA), asig).assign(fuera_flota=False)
    resto = ejec[ejec["_sub"] != cargas.norm(L.SUBRUBRO_FLOTA)]
    extra = _con_gerencia(resto[L.candidatos_ts(resto)], asig).assign(fuera_flota=True)
    return pd.concat([flota, extra], ignore_index=True)


def _datos_empresa(ctx, empresa: str, cfg: dict, asig) -> SimpleNamespace:
    """Todo lo de UNA compañía: presupuesto base, ejecución, presupuesto de flota y de Tecnomecánica/SOAT. Cada compañía se
    carga y se lee por separado; sus movimientos salen marcados con su compañía."""
    leer = ctx.leer
    version = _version_para(ctx, empresa, cfg["anio"])
    ppto_base, ejec = _base(ctx, empresa, cfg["anio"], version)
    ppto = leer("SELECT empresa, placa, numero_interno, tipo, concepto, orden, valor, es_global, texto_original FROM mtto_ppto "
                "WHERE empresa=? AND escenario=?", (empresa, cfg["escenario"]))
    ts_mes = leer("SELECT empresa, placa, concepto, anio, mes, valor, texto_original FROM mtto_ppto_ts WHERE empresa=?", (empresa,))
    ts_per = L.ts_periodo(ts_mes, cfg).assign(empresa=empresa)
    pool = _universo(ejec, asig).assign(empresa=empresa)
    return SimpleNamespace(empresa=empresa, ppto_base=ppto_base, ejec=ejec, ppto=ppto, ts_mes=ts_mes, ts_per=ts_per,
                           pool=pool, ppto_all=pd.concat([ppto, ts_per], ignore_index=True))


def _normalizar_llaves(df: pd.DataFrame) -> pd.DataFrame:
    """Las reasignaciones guardadas en versiones anteriores no llevaban la compañía en la llave. Se les agrega al leerlas,
    sin tocar la base: así siguen aplicando a su movimiento."""
    if df.empty or "llave" not in df or "empresa" not in df:
        return df
    d = df.copy()
    ll, em = d["llave"].fillna("").astype(str), d["empresa"].fillna("").astype(str)
    falta = (ll != "") & (em != "") & pd.Series([not l.startswith(e + "|") for l, e in zip(ll, em)], index=d.index)
    d.loc[falta, "llave"] = em[falta] + "|" + ll[falta]
    return d


def _pool_sap(datos_e: dict) -> pd.DataFrame:
    """Movimientos de SAP de las compañías, tal como están en la base, con su llave única. Se calcula sobre TODO el año: el sufijo
    #2 de un movimiento gemelo no cambia con el corte de mes."""
    todo = pd.concat([d.pool for d in datos_e.values()], ignore_index=True)
    todo["llave"] = L.llave_movimiento(todo)
    return todo


def _pool_salidas(datos_e: dict, asig) -> pd.DataFrame:
    """Movimientos genéricos «Salida de inventario…» de SAP de TODAS las cuentas y rubros de la compañía-año (no solo la flota), con su
    llave y su gerencia. Es de donde la reclasificación toma lo que SISPO viene a reemplazar; solo se arma para el administrador."""
    partes = []
    for e, d in datos_e.items():
        x = d.ejec
        if x.empty:
            continue
        x = x[x["texto"].fillna("").astype(str).str.upper().str.contains("SALIDA", regex=False)]      # filtro barato antes de normalizar
        x = x[x["texto"].map(L.norm_texto).str.contains(L.SISPO_GENERICO, regex=False)]
        if len(x):
            partes.append(_con_gerencia(x, asig).assign(empresa=e, fuera_flota=False))
    if not partes:
        return pd.DataFrame(columns=["empresa", "anio", "mes", "texto", "valor", "llave", "fuera_flota"])
    todo = pd.concat(partes, ignore_index=True)
    todo["llave"] = L.llave_movimiento(todo)
    return todo


def _sispo_filas(lin: pd.DataFrame, rep: pd.DataFrame, solo_gerencia: str | None) -> pd.DataFrame:
    """Líneas de SISPO ya aplicadas, convertidas a movimientos de la misma forma que los de SAP (construcción central en
    `L.sispo_a_movimientos`: el valor queda con el signo de SAP y código, sub-rubro y gerencia salen de lo que reemplazan). Aquí solo
    se completan las columnas propias de esta pantalla y se respeta quién puede ver qué."""
    n = L.sispo_a_movimientos(lin, rep)
    if n.empty:
        return n
    n["_sub"] = n["sub_rubro"].map(cargas.norm)
    n["rubro"] = n["codigo"].map(cargas.rubro_de)
    if solo_gerencia is not None:
        n = n[n["gerencia"] == solo_gerencia]
    n["llave"] = L.llave_movimiento(n)
    return n.reset_index(drop=True)


def _movimientos_maestro(datos_e: dict, cfg: dict, hasta_mes: int, oficial: pd.DataFrame, reglas, manual, catalogo,
                         tipos: dict, excluidos: pd.DataFrame, pool_sap: pd.DataFrame | None = None, sispo=None,
                         solo_gerencia: str | None = None):
    """UNA sola base de movimientos para toda la pantalla, con ambas compañías: placa, llave única, concepto, ámbito.
    Devuelve (todos, vigentes, excluidos): `vigentes` es lo que alimenta cálculos y Excel; los excluidos (eliminados por
    el usuario) quedan fuera por su llave, también en cargas posteriores.
    `sispo` = (líneas aplicadas, movimientos de SAP reemplazados): los movimientos genéricos de SAP salen y entran, en su lugar,
    las líneas de SISPO con su placa y su concepto de la base maestra (los cambios manuales siguen mandando)."""
    m0 = cfg["meses"][0]
    todo = pool_sap if pool_sap is not None else _pool_sap(datos_e)
    if sispo is not None and len(sispo[0]):
        todo = todo[~todo["llave"].isin(set(sispo[1]["llave"]))]
        nuevos = _sispo_filas(sispo[0], sispo[1], solo_gerencia)
        if len(nuevos):
            if pd.api.types.is_datetime64_any_dtype(todo["fecha"]) if len(todo) else False:
                nuevos["fecha"] = pd.to_datetime(nuevos["fecha"], errors="coerce")
            todo = pd.concat([todo, nuevos], ignore_index=True)
    mov = todo[(todo["mes"] >= m0) & (todo["mes"] <= hasta_mes)].reset_index(drop=True)
    mov = L.asignar_placas(mov, oficial["placa"].tolist())
    es_s = mov["sispo"].fillna(False).astype(bool) if "sispo" in mov else pd.Series(False, index=mov.index)
    if es_s.any():                                      # la placa de SISPO sale de su centro de costos, no del texto del artículo
        mov.loc[es_s, "placa"] = mov.loc[es_s, "placa_sispo"]
        mov.loc[es_s, "placa_estado"] = mov.loc[es_s, "placa_estado_sispo"]
    mov = L.aplicar_placas_manuales(mov, manual)
    mov = L.clasificar(mov, reglas, manual, catalogo)
    if es_s.any():                                      # el concepto sale SOLO de la base maestra (o queda pendiente); lo manual manda
        s = es_s & (mov["origen"] != "Manual")
        pend = s & (mov["concepto_sispo"] == L.PENDIENTE)
        mov.loc[s, "concepto"] = mov.loc[s, "concepto_sispo"]
        mov.loc[s, "concepto_sugerido"] = ""
        mov.loc[s, "regla"] = None
        for filtro, conf, tipo, origen in ((s & ~pend, "Alta", "Base maestra SISPO", "SISPO"), (pend, "", "", "Pendiente")):
            mov.loc[filtro, "confianza"] = conf
            mov.loc[filtro, "tipo_clasif"] = tipo
            mov.loc[filtro, "origen"] = origen
    mov = L.asignar_ambito(mov.assign(valor=pd.to_numeric(mov["valor"], errors="coerce").fillna(0.0)), pd.DataFrame(), tipos)
    fuera = set(excluidos["llave"]) if len(excluidos) else set()
    es_fuera = mov["llave"].isin(fuera)
    return mov, mov[~es_fuera].copy(), mov[es_fuera].copy()


def _empresas_por_placa(ppto: pd.DataFrame, mov: pd.DataFrame) -> dict:
    a = pd.concat([ppto[["placa", "empresa"]], mov[["placa", "empresa"]]], ignore_index=True)
    a = a[a["placa"].fillna("") != ""]
    return a.groupby("placa")["empresa"].agg(lambda s: "/".join(sorted(set(s)))).to_dict() if len(a) else {}


# ================================================================== piezas de pantalla
def _tono(est: str):
    return {"Sobre ejecutado": (E.ROJO, E.ROJO_SUAVE), "Sin presupuesto asignado": (E.ROJO, E.ROJO_SUAVE),
            "Presupuesto agotado": (E.AMBAR, E.AMBAR_SUAVE), "Pendiente de clasificación": (E.AMBAR, E.AMBAR_SUAVE),
            "Dentro del presupuesto": (E.VERDE, E.VERDE_SUAVE)}.get(est, (E.AZUL, E.AZUL_SUAVE))


def _kpis(presupuesto: float, ejecutado: float, movimientos: int, hasta: str, sub_pres="Año completo"):
    t = L.con_metricas(pd.DataFrame({"presupuesto": [presupuesto], "ejecutado": [ejecutado],
                                     "movimientos": [movimientos]})).iloc[0]
    color, fondo = _tono(t["estado"])
    st.html(E.kpis([
        dict(label="Presupuesto asignado", valor=E.millones(presupuesto), sub=sub_pres, icono="wallet",
             color=E.AZUL, fondo=E.AZUL_SUAVE),
        dict(label=f"Ejecutado {hasta}", valor=E.millones(ejecutado),
             sub=f"{movimientos:,} movimientos".replace(",", "."), icono="receipt", color=E.AZUL, fondo=E.AZUL_SUAVE),
        dict(label="Saldo disponible", valor=E.millones(t["saldo"]), sub="Presupuesto − ejecutado",
             icono="piggy" if t["saldo"] >= 0 else "alert", color=E.AZUL if t["saldo"] >= 0 else E.ROJO,
             fondo=E.AZUL_SUAVE if t["saldo"] >= 0 else E.ROJO_SUAVE),
        dict(label="% de ejecución", valor=f'<span style="color:{color}">{E.pct_txt(t["pct"])}</span>',
             sub=t["estado"], icono="alert" if t["estado"] in ("Sobre ejecutado", "Sin presupuesto asignado") else "check",
             color=color, fondo=fondo),
    ]))


def _config() -> dict:
    return {
        "empresa": "Compañía",
        "placa": "Placa", "numero_interno": "N° interno", "tipo": "Tipo de vehículo", "concepto": "Concepto",
        "tipo_concepto": "Ámbito", "ambito": "Ámbito", "grupo": "Grupo", "sub_rubro": "Sub-rubro",
        "placas": st.column_config.NumberColumn("N° de placas", format="%d"),
        "texto_original": st.column_config.TextColumn("Descripción original (presupuesto)", width="large"),
        "presupuesto": st.column_config.NumberColumn("Presupuesto asignado", format="accounting"),
        "ejecutado": st.column_config.NumberColumn("Ejecutado", format="accounting"),
        "saldo": st.column_config.NumberColumn("Saldo disponible", format="accounting"),
        "pct": st.column_config.NumberColumn("% ejecución", format="percent"),
        "estado": "Estado", "movimientos": st.column_config.NumberColumn("Movimientos", format="%d"),
    }


def _tabla(df: pd.DataFrame, alto=380):
    st.dataframe(df, hide_index=True, width="stretch", height=alto, column_config=_config())


def _tabla_movimientos(mov: pd.DataFrame, clave: str, archivo: str, vehiculo: bool = False):
    """Detalle de movimientos con búsqueda y descarga. Valores originales de SAP (signo SAP)."""
    cols = [c for c in ["empresa", "fecha", "mes", "tercero", "texto", "documento", "codigo_articulo", "cantidad", "cuenta_nombre",
                        "cc_nombre", "sub_rubro_sap", "valor", "llave"] if c in mov.columns]
    if vehiculo:
        cols = [c for c in ["ambito", "placa", "placa_detectada", "placa_estado", "concepto", "origen", "confianza",
                            "concepto_sugerido"] if c in mov.columns] + cols
    buscar = st.text_input("Buscar", placeholder="Compañía, proveedor, texto, documento, placa, concepto…",
                           label_visibility="collapsed", key=f"mt_buscar_{clave}")
    vista = mov[cols].copy()
    if "fecha" in vista:
        vista["fecha"] = pd.to_datetime(vista["fecha"], errors="coerce")
    if buscar:
        m = vista.astype(str).apply(lambda c: c.str.contains(buscar, case=False, regex=False)).any(axis=1)
        vista = vista[m]
    vista = vista.sort_values("fecha", ascending=False)
    st.html(f'<p class="nf-card-s">{len(vista):,} movimientos · total '.replace(",", ".")
            + f'<b style="color:{E.TEXTO}">{E.pesos(vista["valor"].sum())}</b></p>')
    st.dataframe(vista, hide_index=True, width="stretch", height=420, column_config={
        "empresa": "Compañía",
        "fecha": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"),
        "ambito": "Ámbito", "placa": "Placa (final)", "placa_detectada": "Placa detectada en el texto",
        "placa_estado": "Identificación de placa", "concepto": "Concepto",
        "origen": "Clasificado por", "confianza": "Confianza", "concepto_sugerido": "Sugerido (confianza baja)",
        "tercero": "Proveedor / tercero", "texto": st.column_config.TextColumn("Texto original", width="large"),
        "documento": "Documento", "cuenta_nombre": "Cuenta", "cc_nombre": "Centro de coste",
        "sub_rubro_sap": "Sub-rubro (SAP)", "valor": st.column_config.NumberColumn("Valor", format="accounting"),
        "codigo_articulo": "Referencia (SISPO)", "cantidad": st.column_config.NumberColumn("Cantidad", format="%g"),
        "llave": "ID del movimiento"})
    st.download_button("Descargar movimientos", lambda d=vista: a_excel(d), f"{archivo}.xlsx",
                       icon=":material/download:", key=f"mt_dl_{clave}")


def _auditoria(S, m: pd.DataFrame, clave: str, archivo: str):
    """Tabla plana de los movimientos de la pestaña (con los mismos filtros que las cifras), para auditar y descargar."""
    if not S.es_admin:                                       # el detalle de movimientos es solo del administrador
        return
    with st.expander("Auditoría de movimientos · tabla completa", icon=":material/fact_check:", expanded=False):
        if m.empty:
            st.info("No hay movimientos con estos filtros.")
            return
        st.caption("Estos son los movimientos que componen la ejecución mostrada arriba (misma compañía, placa y concepto). "
                   "Para cambiar su concepto o placa, o eliminarlos, usa «Reclasificación de movimientos» más abajo.")
        _tabla_movimientos(m, clave, archivo, vehiculo=True)


def _sin_ceros(t: pd.DataFrame) -> pd.DataFrame:
    """Quita los renglones con presupuesto = 0, ejecutado = 0 y sin movimientos. Un concepto con movimientos que netean 0
    (p. ej. un gasto y su reverso) se conserva: tiene detalle que explicar."""
    if t.empty:
        return t
    num = lambda c: pd.to_numeric(t[c], errors="coerce").fillna(0.0) if c in t else pd.Series(0.0, index=t.index)
    return t[(num("presupuesto").abs() > 0.5) | (num("ejecutado").abs() > 0.5) | (num("movimientos") > 0)]


def _barras_ppto_ejec(t: pd.DataFrame, nota: str, con_pendiente: bool = False):
    """Presupuesto vs ejecutado por concepto (los 12 mayores). Responde: ¿dónde se concentra el gasto y qué tanto se ha usado?
    «Pendiente» (movimientos sin concepto) se omite, salvo que la tabla de al lado también lo muestre (con_pendiente)."""
    t = t if con_pendiente else t[t["concepto"] != L.PENDIENTE]
    t = t.copy()
    for c in ("presupuesto", "ejecutado"):                   # con filtros que dejan la tabla vacía las columnas llegan como texto
        t[c] = pd.to_numeric(t[c], errors="coerce").fillna(0.0)
    t = t.assign(_m=t[["presupuesto", "ejecutado"]].max(axis=1))
    top = t[t["_m"] > 0].nlargest(12, "_m").iloc[::-1] if len(t) else t
    if top.empty:
        st.caption("Aún no hay presupuesto ni ejecución para graficar.")
        return
    nombres = [c[:48] for c in top["concepto"]]
    fig = go.Figure()
    fig.add_bar(y=nombres, x=top["presupuesto"] / 1e6, name="Presupuesto", orientation="h", marker_color="#c9d4e4",
                hovertemplate="%{y}: %{x:,.1f} M<extra>Presupuesto</extra>")
    fig.add_bar(y=nombres, x=top["ejecutado"] / 1e6, name="Ejecutado", orientation="h", marker_color=E.AZUL,
                hovertemplate="%{y}: %{x:,.1f} M<extra>Ejecutado</extra>")
    fig.update_layout(barmode="group", xaxis=dict(tickformat=",.0f", ticksuffix=" M"), yaxis=dict(type="category"))
    st.plotly_chart(estilo_grafico(fig, 420), width="stretch", config={"displayModeBar": False})
    st.caption(nota)


def _grafico_ts_mensual(S, presupuesto_mes: pd.DataFrame, ejecucion: pd.DataFrame):
    meses = list(range(S.m0, S.m1 + 1))
    pm = presupuesto_mes.groupby("mes")["valor"].sum().reindex(meses, fill_value=0) / 1e6
    em = (ejecucion.groupby("mes")["valor"].sum().reindex(meses, fill_value=0) / 1e6).astype(float)
    em[em.index > S.hasta_mes] = None                       # no se dibuja lo que aún no ha pasado
    x = [f"{MESES[m - 1]} {str(S.cfg['anio'])[2:]}" for m in meses]
    fig = go.Figure()
    fig.add_bar(x=x, y=pm, name="Presupuesto", marker_color="#c9d4e4", marker_line_width=0, hovertemplate="%{y:,.1f} M")
    fig.add_bar(x=x, y=em, name="Ejecutado", marker_color=E.AZUL, marker_line_width=0, hovertemplate="%{y:,.1f} M")
    fig.update_layout(barmode="group", bargap=.25, yaxis=dict(tickformat=",.0f", ticksuffix=" M", rangemode="tozero"))
    st.plotly_chart(estilo_grafico(fig, 300), width="stretch", config={"displayModeBar": False})


# ================================================================== árbol desplegable hasta el movimiento
_CSS_ARBOL = """<style>
.mt-t{font-size:.84rem;width:100%;overflow-x:auto;background:#fff;border:1px solid #e3e9f2;border-radius:12px;color:#1b2a41}
.mt-f{display:grid;grid-template-columns:minmax(340px,4.2fr) 1.1fr 1.1fr 1.1fr .85fr 1.45fr;gap:.6rem;
 padding:.7rem 1rem;align-items:center;border-bottom:1px solid #eef1f6;min-width:980px}
.mt-h{font-weight:500;color:#6b7a90;font-size:.78rem;padding:.65rem 1rem;border-bottom:1px solid #e3e9f2}
.mt-n{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
.mt-t details>summary{list-style:none;cursor:pointer}
.mt-t details>summary::-webkit-details-marker{display:none}
.mt-t details>summary:hover{background:#f6f8fc}
.mt-nom{font-weight:600}
.mt-b1 .mt-nom,.mt-b2 .mt-nom,.mt-b3 .mt-nom{font-weight:500}
.mt-nom::before{content:"";display:inline-block;width:1.1rem}
details>summary .mt-nom::before{content:"›";color:#8a9bb5;font-size:1.15rem;line-height:1;font-weight:400;
 vertical-align:-.08rem;transition:transform .12s}
details[open]>summary .mt-nom::before{transform:rotate(90deg)}
.mt-b1{background:#fafbfd}.mt-b2{background:#f6f8fb}.mt-b3{background:#f2f5f9}
.mt-mov{font-size:.78rem;color:#33415c}.mt-mov .mt-nom,.mt-pp .mt-nom{font-weight:400}
.mt-pp{font-size:.78rem;color:#33415c;font-style:italic}
.mt-orig{display:block;font-size:.7rem;color:#8a9bb5;font-weight:400;margin-top:.1rem;word-break:break-word}
.mt-tag{font-size:.72rem;font-weight:400;color:#7a8aa3;margin-left:.6rem}
.mt-est{font-size:.76rem;font-weight:500;white-space:nowrap}
.mt-est i{display:inline-block;width:.45rem;height:.45rem;border-radius:50%;margin-right:.4rem;vertical-align:.05rem}
.mt-var{font-weight:600}
.mt-tot{font-weight:700;background:#f4f7fc;border-bottom:none}
.mt-filtro{display:inline-block;margin:.2rem 0 .6rem;padding:.35rem .8rem;border-radius:8px;background:#eaf1fb;
 color:#1f3a5f;font-size:.82rem}
</style>"""


def _n(x) -> str:
    return "" if pd.isna(x) else f"{x:,.0f}".replace(",", ".")


def _p(x) -> str:
    return "—" if pd.isna(x) else f"{x * 100:.1f}%".replace(".", ",")


def _m(x, exacto: bool = False) -> str:
    """Pesos como en el resumen general: millones con un decimal («$13,6 M»); por debajo del millón, o con exacto=True, pesos
    completos. El valor exacto siempre está en el atributo title (se ve al pasar el cursor)."""
    if x is None or pd.isna(x):
        return ""
    v = float(x)
    if exacto or abs(v) < 1e6:
        txt = f"${abs(v):,.0f}".replace(",", ".")
    else:
        txt = f"${abs(v) / 1e6:,.1f} M".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("−" if v <= -0.5 else "") + txt


def _cel(x, exacto: bool = False, color: str = "", clase: str = "") -> str:
    """Una celda numérica alineada a la derecha, con el valor exacto en el title."""
    if x is None or pd.isna(x):
        return '<span class="mt-n"></span>'
    estilo = f' style="color:{color}"' if color else ""
    return f'<span class="mt-n {clase}" title="{_n(x)}"{estilo}>{_m(x, exacto)}</span>'


def _celdas(r) -> str:
    """Presupuesto · Ejecutado · Variación (ejecutado − presupuesto: verde si queda por debajo, rojo si se pasa) · % ejecución · Estado."""
    color, _fondo = _tono(r["estado"])
    var = float(r["ejecutado"]) - float(r["presupuesto"])
    c_var = E.ROJO if var > 0.5 else (E.VERDE if var < -0.5 else "")
    return (_cel(r["presupuesto"]) + _cel(r["ejecutado"]) + _cel(var, color=c_var, clase="mt-var")
            + f'<span class="mt-n">{_p(r["pct"])}</span>'
            + f'<span><span class="mt-est" style="color:{color}"><i style="background:{color}"></i>{escape(r["estado"])}</span></span>')


def _pl(n, singular: str, plural: str | None = None) -> str:
    """«1 placa» / «4 placas»."""
    n = int(n)
    return f"{n} {singular if n == 1 else (plural or singular + 's')}"


def _tags(*items) -> str:
    return "".join(f'<span class="mt-tag">{escape(str(t))}</span>' for t in items if t)


def _fila(nombre: str, celdas: str, nivel: int = 0, tags: str = "", orig: str = "", clase: str = "") -> str:
    """Un renglón del árbol. `nivel` indenta y da el tono de fondo (más hondo = más oscuro)."""
    sangria = f' style="padding-left:{1.5 * nivel:.1f}rem"' if nivel else ""
    extra = f'<span class="mt-orig">{escape(orig)}</span>' if orig else ""
    fondo = f" mt-b{min(nivel, 3)}" if nivel else ""
    return (f'<div class="mt-f{fondo} {clase}"><span class="mt-nom"{sangria}>{escape(nombre)}{tags}{extra}</span>{celdas}</div>')


def _bloque(cabeza: str, hijos: list[str], abierto: bool) -> str:
    if not hijos:
        return cabeza
    return f'<details{" open" if abierto else ""}><summary>{cabeza}</summary>{"".join(hijos)}</details>'


def _orig_distinto(orig, concepto) -> str:
    orig = str(orig or "")
    return orig if orig and L.norm_texto(orig) != L.norm_texto(concepto) else ""


def _filas_mov(mv: pd.DataFrame, nivel: int, multi: bool, ver_placa: bool = False) -> list[str]:
    """Renglones de movimientos (fecha · documento · proveedor, y debajo el texto original). El valor va en la columna
    «Ejecutado», así la suma de los renglones es la ejecución del padre."""
    out = []
    if mv is None or mv.empty:
        return out
    for r in mv.sort_values("fecha").to_dict("records"):
        f = pd.to_datetime(r.get("fecha"), errors="coerce")
        cod = str(r.get("codigo_articulo") or "")
        if cod and cod != "nan":                        # línea de SISPO: referencia, artículo y cantidad; debajo, kardex y bodega
            cant = r.get("cantidad")
            partes = ["" if pd.isna(f) else f.strftime("%d/%m/%Y"), cod, str(r.get("articulo") or "")[:70]]
            extra = [f"cant. {cant:g}"] if cant is not None and not pd.isna(cant) else []
            debajo = " · ".join(x for x in (str(r.get("documento") or ""), str(r.get("tercero") or "")) if x)
        else:
            partes = ["" if pd.isna(f) else f.strftime("%d/%m/%Y"), str(r.get("documento") or ""), str(r.get("tercero") or "")]
            extra, debajo = [], str(r.get("texto") or "")
        partes = [x for x in partes if x and x != "nan"]
        placa = r.get("placa") if ver_placa else ""
        etiquetas = _tags(r.get("empresa") if multi else "", f"placa {placa}" if placa and placa != L.SIN_PLACA else "",
                          *extra, r.get("origen"))
        celdas = '<span></span>' + _cel(r.get("valor"), exacto=True) + '<span></span><span></span><span></span>'
        out.append(_fila(" · ".join(partes) or "(sin datos)", celdas, nivel, etiquetas, debajo, "mt-mov"))
    return out


def _filas_ppto_mes(S, placa: str, concepto: str, nivel: int) -> list[str]:
    """Renglones del presupuesto mensual de Tecnomecánica/SOAT: lo que compone el presupuesto del renglón padre."""
    out = []
    for mes, valor in S.ppto_mes.get((placa, concepto), []):
        if not valor:                                       # los meses en cero no aportan al presupuesto
            continue
        celdas = _cel(valor, exacto=True) + '<span></span><span></span><span></span><span></span>'
        out.append(_fila(f"Presupuesto {MESES[mes - 1]} {S.cfg['anio']}", celdas, nivel, clase="mt-pp"))
    return out


def _encabezado_arbol(primera: str) -> str:
    """Encabezado en una sola palabra (el primer nivel); la ruta completa queda en el title y en la nota bajo la tabla."""
    corto = primera.split(" / ")[0]
    return (f'<div class="mt-f mt-h"><span title="{escape(primera)}">{escape(corto)}</span><span class="mt-n">Presupuesto</span>'
            '<span class="mt-n">Ejecutado</span><span class="mt-n">Variación</span><span class="mt-n">% ejecución</span>'
            '<span>Estado</span></div>')


def _fila_total(p: float, e: float, n: int) -> str:
    t = L.con_metricas(pd.DataFrame({"presupuesto": [p], "ejecutado": [e], "movimientos": [n]})).iloc[0]
    return f'<div class="mt-f mt-tot"><span class="mt-nom">Total</span>{_celdas(t)}</div>'


def _cierre(filas: list[str], p: float, e: float, n: int) -> str:
    return _CSS_ARBOL + '<div class="mt-t">' + "".join(filas) + _fila_total(p, e, n) + "</div>"


def _agrupar(mv: pd.DataFrame, claves: list[str]) -> dict:
    if mv.empty:
        return {}
    return {k if isinstance(k, tuple) else (k,): g for k, g in mv.groupby(claves, sort=False)}


def _limpiar_multi(key: str, opciones):
    """Quita de un multiselect las opciones que ya no existen (por ejemplo, al cambiar de escenario)."""
    if key in st.session_state:
        st.session_state[key] = [x for x in st.session_state[key] if x in set(opciones)]


def _totales(nodos: list[dict]):
    return (sum(n["presupuesto"] for n in nodos), sum(n["ejecutado"] for n in nodos), int(sum(n["movimientos"] for n in nodos)))


SIN_EJEC = "Otras placas sin ejecución"


def _sin_ejec(n) -> bool:
    """Placa con presupuesto pero sin un solo movimiento ni valor ejecutado (la «SIN PLACA IDENTIFICADA» nunca se agrupa)."""
    return n["placa"] != L.SIN_PLACA and abs(float(n["ejecutado"])) < 0.5 and int(n["movimientos"]) == 0


def _bloque_otras(S, otras: list, nivel: int, hijos: list[str]) -> str:
    """Una sola fila «Otras placas sin ejecución» con la suma de sus presupuestos; se abre en las placas."""
    t = L.con_metricas(pd.DataFrame({"presupuesto": [float(sum(n["presupuesto"] for n in otras))], "ejecutado": [0.0],
                                     "movimientos": [0]})).iloc[0]
    abierto = bool(getattr(S, "abrir_otras", False))
    return _bloque(_fila(SIN_EJEC, _celdas(t), nivel, _tags(_pl(len(otras), "placa"))), hijos, abierto)


def _bloques_placa(S, nodos, grupos: dict, embebe: bool, nivel0: int, abrir: set, con_ppto_mes: bool, con_tipo: bool = True,
                   agrupar: bool = True) -> list[str]:
    """Bloques placa -> concepto -> movimientos (y, en Tecnomecánica/SOAT, el presupuesto mensual de cada concepto).
    `nivel0` es la sangría de la placa: 0 si es el primer nivel, 1 si cuelga de un tipo o de un grupo.
    Las placas sin ejecución se reúnen al final en «Otras placas sin ejecución» (con el total de su presupuesto)."""
    if agrupar:
        otras = [n for n in nodos if _sin_ejec(n)]
        activas = [n for n in nodos if not _sin_ejec(n)]
        if otras and activas:          # si TODAS las placas están sin ejecución no hace falta un nivel extra: ya están bajo su concepto
            hijos_o = _bloques_placa(S, otras, grupos, embebe, nivel0 + 1, abrir, con_ppto_mes, con_tipo, agrupar=False)
            return (_bloques_placa(S, activas, grupos, embebe, nivel0, abrir, con_ppto_mes, con_tipo, agrupar=False)
                    + [_bloque_otras(S, otras, nivel0, hijos_o)])
    bloques = []
    for n in nodos:
        h = n["hijos"]
        hijos = []
        for r in h.to_dict("records"):
            nietos = _filas_ppto_mes(S, n["placa"], r["concepto"], nivel0 + 2) if con_ppto_mes else []
            if embebe:
                nietos += _filas_mov(grupos.get((n["placa"], r["concepto"])), nivel0 + 2, S.multi)
            cab = _fila(r["concepto"], _celdas(r), nivel0 + 1, "", _orig_distinto(r.get("texto_original"), r["concepto"]))
            hijos.append(_bloque(cab, nietos, (nivel0 + 1) in abrir))
        interno = str(n.get("numero_interno") or "")
        etiqueta = _tags(str(n.get("tipo") or "")[:36] if con_tipo else "", f"N° interno {interno}" if interno and not con_tipo else "",
                         S.emp_placa.get(n["placa"]) if S.multi else "", _pl(len(h), "concepto") if len(h) else "")
        bloques.append(_bloque(_fila(n["placa"], _celdas(n), nivel0, etiqueta), hijos, nivel0 in abrir))
    return bloques


def _html_placa_concepto(S, nodos, mov_amb: pd.DataFrame, abrir: set, con_ppto_mes: bool, titulo: str) -> str:
    """Árbol placa -> concepto -> movimientos."""
    grupos = _agrupar(mov_amb, ["placa", "concepto"])
    filas = [_encabezado_arbol(titulo)] + _bloques_placa(S, nodos, grupos, len(mov_amb) <= MAX_MOV_ARBOL, 0, abrir, con_ppto_mes)
    tp, te, tm = _totales(nodos)
    return _cierre(filas, tp, te, tm)


SIN_CONCEPTO = "Sin concepto presupuestal"


def _categoria_placa(S, n: dict) -> str:
    """Concepto presupuestal de una placa: la descripción (columna TIPO VEHICULO) con que viene en el archivo de presupuesto.
    Una placa que no está en el presupuesto (o la «SIN PLACA IDENTIFICADA») queda en «Sin concepto presupuestal»."""
    if n["placa"] == L.SIN_PLACA:
        return SIN_CONCEPTO
    return S.tipo_ppto.get(n["placa"]) or SIN_CONCEPTO


def _categorias(S, nodos: list[dict], generales: pd.DataFrame | None) -> list[dict]:
    """Conceptos presupuestales de Flota y Equipo de transporte, en orden alfabético («Sin concepto presupuestal» al final).
    Los de vehículo traen sus placas (`nodos`); los generales (sin placa) no tienen placas: sus movimientos cuelgan de ellos."""
    por: dict[str, list] = {}
    for n in nodos:
        por.setdefault(_categoria_placa(S, n), []).append(n)
    cats = [dict(nombre=k, nodos=v, general=False, placas=sum(1 for n in v if n["placa"] != L.SIN_PLACA),
                 presupuesto=float(sum(n["presupuesto"] for n in v)), ejecutado=float(sum(n["ejecutado"] for n in v)),
                 movimientos=int(sum(n["movimientos"] for n in v)), texto_original="") for k, v in por.items()]
    if generales is not None and len(generales):
        cats += [dict(nombre=r["concepto"], nodos=[], general=True, placas=0, presupuesto=float(r["presupuesto"]),
                      ejecutado=float(r["ejecutado"]), movimientos=int(r["movimientos"]), texto_original=str(r.get("texto_original") or ""))
                 for r in generales.to_dict("records")]
    return sorted(cats, key=lambda c: (c["nombre"] == SIN_CONCEPTO, c["nombre"]))


def _df_categorias(cats: list[dict], total: bool = True) -> pd.DataFrame:
    t = pd.DataFrame([dict(categoria=c["nombre"], placas=c["placas"], presupuesto=c["presupuesto"], ejecutado=c["ejecutado"],
                           movimientos=c["movimientos"]) for c in cats],
                     columns=["categoria", "placas", "presupuesto", "ejecutado", "movimientos"])
    if t.empty:
        return t.astype({"placas": int, "presupuesto": float, "ejecutado": float, "movimientos": int}).assign(
            saldo=pd.Series(dtype=float), pct=pd.Series(dtype=float), estado=pd.Series(dtype=str))
    if total:
        t.loc[len(t)] = dict(categoria="Total", placas=int(t["placas"].sum()), presupuesto=t["presupuesto"].sum(),
                             ejecutado=t["ejecutado"].sum(), movimientos=int(t["movimientos"].sum()))
    return L.con_metricas(t)


def _resumen_grupos(nodos: list[dict], clave, total: bool = False) -> pd.DataFrame:
    """Una fila por grupo de placas (tipo de vehículo, con/sin presupuesto…): placas, presupuesto, ejecutado, saldo, % y estado.
    Es la MISMA suma de los nodos del árbol, así la tabla resumen y el árbol siempre coinciden."""
    filas: dict = {}
    for n in nodos:
        g = filas.setdefault(clave(n), dict(placas=0, presupuesto=0.0, ejecutado=0.0, movimientos=0))
        g["placas"] += 0 if n["placa"] == L.SIN_PLACA else 1
        g["presupuesto"] += float(n["presupuesto"])
        g["ejecutado"] += float(n["ejecutado"])
        g["movimientos"] += int(n["movimientos"])
    t = pd.DataFrame([dict(grupo=k, **v) for k, v in filas.items()], columns=["grupo", "placas", "presupuesto", "ejecutado", "movimientos"])
    if t.empty:
        return t.astype({"placas": int, "presupuesto": float, "ejecutado": float, "movimientos": int}).assign(
            saldo=pd.Series(dtype=float), pct=pd.Series(dtype=float), estado=pd.Series(dtype=str))
    if total:
        t.loc[len(t)] = dict(grupo="Total", placas=int(t["placas"].sum()), presupuesto=t["presupuesto"].sum(),
                             ejecutado=t["ejecutado"].sum(), movimientos=int(t["movimientos"].sum()))
    return L.con_metricas(t)


def _html_grupos_ts(S, bloques: list[tuple], mov_amb: pd.DataFrame, abrir: set, titulo: str, concepto: str) -> str:
    """Árbol de UN concepto (Tecnomecánica o SOAT): grupo («Placas con presupuesto» / «Otras placas») -> placa -> presupuesto
    mensual y movimientos. `bloques` = [(nombre del grupo, nodos)]; los grupos vacíos no se dibujan."""
    grupos = _agrupar(mov_amb, ["placa", "concepto"])
    embebe = len(mov_amb) <= MAX_MOV_ARBOL
    filas, todos = [_encabezado_arbol(titulo)], []
    for nombre, nodos in bloques:
        if not nodos:
            continue
        todos += nodos
        tp, te, tm = _totales(nodos)
        t = L.con_metricas(pd.DataFrame({"presupuesto": [tp], "ejecutado": [te], "movimientos": [tm]})).iloc[0]
        hijos = []
        for n in nodos:
            nietos = _filas_ppto_mes(S, n["placa"], concepto, 2)
            if embebe:
                nietos += _filas_mov(grupos.get((n["placa"], concepto)), 2, S.multi)
            etiqueta = _tags(str(n.get("tipo") or "")[:36], S.emp_placa.get(n["placa"]) if S.multi else "")
            hijos.append(_bloque(_fila(n["placa"], _celdas(n), 1, etiqueta), nietos, 1 in abrir))
        filas.append(_bloque(_fila(nombre, _celdas(t), 0, _tags(_pl(len(nodos), "placa"))), hijos, 0 in abrir))
    tp, te, tm = _totales(todos)
    return _cierre(filas, tp, te, tm)


def _html_concepto(S, filas_c: pd.DataFrame, mov_amb: pd.DataFrame, abrir: set, titulo: str) -> str:
    """Árbol concepto de gasto -> movimientos (sin nivel de placa: la placa de cada movimiento va como etiqueta). El valor de cada
    movimiento va en «Ejecutado», así que la suma de los renglones es la ejecución del concepto."""
    grupos = _agrupar(mov_amb, ["concepto"])
    embebe = len(mov_amb) <= MAX_MOV_ARBOL
    filas = [_encabezado_arbol(titulo)]
    for c in filas_c.to_dict("records"):
        hijos = _filas_mov(grupos.get((c["concepto"],)), 1, S.multi, ver_placa=True) if embebe else []
        etiqueta = _tags(_pl(c["movimientos"], "movimiento") if c["movimientos"] else "")
        filas.append(_bloque(_fila(c["concepto"], _celdas(c), 0, etiqueta, _orig_distinto(c.get("texto_original"), c["concepto"])),
                             hijos, 0 in abrir))
    return _cierre(filas, filas_c["presupuesto"].sum(), filas_c["ejecutado"].sum(), int(filas_c["movimientos"].sum()))


def _html_generales(S, g: pd.DataFrame, mov_gen: pd.DataFrame, abrir: set) -> str:
    """Árbol concepto general -> movimientos."""
    grupos = _agrupar(mov_gen, ["concepto"])
    embebe = len(mov_gen) <= MAX_MOV_ARBOL
    filas = [_encabezado_arbol("Concepto general / movimiento")]
    for r in g.to_dict("records"):
        hijos = _filas_mov(grupos.get((r["concepto"],)), 1, S.multi, ver_placa=True) if embebe else []
        etiqueta = _tags(_pl(r["movimientos"], "movimiento") if r["movimientos"] else "")
        filas.append(_bloque(_fila(r["concepto"], _celdas(r), 0, etiqueta, str(r.get("texto_original") or "")), hijos, 0 in abrir))
    return _cierre(filas, g["presupuesto"].sum(), g["ejecutado"].sum(), int(g["movimientos"].sum()))


# ================================================================== reasignar y eliminar movimientos
def _insertar_por_lotes(ctx, sql_cols: str, filas: list[tuple], lote=100):
    """INSERT multi-fila en lotes (menos viajes a la base que uno por fila; lote × columnas < 999 variables)."""
    if not filas:                                       # nada que insertar (p. ej. una cuenta aplicada sin movimientos de SAP que reemplazar)
        return
    n = len(filas[0])
    for i in range(0, len(filas), lote):
        parte = filas[i:i + lote]
        ctx.con.execute(f"INSERT INTO {sql_cols} VALUES " + ",".join(["(" + ",".join("?" * n) + ")"] * len(parte)),
                        tuple(v for f in parte for v in f))


def _exigir_admin(ctx):
    """Candado de permisos en la LÓGICA (no solo en la pantalla): toda escritura en la base y toda exportación con detalle de
    movimientos exige administrador. Si un usuario normal llegara a ejecutar la función por otra ruta, se detiene aquí."""
    if not getattr(ctx, "es_admin", False):
        raise PermissionError("Acción reservada a administradores.")


def _sin_prefijo(empresa: str, llave: str) -> str:
    return llave[len(empresa) + 1:] if llave.startswith(empresa + "|") else llave


def _registrar_log(ctx, log: list[dict], ahora_txt: str):
    if log:
        _insertar_por_lotes(ctx, "mtto_clasif_log(empresa, fecha, usuario, llave, texto, valor, concepto_antes, "
                                 "concepto_despues, placa_antes, placa_despues, accion, motivo)",
                            [(g["empresa"], ahora_txt, ctx.usuario, g["llave"], g["texto"], float(g["valor"]), g["concepto_antes"],
                              g["concepto_despues"], g["placa_antes"], g["placa_despues"], g["accion"], g.get("motivo", ""))
                             for g in log], lote=70)


def _guardar_overrides(ctx, filas: list[dict], quitar: list[tuple], log: list[dict] | None = None):
    """Guarda reasignaciones manuales y deja el rastro. Cada fila: empresa, llave, concepto ('' = el automático), placa
    ('' = la detectada), texto_tipo (para aprender), tercero. `quitar`: lista de (empresa, llave) que vuelven a automático.
    Una reasignación nueva REEMPLAZA la anterior del mismo movimiento (una sola fila por movimiento): no se duplica nada."""
    _exigir_admin(ctx)
    por_empresa: dict[str, set] = {}
    for f in filas:
        por_empresa.setdefault(f["empresa"], set()).add(f["llave"])
    for e, ll in quitar:
        por_empresa.setdefault(e, set()).add(ll)
    for e, llaves in por_empresa.items():
        todas = sorted(llaves | {_sin_prefijo(e, x) for x in llaves})      # también la forma antigua, sin compañía
        for i in range(0, len(todas), 200):
            parte = todas[i:i + 200]
            ctx.con.execute("DELETE FROM mtto_clasif_manual WHERE empresa=? AND origen='Manual' AND llave IN ("
                            + ",".join("?" * len(parte)) + ")", (e, *parte))
    ahora_txt = ahora().strftime("%Y-%m-%d %H:%M")
    if filas:
        _insertar_por_lotes(ctx, "mtto_clasif_manual(empresa, llave, concepto, texto_tipo, tercero, origen, usuario, fecha, placa)",
                            [(f["empresa"], f["llave"], f["concepto"], f["texto_tipo"], L.norm_texto(f["tercero"]), "Manual",
                              ctx.usuario, ahora_txt, f["placa"]) for f in filas], lote=100)
    _registrar_log(ctx, log or [], ahora_txt)
    ctx.con.commit()
    ctx.registrar_carga(ctx.usuario, "Reasignación mantenimiento", "pantalla",
                        f"{len(filas)} reasignados, {len(quitar)} devueltos a automático")


def _excluir(ctx, filas: list[dict], motivo: str):
    """Elimina movimientos de la vista SIN borrarlos de la ejecución: quedan en mtto_excluidos (con su texto, valor y
    motivo) y en el historial. Como se excluyen por su llave, no reaparecen cuando se vuelva a cargar la ejecución."""
    _exigir_admin(ctx)
    ahora_txt = ahora().strftime("%Y-%m-%d %H:%M")
    por_empresa: dict[str, list] = {}
    for f in filas:
        por_empresa.setdefault(f["empresa"], []).append(f["llave"])
    for e, llaves in por_empresa.items():
        for i in range(0, len(llaves), 200):
            parte = llaves[i:i + 200]
            ctx.con.execute("DELETE FROM mtto_excluidos WHERE empresa=? AND llave IN (" + ",".join("?" * len(parte)) + ")",
                            (e, *parte))
    _insertar_por_lotes(ctx, "mtto_excluidos(empresa, llave, fecha_mov, documento, tercero, texto, valor, concepto, placa, "
                             "motivo, usuario, fecha)",
                        [(f["empresa"], f["llave"], str(f["fecha"])[:10], str(f["documento"] or ""), str(f["tercero"] or ""),
                          str(f["texto"] or "")[:500], float(f["valor"]), f["concepto"], f["placa"], motivo, ctx.usuario, ahora_txt)
                         for f in filas], lote=60)
    _registrar_log(ctx, [dict(empresa=f["empresa"], llave=f["llave"], texto=str(f["texto"] or "")[:300], valor=f["valor"],
                              concepto_antes=f["concepto"], concepto_despues="(eliminado)", placa_antes=f["placa"],
                              placa_despues="(eliminado)", accion="Eliminada", motivo=motivo) for f in filas], ahora_txt)
    ctx.con.commit()
    ctx.registrar_carga(ctx.usuario, "Eliminación movimientos mantenimiento", "pantalla",
                        f"{len(filas)} movimientos por {E.pesos(sum(float(f['valor']) for f in filas))}. {motivo}"[:250])


def _restaurar(ctx, filas: list[dict]):
    _exigir_admin(ctx)
    ahora_txt = ahora().strftime("%Y-%m-%d %H:%M")
    por_empresa: dict[str, list] = {}
    for f in filas:
        por_empresa.setdefault(f["empresa"], []).append(f["llave"])
    for e, llaves in por_empresa.items():
        for i in range(0, len(llaves), 200):
            parte = llaves[i:i + 200]
            ctx.con.execute("DELETE FROM mtto_excluidos WHERE empresa=? AND llave IN (" + ",".join("?" * len(parte)) + ")",
                            (e, *parte))
    _registrar_log(ctx, [dict(empresa=f["empresa"], llave=f["llave"], texto=str(f["texto"] or "")[:300], valor=f["valor"],
                              concepto_antes="(eliminado)", concepto_despues=f["concepto"], placa_antes="(eliminado)",
                              placa_despues=f["placa"], accion="Restaurada", motivo="") for f in filas], ahora_txt)
    ctx.con.commit()
    ctx.registrar_carga(ctx.usuario, "Restauración movimientos mantenimiento", "pantalla", f"{len(filas)} movimientos")


def _panel_eliminar(S):
    """Confirmación de la eliminación: muestra qué se va a eliminar, pide el motivo y solo entonces actúa."""
    pend = st.session_state.get("mt_del_pend")
    if not pend:
        return
    d = pd.DataFrame(pend)
    st.warning(f"Vas a eliminar **{len(d)}** movimientos por **{E.pesos(d['valor'].sum())}**. Dejarán de contar en el "
               "presupuesto vs ejecución y en el Excel, y no reaparecerán al cargar la ejecución de nuevo. La ejecución cargada no "
               "se borra: puedes restaurarlos en «Estado → Eliminados». Queda registro con tu usuario, fecha y motivo.",
               icon=":material/delete:")
    st.dataframe(d[["empresa", "fecha", "documento", "tercero", "texto", "valor"]], hide_index=True, width="stretch", height=200,
                 column_config={"empresa": "Compañía", "texto": st.column_config.TextColumn("Texto original", width="large"),
                                "valor": st.column_config.NumberColumn("Valor", format="accounting")})
    motivo = st.text_input("Motivo de la eliminación", placeholder="Ej.: movimiento duplicado, no corresponde a flota…",
                           key="mt_del_motivo")
    b1, b2 = st.columns(2)
    if b1.button("Sí, eliminar estos movimientos", icon=":material/delete:", key="mt_del_ok"):
        _excluir(S.ctx, pend, (motivo or "").strip())
        st.session_state.pop("mt_del_pend", None)
        st.success(f"{len(pend)} movimientos eliminados.")
        st.rerun()
    if b2.button("Cancelar", key="mt_del_cancel"):
        st.session_state.pop("mt_del_pend", None)
        st.rerun()


def _vista_eliminados(S, f_emp: str, buscar: str):
    ex = S.excluidos
    if f_emp != "Todas":
        ex = ex[ex["empresa"] == f_emp]
    if buscar:
        ex = ex[ex.astype(str).apply(lambda c: c.str.contains(buscar, case=False, regex=False)).any(axis=1)]
    if ex.empty:
        st.info("No hay movimientos eliminados con estos filtros.")
        return
    st.caption(f"{len(ex)} movimientos eliminados por {E.pesos(ex['valor'].sum())}. Siguen en la ejecución cargada, pero no cuentan.")
    t = ex[["empresa", "fecha_mov", "documento", "tercero", "texto", "valor", "concepto", "placa", "motivo", "usuario", "fecha",
            "llave"]].reset_index(drop=True)
    t["restaurar"] = False
    cfg = {"empresa": "Compañía", "fecha_mov": "Fecha movimiento", "texto": st.column_config.TextColumn("Texto original", width="large"),
           "valor": st.column_config.NumberColumn("Valor", format="accounting"), "concepto": "Concepto (antes)",
           "placa": "Placa (antes)", "usuario": "Eliminado por", "fecha": "Fecha eliminación", "llave": "ID del movimiento",
           "restaurar": st.column_config.CheckboxColumn("Restaurar")}
    if not S.es_admin:
        st.dataframe(t.drop(columns="restaurar"), hide_index=True, width="stretch", height=360, column_config=cfg)
        return
    ed = st.data_editor(t, hide_index=True, width="stretch", height=360, key=f"mt_rc_ex_{f_emp}_{buscar}", column_config=cfg,
                        disabled=[c for c in t.columns if c != "restaurar"])
    marcados = ed[ed["restaurar"]]
    if st.button(f"Restaurar {len(marcados)} movimientos", icon=":material/restore:", key="mt_rc_restaurar",
                 disabled=marcados.empty):
        _restaurar(S.ctx, [dict(empresa=r.empresa, llave=r.llave, texto=r.texto, valor=r.valor, concepto=r.concepto, placa=r.placa)
                           for r in marcados.itertuples()])
        st.success(f"{len(marcados)} movimientos restaurados.")
        st.rerun()


def _reclasificacion(S):
    if not S.es_admin:
        return
    """TODOS los movimientos de ambas compañías (clasificados o pendientes), con filtros y paginación. El administrador
    cambia concepto y placa, o elimina con confirmación. Nada se borra de la ejecución cargada."""
    with st.expander("Reclasificación de movimientos · todas las compañías", icon=":material/edit_note:",
                     expanded=bool(st.session_state.get("mt_del_pend"))):
        st.caption(f"Todos los movimientos de {' y '.join(EMPRESAS)} del escenario **{S.escenario}** {S.hasta}, clasificados o pendientes. "
                   "Cada uno muestra su compañía, placa, concepto, texto original y clasificación (Placas / Generales / "
                   "Tecnomecánica / SOAT). "
                   + ("Cambia el **concepto** (eso define la clasificación) y/o la **placa**, o marca **Eliminar**. El texto original "
                      f"nunca se modifica. «{AUTO}» y «{AUTO_PLACA}» deshacen un cambio." if S.es_admin
                      else "Solo el administrador puede modificarlos."))
        _panel_eliminar(S)
        c1, c2, c3 = st.columns([1, 1, 1.3])
        f_emp = c1.selectbox("Compañía", ["Todas", *EMPRESAS], key="mt_rc_emp")
        f_grupo = c2.selectbox("Clasificación", ["Todas", "Placas", "Generales", "Tecnomecánica", "SOAT", "Pendiente"], key="mt_rc_grupo")
        f_est = c3.selectbox("Estado", ["Todos", "Pendientes de clasificar", "Reclasificados a mano", "Con placa a mano", "Eliminados"],
                             key="mt_rc_estado")
        c4, c5 = st.columns([1, 2])
        placas_op = ["Todas", *S.placas_edit, L.SIN_PLACA]
        f_placa = c4.selectbox("Placa", placas_op, key="mt_rc_placa")
        presentes = sorted(set(S.mov_todos["concepto"]) | set(S.catalogo))
        f_conc = c5.selectbox("Concepto", ["Todos", *presentes], key="mt_rc_concepto")
        buscar = st.text_input("Buscar", placeholder="Compañía, placa, concepto, proveedor, descripción, documento…",
                               key="mt_rc_buscar", label_visibility="collapsed")
        limitar = bool(S.placa) and st.checkbox(f"Limitar a la placa {S.placa} (filtro global de arriba)", value=True, key="mt_rc_global")
        if f_est == "Eliminados":
            _vista_eliminados(S, f_emp, buscar)
            return
        v = S.mov_todos
        if f_emp != "Todas":
            v = v[v["empresa"] == f_emp]
        if f_grupo != "Todas":
            v = v[v["grupo"] == f_grupo]
        if f_placa != "Todas":
            v = v[v["placa"] == f_placa]
        if limitar:
            v = v[v["placa"] == S.placa]
        if f_conc != "Todos":
            v = v[v["concepto"] == f_conc]
        if f_est == "Pendientes de clasificar":
            v = v[v["concepto"] == L.PENDIENTE]
        elif f_est == "Reclasificados a mano":
            v = v[v["origen"] == "Manual"]
        elif f_est == "Con placa a mano":
            v = v[v["placa_estado"] == "Asignada a mano"]
        if buscar:
            v = v[v[["empresa", "placa", "concepto", "tercero", "texto", "documento"]].astype(str).apply(
                lambda c: c.str.contains(buscar, case=False, regex=False)).any(axis=1)]
        v = v.sort_values(["fecha", "valor"], ascending=False)
        n = len(v)
        paginas = max(1, math.ceil(n / PAGINA))
        try:                                            # el valor guardado puede llegar como texto (p. ej. "2"); se normaliza a entero
            pag_ant = int(st.session_state.get("mt_rc_pag", 1))
        except (TypeError, ValueError):
            pag_ant = 1
        st.session_state["mt_rc_pag"] = pag_ant if 1 <= pag_ant <= paginas else 1
        pag = 1 if paginas == 1 else int(st.selectbox("Página", list(range(1, paginas + 1)), key="mt_rc_pag",
                                                      format_func=lambda i: f"Página {i} de {paginas}"))
        st.html(f'<p class="nf-card-s">{n:,} movimientos · total '.replace(",", ".")
                + f'<b style="color:{E.TEXTO}">{E.pesos(v["valor"].sum())}</b>'
                + (f" · mostrando {PAGINA} por página" if paginas > 1 else "") + "</p>")
        if v.empty:
            st.info("No hay movimientos con estos filtros.")
            return
        page = v.iloc[(pag - 1) * PAGINA: pag * PAGINA]
        base = page[["llave", "empresa", "fecha", "documento", "tercero", "texto", "valor", "grupo", "concepto", "placa",
                     "placa_detectada", "origen"]].reset_index(drop=True)
        base["fecha"] = pd.to_datetime(base["fecha"], errors="coerce")
        cfg_cols = {"empresa": "Compañía", "fecha": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"),
                    "texto": st.column_config.TextColumn("Texto original", width="large"), "tercero": "Proveedor",
                    "documento": "Documento", "valor": st.column_config.NumberColumn("Valor", format="accounting"),
                    "grupo": "Clasificación", "origen": "Clasificado por", "llave": "ID del movimiento",
                    "placa_detectada": "Placa detectada en el texto"}
        if not S.es_admin:
            st.dataframe(base, hide_index=True, width="stretch", height=420, column_config=cfg_cols,
                         column_order=["empresa", "fecha", "documento", "tercero", "texto", "valor", "grupo", "concepto", "placa",
                                       "origen", "llave"])
            return
        op_placa = [AUTO_PLACA, *sorted(set(S.placas_edit) | set(base["placa"]) | {L.SIN_PLACA})]
        op_conc = list(dict.fromkeys([AUTO, L.PENDIENTE, *S.catalogo, *base["concepto"]]))
        ed_base = base.assign(eliminar=False)
        ed = st.data_editor(
            ed_base, hide_index=True, width="stretch", height=440, key=f"mt_rc_ed_{f_emp}_{f_grupo}_{f_est}_{f_placa}_{f_conc}_{buscar}_{pag}",
            disabled=["llave", "empresa", "fecha", "documento", "tercero", "texto", "valor", "grupo", "placa_detectada", "origen"],
            column_order=["empresa", "fecha", "documento", "tercero", "texto", "valor", "grupo", "concepto", "placa", "origen",
                          "eliminar"],
            column_config={**cfg_cols,
                           "concepto": st.column_config.SelectboxColumn("Concepto asignado", options=op_conc, width="large"),
                           "placa": st.column_config.SelectboxColumn("Placa", options=op_placa),
                           "eliminar": st.column_config.CheckboxColumn("Eliminar")})
        cambios = ed_base.index[(ed["concepto"] != ed_base["concepto"]) | (ed["placa"] != ed_base["placa"])]
        marcados = ed_base.index[ed["eliminar"]]
        aprender = st.checkbox("Aprender de estos cambios: aplicar el mismo concepto a otros movimientos con el mismo texto y proveedor",
                               value=True, key="mt_rc_aprender",
                               help="Desmárcalo si el cambio es un caso puntual (por ejemplo, un servicio recurrente que esta vez "
                                    "corresponde a otra placa). La reasignación de este movimiento se guarda igual.")
        b1, b2 = st.columns([1, 1])
        guardar = b1.button(f"Guardar {len(cambios)} cambios", icon=":material/save:", key="mt_rc_guardar", disabled=len(cambios) == 0)
        eliminar = b2.button(f"Eliminar {len(marcados)} seleccionados…", icon=":material/delete:", key="mt_rc_eliminar",
                             disabled=len(marcados) == 0)
        if eliminar:
            st.session_state["mt_del_pend"] = [
                dict(empresa=r.empresa, llave=r.llave, fecha=str(r.fecha)[:10], documento=r.documento, tercero=r.tercero,
                     texto=r.texto, valor=float(r.valor), concepto=r.concepto, placa=r.placa)
                for r in ed_base.loc[marcados].itertuples()]
            st.rerun()
        if guardar:
            filas, quitar, log, ignorados = [], [], [], 0
            for i in cambios:
                b, e = ed_base.loc[i], ed.loc[i]
                c_prev, p_prev = S.manual_prev.get(b["llave"], ("", ""))
                c_new, p_new = c_prev, p_prev
                if e["concepto"] != b["concepto"]:
                    if e["concepto"] == AUTO:
                        c_new = ""
                    elif e["concepto"] in S.catalogo:
                        c_new = e["concepto"]
                    else:                                         # «pendiente» no se puede asignar a mano
                        ignorados += 1
                        continue
                if e["placa"] != b["placa"]:
                    p_new = "" if e["placa"] == AUTO_PLACA else e["placa"]
                if c_new == "" and p_new == "":
                    quitar.append((b["empresa"], b["llave"]))
                else:
                    filas.append(dict(empresa=b["empresa"], llave=b["llave"], concepto=c_new, placa=p_new,
                                      texto_tipo=L.firma_texto(b["texto"], b["placa_detectada"]) if (c_new and aprender) else "",
                                      tercero=b["tercero"] if pd.notna(b["tercero"]) else ""))
                log.append(dict(empresa=b["empresa"], llave=b["llave"], texto=str(b["texto"])[:300], valor=b["valor"],
                                concepto_antes=b["concepto"], concepto_despues=c_new or "(automático)",
                                placa_antes=b["placa"], placa_despues=p_new or "(detectada en el texto)",
                                accion="Revertida" if (c_new == "" and p_new == "") else "Reasignada"))
            _guardar_overrides(S.ctx, filas, quitar, log)
            if ignorados:
                st.warning(f"{ignorados} cambios a «{L.PENDIENTE}» se ignoraron; usa «{AUTO}» para deshacer.")
            st.success(f"{len(filas)} movimientos reasignados, {len(quitar)} devueltos a automático.")
            st.rerun()


def _reclasificar_grupo(S):
    """Pendientes agrupados por texto y proveedor (de ambas compañías): un solo clic reasigna el grupo y deja el aprendizaje."""
    if not S.es_admin:
        st.info("Solo el administrador puede reclasificar.")
        return
    pend = S.mov_todos[S.mov_todos["concepto"] == L.PENDIENTE].copy()
    if pend.empty:
        st.success("No hay movimientos pendientes de clasificación.")
        return
    pend["firma"] = [L.firma_texto(t, p) for t, p in zip(pend["texto"], pend["placa_detectada"])]
    pend["ter"] = pend["tercero"].fillna("").map(L.norm_texto)
    g = pend.groupby(["firma", "ter"]).agg(n=("valor", "size"), valor=("valor", "sum")).reset_index() \
        .sort_values("valor", ascending=False).head(200)
    etiquetas = [f"{r.firma[:60] or '(sin texto)'} · {r.ter[:26] or 'sin proveedor'} · {r.n} mov. · {_n(r.valor)}"
                 for r in g.itertuples()]
    c1, c2 = st.columns([2.2, 1.4], vertical_alignment="bottom")
    i = etiquetas.index(c1.selectbox("Grupo", etiquetas, key="mt_rec_grupo"))
    concepto = c2.selectbox("Asignar al concepto", S.catalogo, key="mt_rec_concepto_g")
    fila = g.iloc[i]
    sel = pend[(pend["firma"] == fila["firma"]) & (pend["ter"] == fila["ter"])]
    st.dataframe(sel[["empresa", "fecha", "placa", "tercero", "texto", "valor"]].head(8), hide_index=True, width="stretch",
                 column_config={"empresa": "Compañía", "valor": st.column_config.NumberColumn("Valor", format="accounting"),
                                "texto": st.column_config.TextColumn("Texto original", width="large")})
    if st.button(f"Aplicar a los {len(sel)} movimientos", icon=":material/rule:", key="mt_rec_aplicar_g"):
        filas = [dict(empresa=r.empresa, llave=r.llave, concepto=concepto, placa=S.manual_prev.get(r.llave, ("", ""))[1],
                      texto_tipo=fila["firma"], tercero=r.tercero if pd.notna(r.tercero) else "") for r in sel.itertuples()]
        log = [dict(empresa=r.empresa, llave=r.llave, texto=str(r.texto)[:300], valor=r.valor, concepto_antes=L.PENDIENTE,
                    concepto_despues=concepto, placa_antes=r.placa, placa_despues=r.placa, accion="Reasignada (grupo)")
               for r in sel.itertuples()]
        _guardar_overrides(S.ctx, filas, [], log)
        st.success(f"{len(sel)} movimientos clasificados como «{concepto}». Los futuros con este mismo texto y proveedor "
                   "se clasificarán igual.")
        st.rerun()


def _historial(S):
    h = S.ctx.leer("SELECT empresa, fecha, usuario, accion, texto, valor, concepto_antes, concepto_despues, placa_antes, "
                   "placa_despues, motivo, llave FROM mtto_clasif_log ORDER BY fecha DESC")
    if h.empty:
        st.info("Aún no hay reasignaciones. Cada cambio de concepto o placa, y cada eliminación, quedará aquí con usuario y fecha.")
        return
    h = h[h["empresa"].isin(S.sel)]
    st.dataframe(h.head(500), hide_index=True, width="stretch", height=360, column_config={
        "empresa": "Compañía", "texto": st.column_config.TextColumn("Texto original", width="large"),
        "valor": st.column_config.NumberColumn("Valor", format="accounting"),
        "concepto_antes": "Concepto antes", "concepto_despues": "Concepto después",
        "placa_antes": "Placa antes", "placa_despues": "Placa después", "llave": "ID del movimiento"})
    st.caption(f"{len(h)} registros de {' + '.join(S.sel)} (se muestran los 500 más recientes).")


def _memoria(S):
    ctx, es_admin, manual, catalogo = S.ctx, S.es_admin, S.manual, S.catalogo
    n1, n2, conflictos = L.construir_memoria(manual, catalogo)
    st.markdown(f"**{len(n1)}** textos con proveedor · **{len(n2)}** textos sin importar el proveedor · "
                f"**{conflictos}** con conceptos contradictorios (no se usan hasta resolverlos). "
                "La memoria es común a las dos compañías: lo que se aprende en una se aplica en la otra.")
    m = manual[manual["texto_tipo"].fillna("").ne("") & manual["concepto"].fillna("").ne("")]
    if m.empty:
        st.info("Aún no hay reasignaciones ni histórico. Cada vez que reasignes un concepto, el sistema aprende.")
        return
    t = m.groupby(["origen", "texto_tipo", "tercero", "concepto"]).size().rename("casos").reset_index() \
        .sort_values("casos", ascending=False)
    t["en_conflicto"] = t.duplicated(["texto_tipo", "tercero"], keep=False)
    t["olvidar"] = False
    ed = st.data_editor(t, hide_index=True, width="stretch", height=380, key="mt_mem_ed",
                        disabled=["origen", "texto_tipo", "tercero", "concepto", "casos", "en_conflicto"] if es_admin else True,
                        column_config={"texto_tipo": st.column_config.TextColumn("Texto (números como #)", width="large"),
                                       "tercero": "Proveedor", "concepto": st.column_config.TextColumn("Concepto", width="large"),
                                       "origen": "Origen", "casos": "Casos", "en_conflicto": "En conflicto",
                                       "olvidar": st.column_config.CheckboxColumn("Olvidar")})
    st.caption("«Olvidar» borra esas reasignaciones: los movimientos vuelven a clasificarse con las reglas. "
               "Un texto en conflicto tiene dos conceptos distintos en el histórico.")
    borrar = ed[ed["olvidar"]]
    if es_admin and st.button(f"Olvidar {len(borrar)} entradas", icon=":material/delete:", key="mt_mem_borrar",
                              disabled=borrar.empty):
        _exigir_admin(ctx)
        for r in borrar.itertuples():
            ctx.con.execute("DELETE FROM mtto_clasif_manual WHERE origen=? AND texto_tipo=? AND tercero=? AND concepto=?",
                            (r.origen, r.texto_tipo, r.tercero, r.concepto))
        ctx.con.commit()
        ctx.registrar_carga(ctx.usuario, "Memoria mantenimiento", "pantalla", f"{len(borrar)} entradas olvidadas")
        st.rerun()


def _avanzado(S, reglas, sin_guardar):
    if not S.es_admin:
        return
    with st.expander("Clasificación avanzada · pendientes por grupo, reglas, memoria e historial", icon=":material/tune:"):
        st.caption("Orden con que se clasifica cada movimiento: 1) reasignado a mano · 2) histórico aprendido "
                   "(mismo texto y proveedor ya reasignados siempre igual) · 3) reglas · 4) pendiente.")
        t_grupo, t_reg, t_mem, t_his = st.tabs(["Pendientes por grupo", "Reglas", "Memoria aprendida", "Historial de cambios"])
        with t_grupo:
            _reclasificar_grupo(S)
        with t_reg:
            _reglas(S, reglas, sin_guardar)
        with t_mem:
            _memoria(S)
        with t_his:
            _historial(S)


def _reglas(S, reglas, sin_guardar):
    ctx, es_admin, catalogo = S.ctx, S.es_admin, S.catalogo
    st.markdown(
        "**Cómo escribir un patrón** (se compara sin tildes ni mayúsculas con la columna *texto*): "
        "`ACEITE+MOTOR` = el texto debe tener ACEITE **y** MOTOR; `BANDA|VALVULA` = BANDA **o** VALVULA. "
        "El *proveedor* se busca en la columna tercero. Si pones proveedor y patrón, deben cumplirse los dos. "
        "Confianza **Baja**: el movimiento queda pendiente, con el concepto como sugerencia.")
    if sin_guardar:
        st.info("Estas son las reglas iniciales y aún no están guardadas.", icon=":material/info:")
    vista = reglas.copy()
    vista["activa"] = pd.to_numeric(vista["activa"], errors="coerce").fillna(1).astype(int).astype(bool)
    vista["orden"] = pd.to_numeric(vista["orden"], errors="coerce")
    ed = st.data_editor(vista, num_rows="dynamic", hide_index=True, width="stretch", key="mt_reglas_ed",
                        disabled=not es_admin, column_config={
                            "orden": st.column_config.NumberColumn("Orden", step=1),
                            "patron": "Texto o patrón", "proveedor": "Proveedor",
                            "concepto": st.column_config.SelectboxColumn("Concepto presupuestal", options=catalogo, width="large"),
                            "confianza": st.column_config.SelectboxColumn("Confianza", options=L.CONFIANZAS),
                            "observacion": st.column_config.TextColumn("Observación", width="large"),
                            "activa": st.column_config.CheckboxColumn("Activa")})
    if es_admin and st.button("Guardar reglas", icon=":material/save:", key="mt_guardar_reglas"):
        _guardar_reglas(ctx, ed, catalogo)
    elif not es_admin:
        st.caption("Solo el administrador puede modificar las reglas.")
    st.markdown("**Textos pendientes de clasificar** (agrupados por texto sin placa ni números y proveedor, de mayor a menor valor)")
    p = L.textos_pendientes(S.mov_todos)
    if p.empty:
        st.success("No hay movimientos pendientes de clasificación.")
    else:
        st.dataframe(p.head(60), hide_index=True, width="stretch", height=300, column_config={
            "texto_tipo": st.column_config.TextColumn("Texto (números como #)", width="large"),
            "tercero": "Proveedor / tercero", "movimientos": "Movimientos",
            "valor": st.column_config.NumberColumn("Valor", format="accounting")})
        st.caption("Úsalos para crear reglas nuevas: copia las palabras que identifican el concepto.")


def _guardar_reglas(ctx, df: pd.DataFrame, catalogo: list[str]):
    _exigir_admin(ctx)
    df = df.dropna(how="all").copy()
    for c in ("patron", "proveedor", "observacion"):
        df[c] = df[c].fillna("").astype(str).str.strip()
    df["confianza"] = df["confianza"].fillna("Media")
    sin_concepto = df["concepto"].isna() | ~df["concepto"].isin(catalogo)
    vacias = (df["patron"] == "") & (df["proveedor"] == "")
    if (sin_concepto | vacias).any():
        st.error(f"{int((sin_concepto | vacias).sum())} reglas incompletas: cada una necesita un concepto del presupuesto "
                 "y un patrón o un proveedor. No se guardó nada.")
        return
    df["orden"] = pd.to_numeric(df["orden"], errors="coerce")
    faltan = df["orden"].isna()
    df.loc[faltan, "orden"] = (df["orden"].max() if df["orden"].notna().any() else 0) + 10 * (faltan.cumsum()[faltan])
    ctx.con.execute("DELETE FROM mtto_reglas")
    for r in df.itertuples(index=False):
        ctx.con.execute("INSERT INTO mtto_reglas(orden, patron, proveedor, concepto, confianza, observacion, activa) "
                        "VALUES(?,?,?,?,?,?,?)",
                        (float(r.orden), r.patron, r.proveedor, r.concepto, r.confianza, r.observacion,
                         1 if r.activa else 0))
    ctx.con.commit()
    ctx.registrar_carga(ctx.usuario, "Reglas mantenimiento", "editor", f"{len(df)} reglas")
    st.success(f"{len(df)} reglas guardadas.")
    st.rerun()


# ================================================================== pestañas
def _movs(S, ambito: str) -> pd.DataFrame:
    m = S.vistas["movimientos"]
    return m[m["ambito"] == ambito]


def _toggle_abrir(S, key: str, etiqueta: str = "Desplegar conceptos", hondo: int = 2) -> set:
    """Niveles que se abren de entrada: con una placa seleccionada, hasta los movimientos (`hondo` = niveles que se abren);
    si no, solo el primer nivel."""
    abrir = st.toggle(etiqueta, value=bool(S.placa), key=key)
    return (set(range(hondo)) if S.placa else {0}) if abrir else set()


def _aviso_arbol(m: pd.DataFrame):
    if len(m) > MAX_MOV_ARBOL:
        st.info(f"Hay {len(m):,} movimientos en esta vista: para que la página siga ágil, el detalle de cada concepto no se incrusta "
                "en el árbol. Filtra por placa o concepto, o ábrelo completo en «Auditoría de movimientos».".replace(",", "."),
                icon=":material/info:")


def _flota_datos(S, conc_f=None, est_f=None):
    """Conceptos presupuestales de Flota (vehículos + generales) con los filtros aplicados. Los generales no se atribuyen a una
    placa ni tienen concepto de gasto: se ocultan cuando hay filtro de placa o de concepto de gasto."""
    v = S.vistas
    nodos = L.arbol_placas(v, None, conc_f or None, est_f or None)
    g = v["globales"]
    if S.placa or conc_f:
        g = g.iloc[0:0]
    elif est_f:
        g = g[g["estado"].isin(est_f)]
    return nodos, g, _categorias(S, nodos, g)


def _excel_informe_meses(S, t, g_gen, m, mg, conc_f, est_f, antes: dict | None = None, despues: dict | None = None,
                         titulo_extra: str = "", diseno: bool = False, resaltar: str | None = None, resumen=None,
                         rotulo_ppto: str = "Presupuesto", extra_placa: pd.DataFrame | None = None, extra_solo_sobre: tuple = ()) -> bytes:
    """Informe «ejecución vs presupuesto» de Flota con la ejecución SEPARADA POR MES (ejecución septiembre, ejecución octubre…; sin el
    acumulado). Son las mismas tablas de la pantalla —por concepto de gasto y, para el administrador, por placa y por placa y
    concepto— con los mismos filtros (concepto, estado, placa). Los movimientos salen de la misma base de la pantalla, así que las
    columnas de meses suman el «Ejecutado» que se ve. `antes` / `despues` = hojas adicionales (nombre -> tabla) que van antes o después
    de las tres del informe (las usa el informe reclasificado)."""
    v, cfg, hasta_mes = S.vistas, S.cfg, S.hasta_mes
    conceptos = pd.concat([t.assign(ambito=L.AMB_PLACAS), g_gen.assign(ambito=L.AMB_GEN)], ignore_index=True)
    mov_c = pd.concat([m[m["concepto"].isin(set(t["concepto"]))], mg[mg["concepto"].isin(set(g_gen["concepto"]))]], ignore_index=True) \
        if len(g_gen) else m[m["concepto"].isin(set(t["concepto"]))]
    hojas = {"Por concepto de gasto": L.tabla_por_mes(conceptos.drop(columns=["tipo_concepto", "texto_original"], errors="ignore"),
                                                      mov_c, ["ambito", "concepto"], cfg, hasta_mes)}
    if S.es_admin:
        ppc = v["por_placa_concepto"]
        if conc_f:
            ppc = ppc[ppc["concepto"].isin(conc_f)]
        if est_f:
            ppc = ppc[ppc["estado"].isin(est_f)]
        ppc = _sin_ceros(ppc)
        if conc_f or est_f:                              # con filtros, el total por placa sale de lo que queda a la vista
            meta = ppc.drop_duplicates("placa").set_index("placa")[["numero_interno", "tipo"]]
            pp = L.con_metricas(ppc.groupby("placa")[["presupuesto", "ejecutado", "movimientos"]].sum().join(meta).reset_index()
                                .assign(concepto=""))
        else:
            pp = _sin_ceros(v["por_placa"])
        mov_p = m[m["placa"].isin(set(pp["placa"]))]
        hojas["Por placa"] = L.tabla_por_mes(pp.drop(columns=["saldo", "concepto"], errors="ignore"), mov_p, ["placa"], cfg, hasta_mes)
        hojas["Por placa y concepto"] = L.tabla_por_mes(ppc.drop(columns=["saldo", "texto_original"], errors="ignore"), m,
                                                        ["placa", "concepto"], cfg, hasta_mes)
    titulo = (f"Ejecución vs presupuesto por mes · Flota y Equipo de transporte · {S.escenario} · {S.hasta}"
              + (f" · placa {S.placa}" if S.placa else "") + " · Variación = ejecución − presupuesto" + titulo_extra)
    todas = {**(antes or {}), **hojas, **(despues or {})}
    return L.escribir_informe_diseno(todas, resaltar, resumen, rotulo_ppto, extra_placa, extra_solo_sobre) if diseno else L.escribir_informe_mensual(todas, titulo)


CONCEPTO_OTROS = "OTROS"


def _excel_informe_reclasificado(S, mes: int = 0) -> bytes:
    """Informe «ejecución vs presupuesto por mes» de Flota con las referencias cuyo concepto del maestro es «OTROS» mostradas como un
    concepto propio (sin presupuesto asignado) en vez de PENDIENTE DE CLASIFICACIÓN. Solo cambia el Excel: la pantalla y la base no se
    modifican. Mismas tres hojas del informe por mes (concepto, placa, placa y concepto), más «Resumen», «Reclasificación» (qué entró
    a cada concepto desde SISPO) y «Pendientes por referencia» (lo que de verdad sigue sin concepto, sin OTROS). Sin filtros.
    `mes` = 0: todos los meses hasta el corte. `mes` = n: solo ese mes, con la ejecución de ese mes y su presupuesto. OJO: el presupuesto
    por placa y concepto se carga para TODO el escenario (no por mes); el del mes se obtiene prorrateando con el peso de ese mes en el
    presupuesto mensual del sub-rubro Flota (que sí es mensual). El «Resumen» por sub-rubro usa el presupuesto mensual real."""
    pr, er = _rubro_sel(S)
    cfg = dict(S.cfg, etiqueta=str(S.escenario).split()[0])
    nombres, orden, p_, e_ = L.serie_subrubros(pr, er, cfg)
    ppto_u, factor = S.ppto_sel, None
    if mes:
        kf = L.norm_texto(L.SUBRUBRO_FLOTA)
        m0_, m1_ = cfg["meses"]
        tot_f = sum(float(p_.get((kf, mm), 0.0)) for mm in range(m0_, m1_ + 1))
        factor = float(p_.get((kf, mes), 0.0)) / tot_f if tot_f > 0 else 0.0
        prorratea = ppto_u["es_global"].isin([0, 1])                                  # placas y generales; Tecno/SOAT ya son mensuales
        ppto_u = ppto_u.assign(valor=ppto_u["valor"].where(~prorratea, ppto_u["valor"] * factor))
    m = S.mov_todos[S.mov_todos["empresa"].isin(S.sel)]
    if mes:
        m = m[pd.to_numeric(m["mes"], errors="coerce") == mes]
    sin_clas = m["concepto"] == L.PENDIENTE                                          # todo lo que no quedó en un concepto del presupuesto
    m2 = m.assign(concepto=m["concepto"].mask(sin_clas, CONCEPTO_OTROS))
    v2 = L.calcular(ppto_u, m2, S.oficial, S.placa, S.tipos)
    S2 = SimpleNamespace(**{**vars(S), "vistas": v2})
    if mes:                                              # las columnas de «tabla_por_mes» salen de cfg y hasta_mes
        S2.cfg, S2.hasta_mes, S2.hasta = dict(S.cfg, meses=(mes, S.cfg["meses"][1])), mes, f"{MESES_L[mes - 1]} {S.cfg['anio']}"
    t = _sin_ceros(v2["por_concepto"][v2["por_concepto"]["tipo_concepto"].isin([L.AMB_PLACAS, "Pendiente", CONCEPTO_OTROS])])
    _, g_gen, _ = _flota_datos(S2)
    g_gen = _sin_ceros(g_gen)
    mp, mg = _movs(S2, L.AMB_PLACAS), _movs(S2, L.AMB_GEN)

    # ---- «Por placa»: ejecutado por origen (SISPO / SAP) y, en las placas sobre ejecutadas, los conceptos con mayor ejecución
    mpv = mp.assign(valor=pd.to_numeric(mp["valor"], errors="coerce").fillna(0.0))
    es_sp = mpv["sispo"].fillna(False).astype(bool) if "sispo" in mpv else pd.Series(False, index=mpv.index)
    v_sispo, v_sap = mpv[es_sp].groupby("placa")["valor"].sum(), mpv[~es_sp].groupby("placa")["valor"].sum()
    g_pc = mpv.groupby(["placa", "concepto"])["valor"].sum().reset_index()
    g_pc = g_pc[g_pc["valor"] > 0].sort_values(["placa", "valor"], ascending=[True, False])
    top = {pl: list(zip(x["concepto"], x["valor"]))[:3] for pl, x in g_pc.groupby("placa")}
    cols_top = [c for i in (1, 2, 3) for c in (f"Concepto {i} de mayor ejecución", f"Valor concepto {i}")]
    filas_ex = []
    for pl in sorted(set(mpv["placa"])):
        fila_ = {"Placa": pl, "Ejecutado SISPO": float(v_sispo.get(pl, 0.0)), "Ejecutado SAP": float(v_sap.get(pl, 0.0))}
        for i, (cn, vl) in enumerate(top.get(pl, [])[:3], start=1):
            fila_[f"Concepto {i} de mayor ejecución"], fila_[f"Valor concepto {i}"] = cn, float(vl)
        filas_ex.append(fila_)
    extra_placa = pd.DataFrame(filas_ex).reindex(columns=["Placa", "Ejecutado SISPO", "Ejecutado SAP", *cols_top])
    # ---- «Detalle sobre ejecución»: de cada placa sobre ejecutada, los movimientos más representativos de SISPO (salidas de inventario)
    # y de SAP (todo lo demás: facturas, servicios…); lo que no entra en el top va en una fila «Resto» para que cada origen cuadre
    N_TOP = 5
    pp_ = v2["por_placa"]
    sobre = pp_[pp_["estado"] == "Sobre ejecutado"].assign(_exceso=lambda d: d["ejecutado"] - d["presupuesto"]).sort_values("_exceso", ascending=False)
    filas_d = []
    for r_ in sobre.itertuples():
        mpl = mpv[mpv["placa"] == r_.placa]
        tot_pl = float(mpl["valor"].sum())
        for origen_, sel_ in (("SISPO", es_sp[mpl.index]), ("SAP", ~es_sp[mpl.index])):
            x = mpl[sel_.values]
            if x.empty:
                continue
            x = x.sort_values("valor", ascending=False)
            for mv_ in x.head(N_TOP).itertuples():
                es_k = origen_ == "SISPO"
                filas_d.append(dict(
                    Placa=r_.placa, Tipo=getattr(r_, "tipo", ""), Origen=origen_, Concepto=mv_.concepto,
                    Fecha=str(mv_.fecha)[:10], Documento=str(mv_.documento),
                    Referencia=str(getattr(mv_, "codigo_articulo", "") or "") if es_k else "",
                    **{"Artículo / texto": str(mv_.texto), "Proveedor / bodega": str(mv_.tercero or "")},
                    Cantidad=float(mv_.cantidad) if es_k and pd.notna(getattr(mv_, "cantidad", None)) else None,
                    Valor=float(mv_.valor), **{"% de la ejecución de la placa": float(mv_.valor) / tot_pl if tot_pl else None}))
            if len(x) > N_TOP:
                rest = x.iloc[N_TOP:]
                filas_d.append(dict(
                    Placa=r_.placa, Tipo=getattr(r_, "tipo", ""), Origen=origen_, Concepto="", Fecha="", Documento="", Referencia="",
                    **{"Artículo / texto": f"Resto: {len(rest)} movimientos de menor valor", "Proveedor / bodega": ""},
                    Cantidad=None, Valor=float(rest["valor"].sum()),
                    **{"% de la ejecución de la placa": float(rest["valor"].sum()) / tot_pl if tot_pl else None}))
    detalle = pd.DataFrame(filas_d, columns=["Placa", "Tipo", "Origen", "Concepto", "Fecha", "Documento", "Referencia", "Artículo / texto",
                                             "Proveedor / bodega", "Cantidad", "Valor", "% de la ejecución de la placa"])
    # ---- hoja «Resumen»: por sub-rubro y mes, y origen de la ejecución de Flota
    meses = [mes] if mes else list(range(S.m0, S.hasta_mes + 1))
    mv = S.vistas["movimientos"]
    mv = mv[mv["ambito"].isin([L.AMB_PLACAS, L.AMB_GEN])]
    mv_s = mv["sispo"].fillna(False).astype(bool) if "sispo" in mv else pd.Series(False, index=mv.index)
    por_mes = lambda d: {int(k): float(x) for k, x in d.groupby("mes")["valor"].sum().items()}
    origen = [("Salidas de inventario (SISPO, con detalle por referencia)", por_mes(mv[mv_s]), "Kardex de SISPO ya reclasificado por placa y concepto."),
              ("Otros movimientos de SAP (facturas, servicios…)", por_mes(mv[~mv_s]), "")]
    sub = f"Escenario {S.escenario} · " + (f"solo {MESES_L[mes - 1]} · " if mes else "") + " y ".join(MESES_L[m - 1].capitalize() for m in meses) + " por separado · Valores en pesos"
    fn_resumen = lambda ws: L.hoja_resumen_mantenimiento(ws, "Ejecución vs presupuesto · Mantenimiento · " + ", ".join(S.sel), sub, meses, nombres,
                                                         orden, p_, e_, cfg, origen, [], None)
    return _excel_informe_meses(S2, t, g_gen, mp, mg, [], [], antes={"Resumen": pd.DataFrame()},      # el contenido lo escribe fn_resumen
                                despues={"Detalle sobre ejecución": detalle}, titulo_extra=" · OTROS como concepto propio",
                                diseno=True, resaltar=CONCEPTO_OTROS, resumen=fn_resumen, extra_placa=extra_placa, extra_solo_sobre=tuple(cols_top),
                                rotulo_ppto=f"Presupuesto {MESES_L[mes - 1]}\n(prorrateado)" if mes else "Presupuesto")


def _tab_flota(S):
    """Flota y Equipo de transporte. TODOS los perfiles ven la tabla por concepto de gasto (concepto -> movimientos), los indicadores,
    los conceptos generales y el gráfico, con los datos que su perfil puede ver (cada gerencia, solo lo suyo). Solo el administrador
    tiene además la carga mensual de SISPO, el filtro/vista por placa, la auditoría y la reclasificación."""
    if S.es_admin:
        _sispo_panel(S)                                 # carga mensual del SISPO: arriba y abierto, es lo primero que se usa
    v = S.vistas
    ppc = v["por_placa_concepto"]
    opciones = sorted(set(ppc["concepto"]))
    _limpiar_multi("mt_f_concepto", opciones)
    with st.container(key="mt_card_filtros_placas"):
        f1, f2 = st.columns([1.6, 1.4], vertical_alignment="bottom")
        conc_f = f1.multiselect("Concepto de gasto", opciones, placeholder="Todos los conceptos de gasto", key="mt_f_concepto",
                                help="Llantas, aceite, frenos… Con este filtro (o con una placa) no se muestran los conceptos "
                                     "generales, que no tienen placa ni concepto de gasto.")
        est_f = f2.multiselect("Estado de ejecución", ESTADOS, placeholder="Todos los estados", key="mt_f_estado")
    nodos, g_gen, cats = _flota_datos(S, conc_f, est_f)
    g_gen = _sin_ceros(g_gen)                           # los conceptos generales en 0 (presupuesto y ejecución) tampoco se muestran
    # «Otras placas sin ejecución» se abre sola cuando la persona busca justamente esas placas (una placa o el estado «Sin ejecución»)
    S.abrir_otras = bool(S.placa) or bool(est_f and set(est_f) <= {"Sin ejecución"})
    t_cats = _df_categorias(cats)
    tot = t_cats.iloc[-1] if len(t_cats) else dict(presupuesto=0.0, ejecutado=0.0, movimientos=0)
    _kpis(float(tot["presupuesto"]), float(tot["ejecutado"]), int(tot["movimientos"]), S.hasta,
          f"Escenario {S.escenario}" + (f" · placa {S.placa}" if S.placa else ""))
    mp, mg = _movs(S, L.AMB_PLACAS), _movs(S, L.AMB_GEN)
    pend = mp[mp["concepto"] == L.PENDIENTE]
    if not pend.empty:
        st.warning(f"{len(pend)} movimientos por {E.pesos(pend['valor'].sum())} están pendientes de clasificación: siguen "
                   "sumando en la ejecución, pero no en ningún concepto de gasto. "
                   + ("Asígnalos abajo, en «Reclasificación de movimientos»." if S.es_admin else "El administrador los clasifica."),
                   icon=":material/rule:")
    if S.placa:
        st.info(f"Con la placa **{S.placa}** seleccionada no se muestran los conceptos generales: son gastos de la operación y no "
                "se atribuyen a ninguna placa.", icon=":material/info:")
    # una vista guardada en la sesión que ya no existe («Por concepto presupuestal») se descarta antes de crear el control
    if st.session_state.get("mt_vista_placas") not in ("Por placa", "Por concepto de gasto", None):
        del st.session_state["mt_vista_placas"]
    if S.es_admin:
        vista = st.segmented_control("Vista", ["Por concepto de gasto", "Por placa"], default="Por concepto de gasto",
                                     key="mt_vista_placas") or "Por concepto de gasto"
    else:                                               # el resto de perfiles ve la vista por concepto de gasto (sin nivel de placa)
        vista = "Por concepto de gasto"
    m = mp[mp["concepto"].isin(conc_f)] if conc_f else mp
    # conceptos de gasto con filtros aplicados, sin los que no tienen presupuesto, ejecución ni movimientos (todo en 0: no dicen nada).
    # Alimenta la tabla por concepto de gasto y el gráfico.
    t = _sin_ceros(v["por_concepto"][v["por_concepto"]["tipo_concepto"].isin([L.AMB_PLACAS, "Pendiente"])])
    if conc_f:
        t = t[t["concepto"].isin(conc_f)]
    if est_f:
        t = t[t["estado"].isin(est_f)]
    if vista == "Por placa":
        with st.container(key="mt_card_placa"):
            titulo_tarjeta("Presupuesto y ejecución por placa", f"{S.escenario} · ejecutado {S.hasta} · placa → concepto de gasto → movimiento")
            abrir = _toggle_abrir(S, "mt_abrir_placas", "Desplegar conceptos", hondo=2)
            _aviso_arbol(m)
            st.html(_html_placa_concepto(S, nodos, m, abrir, False, "Placa / concepto de gasto / movimiento"))
            st.caption("Cada placa se abre en sus conceptos de gasto y en los movimientos que los componen (fecha, documento, proveedor y "
                       "texto original; el valor va en «Ejecutado»). Las placas sin ejecución se reúnen en «Otras placas sin ejecución». "
                       "«SIN PLACA IDENTIFICADA» reúne lo ejecutado cuyo texto no trae una placa del listado oficial.")
    else:
        with st.container(key="mt_card_conc"):
            titulo_tarjeta("Presupuesto y ejecución por concepto de gasto",
                           f"{S.escenario} · ejecutado {S.hasta}" + (f" · solo la placa {S.placa}" if S.placa else " · todas las placas")
                           + " · concepto de gasto → movimiento")
            abrir = _toggle_abrir(S, "mt_abrir_conc", "Desplegar movimientos", hondo=1)
            _aviso_arbol(m)
            st.html(_html_concepto(S, t, m, abrir, "Concepto de gasto / movimiento"))
            with st.expander("Ver como tabla", icon=":material/table:"):
                _tabla(t, 420)
            # descarga: todos los movimientos de los conceptos que se ven (con los filtros de concepto, estado y placa); si los
            # conceptos generales se ven, también van (columna «Ámbito»), así el Excel suma lo mismo que el indicador de ejecución
            conceptos_vis = set(t["concepto"])
            md = m[m["concepto"].isin(conceptos_vis)]
            if len(g_gen):
                md = pd.concat([md, mg[mg["concepto"].isin(set(g_gen["concepto"]))]], ignore_index=True)
            nombre_xl = "movimientos_por_concepto_de_gasto_" + str(S.hasta).replace(" ", "_") + (f"_{S.placa}" if S.placa else "") + ".xlsx"
            st.download_button(f"Descargar movimientos ({len(md):,})".replace(",", "."), lambda: _excel_movs_concepto(
                                   pd.concat([t.assign(ambito=L.AMB_PLACAS), g_gen.assign(ambito=L.AMB_GEN)], ignore_index=True), md), nombre_xl,
                               icon=":material/download:", key="mt_dl_mov_concepto", disabled=md.empty,
                               help="Un renglón por movimiento, con su concepto de gasto, placa, documento, proveedor y texto: una hoja con todos, "
                                    "UNA HOJA POR MES, el resumen por concepto y el resumen concepto × mes. Respeta los filtros de concepto, "
                                    "estado y placa.")
            st.caption("Cada concepto de gasto se abre en los movimientos clasificados en él (la placa de cada uno va como etiqueta"
                       + ("; para ver el detalle por placa usa la vista «Por placa»" if S.es_admin else "") + "). «Pendiente» = movimientos "
                       "sin concepto asignado.")
    nombre_inf = "ejecucion_vs_presupuesto_por_mes_" + str(S.hasta).replace(" ", "_").replace("–", "-") + (f"_{S.placa}" if S.placa else "") + ".xlsx"
    st.download_button("Descargar informe ejecución vs presupuesto por mes", lambda: _excel_informe_meses(S, t, g_gen, m, mg, conc_f, est_f),
                       nombre_inf, icon=":material/download:", key="mt_dl_informe_meses", disabled=t.empty and g_gen.empty,
                       help="Las tablas de esta página (por concepto de gasto" + (", por placa y por placa y concepto" if S.es_admin else "") +
                            ") con presupuesto, ejecución de CADA mes por separado (p. ej. septiembre y octubre), variación y % de ejecución. "
                            "No trae el acumulado. Respeta los filtros de concepto, estado y placa.")
    if S.es_admin:                                      # trae detalle de SISPO de todas las gerencias: solo administrador
        opc_mes = [0, *range(S.m0, S.hasta_mes + 1)]
        mes_inf = st.selectbox("Mes del informe reclasificado", opc_mes, index=len(opc_mes) - 1, key="mt_mes_informe_recl",
                               format_func=lambda x: "Todos los meses (columna por mes)" if x == 0 else MESES_L[x - 1].capitalize(),
                               help="Un mes: solo la ejecución de ese mes y su presupuesto (el de placa y concepto se prorratea; el Resumen usa el "
                                    "presupuesto mensual real). «Todos los meses»: el informe completo con una columna por mes.")
        st.download_button("Descargar informe reclasificado (OTROS como concepto)", lambda: _excel_informe_reclasificado(S, mes_inf),
                           "ejecucion_reclasificada_" + (f"{MESES_L[mes_inf - 1]}_{S.cfg['anio']}" if mes_inf else str(S.hasta).replace(" ", "_").replace("–", "-")) + ".xlsx", icon=":material/download:",
                           key="mt_dl_informe_reclasificado", disabled=S.mov_todos.empty,
                           help="El informe por mes, sin filtros, con todo lo que no quedó en un concepto del presupuesto como «OTROS» "
                                "(sin presupuesto), saldo y % de avance con fórmulas, barras de avance y negativos en rojo, más las hojas "
                                "«Resumen», «Reclasificación» y «Pendientes por referencia». Solo cambia el Excel: la pantalla no se modifica.")
    if len(g_gen):                                      # los conceptos generales no tienen placa: se ven aparte, así nada queda fuera de la pestaña
        with st.expander(f"Conceptos generales, sin placa ({len(g_gen)})", icon=":material/category:"):
            _aviso_arbol(mg)
            st.html(_html_generales(S, g_gen, mg, _toggle_abrir(S, "mt_abrir_gen", "Desplegar movimientos", hondo=1)))
            st.caption("Gastos de la operación que no se atribuyen a ninguna placa (In House, grúa, lavado…). Suman en los indicadores "
                       "de arriba, pero no en la tabla por concepto de gasto ni por placa.")
    if S.es_admin and not (conc_f or est_f or S.placa):
        r = v["resumen"]
        esperado = r.loc[r["ambito"].isin([L.AMB_PLACAS, L.AMB_GEN]), ["presupuesto", "ejecutado"]].sum()
        dp, de = tot["presupuesto"] - esperado["presupuesto"], tot["ejecutado"] - esperado["ejecutado"]
        if abs(dp) >= 1 or abs(de) >= 1:
            st.error(f"Los indicadores no coinciden con Placas + Generales: presupuesto {E.pesos(dp)}, ejecutado {E.pesos(de)}.")
    with st.container(key="mt_card_graf_placas"):
        titulo_tarjeta("Conceptos de gasto que más pesan", "Presupuesto vs ejecutado (millones de pesos)")
        _barras_ppto_ejec(t[["concepto", "presupuesto", "ejecutado"]],
                          "Los 12 conceptos de gasto con mayor presupuesto o ejecución" + (f" (placa {S.placa})." if S.placa else "."),
                          con_pendiente=True)
    _auditoria(S, pd.concat([m, mg]) if not (S.placa or conc_f) else m, "flota", "mantenimiento_flota")


GRUPO_CON, GRUPO_SIN = "Placas con presupuesto", "Otras placas"


def _tab_ts(S):
    """Tecnomecánica y SOAT: dos conceptos independientes, cada uno con sus propios indicadores, grupos y detalle por placa."""
    if S.ts_scen.empty:
        st.warning("Aún no hay presupuesto de Tecnomecánica y SOAT cargado. El administrador lo carga en «Datos base». "
                   "Mientras tanto se muestra solo la ejecución.", icon=":material/warning:")
    d = S.vistas["ts_placa_concepto"]
    t = pd.DataFrame({"concepto": L.CONCEPTOS_TS}).merge(
        d.groupby("concepto", as_index=False).agg(presupuesto=("presupuesto", "sum"), ejecutado=("ejecutado", "sum"),
                                                   movimientos=("movimientos", "sum")), how="left")
    t[["presupuesto", "ejecutado", "movimientos"]] = t[["presupuesto", "ejecutado", "movimientos"]].fillna(0)
    t["movimientos"] = t["movimientos"].astype(int)
    t = L.con_metricas(t)[["concepto", "presupuesto", "ejecutado", "saldo", "pct", "estado", "movimientos"]]
    with st.container(key="mt_card_ts_res"):
        titulo_tarjeta("Tecnomecánica y SOAT por separado",
                       f"{S.escenario} · ejecutado {S.hasta}" + (f" · placa {S.placa}" if S.placa else " · todas las placas"))
        _tabla(t, 120)
    t_tec, t_soat = st.tabs(["Tecnomecánica", "SOAT"])
    with t_tec:
        _ts_concepto(S, L.CONCEPTO_TECNO, "tec")
    with t_soat:
        _ts_concepto(S, L.CONCEPTO_SOAT, "soat")


def _ts_concepto(S, concepto: str, k: str):
    """Una sola categoría (Tecnomecánica o SOAT): indicadores, vehículos con y sin presupuesto, detalle por placa y mes a mes."""
    v = S.vistas
    est_f = st.multiselect("Estado de ejecución", ESTADOS, placeholder="Todos los estados", key=f"mt_f_estado_ts_{k}")
    todos = L.arbol_ts(v, None, [concepto], est_f or None)
    # solo placas con algo que mostrar (presupuesto, ejecución o movimientos); una placa «tiene presupuesto» si su presupuesto > 0
    nodos = [n for n in todos if n["presupuesto"] != 0 or n["ejecutado"] != 0 or n["movimientos"] > 0]
    con = [n for n in nodos if n["presupuesto"] > 0]
    sin = [n for n in nodos if n["presupuesto"] <= 0]
    tp, te, tm = _totales(nodos)
    _kpis(tp, te, tm, S.hasta, f"Escenario {S.escenario}" + (f" · placa {S.placa}" if S.placa else ""))

    resumen = _resumen_grupos(nodos, lambda n: GRUPO_CON if n["presupuesto"] > 0 else GRUPO_SIN, total=True)
    orden = {GRUPO_CON: 0, GRUPO_SIN: 1, "Total": 2}
    resumen = resumen.sort_values("grupo", key=lambda c: c.map(orden)).reset_index(drop=True)
    with st.container(key=f"mt_card_ts_grupos_{k}"):
        titulo_tarjeta(f"{concepto.capitalize() if concepto != L.CONCEPTO_SOAT else 'SOAT'}: vehículos con y sin presupuesto asignado",
                       f"{S.escenario} · ejecutado {S.hasta} · «{GRUPO_SIN}» = placas sin presupuesto de este concepto (con o sin ejecución)")
        _tabla(resumen[["grupo", "placas", "presupuesto", "ejecutado", "saldo", "pct", "estado", "movimientos"]], 40 + 35 * len(resumen))
    e = _movs(S, L.AMB_TS)
    e = e[e["concepto"] == concepto]
    if S.es_admin:
        with st.container(key=f"mt_card_ts_placa_{k}"):
            titulo_tarjeta("Detalle por vehículo", f"{GRUPO_CON} / {GRUPO_SIN} → placa → presupuesto mensual y movimientos")
            abrir = _toggle_abrir(S, f"mt_abrir_ts_{k}", hondo=2)
            _aviso_arbol(e)
            st.html(_html_grupos_ts(S, [(GRUPO_CON, con), (GRUPO_SIN, sin)], e, abrir, "Grupo / placa / detalle", concepto))
            st.caption("De entrada solo se ven los dos grupos; cada uno se abre en sus placas y cada placa en su presupuesto mes a mes "
                       "(en cursiva) y en los movimientos que componen su ejecución. «SIN PLACA IDENTIFICADA» (dentro de "
                       f"«{GRUPO_SIN}»): movimientos cuyo texto no trae una placa del listado oficial; asígnales la placa en "
                       "«Reclasificación de movimientos».")
    pm = S.ts_scen[S.ts_scen["concepto"] == concepto]
    if S.placa:
        pm = pm[pm["placa"] == S.placa]
    with st.container(key=f"mt_card_ts_mes_{k}"):
        titulo_tarjeta("Mes a mes", "Presupuesto vs ejecutado (millones de pesos)")
        _grafico_ts_mensual(S, pm, e)
    _auditoria(S, e, f"ts_{k}", f"mantenimiento_{k}")


# ================================================================== reclasificación mensual con SISPO (administrador)
def _pesos2(c) -> str:
    """Centavos (o pesos con decimales) en formato colombiano: 1.053.124.780,14"""
    v = float(c)
    return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _a_db(v):
    """Valor listo para la base: NaN/NaT -> None."""
    return None if (not isinstance(v, str) and pd.isna(v)) else (v.item() if hasattr(v, "item") else v)


def _sispo_cargar_maestro(ctx, m: pd.DataFrame, reemplazar: bool, archivo: str):
    """Guarda la base maestra de referencias. Por defecto ACTUALIZA (agrega las nuevas y corrige el concepto de las existentes,
    sin tocar las demás); «reemplazar» deja solo lo que trae el archivo."""
    _exigir_admin(ctx)
    ahora_txt = ahora().strftime("%Y-%m-%d %H:%M")
    if reemplazar:
        ctx.con.execute("DELETE FROM mtto_sispo_ref")
    else:
        for i in range(0, len(m), 200):
            parte = list(m["codigo"].iloc[i:i + 200])
            ctx.con.execute("DELETE FROM mtto_sispo_ref WHERE codigo IN (" + ",".join("?" * len(parte)) + ")", tuple(parte))
    _insertar_por_lotes(ctx, "mtto_sispo_ref(codigo, articulo, concepto, usuario, fecha)",
                        [(r.codigo, r.articulo, r.concepto, ctx.usuario, ahora_txt) for r in m.itertuples()])
    ctx.con.commit()
    ctx.registrar_carga(ctx.usuario, "Base maestra SISPO", archivo, f"{len(m)} referencias" + (" (reemplazo total)" if reemplazar else ""))


def _cond_cuenta(col: str, cuenta: str) -> tuple[str, tuple]:
    """Condición SQL «es de esta cuenta». Lo guardado antes de existir las cuentas no la trae (NULL o vacío): es de la flota."""
    if cuenta == L.SISPO_CUENTA:
        return f"({col}=? OR {col} IS NULL OR {col}='')", (cuenta,)
    return f"{col}=?", (cuenta,)


def _sispo_guardar_lineas(ctx, empresa: str, periodo: int, cuenta: str, lin: pd.DataFrame):
    """Reescribe las líneas de UNA cuenta de un periodo (borra las anteriores y las inserta): aplicar dos veces el mismo archivo
    deja lo mismo, y reprocesar una cuenta no toca las demás."""
    cond, par = _cond_cuenta("cuenta_nombre", cuenta)
    ctx.con.execute(f"DELETE FROM mtto_sispo_lineas WHERE empresa=? AND periodo=? AND {cond}", (empresa, int(periodo), *par))
    _insertar_por_lotes(ctx, "mtto_sispo_lineas(empresa, periodo, id_kardex, bodega, fecha_mov, codigo, articulo, cantidad, "
                             "costo_unidad, costo_sin_iva, costo_con_iva, centro_costo, cc_codigo, cc_nombre, cuenta_nombre, placa, "
                             "placa_estado, concepto, concepto_estado, valor_sap)",
                        [(empresa, int(periodo), r.id_kardex, r.bodega, _a_db(r.fecha_mov), r.codigo, r.articulo, _a_db(r.cantidad),
                          _a_db(r.costo_unidad), _a_db(r.costo_sin_iva), _a_db(r.costo_con_iva), r.centro_costo, r.cc_codigo,
                          r.cc_nombre, cuenta, r.placa, r.placa_estado, r.concepto, r.concepto_estado, _a_db(r.valor_sap))
                         for r in lin.itertuples()], lote=40)


def _sispo_aplicar(ctx, empresa: str, periodo: int, cuenta: str, lin: pd.DataFrame, sap: pd.DataFrame, conc: dict, archivo: str,
                   huella: str, huella_m: str, generico: str):
    """Aplica (o reprocesa) la reclasificación de UNA cuenta de UN periodo. Idempotente: reemplaza lo aplicado antes para esa
    compañía, periodo y cuenta, sin tocar las otras cuentas. Solo se ejecuta si la conciliación cuadra al centavo y si todo lo
    que se va a guardar es de esa cuenta; las validaciones se repiten aquí, no solo en la pantalla."""
    _exigir_admin(ctx)
    if not conc["cuadra"]:
        raise ValueError(f"{cuenta}: el total de SISPO no coincide con el de SAP; no se aplica la reclasificación.")
    if len(lin) and lin["id_kardex"].astype(str).str.startswith(L.PREFIJO_ID_SINTETICO).any():
        raise ValueError(f"{cuenta}: el SISPO no trae «IdKardex» (los identificadores «{L.PREFIJO_ID_SINTETICO}n» son provisionales); no se aplica.")
    if len(lin) and not (lin["cuenta"] == cuenta).all():
        raise ValueError(f"{cuenta}: hay salidas de otra cuenta en la lista a aplicar; no se aplica.")
    if len(sap) and (L.cuenta_sap(sap) != cuenta).any():
        raise ValueError(f"{cuenta}: hay movimientos de SAP de otra cuenta en lo que se va a reemplazar; no se aplica.")
    ahora_txt = ahora().strftime("%Y-%m-%d %H:%M")
    for t in ("mtto_sispo_reemplazados", "mtto_sispo_periodo"):
        cond, par = _cond_cuenta("cuenta", cuenta)
        ctx.con.execute(f"DELETE FROM {t} WHERE empresa=? AND periodo=? AND {cond}", (empresa, int(periodo), *par))
    guardar = lin.assign(fecha_mov=lin["fecha"].dt.strftime("%Y-%m-%d %H:%M:%S"), cuenta_nombre=cuenta, valor_sap=-lin["costo_con_iva"])
    _sispo_guardar_lineas(ctx, empresa, periodo, cuenta, guardar)
    vacia = pd.Series("", index=sap.index)
    col = lambda c: (sap[c] if c in sap else vacia).fillna("")
    _insertar_por_lotes(ctx, "mtto_sispo_reemplazados(empresa, periodo, llave, fecha_mov, documento, cc, texto, valor, gerencia, codigo, "
                             "cuenta, sub_rubro, sub_rubro_sap)",
                        [(empresa, int(periodo), str(l), str(f)[:10], str(d), str(c), str(t), float(v), str(g), str(k), cuenta, str(sr), str(ss))
                         for l, f, d, c, t, v, g, k, sr, ss in zip(col("llave"), col("fecha"), col("documento"), col("cc"), col("texto"),
                                                                   col("valor"), col("gerencia"), col("codigo"), col("sub_rubro"),
                                                                   col("sub_rubro_sap"))])
    n_pend = int((lin["concepto"] == L.PENDIENTE).sum())
    ctx.con.execute("INSERT INTO mtto_sispo_periodo(empresa, periodo, archivo, huella, huella_maestro, usuario, fecha, n_lineas, "
                    "n_pendientes, valor_sispo, valor_sap, diferencia, estado, n_reemplazados, generico, cuenta) "
                    "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                    (empresa, int(periodo), archivo, huella, huella_m, ctx.usuario, ahora_txt, len(lin), n_pend,
                     conc["sispo_c"] / 100, conc["sap_c"] / 100, conc["diferencia_c"] / 100,
                     "Conciliado", len(sap), generico, cuenta))
    ctx.con.commit()
    ctx.registrar_carga(ctx.usuario, "Reclasificación SISPO", archivo,
                        f"{empresa} {periodo} · {cuenta}: {len(lin)} líneas · {len(sap)} movimientos de SAP reemplazados")


def _sispo_revertir(ctx, empresa: str, periodo: int, cuenta: str):
    """Quita la reclasificación de UNA cuenta de un periodo: vuelven sus movimientos genéricos de SAP tal como estaban."""
    _exigir_admin(ctx)
    for t, col in (("mtto_sispo_lineas", "cuenta_nombre"), ("mtto_sispo_reemplazados", "cuenta"), ("mtto_sispo_periodo", "cuenta")):
        cond, par = _cond_cuenta(col, cuenta)
        ctx.con.execute(f"DELETE FROM {t} WHERE empresa=? AND periodo=? AND {cond}", (empresa, int(periodo), *par))
    ctx.con.commit()
    ctx.registrar_carga(ctx.usuario, "Reversión reclasificación SISPO", "pantalla", f"{empresa} {periodo} · {cuenta}")


def _sispo_reasignar(ctx, empresa: str, periodo: int, cuenta: str, maestro: pd.DataFrame, catalogos: dict) -> tuple[int, int]:
    """Vuelve a cruzar las líneas YA aplicadas de una cuenta y periodo con la base maestra vigente (sin necesitar el SISPO otra
    vez). Los valores no cambian, solo el concepto; devuelve (líneas que cambiaron de concepto, pendientes que quedan)."""
    _exigir_admin(ctx)
    cond, par = _cond_cuenta("cuenta_nombre", cuenta)
    lin = ctx.leer(f"SELECT * FROM mtto_sispo_lineas WHERE empresa=? AND periodo=? AND {cond}", (empresa, int(periodo), *par))
    if lin.empty:
        return 0, 0
    nuevo = L.sispo_cruzar_cuentas(lin.assign(cuenta=cuenta), maestro, catalogos)
    cambios = int((nuevo["concepto"] != lin["concepto"]).sum())
    _sispo_guardar_lineas(ctx, empresa, periodo, cuenta, nuevo)
    n_pend = int((nuevo["concepto"] == L.PENDIENTE).sum())
    cond, par = _cond_cuenta("cuenta", cuenta)
    ctx.con.execute(f"UPDATE mtto_sispo_periodo SET huella_maestro=?, n_pendientes=? WHERE empresa=? AND periodo=? AND {cond}",
                    (L.sispo_huella_maestro(maestro), n_pend, empresa, int(periodo), *par))
    ctx.con.commit()
    ctx.registrar_carga(ctx.usuario, "Reasignación conceptos SISPO", "pantalla", f"{empresa} {periodo} · {cuenta}: {cambios} líneas con concepto nuevo")
    return cambios, n_pend


def _excel_simple(hojas: dict) -> bytes:
    from io import BytesIO
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for nombre, df in hojas.items():
            df.to_excel(w, sheet_name=nombre[:31], index=False)
    return buf.getvalue()


def _excel_historico_sispo(ctx) -> bytes:
    """Histórico acumulado del SISPO aplicado: TODAS las compañías, periodos y cuentas que hay en la base (mtto_sispo_lineas /
    mtto_sispo_periodo). Se arma al momento de descargar, así siempre trae lo último y crece con cada periodo que se aplica. Hojas:
    «Histórico SISPO» (una fila por salida), «Resumen» (compañía × periodo × cuenta × concepto) y «Periodos aplicados»."""
    lin = ctx.leer("SELECT * FROM mtto_sispo_lineas ORDER BY periodo, empresa, fecha_mov, id_kardex")
    per = ctx.leer("SELECT empresa, periodo, cuenta, archivo, usuario, fecha, n_lineas, n_pendientes, valor_sispo, valor_sap, diferencia, estado, "
                   "n_reemplazados FROM mtto_sispo_periodo ORDER BY periodo, empresa, cuenta")
    if len(lin):
        lin["cuenta_nombre"] = lin["cuenta_nombre"].map(lambda x: L.cuenta_de(x) or L.SISPO_CUENTA)
        lin["fecha_mov"] = pd.to_datetime(lin["fecha_mov"], errors="coerce")
    if len(per):
        per["cuenta"] = per["cuenta"].map(lambda x: L.cuenta_de(x) or L.SISPO_CUENTA)
    nombres = {"empresa": "Compañía", "periodo": "Periodo", "cuenta_nombre": "Cuenta", "id_kardex": "IdKardex", "bodega": "Bodega",
               "fecha_mov": "Fecha", "codigo": "Referencia", "articulo": "Artículo", "cantidad": "Cantidad", "costo_unidad": "Costo unidad",
               "costo_sin_iva": "Costo sin IVA", "costo_con_iva": "Costo con IVA (SISPO)", "centro_costo": "Centro de costo (SISPO)",
               "cc_codigo": "Código centro de costo", "cc_nombre": "Nombre centro de costo", "placa": "Placa", "placa_estado": "Estado de la placa",
               "concepto": "Concepto", "concepto_estado": "Estado del concepto", "valor_sap": "Valor (signo SAP)"}
    hist = lin[[c for c in nombres if c in lin.columns]].rename(columns=nombres) if len(lin) else pd.DataFrame(columns=list(nombres.values()))
    if len(lin):
        res = (lin.groupby(["empresa", "periodo", "cuenta_nombre", "concepto"], dropna=False)
               .agg(salidas=("id_kardex", "size"), referencias=("codigo", "nunique"), placas=("placa", "nunique"), valor=("valor_sap", "sum"))
               .reset_index().rename(columns={"empresa": "Compañía", "periodo": "Periodo", "cuenta_nombre": "Cuenta", "concepto": "Concepto",
                                              "salidas": "Salidas", "referencias": "Referencias", "placas": "Placas", "valor": "Valor (signo SAP)"}))
    else:
        res = pd.DataFrame(columns=["Compañía", "Periodo", "Cuenta", "Concepto", "Salidas", "Referencias", "Placas", "Valor (signo SAP)"])
    per = per.rename(columns={"empresa": "Compañía", "periodo": "Periodo", "cuenta": "Cuenta", "archivo": "Archivo", "usuario": "Usuario",
                              "fecha": "Aplicado", "n_lineas": "Salidas", "n_pendientes": "Pendientes", "valor_sispo": "Valor SISPO",
                              "valor_sap": "Valor SAP", "diferencia": "Diferencia", "estado": "Estado", "n_reemplazados": "Mov. SAP reemplazados"})
    return _excel_simple({"Histórico SISPO": hist, "Resumen": res, "Periodos aplicados": per})


def _excel_movs_concepto(t: pd.DataFrame, mov: pd.DataFrame) -> bytes:
    """Excel de los movimientos que componen la tabla por concepto de gasto (los mismos filtros que se ven en pantalla). Hojas:
    «Movimientos» (todos, un renglón por movimiento), UNA HOJA POR MES con los movimientos de ese mes, «Resumen por concepto» y
    «Resumen por mes» (concepto × mes). «Valor» conserva el signo de SAP, igual que la columna «Ejecutado»."""
    nombres = {"empresa": "Compañía", "ambito": "Ámbito", "concepto": "Concepto de gasto", "placa": "Placa", "fecha": "Fecha",
               "documento": "Documento", "tercero": "Proveedor / tercero", "texto": "Texto original", "codigo_articulo": "Referencia (SISPO)",
               "articulo": "Artículo (SISPO)", "cantidad": "Cantidad", "valor": "Valor", "origen": "Clasificado por",
               "cuenta_nombre": "Cuenta", "cc_nombre": "Centro de coste", "llave": "ID del movimiento"}
    d = mov.copy()
    if "fecha" in d:
        d["fecha"] = pd.to_datetime(d["fecha"], errors="coerce").dt.date
    d = d.sort_values(["concepto", "fecha"], na_position="last") if {"concepto", "fecha"} <= set(d.columns) else d
    cols = [c for c in nombres if c in d.columns]
    r = t[[c for c in ("ambito", "concepto", "presupuesto", "ejecutado", "saldo", "pct", "estado", "movimientos") if c in t.columns]].rename(columns={
        "ambito": "Ámbito", "concepto": "Concepto de gasto", "presupuesto": "Presupuesto", "ejecutado": "Ejecutado", "saldo": "Saldo", "pct": "% ejecución",
        "estado": "Estado", "movimientos": "Movimientos"})
    hojas = {"Movimientos": d[cols].rename(columns=nombres)}
    if {"anio", "mes", "valor", "concepto"} <= set(d.columns) and len(d):
        d = d.assign(_per=pd.to_numeric(d["anio"], errors="coerce").fillna(0).astype(int) * 100 + pd.to_numeric(d["mes"], errors="coerce").fillna(0).astype(int))
        etiqueta = lambda per: f"{MESES_L[per % 100 - 1].capitalize()} {per // 100}" if 1 <= per % 100 <= 12 else f"Periodo {per}"
        for per in sorted(d["_per"].unique()):                  # una hoja por mes, en orden cronológico
            hojas[etiqueta(per)] = d[d["_per"] == per][cols].rename(columns=nombres)
        piv = d.pivot_table(index="concepto", columns="_per", values="valor", aggfunc="sum", fill_value=0.0)
        piv.columns = [etiqueta(c) for c in piv.columns]
        piv["Total"] = piv.sum(axis=1)
        hojas["Resumen por mes"] = piv.reset_index().rename(columns={"concepto": "Concepto de gasto"})
    hojas["Resumen por concepto"] = r
    return _excel_simple(hojas)


def _catalogos_cuentas(S) -> dict:
    """Conceptos del presupuesto de cada cuenta. Hoy solo la flota tiene catálogo propio (sus conceptos por placa); las demás
    cuentas no lo tienen, y L.sispo_cruzar_cuentas las trata como tales (concepto del maestro, sin cruzar de cuenta)."""
    return {L.SISPO_CUENTA: S.catalogo}


def _sispo_analizar(S, up, maestro: pd.DataFrame):
    """Lee el SISPO y arma todo el análisis de las siete cuentas (memoria por archivo: no se vuelve a leer en cada interacción)."""
    llave = (getattr(up, "name", ""), len(up.getvalue()), L.sispo_huella_maestro(maestro), tuple(S.catalogo).__hash__(),
             tuple(sorted(S.registro_placas.itertuples(index=False, name=None))).__hash__(), "cuentas7")
    memo = st.session_state.get("_mt_sispo_memo")
    if memo and memo[0] == llave:
        return memo[1]
    d = L.sispo_leer(up)
    lin, info = L.sispo_filtrar(d, L.SISPO_CUENTAS)
    res = dict(info=info, periodos=L.sispo_periodos(lin), empresa_probable=L.sispo_empresa_probable(lin) if len(lin) else None, original=d)
    lin = L.sispo_placas(lin, S.registro_placas)
    lin = L.sispo_cruzar_cuentas(lin, maestro, _catalogos_cuentas(S))
    res.update(lin=lin, n_referencias=int(d.loc[d["codigo"] != "", "codigo"].nunique()))
    st.session_state["_mt_sispo_memo"] = (llave, res)
    return res


def _sispo_panel(S):
    """Pantalla del proceso mensual: base maestra, SISPO del mes, validación, conciliación y aplicación. Solo administrador."""
    _exigir_admin(S.ctx)
    ctx = S.ctx
    with st.expander("Cargar SISPO del mes y reemplazar las salidas de inventario (administrador)", icon=":material/inventory_2:", expanded=True):
        resultado = st.session_state.pop("_mt_sispo_msg", None)          # resultado de la última aplicación (el rerun borraría un st.success directo)
        if resultado:
            (st.success if resultado[0] == "ok" else st.error)(resultado[1], icon=":material/check_circle:" if resultado[0] == "ok" else ":material/error:")
        st.caption("Reemplaza el movimiento genérico «Salida de inventario <mes>» de SAP por el detalle de SISPO: una línea por referencia "
                   "y placa, con el concepto de la base maestra. Se hace CUENTA POR CUENTA (" + "; ".join(L.SISPO_CUENTAS) + "): cada "
                   "salida de SISPO solo se asigna a la cuenta de su propio «Nombre Cta» y reemplaza únicamente los genéricos de SAP "
                   "de esa misma clase de coste. Cada mes se adjunta el SISPO de ESE mes; no se reutiliza el anterior. Una cuenta solo "
                   "se aplica si su valor total de SISPO coincide al centavo con lo que SAP tiene como genérico (se compara solo por valor, "
                   "no por centro de coste); las que no cuadran quedan sin tocar y se informan.")
        maestro = ctx.leer("SELECT codigo, articulo, concepto, fecha FROM mtto_sispo_ref")
        st.markdown("**1. Base maestra de referencias y conceptos**")
        if maestro.empty:
            st.warning("Aún no hay base maestra cargada: adjunta el Excel con «Codigo Articulo», «Articulo» y «CONCEPTO ASIGNADO».",
                       icon=":material/warning:")
        else:
            st.info(f"Base vigente: {len(maestro):,} referencias · {maestro['concepto'].nunique()} conceptos · última actualización "
                    f"{maestro['fecha'].max()}.".replace(",", "."), icon=":material/database:")
        up_m = st.file_uploader("Base maestra (.xlsx): adjúntala la primera vez o cuando quieras actualizarla", type=["xlsx"],
                                key="mt_sispo_up_maestro")
        if up_m:
            try:
                nuevo, im = L.sispo_maestro_leer(up_m)
            except Exception as e:
                st.error(f"No pude leer la base maestra: {e}")
                nuevo, im = None, None
            if nuevo is not None:
                if im["conflictos"]:
                    st.error(f"{len(im['conflictos'])} referencias tienen dos conceptos distintos en el archivo (no se adivina cuál): "
                             + ", ".join(im["conflictos"][:10]), icon=":material/error:")
                if im["sin_dato"]:
                    st.warning(f"{im['sin_dato']} filas sin código o sin concepto no se cargan.")
                c_m1, c_m2 = st.columns([1, 2], vertical_alignment="center")
                reemplazar = c_m2.checkbox("Reemplazar toda la base (borra las referencias que no estén en este archivo)", key="mt_sispo_reemplazar")
                if c_m1.button(f"Cargar {len(nuevo)} referencias", key="mt_sispo_cargar_maestro", disabled=bool(im["conflictos"]),
                               icon=":material/upload:"):
                    _sispo_cargar_maestro(ctx, nuevo, reemplazar, getattr(up_m, "name", "base maestra"))
                    st.success("Base maestra actualizada. Los periodos ya aplicados conservan sus conceptos hasta que los reasignes (abajo).")
                    st.rerun()
        st.markdown("**2. SISPO del mes**")
        up = st.file_uploader("SISPO del periodo que vas a procesar (.xlsx)", type=["xlsx"], key="mt_sispo_up")
        if up is None:
            st.info("Adjunta el SISPO del mes. No se reutiliza el de meses anteriores.", icon=":material/upload_file:")
        elif maestro.empty:
            st.error("Primero carga la base maestra.")
        else:
            try:
                a = _sispo_analizar(S, up, maestro[["codigo", "articulo", "concepto"]])
            except Exception as e:
                st.error(f"No pude leer el SISPO: {e}")
                a = None
            if a is not None:
                _sispo_resultado(S, a, getattr(up, "name", "SISPO"), maestro, up)
        _sispo_periodos(S, maestro)


def _sispo_operaciones(S, up, empresa: str, periodo) -> dict | None:
    """El MISMO archivo que se aplica en Mantenimiento alimenta también el SISPO de Operaciones (rubro 3.1): se carga una sola vez.
    Muestra lo que Operaciones tomaría y devuelve el análisis y su estado («nuevo», «igual» o «distinto»), o None si no se puede
    (el módulo Operaciones no está disponible, no hay periodo único o el archivo no se pudo leer). Las lecturas se guardan en memoria
    por archivo para no releer el Excel en cada interacción."""
    if up is None or not periodo:
        return None
    import importlib
    PO, falla = None, None
    for nombre in ("operaciones", "ppto_oper"):                 # import perezoso: Mantenimiento sigue funcionando sin Operaciones
        try:
            PO = importlib.import_module(nombre)
            break
        except ModuleNotFoundError as e:
            if e.name != nombre:                                # el módulo existe pero le falta una dependencia: se informa esa falla
                falla = e
                break
        except Exception as e:
            falla = e
            break
    if PO is None:
        st.warning("No pude cargar el módulo Operaciones" + (f" ({falla})" if falla else " (busqué «operaciones.py» y «ppto_oper.py» junto a este módulo)")
                   + ": el SISPO se aplicará solo en Mantenimiento.", icon=":material/warning:")
        return None
    try:
        PO._asegurar_tablas(S.ctx)
        llave = ("ops", getattr(up, "name", ""), len(up.getvalue()),
                 tuple(S.ctx.leer("SELECT codigo, item FROM oper_item_map").itertuples(index=False, name=None)).__hash__(),
                 tuple(S.ctx.leer("SELECT item FROM oper_ppto")["item"]).__hash__())
        memo = st.session_state.get("_mt_sispo_ops_memo")
        if memo and memo[0] == llave:
            a_ops = memo[1]
        else:
            a_ops = PO.sispo_analizar(S.ctx, up)
            st.session_state["_mt_sispo_ops_memo"] = (llave, a_ops)
        estado = PO.sispo_estado(S.ctx, empresa, periodo, a_ops)
    except Exception as e:
        st.error(f"No pude preparar el SISPO de Operaciones: {e}. Mantenimiento se puede aplicar igual; Operaciones queda sin este periodo.",
                 icon=":material/error:")
        return None
    with st.expander("Lo que tomará Operaciones (rubro 3.1) de este mismo archivo", icon=":material/sync_alt:", expanded=False):
        st.dataframe(a_ops["resumen"].rename(columns={"hoja": "Hoja", "lineas": "Salidas", "referencias": "Referencias", "cantidad": "Cantidad",
                                                      "valor": "Valor"}), hide_index=True, width="stretch", column_config={
            "Salidas": st.column_config.NumberColumn(format="%d"), "Referencias": st.column_config.NumberColumn(format="%d"),
            "Cantidad": st.column_config.NumberColumn(format="localized"), "Valor": st.column_config.NumberColumn(format="accounting")})
        st.dataframe(pd.DataFrame(a_ops["embudo"], columns=["Filtro de Operaciones (acumulado)", "Filas que quedan"]), hide_index=True, width="stretch",
                     column_config={"Filas que quedan": st.column_config.NumberColumn(format="%d")})
        st.caption("Operaciones usa solo las salidas de sus 5 centros de costo, grupos 5 y 7 (o ítems del presupuesto) y cuentas ÚTILES, PAPELERÍA Y OTROS u "
                   "OTROS, sin Staf Mantenimiento. Papelería va en cantidades. El reparto solo se aplica en Operaciones cuando la cuenta OTROS cuadra al "
                   "centavo con SAP (se ve en la tabla de validación de arriba).")
        if a_ops["huerfanos"]:
            st.warning("Valores de «Concepto» que no coinciden con ningún ítem del presupuesto de Operaciones (esas líneas se clasifican por la tabla "
                       "de equivalencias o por grupo): " + ", ".join(a_ops["huerfanos"]), icon=":material/rule:")
    texto = {"nuevo": "Operaciones: este periodo todavía no está guardado; se guardará al aplicar.",
             "igual": "Operaciones: ya tiene este periodo con el mismo archivo; no hay nada que cambiar allí.",
             "distinto": "Operaciones: ya tiene este periodo, pero distinto; se reemplazará (no se duplica)."}[estado]
    st.info(texto, icon=":material/sync_alt:")
    for b in a_ops["bloqueos"]:
        st.error(f"Operaciones no se guardará: {b}", icon=":material/block:")
    return dict(PO=PO, a=a_ops, estado=estado, bloqueos=a_ops["bloqueos"])


def _sispo_resultado(S, a: dict, archivo: str, maestro: pd.DataFrame, up=None):
    ctx = S.ctx
    lin, info = a["lin"], a["info"]
    if lin.empty:
        st.error(f"El SISPO no tiene salidas (costo con IVA negativo) en ninguna de las {len(L.SISPO_CUENTAS)} cuentas "
                 f"({info['total']} registros leídos).")
        return
    periodos = a["periodos"]
    c_e, c_p = st.columns(2)
    inferida = a["empresa_probable"]
    empresa = c_e.selectbox("Compañía del SISPO", list(EMPRESAS), index=list(EMPRESAS).index(inferida) if inferida else
                            list(EMPRESAS).index(S.sel[0]), key="mt_sispo_empresa",
                            help="Se propone según «EmpresaAgencia» del archivo; si no coincide, la aplicación no deja continuar.")
    bloqueos = []
    if inferida and inferida != empresa:
        bloqueos.append(f"El archivo parece de {inferida}, pero elegiste {empresa}.")
    if len(periodos) != 1:
        bloqueos.append("El SISPO debe traer un solo periodo de cierre; trae: " + (", ".join(map(str, periodos)) or "ninguno") + ".")
        periodo = None
    else:
        periodo = periodos[0]
        c_p.metric("Periodo (cierre)", f"{MESES_L[periodo % 100 - 1].capitalize()} {periodo // 100}")
        if periodo // 100 != int(S.cfg["anio"]):
            bloqueos.append(f"El periodo es de {periodo // 100}, pero el escenario elegido arriba es de {S.cfg['anio']}: cámbialo para procesarlo.")

    # ---- UNA conciliación por cuenta con la misma función (la regla es la de Flota)
    por = L.sispo_conciliar_cuentas(lin, S.sap_salidas, empresa, periodo)
    for v in por.values():                      # si el genérico está en varios sub-rubros, cada uno debe quedar igual que en SAP
        if v["estado"] == "Cuadra":
            t = L.sispo_cuadre_destinos(v["lin"], v["sap"])
            if (t["diferencia_c"] != 0).any():
                v["estado"], v["destinos"] = "No cuadra por sub-rubro", t
    problemas = L.sispo_validar_cuentas(lin, a["original"], por)
    n_pend_cta = {c: int((v["lin"]["concepto"] == L.PENDIENTE).sum()) for c, v in por.items()}
    reg = info["por_cuenta"].set_index("cuenta")["registros"]

    st.markdown("**3. Validación antes de aplicar**")
    filas = []
    for cta, v in por.items():
        c = v["conc"]
        filas.append({"Cuenta": cta, "Registros SISPO": int(reg.get(cta, 0)), "Salidas (costo < 0)": c["n_sispo"],
                      "Valor SISPO": c["sispo_c"] / 100, "Mov. SAP genéricos": c["n_sap"], "Valor SAP": c["sap_c"] / 100,
                      "Diferencia (SAP + SISPO)": c["diferencia_c"] / 100, "Pendientes": n_pend_cta[cta], "Estado": v["estado"]})
    t_cta = pd.DataFrame(filas)
    tot = {"Cuenta": "TOTAL", **{k: t_cta[k].sum() for k in t_cta.columns if k not in ("Cuenta", "Estado")}, "Estado": ""}
    t_cta = pd.concat([t_cta, pd.DataFrame([tot])], ignore_index=True)
    st.dataframe(t_cta, hide_index=True, width="stretch", column_config={
        "Registros SISPO": st.column_config.NumberColumn(format="%d"), "Salidas (costo < 0)": st.column_config.NumberColumn(format="%d"),
        "Mov. SAP genéricos": st.column_config.NumberColumn(format="%d"), "Pendientes": st.column_config.NumberColumn(format="%d"),
        "Valor SISPO": st.column_config.NumberColumn(format="accounting"), "Valor SAP": st.column_config.NumberColumn(format="accounting"),
        "Diferencia (SAP + SISPO)": st.column_config.NumberColumn(format="accounting")})
    st.caption("SAP carga estas salidas con signo positivo (costo) y SISPO las trae negativas (salida de inventario): cada cuenta concilia "
               "SAP contra −SISPO, en centavos exactos. Los valores quedan con el signo de SAP. El TOTAL suma todas las cuentas, "
               "incluso las que no se pueden aplicar.")
    st.dataframe(pd.DataFrame([
        ("Total de registros SISPO", f"{info['total']:,}".replace(",", ".")),
        (f"Registros de las {len(L.SISPO_CUENTAS)} cuentas", f"{info['en_cuenta']:,}".replace(",", ".")),
        ("Registros de otras cuentas (no se tocan)", (", ".join(f"{k}: {v}" for k, v in info["otras_cuentas"].items()) or "ninguno")),
        ("Registros con Costo Total Con IVA negativo", f"{info['negativos']:,}".replace(",", ".")),
        ("Referencias distintas (en las salidas filtradas)", f"{lin.loc[lin['codigo'] != '', 'codigo'].nunique():,}".replace(",", ".")),
        ("Referencias encontradas en la base maestra", f"{lin.loc[lin['concepto'] != L.PENDIENTE, 'codigo'].nunique():,}".replace(",", ".")),
        ("Referencias pendientes de clasificación", f"{lin.loc[(lin['concepto'] == L.PENDIENTE) & (lin['codigo'] != ''), 'codigo'].nunique():,} "
                                                    f"({sum(n_pend_cta.values()):,} salidas)".replace(",", ".")),
        (f"Placas identificadas ({L.SISPO_CUENTA})", f"{lin.loc[(lin['cuenta'] == L.SISPO_CUENTA) & (lin['placa'] != L.SIN_PLACA), 'placa'].nunique():,}"
                                                     .replace(",", ".")),
    ], columns=["Indicador", "Valor"]), hide_index=True, width="stretch")

    # ---- qué pasa con cada cuenta que no se puede aplicar
    for cta, v in por.items():
        if v["estado"] == "No cuadra":
            c = v["conc"]
            st.error(f"{cta}: el total no cuadra. SAP {_pesos2(c['sap_c'] / 100)} contra SISPO {_pesos2(-c['sispo_c'] / 100)}; diferencia "
                     f"{_pesos2(c['diferencia_c'] / 100)}. No se aplica hasta resolverlo; el genérico de SAP queda como está.", icon=":material/error:")
        elif v["estado"] == "No cuadra por sub-rubro":
            st.error(f"{cta}: el total cuadra, pero SAP tiene el genérico en más de un sub-rubro del presupuesto y el reparto por centro de coste "
                     "no deja cada uno igual; aplicarlo cambiaría el presupuesto ejecutado de esos sub-rubros. No se aplica.", icon=":material/error:")
            st.dataframe(v["destinos"].assign(sap=lambda t: t["sap_c"] / 100, sispo=lambda t: t["sispo_c"] / 100,
                                              diferencia=lambda t: t["diferencia_c"] / 100)[["codigo", "sub_rubro", "sap", "sispo", "diferencia"]],
                         hide_index=True, width="stretch")
        elif v["estado"] == "Sin genérico en SAP":
            st.warning(f"{cta}: SISPO trae {v['conc']['n_sispo']} salidas por {_pesos2(-v['conc']['sispo_c'] / 100)}, pero no encontré en SAP "
                       f"movimientos «Salida de inventario» de esa cuenta, compañía y mes. No hay con qué conciliar: no se aplica.", icon=":material/warning:")
        elif v["estado"] == "Sin salidas en SISPO" and (reg.get(cta, 0) or v["conc"]["n_sap"]):
            st.warning(f"{cta}: SISPO trae {int(reg.get(cta, 0))} registros pero ninguno con costo negativo"
                       + (f", y SAP tiene {_pesos2(v['conc']['sap_c'] / 100)} como genérico sin detalle. No se aplica." if v["conc"]["n_sap"] else ". No hay nada que reclasificar."),
                       icon=":material/warning:")
    for p_ in problemas:
        bloqueos.append(p_)
    faltan_col = L.sispo_sin_columnas_clave(a["original"])
    if faltan_col["sin_id"] or faltan_col["sin_codigo"]:         # el lector las tolera (para poder diagnosticar), pero aplicar exige ambas
        nombres = " ni ".join(n for n, f in (("«IdKardex»", faltan_col["sin_id"]), ("«Codigo Articulo»", faltan_col["sin_codigo"])) if f)
        bloqueos.append(f"El archivo no trae {nombres}: son obligatorias para aplicar (el IdKardex identifica cada salida y la referencia "
                        "es lo que se cruza con la base maestra). Descarga el SISPO completo e inténtalo de nuevo.")
    if info["con_mas_de_2_decimales"]:
        bloqueos.append(f"{info['con_mas_de_2_decimales']} valores del SISPO traen más de 2 decimales: no se redondean, revisa el archivo.")
    if info["repetidos"]:
        bloqueos.append(f"{info['repetidos']} IdKardex aparecen repetidos: revisa el archivo para no duplicar movimientos.")

    with st.expander("Distribución de valores por cuenta y concepto", icon=":material/stacked_bar_chart:", expanded=True):
        st.dataframe(L.sispo_distribucion(lin, por_cuenta=True), hide_index=True, width="stretch", column_config={
            "cuenta": "Cuenta", "concepto": "Concepto", "lineas": st.column_config.NumberColumn("Salidas", format="%d"),
            "referencias": st.column_config.NumberColumn("Referencias", format="%d"),
            "placas": st.column_config.NumberColumn("Placas", format="%d"),
            "valor": st.column_config.NumberColumn("Valor SISPO", format="accounting")})
    pend = L.sispo_pendientes(lin)
    if len(pend):
        pv = pend["Valor SISPO"].sum()
        st.warning(f"{len(pend)} referencias ({sum(n_pend_cta.values())} salidas, {_pesos2(pv)}) quedan como PENDIENTE DE CLASIFICACIÓN: no están "
                   "en la base maestra o su concepto pertenece a otra cuenta; no se asignan por aproximación. El Excel trae las mismas columnas "
                   "de la base maestra (más la cuenta): complétalo con el concepto, cárgalo arriba y usa «Reasignar conceptos».",
                   icon=":material/rule:")
        st.download_button("Descargar referencias pendientes", _excel_simple({"Pendientes": pend}), f"referencias_pendientes_{periodo or ''}.xlsx",
                           icon=":material/download:", key="mt_sispo_dl_pend")
    sinp = lin[(lin["cuenta"] == L.SISPO_CUENTA) & (lin["placa"] == L.SIN_PLACA)]
    if len(sinp):
        with st.expander(f"Salidas de {L.SISPO_CUENTA} sin placa identificada ({len(sinp)})", icon=":material/help:"):
            g = sinp.groupby(["centro_costo", "placa_estado"]).agg(salidas=("id_kardex", "size"), valor=("costo_con_iva", "sum")).reset_index()
            st.dataframe(g, hide_index=True, width="stretch", column_config={"valor": st.column_config.NumberColumn(format="accounting")})
            st.caption("Los códigos de centro de costos (DB…) no se toman como placa. Se conservan en «SIN PLACA IDENTIFICADA» y puedes "
                       "asignarles la placa en «Reclasificación de movimientos».")

    # ---- aplicar
    st.markdown("**4. Aplicar**")
    huella_m = L.sispo_huella_maestro(maestro[["codigo", "concepto"]])
    previo = ctx.leer("SELECT cuenta, huella, huella_maestro, fecha, usuario, estado FROM mtto_sispo_periodo WHERE empresa=? AND periodo=?",
                      (empresa, int(periodo or 0)))
    previo = {(L.cuenta_de(r.cuenta) or L.SISPO_CUENTA): r for r in previo.itertuples()}
    por_aplicar, plan = [], []
    for cta, v in por.items():
        if v["estado"] != "Cuadra":
            continue
        h, p0 = L.sispo_huella(v["lin"]), previo.get(cta)
        if p0 is not None and p0.huella == h and p0.huella_maestro == huella_m and p0.estado != ESTADO_APLICACION_SIN_SAP_HEREDADA:
            plan.append((cta, "Ya aplicada con este mismo archivo y base maestra: no hay nada que cambiar", f"{p0.fecha} · {p0.usuario}"))
        else:
            por_aplicar.append((cta, h))
            accion = "Aplicar" if p0 is None else (
                "Reprocesar: la aplicación anterior fue sin conciliar con SAP; se reemplaza (no se duplica)"
                if p0.estado == ESTADO_APLICACION_SIN_SAP_HEREDADA else "Reprocesar: reemplaza lo aplicado antes (no se duplica)")
            plan.append((cta, accion, "" if p0 is None else f"{p0.fecha} · {p0.usuario}"))
    if plan:
        st.dataframe(pd.DataFrame(plan, columns=["Cuenta", "Acción", "Aplicada antes"]), hide_index=True, width="stretch")
    else:
        st.info("Ninguna cuenta está lista para aplicar (ver los avisos de arriba).", icon=":material/info:")
    ops = None if bloqueos else _sispo_operaciones(S, up, empresa, periodo)       # mismo archivo, mismo botón
    for b in bloqueos:
        st.error(b, icon=":material/block:")
    pend_a_aplicar = sum(n_pend_cta[c] for c, _ in por_aplicar)
    ok_pend = True
    if pend_a_aplicar:
        ok_pend = st.checkbox(f"Confirmo aplicar dejando {pend_a_aplicar} salidas como PENDIENTE DE CLASIFICACIÓN", key="mt_sispo_ok_pend")
    ops_hacer = bool(ops) and ops["estado"] != "igual" and not ops["bloqueos"]
    n_mtto = len(por_aplicar) if ok_pend else 0              # sin confirmar los pendientes, Mantenimiento no se aplica
    etiqueta = (f"Aplicar reclasificación ({n_mtto} cuenta{'s' if n_mtto != 1 else ''}" + (" + Operaciones)" if ops_hacer else ")"))
    listo = (bool(por_aplicar) and ok_pend) or ops_hacer
    if st.button(etiqueta, key="mt_sispo_aplicar", type="primary", icon=":material/done_all:", disabled=bool(bloqueos) or not listo):
        hechas = []
        if ok_pend:                                          # sin la confirmación de pendientes no se toca Mantenimiento
            for cta, h in por_aplicar:
                v = por[cta]
                gen = " / ".join(sorted({str(t) for t in v["sap"]["texto"]}))
                _sispo_aplicar(ctx, empresa, periodo, cta, v["lin"], v["sap"], v["conc"], archivo, h, huella_m, gen)
                hechas.append(cta)
        falla_ops = None
        if ops_hacer:
            try:
                n_ops = ops["PO"].sispo_guardar(ctx, empresa, periodo, ops["a"], archivo)
                leido = ctx.leer("SELECT n_lineas FROM oper_sispo_periodo WHERE empresa=? AND periodo=?", (empresa, int(periodo)))
                if leido.empty:                                  # comprobación de lectura: que de verdad quedó en la base
                    falla_ops = "se ejecutó el guardado pero el periodo no aparece en la tabla de Operaciones."
                else:
                    hechas.append(f"Operaciones ({n_ops} líneas, verificado en la base)")
            except Exception as e:
                falla_ops = str(e)
        if falla_ops:
            st.session_state["_mt_sispo_msg"] = ("error", ("Aplicado en Mantenimiento: " + ("; ".join(hechas) if hechas else "nada nuevo")
                                                           + f". Operaciones NO se guardó: {falla_ops}"))
        else:
            st.session_state["_mt_sispo_msg"] = ("ok", "Aplicado en: " + "; ".join(hechas) + ". Flota, Maquinaria y Reparaciones locativas se ven en sus "
                                                 "pestañas; el detalle de todas las cuentas está en «Periodos aplicados».")
        st.rerun()


def _sispo_periodos(S, maestro: pd.DataFrame):
    """Periodos y cuentas ya aplicados: estado vigente de la conciliación, trazabilidad, reasignación de conceptos y reversión."""
    ctx = S.ctx
    ap = ctx.leer("SELECT empresa, periodo, cuenta, archivo, usuario, fecha, n_lineas, n_pendientes, valor_sispo, valor_sap, diferencia, estado, "
                  "n_reemplazados FROM mtto_sispo_periodo ORDER BY periodo DESC, empresa, cuenta")
    st.markdown("**5. Periodos aplicados**")
    if ap.empty:
        st.caption("Aún no se ha aplicado ningún periodo.")
        return
    ap["cuenta"] = ap["cuenta"].map(lambda x: L.cuenta_de(x) or L.SISPO_CUENTA)
    vigente = []
    for r in ap.itertuples():
        sap = L.sispo_sap_generico(S.sap_salidas, r.empresa, r.periodo, r.cuenta) if r.periodo // 100 == int(S.cfg["anio"]) else None
        actual = None if sap is None else round(float(pd.to_numeric(sap["valor"], errors="coerce").fillna(0).sum()), 2) if len(sap) else 0.0
        # los movimientos reemplazados ya no están en las pantallas, pero sap_salidas es el original de SAP
        vigente.append("Año distinto al escenario: no verificable" if actual is None else
                       "⚠ Aplicado sin conciliar con SAP (prueba anterior): revierte la cuenta y vuelve a aplicarla"
                       if r.estado == ESTADO_APLICACION_SIN_SAP_HEREDADA else
                       ("✔ Vigente" if abs(actual - r.valor_sap) < 0.005 else f"⚠ SAP cambió: hoy {_pesos2(actual)}"))
    t = ap.assign(**{"Conciliación vigente": vigente})
    st.dataframe(t, hide_index=True, width="stretch", column_config={
        "empresa": "Compañía", "periodo": st.column_config.NumberColumn("Periodo", format="%d"), "cuenta": "Cuenta", "archivo": "Archivo",
        "usuario": "Usuario", "fecha": "Aplicado", "n_lineas": st.column_config.NumberColumn("Salidas", format="%d"),
        "n_pendientes": st.column_config.NumberColumn("Pendientes", format="%d"),
        "valor_sispo": st.column_config.NumberColumn("Valor SISPO", format="accounting"),
        "valor_sap": st.column_config.NumberColumn("Valor SAP", format="accounting"),
        "diferencia": st.column_config.NumberColumn("Diferencia", format="accounting"), "estado": "Estado",
        "n_reemplazados": st.column_config.NumberColumn("Mov. SAP reemplazados", format="%d")})
    st.download_button("Descargar histórico SISPO (todos los periodos aplicados)", lambda: _excel_historico_sispo(ctx),
                       "historico_sispo.xlsx", icon=":material/download:", key="mt_sispo_dl_historico",
                       help="Excel con todo el SISPO aplicado hasta hoy (todas las compañías, periodos y cuentas). Se arma desde la base en el "
                            "momento de descargar, así que cada periodo que apliques queda incluido.")
    opciones = [f"{r.empresa} · {r.periodo} · {r.cuenta}" for r in ap.itertuples()]
    elegido = st.selectbox("Periodo y cuenta a gestionar", opciones, key="mt_sispo_gestion")
    emp, per, cta = elegido.split(" · ", 2)
    per = int(per)
    lin = ctx.leer("SELECT * FROM mtto_sispo_lineas WHERE empresa=? AND periodo=?", (emp, per))
    lin["cuenta_nombre"] = lin["cuenta_nombre"].map(lambda x: L.cuenta_de(x) or L.SISPO_CUENTA)
    gen = ctx.leer("SELECT cuenta, generico FROM mtto_sispo_periodo WHERE empresa=? AND periodo=?", (emp, per))
    gen = {(L.cuenta_de(r.cuenta) or L.SISPO_CUENTA): str(r.generico) for r in gen.itertuples()}
    traza = pd.concat([L.sispo_filas_trazabilidad(g.assign(cuenta=c), per, "Conciliado", gen.get(c, "")) for c, g in lin.groupby("cuenta_nombre")],
                      ignore_index=True) if len(lin) else pd.DataFrame()
    lin_c = lin[lin["cuenta_nombre"] == cta]
    with st.expander(f"Detalle aplicado · {cta} · {per}", icon=":material/stacked_bar_chart:", expanded=False):
        dist = L.sispo_distribucion(lin_c.assign(centavos=lin_c["costo_con_iva"].map(L.centavos), cuenta=cta))
        st.dataframe(dist.assign(valor=-dist["valor"]), hide_index=True, width="stretch", column_config={
            "concepto": "Concepto", "lineas": st.column_config.NumberColumn("Salidas", format="%d"),
            "referencias": st.column_config.NumberColumn("Referencias", format="%d"), "placas": st.column_config.NumberColumn("Placas", format="%d"),
            "valor": st.column_config.NumberColumn("Valor (signo SAP)", format="accounting")})
    c1, c2, c3 = st.columns(3)
    c1.download_button("Descargar trazabilidad (todas las cuentas del periodo)", _excel_simple({"Trazabilidad": traza}),
                       f"trazabilidad_sispo_{emp}_{per}.xlsx", icon=":material/download:", key="mt_sispo_dl_traza")
    if c2.button("Reasignar conceptos con la base maestra vigente", key="mt_sispo_reasignar", icon=":material/sync:",
                 help="Útil después de actualizar la base maestra: vuelve a cruzar las salidas ya aplicadas de esta cuenta sin necesitar el SISPO otra vez."):
        cambios, pend = _sispo_reasignar(ctx, emp, per, cta, maestro[["codigo", "articulo", "concepto"]], _catalogos_cuentas(S))
        st.success(f"{cambios} salidas cambiaron de concepto; quedan {pend} pendientes. Los valores no cambian.")
        st.rerun()
    ok = c3.checkbox("Confirmo revertir esta cuenta y periodo", key="mt_sispo_ok_revertir")
    if c3.button("Revertir cuenta y periodo", key="mt_sispo_revertir", icon=":material/undo:", disabled=not ok,
                 help="Quita el detalle de SISPO de esa cuenta y periodo y vuelven sus movimientos genéricos de SAP. Las demás cuentas no se tocan."):
        _sispo_revertir(ctx, emp, per, cta)
        st.success("Cuenta revertida: vuelven sus movimientos genéricos de SAP.")
        st.rerun()


# ================================================================== rubro completo: Mantenimiento y Contenedores
def _rubro(S, empresa: str, detalle_flota: bool = False):
    """Presupuesto y ejecución (sin raw) de todos los sub-rubros del rubro de la flota para UNA compañía: es la fuente común de
    la pestaña Mantenimiento, de Contenedores y del soporte de ejecución. Respeta la gerencia de quien consulta y excluye los
    movimientos eliminados. Se calcula una sola vez por pantalla.
    `detalle_flota=True` (solo el soporte de ejecución) reemplaza también los genéricos de la cuenta de la flota por el detalle de
    SISPO; las pantallas no lo piden porque la pestaña Flota ya trae ese detalle por placa y no debe contarse dos veces."""
    ck = (empresa, bool(detalle_flota))
    if ck in S.rubro_cache:
        return S.rubro_cache[ck]
    d, cfg = S.datos_e[empresa], S.cfg
    flota = cargas.norm(L.SUBRUBRO_FLOTA)
    staf = d.ejec["_staf_otros"].fillna(False).astype(bool) if "_staf_otros" in d.ejec else pd.Series(False, index=d.ejec.index)
    # los códigos del rubro salen del presupuesto y de los movimientos propios de flota; los OTROS/Staf Mantenimiento no definen
    # códigos (traerían todo lo demás de su código), se agregan fila a fila
    codigos = set(d.ppto_base.loc[d.ppto_base["_sub"] == flota, "codigo"]) | set(d.ejec.loc[(d.ejec["_sub"] == flota) & ~staf, "codigo"])
    ppto = d.ppto_base[d.ppto_base["codigo"].isin(codigos)]
    ejec = d.ejec[d.ejec["codigo"].isin(codigos) | staf].copy()
    if len(ejec) and (len(S.sispo_otras[0]) or (detalle_flota and len(S.sispo_flota[0]))):
        ejec["llave"] = L.llave_movimiento(ejec.assign(empresa=empresa))
        if detalle_flota:                               # soporte: cada centro de costo que cuadre con SISPO sale con su detalle (flota incluida)
            for fuente in (S.sispo_otras, S.sispo_flota):
                ejec = L.sispo_superponer(ejec, *fuente, empresa, parcial=True)
        else:                                           # cuentas reclasificadas con SISPO (Maquinaria, Reparaciones locativas…): detalle en vez del genérico
            ejec = L.sispo_superponer(ejec, *S.sispo_otras, empresa)
    if not S.ver_todo:                                       # quien no ve todo, solo lo suyo
        ppto = ppto[asignacion.columna(ppto, S.asig) == S.ctx.usuario] if len(ppto) else ppto
        ejec = _con_gerencia(ejec, S.asig)
        ejec = ejec[ejec["gerencia"] == S.ctx.usuario]
    if len(ejec):
        ejec["llave"] = L.llave_movimiento(ejec.assign(empresa=empresa))
        ejec = ejec[~ejec["llave"].isin(set(S.excluidos["llave"]))]
    S.rubro_cache[ck] = (ppto, ejec)
    return S.rubro_cache[ck]


def _rubro_sel(S):
    """El rubro de las compañías elegidas, sumado."""
    partes = [_rubro(S, e) for e in S.sel]
    return (pd.concat([p for p, _ in partes], ignore_index=True), pd.concat([x for _, x in partes], ignore_index=True))


def _fila_total_df(t: pd.DataFrame, col: str, nombre: str = "Total") -> pd.DataFrame:
    """Agrega una fila de total (presupuesto, ejecutado y movimientos sumados; saldo, % y estado recalculados)."""
    base = t[[col, "presupuesto", "ejecutado", "movimientos"]].copy()
    base.loc[len(base)] = {col: nombre, "presupuesto": base["presupuesto"].sum(), "ejecutado": base["ejecutado"].sum(),
                           "movimientos": int(base["movimientos"].sum())}
    return L.con_metricas(base)[[col, "presupuesto", "ejecutado", "saldo", "pct", "estado", "movimientos"]]


def _control_staf(S):
    """Control (solo administrador) de las líneas de SAP con clase de coste OTROS y objeto «Staf Mantenimiento», que se suman a
    «Flota y Equipo de transporte» y por tanto salen en el soporte de ejecución. Muestra, por compañía y mes, cuántas hay y cuánto
    suman; y si no hay ninguna, qué se encontró parecido."""
    if not S.ctx.es_admin:
        return
    filas, parecidas = [], []
    for e in S.sel:
        d = S.datos_e[e]
        x = d.ejec[d.ejec["_staf_otros"].fillna(False).astype(bool)] if "_staf_otros" in d.ejec else d.ejec.iloc[0:0]
        en_informe = _rubro(S, e)[1]
        en_informe = en_informe[en_informe["_staf_otros"].fillna(False).astype(bool)] if "_staf_otros" in en_informe and len(en_informe) else en_informe.iloc[0:0]
        for mes, g in x.groupby("mes"):
            filas.append({"Compañía": e, "Mes": MESES_L[int(mes) - 1].capitalize() if 1 <= int(mes) <= 12 else str(mes),
                          "Líneas en SAP": len(g), "Valor en SAP": float(g["valor"].sum()),
                          "Líneas en el informe": int((en_informe["mes"] == mes).sum()) if len(en_informe) else 0,
                          "Valor en el informe": float(en_informe.loc[en_informe["mes"] == mes, "valor"].sum()) if len(en_informe) else 0.0})
        if x.empty:
            cerca = d.ejec[d.ejec["cc_nombre"].map(L._canon).str.contains("STAF", regex=False)] if "cc_nombre" in d.ejec and len(d.ejec) else d.ejec.iloc[0:0]
            for (cl, ob), g in cerca.groupby([cerca["cuenta_nombre"].fillna(""), cerca["cc_nombre"].fillna("")]):
                parecidas.append({"Compañía": e, "Denom.clase de coste": cl, "Denominación del objeto": ob, "Líneas": len(g), "Valor": float(g["valor"].sum())})
    with st.expander("Control: OTROS · Staf Mantenimiento incluidos en Flota", icon=":material/fact_check:"):
        st.caption("Líneas de SAP con «Denom.clase de coste» = OTROS y «Denominación del objeto» = Staf Mantenimiento. Se suman a «Flota y "
                   "Equipo de transporte» (Resumen, pestaña Flota y soporte de ejecución). No traen placa ni presupuesto.")
        if filas:
            df = pd.DataFrame(filas)
            st.dataframe(df, hide_index=True, width="stretch", height=40 + 35 * (len(df) + 1),
                         column_config={c: st.column_config.NumberColumn(c, format="accounting") for c in ("Valor en SAP", "Valor en el informe")})
        else:
            st.warning("No se encontró ninguna línea con clase de coste OTROS y objeto Staf Mantenimiento en la ejecución cargada.")
        if parecidas:
            st.caption("Líneas con «Staf» en el objeto pero con otra clase de coste (no se incluyen):")
            st.dataframe(pd.DataFrame(parecidas), hide_index=True, width="stretch", height=40 + 35 * (len(parecidas) + 1))
        if not filas:                                   # diagnóstico: qué trae realmente la base de ejecución cargada
            st.caption("Diagnóstico de la ejecución cargada (sin filtros de código ni de rubro):")
            meses_, otros_ = [], []
            for e in S.sel:
                meses_.append(S.ctx.leer("SELECT empresa, mes, COUNT(*) AS lineas, SUM(valor) AS valor FROM ejecucion WHERE anio=? AND empresa=? "
                                         "GROUP BY empresa, mes ORDER BY mes", (int(S.cfg["anio"]), e)))
                otros_.append(S.ctx.leer(
                    "SELECT empresa, mes, cuenta_nombre, sub_rubro, cc_nombre, codigo, COUNT(*) AS lineas, SUM(valor) AS valor FROM ejecucion "
                    "WHERE anio=? AND empresa=? AND (UPPER(COALESCE(cuenta_nombre,'')) LIKE '%OTRO%' OR UPPER(COALESCE(sub_rubro,'')) LIKE '%OTRO%') "
                    "GROUP BY empresa, mes, cuenta_nombre, sub_rubro, cc_nombre, codigo ORDER BY mes, valor DESC", (int(S.cfg["anio"]), e)))
            m_ = pd.concat(meses_, ignore_index=True)
            st.caption("Líneas de SAP cargadas por compañía y mes:")
            st.dataframe(m_, hide_index=True, width="stretch", height=40 + 35 * (len(m_) + 1))
            o_ = pd.concat(otros_, ignore_index=True)
            if o_.empty:
                st.caption("No hay ninguna línea cuya clase de coste o sub-rubro contenga «OTRO».")
            else:
                st.caption("Líneas cuya clase de coste o sub-rubro contiene «OTRO» (con su objeto y código):")
                st.dataframe(o_.head(60), hide_index=True, width="stretch", height=380)


def _descarga_soporte(S):
    """Soporte de ejecución (Excel) para cualquier usuario: presupuesto de todos los meses y ejecución acumulada en la primera
    hoja; en las demás, los movimientos."""
    with st.container(key="mt_card_soporte"):
        titulo_tarjeta("Soporte de ejecución de mantenimiento", "Excel por compañía · «Presupuesto»: presupuesto de todos los meses y "
                       "ejecución acumulada · demás hojas: movimientos")
        c1, c2 = st.columns([1, 1.4], vertical_alignment="bottom")
        mes_sop = c1.selectbox("Mes del soporte de ejecución", list(range(S.m0, S.m1 + 1)), index=S.ult - S.m0,
                               key=f"mt_mes_soporte_{S.escenario}", format_func=lambda m: MESES_L[m - 1].capitalize(),
                               help="La hoja «Presupuesto» acumula la ejecución desde el inicio del escenario hasta este "
                                    "mes. Es independiente de «Ejecución hasta».")
        todos = c2.radio("Movimientos de cada sub-rubro", ["Solo el mes elegido", "Todos los meses hasta el elegido"], horizontal=True,
                         key=f"mt_detalle_soporte_{S.escenario}",
                         help="«Todos los meses…» agrega a cada hoja de sub-rubro los movimientos de los meses anteriores (p. ej. septiembre y "
                              "octubre) con una columna «Mes».") == "Todos los meses hasta el elegido"
        cols_dl = st.columns(len(S.sel))
        for cdl, e in zip(cols_dl, S.sel):
            cdl.download_button(f"Descargar soporte de ejecución · {e}", lambda e=e: _soporte(S, e, mes_sop, todos),
                                f"{mes_sop:02d}_{e}_Soporte_Ejecucion_mantenimiento.xlsx", icon=":material/download:",
                                key=f"mt_soporte_{e}",
                                help="Hoja «Presupuesto»: presupuesto de todos los meses y ejecución acumulada hasta el mes elegido, "
                                     "por sub-rubro. "
                                     "Una hoja por sub-rubro con los movimientos de SAP.")


def _tab_mantenimiento(S):
    """Resumen ejecutivo del rubro: sub-rubros, compañías y mes a mes, sin ningún movimiento a la vista."""
    ppto, ejec = _rubro_sel(S)
    t = L.tabla_subrubros(ppto, ejec, S.cfg, S.hasta_mes)
    if t.empty:
        st.info("El rubro no tiene presupuesto ni ejecución en este escenario.")
        return
    tp, te, tm = float(t["presupuesto"].sum()), float(t["ejecutado"].sum()), int(t["movimientos"].sum())
    _kpis(tp, te, tm, S.hasta, f"Escenario {S.escenario} · rubro completo")
    if S.placa:
        st.caption(f"El filtro de placa ({S.placa}) no cambia el rubro completo de esta pestaña; solo afecta la tabla de ámbitos de abajo.")
    with st.container(key="mt_card_rubro"):
        titulo_tarjeta("Presupuesto y ejecución por sub-rubro", f"{S.escenario} · presupuesto del escenario completo · ejecutado {S.hasta}")
        _tabla(_fila_total_df(t, "sub_rubro"), 40 + 35 * (len(t) + 1))
    with st.container(key="mt_card_rubro_graf"):
        titulo_tarjeta("Dónde se concentra el presupuesto", "Presupuesto vs ejecutado por sub-rubro (millones de pesos)")
        _barras_ppto_ejec(t.rename(columns={"sub_rubro": "concepto"}), "Sub-rubros del rubro Reparación y Conservación.")
    if S.multi:
        filas = []
        for e in S.sel:
            pe, ee = _rubro(S, e)
            te_ = L.tabla_subrubros(pe, ee, S.cfg, S.hasta_mes)
            filas.append(dict(empresa=e, presupuesto=float(te_["presupuesto"].sum()), ejecutado=float(te_["ejecutado"].sum()),
                              movimientos=int(te_["movimientos"].sum())))
        with st.container(key="mt_card_rubro_emp"):
            titulo_tarjeta("Por compañía", f"{S.escenario} · ejecutado {S.hasta}")
            tc = _fila_total_df(pd.DataFrame(filas), "empresa", "Consolidado")
            _tabla(tc, 40 + 35 * len(tc))
    with st.container(key="mt_card_rubro_mes"):
        titulo_tarjeta("Mes a mes", "Presupuesto vs ejecutado del rubro (millones de pesos)")
        _grafico_ts_mensual(S, ppto, ejec)
    _control_staf(S)
    _descarga_soporte(S)
    with st.expander("Flota y Equipo de transporte por ámbito (Placas · Generales · Tecnomecánica y SOAT)", icon=":material/table:"):
        _tabla(S.vistas["resumen"], 190)
        st.caption("Placas = presupuesto por vehículo. Generales = conceptos sin placa. Tecnomecánica y SOAT = presupuesto mensual por "
                   "placa. El saldo es presupuesto − ejecutado.")


def _tab_contenedores(S):
    """Vista simple de Contenedores: cuatro indicadores, tabla mes a mes de presupuesto vs ejecutado y un gráfico."""
    ppto, ejec = _rubro_sel(S)
    es_c = lambda df: df["sub_rubro"].fillna("").astype(str).map(L.norm_texto).str.contains("CONTENEDOR")
    pc, ec = ppto[es_c(ppto)].assign(sub_rubro="Contenedores"), ejec[es_c(ejec)].assign(sub_rubro="Contenedores")
    pc = pc[(pc["mes"] >= S.m0) & (pc["mes"] <= S.m1)]
    if pc.empty and ec.empty:
        st.info(f"No hay presupuesto ni ejecución de Contenedores en {S.escenario}.", icon=":material/info:")
        return
    t = L.tabla_mensual(pc, ec, S.cfg, S.hasta_mes, "Contenedores")
    tot = t[t["mes"] == 0].iloc[0]
    n_mov = int(((ec["mes"] >= S.m0) & (ec["mes"] <= S.hasta_mes)).sum()) if len(ec) else 0
    _kpis(float(tot["presupuesto"]), float(tot["ejecutado"]), n_mov, S.hasta, f"Escenario {S.escenario}")
    vista = t.assign(Mes=t["mes"].map(lambda m: "Total" if m == 0 else f"{MESES_L[m - 1].capitalize()} {S.cfg['anio']}"))
    with st.container(key="mt_card_cont_tabla"):
        titulo_tarjeta("Contenedores · presupuesto vs ejecutado", f"{S.escenario} · ejecutado {S.hasta}")
        st.dataframe(vista[["Mes", "presupuesto", "ejecutado", "saldo", "pct"]], hide_index=True, width="stretch",
                     height=40 + 35 * len(vista), column_config={
                         "Mes": "Mes",
                         "presupuesto": st.column_config.NumberColumn("Presupuesto asignado", format="accounting"),
                         "ejecutado": st.column_config.NumberColumn("Valor ejecutado", format="accounting"),
                         "saldo": st.column_config.NumberColumn("Saldo disponible", format="accounting"),
                         "pct": st.column_config.NumberColumn("% de ejecución", format="percent")})
    with st.container(key="mt_card_cont_graf"):
        titulo_tarjeta("Mes a mes", "Presupuesto vs valor ejecutado (millones de pesos)")
        _grafico_ts_mensual(S, pc, ec)


# ================================================================== carga de datos base (administrador)
def _columna(df: pd.DataFrame, *claves):
    for c in df.columns:
        n = L.sin_tildes(c).upper()
        if any(k in n for k in claves):
            return c
    return None


def _notas_encabezado(archivo, hoja: str) -> dict:
    """Notas (comentarios) de las celdas del encabezado de una hoja: forman parte del texto original del presupuesto."""
    try:
        archivo.seek(0)
        ws = openpyxl.load_workbook(archivo)[hoja]
        archivo.seek(0)
        return {" ".join(str(c.value).split()): " ".join(c.comment.text.split()) for c in ws[1] if c.comment and c.value}
    except Exception:
        return {}


def _admin_datos(ctx):
    _exigir_admin(ctx)
    oficial = set(ctx.leer("SELECT placa FROM mtto_placas")["placa"])
    with st.expander("Datos base · placas, presupuesto de flota, histórico y Tecnomecánica/SOAT (administrador)",
                     icon=":material/settings:"):
        _admin_placas(ctx)
        st.divider()
        st.markdown("**Por compañía.** El presupuesto, el histórico y Tecnomecánica/SOAT se cargan por separado para cada compañía; "
                    "cargar una no toca los datos de la otra.")
        pestanas = st.tabs([f"{e} · {EMPRESAS[e]}" for e in EMPRESAS])
        for tab, e in zip(pestanas, EMPRESAS):
            with tab:
                _admin_ppto_flota(ctx, e, oficial)
                _admin_historico(ctx, e)
                _admin_ts(ctx, e, oficial)


def _admin_placas(ctx):
    _exigir_admin(ctx)
    st.markdown("**1. Listado oficial de placas** (único para todas las compañías)  \nExcel con las columnas PLACA, NÚMERO INTERNO y "
                "TIPO DE VEHÍCULO. Reemplaza el listado actual; las placas repetidas se unifican.")
    f = st.file_uploader("Listado de placas (.xlsx)", type=["xlsx"], key="mt_up_placas")
    if f and st.button("Guardar listado de placas", key="mt_guardar_placas"):
        df = pd.read_excel(f, dtype=str)
        c_pl, c_in, c_ti = _columna(df, "PLACA"), _columna(df, "INTERNO"), _columna(df, "TIPO")
        if not c_pl:
            st.error("No encuentro la columna PLACA.")
            return
        t = pd.DataFrame({"placa": df[c_pl].map(L.norm_placa),
                          "numero_interno": df[c_in].str.strip() if c_in else None,
                          "tipo": df[c_ti].str.strip() if c_ti else None})
        t = t[t["placa"] != ""]
        n_filas = len(t)
        t = t.groupby("placa", sort=False).agg("first").reset_index()
        ctx.con.execute("DELETE FROM mtto_placas")
        _insertar_por_lotes(ctx, "mtto_placas(placa, numero_interno, tipo)",
                            [(r.placa, r.numero_interno if pd.notna(r.numero_interno) else "",
                              r.tipo if pd.notna(r.tipo) else "") for r in t.itertuples(index=False)])
        ctx.con.commit()
        ctx.registrar_carga(ctx.usuario, "Placas flota", f.name, f"{len(t)} placas únicas de {n_filas} filas")
        st.success(f"{len(t)} placas únicas guardadas ({n_filas - len(t)} filas repetidas unificadas).")
        st.rerun()


def _admin_ppto_flota(ctx, empresa, oficial):
    _exigir_admin(ctx)
    st.markdown(f"**2. Presupuesto de la flota · {empresa}**  \nExcel con una hoja por escenario (cada fila un "
                "vehículo y cada columna un concepto; las filas sin placa son conceptos generales). Se transforma a "
                "formato largo, conserva el texto original (encabezado y notas de celda) y solo se guarda si la suma "
                f"coincide con el total del archivo. Reemplaza el escenario cargado para {empresa}.")
    g = st.file_uploader("Presupuesto de la flota (.xlsx)", type=["xlsx"], key=f"mt_up_ppto_{empresa}")
    if not g:
        return
    xl = pd.ExcelFile(g)
    listos = {}
    for esc, cfg in L.ESCENARIOS.items():
        opciones = ["(no cargar)"] + xl.sheet_names
        hoja = st.selectbox(f"Hoja para «{esc}»", opciones, key=f"mt_hoja_{esc}_{empresa}",
                            index=opciones.index(cfg["hoja"]) if cfg["hoja"] in opciones else 0)
        if hoja == "(no cargar)":
            continue
        try:
            largo, info = L.leer_presupuesto_hoja(xl.parse(hoja), _notas_encabezado(g, hoja))
        except ValueError as e:
            st.error(f"{esc}: {e}")
            continue
        ok = abs(info["dif_total"]) < 1 and not info["filas_que_no_cuadran"] and not info["conceptos_que_no_cuadran"]
        (st.success if ok else st.error)(
            f"**{esc}** (hoja {hoja}): {info['n_placas']} vehículos, {info['n_globales']} conceptos generales, "
            f"{info['n_conceptos']} columnas de concepto. Total del archivo {E.pesos(info['total_archivo'] or info['total_filas'])} · "
            f"transformado {E.pesos(info['total_transformado'])}.")
        if info["duplicadas"]:
            st.warning("Placas repetidas en el archivo: " + ", ".join(info["duplicadas"]))
        fuera = [p for p in info["placas"] if p not in oficial]
        if fuera:
            st.warning("Placas del presupuesto que no están en el listado oficial (se cargan igual, pero sus "
                       "movimientos no se podrán identificar solos): " + ", ".join(fuera))
        if ok:
            listos[esc] = largo
    if listos and st.button(f"Guardar presupuesto de la flota · {empresa}", key=f"mt_guardar_ppto_{empresa}", icon=":material/save:"):
        for esc, largo in listos.items():
            ctx.con.execute("DELETE FROM mtto_ppto WHERE empresa=? AND escenario=?", (empresa, esc))
            _insertar_por_lotes(ctx, "mtto_ppto(empresa, escenario, placa, numero_interno, tipo, concepto, orden, "
                                     "valor, es_global, texto_original)",
                                [(empresa, esc, r.placa, r.numero_interno, r.tipo, r.concepto, int(r.orden),
                                  float(r.valor), int(r.es_global), r.texto_original) for r in largo.itertuples(index=False)],
                                lote=80)
            ctx.registrar_carga(ctx.usuario, "Presupuesto flota", g.name,
                                f"{empresa} {esc}: {len(largo)} filas, {E.pesos(largo['valor'].sum())}")
        ctx.con.commit()
        st.success(f"Presupuesto de {empresa} guardado: " + ", ".join(listos))
        st.rerun()


def _admin_historico(ctx, empresa):
    _exigir_admin(ctx)
    st.divider()
    st.markdown(f"**3. Histórico de clasificación (opcional) · {empresa}**  \nExcel con columnas TEXTO, PROVEEDOR (opcional) y "
                "CONCEPTO, con movimientos de años anteriores ya clasificados. El concepto debe llamarse igual que en el presupuesto. "
                "El sistema aprende de él (para las dos compañías): reemplaza el histórico que cargaste antes para "
                f"{empresa}, pero no toca tus reasignaciones manuales.")
    h = st.file_uploader("Histórico de clasificación (.xlsx)", type=["xlsx"], key=f"mt_up_hist_{empresa}")
    if not (h and st.button(f"Cargar histórico · {empresa}", key=f"mt_cargar_hist_{empresa}")):
        return
    df = pd.read_excel(h, dtype=str)
    c_t, c_p, c_c = _columna(df, "TEXTO"), _columna(df, "PROVEEDOR", "TERCERO"), _columna(df, "CONCEPTO")
    if not (c_t and c_c):
        st.error("Necesito las columnas TEXTO y CONCEPTO.")
        return
    cat = ctx.leer("SELECT DISTINCT concepto FROM mtto_ppto")["concepto"].tolist() + L.CONCEPTOS_TS
    por_nombre = {L.norm_texto(c): c for c in cat}
    oficial = ctx.leer("SELECT placa FROM mtto_placas")["placa"].tolist()
    d = pd.DataFrame({"texto": df[c_t], "tercero": df[c_p] if c_p else "", "concepto_h": df[c_c].map(L.norm_texto)})
    d = L.asignar_placas(d, oficial)
    d["concepto"] = d["concepto_h"].map(por_nombre)
    d["texto_tipo"] = [L.firma_texto(t, p) for t, p in zip(d["texto"], d["placa"])]
    malos = d[d["concepto"].isna()]
    d = d[d["concepto"].notna() & (d["texto_tipo"] != "")]
    ctx.con.execute("DELETE FROM mtto_clasif_manual WHERE empresa=? AND origen='Archivo histórico'", (empresa,))
    ahora_txt = ahora().strftime("%Y-%m-%d %H:%M")
    if len(d):
        _insertar_por_lotes(ctx, "mtto_clasif_manual(empresa, llave, concepto, texto_tipo, tercero, origen, usuario, fecha)",
                            [(empresa, "", r.concepto, r.texto_tipo, L.norm_texto(r.tercero), "Archivo histórico",
                              ctx.usuario, ahora_txt) for r in d.itertuples()])
    ctx.con.commit()
    ctx.registrar_carga(ctx.usuario, "Histórico clasificación", h.name, f"{empresa}: {len(d)} filas usadas de {len(df)}")
    st.success(f"{len(d)} filas del histórico cargadas.")
    if len(malos):
        st.warning(f"{len(malos)} filas con un concepto que no existe en el presupuesto no se cargaron: "
                   + ", ".join(sorted(set(malos['concepto_h']))[:15]))


def _admin_ts(ctx, empresa, oficial):
    _exigir_admin(ctx)
    st.divider()
    st.markdown(f"**4. Presupuesto de Tecnomecánica y SOAT · {empresa}**  \nExcel con una hoja por concepto (placa en filas, meses en "
                "columnas, fila «TOTAL MES»). Se guardan todos los meses (UPA = sep–dic 2026; Normal 2027 = ene–dic 2027) y se "
                "valida cada mes contra «TOTAL MES». Ojo: la columna TOTAL del archivo suma solo 2027. Reemplaza lo cargado "
                f"de los conceptos que subas para {empresa}.")
    t = st.file_uploader("Tecnomecánica y SOAT (.xlsx)", type=["xlsx"], key=f"mt_up_ts_{empresa}")
    if not t:
        return
    xt = pd.ExcelFile(t)
    elegidas = {}
    for concepto in L.CONCEPTOS_TS:
        opciones = ["(no cargar)"] + xt.sheet_names
        clave = L.norm_texto(concepto)[:4]                                  # TECN / SOAT
        sugerida = next((h for h in xt.sheet_names if clave in L.norm_texto(h)), None)
        hoja = st.selectbox(f"Hoja de {concepto}", opciones, key=f"mt_ts_hoja_{concepto}_{empresa}",
                            index=opciones.index(sugerida) if sugerida else 0)
        if hoja != "(no cargar)":
            elegidas[concepto] = (hoja, xt.parse(hoja))
    if not elegidas:
        return
    try:
        largo, info = L.leer_ts(elegidas)
    except ValueError as e:
        st.error(str(e))
        return
    ok = True
    for concepto, i in info["hojas"].items():
        bien = not i["meses_que_no_cuadran"] and not i["filas_que_no_cuadran"] and i["tiene_total_mes"]
        ok &= bien
        por_anio = " · ".join(f"{a}: {E.pesos(v)}" for a, v in sorted(i["por_anio"].items()))
        (st.success if bien else st.error)(
            f"**{concepto}** (hoja {i['hoja']}): {i['n_placas']} placas · {por_anio}."
            + ("" if bien else f" Meses que no cuadran con «TOTAL MES»: {', '.join(i['meses_que_no_cuadran']) or '—'}; "
                               f"filas cuyo TOTAL no cuadra: {i['filas_que_no_cuadran']}."))
    if info["duplicadas"]:
        st.warning("Placas repetidas en el archivo: " + ", ".join(sorted(set(info["duplicadas"]))))
    fuera = [p for p in info["placas"] if p not in oficial]
    if fuera:
        st.warning("Placas que no están en el listado oficial (se cargan igual, pero sus movimientos no se podrán identificar "
                   "solos): " + ", ".join(fuera))
    if ok and st.button(f"Guardar Tecnomecánica y SOAT · {empresa}", key=f"mt_guardar_ts_{empresa}", icon=":material/save:"):
        conceptos = list(info["hojas"])
        ctx.con.execute("DELETE FROM mtto_ppto_ts WHERE empresa=? AND concepto IN (" + ",".join("?" * len(conceptos)) + ")",
                        (empresa, *conceptos))
        _insertar_por_lotes(ctx, "mtto_ppto_ts(empresa, placa, concepto, anio, mes, valor, texto_original)",
                            [(empresa, r.placa, r.concepto, int(r.anio), int(r.mes), float(r.valor), r.texto_original)
                             for r in largo.itertuples(index=False)])
        ctx.con.commit()
        ctx.registrar_carga(ctx.usuario, "Presupuesto Tecnomecánica/SOAT", t.name,
                            f"{empresa}: {len(largo)} filas, {E.pesos(largo['valor'].sum())}")
        st.success(f"Presupuesto de Tecnomecánica y SOAT de {empresa} guardado.")
        st.rerun()


# ================================================================== diagnóstico de nombres de sub-rubro
def _diagnostico(ppto: pd.DataFrame, ejec: pd.DataFrame):
    """Muestra qué sub-rubros parecidos a los del módulo existen en la base (detecta diferencias de nombre) y de dónde
    salen los movimientos de SOAT y tecnomecánica (por su texto). Separado por compañía."""
    patron = "flota|transport|soat|segur|tecno"
    claves = ["empresa", "codigo", "sub_rubro"]
    p = ppto[ppto["_sub"].str.contains(patron)].groupby(claves)["valor"].agg(presupuesto="sum", filas_presupuesto="size")
    e = ejec[ejec["_sub"].str.contains(patron)].groupby(claves)["valor"].agg(ejecutado="sum", movimientos="size")
    t = p.join(e, how="outer").reset_index()
    if t.empty:
        st.warning("No hay ningún sub-rubro con «flota», «transporte», «soat», «seguro» o «tecno» en el presupuesto ni "
                   "en la ejecución de este año.")
    else:
        st.dataframe(t, hide_index=True, width="stretch", column_config={"empresa": "Compañía"})
    cand = ejec[L.candidatos_ts(ejec)]
    st.markdown("**Movimientos cuyo texto o cuenta habla de SOAT o tecnomecánica, por sub-rubro**")
    if cand.empty:
        st.caption("Ninguno en el año seleccionado.")
    else:
        st.dataframe(cand.groupby(claves)["valor"].agg(movimientos="size", valor="sum").reset_index(),
                     hide_index=True, width="stretch", column_config={"empresa": "Compañía"})
    st.caption(f"El módulo busca el sub-rubro «{L.SUBRUBRO_FLOTA}» (SUBRUBRO_FLOTA en logica_m.py) y trae además los movimientos "
               "de SOAT/tecnomecánica de cualquier otro sub-rubro. Si en la base el nombre es otro, corrígelo ahí.")


def _soporte(S, empresa: str, mes: int, detalle_meses: bool = False) -> bytes:
    """Excel «Soporte de ejecución de mantenimiento» de UNA compañía: hoja Presupuesto (presupuesto de todos los meses y ejecución
    acumulada hasta el mes elegido, por sub-rubro) y una hoja por sub-rubro con los movimientos de SAP del mes elegido. Lo puede
    descargar cualquier usuario (quien no ve todo recibe solo lo de su gerencia). Respeta los movimientos eliminados; no depende
    del filtro de placa (es el soporte del rubro completo)."""
    ctx, cfg = S.ctx, S.cfg                                  # lo puede descargar cualquier usuario; _rubro respeta su gerencia
    ppto, ejec = _rubro(S, empresa, detalle_flota=True)
    ejec = ejec.copy()
    if len(ejec):
        crudo = ctx.leer("SELECT id, raw FROM ejecucion WHERE anio=? AND empresa=?", (int(cfg["anio"]), empresa))
        ejec = ejec.merge(crudo, on="id", how="left") if "id" in ejec else ejec.assign(raw="")
        num_cuenta = L.numeros_cuenta(crudo)                 # el detalle de SISPO usa el número de cuenta de SAP que reemplaza
    else:
        ejec = ejec.assign(raw="")
        num_cuenta = {}
    return L.exportar_soporte(ppto, ejec, dict(cfg, escenario=S.escenario), mes, "UPA" if S.escenario.startswith("UPA") else "PRESUPUESTO",
                              detalle_meses=detalle_meses, num_cuenta=num_cuenta)


def _excel_control(S, ts_mes: pd.DataFrame) -> bytes:
    """Excel de control de 3 hojas (trae el detalle de movimientos): solo administrador, también en la lógica."""
    _exigir_admin(S.ctx)
    return L.exportar_excel({e: S.vistas_emp[e] for e in S.sel}, ts_mes, S.cfg)


def _estado_cargas(S) -> pd.DataFrame:
    """Qué hay cargado de cada compañía: sirve para ver que ALDC y ALSM se cargaron cada una por su lado."""
    filas = []
    for e, d in S.datos_e.items():
        mv = S.mov_pre[S.mov_pre["empresa"] == e]
        filas.append({"Compañía": e, "Movimientos en el periodo": len(mv), "Valor movimientos": float(mv["valor"].sum()),
                      f"Filas ppto flota ({S.escenario})": int((d.ppto["es_global"] != 2).sum()) if len(d.ppto) else 0,
                      "Placas con ppto Tecnomecánica/SOAT": int(d.ts_per["placa"].nunique()) if len(d.ts_per) else 0,
                      "Movimientos excluidos": int((S.mov_excl["empresa"] == e).sum())})
    return pd.DataFrame(filas)


# ================================================================== pantalla principal
def _mostrar(ctx):
    leer, usuario, es_admin = ctx.leer, ctx.usuario, ctx.es_admin
    ver_todo = getattr(ctx, "ver_todo", es_admin)
    _asegurar_tablas(ctx)
    zona_encabezado = st.container()

    # Barra de control propia de Mantenimiento (reemplaza a «Filtros del informe», que aquí no aplican: el año y la
    # versión los define el escenario). El filtro de placa es GLOBAL: está arriba de las pestañas y cambia los datos.
    with st.container(key="mt_barra"):
        if es_admin:
            c_emp, c_esc, c_hasta, c_placa = st.columns([0.9, 1.0, 0.9, 1.5], vertical_alignment="bottom")
        else:                                           # el usuario normal ve el consolidado: sin filtro por placa
            c_emp, c_esc, c_hasta = st.columns([1, 1, 1], vertical_alignment="bottom")
            c_placa = None
        vista_emp = c_emp.selectbox("Compañía", [TODAS, *EMPRESAS], key="mt_empresa",
                                    help="Consolidado = ALDC + ALSM. " + " · ".join(f"{e}: {EMPRESAS[e]}" for e in EMPRESAS))
        escenario = c_esc.selectbox("Escenario presupuestal", list(L.ESCENARIOS), key="mt_escenario",
                                    help="UPA 2026: sep–dic 2026. Normal 2027: ene–dic 2027. Sus cifras nunca se mezclan.")
    cfg = dict(L.ESCENARIOS[escenario], escenario=escenario)
    sel = list(EMPRESAS) if vista_emp == TODAS else [vista_emp]
    multi = len(sel) > 1

    # ------------------------------------------------------------------ datos de cada compañía
    asig = asignacion.mapa(leer)
    datos_e = {e: _datos_empresa(ctx, e, cfg, asig) for e in EMPRESAS}
    oficial = leer("SELECT placa, numero_interno, tipo FROM mtto_placas")
    if not ver_todo:                                    # solo ve la flota quien la tiene asignada
        tiene = False
        for d in datos_e.values():
            d.pool = d.pool[d.pool["gerencia"] == usuario]
            subs = d.ppto_base[d.ppto_base["sub_rubro"].fillna("").map(cargas.norm) == cargas.norm(L.SUBRUBRO_FLOTA)]
            tiene |= d.pool.shape[0] > 0 or (not subs.empty and (asignacion.columna(subs, asig) == usuario).any())
        if not tiene:
            st.info("Tu gerencia no tiene asignado el sub-rubro «Flota y Equipo de transporte».")
            return
    m0, m1 = cfg["meses"]
    pool_todo = pd.concat([d.pool for d in datos_e.values()], ignore_index=True)
    en_ventana = pool_todo[(pool_todo["mes"] >= m0) & (pool_todo["mes"] <= m1)]
    ult = int(en_ventana["mes"].max()) if not en_ventana.empty else m1
    hasta_mes = c_hasta.selectbox("Ejecución hasta", list(range(m0, m1 + 1)), index=ult - m0, key=f"mt_hasta_{escenario}",
                                  format_func=lambda m: MESES_L[m - 1].capitalize())
    if m0 == hasta_mes:
        hasta = f"{MESES_L[hasta_mes - 1]} {cfg['anio']}"
    else:
        hasta = f"{MESES_L[m0 - 1]}–{MESES_L[hasta_mes - 1]} {cfg['anio']}" if m0 != 1 else f"a {MESES_L[hasta_mes - 1]} {cfg['anio']}"

    ppto_emp = {e: d.ppto_all for e, d in datos_e.items()}
    ppto_todo = pd.concat(list(ppto_emp.values()), ignore_index=True)
    if not (ppto_todo["es_global"] != 2).any():
        zona_encabezado.html(E.encabezado(
            "Mantenimiento" + ("" if ver_todo else f" · {usuario}"),
            "Mantenimiento · Flota y Equipo de Transporte · Tecnomecánica y SOAT · Contenedores: presupuesto vs ejecución · valores en pesos", [(vista_emp, "filtro")]))
        st.warning(f"No hay presupuesto de flota cargado para **{escenario}** en ninguna compañía. El administrador lo carga abajo, "
                   "en «Datos base».", icon=":material/warning:")
        if es_admin:
            _admin_datos(ctx)
        return
    if oficial.empty:
        st.warning("No hay listado oficial de placas: todos los movimientos aparecen como «SIN PLACA IDENTIFICADA». "
                   "El administrador lo carga en «Datos base».", icon=":material/warning:")

    # selector GLOBAL de placa: filtra los datos de Placas y Tecnomecánica/SOAT (los generales no se atribuyen a placas)
    ts_placas = {p for d in datos_e.values() for p in d.ts_per["placa"]}
    placas_edit = sorted((set(ppto_todo.loc[ppto_todo["es_global"] == 0, "placa"]) | ts_placas | set(oficial["placa"])) - {""})
    placas_op = [TODAS_PLACAS, *placas_edit, L.SIN_PLACA]
    if st.session_state.get("mt_f_placa") not in placas_op:
        st.session_state["mt_f_placa"] = TODAS_PLACAS
    if es_admin:
        placa_f = c_placa.selectbox("Placa", placas_op, key="mt_f_placa",
                                    help="Filtro global: cambia las cifras de Flota y Tecnomecánica/SOAT (tablas, gráficos, detalle y Excel). "
                                         "Los conceptos generales no se atribuyen a una placa.")
        placa = None if placa_f == TODAS_PLACAS else placa_f
    else:
        placa = None
    zona_encabezado.html(E.encabezado(
        "Mantenimiento" + ("" if ver_todo else f" · {usuario}"),
        "Mantenimiento · Flota y Equipo de Transporte · Tecnomecánica y SOAT · Contenedores: presupuesto vs ejecución · valores en pesos",
        [(vista_emp if not multi else f"Consolidado · {' + '.join(sel)}", "filtro")] + ([(f"Placa {placa}", "filtro")] if placa else [])))
    if placa:
        st.html(f'<div class="mt-filtro" style="display:inline-block;margin:.2rem 0 .6rem;padding:.35rem .8rem;border-radius:8px;'
                f'background:#eaf1fb;color:#1f3a5f;font-size:.82rem"><b>Filtro activo: placa {escape(placa)}.</b> Presupuesto, ejecución, saldo, % y detalle de '
                "Flota y Tecnomecánica/SOAT son solo de esta placa. Los conceptos generales no se atribuyen a placas: no se muestran "
                f"y no entran al total de la placa. Para volver al consolidado elige «{TODAS_PLACAS}».</div>")

    # ------------------------------------------------------------------ clasificación y cálculo (una sola base de movimientos)
    catalogo = L.catalogo_conceptos(ppto_todo)
    tipos = L.tipos_concepto(ppto_todo)
    reglas_db = leer("SELECT orden, patron, proveedor, concepto, confianza, observacion, activa FROM mtto_reglas ORDER BY orden")
    reglas = L.asegurar_regla_lavado(
        L.asegurar_reglas_ts(reglas_db if not reglas_db.empty else L.reglas_iniciales(catalogo), catalogo), catalogo)
    manual = _normalizar_llaves(leer("SELECT empresa, llave, concepto, texto_tipo, tercero, origen, placa FROM mtto_clasif_manual"))
    excluidos = _normalizar_llaves(leer("SELECT empresa, llave, fecha_mov, documento, tercero, texto, valor, concepto, placa, motivo, "
                                        "usuario, fecha FROM mtto_excluidos"))
    # SISPO: líneas ya aplicadas del año del escenario y movimientos genéricos de SAP que reemplazan
    pool_sap = _pool_sap(datos_e)
    lin_sispo = leer("SELECT * FROM mtto_sispo_lineas WHERE periodo >= ? AND periodo <= ?", (int(cfg["anio"]) * 100 + 1, int(cfg["anio"]) * 100 + 12))
    rep_sispo = leer("SELECT empresa, periodo, cuenta, llave, gerencia, cc, codigo, sub_rubro, sub_rubro_sap, valor "
                     "FROM mtto_sispo_reemplazados WHERE periodo >= ? AND periodo <= ?",
                     (int(cfg["anio"]) * 100 + 1, int(cfg["anio"]) * 100 + 12))
    # registros que dejó la prueba de octubre (aplicados sin conciliar con SAP): no entran a ningún cálculo hasta revertirlos y reaplicarlos
    her = leer("SELECT empresa, periodo, cuenta FROM mtto_sispo_periodo WHERE estado=?", (ESTADO_APLICACION_SIN_SAP_HEREDADA,))
    if len(her) and len(lin_sispo):
        k_her = {(r.empresa, int(r.periodo), L.cuenta_de(r.cuenta) or L.SISPO_CUENTA) for r in her.itertuples()}
        k_lin = [(e, int(p_), L.cuenta_de(c) or L.SISPO_CUENTA) for e, p_, c in zip(lin_sispo["empresa"], lin_sispo["periodo"], lin_sispo["cuenta_nombre"])]
        lin_sispo = lin_sispo[[k not in k_her for k in k_lin]]
        if len(rep_sispo):
            k_rep = [(e, int(p_), L.cuenta_de(c) or L.SISPO_CUENTA) for e, p_, c in zip(rep_sispo["empresa"], rep_sispo["periodo"], rep_sispo["cuenta"])]
            rep_sispo = rep_sispo[[k not in k_her for k in k_rep]]
    # la página Flota solo lleva lo de la cuenta de la flota; las demás cuentas se ven en el rubro (Mantenimiento, Contenedores, soporte)
    sispo_flota = L.sispo_de_cuentas(lin_sispo, rep_sispo, [L.SISPO_CUENTA])
    sispo_otras = L.sispo_de_cuentas(lin_sispo, rep_sispo, [c for c in L.SISPO_CUENTAS if c != L.SISPO_CUENTA])
    sap_salidas = _pool_salidas(datos_e, asig) if es_admin else pd.DataFrame()
    mov_pre, mov_todos, mov_excl = _movimientos_maestro(datos_e, cfg, hasta_mes, oficial, reglas, manual, catalogo, tipos, excluidos,
                                                        pool_sap=pool_sap, sispo=sispo_flota,
                                                        solo_gerencia=None if ver_todo else usuario)

    vistas_emp = {e: L.calcular(ppto_emp[e], mov_todos[mov_todos["empresa"] == e], oficial, placa, tipos) for e in EMPRESAS}
    vistas = vistas_emp[sel[0]] if not multi else L.calcular(
        pd.concat([ppto_emp[e] for e in sel], ignore_index=True), mov_todos[mov_todos["empresa"].isin(sel)], oficial, placa, tipos)
    ppto_sel = pd.concat([ppto_emp[e] for e in sel], ignore_index=True)
    ppto_val = ppto_sel[(ppto_sel["es_global"] == 1) | (ppto_sel["placa"] == placa)] if placa else ppto_sel

    def alcance(df):                                    # mismo recorte que hace calcular(): compañías elegidas y placa
        d = df[df["empresa"].isin(sel)]
        return d[(d["ambito"] == L.AMB_GEN) | (d["placa"] == placa)] if placa else d

    n1, n2, conflictos = L.construir_memoria(manual, catalogo)
    man = manual[(manual["origen"] == "Manual") & manual["llave"].fillna("").ne("")]
    mv_v = vistas["movimientos"]
    valid = None
    if es_admin:                                        # las validaciones muestran detalle interno: solo administrador
        valid = L.validaciones(vistas, oficial, ppto_val, dict(
            manuales=int((mv_v["origen"] == "Manual").sum()), placas_manuales=int((mv_v["placa_estado"] == "Asignada a mano").sum()),
            firmas=len(n1) + len(n2), conflictos=conflictos, llaves_repetidas=int(man["llave"].duplicated().sum()),
            huerfanas=int((~man["llave"].isin(set(mov_pre["llave"]))).sum()),
            fuente_total=float(alcance(mov_pre)["valor"].sum()), excluidos_valor=float(alcance(mov_excl)["valor"].sum()),
            excluidos_n=len(alcance(mov_excl))))
        if multi:                                       # el consolidado debe ser exactamente la suma de las compañías
            tot = lambda v_: v_["resumen"].iloc[-1]
            dp = float(tot(vistas)["presupuesto"] - sum(tot(vistas_emp[e])["presupuesto"] for e in sel))
            de = float(tot(vistas)["ejecutado"] - sum(tot(vistas_emp[e])["ejecutado"] for e in sel))
            valid.loc[len(valid)] = ["Consolidado = suma de las compañías", "✔ Correcto" if abs(dp) < 1 and abs(de) < 1 else "⚠ Revisar",
                                     f"Diferencia en presupuesto: {E.pesos(dp)}; en ejecución: {E.pesos(de)}. El Excel se arma por compañía."]

    ts_mes_sel = pd.concat([datos_e[e].ts_mes for e in sel], ignore_index=True)
    ts_scen = ts_mes_sel[(ts_mes_sel["anio"] == cfg["anio"]) & (ts_mes_sel["mes"] >= m0) & (ts_mes_sel["mes"] <= m1)]
    ts_per_sel = pd.concat([datos_e[e].ts_per for e in sel], ignore_index=True)
    ppto_mes = {k: sorted(g.groupby("mes")["valor"].sum().items()) for k, g in ts_scen.groupby(["placa", "concepto"])} if len(ts_scen) else {}

    tipo_ppto = {}                                      # placa -> concepto presupuestal (columna TIPO VEHICULO del archivo)
    for r in ppto_sel[ppto_sel["es_global"] == 0].itertuples():
        if r.placa and str(r.tipo or "").strip():
            tipo_ppto.setdefault(r.placa, " ".join(str(r.tipo).split()))
    S = SimpleNamespace(
        ctx=ctx, asig=asig, sap_pool=pool_sap, sap_salidas=sap_salidas, oficial=oficial, tipos=tipos, ppto_sel=ppto_sel, lin_sispo=sispo_flota[0], sispo_flota=sispo_flota, sispo_otras=sispo_otras, registro_placas=oficial[["placa", "numero_interno"]], ver_todo=ver_todo, sel=sel, multi=multi, escenario=escenario, cfg=cfg, m0=m0, m1=m1, hasta=hasta, hasta_mes=hasta_mes,
        placa=placa, vistas=vistas, vistas_emp=vistas_emp, catalogo=catalogo, placas_edit=placas_edit, manual=manual,
        es_admin=es_admin, ts_scen=ts_scen, ppto_mes=ppto_mes, datos_e=datos_e, mov_pre=mov_pre, mov_todos=mov_todos,
        mov_excl=mov_excl, excluidos=excluidos, rubro_cache={}, tipo_ppto=tipo_ppto, ult=ult,
        emp_placa=_empresas_por_placa(ppto_sel, mov_todos[mov_todos["empresa"].isin(sel)]),
        ts_original=ts_per_sel.drop_duplicates("concepto").set_index("concepto")["texto_original"] if len(ts_per_sel) else {},
        manual_prev={r.llave: (r.concepto or "", r.placa or "") for r in man.itertuples()})

    # ------------------------------------------------------------------ validación cruzada del rubro (administrador)
    if es_admin and not placa:
        pr, er = _rubro_sel(S)
        tr = L.tabla_subrubros(pr, er, cfg, hasta_mes)
        fl = tr[tr["sub_rubro"].map(L.norm_texto) == L.norm_texto(L.SUBRUBRO_FLOTA)]
        mod = vistas["resumen"].iloc[-1]
        dp_ = float(fl["presupuesto"].sum() - mod["presupuesto"])
        de_ = float(fl["ejecutado"].sum() - mod["ejecutado"])
        valid.loc[len(valid)] = ["Flota del rubro = Flota del módulo (Placas + Generales + Tecnomecánica/SOAT)",
                                 "✔ Correcto" if abs(dp_) < 1 and abs(de_) < 1 else "⚠ Revisar",
                                 f"Diferencia en presupuesto: {E.pesos(dp_)}; en ejecución: {E.pesos(de_)}. La pestaña Mantenimiento lee el rubro "
                                 "cargado en Presupuesto; Flota lee el detalle por placa cargado en «Datos base»."]

    # ------------------------------------------------------------------ descargas y validaciones (solo administrador)
    if es_admin:
        with st.expander("Descargas y validaciones (administrador)", icon=":material/download:"):
            nombre = f"mantenimiento_{escenario.replace(' ', '_')}_{'consolidado' if multi else sel[0]}" + (f"_{placa}" if placa else "") + ".xlsx"
            st.download_button("Descargar Excel de control", lambda: _excel_control(S, ts_mes_sel), nombre,
                               icon=":material/download:", key="mt_excel",
                               help="Tres hojas: VEHÍCULOS (Placas + Generales), TECNOMECÁNICAS y SOAT, con el detalle de movimientos. "
                                    "Respeta la compañía y la placa elegidas; sin filtros, el consolidado completo.")
            st.markdown("**Validaciones de los datos**")
            st.dataframe(valid, hide_index=True, width="stretch")
            st.markdown("**Qué hay cargado de cada compañía**")
            st.dataframe(_estado_cargas(S), hide_index=True, width="stretch", column_config={
                "Valor movimientos": st.column_config.NumberColumn(format="accounting")})
            sin_mov = [e for e in EMPRESAS if S.mov_pre[S.mov_pre["empresa"] == e].empty]
            if sin_mov:
                st.warning("No hay movimientos de " + ", ".join(sin_mov) + f" en {escenario} ({hasta}). Se muestran sus presupuestos, "
                           "pero la ejecución sale en cero hasta que se cargue en el módulo de ejecución.", icon=":material/warning:")
            st.caption("El Excel se genera con las mismas tablas de la pantalla, por eso sus totales coinciden.")

    # ------------------------------------------------------------------ pestañas
    nombres_tabs = ["Mantenimiento", "Flota y Equipo de Transporte", "Tecnomecánica y SOAT", "Contenedores"]
    t_m, t_f, t_t, t_c = st.tabs(nombres_tabs)
    with t_m:
        _tab_mantenimiento(S)
    with t_f:
        _tab_flota(S)
    with t_t:
        _tab_ts(S)
    with t_c:
        _tab_contenedores(S)

    # lo que sigue es solo del administrador (la lógica de cada función lo vuelve a exigir)
    if es_admin:
        _reclasificacion(S)
        _avanzado(S, reglas, bool(reglas_db.empty))
        _admin_datos(ctx)
        with st.expander("¿Qué sub-rubros hay en la base?", icon=":material/search:"):
            _diagnostico(pd.concat([d.ppto_base for d in datos_e.values()], ignore_index=True),
                         pd.concat([d.ejec for d in datos_e.values()], ignore_index=True))


def _verificar_logica() -> bool:
    """`logica_m` es otro archivo: si la app tiene en memoria una copia vieja, incompleta o un módulo vacío con ese nombre, le
    faltan nombres que este módulo usa. Primero se intenta recargarlo; si no se puede (p. ej. no encuentra el archivo en el path), se
    carga directamente desde el archivo que está junto a este módulo. Si aun así faltan nombres, se muestra qué se cargó."""
    import importlib
    import importlib.util
    import os
    import sys
    global L
    requeridos = ("ESCENARIOS", "es_staf_otros", "ids_staf_otros_raw", "ids_staf_otros_ejec", "asegurar_regla_lavado", "exportar_soporte", "numeros_cuenta")
    faltan = lambda m: [n for n in requeridos if not hasattr(m, n)]
    if not faltan(L):
        return True
    previo = f"{type(L).__name__} · archivo={getattr(L, '__file__', None)} · spec={getattr(L, '__spec__', None)} · nombres={len(vars(L))}"
    try:
        importlib.reload(L)
    except Exception:
        pass
    if not faltan(L):
        return True
    ruta = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logica_m.py")
    try:
        spec = importlib.util.spec_from_file_location("logica_m", ruta)
        nuevo = importlib.util.module_from_spec(spec)
        sys.modules["logica_m"] = nuevo
        spec.loader.exec_module(nuevo)
        if not faltan(nuevo):
            L = nuevo
            return True
        detalle = f"Se cargó {ruta} pero le faltan: {faltan(nuevo)}"
    except Exception:
        detalle = "Al cargar " + ruta + " falló:\n" + traceback.format_exc()
    st.error("La app no está usando el logica_m.py actual (faltan: " + ", ".join(faltan(L)) + ").")
    st.code(f"Módulo que había en memoria: {previo}\nTamaño del archivo junto a este módulo: "
            f"{os.path.getsize(ruta) if os.path.exists(ruta) else 'no existe'} bytes (el actual pesa 133.025)\n{detalle}")
    return False


def mostrar(ctx):
    """Punto de entrada. Si algo falla, muestra el error completo (Streamlit lo oculta en la nube)."""
    try:
        if not _verificar_logica():
            return
        _mostrar(ctx)
    except Exception:
        st.error("El módulo de Mantenimiento tuvo un error. Copia este detalle y compártelo para corregirlo.")
        st.code(traceback.format_exc())
