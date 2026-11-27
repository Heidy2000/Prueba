"""
Presupuestos: versiones, escenarios y selección compartida (lógica pura, sin Streamlit ni base de datos).

Una sola fuente de verdad para todos los módulos (Presupuesto, Mantenimiento, Operaciones, Combustible, Corte y poda, Nómina y la presentación):

- VERSIONES: «Inicial» (el aprobado), «4+8» (4 meses reales + 8 presupuestados) y «8+4» (8 meses reales + 4 presupuestados).
  La versión 8+4 es la que antes se llamaba «UPA». Las tres cubren los 12 meses de 2026.
- ESCENARIOS: lo que eligen Mantenimiento y Operaciones: una de las tres versiones del año 2026 (las tres cubren los 12 meses del año) o el
  presupuesto «Inicial 2027» (el año siguiente solo tiene la versión Inicial, que se guarda sin versión).
- La versión elegida vive en `st.session_state["ppto_version"]` (ver ui.selector_presupuesto): se elige en cualquier módulo y los demás la
  siguen, y la presentación la usa como base de todos sus comparativos.
"""
import pandas as pd

INICIAL, V48, V84 = "Inicial", "4+8", "8+4"
VERSIONES = (INICIAL, V48, V84)
DEFECTO = V84                                    # la más reciente: el cierre esperado del año
MESES_REALES = {INICIAL: 0, V48: 4, V84: 8}      # meses del año que ya son ejecución real dentro de cada versión
AYUDA = ("Inicial: el aprobado. 4+8: 4 meses reales + 8 presupuestados. 8+4: 8 meses reales + 4 presupuestados (cierre esperado del año).")
ANIO_VERSIONES = 2026                            # año al que pertenecen las tres versiones en Mantenimiento y Operaciones
ESCENARIO_SIGUIENTE = "Inicial 2027"             # presupuesto del año siguiente: solo existe el Inicial, sin versiones
CLAVE_VERSION = "ppto_version"                   # clave de session_state con la versión elegida (compartida por todos los módulos)


def etiqueta(version) -> str:
    """«Presupuesto 8+4». Para una versión vacía (datos sin versión) devuelve «Presupuesto»."""
    return f"Presupuesto {version}" if version else "Presupuesto"


def canonica(nombre) -> str:
    """El nombre oficial de una versión escrito de cualquier forma («inicial», «PPTO 4 + 8», «8+4»): «Inicial», «4+8» u «8+4». Si no es ninguna,
    el texto tal cual."""
    t = str(nombre).strip()
    clave = t.upper().replace(" ", "")
    return next((v for v in VERSIONES if v.upper() == clave), t)


def ordenar(versiones) -> list:
    """Versiones sin repetir, en el orden Inicial → 4+8 → 8+4 (las desconocidas, al final)."""
    orden = {v: i for i, v in enumerate(VERSIONES)}
    return sorted({str(v) for v in versiones if v is not None and str(v) != ""}, key=lambda v: (orden.get(v, 9), v))


def elegir(disponibles, preferida=None) -> str:
    """La versión a usar entre las disponibles: la preferida si existe; si no, la 8+4; si no, la más reciente; sin ninguna, «Inicial»."""
    vs = ordenar(disponibles)
    if preferida in vs:
        return preferida
    return DEFECTO if DEFECTO in vs else (vs[-1] if vs else INICIAL)


def versiones_de(leer, anio, empresa) -> list:
    """Versiones del presupuesto general (tabla `presupuesto`) cargadas para un año y una compañía, ordenadas."""
    return ordenar(leer("SELECT DISTINCT version FROM presupuesto WHERE anio=? AND empresa=?", (int(anio), empresa))["version"].tolist())


def version_para(leer, empresa, anio, preferida=None) -> str:
    """Versión de la tabla `presupuesto` que se usa para un año y compañía (la preferida si está cargada)."""
    return elegir(versiones_de(leer, anio, empresa), preferida)


# ------------------------------------------------------------------ escenarios (Mantenimiento y Operaciones)
def _escenarios() -> dict:
    esc = {etiqueta(v): dict(anio=ANIO_VERSIONES, meses=(1, 12), hoja=str(ANIO_VERSIONES) if v == V84 else "", version=v)
           for v in VERSIONES}
    esc[ESCENARIO_SIGUIENTE] = dict(anio=ANIO_VERSIONES + 1, meses=(1, 12), hoja=str(ANIO_VERSIONES + 1), version=None)
    return esc


ESCENARIOS = _escenarios()                       # nombre -> dict(anio, meses, hoja del Excel por defecto, version)
ESCENARIO_ANTERIOR = {"UPA 2026": etiqueta(V84), "Normal 2027": ESCENARIO_SIGUIENTE}  # nombre con que se guardó antes: se actualiza en la base (ver mod_mantenimiento)


def escenario_de_version(version) -> str | None:
    """Nombre del escenario de una versión (None si la versión no tiene escenario)."""
    return next((n for n, c in ESCENARIOS.items() if c["version"] == version), None)


def escenario_para(anio, mes, version=None) -> str | None:
    """Escenario de la `version` pedida que cubre un mes de un año (None si esa versión no cubre ese mes; nunca se cambia en silencio por otra versión). Un año sin versiones (2027) usa su único escenario; sin versión pedida, la por defecto."""
    anio, mes = int(anio), int(mes)
    cands = [(n, c) for n, c in ESCENARIOS.items() if c["anio"] == anio]
    for n, c in cands:                                  # año sin versiones
        if c["version"] is None and c["meses"][0] <= mes <= c["meses"][1]:
            return n
    pedida = version if version in VERSIONES else DEFECTO
    for n, c in cands:
        if c["version"] == pedida:
            return n if c["meses"][0] <= mes <= c["meses"][1] else None
    return None


def escenario_anio(anio, version=None) -> str | None:
    """Escenario de un año (sin importar el mes): el de la `version` pedida; si el año no tiene versiones, su único escenario; si no hay, el
    de la versión por defecto. None si ningún escenario es de ese año."""
    cands = [(n, c) for n, c in ESCENARIOS.items() if c["anio"] == int(anio)]
    for n, c in cands:
        if c["version"] == version:
            return n
    for n, c in cands:
        if c["version"] in (None, DEFECTO):
            return n
    return cands[0][0] if cands else None


def version_de_escenario(escenario) -> str | None:
    """Versión que representa un escenario (None para «Inicial 2027»)."""
    return ESCENARIOS.get(escenario, {}).get("version")


def rotulo(escenario) -> str:
    """Rótulo corto para los archivos de descarga: «8+4», «4+8», «Inicial» o «PRESUPUESTO» (año sin versiones)."""
    return version_de_escenario(escenario) or "PRESUPUESTO"


# ------------------------------------------------------------------ tablas con versión opcional (Nómina, Corte y poda, Operaciones)
def filtrar(df: pd.DataFrame, version):
    """Filas de un presupuesto guardado con columna `version` que corresponden a la versión elegida. La versión solo aplica a las filas del año
    con versiones (2026): las de otros años (el Inicial 2027) se dejan como están.
    Devuelve (filas, usada): `usada` = la versión pedida si hay filas de ella; "" si solo hay presupuesto sin versión (cargas anteriores a las
    versiones: se usa para cualquiera, y la pantalla lo avisa); None si solo hay otras versiones (no hay nada que comparar)."""
    if df is None or len(df) == 0 or "version" not in df.columns:
        return df, version
    if "anio" in df.columns:
        es26 = df["anio"] == ANIO_VERSIONES
        otros, df = df[~es26], df[es26]
        if df.empty:
            return otros, version
    else:
        otros = df.iloc[0:0]
    v = df["version"].fillna("").astype(str)
    exacto, sin = df[v == str(version)], df[v == ""]
    if len(exacto):
        sel, usada = exacto, version
    elif len(sin):
        sel, usada = sin, ""
    else:
        sel, usada = df.iloc[0:0], None
    return pd.concat([sel, otros], ignore_index=True), usada


def filtrar_escenario(df: pd.DataFrame, cfg: dict):
    """`filtrar` para un escenario de Mantenimiento/Operaciones: «Inicial 2027» (sin versión) no filtra nada. Devuelve (filas, usada)."""
    return filtrar(df, cfg["version"]) if cfg.get("version") else (df, "")


def condicion_version(version: str) -> tuple:
    """Condición SQL (texto, parámetros) de las filas de una versión. Lo guardado sin versión era el UPA (8+4) o el año siguiente: se incluye al
    reemplazar la 8+4 o un año sin versiones."""
    if version in (V84, ""):
        return "(version=? OR version IS NULL OR version='')", (version,)
    return "version=?", (version,)


def version_fila(anio, version) -> str:
    """Versión con que se guarda una fila de presupuesto: la elegida si es del año con versiones; vacía (sin versión) en los demás años."""
    return version if int(anio) == ANIO_VERSIONES else ""


def aviso_base(leer, cfg: dict, empresas) -> str:
    """Aviso cuando el presupuesto por sub-rubro (tabla `presupuesto`, que se carga en el módulo Presupuesto) no tiene la versión del escenario
    para alguna compañía y se usó otra: '' si todo coincide."""
    v = cfg.get("version")
    if not v:
        return ""
    otras = {}
    for e in empresas:
        vs = versiones_de(leer, cfg["anio"], e)
        if v not in vs:
            otras[e] = elegir(vs, v)
    if not otras:
        return ""
    return (f"El {etiqueta(v)} no está cargado en el módulo Presupuesto para " + ", ".join(otras)
            + ": el presupuesto por sub-rubro sale de " + ", ".join(f"{etiqueta(o)} ({e})" for e, o in otras.items()) + ".")


def versiones_guardadas(df: pd.DataFrame) -> list:
    """Versiones presentes en una tabla con columna `version` (las filas sin versión no cuentan)."""
    if df is None or len(df) == 0 or "version" not in df.columns:
        return []
    return ordenar(df["version"].dropna().astype(str).unique())


def nota_version(usada, version) -> str:
    """Aviso para la pantalla cuando el presupuesto usado no es exactamente la versión elegida ('' si lo es)."""
    if usada == version:
        return ""
    if usada == "":
        return f"Este presupuesto se cargó sin versión: se usa para {etiqueta(version)} y para las demás."
    return f"No hay presupuesto cargado para la versión {version} en este módulo (Administración)."
