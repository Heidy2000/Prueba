"""
Lógica del módulo Mantenimiento (sin Streamlit, para poder probarla sola)
=========================================================================
Rubro 3.5 · Reparación y Conservación, sub-rubro «Flota y Equipo de transporte», más Tecnomecánica y SOAT.

Tres ámbitos. El ámbito de un movimiento lo define SU CONCEPTO (un solo concepto por movimiento, por eso no se
puede contar dos veces):
- Placas                 <- conceptos de vehículo del presupuesto de flota (filas con placa).
- Generales              <- conceptos del presupuesto de flota SIN placa (In House, grúa, lavado…).
- Tecnomecánica y SOAT   <- presupuesto mensual por placa del archivo «Soat y Tecno».

Piezas:
- Presupuesto: lee el Excel de la flota (hojas 2026 y 2027), lo pasa de formato horizontal (una columna por concepto)
  a formato largo (placa, concepto, valor) y valida que la suma coincida con el archivo. Conserva el texto original.
  Lee el Excel de Tecnomecánica y SOAT (mensual) y valida contra su fila «TOTAL MES».
- Placas: identifica la placa de cada movimiento en la columna texto, solo contra el listado oficial. Se puede
  corregir a mano por movimiento.
- Clasificación: asigna a cada movimiento un concepto con: reclasificación manual > histórico aprendido > reglas
  editables > pendiente.
- Cálculo: presupuesto vs ejecución por ámbito, concepto, placa y placa-concepto, con saldo, % y estado.
- Excel: tres hojas (VEHÍCULOS, TECNOMECÁNICAS, SOAT) armadas con los mismos cálculos que muestra la pantalla.
"""
import re
import unicodedata
from datetime import date, datetime
from io import BytesIO

import pandas as pd

# escenario -> año de la ejecución, meses que cubre y hoja del Excel de presupuesto
# UPA 2026 = últimos cuatro meses de 2026 (septiembre a diciembre): SUPUESTO, confirmar con Financiera
ESCENARIOS = {
    "UPA 2026": dict(anio=2026, meses=(9, 12), hoja="2026"),
    "Normal 2027": dict(anio=2027, meses=(1, 12), hoja="2027"),
}
SUBRUBRO_FLOTA = "Flota y Equipo de transporte"
CONCEPTO_TECNO = "TECNOMECANICA"
CONCEPTO_SOAT = "SOAT"
CONCEPTOS_TS = [CONCEPTO_TECNO, CONCEPTO_SOAT]
AMB_PLACAS, AMB_GEN, AMB_TS = "Placas", "Generales", "Tecnomecánica y SOAT"
# movimientos de otro sub-rubro que se traen a la página solo si su texto o cuenta habla de SOAT o tecnomecánica
PATRON_TS = re.compile(r"SOAT|SEGURO OBLIGATORIO|TECNOMEC|TECNO MEC|REVISION TECNICO|REV TECNICO")
PENDIENTE = "PENDIENTE DE CLASIFICACIÓN"
SIN_PLACA = "SIN PLACA IDENTIFICADA"
VENTANA_FINAL = 20            # placas de formato no estándar: solo se aceptan en los últimos N caracteres del texto
CONFIANZAS = ["Alta", "Media", "Baja"]
COLS_REGLAS = ["orden", "patron", "proveedor", "concepto", "confianza", "observacion", "activa"]

# placa colombiana estándar: carro/moto/volqueta (ABC123), remolque antiguo (123ABC) y moto (ABC12D)
PLACA_ESTANDAR = re.compile(r"^(?:[A-Z]{3}\d{3}|\d{3}[A-Z]{3}|[A-Z]{3}\d{2}[A-Z])$")

# columnas que acompañan a cada movimiento (si existen) para rastrear el valor hasta SAP
COLS_DETALLE = ["clave", "fecha", "anio", "mes", "documento", "tercero", "texto", "cuenta", "cuenta_nombre", "cc",
                "cc_nombre", "codigo", "rubro", "sub_rubro_sap", "sub_rubro", "placa_detectada", "placa", "placa_estado",
                "concepto", "ambito", "concepto_sugerido", "confianza", "regla", "tipo_clasif", "origen", "llave", "valor"]


# ================================================================== texto y placas
def sin_tildes(x) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", str(x)) if unicodedata.category(c) != "Mn")


def norm_texto(x) -> str:
    """Mayúsculas, sin tildes y con espacios simples. Vacío si no hay valor."""
    if x is None or (not isinstance(x, str) and pd.isna(x)):
        return ""
    return re.sub(r"\s+", " ", sin_tildes(x).upper()).strip()


def _como_texto(x) -> str:
    """Convierte 25.0 en '25' (Excel entrega enteros como decimales) y deja el resto igual."""
    if isinstance(x, float) and x.is_integer():
        return str(int(x))
    return "" if x is None or (not isinstance(x, str) and pd.isna(x)) else str(x)


def norm_placa(x) -> str:
    """Placa en mayúsculas y sin espacios, guiones ni símbolos: 'GUQ 558 ' -> 'GUQ558'; 'ESN716*' -> 'ESN716'."""
    return re.sub(r"[^A-Z0-9]", "", sin_tildes(_como_texto(x)).upper())


def _compactar(texto: str) -> str:
    """Une placas escritas con espacio o guion ('ELQ 413', 'ELQ-413') para poder compararlas con el listado."""
    t = re.sub(r"\b([A-Z]{3})[\s\-]+(\d{3})\b", r"\1\2", texto)
    return re.sub(r"\b(\d{3})[\s\-]+([A-Z]{3})\b", r"\1\2", t)


def _patron(lista: list[str]):
    if not lista:
        return None
    lista = sorted(lista, key=len, reverse=True)
    return re.compile(r"(?<![A-Z0-9])(" + "|".join(map(re.escape, lista)) + r")(?![A-Z0-9])")


def buscador_placas(placas: list[str]):
    """Función texto -> (placa, estado). Solo reconoce placas del listado oficial, nunca por aproximación.
    - Placas estándar: se buscan en todo el texto.
    - Placas de otro formato (códigos internos, remolques…): solo en los últimos VENTANA_FINAL caracteres,
      porque un código corto puede aparecer en el texto sin ser una placa.
    - Si hay varias placas distintas se toma la última (la más cercana al final) y se marca para revisión."""
    placas = sorted({p for p in placas if p})
    re_std = _patron([p for p in placas if PLACA_ESTANDAR.match(p)])
    re_raro = _patron([p for p in placas if not PLACA_ESTANDAR.match(p)])

    def buscar(texto) -> tuple[str, str]:
        t = _compactar("" if pd.isna(texto) else str(texto).upper())
        hallazgos = [(m.start(), m.group(1)) for m in re_std.finditer(t)] if re_std else []
        if re_raro:
            hallazgos += [(m.start(), m.group(1)) for m in re_raro.finditer(t, max(0, len(t) - VENTANA_FINAL))]
        if not hallazgos:
            return SIN_PLACA, "Sin placa"
        placa = max(hallazgos)[1]
        return placa, ("Identificada" if len({p for _, p in hallazgos}) == 1
                       else "Varias placas (se tomó la última)")
    return buscar


def asignar_placas(ej: pd.DataFrame, placas: list[str]) -> pd.DataFrame:
    """Agrega las columnas placa y placa_estado a los movimientos, según la columna texto."""
    ej = ej.copy()
    if ej.empty:
        return ej.assign(placa=pd.Series(dtype=str), placa_estado=pd.Series(dtype=str))
    buscar = buscador_placas(placas)
    textos = ej["texto"].fillna("").astype(str)          # el texto vacío de SAP puede venir como NaN/None
    cache = {t: buscar(t) for t in textos.unique()}
    ej["placa"] = textos.map(lambda t: cache[t][0])
    ej["placa_estado"] = textos.map(lambda t: cache[t][1])
    return ej


# ================================================================== presupuesto (Excel horizontal -> largo)
def leer_presupuesto_hoja(crudo: pd.DataFrame, notas: dict | None = None) -> tuple[pd.DataFrame, dict]:
    """Transforma una hoja del Excel de la flota.
    Cada fila es un vehículo y cada columna un concepto; las filas sin placa son conceptos globales (In House,
    grúa, lavado…) y se conservan como globales, con el nombre del archivo, sin repartirlos entre vehículos.
    No se descarta ninguna fila, ni siquiera las que valen cero. Las cantidades de llantas y los valores unitarios
    no son dinero y no se cargan (el dinero está en los totales de llantas).
    Texto original: el archivo no trae una descripción libre por celda; se conserva el encabezado de la columna (más la
    nota de la celda, si existe, en `notas`) y, para los globales, la fila del Excel con su rótulo.
    La clasificación inicial sale de la ESTRUCTURA del archivo (fila con placa = vehículo; fila sin placa = global),
    no del nombre del concepto."""
    cols = [" ".join(str(c).split()) for c in crudo.columns]
    try:
        i_nuevas, i_reenc = cols.index("TOTAL LLANTAS NUEVAS"), cols.index("TOTAL LLANTAS REENCAUCHADAS")
        i_total = next(i for i, c in enumerate(cols) if c.startswith("TOTAL POR VEHICULO"))
        i_placa, i_tipo, i_int = cols.index("PLACA"), cols.index("TIPO VEHICULO"), 0
    except (ValueError, StopIteration) as e:
        raise ValueError("La hoja no tiene la estructura esperada: faltan las columnas PLACA, TIPO VEHICULO, "
                         "TOTAL LLANTAS NUEVAS, TOTAL LLANTAS REENCAUCHADAS o TOTAL POR VEHICULO.") from e
    idx = [i_nuevas, i_reenc] + list(range(i_reenc + 1, i_total))
    nombres = {i_nuevas: "LLANTAS NUEVAS", i_reenc: "LLANTAS REENCAUCHADAS", **{i: cols[i] for i in idx[2:]}}

    es_total = crudo.iloc[:, 0].astype(str).str.strip().str.upper() == "TOTAL"
    fila_total = crudo[es_total]
    datos = crudo[~es_total & (crudo.iloc[:, i_placa].notna() | crudo.iloc[:, i_tipo].notna())]
    valores = datos.iloc[:, idx].apply(pd.to_numeric, errors="coerce").fillna(0)
    valores.columns = [nombres[i] for i in idx]
    orden = {nombres[i]: n for n, i in enumerate(idx)}
    notas = {" ".join(str(k).split()): " ".join(str(v).split()) for k, v in (notas or {}).items()}
    originales = {nombres[i]: cols[i] + (f" · Nota del archivo: {notas[cols[i]]}" if cols[i] in notas else "") for i in idx}

    placa = datos.iloc[:, i_placa].map(norm_placa)
    es_global = (placa == "").to_numpy()
    base = pd.DataFrame({"placa": placa.to_numpy(), "numero_interno": datos.iloc[:, i_int].map(_como_texto).to_numpy(),
                         "tipo": datos.iloc[:, i_tipo].map(lambda x: " ".join(_como_texto(x).split())).to_numpy()})

    por_vehiculo = pd.concat([base[~es_global].reset_index(drop=True), valores[~es_global].reset_index(drop=True)],
                             axis=1).melt(id_vars=["placa", "numero_interno", "tipo"], var_name="concepto",
                                          value_name="valor")
    por_vehiculo = por_vehiculo.assign(es_global=0)          # se conservan también los ceros: nada se descarta
    por_vehiculo["orden"] = por_vehiculo["concepto"].map(orden)
    por_vehiculo["texto_original"] = por_vehiculo["concepto"].map(originales)

    glob = base[es_global].assign(valor=valores[es_global].sum(axis=1).to_numpy())
    filas_excel = [int(i) + 2 for i in datos.index[es_global]]              # fila 1 = encabezado
    rotulos = [str(x).strip() for x in datos.iloc[:, i_tipo][es_global]]
    glob = pd.DataFrame({"placa": "", "numero_interno": "", "tipo": "", "concepto": glob["tipo"].to_numpy(),
                         "valor": glob["valor"].to_numpy(), "es_global": 1,
                         "texto_original": [f"Fila {n} del Excel · columna TIPO VEHICULO: «{r}» (suma de sus columnas de concepto)"
                                            for n, r in zip(filas_excel, rotulos)]})
    glob["orden"] = 100 + glob.reset_index(drop=True).index
    largo = pd.concat([por_vehiculo, glob], ignore_index=True)[
        ["placa", "numero_interno", "tipo", "concepto", "orden", "valor", "es_global", "texto_original"]]

    total_col = pd.to_numeric(datos.iloc[:, i_total], errors="coerce").fillna(0)
    info = dict(
        total_archivo=float(pd.to_numeric(fila_total.iloc[0, i_total], errors="coerce")) if len(fila_total) else None,
        total_filas=float(total_col.sum()), total_transformado=float(largo["valor"].sum()),
        n_placas=int((~es_global).sum()), n_globales=int(es_global.sum()),
        globales=glob["concepto"].tolist(), n_conceptos=len(idx),
        filas_que_no_cuadran=int(((valores.sum(axis=1) - total_col).abs() > 1).sum()),
        placas=sorted(set(base.loc[~es_global, "placa"])),
        duplicadas=sorted(set(base.loc[~es_global, "placa"][base.loc[~es_global, "placa"].duplicated()])))
    if len(fila_total):
        tot = pd.to_numeric(fila_total.iloc[0, idx], errors="coerce").fillna(0).to_numpy()
        dif = (valores.sum().to_numpy() - tot)
        info["conceptos_que_no_cuadran"] = [nombres[i] for i, d in zip(idx, dif) if abs(d) > 1]
    else:
        info["conceptos_que_no_cuadran"] = []
    info["dif_total"] = info["total_transformado"] - (info["total_archivo"] if info["total_archivo"] is not None
                                                      else info["total_filas"])
    return largo, info


# ================================================================== presupuesto de Tecnomecánica y SOAT (mensual)
def _es_mes(c) -> bool:
    return isinstance(c, (pd.Timestamp, datetime, date)) or (isinstance(c, str) and not pd.isna(pd.to_datetime(c, errors="coerce"))
                                                              and bool(re.match(r"^\d{4}-\d{2}", c)))


def leer_ts(hojas: dict[str, tuple[str, pd.DataFrame]]) -> tuple[pd.DataFrame, dict]:
    """Presupuesto de Tecnomecánica y SOAT. `hojas` = {concepto: (nombre de la hoja, hoja leída con encabezado)}.
    Cada hoja: una fila por placa, una columna por mes (se conservan TODOS los meses del archivo, incluidos los
    ceros), una columna TOTAL y una fila «TOTAL MES». OJO: la columna TOTAL del archivo suma solo 2027; el UPA
    (sep–dic 2026) se lee mes a mes. Se valida cada mes contra la fila «TOTAL MES».
    Devuelve (largo: placa, concepto, anio, mes, valor, texto_original ; info)."""
    partes, info = [], {"hojas": {}, "duplicadas": [], "placas": set()}
    for concepto, (hoja, crudo) in hojas.items():
        cols = list(crudo.columns)
        i_placa = next((i for i, c in enumerate(cols) if norm_texto(c) == "PLACA"), 0)
        meses = [(i, pd.to_datetime(c)) for i, c in enumerate(cols) if _es_mes(c)]
        if not meses:
            raise ValueError(f"La hoja «{hoja}» no tiene columnas de meses (fechas) en el encabezado.")
        i_total = next((i for i, c in enumerate(cols) if norm_texto(c) == "TOTAL"), None)
        rotulo = crudo.iloc[:, i_placa].map(norm_texto)
        fila_mes = crudo[rotulo == "TOTAL MES"]
        datos = crudo[crudo.iloc[:, i_placa].notna() & (rotulo != "TOTAL MES") & (rotulo != "")]
        placa = datos.iloc[:, i_placa].map(norm_placa)
        datos, placa = datos[placa != ""], placa[placa != ""]
        val = datos.iloc[:, [i for i, _ in meses]].apply(pd.to_numeric, errors="coerce").fillna(0)
        val.columns = range(len(meses))
        largo = pd.concat([pd.DataFrame({"placa": placa.to_numpy(), "concepto": concepto,
                                          "anio": f.year, "mes": f.month, "valor": val[k].to_numpy()})
                           for k, (_, f) in enumerate(meses)], ignore_index=True)
        largo["texto_original"] = f"Hoja «{hoja}» del archivo de Tecnomecánica y SOAT"
        partes.append(largo)
        # validaciones contra el propio archivo
        tot_mes = pd.to_numeric(fila_mes.iloc[0, [i for i, _ in meses]], errors="coerce").fillna(0).to_numpy() \
            if len(fila_mes) else None
        malos = [f"{f:%Y-%m}" for (k, (_, f)), d in zip(enumerate(meses), (val.sum().to_numpy() - tot_mes)
                 if tot_mes is not None else [0] * len(meses)) if abs(d) > 1]
        ultimo_anio = max(f.year for _, f in meses)
        cols_ult = [k for k, (_, f) in enumerate(meses) if f.year == ultimo_anio]
        filas_mal = 0
        if i_total is not None:
            tot = pd.to_numeric(datos.iloc[:, i_total], errors="coerce").fillna(0)
            filas_mal = int(((val[cols_ult].sum(axis=1) - tot).abs() > 1).sum())
        info["hojas"][concepto] = dict(
            hoja=hoja, n_placas=int(len(placa)), total=float(largo["valor"].sum()),
            por_anio={int(a): float(v) for a, v in largo.groupby("anio")["valor"].sum().items()},
            meses_que_no_cuadran=malos, filas_que_no_cuadran=filas_mal, anio_total=int(ultimo_anio),
            tiene_total_mes=bool(len(fila_mes)))
        info["duplicadas"] += sorted(set(placa[placa.duplicated()]))
        info["placas"] |= set(placa)
    info["placas"] = sorted(info["placas"])
    return pd.concat(partes, ignore_index=True), info


def ts_periodo(ts_mensual: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    """Presupuesto de Tecnomecánica y SOAT del escenario, en el mismo formato largo del presupuesto de flota
    (es_global = 2): suma los meses del escenario (UPA = sep–dic 2026; Normal 2027 = ene–dic 2027)."""
    cols = ["placa", "numero_interno", "tipo", "concepto", "orden", "valor", "es_global", "texto_original"]
    if ts_mensual is None or ts_mensual.empty:
        return pd.DataFrame(columns=cols).astype({"valor": float, "orden": float, "es_global": int})
    m0, m1 = cfg["meses"]
    t = ts_mensual[(ts_mensual["anio"] == cfg["anio"]) & (ts_mensual["mes"] >= m0) & (ts_mensual["mes"] <= m1)]
    t = t.groupby(["placa", "concepto"], as_index=False).agg(valor=("valor", "sum"), texto_original=("texto_original", "first"))
    t["numero_interno"], t["tipo"], t["es_global"] = "", "", 2
    t["orden"] = t["concepto"].map({c: 200 + i for i, c in enumerate(CONCEPTOS_TS)}).fillna(299)
    return t[cols]


def candidatos_ts(ej: pd.DataFrame) -> pd.Series:
    """Movimientos cuyo texto o nombre de cuenta habla de SOAT o tecnomecánica (para traerlos aunque estén en otro sub-rubro)."""
    if ej.empty:
        return pd.Series(False, index=ej.index)
    base = ej["texto"].fillna("").astype(str) + " " + (ej["cuenta_nombre"].fillna("").astype(str) if "cuenta_nombre" in ej else "")
    cache = {t: bool(PATRON_TS.search(norm_texto(t))) for t in base.unique()}
    return base.map(cache)


def catalogo_conceptos(ppto: pd.DataFrame) -> list[str]:
    """Conceptos del presupuesto, en el orden del archivo (vehículo, luego globales, luego Tecnomecánica y SOAT).
    Tecnomecánica y SOAT siempre están, aunque aún no se haya cargado su presupuesto."""
    cat = ppto.sort_values(["es_global", "orden"]).drop_duplicates("concepto")["concepto"].tolist()
    return cat + [c for c in CONCEPTOS_TS if c not in cat]


# ================================================================== reglas de clasificación
def reglas_iniciales(catalogo: list[str]) -> pd.DataFrame:
    """Reglas de partida: SOLO los ejemplos del requerimiento, ligados a conceptos que existen en el presupuesto.
    Son preliminares (confianza Media): hay que revisarlas con los textos reales."""
    def concepto(prefijo):
        return next((c for c in catalogo if norm_texto(c).startswith(prefijo)), None)

    semillas = [
        ("SOAT|SEGURO OBLIGATORIO", "", "SOAT", "Regla estructural: el texto menciona SOAT."),
        ("TECNOMEC|TECNO MEC|REVISION TECNICO|REV TECNICO", "", "TECNOMECANICA",
         "Regla estructural: el texto menciona la revisión tecnomecánica."),
        ("", "ARDILLANTAS", "IN HOUSE LLANTAS", "Proveedor In House de llantas."),
        ("IN HOUSE+SCANIA|SCANIA IN HOUSE", "", "IN HOUSE TECNICO SCANIA", ""),
        ("REENCAUCHE|REENCAUCHADA", "", "LLANTAS REENCAUCHADAS", "Va antes de «llantas nuevas»."),
        ("LLANTA+NUEVA", "", "LLANTAS NUEVAS", ""),
        ("ACEITE+HIDRAULICO", "", "ACEITE HIDRAULICO", ""),
        ("ACEITE+MOTOR", "", "ACEITE MOTOR", ""),
        ("FILTRO|FILTRACION", "", "FILTRACION", ""),
        ("REPARACION+MOTOR", "", "CAMPANA DE DISTRIBUCION", "Concepto: campaña de distribución… reparaciones de motor."),
        ("BANDA|VALVULA", "", "FRENOS", "Ojo: «calibración de válvulas» es otro concepto; si aparece, crear una regla "
                                        "previa para ese caso."),
        ("RADIADOR|REFRIGERA", "", "REFRIGERANTE", ""),
    ]
    filas = []
    for patron, prov, pref, obs in semillas:
        c = concepto(pref)
        if c:
            filas.append(dict(orden=(len(filas) + 1) * 10, patron=patron, proveedor=prov, concepto=c,
                              confianza="Media", observacion=("Regla inicial del requerimiento, validar. " + obs).strip(),
                              activa=1))
    return pd.DataFrame(filas, columns=COLS_REGLAS)


def asegurar_reglas_ts(reglas: pd.DataFrame, catalogo: list[str]) -> pd.DataFrame:
    """Si la tabla de reglas guardada es anterior a Tecnomecánica y SOAT, antepone sus dos reglas (editables).
    Solo las agrega cuando NINGUNA regla apunta ya a esos conceptos."""
    ya = set(reglas["concepto"].dropna()) if len(reglas) else set()
    faltan = [c for c in CONCEPTOS_TS if c not in ya and c in catalogo]
    if not faltan:
        return reglas
    base = {CONCEPTO_SOAT: ("SOAT|SEGURO OBLIGATORIO", "Regla estructural: el texto menciona SOAT."),
            CONCEPTO_TECNO: ("TECNOMEC|TECNO MEC|REVISION TECNICO|REV TECNICO",
                             "Regla estructural: el texto menciona la revisión tecnomecánica.")}
    nuevas = pd.DataFrame([dict(orden=-20 + 10 * i, patron=base[c][0], proveedor="", concepto=c, confianza="Alta",
                                observacion=base[c][1], activa=1) for i, c in enumerate(faltan)], columns=COLS_REGLAS)
    return pd.concat([nuevas, reglas], ignore_index=True)


def concepto_lavado(catalogo: list[str]) -> str | None:
    """Concepto general de lavado del presupuesto («LAVADO», «SERVICIO(S) DE LAVADO»), o None si no existe."""
    return next((c for c in catalogo if norm_texto(c) in ("LAVADO", "SERVICIO DE LAVADO", "SERVICIOS DE LAVADO")), None)


# Patrón deliberadamente estricto: la comparación es por subcadena, y «LAVADO» a secas también aparece en
# «HIDROLAVADORA» (mantenimiento de equipo, otro concepto) o en «lavado de catalizador» (reparación de una pieza).
PATRON_LAVADO = ("SERVICIO DE LAVADO|SERVICIOS DE LAVADO|SERV+LAVADA|SERV+LAVADO+CAMION|SERV+LAVADO+VEHICULO")


def asegurar_regla_lavado(reglas: pd.DataFrame, catalogo: list[str]) -> pd.DataFrame:
    """Los servicios de lavado de vehículos van al concepto general de lavado. Antepone la regla (editable) solo si
    el presupuesto tiene ese concepto y NINGUNA regla apunta ya a él; si el usuario ya definió la suya, no se toca."""
    c = concepto_lavado(catalogo)
    if not c or (len(reglas) and c in set(reglas["concepto"].dropna())):
        return reglas
    nueva = pd.DataFrame([dict(orden=-5, patron=PATRON_LAVADO, proveedor="", concepto=c, confianza="Media",
                               observacion="Servicio de lavado de vehículos -> concepto de lavado. Patrón estricto: no "
                                           "incluye «hidrolavadora» ni «lavado de catalizador».", activa=1)],
                         columns=COLS_REGLAS)
    return pd.concat([nueva, reglas], ignore_index=True)


def tipo_regla(patron, proveedor) -> str:
    p, v = bool(norm_texto(patron)), bool(norm_texto(proveedor))
    return "Proveedor + texto" if p and v else ("Proveedor" if v else ("Texto" if p else "Incompleta"))


def _compilar_reglas(reglas: pd.DataFrame) -> list[dict]:
    """patron: alternativas separadas por | ; dentro de cada una, términos obligatorios separados por +.
    Ejemplo «ACEITE+MOTOR|LUBRICANTE MOTOR» = (ACEITE y MOTOR) o (LUBRICANTE MOTOR). Se compara sin tildes."""
    out = []
    r = reglas.copy()
    r["orden"] = pd.to_numeric(r["orden"], errors="coerce").fillna(9999)
    for x in r.sort_values("orden", kind="stable").to_dict("records"):
        if int(pd.to_numeric(x.get("activa"), errors="coerce") if pd.notna(x.get("activa")) else 1) != 1:
            continue
        alts = [[t.strip() for t in a.split("+") if t.strip()] for a in norm_texto(x["patron"]).split("|")]
        alts = [a for a in alts if a]
        prov = norm_texto(x["proveedor"])
        if (alts or prov) and x.get("concepto"):
            out.append(dict(orden=x["orden"], alts=alts, prov=prov, concepto=x["concepto"],
                            confianza=x.get("confianza") or "Media", tipo=tipo_regla(x["patron"], x["proveedor"])))
    return out


def firma_texto(texto, placa="") -> str:
    """«Firma» del movimiento: el texto sin la placa y con los números como #. Dos movimientos con la misma firma
    describen el mismo tipo de servicio (p. ej. «CAMBIO ACEITE MOTOR #» aunque cambie la placa o la factura)."""
    t = norm_texto(texto)
    if placa and placa != SIN_PLACA:
        t = t.replace(placa, "")
    return re.sub(r"\s+", " ", re.sub(r"\d+", "#", t)).strip()


def llave_movimiento(ej: pd.DataFrame) -> pd.Series:
    """Identificador ÚNICO y estable de cada movimiento: «compañía|clave de SAP» (o, si no hay clave,
    «compañía|fecha|documento|cuenta|valor|texto»). Si dos movimientos fueran idénticos en todo, el segundo lleva el
    sufijo #2, #3… para que cada uno tenga su propia llave. Nunca depende de la clasificación."""
    def col(c):
        return ej[c].fillna("").astype(str) if c in ej else pd.Series("", index=ej.index)
    if ej.empty:
        return pd.Series(dtype=str)
    valor = pd.to_numeric(ej["valor"], errors="coerce").fillna(0).round(2).astype(str)
    compuesta = col("fecha") + "|" + col("documento") + "|" + col("cuenta") + "|" + valor + "|" + ej["texto"].map(norm_texto)
    clave = col("clave").str.strip()
    base = clave.where(clave != "", compuesta)
    if "empresa" in ej:
        base = col("empresa") + "|" + base
    n = base.groupby(base).cumcount()
    return base.where(n == 0, base + "#" + (n + 1).astype(str))


def aplicar_placas_manuales(ej: pd.DataFrame, manual: pd.DataFrame | None) -> pd.DataFrame:
    """Guarda la placa detectada en el texto (`placa_detectada`) y aplica las placas corregidas a mano (por llave).
    La firma de aprendizaje se calcula siempre con la placa detectada, nunca con la corregida."""
    ej = ej.copy()
    ej["placa_detectada"] = ej["placa"]
    if ej.empty or manual is None or manual.empty or "placa" not in manual:
        return ej
    m = manual[manual["llave"].fillna("").ne("") & manual["placa"].fillna("").ne("")]
    por_llave = dict(zip(m["llave"], m["placa"]))
    if not por_llave:
        return ej
    nueva = (ej["llave"] if "llave" in ej else llave_movimiento(ej)).map(por_llave)
    hay = nueva.notna()
    ej.loc[hay, "placa"] = nueva[hay]
    ej.loc[hay, "placa_estado"] = "Asignada a mano"
    return ej


def construir_memoria(manual: pd.DataFrame, catalogo: list[str] | None = None):
    """Aprende del histórico de reclasificaciones (manuales y de archivo): una firma solo se usa si TODAS sus
    apariciones llevan el mismo concepto. Nivel 1 = firma + proveedor; nivel 2 = solo firma (mínimo 2 casos).
    Devuelve (nivel1, nivel2, firmas_en_conflicto)."""
    if manual is None or manual.empty:
        return {}, {}, 0
    m = manual[manual["texto_tipo"].fillna("").ne("") & manual["concepto"].fillna("").ne("")].copy()
    if catalogo is not None:
        m = m[m["concepto"].isin(catalogo)]
    m["tercero"] = m["tercero"].fillna("").map(norm_texto)
    n1, n2, conflictos = {}, {}, 0
    for clave, g in m.groupby(["texto_tipo", "tercero"]):
        cs = set(g["concepto"])
        if len(cs) == 1:
            n1[clave] = next(iter(cs))
        else:
            conflictos += 1
    for tt, g in m.groupby("texto_tipo"):
        cs = set(g["concepto"])
        if len(cs) == 1 and len(g) >= 2:
            n2[tt] = next(iter(cs))
        elif len(cs) > 1:
            conflictos += 1
    return n1, n2, conflictos


def clasificar(ej: pd.DataFrame, reglas: pd.DataFrame, manual: pd.DataFrame | None = None,
               catalogo: list[str] | None = None) -> pd.DataFrame:
    """Asigna concepto a cada movimiento. Prioridad:
    1) Manual: el movimiento fue reclasificado a mano (por su llave).
    2) Histórico: su firma (texto + proveedor) ya fue reclasificada siempre al mismo concepto.
    3) Regla: la primera de la tabla de reglas (por orden) que coincida; confianza Baja deja pendiente con sugerencia.
    4) Sin nada claro -> PENDIENTE DE CLASIFICACIÓN. Columna `origen` dice cuál de las cuatro fue."""
    ej = ej.copy()
    compiladas = _compilar_reglas(reglas)
    manual = manual if manual is not None else pd.DataFrame(columns=["llave", "concepto", "texto_tipo", "tercero"])
    ok = (lambda c: True) if catalogo is None else (lambda c: c in catalogo)
    por_llave = {r["llave"]: r["concepto"] for r in manual.to_dict("records")
                 if r.get("llave") and r.get("concepto") and ok(r["concepto"])}      # sin concepto = solo se corrigió la placa
    n1, n2, _ = construir_memoria(manual, catalogo)
    llave = (ej["llave"] if "llave" in ej else llave_movimiento(ej)) if len(ej) else pd.Series(dtype=str)
    texto_n = ej["texto"].map(norm_texto) if len(ej) else pd.Series(dtype=str)
    terc_n = ej["tercero"].map(norm_texto) if len(ej) else pd.Series(dtype=str)
    placas = ej["placa_detectada"] if "placa_detectada" in ej else (ej["placa"] if "placa" in ej else pd.Series("", index=ej.index))
    firmas = [firma_texto(t, p) for t, p in zip(ej["texto"], placas)] if len(ej) else []

    def por_reglas(t, v):
        for r in compiladas:
            if r["prov"] and r["prov"] not in v:
                continue
            if r["alts"] and not any(all(x in t for x in a) for a in r["alts"]):
                continue
            if r["confianza"] == "Baja":
                return PENDIENTE, r["concepto"], r["orden"], "Baja", r["tipo"], "Pendiente"
            return r["concepto"], "", r["orden"], r["confianza"], r["tipo"], "Regla"
        return PENDIENTE, "", None, "", "", "Pendiente"

    cache, res = {}, []
    for lk, t, v, f in zip(llave, texto_n, terc_n, firmas):
        if lk in por_llave:
            res.append((por_llave[lk], "", None, "Alta", "Reclasificación manual", "Manual"))
        elif f and (f, v) in n1:
            res.append((n1[(f, v)], "", None, "Alta", "Histórico (texto + proveedor)", "Histórico"))
        elif f and f in n2:
            res.append((n2[f], "", None, "Media", "Histórico (texto)", "Histórico"))
        else:
            if (t, v) not in cache:
                cache[(t, v)] = por_reglas(t, v)
            res.append(cache[(t, v)])
    for i, c in enumerate(["concepto", "concepto_sugerido", "regla", "confianza", "tipo_clasif", "origen"]):
        ej[c] = [r[i] for r in res]
    return ej


# ================================================================== indicadores
def estado(presupuesto, ejecutado, movimientos, concepto="") -> str:
    if concepto == PENDIENTE:
        return "Pendiente de clasificación"
    if not movimientos:
        return "Sin ejecución"
    if pd.isna(presupuesto) or presupuesto <= 0:          # sin presupuesto no hay % (evita dividir por cero)
        return "Sin presupuesto asignado"
    if abs(ejecutado - presupuesto) < 0.5:                 # igualdad al peso
        return "Presupuesto agotado"
    return "Dentro del presupuesto" if ejecutado < presupuesto else "Sobre ejecutado"


def con_metricas(df: pd.DataFrame) -> pd.DataFrame:
    """Agrega saldo, % de ejecución y estado a una tabla con presupuesto, ejecutado y movimientos."""
    df = df.copy()
    for c in ("presupuesto", "ejecutado"):                    # tablas vacías llegan como «object»: la división fallaría
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0).astype(float)
    df["saldo"] = df["presupuesto"] - df["ejecutado"]
    df["pct"] = (df["ejecutado"] / df["presupuesto"]).where(df["presupuesto"] > 0)
    conceptos = df["concepto"] if "concepto" in df else [""] * len(df)
    df["estado"] = [estado(p, e, n, c) for p, e, n, c in zip(df["presupuesto"], df["ejecutado"], df["movimientos"],
                                                              conceptos)]
    return df


def _unir(ppto: pd.DataFrame, ej: pd.DataFrame, claves: list[str]) -> pd.DataFrame:
    """Presupuesto y ejecución por las claves dadas (unión externa: aparece lo que solo tiene uno de los dos)."""
    p = ppto.groupby(claves)["valor"].sum().rename("presupuesto")
    e = ej.groupby(claves)["valor"].agg(ejecutado="sum", movimientos="size")
    if p.empty and e.empty:
        return pd.DataFrame(columns=claves + ["presupuesto", "ejecutado", "movimientos"])
    t = pd.concat([p, e], axis=1).reset_index()
    t["presupuesto"] = t["presupuesto"].fillna(0.0)          # 0 = sin presupuesto asignado
    t["ejecutado"] = t["ejecutado"].fillna(0.0)
    t["movimientos"] = t["movimientos"].fillna(0).astype(int)
    return t


def tipos_concepto(ppto: pd.DataFrame) -> dict:
    """concepto -> es_global (0 vehículo, 1 general, 2 Tecnomecánica/SOAT) según el presupuesto de TODAS las compañías.
    Sirve para que un movimiento caiga en el mismo ámbito aunque su compañía aún no tenga presupuesto cargado."""
    if ppto is None or ppto.empty:
        return {}
    t = pd.to_numeric(ppto["es_global"], errors="coerce").fillna(0).astype(int)
    return dict(zip(ppto["concepto"], t))


def _conjuntos_tipo(ppto: pd.DataFrame, tipos: dict | None):
    tipos = tipos if tipos is not None else tipos_concepto(ppto)
    generales = {c for c, t in tipos.items() if t == 1}
    ts = {c for c, t in tipos.items() if t == 2} | set(CONCEPTOS_TS)
    return generales, ts


def asignar_ambito(ej: pd.DataFrame, ppto: pd.DataFrame, tipos: dict | None = None) -> pd.DataFrame:
    """Agrega `ambito` (Placas / Generales / Tecnomecánica y SOAT) y `grupo` (Placas / Generales / Tecnomecánica / SOAT /
    Pendiente) a cada movimiento. Los dos salen de SU CONCEPTO: un movimiento tiene un solo concepto, así que cae en
    un solo lugar. `tipos` (concepto -> es_global) permite usar el catálogo de todas las compañías."""
    generales, ts = _conjuntos_tipo(ppto, tipos)
    amb = [AMB_GEN if c in generales else (AMB_TS if c in ts else AMB_PLACAS) for c in ej["concepto"]]
    nombre_ts = {CONCEPTO_TECNO: "Tecnomecánica", CONCEPTO_SOAT: "SOAT"}
    grupo = ["Pendiente" if c == PENDIENTE else (nombre_ts.get(c, AMB_TS) if a == AMB_TS else a)
             for c, a in zip(ej["concepto"], amb)]
    return ej.assign(ambito=amb, grupo=grupo)


def calcular(ppto: pd.DataFrame, ej: pd.DataFrame, oficial: pd.DataFrame, placa: str | None = None,
             tipos: dict | None = None) -> dict:
    """Todas las vistas del escenario a partir del presupuesto en formato largo (es_global 0 = vehículo, 1 = general,
    2 = Tecnomecánica/SOAT) y los movimientos ya clasificados. El ámbito de cada movimiento lo define su concepto.
    Los conceptos generales nunca se reparten entre vehículos.
    `placa`: filtra los DATOS antes de calcular (no solo la pantalla): se quedan el presupuesto y los movimientos de esa
    placa y los conceptos generales, que no se atribuyen a ninguna placa (el resumen los excluye del total de la placa)."""
    ppto = ppto.copy()
    ppto["valor"] = pd.to_numeric(ppto["valor"], errors="coerce").fillna(0.0)
    ppto["es_global"] = pd.to_numeric(ppto["es_global"], errors="coerce").fillna(0).astype(int)
    if "texto_original" not in ppto:
        ppto["texto_original"] = ""
    ppto["texto_original"] = ppto["texto_original"].fillna("")
    ej = asignar_ambito(ej.assign(valor=pd.to_numeric(ej["valor"], errors="coerce").fillna(0.0)), ppto, tipos)
    generales, ts = _conjuntos_tipo(ppto, tipos)
    if placa:
        ppto = ppto[(ppto["es_global"] == 1) | (ppto["placa"] == placa)]
        ej = ej[(ej["ambito"] == AMB_GEN) | (ej["placa"] == placa)]
    p_veh, p_glo, p_ts = (ppto[ppto["es_global"] == k] for k in (0, 1, 2))
    e_veh, e_glo, e_ts = (ej[ej["ambito"] == a] for a in (AMB_PLACAS, AMB_GEN, AMB_TS))
    original = ppto.drop_duplicates("concepto").set_index("concepto")["texto_original"]

    # por concepto (todos los ámbitos)
    pc = _unir(ppto, ej, ["concepto"])
    pc["tipo_concepto"] = [AMB_GEN if c in generales else (AMB_TS if c in ts else ("Pendiente" if c == PENDIENTE else AMB_PLACAS))
                           for c in pc["concepto"]]
    pc["texto_original"] = pc["concepto"].map(original).fillna("")
    pc = con_metricas(pc)
    pc["_o"] = pc["tipo_concepto"].map({AMB_PLACAS: 0, AMB_GEN: 1, AMB_TS: 2, "Pendiente": 3})
    pc = pc.sort_values(["_o", "presupuesto"], ascending=[True, False])[
        ["concepto", "tipo_concepto", "texto_original", "presupuesto", "ejecutado", "saldo", "pct", "estado", "movimientos"]]

    # por placa y por placa-concepto, para Placas y para Tecnomecánica y SOAT
    meta = pd.concat([p_veh[["placa", "numero_interno", "tipo"]], oficial[["placa", "numero_interno", "tipo"]],
                      p_ts[["placa", "numero_interno", "tipo"]]]).drop_duplicates("placa").set_index("placa")

    def por_placa(p, e):
        t = _unir(p, e, ["placa"]).join(meta, on="placa")
        t.loc[t["placa"] == SIN_PLACA, ["numero_interno", "tipo"]] = ["", "Sin placa identificada"]
        return con_metricas(t.sort_values("ejecutado", ascending=False))[
            ["placa", "numero_interno", "tipo", "presupuesto", "ejecutado", "saldo", "pct", "estado", "movimientos"]]

    def por_placa_concepto(p, e):
        t = _unir(p, e, ["placa", "concepto"]).join(meta[["numero_interno", "tipo"]], on="placa")
        t["texto_original"] = t["concepto"].map(original).fillna("")
        return con_metricas(t.sort_values(["placa", "presupuesto"], ascending=[True, False]))[
            ["placa", "numero_interno", "tipo", "concepto", "texto_original", "presupuesto", "ejecutado", "saldo", "pct",
             "estado", "movimientos"]]

    pp, ppc = por_placa(p_veh, e_veh), por_placa_concepto(p_veh, e_veh)
    tp, tpc = por_placa(p_ts, e_ts), por_placa_concepto(p_ts, e_ts)
    glob = pc[pc["tipo_concepto"] == AMB_GEN].drop(columns="tipo_concepto")
    tsc = pc[pc["tipo_concepto"] == AMB_TS].drop(columns="tipo_concepto")

    # nivel 1: los ámbitos (con una placa, los generales no se atribuyen: no entran al total de la placa)
    filas = [dict(ambito=AMB_PLACAS, presupuesto=p_veh["valor"].sum(), ejecutado=e_veh["valor"].sum(), movimientos=len(e_veh))]
    if not placa:
        filas.append(dict(ambito=AMB_GEN, presupuesto=p_glo["valor"].sum(), ejecutado=e_glo["valor"].sum(), movimientos=len(e_glo)))
    filas.append(dict(ambito=AMB_TS, presupuesto=p_ts["valor"].sum(), ejecutado=e_ts["valor"].sum(), movimientos=len(e_ts)))
    res = pd.DataFrame(filas)
    res.loc[len(res)] = dict(ambito=f"Total de la placa {placa}" if placa else "Total mantenimiento",
                             presupuesto=res["presupuesto"].sum(), ejecutado=res["ejecutado"].sum(),
                             movimientos=int(res["movimientos"].sum()))
    res = con_metricas(res)[["ambito", "presupuesto", "ejecutado", "saldo", "pct", "estado", "movimientos"]]

    total_e = float(ej["valor"].sum())
    conc = pd.DataFrame([
        ("Presupuesto cargado vs suma por placa + generales + Tecnomecánica/SOAT", ppto["valor"].sum(),
         pp["presupuesto"].sum() + glob["presupuesto"].sum() + tp["presupuesto"].sum()),
        ("Ejecución (movimientos) vs suma por concepto", total_e, pc["ejecutado"].sum()),
        ("Ejecución (movimientos) vs suma por placa + generales + Tecnomecánica/SOAT", total_e,
         pp["ejecutado"].sum() + glob["ejecutado"].sum() + tp["ejecutado"].sum()),
        ("Ejecución (movimientos) vs suma por placa-concepto + generales + Tecnomecánica/SOAT", total_e,
         ppc["ejecutado"].sum() + glob["ejecutado"].sum() + tpc["ejecutado"].sum()),
        ("Ejecución clasificada + pendiente vs ejecución total", ej.loc[ej["concepto"] != PENDIENTE, "valor"].sum()
         + ej.loc[ej["concepto"] == PENDIENTE, "valor"].sum(), total_e),
    ], columns=["validacion", "fuente", "calculado"])
    conc["diferencia"] = conc["fuente"] - conc["calculado"]
    return dict(resumen=res, por_concepto=pc, por_placa=pp, por_placa_concepto=ppc, globales=glob,
                ts_concepto=tsc, ts_placa=tp, ts_placa_concepto=tpc, conciliacion=conc, movimientos=ej, placa=placa)


def _arbol(pp: pd.DataFrame, ppc: pd.DataFrame, placa: str | None, conceptos: list[str] | None,
           estados: list[str] | None) -> list[dict]:
    """Un nodo por PLACA y, dentro de cada una, sus conceptos (solo los que tienen presupuesto o ejecución).
    Filtros: placa (solo ese vehículo), conceptos (recalcula los totales de cada placa con esos conceptos) y
    estados (estado de la placa)."""
    meta = pp.set_index("placa")[["numero_interno", "tipo"]]
    if conceptos:
        ppc = ppc[ppc["concepto"].isin(conceptos)]
    if placa:
        ppc = ppc[ppc["placa"] == placa]
    if conceptos or placa:
        pp = con_metricas(ppc.groupby("placa", as_index=False).agg(
            presupuesto=("presupuesto", "sum"), ejecutado=("ejecutado", "sum"), movimientos=("movimientos", "sum"))
            .join(meta, on="placa", how="left") if len(ppc) else
            pd.DataFrame(columns=["placa", "presupuesto", "ejecutado", "movimientos", "numero_interno", "tipo"]))
    if estados:
        pp = pp[pp["estado"].isin(estados)]
    cols = ["concepto", "texto_original", "presupuesto", "ejecutado", "saldo", "pct", "estado", "movimientos"]
    nodos = []
    for r in pp.sort_values(["ejecutado", "presupuesto"], ascending=False).to_dict("records"):
        h = ppc[(ppc["placa"] == r["placa"]) & ((ppc["presupuesto"] != 0) | (ppc["ejecutado"] != 0) | (ppc["movimientos"] > 0))]
        nodos.append(dict(r, nivel="placa", hijos=h.sort_values(["ejecutado", "presupuesto"], ascending=False)[cols]))
    return nodos


def arbol_placas(vistas: dict, placa: str | None = None, conceptos: list[str] | None = None,
                 estados: list[str] | None = None) -> list[dict]:
    """Pestaña Placas: placa -> conceptos de vehículo. Los conceptos generales tienen su propia pestaña."""
    return _arbol(vistas["por_placa"], vistas["por_placa_concepto"], placa, conceptos, estados)


def arbol_ts(vistas: dict, placa: str | None = None, conceptos: list[str] | None = None,
             estados: list[str] | None = None) -> list[dict]:
    """Pestaña Tecnomecánica y SOAT: placa -> [Tecnomecánica, SOAT]."""
    return _arbol(vistas["ts_placa"], vistas["ts_placa_concepto"], placa, conceptos, estados)


def validaciones(vistas: dict, oficial: pd.DataFrame, ppto: pd.DataFrame, info_extra: dict | None = None) -> pd.DataFrame:
    """Chequeos del requerimiento. ⚠ = revisar. No corrige nada."""
    ej, filas = vistas["movimientos"], []
    info_extra = info_extra or {}

    def add(nombre, ok, detalle):
        filas.append({"validacion": nombre, "resultado": "✔ Correcto" if ok else "⚠ Revisar", "detalle": detalle})

    c = vistas["conciliacion"]
    add("Presupuesto cargado = suma por placa + generales + Tecnomecánica/SOAT", abs(c.loc[0, "diferencia"]) < 1,
        f"Diferencia: {_pesos(c.loc[0, 'diferencia'])}.")
    add("Ejecución por concepto = movimientos fuente", abs(c.loc[1, "diferencia"]) < 1,
        f"Diferencia: {_pesos(c.loc[1, 'diferencia'])}.")
    add("Ejecución por placa (con «sin placa») + generales + Tecnomecánica/SOAT = total", abs(c.loc[2, "diferencia"]) < 1,
        f"Diferencia: {_pesos(c.loc[2, 'diferencia'])}.")
    add("Cada movimiento está en un solo ámbito (Placas, Generales o Tecnomecánica/SOAT)",
        abs(c.loc[3, "diferencia"]) < 1 and abs(c.loc[2, "diferencia"]) < 1,
        f"Suma de ámbitos vs total de movimientos: diferencia {_pesos(c.loc[3, 'diferencia'])}.")
    if "clave" in ej and ej["clave"].fillna("").ne("").any():
        con_clave = ej[ej["clave"].fillna("") != ""]
        cols_dup = ["empresa", "clave"] if "empresa" in con_clave else ["clave"]      # la misma clave en dos compañías no es duplicado
        dup, crit = int(con_clave.duplicated(subset=cols_dup).sum()), " + ".join(cols_dup)
    else:
        cols = [x for x in ["fecha", "documento", "cuenta", "tercero", "texto", "valor", "cc"] if x in ej]
        dup, crit = int(ej.duplicated(subset=cols).sum()), "fecha, documento, cuenta, tercero, texto, valor y cc"
    add("Movimientos duplicados", dup == 0, f"{dup} repetidos según {crit}. No se eliminan.")
    if "llave" in ej:
        llr = int(ej["llave"].duplicated().sum())
        add("Identidad de los movimientos (llave única)", llr == 0,
            f"{llr} movimientos comparten llave con otro: una reasignación los afectaría a todos.")
    if "empresa" in ej:
        por_emp = ej.groupby("empresa")["valor"].agg(["size", "sum"])
        add("Cada movimiento conserva su compañía", bool(ej["empresa"].fillna("").ne("").all()),
            " · ".join(f"{e}: {int(r['size'])} movimientos, {_pesos(r['sum'])}" for e, r in por_emp.iterrows()) or "Sin movimientos.")
    if "fuente_total" in info_extra:
        dif = info_extra["fuente_total"] - (float(ej["valor"].sum()) + info_extra.get("excluidos_valor", 0.0))
        add("Movimientos de la fuente = mostrados + excluidos", abs(dif) < 1,
            f"Fuente {_pesos(info_extra['fuente_total'])} = mostrados {_pesos(ej['valor'].sum())} + excluidos "
            f"{_pesos(info_extra.get('excluidos_valor', 0.0))} ({info_extra.get('excluidos_n', 0)} movimientos). Diferencia: {_pesos(dif)}.")
    desconocidas = sorted(set(ej["placa"]) - set(oficial["placa"]) - {SIN_PLACA})
    add("Placas de los movimientos pertenecen al listado oficial", not desconocidas,
        "Todas." if not desconocidas else "Fuera del listado (asignadas a mano): " + ", ".join(desconocidas))
    pend = ej[ej["concepto"] == PENDIENTE]
    add("Movimientos pendientes de clasificación", pend.empty,
        f"{len(pend)} movimientos por {_pesos(pend['valor'].sum())}; se muestran en Placas, bajo su placa o «{SIN_PLACA}».")
    sp = ej[(ej["placa"] == SIN_PLACA) & (ej["ambito"] != AMB_GEN)]
    add("Movimientos de placas y Tecnomecánica/SOAT sin placa identificada", sp.empty,
        f"{len(sp)} movimientos por {_pesos(sp['valor'].sum())}; se conservan como «{SIN_PLACA}».")
    vp = int((ej["placa_estado"] == "Varias placas (se tomó la última)").sum())
    add("Movimientos con varias placas en el texto", vp == 0, f"{vp} movimientos; se asignó la última placa.")
    glob = set(ppto.loc[ppto["es_global"] == 1, "concepto"])
    mal = vistas["por_placa_concepto"][vistas["por_placa_concepto"]["concepto"].isin(glob)]
    add("Conceptos generales no se reparten entre vehículos", mal.empty, "Se comparan como conceptos propios, sin repartirse entre placas.")

    # el desglose por placa de cada concepto debe sumar el total del concepto (vehículos y Tecnomecánica/SOAT)
    d = pd.concat([vistas["por_placa_concepto"], vistas["ts_placa_concepto"]])
    pcn = vistas["por_concepto"][vistas["por_concepto"]["tipo_concepto"].isin([AMB_PLACAS, AMB_TS, "Pendiente"])] \
        .set_index("concepto")[["presupuesto", "ejecutado"]]
    b = d.groupby("concepto")[["presupuesto", "ejecutado"]].sum().reindex(pcn.index).fillna(0)
    mal_d = pcn.index[((pcn - b).abs() >= 1).any(axis=1)].tolist()
    add("Desglose por placa de cada concepto = total del concepto", not mal_d,
        "Cuadra en todos los conceptos." if not mal_d else "No cuadra: " + ", ".join(mal_d))
    sin_ppto = vistas["por_concepto"]
    add("Conceptos del presupuesto en la vista por concepto", set(ppto["concepto"]) <= set(sin_ppto["concepto"]),
        f"{ppto['concepto'].nunique()} conceptos cargados ({int(ppto.loc[ppto['es_global'] == 1, 'concepto'].nunique())} "
        f"generales, {int(ppto.loc[ppto['es_global'] == 2, 'concepto'].nunique())} de Tecnomecánica/SOAT); "
        f"la vista muestra {sin_ppto['concepto'].nunique()} (incluye pendientes si los hay).")

    # calidad del presupuesto: casos sin clasificación clara, para revisión
    ofi = set(oficial["placa"])
    fuera = sorted(set(ppto.loc[ppto["es_global"] == 0, "placa"]) - ofi)
    add("Placas del presupuesto de flota en el listado oficial", not fuera,
        "Todas." if not fuera else "No están en el listado oficial (sus movimientos no se pueden identificar solos): "
                                    + ", ".join(fuera))
    fuera_ts = sorted(set(ppto.loc[ppto["es_global"] == 2, "placa"]) - ofi)
    add("Placas de Tecnomecánica/SOAT en el listado oficial", not fuera_ts,
        "Todas." if not fuera_ts else "No están en el listado oficial: " + ", ".join(fuera_ts))
    if (ppto["es_global"] == 2).any():
        sin_ts = sorted(set(ppto.loc[ppto["es_global"] == 0, "placa"]) - set(ppto.loc[ppto["es_global"] == 2, "placa"]))
        add("Placas del presupuesto de flota con presupuesto de Tecnomecánica/SOAT", not sin_ts,
            f"{len(sin_ts)} placas de la flota no aparecen en el archivo de Tecnomecánica/SOAT"
            + (": " + ", ".join(sin_ts[:25]) + ("…" if len(sin_ts) > 25 else "") if sin_ts else "."))
        solo_ts = sorted(set(ppto.loc[ppto["es_global"] == 2, "placa"]) - set(ppto.loc[ppto["es_global"] == 0, "placa"]))
        add("Placas de Tecnomecánica/SOAT que están en el presupuesto de flota", not solo_ts,
            "Todas." if not solo_ts else "Solo en Tecnomecánica/SOAT: " + ", ".join(solo_ts))
    if "fuera_flota" in ej:
        ff = ej[(ej["fuera_flota"] == True) & (ej["ambito"] == AMB_TS)]  # noqa: E712
        add("Tecnomecánica/SOAT traídos de otro sub-rubro (por su texto)", ff.empty,
            f"{len(ff)} movimientos por {_pesos(ff['valor'].sum())} vienen de un sub-rubro distinto de «{SUBRUBRO_FLOTA}»"
            + (" (ver sub_rubro_sap en Movimientos)." if len(ff) else "."))
    sug = int((ej["concepto_sugerido"] != "").sum()) if "concepto_sugerido" in ej else 0
    add("Reglas de confianza baja", sug == 0, f"{sug} movimientos con concepto sugerido pero pendientes por confianza baja.")
    if info_extra:
        add("Aprendizaje del histórico", info_extra.get("conflictos", 0) == 0,
            f"{info_extra.get('manuales', 0)} movimientos con concepto reasignado a mano · {info_extra.get('placas_manuales', 0)} "
            f"con placa asignada a mano · {info_extra.get('firmas', 0)} textos aprendidos "
            f"· {info_extra.get('conflictos', 0)} textos con conceptos contradictorios (no se usan hasta resolverlos).")
        add("Una sola asignación manual por movimiento", info_extra.get("llaves_repetidas", 0) == 0,
            f"{info_extra.get('llaves_repetidas', 0)} movimientos con más de una fila de reasignación.")
        add("Reasignaciones manuales de movimientos fuera del periodo mostrado", True,
            f"{info_extra.get('huerfanas', 0)} reasignaciones guardadas no corresponden a movimientos del periodo/corte actual "
            "(se conservan; vuelven a aplicarse si el movimiento reaparece).")
    nuevas = [x for x in set(ej["concepto"]) - {PENDIENTE} if x not in set(ppto["concepto"])]
    add("Conceptos usados tienen presupuesto en este escenario", not nuevas,
        "Todos." if not nuevas else "Sin presupuesto en este escenario: " + ", ".join(nuevas))
    return pd.DataFrame(filas).rename(columns={"validacion": "Validación", "resultado": "Resultado", "detalle": "Detalle"})


def textos_pendientes(ej: pd.DataFrame) -> pd.DataFrame:
    """Movimientos pendientes agrupados por firma de texto y proveedor, de mayor a menor valor."""
    p = ej[ej["concepto"] == PENDIENTE]
    if p.empty:
        return pd.DataFrame(columns=["texto_tipo", "tercero", "movimientos", "valor"])
    return (p.assign(texto_tipo=[firma_texto(t, pl) for t, pl in zip(p["texto"], p["placa_detectada"] if "placa_detectada" in p else p["placa"])],
                     tercero=p["tercero"].fillna(""))
            .groupby(["texto_tipo", "tercero"]).agg(movimientos=("valor", "size"), valor=("valor", "sum"))
            .reset_index().sort_values("valor", ascending=False))


def _pesos(x) -> str:
    return f"${x:,.0f}".replace(",", ".")


# ================================================================== Excel (3 hojas: VEHÍCULOS, TECNOMECÁNICAS, SOAT)
MESES_ABR = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
_ID = ["Compañía", "Grupo", "Placa", "N° interno", "Tipo de vehículo", "Concepto", "Descripción original (presupuesto)"]
_RES = ["Presupuesto", "Ejecutado", "Saldo", "% ejecución", "Estado"]
_MOV = ["Fecha", "Documento", "Proveedor", "Texto original del movimiento", "Valor movimiento", "Cuenta", "Centro de coste",
        "Sub-rubro SAP", "Clasificado por", "Placa detectada en el texto", "ID del movimiento"]


def _filas_resumen(df: pd.DataFrame, empresa: str, grupo: pd.Series | str) -> pd.DataFrame:
    """Filas «Resumen» (presupuesto, ejecución, saldo, %) de una tabla por placa-concepto o por concepto."""
    d = df.copy()
    for c in ("placa", "numero_interno", "tipo"):
        if c not in d:
            d[c] = ""
    return pd.DataFrame({
        "Compañía": empresa, "Grupo": grupo, "Placa": d["placa"].to_numpy(), "N° interno": d["numero_interno"].to_numpy(),
        "Tipo de vehículo": d["tipo"].to_numpy(), "Concepto": d["concepto"].to_numpy(),
        "Descripción original (presupuesto)": d["texto_original"].to_numpy(), "Presupuesto": d["presupuesto"].to_numpy(),
        "Ejecutado": d["ejecutado"].to_numpy(), "Saldo": d["saldo"].to_numpy(), "% ejecución": d["pct"].to_numpy(),
        "Estado": d["estado"].to_numpy()})


def _filas_movimiento(mv: pd.DataFrame, grupo: pd.Series | str, placa: pd.Series) -> pd.DataFrame:
    return pd.DataFrame({
        "Compañía": mv["empresa"].to_numpy(), "Grupo": grupo if isinstance(grupo, str) else grupo.to_numpy(),
        "Placa": placa.to_numpy(), "Concepto": mv["concepto"].to_numpy(),
        "Fecha": pd.to_datetime(mv["fecha"], errors="coerce").to_numpy(),
        "Documento": mv["documento"].to_numpy() if "documento" in mv else "", "Proveedor": mv["tercero"].to_numpy(),
        "Texto original del movimiento": mv["texto"].to_numpy(), "Valor movimiento": mv["valor"].to_numpy(),
        "Cuenta": mv["cuenta_nombre"].to_numpy() if "cuenta_nombre" in mv else "",
        "Centro de coste": mv["cc_nombre"].to_numpy() if "cc_nombre" in mv else "",
        "Sub-rubro SAP": mv["sub_rubro_sap"].to_numpy() if "sub_rubro_sap" in mv else "",
        "Clasificado por": mv["origen"].to_numpy(), "Placa detectada en el texto": mv["placa_detectada"].to_numpy(),
        "ID del movimiento": mv["llave"].to_numpy()})


def _armar_hoja(res: pd.DataFrame, mov: pd.DataFrame, columnas: list[str]) -> pd.DataFrame:
    """Cada fila «Resumen» va seguida de los movimientos que la componen (Nivel = Movimiento). Presupuesto, ejecutado y
    saldo solo están en las filas de resumen y «Valor movimiento» solo en las de movimiento: sumar cada columna no duplica."""
    if res.empty:                                           # sin resumen no hay a qué colgar movimientos
        res = pd.DataFrame(columns=["Compañía", "Grupo", "Placa", "Concepto", "N° interno", "Tipo de vehículo",
                                    "Descripción original (presupuesto)"])
        if not mov.empty:                                   # movimientos huérfanos: se listan igual, sin fila de resumen
            return mov.assign(Nivel="Movimiento").reindex(columns=columnas).reset_index(drop=True)
    res = res.reset_index(drop=True).assign(Nivel="Resumen", _n=0)
    res["_o"] = res.index
    clave = ["Compañía", "Grupo", "Placa", "Concepto"]
    m = mov.merge(res[clave + ["_o", "N° interno", "Tipo de vehículo", "Descripción original (presupuesto)"]],
                  on=clave, how="left").assign(Nivel="Movimiento", _n=1)
    out = pd.concat([res, m], ignore_index=True).sort_values(["_o", "_n", "Fecha"], kind="stable")
    return out.reindex(columns=columnas).reset_index(drop=True)


def tablas_exportacion(vistas_emp: dict, ts_mensual: pd.DataFrame, cfg: dict) -> dict[str, pd.DataFrame]:
    """Arma las tres hojas a partir de las MISMAS vistas que muestra la pantalla (una por compañía). Respeta los filtros:
    compañía (solo las que vienen en `vistas_emp`) y placa (si la vista está filtrada, los generales se excluyen)."""
    veh_r, veh_m, tec_r, tec_m, soa_r, soa_m = [], [], [], [], [], []
    m0, m1 = cfg["meses"]
    nombres_mes = [f"Ppto {MESES_ABR[m - 1]} {str(cfg['anio'])[2:]}" for m in range(m0, m1 + 1)]
    for e, v in vistas_emp.items():
        mv = v["movimientos"]
        act = lambda d: d[(d["presupuesto"] != 0) | (d["ejecutado"] != 0) | (d["movimientos"] > 0)]
        # ---- VEHÍCULOS = Placas + Generales
        ppc = act(v["por_placa_concepto"])
        veh_r.append(_filas_resumen(ppc, e, "Placas"))
        mp = mv[mv["ambito"] == AMB_PLACAS]
        veh_m.append(_filas_movimiento(mp, "Placas", mp["placa"]))
        if not v["placa"]:
            gl = act(v["globales"])
            veh_r.append(_filas_resumen(gl, e, "Generales"))
            mg = mv[mv["ambito"] == AMB_GEN]
            veh_m.append(_filas_movimiento(mg, "Generales", pd.Series("", index=mg.index)))
        # ---- Tecnomecánica y SOAT
        tpc = act(v["ts_placa_concepto"])
        for concepto, res_l, mov_l in ((CONCEPTO_TECNO, tec_r, tec_m), (CONCEPTO_SOAT, soa_r, soa_m)):
            res_l.append(_filas_resumen(tpc[tpc["concepto"] == concepto], e, concepto))
            mt = mv[(mv["ambito"] == AMB_TS) & (mv["concepto"] == concepto)]
            mov_l.append(_filas_movimiento(mt, concepto, mt["placa"]))

    def unir(listas):
        return pd.concat(listas, ignore_index=True) if listas else pd.DataFrame()

    out = {}
    out["VEHÍCULOS"] = _armar_hoja(unir(veh_r), unir(veh_m), _ID + ["Nivel"] + _RES + _MOV)
    for nombre, concepto, rl, ml in (("TECNOMECÁNICAS", CONCEPTO_TECNO, tec_r, tec_m), ("SOAT", CONCEPTO_SOAT, soa_r, soa_m)):
        res = unir(rl)
        if ts_mensual is not None and len(ts_mensual) and len(res):          # presupuesto mes a mes, en la fila de resumen
            t = ts_mensual[(ts_mensual["concepto"] == concepto) & (ts_mensual["anio"] == cfg["anio"])
                           & (ts_mensual["mes"] >= m0) & (ts_mensual["mes"] <= m1)]
            t = t.assign(col=[f"Ppto {MESES_ABR[m - 1]} {str(cfg['anio'])[2:]}" for m in t["mes"]])
            pv = t.pivot_table(index=["empresa", "placa"], columns="col", values="valor", aggfunc="sum").reset_index() \
                .rename(columns={"empresa": "Compañía", "placa": "Placa"})
            res = res.merge(pv, on=["Compañía", "Placa"], how="left")
        out[nombre] = _armar_hoja(res, unir(ml), [c for c in _ID if c != "Grupo"] + ["Nivel"] + nombres_mes + _RES + _MOV)
    return out


def escribir_excel(hojas: dict[str, pd.DataFrame]) -> bytes:
    """Escribe las hojas con encabezados, filtros, panel fijo, formatos de fecha / pesos / porcentaje y las filas de
    movimiento agrupadas (esquema) bajo su fila de resumen."""
    from openpyxl.styles import Alignment, Font, PatternFill
    azul = PatternFill("solid", fgColor="1F3A5F")
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl", datetime_format="DD/MM/YYYY") as xw:
        for nombre, df in hojas.items():
            df.to_excel(xw, sheet_name=nombre, index=False)
            ws = xw.sheets[nombre]
            ws.sheet_properties.outlinePr.summaryBelow = False
            for j, c in enumerate(df.columns, start=1):
                celda = ws.cell(row=1, column=j)
                celda.font, celda.fill = Font(bold=True, color="FFFFFF"), azul
                celda.alignment = Alignment(wrap_text=True, vertical="center")
                letra = celda.column_letter
                ancho = 60 if c in ("Texto original del movimiento", "Descripción original (presupuesto)") else \
                    (34 if c in ("ID del movimiento", "Concepto", "Proveedor", "Cuenta") else max(13, min(22, len(str(c)) + 2)))
                ws.column_dimensions[letra].width = ancho
                if c == "% ejecución":
                    fmt = "0.0%"
                elif c == "Fecha":
                    fmt = "DD/MM/YYYY"
                elif c in ("Presupuesto", "Ejecutado", "Saldo", "Valor movimiento") or str(c).startswith("Ppto "):
                    fmt = "#,##0"
                else:
                    fmt = None
                if fmt:
                    for i in range(2, len(df) + 2):
                        ws.cell(row=i, column=j).number_format = fmt
            nivel = df["Nivel"].tolist() if "Nivel" in df else []
            for i, n in enumerate(nivel, start=2):
                if n == "Movimiento":
                    ws.row_dimensions[i].outlineLevel = 1
                else:
                    for j in range(1, len(df.columns) + 1):
                        ws.cell(row=i, column=j).font = Font(bold=True)
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
    return buf.getvalue()


MESES_NOMBRE = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
_ROTULOS_MES = {"concepto": "Concepto de gasto", "placa": "Placa", "numero_interno": "N° interno", "tipo": "Tipo de vehículo",
                "ambito": "Ámbito", "presupuesto": "Presupuesto", "pct": "% ejecución", "estado": "Estado", "movimientos": "Movimientos"}


def tabla_por_mes(tabla: pd.DataFrame, mov: pd.DataFrame, claves: list[str], cfg: dict, hasta_mes: int, total: str | None = "TOTAL") -> pd.DataFrame:
    """La tabla de pantalla (presupuesto, ejecutado, saldo, %, estado) con la ejecución SEPARADA POR MES en vez del acumulado: una
    columna «Ejecución <mes>» por cada mes del escenario hasta `hasta_mes` (p. ej. «Ejecución septiembre» y «Ejecución octubre»).
    `mov` son los movimientos que componen la tabla (mes, valor y las `claves` de la tabla). «Variación» = ejecución − presupuesto
    (el mismo signo que la pantalla). Las columnas de meses suman exactamente el ejecutado de la tabla. Con `total`, agrega esa fila."""
    m0 = int(cfg["meses"][0])
    meses = list(range(m0, int(hasta_mes) + 1))
    t = tabla.copy()
    for c in claves:
        if c not in t:
            t[c] = ""
    mv = mov[(mov["mes"] >= m0) & (mov["mes"] <= hasta_mes)]
    if len(mv):
        mv = mv.assign(valor=pd.to_numeric(mv["valor"], errors="coerce").fillna(0.0))
        piv = mv.groupby(claves + ["mes"])["valor"].sum().unstack("mes")
        piv = piv.reindex(columns=meses).fillna(0.0)
        piv.columns = [f"Ejecución {MESES_NOMBRE[m - 1]}" for m in meses]
        t = t.merge(piv.reset_index(), on=claves, how="left")
    cols_mes = [f"Ejecución {MESES_NOMBRE[m - 1]}" for m in meses]
    for c in cols_mes:
        if c not in t:
            t[c] = 0.0
        t[c] = t[c].fillna(0.0).astype(float)
    t["ejecutado"] = pd.to_numeric(t["ejecutado"], errors="coerce").fillna(0.0).astype(float)
    t["presupuesto"] = pd.to_numeric(t["presupuesto"], errors="coerce").fillna(0.0).astype(float)
    t["variacion"] = t["ejecutado"] - t["presupuesto"]
    ident = [c for c in ("ambito", "placa", "numero_interno", "tipo", "concepto") if c in t.columns and (c in claves or c in tabla.columns)]
    orden = ident + ["presupuesto"] + cols_mes + ["variacion", "pct", "estado", "movimientos"]
    out = t.reindex(columns=orden)
    if total and len(out):
        fila = {c: "" for c in out.columns}
        fila[ident[0]] = total
        for c in ["presupuesto"] + cols_mes + ["variacion", "movimientos"]:
            fila[c] = float(pd.to_numeric(out[c], errors="coerce").fillna(0.0).sum())
        fila["pct"] = (fila["variacion"] + fila["presupuesto"]) / fila["presupuesto"] if fila["presupuesto"] > 0 else float("nan")
        out = pd.concat([out, pd.DataFrame([fila])], ignore_index=True)
    return out.rename(columns={**_ROTULOS_MES, "variacion": "Variación"})


def escribir_informe_mensual(hojas: dict[str, pd.DataFrame], titulo: str = "") -> bytes:
    """Excel del informe «ejecución vs presupuesto por mes»: una hoja por tabla, pesos sin decimales, % con un decimal, última fila
    (TOTAL) en negrita, encabezado fijo y filtros."""
    from openpyxl.styles import Alignment, Font, PatternFill
    azul = PatternFill("solid", fgColor="1F3A5F")
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for nombre, df in hojas.items():
            df.to_excel(xw, sheet_name=nombre[:31], index=False, startrow=2 if titulo else 0)
            ws = xw.sheets[nombre[:31]]
            f0 = 3 if titulo else 1
            if titulo:
                ws.cell(row=1, column=1, value=titulo).font = Font(bold=True, size=12)
            for j, c in enumerate(df.columns, start=1):
                h = ws.cell(row=f0, column=j)
                h.font, h.fill, h.alignment = Font(bold=True, color="FFFFFF"), azul, Alignment(wrap_text=True, vertical="center", horizontal="center")
                letra = h.column_letter
                if c in ("Concepto de gasto",):
                    ws.column_dimensions[letra].width = 58
                elif c in ("Placa", "N° interno", "Tipo de vehículo", "Ámbito", "Estado"):
                    ws.column_dimensions[letra].width = 22 if c in ("Estado", "Tipo de vehículo") else 14
                else:
                    ws.column_dimensions[letra].width = 20
                fmt = "0.0%" if c == "% ejecución" else ("#,##0" if (c in ("Presupuesto", "Variación") or str(c).startswith("Ejecución")) else
                                                        ("#,##0" if c == "Movimientos" else None))
                if fmt:
                    for i in range(f0 + 1, f0 + len(df) + 1):
                        ws.cell(row=i, column=j).number_format = fmt
            if len(df) and str(df.iloc[-1, 0]) == "TOTAL":
                for j in range(1, len(df.columns) + 1):
                    ws.cell(row=f0 + len(df), column=j).font = Font(bold=True)
            ws.freeze_panes = ws.cell(row=f0 + 1, column=2)
            ws.auto_filter.ref = f"A{f0}:{ws.cell(row=f0 + len(df), column=len(df.columns)).coordinate}"
    return buf.getvalue()


_FMT_PESOS = '"$"\\ #,##0.00;[Red]\\-"$"\\ #,##0.00'          # negativos en rojo
_ESTADO_COLOR = {"Sobre ejecutado": "FADBD8", "Sin presupuesto asignado": "FADBD8", "Pendiente de clasificación": "FFF3CD",
                 "Presupuesto agotado": "FFF3CD", "Dentro del presupuesto": "DDF3E4", "Sin ejecución": "F2F2F2"}
_TIPO_HOJA = {"Por concepto de gasto": "concepto", "Por placa": "placa", "Por placa y concepto": "placa_concepto"}
_COLOR_PESTANA = {"Resumen": "2E7D32", "Reclasificación": "2E7D32", "Por concepto de gasto": "1F3A5F", "Por placa": "1F3A5F",
                  "Por placa y concepto": "7F8C8D", "Pendientes por referencia": "C0392B"}


def escribir_informe_diseno(hojas: dict[str, pd.DataFrame], resaltar: str | None = None, resumen=None,
                            rotulo_ppto: str = "Presupuesto", extra_placa: pd.DataFrame | None = None,
                            extra_solo_sobre: tuple = ()) -> bytes:
    """Informe «ejecución vs presupuesto por mes» con el diseño del libro de seguimiento: encabezado azul, «Saldo por ejecutar
    (ppto − ejecución)» y «% del presupuesto usado» como FÓRMULAS (con barra azul de avance), estado con color, negativos en rojo,
    secciones y subtotales en «Por concepto de gasto» y total en «Por placa». Las hojas «Por concepto de gasto», «Por placa» y
    «Por placa y concepto» salen de `tabla_por_mes`; las demás se escriben tal cual (encabezado, anchos, pesos con rojo).
    `extra_placa` = columnas adicionales para «Por placa» (una fila por «Placa»); las de `extra_solo_sobre` solo se llenan en las placas con
    estado «Sobre ejecutado». `resaltar` = concepto cuya fila se pinta de amarillo (p. ej. OTROS). `resumen` = función que llena la hoja «Resumen» (ver
    `hoja_resumen_mantenimiento`); si no se da, esa hoja se escribe como una tabla simple."""
    from openpyxl import Workbook
    from openpyxl.formatting.rule import CellIsRule, DataBarRule
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter as col_l
    azul, claro, amar = PatternFill("solid", fgColor="1F3A5F"), PatternFill("solid", fgColor="E8EEF7"), PatternFill("solid", fgColor="FFF3CD")
    wb = Workbook()
    wb.remove(wb.active)

    def encabezado(ws, textos):
        for j, t in enumerate(textos, start=1):
            c = ws.cell(row=1, column=j, value=t)
            c.font, c.fill, c.alignment = Font(bold=True, color="FFFFFF"), azul, Alignment(wrap_text=True, vertical="center", horizontal="center")
        ws.row_dimensions[1].height = 36

    for nombre, df in hojas.items():
        ws = wb.create_sheet(nombre[:31])
        ws.sheet_properties.tabColor = _COLOR_PESTANA.get(nombre, "7F8C8D")
        tipo = _TIPO_HOJA.get(nombre)
        if nombre == "Resumen" and resumen is not None:
            resumen(ws)
            continue
        if tipo is None:                                                     # hoja simple
            encabezado(ws, [str(c) for c in df.columns])
            for i, fila in enumerate(df.itertuples(index=False), start=2):
                for j, v in enumerate(fila, start=1):
                    v = None if (not isinstance(v, str) and pd.isna(v)) else v
                    c = ws.cell(row=i, column=j, value=v)
                    nom_c = str(df.columns[j - 1])
                    if isinstance(v, float):
                        c.number_format = "0.0%" if nom_c.startswith("%") else ("#,##0.##" if nom_c.startswith("Cantidad") else _FMT_PESOS)
                    elif isinstance(v, str) and len(v) > 70:
                        c.alignment = Alignment(wrap_text=True, vertical="top")
            for j, c in enumerate(df.columns, start=1):
                largo = max([len(str(c))] + [len(str(x)) for x in df[c].head(200) if isinstance(x, str)] or [10])
                ws.column_dimensions[col_l(j)].width = min(max(14, largo + 2), 70 if j > 1 else 48)
            ws.freeze_panes = "A2"
            if nombre == "Detalle sobre ejecución" and len(df):
                ws.auto_filter.ref = f"A1:{col_l(len(df.columns))}{len(df) + 1}"
                for i in range(2, len(df) + 2):                      # origen con color; la fila «Resto» en gris cursiva
                    ws.cell(row=i, column=3).fill = PatternFill("solid", fgColor="DDEBF7" if ws.cell(row=i, column=3).value == "SISPO" else "FFF2CC")
                    if str(ws.cell(row=i, column=8).value).startswith("Resto"):
                        for j in range(1, len(df.columns) + 1):
                            ws.cell(row=i, column=j).font = Font(italic=True, color="595959")
                ws.freeze_panes = "D2"
            continue
        cols = list(df.columns)
        ip = cols.index("Presupuesto")
        ejec = [c for c in cols if str(c).startswith("Ejecución")]
        claves = [c for c in cols[:ip] if c != "Ámbito"]
        df = df[df[cols[0]].astype(str) != "TOTAL"] if str(df.iloc[-1, 0]) == "TOTAL" else df
        nk, ne = len(claves), len(ejec)
        c_pre, c_e0 = nk + 1, nk + 2
        c_sal, c_pct, c_est, c_mov = c_e0 + ne, c_e0 + ne + 1, c_e0 + ne + 2, c_e0 + ne + 3
        L_ = {k: col_l(v) for k, v in dict(pre=c_pre, e0=c_e0, e1=c_e0 + ne - 1, sal=c_sal, pct=c_pct).items()}
        encabezado(ws, [*claves, rotulo_ppto, *[str(c).replace("Ejecución ", "Ejecución\n") for c in ejec],
                        "Saldo por ejecutar\n(ppto − ejecución)", "% del presupuesto\nusado", "Estado", "Movimientos"])

        def fila(r, reg, total=False):
            for j, k in enumerate(claves, start=1):
                ws.cell(row=r, column=j, value=reg[k] if not pd.isna(reg[k]) else None)
            ws.cell(row=r, column=c_pre, value=float(reg["Presupuesto"]))
            for j, e in enumerate(ejec):
                ws.cell(row=r, column=c_e0 + j, value=float(reg[e]))
            p, suma = f"{L_['pre']}{r}", f"SUM({L_['e0']}{r}:{L_['e1']}{r})"
            ws.cell(row=r, column=c_sal, value=f'=IF({p}=0,"n/a",{p}-{suma})')
            ws.cell(row=r, column=c_pct, value=f"=IF({p}=0,0,{suma}/{p})")
            ws.cell(row=r, column=c_est, value=reg["Estado"] if not pd.isna(reg["Estado"]) else None)
            ws.cell(row=r, column=c_mov, value=int(reg["Movimientos"]))
            if resaltar and nk and str(reg[claves[-1]]) == resaltar:
                for j in range(1, c_mov + 1):
                    ws.cell(row=r, column=j).fill = amar

        def total(r, rotulo, filas_ref, a=None):
            """Fila de total: `filas_ref` = lista de rangos (a, b) a sumar, o de filas sueltas."""
            ws.cell(row=r, column=1, value=rotulo)
            for j in [c_pre, *range(c_e0, c_e0 + ne), c_mov]:
                L1 = col_l(j)
                ws.cell(row=r, column=j, value="=" + "+".join(f"SUM({L1}{x}:{L1}{y})" for x, y in filas_ref))
            p, suma = f"{L_['pre']}{r}", f"SUM({L_['e0']}{r}:{L_['e1']}{r})"
            ws.cell(row=r, column=c_sal, value=f"={p}-{suma}")
            ws.cell(row=r, column=c_pct, value=f'=IF({p}=0,"",{suma}/{p})')
            for j in range(1, c_mov + 1):
                ws.cell(row=r, column=j).font, ws.cell(row=r, column=j).fill = Font(bold=True), claro

        rangos_datos, r = [], 2
        if tipo == "concepto":
            subt = []
            for rotulo, amb in (("A. Conceptos de gasto atribuidos a placas", AMB_PLACAS),
                                ("B. Conceptos generales (sin placa: In House, grúa, lavado…)", AMB_GEN)):
                sub = df[df["Ámbito"] == amb]
                if sub.empty:
                    continue
                for j in range(1, c_mov + 1):
                    ws.cell(row=r, column=j).fill = claro
                ws.cell(row=r, column=1, value=rotulo).font = Font(bold=True, color="1F3A5F")
                r += 1
                a = r
                for reg in sub.to_dict("records"):
                    fila(r, reg); r += 1
                rangos_datos.append((a, r - 1))
                total(r, "Subtotal " + ("placas" if amb == AMB_PLACAS else "generales"), [(a, r - 1)]); subt.append(r); r += 2
            if subt:
                ws.cell(row=r, column=1, value="TOTAL FLOTA")
                for j in [c_pre, *range(c_e0, c_e0 + ne), c_mov]:
                    ws.cell(row=r, column=j, value="=" + "+".join(f"{col_l(j)}{x}" for x in subt))
                p, suma = f"{L_['pre']}{r}", f"SUM({L_['e0']}{r}:{L_['e1']}{r})"
                ws.cell(row=r, column=c_sal, value=f"={p}-{suma}")
                ws.cell(row=r, column=c_pct, value=f'=IF({p}=0,"",{suma}/{p})')
                for j in range(1, c_mov + 1):
                    ws.cell(row=r, column=j).font, ws.cell(row=r, column=j).fill = Font(bold=True), claro
            ws.freeze_panes = "B2"
            ultima = r
        else:
            a = r
            for reg in df.to_dict("records"):
                fila(r, reg); r += 1
            rangos_datos.append((a, r - 1))
            ultima = r - 1
            if tipo == "placa" and r > a:
                total(r, "TOTAL PLACAS", [(a, r - 1)])
                ultima = r
            ws.freeze_panes = ws.cell(row=2, column=min(nk + 1, 4))
            ult_col = c_mov
            if tipo == "placa" and extra_placa is not None and len(extra_placa):
                ex = extra_placa.set_index("Placa")
                cols_ex = [c for c in extra_placa.columns if c != "Placa"]
                ult_col = c_mov + len(cols_ex)
                for j, c in enumerate(cols_ex, start=1):
                    h = ws.cell(row=1, column=c_mov + j, value=c)
                    h.font, h.fill, h.alignment = Font(bold=True, color="FFFFFF"), azul, Alignment(wrap_text=True, vertical="center", horizontal="center")
                    texto = isinstance(ex[c].dropna().iloc[0], str) if ex[c].notna().any() else False
                    ws.column_dimensions[col_l(c_mov + j)].width = 52 if texto else 19
                for i, reg in enumerate(df.to_dict("records")):
                    rr, pl = a + i, reg[claves[0]]
                    sobre = reg["Estado"] == "Sobre ejecutado"
                    for j, c in enumerate(cols_ex, start=1):
                        if pl not in ex.index or (c in extra_solo_sobre and not sobre):
                            continue
                        v = ex.at[pl, c]
                        if isinstance(v, str) and not v:
                            continue
                        if not isinstance(v, str) and pd.isna(v):
                            continue
                        cel = ws.cell(row=rr, column=c_mov + j, value=v)
                        if not isinstance(v, str):
                            cel.number_format, cel.alignment = _FMT_PESOS, Alignment(horizontal="right")
                if r > a:                                    # total de las columnas de valor (SISPO y SAP)
                    for j, c in enumerate(cols_ex, start=1):
                        if c in ("Ejecutado SISPO", "Ejecutado SAP"):
                            x = ws.cell(row=r, column=c_mov + j, value=f"=SUM({col_l(c_mov + j)}{a}:{col_l(c_mov + j)}{r - 1})")
                            x.font, x.fill, x.number_format = Font(bold=True), claro, _FMT_PESOS
            ws.auto_filter.ref = f"A1:{col_l(ult_col)}{max(a, r - 1)}"
        # formato: pesos con rojo para negativos, barra de avance, color del estado
        for fila_ in ws.iter_rows(min_row=2, max_row=ultima):
            for c in fila_:
                if c_pre <= c.column <= c_sal:
                    c.number_format = _FMT_PESOS
                    c.alignment = Alignment(horizontal="right")
                elif c.column == c_pct:
                    c.number_format = "0%"
                elif c.column == c_mov:
                    c.number_format = "#,##0"
        for x, y in rangos_datos:
            rg_s, rg_p, rg_e = f"{L_['sal']}{x}:{L_['sal']}{y}", f"{L_['pct']}{x}:{L_['pct']}{y}", f"{col_l(c_est)}{x}:{col_l(c_est)}{y}"
            ws.conditional_formatting.add(rg_s, CellIsRule(operator="lessThan", formula=["0"], font=Font(color="C0392B", bold=True)))
            ws.conditional_formatting.add(rg_p, DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1.5, color="5B9BD5"))
            for estado, color in _ESTADO_COLOR.items():
                ws.conditional_formatting.add(rg_e, CellIsRule(operator="equal", formula=[f'"{estado}"'],
                                                               fill=PatternFill(bgColor=color, fill_type="solid")))
        anchos = {c_est: 26, c_mov: 13, c_sal: 19, c_pct: 20}
        for j in range(1, c_mov + 1):
            ws.column_dimensions[col_l(j)].width = anchos.get(j, 16.5 if j >= c_pre else (62 if claves[j - 1] == "Concepto de gasto" else
                                                                                           (11 if claves[j - 1] in ("Placa", "N° interno") else 32)))
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def hoja_resumen_mantenimiento(ws, titulo: str, subtitulo: str, meses: list[int], nombres: dict, orden: list, p, e, cfg: dict,
                               origen: list[tuple[str, dict, str]], avisos: list[tuple[str, str]], indicadores: pd.DataFrame | None = None):
    """Hoja «Resumen» del informe: 1) por sub-rubro, presupuesto del mes vs ejecución del mes y % (con color), más el presupuesto del
    escenario completo; 2) de dónde sale la ejecución de Flota cada mes; 3) lo que hay que mirar; 4) indicadores de la reclasificación.
    `p`/`e` = Series por (sub-rubro, mes) de `serie_subrubros`; `origen` = [(nombre, {mes: valor}, nota)]; `avisos` = [(texto, color)]
    con color «ambar» o «gris»; el total de cada bloque y los % son fórmulas."""
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter as col_l
    azul, claro = PatternFill("solid", fgColor="1F3A5F"), PatternFill("solid", fgColor="E8EEF7")
    gris, verde = PatternFill("solid", fgColor="F2F2F2"), PatternFill("solid", fgColor="DDF3E4")
    rojo, ambar = PatternFill("solid", fgColor="FADBD8"), PatternFill("solid", fgColor="FFF3CD")
    M0 = "#,##0"
    nm = len(meses)
    ncol = max(1 + 3 * nm + 1, 4)

    def hdr(r, textos):
        for j, t in enumerate(textos, start=1):
            c = ws.cell(row=r, column=j, value=t)
            c.font, c.fill = Font(bold=True, color="FFFFFF"), azul
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[r].height = 32

    ws.sheet_properties.tabColor = "2E7D32"
    ws["A1"] = titulo
    ws["A1"].font = Font(bold=True, size=16, color="1F3A5F")
    ws["A2"] = subtitulo
    ws["A2"].font = Font(italic=True, color="595959")
    r = 4
    ws.cell(row=r, column=1, value="1. Por sub-rubro: presupuesto del mes vs lo ejecutado en el mes").font = Font(bold=True, size=12)
    r += 1
    cab = ["Sub-rubro"]
    for m in meses:
        cab += [f"Presupuesto {MESES_NOMBRE[m - 1]}", f"Ejecución {MESES_NOMBRE[m - 1]}", f"% {MESES_NOMBRE[m - 1]}"]
    cab += [f"Presupuesto {cfg.get('etiqueta', 'del escenario')} completo ({MESES_NOMBRE[cfg['meses'][0] - 1][:3]}–{MESES_NOMBRE[cfg['meses'][1] - 1][:3]})"]
    hdr(r, cab)
    r += 1
    f0 = r
    orden = sorted(orden, key=lambda k: -float(e[k].sum()) if k in e.index.get_level_values(0) else 0.0)
    m0_, m1_ = cfg["meses"]
    for k in orden:
        ws.cell(row=r, column=1, value=nombres[k]).font = Font(bold=True)
        for i, m in enumerate(meses):
            c0 = 2 + 3 * i
            pv, ev = float(p.get((k, m), 0.0)), float(e.get((k, m), 0.0))
            ws.cell(row=r, column=c0, value=pv)
            ws.cell(row=r, column=c0 + 1, value=ev)
            a, b = col_l(c0), col_l(c0 + 1)
            pc = ws.cell(row=r, column=c0 + 2, value=f'=IF({a}{r}=0,"",{b}{r}/{a}{r})')
            pc.number_format = "0%"
            pc.fill = rojo if (pv > 0 and ev / pv > 1) else (verde if pv > 0 else gris)
        ws.cell(row=r, column=ncol, value=float(sum(p.get((k, m), 0.0) for m in range(m0_, m1_ + 1))))
        for c in [2 + 3 * i for i in range(nm)] + [3 + 3 * i for i in range(nm)] + [ncol]:
            ws.cell(row=r, column=c).number_format = M0
        r += 1
    ws.cell(row=r, column=1, value="TOTAL").font = Font(bold=True)
    ws.cell(row=r, column=1).fill = claro
    for c in [2 + 3 * i for i in range(nm)] + [3 + 3 * i for i in range(nm)] + [ncol]:
        x = ws.cell(row=r, column=c, value=f"=SUM({col_l(c)}{f0}:{col_l(c)}{r - 1})")
        x.font, x.number_format, x.fill = Font(bold=True), M0, claro
    for i in range(nm):
        c = 4 + 3 * i
        a, b = col_l(c - 2), col_l(c - 1)
        x = ws.cell(row=r, column=c, value=f'=IF({a}{r}=0,"",{b}{r}/{a}{r})')
        x.number_format, x.font, x.fill = "0%", Font(bold=True), claro
    fin_tabla = r
    r += 2
    ws.cell(row=r, column=1, value="2. Flota y Equipo de transporte: de dónde sale la ejecución de cada mes").font = Font(bold=True, size=12)
    r += 1
    hdr(r, ["Origen", *[MESES_NOMBRE[m - 1].capitalize() for m in meses], "Detalle"])
    r += 1
    a0 = r
    for nom, vals, nota in origen:
        ws.cell(row=r, column=1, value=nom)
        for j, m in enumerate(meses):
            ws.cell(row=r, column=2 + j, value=float(vals.get(m, 0.0))).number_format = M0
        ws.cell(row=r, column=2 + nm, value=nota).font = Font(italic=True, color="595959")
        r += 1
    ws.cell(row=r, column=1, value="Total Flota").font = Font(bold=True)
    ws.cell(row=r, column=1).fill = claro
    for j in range(nm):
        x = ws.cell(row=r, column=2 + j, value=f"=SUM({col_l(2 + j)}{a0}:{col_l(2 + j)}{r - 1})")
        x.font, x.number_format, x.fill = Font(bold=True), M0, claro
    r += 2
    if avisos:
        ws.cell(row=r, column=1, value="3. Lo que hay que mirar").font = Font(bold=True, size=12)
        r += 1
        for texto, color in avisos:
            c = ws.cell(row=r, column=1, value=texto)
            c.fill, c.alignment = (ambar if color == "ambar" else gris), Alignment(wrap_text=True, vertical="top")
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncol)
            ws.row_dimensions[r].height = 36
            r += 1
        r += 1
    if indicadores is not None and len(indicadores):
        ws.cell(row=r, column=1, value="4. Reclasificación a OTROS (Flota)").font = Font(bold=True, size=12)
        r += 1
        hdr(r, ["Indicador", "Valor"])
        r += 1
        for ind, val in indicadores.itertuples(index=False):
            ws.cell(row=r, column=1, value=ind)
            c = ws.cell(row=r, column=2, value=val)
            if isinstance(val, float):
                c.number_format = _FMT_PESOS
            if isinstance(val, str) and len(val) > 40:
                ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=ncol)
                c.alignment = Alignment(wrap_text=True, vertical="top")
                ws.row_dimensions[r].height = 80
            r += 1
        r += 1
    ws.column_dimensions["A"].width = 52
    for j in range(2, ncol + 1):
        ws.column_dimensions[col_l(j)].width = 19
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "B4"


def exportar_excel(vistas_emp: dict, ts_mensual: pd.DataFrame, cfg: dict) -> bytes:
    return escribir_excel(tablas_exportacion(vistas_emp, ts_mensual, cfg))


# ================================================================== Excel «Soporte de ejecución de mantenimiento»
# Réplica del archivo que hoy se arma a mano: hoja «Presupuesto» (UPA vs ejecución por sub-rubro del rubro, por mes) y una
# hoja por sub-rubro con los movimientos de SAP tal como llegaron (mismas 23 columnas).
SAP_COLUMNAS = ["Nº docum.r", "Denom.clase de coste", "Fecha de documento", "Nº docum.refer.", "Valor/mon.inf.",
                "Texto de cabecera de documento", "Cta.contrapartida", "Denominacion cuenta contrapartida", "Centro de coste",
                "Anulado", "De período", "Denom.cuenta contrapartida", "TEXTO", "Denominación del objeto", "Documento anulación",
                "Ejercicio", "Fe.contabilización", "Fecha entrada", "Nº ref.anulación", "Objeto", "Sociedad GL asociada",
                "Sociedad referencia", "Usuario"]
SAP_FECHAS = {"Fecha de documento", "Fe.contabilización", "Fecha entrada"}
# Si el movimiento no trae la fila original de SAP (columna raw), se completa con lo que sí guarda la base
SAP_RESPALDO = {"Nº docum.r": "cuenta", "Denom.clase de coste": "cuenta_nombre", "Fecha de documento": "fecha",
                "Nº docum.refer.": "documento", "Valor/mon.inf.": "valor", "Texto de cabecera de documento": "contrapartida",
                "Cta.contrapartida": "contrapartida", "Denominacion cuenta contrapartida": "tercero",
                "Centro de coste": "cc_sap", "Anulado": "anulado", "De período": "mes", "Denom.cuenta contrapartida": "tercero",
                "TEXTO": "texto", "Denominación del objeto": "cc_nombre", "Ejercicio": "anio", "Objeto": "cc"}
MESES_MAYUS = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE",
               "DICIEMBRE"]


def _raw_a_dict(x) -> dict:
    """La fila original de SAP guardada como texto (JSON o diccionario de Python) -> dict. Si no se entiende, {}."""
    if isinstance(x, dict):
        return x
    s = "" if x is None or (isinstance(x, float) and pd.isna(x)) else str(x).strip()
    if not s:
        return {}
    import ast
    import json
    for leer in (json.loads, ast.literal_eval):
        try:
            d = leer(s)
            return d if isinstance(d, dict) else {}
        except Exception:
            continue
    return {}


# Columnas que se agregan a la derecha de las 23 de SAP cuando la hoja trae detalle de SISPO
SISPO_EXTRA = ["Fuente", "Código artículo", "Cantidad", "Placa", "Concepto de mantenimiento", "Bodega"]
# De la fila de SAP que SISPO reemplazó solo se conserva lo que es de la cuenta (no lo del documento genérico)
_SAP_CUENTA = {"Nº docum.r", "Denom.clase de coste", "Sociedad GL asociada", "Sociedad referencia"}


def _vacio(v) -> bool:
    return v is None or v == "" or (isinstance(v, float) and pd.isna(v))


def numeros_cuenta(ej: pd.DataFrame) -> dict:
    """Número de cuenta de SAP («Nº docum.r») por cuenta de SISPO, tomado de los movimientos de SAP que sí traen su fila original."""
    out = {}
    clave = {sin_tildes(c).lower().replace(" ", ""): c for c in SAP_COLUMNAS}
    for r in ej.to_dict("records"):
        if not _vacio(r.get("sispo")) and bool(r.get("sispo")):
            continue
        bruto = _raw_a_dict(r.get("raw"))
        crudo = {clave.get(sin_tildes(str(k)).lower().replace(" ", "")): v for k, v in bruto.items()}
        num = crudo.get("Nº docum.r")
        if _vacio(num):                                      # en el archivo de SAP la cuenta puede venir como «Referencia»
            num = bruto.get("Referencia")
        if _vacio(num) and str(r.get("cuenta") or "") not in ("", "SISPO"):
            num = r.get("cuenta")
        nom = crudo.get("Denom.clase de coste") or r.get("cuenta_nombre")
        if _vacio(num) or _vacio(nom):
            continue
        out.setdefault(cuenta_de(nom) or str(nom), str(int(num)) if isinstance(num, float) and float(num).is_integer() else str(num))
    return out


def filas_sap(ej: pd.DataFrame, num_cuenta: dict | None = None) -> pd.DataFrame:
    """Movimientos -> tabla con las 23 columnas originales de SAP. Toma cada dato de la fila original (`raw`) y, si falta,
    de la columna equivalente de la base. «Valor/mon.inf.» siempre sale de `valor`: así la hoja suma lo mismo que la
    pestaña Presupuesto y la pantalla.
    Las filas de SISPO (`sispo` verdadero) no son un documento de SAP: salen con su propio detalle (documento «Kardex n», TEXTO =
    artículo, centro de costo de SISPO, fecha del movimiento) y solo conservan de SAP la cuenta. Si hay alguna, se agregan a la
    derecha las columnas de SISPO_EXTRA (fuente, código del artículo, cantidad, placa, concepto y bodega)."""
    clave = {sin_tildes(c).lower().replace(" ", ""): c for c in SAP_COLUMNAS}
    hay_sispo = "sispo" in ej and bool(ej["sispo"].fillna(False).astype(bool).any())
    filas = []
    for r in ej.to_dict("records"):
        crudo = {clave.get(sin_tildes(str(k)).lower().replace(" ", "")): v for k, v in _raw_a_dict(r.get("raw")).items()}
        es_sispo = bool(r.get("sispo")) if not _vacio(r.get("sispo")) else False
        fila = {}
        for c in SAP_COLUMNAS:
            v = crudo.get(c)
            if es_sispo and c not in _SAP_CUENTA:
                v = None                                     # lo heredado del genérico de SAP no aplica al detalle
            if _vacio(v):
                v = r.get(SAP_RESPALDO.get(c, ""), "") if not (es_sispo and c in _SAP_CUENTA) else ""
            fila[c] = "" if _vacio(v) else v
        if es_sispo:
            f = r.get("fecha")
            fila.update({"Fecha de documento": f, "Fe.contabilización": f, "Fecha entrada": f,
                         "Nº docum.refer.": r.get("documento", ""), "TEXTO": r.get("texto", ""),
                         "Texto de cabecera de documento": "Salida de inventario (detalle SISPO)",
                         "Centro de coste": r.get("cc", ""), "Objeto": r.get("cc", ""),
                         "Denominación del objeto": r.get("cc_nombre", ""), "De período": r.get("mes", ""),
                         "Ejercicio": r.get("anio", ""), "Usuario": "SISPO"})
            if _vacio(fila["Denom.clase de coste"]):
                fila["Denom.clase de coste"] = r.get("cuenta_sispo", "")
            if _vacio(fila["Nº docum.r"]) and num_cuenta:
                fila["Nº docum.r"] = num_cuenta.get(cuenta_de(r.get("cuenta_sispo")) or r.get("cuenta_sispo"), "")
        fila["Valor/mon.inf."] = float(r.get("valor") or 0)
        if hay_sispo:
            fila.update({"Fuente": "SISPO" if es_sispo else "SAP",
                         "Código artículo": r.get("codigo_articulo", "") if es_sispo else "",
                         "Cantidad": r.get("cantidad", "") if es_sispo else "",
                         "Placa": r.get("placa_sispo", "") if es_sispo else "",
                         "Concepto de mantenimiento": r.get("concepto_sispo", "") if es_sispo else "",
                         "Bodega": r.get("tercero", "") if es_sispo else ""})
        filas.append(fila)
    t = pd.DataFrame(filas, columns=SAP_COLUMNAS + (SISPO_EXTRA if hay_sispo else []))
    for c in SAP_FECHAS:
        t[c] = pd.to_datetime(t[c], errors="coerce")
    for c in t.columns:
        if c not in SAP_FECHAS and c not in ("Valor/mon.inf.", "Cantidad"):
            t[c] = t[c].map(lambda v: "" if _vacio(v) else (
                str(int(v)) if isinstance(v, float) and float(v).is_integer() else str(v)))
    if "Cantidad" in t:
        t["Cantidad"] = t["Cantidad"].map(lambda v: None if _vacio(v) else float(v))
    return t


def _nombre_hoja(nombre: str, usados: set) -> str:
    base = re.sub(r"[\[\]:*?/\\]", " ", str(nombre)).strip()[:31] or "Hoja"
    n, i = base, 2
    while n.lower() in usados:
        n = f"{base[:28]}_{i}"
        i += 1
    usados.add(n.lower())
    return n


def serie_subrubros(ppto: pd.DataFrame, ejec: pd.DataFrame, cfg: dict):
    """Presupuesto y ejecución del rubro por sub-rubro y mes (una sola fuente para el soporte, la pestaña Mantenimiento y
    Contenedores). `ppto` (sub_rubro, mes, valor) y `ejec` (sub_rubro, mes, valor…). Devuelve (nombres, orden, p, e) con
    `nombres` k->nombre, `orden` lista de k (alfabético) y p / e Series indexadas por (k, mes)."""
    m0, m1 = cfg["meses"]
    ppto = ppto[(ppto["mes"] >= m0) & (ppto["mes"] <= m1)]
    ejec = ejec[(ejec["mes"] >= m0) & (ejec["mes"] <= m1)]
    nombres = {}
    for d in (ppto, ejec):                                   # una fila por sub-rubro (mismo nombre sin tildes ni mayúsculas)
        for x in d["sub_rubro"].fillna("").astype(str):
            nombres.setdefault(norm_texto(x), x.strip())
    nombres.pop("", None)
    orden = sorted(nombres, key=lambda k: nombres[k].lower())
    p = ppto.assign(_k=ppto["sub_rubro"].fillna("").astype(str).map(norm_texto)).groupby(["_k", "mes"])["valor"].sum()
    e = ejec.assign(_k=ejec["sub_rubro"].fillna("").astype(str).map(norm_texto)).groupby(["_k", "mes"])["valor"].sum()
    return nombres, orden, p, e


def tabla_subrubros(ppto: pd.DataFrame, ejec: pd.DataFrame, cfg: dict, hasta_mes: int) -> pd.DataFrame:
    """Resumen por sub-rubro: presupuesto del escenario completo, ejecutado hasta el corte, saldo, % y estado."""
    nombres, orden, p, e = serie_subrubros(ppto, ejec, cfg)
    m0, _ = cfg["meses"]
    ej = ejec[(ejec["mes"] >= m0) & (ejec["mes"] <= hasta_mes)]
    n = ej.assign(_k=ej["sub_rubro"].fillna("").astype(str).map(norm_texto)).groupby("_k").size()
    filas = [dict(sub_rubro=nombres[k], presupuesto=float(p[k].sum()) if k in p.index.get_level_values(0) else 0.0,
                  ejecutado=float(e[k][e[k].index <= hasta_mes].sum()) if k in e.index.get_level_values(0) else 0.0,
                  movimientos=int(n.get(k, 0))) for k in orden]
    t = pd.DataFrame(filas, columns=["sub_rubro", "presupuesto", "ejecutado", "movimientos"])
    return con_metricas(t.assign(concepto=t["sub_rubro"]))[["sub_rubro", "presupuesto", "ejecutado", "saldo", "pct", "estado", "movimientos"]]


def tabla_mensual(ppto: pd.DataFrame, ejec: pd.DataFrame, cfg: dict, hasta_mes: int, sub_rubro: str) -> pd.DataFrame:
    """Un sub-rubro mes a mes: presupuesto, ejecutado (hasta el corte), saldo y %. Última fila = total."""
    _, _, p, e = serie_subrubros(ppto, ejec, cfg)
    k = norm_texto(sub_rubro)
    m0, m1 = cfg["meses"]
    filas = []
    for m in range(m0, m1 + 1):
        filas.append(dict(mes=m, presupuesto=float(p.get((k, m), 0.0)), ejecutado=float(e.get((k, m), 0.0)) if m <= hasta_mes else 0.0,
                          movimientos=1 if m <= hasta_mes else 0))
    t = pd.DataFrame(filas)
    t.loc[len(t)] = dict(mes=0, presupuesto=t["presupuesto"].sum(), ejecutado=t["ejecutado"].sum(), movimientos=int(t["movimientos"].sum()))
    t["saldo"] = t["presupuesto"] - t["ejecutado"]
    t["pct"] = (t["ejecutado"] / t["presupuesto"]).where(t["presupuesto"] > 0)
    return t[["mes", "presupuesto", "ejecutado", "saldo", "pct"]]


def exportar_soporte(ppto: pd.DataFrame, ejec: pd.DataFrame, cfg: dict, mes: int, rotulo: str = "UPA",
                     detalle_meses: bool = False, num_cuenta: dict | None = None) -> bytes:
    """Excel de soporte: `ppto` (sub_rubro, mes, valor) y `ejec` (movimientos con sub_rubro, mes, valor…) del MISMO rubro,
    compañía y año. Hoja «Presupuesto»: el presupuesto muestra TODOS los meses del escenario y la ejecución es la ACUMULADA
    desde el primer mes del escenario hasta `mes` (los meses posteriores quedan en cero). Hojas de cada sub-rubro: solo los
    movimientos de SAP del mes `mes`. Los totales son fórmulas, como en el archivo original.
    `num_cuenta` (opcional): número de cuenta de SAP por cuenta de SISPO, para las filas de detalle de SISPO cuyo genérico ya no está en `ejec`.
    Con `detalle_meses=True` las hojas de cada sub-rubro traen los movimientos de todos los meses hasta `mes` (con la columna «Mes»), en vez de solo los de `mes`."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter as col

    m0, m1 = cfg["meses"]
    meses = list(range(m0, m1 + 1))
    n = len(meses)
    acumulada = ejec[(ejec["mes"] >= m0) & (ejec["mes"] <= mes)]          # hoja Presupuesto
    ejec = acumulada if detalle_meses else ejec[ejec["mes"] == mes]       # hojas de detalle: el mes elegido (o todos hasta él)
    nombres, orden, p, e = serie_subrubros(ppto, acumulada, cfg)
    num_cuenta = {**numeros_cuenta(acumulada), **(num_cuenta or {})}

    wb = Workbook()
    ws = wb.active
    ws.title = "Presupuesto"
    dinero = '_-* #,##0_-;\\-* #,##0_-;_-* "-"??_-;_-@_-'
    azul = PatternFill("solid", fgColor="002060")
    c_p0, c_pt = 2, 2 + n                                    # presupuesto: meses y total
    c_e0 = c_pt + 6                                          # ejecución: deja 5 columnas libres, como el original
    c_et = c_e0 + n
    for c0, ct, titulo in ((c_p0, c_pt, rotulo), (c_e0, c_et, "EJECUCION")):
        ws.merge_cells(start_row=4, start_column=c0, end_row=4, end_column=ct)
        celda = ws.cell(row=4, column=c0, value=titulo)
        celda.fill, celda.font, celda.alignment = azul, Font(bold=True, color="FFFFFF"), Alignment(horizontal="center")
        for j, m in enumerate(meses):
            ws.cell(row=5, column=c0 + j, value=MESES_MAYUS[m - 1]).font = Font(bold=True)
            ws.cell(row=5, column=c0 + j).alignment = Alignment(horizontal="center")
        ws.cell(row=5, column=ct, value="TOTAL").font = Font(bold=True)
        ws.cell(row=5, column=ct).alignment = Alignment(horizontal="center")
    fila0 = 6
    for i, k in enumerate(orden):
        r = fila0 + i
        ws.cell(row=r, column=1, value=nombres[k])
        for j, m in enumerate(meses):
            ws.cell(row=r, column=c_p0 + j, value=float(p.get((k, m), 0.0)))
            ws.cell(row=r, column=c_e0 + j, value=float(e.get((k, m), 0.0)) if m <= mes else 0.0)
        for c0, ct in ((c_p0, c_pt), (c_e0, c_et)):
            ws.cell(row=r, column=ct, value=f"=+SUM({col(c0)}{r}:{col(ct - 1)}{r})")
    rt = fila0 + len(orden)
    ws.cell(row=rt, column=1, value="TOTALES").font = Font(bold=True)
    for c0, ct in ((c_p0, c_pt), (c_e0, c_et)):
        for c in range(c0, ct + 1):
            ws.cell(row=rt, column=c, value=f"=+SUM({col(c)}{fila0}:{col(c)}{rt - 1})" if orden else 0).font = Font(bold=True)
    for r in range(fila0, rt + 1):
        for c in list(range(c_p0, c_pt + 1)) + list(range(c_e0, c_et + 1)):
            ws.cell(row=r, column=c).number_format = dinero
    ws.column_dimensions["A"].width = 34
    for c in list(range(c_p0, c_pt + 1)) + list(range(c_e0, c_et + 1)):
        ws.column_dimensions[col(c)].width = 16
    ws.freeze_panes = "B6"

    usados = {"presupuesto"}
    for k in orden:                                          # una hoja por sub-rubro CON movimientos
        mv = ejec[ejec["sub_rubro"].fillna("").astype(str).map(norm_texto) == k]
        if mv.empty:
            continue
        mv = mv.assign(_t=mv["tercero"].fillna("").astype(str).map(norm_texto)).sort_values(["_t", "fecha", "documento"], kind="stable")
        if detalle_meses:                                    # todos los meses hasta el elegido: ordenados por mes y con su columna «Mes»
            mv = mv.assign(_t=mv["tercero"].fillna("").astype(str).map(norm_texto)).sort_values(["mes", "_t", "fecha", "documento"], kind="stable")
        t = filas_sap(mv, num_cuenta)
        columnas = list(t.columns)
        if detalle_meses:
            t = t.assign(**{"Mes": [f"{MESES_MAYUS[int(m) - 1].capitalize()} {int(a)}" for m, a in zip(mv["mes"], mv["anio"] if "anio" in mv else [cfg["anio"]] * len(mv))]})
            columnas = columnas + ["Mes"]
        hoja = wb.create_sheet(_nombre_hoja(nombres[k], usados))
        hoja.append(columnas)
        for fila in t.itertuples(index=False):
            hoja.append([None if (isinstance(v, pd.Timestamp) and pd.isna(v)) else (v.to_pydatetime() if isinstance(v, pd.Timestamp) else v)
                         for v in fila])
        for j, c in enumerate(columnas, start=1):
            hoja.cell(row=1, column=j).font = Font(bold=True)
            if c in SAP_FECHAS:
                for r in range(2, len(t) + 2):
                    hoja.cell(row=r, column=j).number_format = "mm-dd-yy"
            elif c == "Valor/mon.inf.":
                for r in range(2, len(t) + 2):
                    hoja.cell(row=r, column=j).number_format = '_-* #,##0.00_-;\\-* #,##0.00_-;_-* "-"??_-;_-@_-'
            hoja.column_dimensions[col(j)].width = {"TEXTO": 56, "Denom.cuenta contrapartida": 34, "Denominacion cuenta contrapartida": 34,
                                                    "Valor/mon.inf.": 16, "Denom.clase de coste": 26,
                                                    "Denominación del objeto": 24, "Concepto de mantenimiento": 28, "Bodega": 22}.get(c, 14)
        hoja.freeze_panes = "A2"
        hoja.auto_filter.ref = hoja.dimensions
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ================================================================== reclasificación mensual con SISPO
# SAP registra las salidas de inventario del mes como UN concepto genérico («Salida de inventario <mes> <año>», una fila por
# centro de coste). SISPO (el kardex) trae el detalle: una fila por referencia y por placa. Este bloque lee el SISPO, aplica los
# filtros acordados, identifica la placa, cruza cada referencia con la base maestra (código -> concepto de mantenimiento) y
# concilia en centavos contra lo que SAP tiene como genérico. Todo es lógica pura (sin Streamlit ni base de datos).
SISPO_CUENTAS = ("FLOTA Y EQUIPO DE TR", "MAQUINARIA Y EQUIPO", "COMBUSTIBLES Y LUBRI", "DOTACIÓN PERSONAL",
                 "ÚTILES, PAPELERÍA Y", "OTROS", "REPARACIONES LOCATIV")   # «Nombre Cta» que se reclasifican (tal como los trae SISPO)
SISPO_CUENTA = SISPO_CUENTAS[0]                 # la cuenta de la flota: es la que alimenta la página Flota
SISPO_GENERICO = "SALIDA DE INVENTARIO"         # texto (normalizado) con que SAP agrupa esas salidas
SISPO_REQUERIDAS = {                            # clave normalizada de la columna -> nombre interno
    "FECHAMOVIMIENTO": "fecha", "ARTICULO": "articulo",
    "CANTIDADSALIDA": "cantidad", "COSTOUNIDADSINIVA": "costo_unidad", "COSTOTOTALSINIVA": "costo_sin_iva",
    "COSTOTOTALCONIVA": "costo_con_iva", "CENTROCOSTO": "centro_costo", "CENTRODECOSTOS": "cc_codigo",
    "NOMBRECTA": "cuenta", "BODEGA": "bodega"}
# «IdKardex» y «Codigo Articulo» no son obligatorias al LEER (así la pantalla puede diagnosticar el archivo), pero sí para APLICAR. Si
# faltan: el IdKardex se reemplaza por «FILA-n» (n = número de fila del Excel, único y estable) y la referencia queda vacía, o sea
# PENDIENTE DE CLASIFICACIÓN. La pantalla bloquea la aplicación y `_sispo_aplicar` rechaza los identificadores «FILA-n».
SISPO_OPCIONALES = {"CIERRE": "cierre", "EMPRESAAGENCIA": "empresa_txt", "TIPOMOVIMIENTO": "tipo_mov",
                    "IDKARDEX": "id_kardex", "CODIGOARTICULO": "codigo"}
PREFIJO_ID_SINTETICO = "FILA-"


def _canon(x) -> str:
    """Solo letras y números en mayúsculas y sin tildes: compara nombres sin importar espacios, signos ni mayúsculas."""
    return re.sub(r"[^A-Z0-9]", "", sin_tildes(_como_texto(x)).upper())


def _num_co(x) -> float:
    """Número desde Excel: acepta números y textos como «$ 20.193.000,00», «-2,00» o «1,234.50»."""
    if x is None or (not isinstance(x, str) and pd.isna(x)):
        return float("nan")
    if isinstance(x, (int, float)):
        return float(x)
    t = re.sub(r"[^\d,.\-]", "", str(x))
    if t in ("", "-", ".", ","):
        return float("nan")
    if "," in t and "." in t:
        t = t.replace(".", "").replace(",", ".") if t.rfind(",") > t.rfind(".") else t.replace(",", "")
    elif "," in t:                                     # solo comas: decimal si hay una y no le siguen 3 dígitos; si no, miles
        t = t.replace(",", ".") if t.count(",") == 1 and len(t.split(",")[1]) != 3 else t.replace(",", "")
    elif t.count(".") > 1:
        t = t.replace(".", "")
    try:
        return float(t)
    except ValueError:
        return float("nan")


def _cod_articulo(x) -> str:
    t = _como_texto(x).strip()
    return t[:-2] if re.fullmatch(r"\d+\.0", t) else t


def centavos(x) -> int:
    """Pesos -> centavos enteros: las sumas y comparaciones se hacen en enteros, sin errores de coma flotante."""
    return int(round(float(x) * 100))


def sispo_leer(archivo) -> pd.DataFrame:
    """Lee el SISPO: busca la fila de encabezados (el archivo puede traer filas de filtro encima), valida que estén las columnas
    necesarias y devuelve los registros con nombres internos y números ya convertidos. Las filas vacías se descartan."""
    crudo = pd.read_excel(archivo, header=None, dtype=object)
    clave_enc = {"NOMBRECTA", "COSTOTOTALCONIVA"}                 # la fila de encabezados es la que trae estas dos columnas
    fila = next((i for i in range(min(len(crudo), 30)) if clave_enc <= {_canon(c) for c in crudo.iloc[i].tolist()}), None)
    if fila is None:
        raise ValueError("No encontré la fila de encabezados del SISPO (debe tener «Nombre Cta» y «Costo Total Con Iva»).")
    enc = [_canon(c) for c in crudo.iloc[fila].tolist()]
    datos = crudo.iloc[fila + 1:].reset_index(drop=True)
    mapa = {}
    for clave, interno in {**SISPO_REQUERIDAS, **SISPO_OPCIONALES}.items():
        if clave in enc:
            mapa[interno] = enc.index(clave)
    j = next((k for k, c in enumerate(enc) if c.startswith("NOMBRECENTRODE")), None)    # el encabezado viene cortado en el archivo
    if j is not None:
        mapa["cc_nombre"] = j
    if "codigo" not in mapa:                              # el código del artículo da el concepto: se busca también con otros nombres
        usados = set(mapa.values())
        cand = [k for k, c in enumerate(enc) if k not in usados and c and not c.endswith("ID") and (
            c in ("CODIGO", "COD", "CODART", "CODARTICULO", "REFERENCIA", "REF", "CODIGOREFERENCIA", "CODIGOPRODUCTO", "CODIGOITEM", "ITEM", "SKU")
            or (c.startswith(("CODIGO", "COD")) and any(t in c for t in ("ARTIC", "REFER", "PRODUC", "ITEM"))))]
        if len(cand) == 1:                                # si no se reconoce ninguna, la referencia queda vacía (PENDIENTE): no se exige
            mapa["codigo"] = cand[0]
    faltan = [k for k, v in SISPO_REQUERIDAS.items() if v not in mapa]
    if faltan:
        raise ValueError("Al SISPO le faltan columnas: " + ", ".join(faltan) + ".")
    d = pd.DataFrame({k: datos.iloc[:, i] for k, i in mapa.items()})
    for c in ("cierre", "empresa_txt", "tipo_mov", "cc_nombre", "codigo"):
        if c not in d:
            d[c] = ""
    if "id_kardex" in d:
        d = d[d["id_kardex"].notna() | d["cuenta"].notna()].copy()            # filas realmente vacías fuera
    else:                                                                     # sin IdKardex: id sintético = fila del Excel
        d["id_kardex"] = [f"{PREFIJO_ID_SINTETICO}{fila + 2 + k}" for k in range(len(d))]
        d = d[d["cuenta"].notna()].copy()                                     # sin id, una fila sin «Nombre Cta» está vacía
    for c in ("id_kardex", "centro_costo", "cc_codigo", "cc_nombre", "cuenta", "bodega", "articulo", "empresa_txt", "tipo_mov"):
        d[c] = d[c].map(lambda x: " ".join(_como_texto(x).split()))
    d["codigo"] = d["codigo"].map(_cod_articulo)
    d["cierre"] = d["cierre"].map(lambda x: _como_texto(x).strip()[:-2] if re.fullmatch(r"\d+\.0", _como_texto(x).strip()) else _como_texto(x).strip())
    d["fecha"] = pd.to_datetime(d["fecha"], errors="coerce")
    for c in ("cantidad", "costo_unidad", "costo_sin_iva", "costo_con_iva"):
        d[c] = d[c].map(_num_co)
    return d.reset_index(drop=True)


def sispo_sin_columnas_clave(d: pd.DataFrame) -> dict:
    """Qué columnas de referencia NO trajo el archivo: `sin_id` (todos los IdKardex son sintéticos «FILA-n») y `sin_codigo`
    (ninguna fila trae Codigo Articulo). Se calcula con los datos leídos, no con atributos que se pierden al copiar."""
    sin_id = bool(len(d)) and bool(d["id_kardex"].astype(str).str.startswith(PREFIJO_ID_SINTETICO).all())
    sin_codigo = bool(len(d)) and bool((d["codigo"].astype(str) == "").all())
    return dict(sin_id=sin_id, sin_codigo=sin_codigo)


def clase20(x) -> str:
    """Nombre de una cuenta/clase de coste comparable: SISPO y SAP lo recortan a 20 caracteres, así que se recorta igual antes de
    comparar (ignora mayúsculas, tildes, espacios y signos). La igualdad es EXACTA: «MAQUINARIA» no es «MAQUINARIA Y EQUIPO»."""
    return _canon(" ".join(_como_texto(x).split())[:20])


def cuenta_de(x, cuentas=SISPO_CUENTAS) -> str | None:
    """Nombre oficial de la cuenta (de `cuentas`) al que corresponde `x`, o None si no es ninguna de ellas."""
    k = clase20(x)
    return next((c for c in cuentas if clase20(c) == k), None)


def sispo_filtrar(d: pd.DataFrame, cuentas=(SISPO_CUENTA,)) -> tuple[pd.DataFrame, dict]:
    """Aplica los dos filtros: «Nombre Cta» entre `cuentas` y «Costo Total Con Iva» negativo. Devuelve las líneas (con el valor en
    centavos y `cuenta` con el nombre oficial) y los conteos del resumen, también por cuenta. Cada registro cae en UNA sola cuenta
    (la de su propio «Nombre Cta»); las cuentas del archivo que no están en `cuentas` no se tocan y se informan en `otras_cuentas`."""
    oficial = d["cuenta"].map(lambda x: cuenta_de(x, cuentas))
    f = d[oficial.notna()].assign(cuenta=oficial[oficial.notna()])
    sin_valor = int(f["costo_con_iva"].isna().sum())
    neg = f[f["costo_con_iva"] < 0].copy()
    neg["centavos"] = neg["costo_con_iva"].map(centavos)
    mas_dec = int(((neg["costo_con_iva"] * 100 - neg["centavos"]).abs() > 1e-6).sum())
    por = pd.DataFrame({"cuenta": list(cuentas)})
    por["registros"] = por["cuenta"].map(f["cuenta"].value_counts()).fillna(0).astype(int)
    por["salidas"] = por["cuenta"].map(neg["cuenta"].value_counts()).fillna(0).astype(int)
    por["centavos"] = por["cuenta"].map(neg.groupby("cuenta")["centavos"].sum()).fillna(0).astype("int64")
    otras = d[oficial.isna()]["cuenta"].map(lambda x: " ".join(_como_texto(x).split()) or "(sin cuenta)").value_counts()
    info = dict(total=len(d), en_cuenta=len(f), negativos=len(neg), sin_valor=sin_valor, con_mas_de_2_decimales=mas_dec,
                positivos_en_cuenta=int((f["costo_con_iva"] > 0).sum()), repetidos=int(neg["id_kardex"].duplicated().sum()),
                por_cuenta=por, otras_cuentas=otras.to_dict())
    return neg.reset_index(drop=True), info


def sispo_periodos(lineas: pd.DataFrame) -> list[int]:
    """Periodos (AAAAMM) presentes. Sale de «Cierre»; si no viniera, de la fecha del movimiento."""
    p = lineas["cierre"].where(lineas["cierre"].str.fullmatch(r"\d{6}"), lineas["fecha"].dt.strftime("%Y%m").fillna(""))
    return sorted({int(x) for x in p if str(x).isdigit()})


def sispo_empresa_probable(lineas: pd.DataFrame) -> str | None:
    """ALDC o ALSM según el texto de «EmpresaAgencia»; None si no es claro (entonces se elige a mano)."""
    t = " ".join(sorted({norm_texto(x) for x in lineas["empresa_txt"] if x}))
    if "SANTA MARTA" in t or "ALSM" in t:
        return "ALSM"
    if "D C" in t or "D.C" in t or "BOGOTA" in t or "DISTRITO CAPITAL" in t or "ALDC" in t:
        return "ALDC"
    return None


def sispo_placas(lineas: pd.DataFrame, registro: pd.DataFrame) -> pd.DataFrame:
    """Placa de cada salida, desde «Centro Costo» (formato «[ESN924] - RECOLECCION DOMICILIARIA - ESN924»). El código entre
    corchetes solo cuenta como placa si está en el listado oficial o en el presupuesto; si es el N° interno de UN solo vehículo
    se usa esa placa (se marca aparte). Códigos de centro de costos (DB010600…) o desconocidos NO se toman como placa."""
    registro = registro.assign(placa=registro["placa"].map(norm_placa), numero_interno=registro["numero_interno"].map(_como_texto))
    placas = set(registro["placa"]) - {""}
    internos = {}
    for r in registro[registro["numero_interno"] != ""].itertuples():
        internos.setdefault(r.numero_interno.strip(), set()).add(r.placa)

    def una(texto):
        m = re.match(r"\s*\[([^\]]+)\]", str(texto))
        cod = norm_placa(m.group(1)) if m else ""
        if cod in placas:
            return cod, "Identificada", cod
        ps = internos.get(m.group(1).strip(), set()) if m else set()
        if len(ps) == 1:
            return next(iter(ps)), "Identificada por N° interno", cod
        if m and "/" in m.group(1):                      # «[118ADL/5365]»: placa y N° interno pegados; vale solo si una parte es placa oficial
            for parte in m.group(1).split("/"):
                if norm_placa(parte) in placas:
                    return norm_placa(parte), "Identificada", cod
        return SIN_PLACA, ("Sin placa: código «" + cod + "» no es una placa" if cod else "Sin placa en el centro de costos"), cod

    cache = {t: una(t) for t in lineas["centro_costo"].unique()}
    out = lineas.copy()
    out["placa"] = out["centro_costo"].map(lambda t: cache[t][0])
    out["placa_estado"] = out["centro_costo"].map(lambda t: cache[t][1])
    out["cc_placa_codigo"] = out["centro_costo"].map(lambda t: cache[t][2])
    return out


def sispo_maestro_leer(archivo) -> tuple[pd.DataFrame, dict]:
    """Base maestra: Codigo Articulo, Articulo y CONCEPTO ASIGNADO. Una referencia con dos conceptos distintos es un error (no se
    adivina); las filas sin código o sin concepto se informan y no se cargan."""
    d = pd.read_excel(archivo, dtype=object)
    claves = {_canon(c): c for c in d.columns}
    faltan = [k for k in ("CODIGOARTICULO", "CONCEPTOASIGNADO") if k not in claves]
    if faltan:
        raise ValueError("A la base maestra le faltan columnas: «Codigo Articulo» y/o «CONCEPTO ASIGNADO».")
    m = pd.DataFrame({"codigo": d[claves["CODIGOARTICULO"]].map(_cod_articulo),
                      "articulo": d[claves["ARTICULO"]].map(lambda x: " ".join(_como_texto(x).split())) if "ARTICULO" in claves else "",
                      "concepto": d[claves["CONCEPTOASIGNADO"]].map(lambda x: " ".join(_como_texto(x).split()))})
    sin = m[(m["codigo"] == "") | (m["concepto"] == "")]
    m = m.drop(sin.index)
    distintos = m.groupby("codigo")["concepto"].nunique()
    conflicto = sorted(distintos[distintos > 1].index)
    info = dict(filas=len(d), sin_dato=len(sin), conflictos=conflicto, duplicados_iguales=int(m.duplicated(["codigo", "concepto"]).sum()))
    return m.drop_duplicates(["codigo", "concepto"]).reset_index(drop=True), info


def sispo_cruzar(lineas: pd.DataFrame, maestro: pd.DataFrame, catalogo: list[str] | None, ajenos: dict | None = None) -> pd.DataFrame:
    """Concepto de cada línea según la base maestra (llave: Codigo Articulo). El concepto del maestro se empata con el del
    presupuesto ignorando solo mayúsculas, tildes y signos; si la referencia no está, o su concepto no existe en el presupuesto,
    queda PENDIENTE DE CLASIFICACIÓN (nunca por aproximación).
    `catalogo` = conceptos del presupuesto DE ESA CUENTA. Si la cuenta no tiene catálogo propio (None), vale el concepto del maestro
    siempre que no sea un concepto de otra cuenta (`ajenos`: nombre canónico -> concepto), así un concepto no cruza de cuenta."""
    cat = {_canon(c): c for c in catalogo} if catalogo is not None else None
    ajenos = ajenos or {}
    por_codigo = dict(zip(maestro["codigo"], maestro["concepto"]))
    out = lineas.copy()
    conc, est = [], []
    for cod in out["codigo"]:
        crudo = por_codigo.get(cod) if cod else None
        if not cod:
            conc.append(PENDIENTE); est.append("Sin Codigo Articulo en el archivo")
        elif crudo is None:
            conc.append(PENDIENTE); est.append("Referencia no está en la base maestra")
        elif cat is not None and _canon(crudo) in cat:
            conc.append(cat[_canon(crudo)]); est.append("Asignado")
        elif cat is not None:
            conc.append(PENDIENTE); est.append("Concepto del maestro sin presupuesto: " + crudo)
        elif _canon(crudo) in ajenos:
            conc.append(PENDIENTE); est.append("Concepto de otra cuenta del presupuesto: " + crudo)
        else:
            conc.append(" ".join(crudo.split())); est.append("Asignado")
    out["concepto"], out["concepto_estado"] = conc, est
    return out


def sispo_cruzar_cuentas(lineas: pd.DataFrame, maestro: pd.DataFrame, catalogos: dict) -> pd.DataFrame:
    """Cruza con la base maestra respetando la cuenta de cada línea. `catalogos`: cuenta oficial -> lista de conceptos de su
    presupuesto, o vacío/None si la cuenta no tiene catálogo propio. Una cuenta con catálogo solo acepta conceptos de SU catálogo;
    una sin catálogo acepta el del maestro salvo que pertenezca al catálogo de otra cuenta. Conserva el orden y el índice."""
    if lineas.empty:
        return lineas.assign(concepto=pd.Series(dtype=str), concepto_estado=pd.Series(dtype=str))
    out = lineas.copy()
    out["concepto"], out["concepto_estado"] = "", ""
    for cta in out["cuenta"].unique():
        propio = catalogos.get(cta) or None
        ajenos = {_canon(x): x for c2, v in catalogos.items() if c2 != cta and v for x in v}
        m = out["cuenta"] == cta
        r = sispo_cruzar(out[m], maestro, propio, ajenos)
        out.loc[m, "concepto"], out.loc[m, "concepto_estado"] = r["concepto"], r["concepto_estado"]
    return out


CLASE_STAF_OTROS = "OTROS"
OBJETO_STAF_OTROS = "Staf Mantenimiento"


def es_staf_otros(df: pd.DataFrame) -> pd.Series:
    """Movimientos de SAP cuya «Denom.clase de coste» es OTROS (`cuenta_nombre`, o `sub_rubro_sap` si ahí no viene) y cuya
    «Denominación del objeto» (`cc_nombre`) es «Staf Mantenimiento». El presupuesto de flota no los trae en su sub-rubro, pero son
    gasto del área de mantenimiento: la pantalla Mantenimiento los suma a «Flota y Equipo de transporte». La comparación ignora
    mayúsculas, tildes, espacios y signos. Devuelve una serie booleana alineada con `df`."""
    if df is None or len(df) == 0:
        return pd.Series(False, index=getattr(df, "index", None), dtype=bool)
    vacio = pd.Series("", index=df.index)
    clase = df["cuenta_nombre"] if "cuenta_nombre" in df else vacio
    clase_sap = df["sub_rubro_sap"] if "sub_rubro_sap" in df else vacio
    obj = df["cc_nombre"] if "cc_nombre" in df else vacio
    k_otros, k_obj = _canon(CLASE_STAF_OTROS), _canon(OBJETO_STAF_OTROS)
    es_clase = clase.map(_canon).eq(k_otros) | (clase.map(_canon).eq("") & clase_sap.map(_canon).eq(k_otros))
    return (es_clase & obj.map(_canon).eq(k_obj)).astype(bool)


def ids_staf_otros_raw(crudo: pd.DataFrame) -> set:
    """Respaldo de `es_staf_otros` cuando las columnas de la base no traen la clase o el objeto: mira la fila original de SAP (`raw`,
    columnas «Denom.clase de coste» y «Denominación del objeto»). `crudo` trae `id` y `raw`. Devuelve los `id` que cumplen."""
    ids, k_cl, k_ob = set(), _canon(CLASE_STAF_OTROS), _canon(OBJETO_STAF_OTROS)
    if crudo is None or crudo.empty:
        return ids
    for i, r in zip(crudo["id"], crudo["raw"]):
        d = {_canon(k): v for k, v in _raw_a_dict(r).items()}
        if _canon(d.get("DENOMCLASEDECOSTE")) == k_cl and _canon(d.get("DENOMINACIONDELOBJETO")) == k_ob:
            ids.add(i)
    return ids


def ids_staf_otros_ejec(leer, ejec: pd.DataFrame, anio: int, empresa: str) -> set:
    """`id` de las líneas de `ejec` que son OTROS + «Staf Mantenimiento» (ver `es_staf_otros`). Primero por las columnas de la base; si
    así no sale ninguna, por la fila original de SAP (`raw`), consultando solo las que mencionan Staf Mantenimiento. `leer` es
    `ctx.leer`. Lo usan Mantenimiento (las suma a Flota) y Operaciones (las quita), para que ambas decidan igual."""
    if ejec is None or len(ejec) == 0 or "id" not in ejec:
        return set()
    ids = set(ejec.loc[es_staf_otros(ejec).to_numpy(), "id"])
    if ids:
        return ids
    crudo = leer("SELECT id, raw FROM ejecucion WHERE anio=? AND empresa=? AND LOWER(CAST(raw AS TEXT)) LIKE ?",
                 (int(anio), empresa, "%staf%mantenimiento%"))
    return ids_staf_otros_raw(crudo)


def cuenta_sap(pool: pd.DataFrame) -> pd.Series:
    """Cuenta de SISPO a la que corresponde cada movimiento de SAP, por su clase de coste («Denom.clase de coste»): se lee de
    `cuenta_nombre` y, si ahí no hay una de las cuentas, de `sub_rubro_sap`. Una fila pertenece a UNA sola cuenta (None si a
    ninguna), así un movimiento nunca queda en dos."""
    vacio = pd.Series("", index=pool.index)
    cn = pool["cuenta_nombre"].fillna("") if "cuenta_nombre" in pool else vacio
    ss = pool["sub_rubro_sap"].fillna("") if "sub_rubro_sap" in pool else vacio
    tabla = {v: cuenta_de(v) for v in set(cn) | set(ss)}
    a, b = cn.map(tabla), ss.map(tabla)
    return a.where(a.notna(), b)


def sispo_sap_generico(pool: pd.DataFrame, empresa: str, periodo: int, cuenta: str = SISPO_CUENTA) -> pd.DataFrame:
    """Movimientos de SAP (ya con su llave) que SISPO viene a reemplazar para UNA cuenta: texto «Salida de inventario…», de esa
    compañía y mes, cuya clase de coste de SAP es la cuenta de SISPO (ver `cuenta_sap`). Es el vínculo cuenta-SISPO ↔ SAP: no se
    mira el sub-rubro del presupuesto, porque la homologación puede juntar varias clases en uno.
    Incluye anulados y reversos: su suma neta es lo que SAP realmente tiene cargado."""
    if pool.empty:
        return pool
    anio, mes = divmod(int(periodo), 100)
    es = ((pool["empresa"] == empresa) & (pd.to_numeric(pool["mes"], errors="coerce") == mes)
          & (pd.to_numeric(pool["anio"], errors="coerce") == anio)
          & pool["texto"].map(norm_texto).str.contains(SISPO_GENERICO, regex=False))
    if "fuera_flota" in pool:
        es &= ~pool["fuera_flota"].fillna(False).astype(bool)
    es &= cuenta_sap(pool) == cuenta
    return pool[es].copy()


def sispo_conciliar(lineas: pd.DataFrame, sap: pd.DataFrame) -> dict:
    """Conciliación EXACTA en centavos, solo por VALOR. SAP carga el costo en positivo y SISPO lo trae negativo (salida): el total de
    SAP debe ser igual a −SISPO y lo distribuido por concepto debe sumar lo mismo. No se compara por centro de coste."""
    sap_c = [centavos(v) for v in pd.to_numeric(sap["valor"], errors="coerce").fillna(0.0)] if len(sap) else []
    t_sispo, t_sap = int(lineas["centavos"].sum()) if len(lineas) else 0, int(sum(sap_c))
    dist = lineas.groupby("concepto")["centavos"].sum() if len(lineas) else pd.Series(dtype="int64")
    cuadra_total = len(sap) > 0 and t_sap + t_sispo == 0 and (int(dist.sum()) if len(dist) else 0) == t_sispo
    return dict(sap_c=t_sap, sispo_c=t_sispo, distribuido_c=int(dist.sum()) if len(dist) else 0, diferencia_c=t_sap + t_sispo,
                n_sap=len(sap), n_sispo=len(lineas),
                cuadra=cuadra_total)


def sispo_conciliar_cuentas(lineas: pd.DataFrame, pool: pd.DataFrame, empresa: str, periodo: int | None,
                            cuentas=SISPO_CUENTAS) -> dict:
    """UNA función para todas las cuentas: por cada una toma sus líneas de SISPO, sus movimientos genéricos de SAP y concilia
    (misma regla exacta en centavos que Flota). Devuelve {cuenta: dict(lin, sap, conc, estado)}; `estado` es «Cuadra»,
    «No cuadra», «Sin genérico en SAP» o «Sin salidas en SISPO». Nada se asigna a una cuenta que no sea la de la línea."""
    out = {}
    vacio = pool.iloc[0:0] if len(pool) else pool
    for cta in cuentas:
        lin = lineas[lineas["cuenta"] == cta] if len(lineas) else lineas
        sap = sispo_sap_generico(pool, empresa, periodo, cta) if periodo and len(pool) else vacio
        sap = sap.assign(valor=pd.to_numeric(sap["valor"], errors="coerce").fillna(0.0)) if len(sap) else sap
        conc = sispo_conciliar(lin, sap) if len(lin) else dict(sap_c=int(sum(centavos(v) for v in sap["valor"])) if len(sap) else 0,
                                                               sispo_c=0, distribuido_c=0, n_sap=len(sap), n_sispo=0, cuadra=False,
                                                               diferencia_c=int(sum(centavos(v) for v in sap["valor"])) if len(sap) else 0)
        estado = ("Sin salidas en SISPO" if not len(lin) else "Sin genérico en SAP" if not len(sap) else
                  "Cuadra" if conc["cuadra"] else "No cuadra")
        out[cta] = dict(lin=lin, sap=sap, conc=conc, estado=estado)
    return out


def sispo_validar_cuentas(lineas: pd.DataFrame, original: pd.DataFrame, por_cuenta: dict) -> list[str]:
    """Controles que cruzan cuentas; devuelve los problemas (lista vacía = todo bien):
    · ningún IdKardex en más de una línea, y por tanto en más de una cuenta (nada se cuenta dos veces);
    · ningún movimiento de SAP reemplazado por dos cuentas;
    · lo asignado a cada cuenta es EXACTAMENTE lo que el archivo original tiene en esa cuenta (cantidad de salidas y centavos),
      y sus líneas conservan su valor, cantidad y referencia originales."""
    problemas = []
    if len(lineas) and lineas["id_kardex"].duplicated().any():
        problemas.append(f"{int(lineas['id_kardex'].duplicated().sum())} IdKardex repetidos: un movimiento estaría contado más de una vez.")
    llaves = [set(v["sap"]["llave"]) for v in por_cuenta.values() if len(v["sap"]) and "llave" in v["sap"]]
    for i in range(len(llaves)):
        for j in range(i + 1, len(llaves)):
            if llaves[i] & llaves[j]:
                problemas.append(f"{len(llaves[i] & llaves[j])} movimientos de SAP serían reemplazados por dos cuentas a la vez.")
    orig = original[original["costo_con_iva"] < 0].copy()
    orig["cuenta_of"] = orig["cuenta"].map(lambda x: cuenta_de(x, tuple(por_cuenta)))
    orig = orig[orig["cuenta_of"].notna()]
    for cta, v in por_cuenta.items():
        o = orig[orig["cuenta_of"] == cta]
        l = v["lin"]
        if len(o) != len(l) or sum(centavos(x) for x in o["costo_con_iva"]) != (int(l["centavos"].sum()) if len(l) else 0):
            problemas.append(f"{cta}: lo asignado ({len(l)} salidas) no coincide con el archivo original ({len(o)} salidas).")
        elif len(l):
            a = o.set_index("id_kardex")[["codigo", "cantidad", "costo_con_iva"]]
            b = l.set_index("id_kardex")[["codigo", "cantidad", "costo_con_iva"]]
            if not a.sort_index().equals(b.sort_index()):
                problemas.append(f"{cta}: referencia, cantidad o valor de alguna salida cambió respecto al archivo original.")
    return problemas


def _moda(x) -> str:
    x = x.dropna()
    m = x.mode()
    return m.iat[0] if len(m) else ""


def sispo_destinos(lin: pd.DataFrame, rep: pd.DataFrame) -> pd.DataFrame:
    """Dónde cae cada línea de SISPO en el presupuesto y quién la ve: hereda de los movimientos de SAP que reemplaza (de la MISMA
    compañía y cuenta) el código/sub-rubro del presupuesto y la gerencia. Primero se mira el mismo centro de coste; si SAP lo cargó
    en otro, el destino que pesa más (por valor) y la gerencia predominante. Devuelve codigo, sub_rubro, sub_rubro_sap y gerencia,
    con el mismo índice de `lin`. Un SAP antiguo sin sub-rubro guardado se toma como el de la flota."""
    cols = ["codigo", "sub_rubro", "sub_rubro_sap"]
    r = rep.copy()
    for c in (*cols, "gerencia", "cc"):
        if c not in r:
            r[c] = ""
    for c in cols:
        r[c] = r[c].fillna("").astype(str)
    r["sub_rubro"] = r["sub_rubro"].where(r["sub_rubro"] != "", SUBRUBRO_FLOTA)
    r["sub_rubro_sap"] = r["sub_rubro_sap"].where(r["sub_rubro_sap"] != "", r["sub_rubro"])
    r["_k"] = r["cc"].fillna("").astype(str).str.upper()
    r["_a"] = pd.to_numeric(r["valor"], errors="coerce").abs().fillna(0.0) if "valor" in r else 0.0

    def mejor(g):
        t = g.groupby(cols)["_a"].sum().reset_index().sort_values(["_a", *cols], ascending=[False, True, True, True])
        return tuple(t.iloc[0][cols]) if len(t) else ("", SUBRUBRO_FLOTA, SUBRUBRO_FLOTA)
    por_cc = {k: mejor(g) for k, g in r.groupby("_k")}
    glob = mejor(r)
    ger_cc = r.groupby("_k")["gerencia"].agg(_moda).to_dict()
    ger_glob = _moda(r["gerencia"])
    k = lin["cc_codigo"].fillna("").astype(str).str.upper()
    dest = [por_cc.get(x, glob) for x in k]
    return pd.DataFrame({"codigo": [d[0] for d in dest], "sub_rubro": [d[1] for d in dest], "sub_rubro_sap": [d[2] for d in dest],
                         "gerencia": [ger_cc.get(x, ger_glob) for x in k]}, index=lin.index)


def sispo_cuadre_destinos(lin: pd.DataFrame, rep: pd.DataFrame) -> pd.DataFrame:
    """Por destino del presupuesto (código + sub-rubro): lo que SAP tenía como genérico contra lo que quedaría con el detalle de
    SISPO. Si una cuenta tuviera genéricos en más de un destino y el reparto por centro de coste no los deja iguales, el presupuesto
    de ese destino cambiaría: la diferencia sale aquí (en centavos) para no aplicar nada así."""
    d = sispo_destinos(lin, rep)
    r = rep.copy()
    for c in ("codigo", "sub_rubro"):
        r[c] = r[c].fillna("").astype(str) if c in r else ""
    r["sub_rubro"] = r["sub_rubro"].where(r["sub_rubro"] != "", SUBRUBRO_FLOTA)
    sap = r.assign(c=[centavos(v) for v in pd.to_numeric(r["valor"], errors="coerce").fillna(0.0)]).groupby(["codigo", "sub_rubro"])["c"].sum()
    sis = pd.Series(-lin["centavos"].to_numpy(), index=pd.MultiIndex.from_frame(d[["codigo", "sub_rubro"]])).groupby(level=[0, 1]).sum()
    t = pd.concat([sap.rename("sap_c"), sis.rename("sispo_c")], axis=1).fillna(0).astype("int64")
    t["diferencia_c"] = t["sap_c"] - t["sispo_c"]
    return t.reset_index()


def sispo_a_movimientos(lin: pd.DataFrame, rep: pd.DataFrame, plantilla: pd.DataFrame | None = None) -> pd.DataFrame:
    """Líneas de SISPO ya aplicadas -> movimientos con la forma de los de SAP (UNA sola construcción para la página Flota y para
    el resto de cuentas). El valor queda con el signo de SAP (costo positivo = −Costo Total Con Iva), así la suma coincide con los
    genéricos que reemplazan. Código, sub-rubro y gerencia salen de `sispo_destinos`. Si se pasa `plantilla` (los movimientos de SAP,
    con `llave`), cada fila hereda de un movimiento reemplazado de su mismo destino todas las columnas que SISPO no trae."""
    if lin.empty:
        return pd.DataFrame()
    lin = lin.copy()
    rep = rep.copy()
    lin["_cta"] = lin["cuenta_nombre"].map(lambda x: cuenta_de(x) or SISPO_CUENTA) if "cuenta_nombre" in lin else SISPO_CUENTA
    rep["_cta"] = rep["cuenta"].map(lambda x: cuenta_de(x) or SISPO_CUENTA) if "cuenta" in rep else SISPO_CUENTA
    partes = []
    for (emp, cta), g in lin.groupby(["empresa", "_cta"], sort=False):
        r = rep[(rep["empresa"] == emp) & (rep["_cta"] == cta)]
        d = sispo_destinos(g, r)
        per = g["periodo"].astype(int)
        n = pd.DataFrame({
            "clave": "SISPO-" + per.astype(str) + "-" + g["id_kardex"].astype(str), "empresa": emp,
            "anio": per // 100, "mes": per % 100, "fecha": g["fecha_mov"].astype(str).str[:10],
            "documento": "Kardex " + g["id_kardex"].astype(str), "tercero": g["bodega"], "texto": g["articulo"],
            "cuenta": "SISPO", "cuenta_nombre": "SISPO · " + cta, "cc": g["cc_codigo"], "cc_nombre": g["cc_nombre"],
            "codigo": d["codigo"], "sub_rubro": d["sub_rubro"], "sub_rubro_sap": d["sub_rubro_sap"], "fuera_flota": False,
            "valor": g["valor_sap"], "sispo": True, "codigo_articulo": g["codigo"], "articulo": g["articulo"],
            "cantidad": g["cantidad"].abs(), "placa_sispo": g["placa"], "placa_estado_sispo": g["placa_estado"],
            "concepto_sispo": g["concepto"], "concepto_estado_sispo": g["concepto_estado"], "cuenta_sispo": cta,
            "gerencia": d["gerencia"]})
        if plantilla is not None and len(plantilla) and "llave" in plantilla:
            usados = plantilla[plantilla["llave"].isin(set(r["llave"]))]
            if len(usados):
                for (cod, sub), idx in n.groupby(["codigo", "sub_rubro"]).groups.items():
                    cand = usados[(usados["codigo"].fillna("").astype(str) == cod)]
                    cand = cand.iloc[:1] if len(cand) else usados.iloc[:1]
                    for c in cand.columns:
                        if c not in n.columns and c != "llave":
                            n.loc[idx, c] = cand[c].iloc[0]
        partes.append(n)
    return pd.concat(partes, ignore_index=True)


def sispo_de_cuentas(lin: pd.DataFrame, rep: pd.DataFrame, cuentas) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Solo lo de las `cuentas` indicadas, tanto de las líneas aplicadas como de los movimientos de SAP que reemplazan. Un registro
    guardado antes de existir las cuentas (sin cuenta) es de la flota."""
    cuentas = set(cuentas)
    cl = lin["cuenta_nombre"].map(lambda x: cuenta_de(x) or SISPO_CUENTA) if len(lin) and "cuenta_nombre" in lin else pd.Series(SISPO_CUENTA, index=lin.index)
    cr = rep["cuenta"].map(lambda x: cuenta_de(x) or SISPO_CUENTA) if len(rep) and "cuenta" in rep else pd.Series(SISPO_CUENTA, index=rep.index)
    return lin[cl.isin(cuentas)], rep[cr.isin(cuentas)]


def _sispo_cc_conciliados(lin: pd.DataFrame, rep: pd.DataFrame, presentes: set) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Parte de `lin`/`rep` que se puede superponer por centro de costo: solo los genéricos que están en `presentes` y solo las
    parejas (cuenta, centro de costo) donde lo de SISPO y lo de SAP suman lo mismo al centavo."""
    if lin.empty or rep.empty:
        return lin.iloc[0:0], rep.iloc[0:0]
    rep = rep[rep["llave"].isin(presentes)]
    if rep.empty:
        return lin.iloc[0:0], rep
    cl = lin["cuenta_nombre"].map(lambda x: cuenta_de(x) or SISPO_CUENTA) if "cuenta_nombre" in lin else pd.Series(SISPO_CUENTA, index=lin.index)
    cr = rep["cuenta"].map(lambda x: cuenta_de(x) or SISPO_CUENTA) if "cuenta" in rep else pd.Series(SISPO_CUENTA, index=rep.index)
    kl = pd.Series(list(zip(cl, lin["cc_codigo"].fillna("").astype(str).str.strip().str.upper())), index=lin.index)
    kr = pd.Series(list(zip(cr, rep["cc"].fillna("").astype(str).str.strip().str.upper())), index=rep.index)
    sis = pd.Series([centavos(v) for v in pd.to_numeric(lin["valor_sap"], errors="coerce").fillna(0.0)], index=lin.index).groupby(kl).sum()
    sap = pd.Series([centavos(v) for v in pd.to_numeric(rep["valor"], errors="coerce").fillna(0.0)], index=rep.index).groupby(kr).sum()
    buenas = {k for k in sap.index if k in sis.index and int(sap[k]) == int(sis[k])}
    return lin[kl.isin(buenas)], rep[kr.isin(buenas)]


def sispo_superponer(ejec: pd.DataFrame, lin: pd.DataFrame, rep: pd.DataFrame, empresa: str, cuentas=None,
                     parcial: bool = False) -> pd.DataFrame:
    """Reemplazo reutilizable para cualquier pantalla que lea la ejecución de SAP: quita de `ejec` (con `llave`) los movimientos
    genéricos que SISPO reemplazó y pone en su lugar el detalle de SISPO, dentro del mismo código y sub-rubro del presupuesto. El
    total no cambia (la reclasificación solo se aplicó porque cuadraba al centavo). `cuentas` limita qué cuentas se superponen.
    Por seguridad, una cuenta solo se superpone si TODOS sus genéricos están en `ejec`; si no, se deja como está (así el total
    de esa pantalla nunca queda a medias).
    Con `parcial=True` (para una pantalla que solo trae una parte de los genéricos, como el soporte de un rubro) se reemplaza cada
    pareja (cuenta, centro de costo) cuyo genérico esté en `ejec` y cuyo total de SISPO coincida al centavo con el de SAP; lo
    demás queda como estaba."""
    if ejec.empty or lin.empty or rep.empty:
        return ejec
    lin, rep = sispo_de_cuentas(lin, rep, SISPO_CUENTAS if cuentas is None else cuentas)
    lin, rep = lin[lin["empresa"] == empresa], rep[rep["empresa"] == empresa]
    if lin.empty or rep.empty:
        return ejec
    if parcial:
        lin, rep = _sispo_cc_conciliados(lin, rep, set(ejec["llave"]))
        if lin.empty or rep.empty:
            return ejec
    else:
        cr = rep["cuenta"].map(lambda x: cuenta_de(x) or SISPO_CUENTA) if "cuenta" in rep else pd.Series(SISPO_CUENTA, index=rep.index)
        cl = lin["cuenta_nombre"].map(lambda x: cuenta_de(x) or SISPO_CUENTA) if "cuenta_nombre" in lin else pd.Series(SISPO_CUENTA, index=lin.index)
        presentes = set(ejec["llave"])
        completas = [c for c in cr.unique() if rep.loc[cr == c, "llave"].isin(presentes).all() and (cl == c).any()]
        if not completas:
            return ejec
        lin, rep = lin[cl.isin(completas)], rep[cr.isin(completas)]
    nuevos = sispo_a_movimientos(lin, rep, plantilla=ejec)
    base = ejec[~ejec["llave"].isin(set(rep["llave"]))]
    if len(base) and pd.api.types.is_datetime64_any_dtype(base["fecha"]):
        nuevos["fecha"] = pd.to_datetime(nuevos["fecha"], errors="coerce")
    nuevos["llave"] = llave_movimiento(nuevos.assign(empresa=empresa))
    return pd.concat([base, nuevos], ignore_index=True)


def sispo_distribucion(lineas: pd.DataFrame, por_cuenta: bool = False) -> pd.DataFrame:
    """Valor (signo SISPO) por concepto de mantenimiento, con número de líneas, referencias y placas. Con `por_cuenta`, también
    separado por cuenta (un mismo concepto en dos cuentas son dos filas)."""
    llaves = ["cuenta", "concepto"] if por_cuenta and "cuenta" in lineas else ["concepto"]
    if lineas.empty:
        return pd.DataFrame(columns=[*llaves, "lineas", "referencias", "placas", "valor"])
    g = lineas.groupby(llaves).agg(lineas=("id_kardex", "size"), referencias=("codigo", "nunique"),
                                       placas=("placa", lambda s: s[s != SIN_PLACA].nunique()), centavos=("centavos", "sum"))
    g["valor"] = g["centavos"] / 100
    return g.drop(columns="centavos").reset_index().sort_values("valor")


def sispo_pendientes(lineas: pd.DataFrame) -> pd.DataFrame:
    """Referencias sin concepto, de mayor a menor valor: sirve para completar la base maestra (tiene las mismas columnas)."""
    p = lineas[lineas["concepto"] == PENDIENTE]
    if p.empty:
        return pd.DataFrame(columns=["Codigo Articulo", "Articulo", "CONCEPTO ASIGNADO", "Motivo", "Salidas", "Valor SISPO", "Cuenta"])
    p = p.assign(_k=p["codigo"].where(p["codigo"] != "", "(sin código) " + p["articulo"]))      # sin código: se agrupa por artículo
    cuenta = ("cuenta", lambda x: " / ".join(sorted(set(x)))) if "cuenta" in p else ("_k", lambda x: "")
    g = p.groupby("_k").agg(codigo=("codigo", "first"), Articulo=("articulo", "first"), Motivo=("concepto_estado", "first"),
                            Salidas=("id_kardex", "size"), centavos=("centavos", "sum"), Cuenta=cuenta).reset_index(drop=True)
    g["Valor SISPO"] = g["centavos"] / 100
    g["CONCEPTO ASIGNADO"] = ""
    return g.rename(columns={"codigo": "Codigo Articulo"}).sort_values("Valor SISPO")[
        ["Codigo Articulo", "Articulo", "CONCEPTO ASIGNADO", "Motivo", "Salidas", "Valor SISPO", "Cuenta"]].reset_index(drop=True)


def sispo_huella(lineas: pd.DataFrame) -> str:
    """Huella del CONTENIDO procesado (no del archivo): el mismo SISPO da siempre la misma huella."""
    import hashlib
    filas = sorted((r.id_kardex, r.codigo, r.centavos, r.cc_codigo, r.centro_costo) for r in lineas.itertuples())
    return hashlib.sha256(repr(filas).encode("utf-8")).hexdigest()[:16]


def sispo_huella_maestro(maestro: pd.DataFrame) -> str:
    import hashlib
    return hashlib.sha256(repr(sorted(zip(maestro["codigo"], maestro["concepto"]))).encode("utf-8")).hexdigest()[:16]


def sispo_filas_trazabilidad(lineas: pd.DataFrame, periodo: int, estado: str, generico: str) -> pd.DataFrame:
    """Registro de control (sección «Control y trazabilidad»): una fila por salida de SISPO."""
    return pd.DataFrame({
        "Período": periodo, "Cuenta": lineas["cuenta"] if "cuenta" in lineas else (lineas["cuenta_nombre"] if "cuenta_nombre" in lineas else SISPO_CUENTA),
        "Concepto genérico original": generico, "IdKardex": lineas["id_kardex"], "Placa": lineas["placa"],
        "Referencia": lineas["codigo"], "Artículo": lineas["articulo"], "Cantidad": lineas["cantidad"],
        "Concepto nuevo": lineas["concepto"], "Estado del concepto": lineas["concepto_estado"],
        "Valor original (SISPO, con IVA)": lineas["costo_con_iva"], "Estado de conciliación": estado})
