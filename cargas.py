"""
Lectura de los archivos fuente de Área Limpia.

- leer_presupuestos() / leer_presupuesto(): presupuesto por versión (libro de cierre, valores en miles).
- leer_ejecucion_sap():          extracto de ejecución contable de SAP.
- leer_flujo(): hoja 'Flujo de caja' del libro (conceptos, secciones y proyectado 4+8).
- leer_relacion_combustible():   relaciones de combustible (ALDC / ALSM) por proveedor.
Al cambiar de ERP, estos lectores son los que se adaptan para devolver las mismas columnas.
"""
import re
from datetime import datetime

import pandas as pd

# Clases del PyG (código -> nombre). El rubro se guarda como "3.1 Compras" para que ordene y se lea igual que en SAP.
CLASES = {   # nombres según la tabla de rubros de Financiera (1.0 = ingresos)
    "1.0": "Ingresos",
    "2.0": "Suplidos",
    "3.0": "Personal",
    "3.1": "Compras",
    "3.2": "Subcontratas",
    "3.3": "Servicios Exteriores",
    "3.4": "Alquileres y Arrendamientos",
    "3.5": "Rep. Y Conservación",
    "3.6": "Transportes",
    "3.7": "Seguros y Avales",
    "3.8": "Suplidos",
    "3.9": "Suministros",
    "3.9.8": "Tributos",
    "3.9.9": "Otros Gastos e Ingresos",
    "4.0": "Amortizaciones",
    "4.2": "Otros proyectos",
    "4.1": "Variacion de provisiones",
    "6.0": "Rdo Financiero y Rdo Inmovilizado",
    "7.0": "Impuestos",
}
RUBROS = {k: f"{k} {v}" for k, v in CLASES.items()}
# Subtotales del PyG: nombre -> clases que suma (todo en signo SAP: ingreso negativo, gasto positivo)
EBITDA = [k for k in CLASES if k < "4"]
SUBTOTALES = {
    "EBITDA": EBITDA,
    "Resultado antes de impuestos": EBITDA + ["4.0", "4.1", "4.2", "6.0"],
    "Resultado del ejercicio": list(CLASES),
}

def orden_clase(codigo: str) -> tuple:
    return tuple(int(p) for p in re.findall(r"\d+", str(codigo))) or (99,)


def _codigo(x) -> str:
    """Normaliza el código del PyG (3.1, '3.1', '3.9.8', 1 -> '1.0')."""
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if s.endswith(".0") and s.count(".") > 1:
        s = s[:-2]
    if re.fullmatch(r"\d", s):
        s += ".0"
    return s


def _limpio(x) -> str:
    return " ".join(str(x).split()) if pd.notna(x) else ""


def norm(x) -> str:
    """Nombre normalizado para cruzar: sin tildes, minúsculas, sin signos ni espacios dobles."""
    import unicodedata
    s = unicodedata.normalize("NFKD", str(x)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9 ]", " ", s)).strip()


def rubro_de(cod: str) -> str:
    return RUBROS.get(cod, f"{cod} (código no reconocido)")


def hojas(archivo) -> list[str]:
    return pd.ExcelFile(archivo).sheet_names


# ------------------------------------------------------------------ presupuesto
def leer_presupuesto(archivo, hoja: str | None = None) -> pd.DataFrame:
    """
    Hoja "Gastos de explotacion AAAA" (valores en miles) -> anio, codigo, rubro, sub_rubro, mes, valor (PESOS).
    Toma todas las clases del PyG. Omite las filas que son subtotal de otras de su misma clase
    (ej. 'INGRESOS' = Facturación de servicio + Soterrados + Mixtos + ...), para no contar doble.
    """
    if hoja is None:
        hoja = next(h for h in hojas(archivo)
                    if h.lower().replace("ó", "o").startswith("gastos de explotacion"))
    df = pd.read_excel(archivo, sheet_name=hoja, header=None)

    fila_fechas, cols_mes = None, []
    for i in range(min(15, len(df))):
        fechas = [j for j, v in enumerate(df.iloc[i]) if isinstance(v, (datetime, pd.Timestamp))]
        if len(fechas) >= 12:
            fila_fechas, cols_mes = i, fechas[:12]
            break
    if fila_fechas is None:
        raise ValueError(f"No encontré la fila de meses en la hoja '{hoja}'.")
    anio = pd.Timestamp(df.iat[fila_fechas, cols_mes[0]]).year
    col_cod = next(j for j in range(cols_mes[0]) if df.iloc[:, j].map(_codigo).isin(RUBROS).sum() > 5)
    col_nom = cols_mes[0] - 1

    lineas = []
    for _, r in df.iloc[fila_fechas + 1:].iterrows():
        cod = _codigo(r[col_cod])
        if cod not in RUBROS:
            continue
        vals = [pd.to_numeric(r[c], errors="coerce") for c in cols_mes]
        lineas.append((cod, _limpio(r[col_nom]), [0.0 if pd.isna(v) else float(v) * 1000 for v in vals]))

    # quitar subtotales: una línea igual a la suma de las demás de su clase
    import numpy as np
    finales = []
    for i, (cod, nom, vals) in enumerate(lineas):
        otras = [v for j, (c2, _, v) in enumerate(lineas) if c2 == cod and j != i]
        if otras and np.abs(vals).sum() > 0 and np.allclose(np.sum(otras, axis=0), vals, atol=1):
            continue
        finales.append((cod, nom, vals))

    filas = [(anio, cod, RUBROS[cod], nom, m, v) for cod, nom, vals in finales for m, v in enumerate(vals, 1)]
    out = pd.DataFrame(filas, columns=["anio", "codigo", "rubro", "sub_rubro", "mes", "valor"])
    # nombres repetidos en clases distintas (ej. 'Aguas de bogota' en 1.0 y 3.2) se agrupan por clase
    return out.groupby(["anio", "codigo", "rubro", "sub_rubro", "mes"], as_index=False)["valor"].sum()


# ------------------------------------------------------------------ PyG (validación y clases sin línea de presupuesto)

# ------------------------------------------------------------------ ejecución (BASE SAP)
def _col(df, *opciones, obligatoria=True):
    for o in opciones:
        for c in df.columns:
            if str(c).strip().lower().startswith(o.lower()):
                return c
    if obligatoria:
        raise ValueError(f"Falta la columna '{opciones[0]}' en la BASE.")
    return None



# ------------------------------------------------------------------ dueños por sub-rubro

# ------------------------------------------------------------------ combustible
EMPRESAS = {"AREA LIMPIA DISTRITO CAPITAL": "ALDC", "AREA LIMPIA SERVICIOS MEDIOAMBIENTALES": "ALSM"}


def _empresa_corta(nombre: str) -> str:
    n = _limpio(nombre).upper()
    for largo, corto in EMPRESAS.items():
        if n.startswith(largo):
            return corto
    return n[:20]


COLS_COMBUSTIBLE = ["empresa", "fuente", "fecha", "anio", "mes", "placa", "tipo", "galones", "valor", "contrato",
                    "clave", "documento", "estacion", "hora", "km"]

# Cada proveedor llega en su propia hoja con su propio formato. Se reconocen por sus encabezados, no por el nombre
# de la hoja, así que da igual el orden o si cambia el nombre. (alias de encabezado, sin tildes ni mayúsculas)
FORMATOS_COMBUSTIBLE = {
    "Terpel": {"firma": ["no. venta", "producto"],
               "empresa": ["compania", "cliente"], "fecha": ["fecha"], "hora": ["hora"], "placa": ["placa"],
               "tipo": ["producto"], "galones": ["cantidad"], "valor": ["total venta"], "documento": ["no. venta"],
               "estacion": ["estacion"], "km": ["kilometraje"]},
    "Autogas": {"firma": ["no. recibo", "volumen"],
                "empresa": ["compania", "nombre cliente", "cliente"], "fecha": ["fecha"], "hora": ["hora"],
                "placa": ["placa"], "tipo": ["articulo"], "galones": ["volumen"], "valor": ["valor venta"],
                "documento": ["no. recibo"], "estacion": ["estacion de servicio"], "km": ["km"]},
    "Petrobras": {"firma": ["consumo", "precio galon"],
                  "empresa": ["compania", "cliente"], "fecha": ["fecha"], "placa": ["placa"],
                  "tipo": ["tipo combustible"], "galones": ["consumo"], "valor": ["total consumo"],
                  "precio": ["precio galon"], "documento": ["movil"]},
}


def _n(x) -> str:
    return " ".join(norm(x).lower().split())


def _col_alias(df: pd.DataFrame, alias: list[str]):
    """Primera columna cuyo encabezado (sin tildes/mayúsculas/espacios extra) es igual o empieza por el alias."""
    cols = {c: _n(c).replace(".", ". ").replace("  ", " ").strip() for c in df.columns}
    for a in alias:
        a2 = _n(a).replace(".", ". ").replace("  ", " ").strip()
        for c, cn in cols.items():
            if cn == a2:
                return c
        for c, cn in cols.items():
            if cn.startswith(a2):
                return c
    return None


def _tipo_combustible(x) -> str:
    t = _limpio(x).upper()
    t = re.sub(r"\s*\d+$", "", t)                       # 'UREA 1' -> 'UREA'
    return {"ACPM": "DIESEL", "GASOLINA CORRIENTE": "GASOLINA"}.get(t, t)


def _fechas(s: pd.Series) -> pd.Series:
    """Fechas que pueden venir como fecha de Excel o como texto dd/mm/aaaa."""
    f = pd.to_datetime(s.where(s.map(lambda v: not isinstance(v, str))), errors="coerce")
    txt = s.where(s.map(lambda v: isinstance(v, str)))
    if txt.notna().any():
        f = f.fillna(pd.to_datetime(txt, errors="coerce", dayfirst=True))
    return f


def _empresa_de_archivo(nombre: str) -> str:
    n = (nombre or "").upper()
    return "ALSM" if "ALSM" in n else "ALDC" if "ALDC" in n else ""


def leer_relacion_combustible(archivo, nombre: str = "") -> tuple[pd.DataFrame, dict]:
    """
    Lee una relación de combustible tal como la entrega el área (una hoja por proveedor: Petrobras, Terpel,
    Autogas…) sin que haya que organizar nada. Cada tanqueo recibe una 'clave' única (proveedor + número de venta o
    recibo; Petrobras: fecha + placa + galones + móvil) para no duplicar al volver a cargar.
    De la hoja 'Datos' (si tiene valores) se toman: el mes asignado (cuando se reasignó a mano) y las placas del
    contrato Aguas de Bogotá (columna AGUAS).
    Devuelve (tanqueos, info) — info: hojas reconocidas, placas Aguas, avisos.
    """
    libro = pd.read_excel(archivo, sheet_name=None, dtype=object)
    info = {"hojas": {}, "aguas": set(), "avisos": [], "ignoradas": []}
    partes = []
    for hoja, df in libro.items():
        df = df.dropna(how="all")
        df.columns = [str(c).strip() for c in df.columns]
        formato = next((f for f, d in FORMATOS_COMBUSTIBLE.items()
                        if all(_col_alias(df, [a]) is not None for a in d["firma"])), None)
        if formato is None:
            if _col_alias(df, ["aguas"]) is not None and _col_alias(df, ["placa"]) is not None:
                ag = df[_col_alias(df, ["aguas"])].map(_limpio).str.upper()
                info["aguas"] |= set(ag[(ag != "") & (ag != "NO APLICA")])
                info["_datos"] = df
            elif hoja.strip().lower() not in ("resumen", "datos"):
                info["ignoradas"].append(hoja)
            continue
        d = FORMATOS_COMBUSTIBLE[formato]
        c = {k: _col_alias(df, v) for k, v in d.items() if k != "firma"}
        placa = df[c["placa"]].map(_limpio).str.upper().str.replace(" ", "")
        fecha = _fechas(df[c["fecha"]])
        ok = (placa != "") & fecha.notna()
        galones = pd.to_numeric(df[c["galones"]], errors="coerce")
        valor = pd.to_numeric(df[c["valor"]], errors="coerce") if c.get("valor") else pd.Series(float("nan"), df.index)
        if c.get("precio"):                                  # Petrobras sin 'total consumo': galones x precio
            valor = valor.fillna(galones * pd.to_numeric(df[c["precio"]], errors="coerce"))
        emp_arch = _empresa_de_archivo(nombre)
        empresa = df[c["empresa"]].map(_empresa_corta) if c.get("empresa") else pd.Series(emp_arch, df.index)
        empresa = empresa.where(empresa.isin(["ALDC", "ALSM"]) | (emp_arch == ""), emp_arch)
        texto = lambda k: df[c[k]].map(_limpio) if c.get(k) else pd.Series("", df.index)
        t = pd.DataFrame({
            "empresa": empresa, "fuente": formato, "fecha_dt": fecha, "placa": placa,
            "tipo": df[c["tipo"]].map(_tipo_combustible) if c.get("tipo") else "DIESEL",
            "galones": galones.fillna(0.0), "valor": valor.fillna(0.0),
            "documento": texto("documento").str.replace(r"\.0$", "", regex=True),
            "estacion": texto("estacion"), "hora": texto("hora").str[:5],
            "km": texto("km").str.replace(r"\.0$", "", regex=True), "_fila": range(len(df)),
        })
        t["_ok"] = ok.values
        partes.append(t)
        info["hojas"][hoja] = {"proveedor": formato, "tanqueos": int(ok.sum()), "filas": len(df)}
    if not partes:
        raise ValueError("No encontré hojas de proveedores (Petrobras, Terpel, Autogas) con su formato habitual.")

    # mes asignado en la hoja Datos (mismo orden que las hojas de proveedores) si trae valores calculados
    datos = info.pop("_datos", None)
    mes_datos = {}
    if datos is not None and _col_alias(datos, ["fuente"]) is not None and _col_alias(datos, ["mes"]) is not None:
        dd = datos[[_col_alias(datos, ["fuente"]), _col_alias(datos, ["mes"])]].copy()
        dd.columns = ["fuente", "mes"]
        dd["mes"] = pd.to_datetime(dd["mes"], errors="coerce")
        for fuente, g in dd.groupby(dd["fuente"].map(_limpio), sort=False):
            mes_datos[fuente] = g["mes"].tolist()

    salida = []
    for t in partes:
        fuente = t["fuente"].iloc[0]
        meses = mes_datos.get(fuente)
        if meses and len(meses) == len(t) and pd.notna(pd.Series(meses)).mean() > 0.9:
            t["mes_dt"] = pd.to_datetime(pd.Series(meses, index=t.index)).fillna(t["fecha_dt"])
            cambiados = int((t["mes_dt"].dt.to_period("M") != t["fecha_dt"].dt.to_period("M")).sum())
            if cambiados:
                info["avisos"].append(f"{fuente}: {cambiados} tanqueos asignados a otro mes en la hoja Datos "
                                      "(se respeta ese mes).")
        else:
            t["mes_dt"] = t["fecha_dt"]
        salida.append(t[t["_ok"]])
    out = pd.concat(salida, ignore_index=True)
    out["fecha"] = out["fecha_dt"].dt.strftime("%Y-%m-%d")
    out["anio"] = out["mes_dt"].dt.year.astype(int)
    out["mes"] = out["mes_dt"].dt.month.astype(int)
    out["contrato"] = ""          # se define al cargar con la lista de placas de Aguas de Bogotá
    # clave única por tanqueo; el contador desempata registros idénticos del mismo archivo
    base = out["empresa"] + "|" + out["fuente"] + "|"
    num = out["documento"].where(out["fuente"] != "Petrobras", "")
    base += num.where(num != "", out["fecha"] + "|" + out["placa"] + "|" + out["galones"].round(3).astype(str)
                      + "|" + out["documento"])
    base = base.where(num == "", base + "|" + out["fecha"] + "|" + out["placa"])
    out["clave"] = base + "#" + out.groupby(base).cumcount().astype(str)
    out.loc[out["fuente"] == "Petrobras", "documento"] = ""      # en Petrobras era el número de móvil (solo para la clave)
    return out[COLS_COMBUSTIBLE].reset_index(drop=True), info



# ------------------------------------------------------------------ flujo de caja
_TOTALES = {"total ingresos", "total egresos", "flujo de caja económico", "total financiamiento",
            "flujo de caja operación", "flujo de caja financiero", "saldo final otros componentes",
            "saldo final caja", "saldo inicial", "saldo inicial otros componentes"}
_SECCIONES = {"ingresos": "Ingresos", "egresos": "Egresos", "financiamiento": "Financiamiento"}


def leer_flujo(archivo, hoja: str = "Flujo de caja") -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Devuelve (lineas, proyectado):
      lineas: anio, orden, seccion, padre, concepto, tipo ('linea'|'grupo'|'total'), mes, valor, fila (de Excel)
      proyectado: anio, concepto, proyectado (columna FLUJO 4+8), observacion
    """
    df = pd.read_excel(archivo, sheet_name=hoja, header=None)
    fila_mes = next(i for i in range(15) if str(df.iat[i, 2]).strip().upper() == "ENE")
    anio = None
    m = re.search(r"(20\d\d)", str(getattr(archivo, "name", archivo)))
    anio = int(m.group(1)) if m else datetime.now().year
    cab = [str(x).strip().upper() for x in df.iloc[fila_mes]]
    col_48 = next((j for j, c in enumerate(cab) if "4+8" in c or "PROYECT" in c), None)
    col_obs = col_48 + 2 if col_48 is not None else None

    filas, seccion = [], "Saldos"
    for i in range(fila_mes + 1, len(df)):
        nombre = df.iat[i, 1]
        if not isinstance(nombre, str) or not nombre.strip():
            continue
        nombre = nombre.strip()
        if nombre.upper().startswith("VALIDACION"):
            break
        clave = nombre.lower()
        vals = pd.to_numeric(df.iloc[i, 2:14], errors="coerce")
        if clave in _SECCIONES and vals.isna().all():
            seccion = _SECCIONES[clave]
            continue
        if clave == "contrato comercial socios":
            seccion = "Socios"
        if clave == "saldo inicial otros componentes":
            seccion = "Otros componentes"
        if clave.startswith("saldo final caja"):
            seccion = "Saldos"
        proy = pd.to_numeric(df.iat[i, col_48], errors="coerce") if col_48 is not None else float("nan")
        obs = df.iat[i, col_obs] if col_obs is not None and col_obs < df.shape[1] else None
        filas.append(dict(fila=i, concepto=nombre, seccion=seccion,
                          tipo="total" if clave in _TOTALES else "linea",
                          vals=vals.fillna(0.0).to_numpy(), proy=proy,
                          obs=obs if isinstance(obs, str) else ""))

    # grupos: una línea cuyo valor es la suma exacta de las siguientes de su sección (ej. Gestion Humana)
    import numpy as np
    for k, f in enumerate(filas):
        f.setdefault("padre", "")
        if f["tipo"] != "linea" or not np.abs(f["vals"]).sum():
            continue
        acum = np.zeros(12)
        for j in range(k + 1, len(filas)):
            g = filas[j]
            if g["seccion"] != f["seccion"] or g["tipo"] != "linea" or g.get("padre"):
                break
            acum = acum + g["vals"]
            if j - k >= 2 and np.allclose(acum, f["vals"], atol=5):
                f["tipo"] = "grupo"
                for h in filas[k + 1:j + 1]:
                    h["padre"] = f["concepto"]
                break

    lineas, proy = [], []
    for orden, f in enumerate(filas):
        for mes in range(1, 13):
            lineas.append((anio, orden, f["seccion"], f["padre"], f["concepto"], f["tipo"], mes,
                           float(f["vals"][mes - 1]), f["fila"] + 1))
        if pd.notna(f["proy"]):
            proy.append((anio, f["concepto"], float(f["proy"]), f["obs"]))
    return (pd.DataFrame(lineas, columns=["anio", "orden", "seccion", "padre", "concepto", "tipo", "mes", "valor",
                                          "fila"]),
            pd.DataFrame(proy, columns=["anio", "concepto", "proyectado", "observacion"]))


# ------------------------------------------------------------------ presupuestos por versión (Presupuestos.xlsx)
def linea_norm(x) -> str:
    """Nombre de línea de negocio unificado (MAYÚSCULAS, sin tildes; 'ADMISNITRATIVO' -> 'ADMINISTRATIVO')."""
    t = re.sub(r"\s*\+\s*", " + ", _limpio(x).upper()) if pd.notna(x) else ""
    t = " ".join(t.replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U").split())
    if t.startswith("ADMI") or t.startswith("ADMS"):
        return "ADMINISTRATIVO"
    return t


def leer_distribucion(archivo) -> pd.DataFrame:
    """Hoja con el % de reparto de lo administrativo por línea de negocio y año (ej. 'Hoja2'):
    primera columna = línea, encabezado = años. Devuelve anio, linea, pct."""
    filas = []
    for hoja in hojas(archivo):
        df = pd.read_excel(archivo, sheet_name=hoja, header=None)
        if df.empty or df.shape[1] < 2:
            continue
        enc = df.iloc[0]
        anios = {j: int(v) for j, v in enc.items() if isinstance(v, (int, float)) and pd.notna(v) and 2000 < v < 2100}
        if not anios or len(df) > 60:
            continue
        for _, r in df.iloc[1:].iterrows():
            lin = linea_norm(r.iloc[0])
            if not lin:
                continue
            for j, a in anios.items():
                v = pd.to_numeric(r.iloc[j], errors="coerce")
                if pd.notna(v) and 0 <= v <= 1:
                    filas.append((a, lin, float(v)))
    return pd.DataFrame(filas, columns=["anio", "linea", "pct"])


def leer_ceco_lineas(archivo) -> pd.DataFrame:
    """Hoja de centros de coste con su gerencia y proyecto / línea de negocio (ej. 'CECO ALSM': CE.COSTE,
    SOC, TEXTO BREVE = gerencia, PROYECTO = línea). Devuelve cc, gerencia, linea (vacío si no hay esa hoja)."""
    for hoja in hojas(archivo):
        df = pd.read_excel(archivo, sheet_name=hoja, header=None, dtype=str, nrows=2000)
        for i in range(min(5, len(df))):
            enc = [norm(x) for x in df.iloc[i]]
            c_cc = next((j for j, t in enumerate(enc) if t.replace(" ", "") in ("cecoste", "centrodecoste", "ceco",
                                                                             "centrodecosto")), None)
            c_lin = next((j for j, t in enumerate(enc) if t in ("proyecto", "linea de negocio", "linea")), None)
            if c_cc is None or c_lin is None:
                continue
            c_ger = next((j for j, t in enumerate(enc) if t in ("texto breve", "gerencia")), None)
            t = df.iloc[i + 1:]
            out = pd.DataFrame({"cc": t.iloc[:, c_cc].map(_limpio),
                                "gerencia": t.iloc[:, c_ger].map(_limpio) if c_ger is not None else "",
                                "linea": t.iloc[:, c_lin].map(linea_norm)})
            out = out[(out["cc"] != "") & (out["linea"] != "")]
            return out.drop_duplicates("cc", keep="last").reset_index(drop=True)
    return pd.DataFrame(columns=["cc", "gerencia", "linea"])


def nombre_version(hoja: str) -> str:
    v = re.sub(r"^(PPTO|PRESUPUESTO)\s*", "", hoja.strip(), flags=re.I).strip()
    return v.capitalize() if v.isalpha() else v


def leer_presupuestos(archivo) -> pd.DataFrame:
    """
    Una hoja por versión (ej. 'PPTO INICIAL', 'PPTO 4+8', 'PPTO 8+4') con columnas:
    Codigo rubro | Subrubro | [Línea de negocio] | 12 meses (valores en PESOS).
    Las columnas se ubican por su encabezado; la línea de negocio es opcional (presupuesto de ALSM).
    Devuelve: anio, version, codigo, rubro, sub_rubro, linea, mes, valor.
    """
    out = []
    for hoja in hojas(archivo):
        df = pd.read_excel(archivo, sheet_name=hoja, header=None)
        fila = next((i for i in range(min(10, len(df))) if sum(isinstance(v, (datetime, pd.Timestamp)) for v in df.iloc[i]) >= 12),
                    None)
        if fila is None:
            continue
        cols_mes = [j for j, v in enumerate(df.iloc[fila]) if isinstance(v, (datetime, pd.Timestamp))]
        fechas = {j: pd.Timestamp(df.iat[fila, j]) for j in cols_mes}
        enc = {j: norm(df.iat[fila, j]) for j in range(cols_mes[0]) if pd.notna(df.iat[fila, j])}
        def buscar(*claves):
            return next((j for j, t in enc.items() if any(k in t for k in claves)), None)
        c_lin = buscar("linea", "negocio")
        c_cod = buscar("codigo")
        c_sub = buscar("subrubro", "sub rubro", "sub-rubro")
        if c_cod is None or c_sub is None:          # sin encabezados reconocibles: las dos columnas antes de los meses
            previas = [j for j in range(cols_mes[0]) if j != c_lin]
            c_cod, c_sub = previas[-2], previas[-1]
        # la versión sale del nombre de la hoja ('PPTO 8+4'); si la hoja no lo dice queda vacía y se elige al cargar
        version = nombre_version(hoja) if re.match(r"^\s*(PPTO|PRESUPUESTO)\b", hoja, flags=re.I) else ""
        for _, r in df.iloc[fila + 1:].iterrows():
            cod, sub = _codigo(r[c_cod]), _limpio(r[c_sub])
            if not cod or not sub:
                continue
            lin = linea_norm(r[c_lin]) if c_lin is not None else ""
            for c in cols_mes:
                v = pd.to_numeric(r[c], errors="coerce")
                out.append((fechas[c].year, version, cod, rubro_de(cod), sub, lin, fechas[c].month,
                            0.0 if pd.isna(v) else float(v)))
    cols = ["anio", "version", "codigo", "rubro", "sub_rubro", "linea", "mes"]
    df = pd.DataFrame(out, columns=cols + ["valor"])
    return df.groupby(cols, as_index=False)["valor"].sum()


# ------------------------------------------------------------------ base SAP de ejecución (formato nuevo)
def _encabezado(crudo: pd.DataFrame, marca: str = "Valor/mon") -> tuple[int, list[str]]:
    fila = next(i for i in range(min(15, len(crudo)))
                if any(str(v).strip().startswith(marca) for v in crudo.iloc[i]))
    vistos, cols = {}, []
    for c in crudo.iloc[fila]:
        c = str(c).strip()
        vistos[c] = vistos.get(c, -1) + 1
        cols.append(c if vistos[c] == 0 else f"{c}.{vistos[c]}")
    return fila, cols


# nombres posibles de cada columna en las distintas exportaciones de SAP (se busca primero el nombre exacto)
ALIAS_SAP = {
    "cuenta": ["Referencia", "Clase de coste", "Clase coste"],
    "cuenta_nombre": ["Denom.clase de coste", "Denominación clase de coste"],
    "documento": ["Nº docum.refer.", "N° docum.refer.", "No docum.refer.", "Nº documento"],
    "valor": ["Valor/mon.inf.", "Valor/mon. inf.", "Valor"],
    "contrapartida": ["Cta.contrapartida", "Cuenta contrapartida"],
    "tercero": ["Denom.cuenta contrapartida", "Denominacion cuenta contrapartida",
                "Denominación cuenta contrapartida"],
    "cc_sap": ["Centro de coste", "Centro coste"],
    "texto": ["TEXTO", "Denominación", "Texto"],
    "cc_nombre": ["Denominación del objeto", "Denominacion del objeto"],
    "anulado": ["Anulado"],
    "periodo": ["De período", "De periodo", "Período", "Periodo"],
    "ejercicio": ["Ejercicio"],
    "fecha": ["Fe.contabilización", "Fe.contabilizacion", "Fecha contabilización"],
}


def _buscar(cols: list[str], alias: list[str], hasta: int | None = None) -> str | None:
    """Columna por nombre: primero coincidencia exacta (sin importar mayúsculas), luego 'empieza por'."""
    zona = cols[:hasta] if hasta else cols
    bajos = [c.lower().strip() for c in zona]
    for a in alias:
        if a.lower() in bajos:
            return zona[bajos.index(a.lower())]
    for a in alias:
        for c, b in zip(zona, bajos):
            if b.startswith(a.lower()):
                return c
    return None


def parsear_ejecucion(df: pd.DataFrame, cols: list[str]) -> tuple[pd.DataFrame, list[str]]:
    """Convierte una tabla de SAP (con sus encabezados) en movimientos para la app."""
    import json
    df = df.copy()
    df.columns = cols
    i_mod = next((i for i, c in enumerate(cols) if c.lower().startswith("modifica cod")), None)
    i_val = next((i for i, c in enumerate(cols) if c.upper() == "CODIGO"), None)
    if i_val is None:   # sin encabezado CODIGO: la clasificación validada empieza justo antes de 'BASE GASTOS...'
        j = next((i for i, c in enumerate(cols) if c.upper().startswith("BASE GASTOS")), None)
        i_val = j - 1 if j else None
    fin_sap = min(i for i in [i_mod, i_val, len(cols)] if i is not None)
    sap_cols = cols[:fin_sap]

    def col(clave):
        c = _buscar(cols, ALIAS_SAP[clave], fin_sap)
        return df[c].map(_limpio) if c else pd.Series("", index=df.index)

    c_valor = _buscar(cols, ALIAS_SAP["valor"], fin_sap)
    if c_valor is None or _buscar(cols, ALIAS_SAP["cuenta"], fin_sap) is None:
        raise ValueError("No encuentro las columnas de cuenta de coste ('Referencia' o 'Clase de coste') y valor.")
    df = df[df[c_valor].notna() & (df[c_valor].astype(str).str.strip() != "")]
    fecha = pd.to_datetime(df[_buscar(cols, ALIAS_SAP["fecha"], fin_sap)], errors="coerce")
    periodo = pd.to_numeric(col("periodo"), errors="coerce").fillna(fecha.dt.month)
    ejercicio = pd.to_numeric(col("ejercicio"), errors="coerce").fillna(fecha.dt.year)
    out = pd.DataFrame({
        "anio": ejercicio.astype(int), "mes": periodo.astype(int),
        "fecha": fecha.dt.strftime("%Y-%m-%d"),
        "cuenta": col("cuenta").str.replace(r"\.0$", "", regex=True), "cuenta_nombre": col("cuenta_nombre"),
        "documento": col("documento").str.replace(r"\.0$", "", regex=True),
        "valor": pd.to_numeric(df[c_valor], errors="coerce").fillna(0.0),
        "contrapartida": col("contrapartida").str.replace(r"\.0$", "", regex=True), "tercero": col("tercero"),
        "cc_sap": col("cc_sap"), "texto": col("texto"), "cc_nombre": col("cc_nombre"), "anulado": col("anulado"),
        "mod_cod": df[cols[i_mod]].map(_limpio).map(lambda x: _codigo(x) if x else "") if i_mod is not None else "",
        "mod_sub": df[cols[i_mod + 1]].map(_limpio) if i_mod is not None and i_mod + 1 < len(cols) else "",
        "mod_cc": df[cols[i_mod + 2]].map(_limpio) if i_mod is not None and i_mod + 2 < len(cols) else "",
        "val_cod": df[cols[i_val]].map(_limpio).map(lambda x: _codigo(x) if x else "") if i_val is not None else "",
        "val_sub": df[cols[i_val + 1]].map(_limpio) if i_val is not None and i_val + 1 < len(cols) else "",
        "val_cc": df[cols[i_val + 2]].map(_limpio) if i_val is not None and i_val + 2 < len(cols) else "",
    })
    out["tercero"] = out["tercero"].where(out["tercero"] != "", out["cuenta_nombre"])
    base = (out["anio"].astype(str) + "|" + out["mes"].astype(str) + "|" + out["documento"] + "|" + out["cuenta"]
            + "|" + out["cc_sap"] + "|" + out["valor"].round(2).astype(str) + "|" + out["texto"] + "|" + out["fecha"])
    out["clave"] = base + "#" + base.groupby(base).cumcount().astype(str)
    out["raw"] = df[sap_cols].astype(object).where(df[sap_cols].notna(), "").apply(
        lambda r: json.dumps([v.strftime("%Y-%m-%d") if isinstance(v, (datetime, pd.Timestamp)) else str(v)
                              for v in r], ensure_ascii=False), axis=1)
    return out.reset_index(drop=True), sap_cols


def leer_ejecucion_sap(archivo, hoja: str | None = None) -> tuple[pd.DataFrame, list[str]]:
    """
    Base de ejecución de SAP: sirve la exportación directa de SAP y tu base con ajustes.
    Columnas opcionales: ajustes ('Modifica COD', 'Modifica Base Gasto', 'CC') y clasificación validada
    ('CODIGO', 'BASE GASTOS EXPLOTACIO', 'CC'). Las columnas se ubican por nombre, no por letra.
    """
    if hoja is None:
        hoja = hojas(archivo)[0]
    crudo = pd.read_excel(archivo, sheet_name=hoja, header=None, dtype=object)
    fila, cols = _encabezado(crudo)
    return parsear_ejecucion(crudo.iloc[fila + 1:], cols)
