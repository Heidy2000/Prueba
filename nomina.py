"""
Nómina: lectura de archivos, homologación de cargos y cruce presupuesto vs ejecución (no necesitas modificar
este archivo).

- Presupuesto detallado por cargo: una hoja por año ('3.0 Personal 2026', '3.0 Personal 2027'…) con, por cargo,
  las personas de cada mes y el costo de cada mes. Las filas sin salario son títulos de grupo ('D - Barrido Manual').
- Ejecución: base de nómina de SAP (una fila por empleado y concepto) con centro de coste y CARGO.
- Relación de centros de coste con el área (gerencia).
Los nombres de cargo no son iguales en los dos archivos: cada cargo se lleva a un 'cargo común'. La app propone
la homologación (quita '(A)', turnos, tildes; todos los conductores van a 'Conductor') y Financiera la ajusta.
"""
import difflib
import re

import pandas as pd

from cargas import norm as _norm_min


def norm(x) -> str:
    """Texto sin tildes ni signos, en MAYÚSCULAS."""
    return _norm_min(x).upper()

MESES_TXT = {"ENE": 1, "FEB": 2, "MAR": 3, "ABR": 4, "MAY": 5, "JUN": 6, "JUL": 7, "AGO": 8, "SEP": 9, "OCT": 10,
             "NOV": 11, "DIC": 12}
SIN_PPTO = "(sin presupuesto)"
VACIAS = {"DE", "DEL", "LA", "EL", "LOS", "LAS", "Y", "EN", "PARA"}
TURNOS = {"GUARDIA", "NOCHE", "TARDE", "DIA", "TURNO", "MANANA"}


# ------------------------------------------------------------------ homologación
ABREVIATURAS = {"MTTO": "MANTENIMIENTO", "MTO": "MANTENIMIENTO", "ADTVO": "ADMINISTRATIVO",
                "ADTVA": "ADMINISTRATIVO", "ADMINISTRATIVA": "ADMINISTRATIVO", "ADMON": "ADMINISTRACION",
                "TEC": "TECNICO", "OPERATIVA": "OPERATIVO", "AUX": "AUXILIAR", "PROF": "PROFESIONAL"}
RE_TURNO = re.compile(r"\s*(-\s*)?\(?\s*(GUARDIA|NOCHE|TARDE|DIA|D[IÍ]A|MA[NÑ]ANA)\s*\)?\s*$", re.I)


def clave_cargo(cargo) -> str:
    """'OPERARIO (A) DE BARRIDO' y 'Operario de barrido - Guardia' -> 'OPERARIO BARRIDO'."""
    t = norm(re.sub(r"\((a|o|as|os)\)", " ", str(cargo), flags=re.I))      # marcador de género (a)
    palabras = [ABREVIATURAS.get(p, p) for p in t.split()]
    palabras = [p for p in palabras if p not in VACIAS and p not in TURNOS]
    palabras = [p[:-1] if p.endswith("S") and len(p) > 4 and p not in ("SERVICIOS", "GENERALES", "SOCIALES")
                else p for p in palabras]                                   # plural -> singular
    k = " ".join(palabras)
    return "CONDUCTOR" if k.startswith("CONDUCTOR") else k


def nombre_limpio(cargo: str) -> str:
    """Nombre visible del cargo común: sin el turno ('Operario de barrido - Guardia' -> 'Operario de barrido')."""
    c = " ".join(str(cargo).split())
    for _ in range(2):
        c = RE_TURNO.sub("", c).strip(" -")
    return c[:1].upper() + c[1:]


def comunes_ppto(ppto: pd.DataFrame) -> dict:
    """clave -> nombre visible del cargo común (el nombre más corto del presupuesto con esa clave, sin turno)."""
    out = {}
    for c in ppto["cargo"].dropna().unique():
        k = clave_cargo(c)
        if not k:
            continue
        nombre = "Conductor" if k == "CONDUCTOR" else nombre_limpio(c)
        if k not in out or len(nombre) < len(out[k]):
            out[k] = nombre
    return out


def sugerir(cargo: str, comunes: dict) -> str:
    """Cargo común sugerido ('' si no hay uno suficientemente parecido)."""
    k = clave_cargo(cargo)
    if k in comunes:
        return comunes[k]
    parecido = difflib.get_close_matches(k, list(comunes), n=1, cutoff=0.9)
    return comunes[parecido[0]] if parecido else ""


def homologaciones(leer) -> dict:
    """(fuente, cargo normalizado, llave) -> cargo común guardado. llave = centro de coste (SAP), 'AREA:<área>'
    (presupuesto) o '' (asignación general del cargo, de versiones anteriores)."""
    g = leer("SELECT * FROM nomina_homol")
    if "cc" not in g:
        g["cc"] = ""
    return {(f, norm(c), (k or "").upper()): cc for f, c, k, cc in zip(g["fuente"], g["cargo"], g["cc"].fillna(""),
                                                                     g["cargo_comun"]) if cc}


def buscar_manual(manual: dict, fuente: str, cargo: str, llave: str):
    n = norm(cargo)
    return manual.get((fuente, n, (llave or "").upper())) or manual.get((fuente, n, ""))


def homologar(leer, ppto: pd.DataFrame, ejec: pd.DataFrame) -> tuple[pd.Series, pd.Series]:
    """Cargo común de cada fila del presupuesto y de la ejecución. Manda lo guardado por Financiera en
    Administración para ese cargo en ese centro de coste (o área, en el presupuesto); si no, la sugerencia."""
    comunes = comunes_ppto(ppto)
    manual = homologaciones(leer)

    def m(fuente, df, col_llave, prefijo, defecto):
        pares = df[["cargo", col_llave]].fillna("").drop_duplicates()
        mapa = {(c, k): buscar_manual(manual, fuente, c, prefijo + k) or defecto(c)
                for c, k in zip(pares["cargo"], pares[col_llave])}
        return pd.Series([mapa[(c, k)] for c, k in zip(df["cargo"].fillna(""), df[col_llave].fillna(""))],
                         index=df.index, dtype=object)
    cp = m("ppto", ppto, "area", "AREA:", lambda c: comunes.get(clave_cargo(c), nombre_limpio(c))) \
        if not ppto.empty else pd.Series(dtype=object)
    ce = m("ejec", ejec, "cc", "", lambda c: sugerir(c, comunes) or SIN_PPTO) \
        if not ejec.empty else pd.Series(dtype=object)
    return cp, ce


# ------------------------------------------------------------------ lectura de archivos
def _mes_de(texto) -> int | None:
    t = norm(texto)[:3]
    return MESES_TXT.get(t)


# conceptos comparables entre el presupuesto (que se calcula con salario, horas extra, auxilios y porcentajes) y la
# nómina de SAP (que viene por concepto). Columna del presupuesto -> nombre del concepto.
CONCEPTOS = [
    ("c_sueldo", "Sueldos y salarios"),
    ("c_horas", "Horas extra y recargos"),
    ("c_transporte", "Auxilio de transporte"),
    ("c_bono_cump", "Bonificación de cumplimiento"),
    ("c_extralegal", "Auxilios y bonificaciones extralegales"),
    ("c_aportes", "Aportes (salud, pensión, riesgos, caja, parafiscales)"),
    ("c_prestaciones", "Prestaciones (cesantías, intereses, primas)"),
    ("c_vacaciones", "Vacaciones"),
    ("c_otros", "Otros conceptos"),
]
COLS_CONCEPTO = [c for c, _ in CONCEPTOS]
VERSION_LECTOR = 2   # súbelo cuando cambie la lectura de los archivos de nómina

# texto del concepto de SAP -> concepto comparable (se evalúa en orden; el primero que coincide gana)
_REGLAS_SAP = [
    ("Bonificación de cumplimiento", ["PROVISION BONO"]),
    ("Vacaciones", ["PROVISION VACACION"]),
    ("Prestaciones (cesantías, intereses, primas)", ["PROVISION"]),
    ("Aportes (salud, pensión, riesgos, caja, parafiscales)", ["APORTE", "DIF REDONDEO"]),
    ("Horas extra y recargos", ["HORA", "RECARGO"]),
    ("Auxilio de transporte", ["AUXILIO DE TRANS", "AUXILIO  DE TRANS", "AUX TRANS", "AUXILIO TRANS"]),
    ("Sueldos y salarios", ["APOYO SOSTENIMIENTO"]),
    ("Auxilios y bonificaciones extralegales", ["AUXILIO", "AYUDA", "APOYO", "BONIFICACION"]),
    ("Sueldos y salarios", ["SUELDO", "SALARIO", "VAC ", "INCAP", "LICENCIA", "CALAMIDAD"]),
]


def concepto_sap(texto) -> str:
    """Concepto comparable con el presupuesto para un concepto de la nómina de SAP."""
    t = " ".join(norm(texto).split())
    for nombre, claves in _REGLAS_SAP:
        if any(" ".join(norm(k).split()) in t for k in claves):
            return nombre
    return "Otros conceptos"


def leer_presupuesto(archivo) -> pd.DataFrame:
    """Hojas '… <año>': por cargo, personas y costo por mes. Devuelve anio, mes, area, grupo, cargo, tipologia,
    personas, valor."""
    libro = pd.read_excel(archivo, sheet_name=None, header=None)
    filas = []
    for hoja, df in libro.items():
        if df.empty:
            continue
        enc = [str(x).strip() if pd.notna(x) else "" for x in df.iloc[0]]
        m = re.search(r"(20\d\d)", hoja) or re.search(r"(20\d\d)", " ".join(enc))
        if not m or not enc or norm(enc[0]) not in ("CATEGORIA", "CARGO"):
            continue
        anio = int(m.group(1))
        # columnas de meses: el primer bloque de meses después de los porcentajes = personas, el segundo = costo
        i_pct = max(i for i, c in enumerate(enc) if "%" in c)
        bloques, actual = [], {}
        for i in range(i_pct + 1, len(enc)):
            mes = _mes_de(enc[i]) if enc[i] else None
            if mes:
                if mes in actual:
                    bloques.append(actual)
                    actual = {}
                actual[mes] = i
            elif actual and norm(enc[i]).startswith("TOTAL"):
                bloques.append(actual)
                actual = {}
        if actual:
            bloques.append(actual)
        if len(bloques) < 2:
            continue
        c_pers, c_val = bloques[0], bloques[1]
        # columnas con las que el archivo calcula el costo (para repartirlo por concepto)
        def col(*prefijos):
            return next((i for i, c in enumerate(enc[:i_pct + 1]) if any(norm(c).startswith(p) for p in prefijos)),
                        None)
        i_t, i_bc, i_be = col("SUBSIDIO", "AUXILIO DE TRANSPORTE"), col("BONIFICACION CUMPL"), col("BONIFICACION EXTRA")
        i_seg, i_vac, i_pre = col("SEG"), col("VACACION"), col("PRESTACION")
        c_he = {_mes_de(c): i for i, c in enumerate(enc[:i_pct + 1]) if "HORA EXTRA" in norm(c)}
        i_area = next((i for i, c in enumerate(enc) if norm(c) == "CATEGORIA" and i > 0), None)
        i_tip = next((i for i, c in enumerate(enc) if norm(c).startswith("TIPOLOGIA")), None)
        grupo = ""
        for _, r in df.iloc[1:].iterrows():
            cargo = r.iloc[0]
            if pd.isna(cargo) or not str(cargo).strip():
                continue
            cargo = " ".join(str(cargo).split())
            salario = pd.to_numeric(r.iloc[1], errors="coerce")
            area = str(r.iloc[i_area]).strip() if i_area is not None and pd.notna(r.iloc[i_area]) else ""
            if pd.isna(salario) and not area:
                grupo = cargo                       # fila de título de grupo
                continue
            tip = str(r.iloc[i_tip]).strip() if i_tip is not None and pd.notna(r.iloc[i_tip]) else ""

            def num(i):
                v = pd.to_numeric(r.iloc[i], errors="coerce") if i is not None else None
                return 0.0 if v is None or pd.isna(v) else float(v)
            s_ = 0.0 if pd.isna(salario) else float(salario)
            t_, bc, be, seg, vac, pre = (num(i) for i in (i_t, i_bc, i_be, i_seg, i_vac, i_pre))
            for mes, i in c_val.items():
                personas = num(c_pers.get(mes))
                valor = num(i)
                h = num(c_he.get(mes))
                # misma fórmula del archivo: salario + bono cumplimiento + horas extra + transporte + extralegal
                # + seguridad social sobre (salario+bono+horas) + vacaciones sobre (salario+bono)
                # + prestaciones sobre (salario+bono+horas+transporte), todo por persona
                c = {"c_sueldo": s_, "c_bono_cump": bc, "c_horas": h, "c_transporte": t_, "c_extralegal": be,
                     "c_aportes": (s_ + bc + h) * seg, "c_vacaciones": (s_ + bc) * vac,
                     "c_prestaciones": (s_ + bc + h + t_) * pre}
                c = {k: v * personas for k, v in c.items()}
                c["c_otros"] = valor - sum(c.values())     # lo que la fórmula no explica (normalmente cero)
                filas.append({"anio": anio, "mes": mes, "area": area.upper(), "grupo": grupo, "cargo": cargo,
                              "tipologia": tip, "personas": personas, "valor": valor, **c})
    if not filas:
        raise ValueError("No encontré hojas de presupuesto de nómina (encabezado 'Categoría', meses de personas y "
                         "de costo).")
    out = pd.DataFrame(filas)
    out = out[(out["personas"] != 0) | (out["valor"] != 0)].copy()
    # el área sigue a los centros de coste: los grupos que son un área propia (Almacén, Compras, Aguas de Bogotá)
    gn = out["grupo"].map(norm)
    out.loc[gn.str.contains("AGUAS DE BOGOTA"), "area"] = "AGUAS DE BOGOTA"
    out.loc[gn == "ALMACEN", "area"] = "ALMACEN"
    out.loc[gn == "COMPRAS", "area"] = "COMPRAS"
    return out.reset_index(drop=True)


def _col(df, *nombres):
    for n in nombres:
        for c in df.columns:
            if norm(c) == norm(n):
                return c
    for n in nombres:
        for c in df.columns:
            if norm(c).startswith(norm(n)):
                return c
    return None


def leer_ejecucion(archivo) -> pd.DataFrame:
    """Base de nómina de SAP: anio, mes, cc, cargo, empleado_id, empleado, concepto, cuenta, valor."""
    df = pd.read_excel(archivo, dtype={"Centro de coste": str})
    c = {k: _col(df, *v) for k, v in {
        "cargo": ["CARGO"], "cc": ["Centro de coste"], "id": ["Número de personal", "Numero de personal"],
        "nombre": ["Nombre del empleado"], "concepto": ["Texto expl.CC-nómina", "Texto expl"],
        "cuenta": ["Número cuenta c/txt.", "Numero cuenta"], "importe": ["Importe"],
        "fecha": ["Período En, fecha de fin", "Periodo En, fecha de fin", "Per.Para, fecha fin"],
        "periodo": ["Período para nómina", "Periodo para nomina"]}.items()}
    faltan = [k for k in ("cargo", "cc", "importe") if c[k] is None]
    if faltan:
        raise ValueError("Faltan columnas en la base de nómina: " + ", ".join(faltan))
    fecha = pd.to_datetime(df[c["fecha"]], errors="coerce") if c["fecha"] else pd.Series(pd.NaT, df.index)
    if fecha.isna().all() and c["periodo"]:                      # 092026 -> sep 2026
        p = df[c["periodo"]].map(lambda v: f"{int(v):06d}" if pd.notna(v) and v else "")
        fecha = pd.to_datetime("01" + p, format="%d%m%Y", errors="coerce")
    # un archivo de nómina es un mes contable: las pocas filas de ajustes de meses anteriores van al mes principal
    principal = fecha.dt.to_period("M").mode()
    if len(principal):
        fecha = pd.Series(principal.iloc[0].to_timestamp(), index=df.index)
    txt = lambda k: df[c[k]].fillna("").astype(str).str.strip() if c[k] else ""
    out = pd.DataFrame({
        "anio": fecha.dt.year, "mes": fecha.dt.month, "cc": txt("cc").str.upper(),
        "cargo": txt("cargo").map(lambda s: " ".join(s.split())),
        "empleado_id": df[c["id"]].map(lambda v: str(int(v)) if pd.notna(v) else "") if c["id"] else "",
        "empleado": txt("nombre"), "concepto": txt("concepto"), "cuenta": txt("cuenta"),
        "valor": pd.to_numeric(df[c["importe"]], errors="coerce").fillna(0.0)})
    out = out[out["anio"].notna() & (out["valor"] != 0)]
    out = out[~out["cargo"].isin(["", "0", "nan"]) | (out["empleado_id"] != "")]
    out.loc[out["cargo"].isin(["", "0", "nan"]), "cargo"] = "(sin cargo)"
    out[["anio", "mes"]] = out[["anio", "mes"]].astype(int)
    return out.reset_index(drop=True)


def leer_ceco(archivo) -> pd.DataFrame:
    df = pd.read_excel(archivo, dtype=str)
    c_cc = _col(df, "CE.COSTE", "Centro de coste", "CECO")
    c_area = _col(df, "TEXTO BREVE", "Area", "Gerencia")
    if not c_cc or not c_area:
        raise ValueError("El archivo debe tener una columna de centro de coste y otra de área (TEXTO BREVE).")
    out = pd.DataFrame({"cc": df[c_cc].fillna("").str.strip().str.upper(),
                        "area": df[c_area].fillna("").str.strip().str.upper()})
    return out[(out["cc"] != "") & (out["area"] != "")].drop_duplicates("cc")
