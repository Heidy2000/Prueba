"""
Cartera de usuarios: base mensual "Reporte Cartera AñoMesTarifa" del sistema comercial (una fila por cuenta contrato
y mes de tarifa adeudado), en .txt (separado por |), .csv, Excel (hoja "Base") o comprimida (.zip, .bz2, .gz).

Replica el informe de cartera que se llevaba en Power BI (INFORME CARTERA PBI):
- Cada componente de la tarifa trae tres columnas: SERV_x (servicio facturado), SC_x (subsidio − / contribución +)
  e INT_x (intereses de mora). El saldo del componente es la suma de las tres.
- "Componentes propios" = medida COMPONENTES ACTUAL del PBI: barrido y limpieza (TBL), limpieza urbana (TLU),
  recolección y transporte (TRT), comercialización (TCS), CCS no aprovechable (CCSA_NPA) y servicios especiales.
- "Todos los componentes" = medida COMPONENTES ACTUAL RECAUDO = TOTAL_CARTERA: suma además aprovechamiento (APR),
  disposición final (TDF), tratamiento de lixiviados (TTL), VIAT, CCS aprovechable (CCSA_PA), el ajuste a la decena
  y los saldos a favor, que no son de la compañía.
- Variación ("Pago" en el PBI) por cuenta = saldo del mes actual − saldo del mes anterior: negativa = pagó o abonó,
  positiva = aumentó.
- Edad: meses entre el mes de corte y el mes de tarifa (ANIO_MES_TARIFA), igual que la TABLA EDAD CARTERA.

Al cargar, la base se resume: saldo por cuenta y tipo de cartera (con su saldo en cada edad) en un archivo
comprimido por mes (cartera_archivo), saldos por tipo, edad y componente (cartera_edades), usuarios con saldo por
tipo y edad (cartera_usuarios) y totales del mes (cartera_meses).
"""
import bz2
import gzip
import io
import re
import zipfile
from datetime import date, datetime

import numpy as np
import pandas as pd

# código: (nombre, propio de la compañía). None = no es componente del ingreso (ajustes y saldos a favor)
COMPONENTES = {
    "TBL": ("Barrido y limpieza", True),
    "TLU": ("Limpieza urbana", True),
    "TRT": ("Recolección y transporte", True),
    "TCS": ("Comercialización", True),
    "CCSA_NPA": ("CCS no aprovechable", True),
    "SERVESPECIALES": ("Servicios especiales", True),
    "APR": ("Aprovechamiento", False),
    "TDF": ("Disposición final", False),
    "TTL": ("Tratamiento de lixiviados", False),
    "VIAT": ("VIAT", False),
    "CCSA_PA": ("CCS aprovechable", False),
    "AJUSTE": ("Ajuste a la decena", None),
    "SALDOS_A_FAVOR": ("Saldos a favor", None),
}
PROPIOS = [c for c, (_, p) in COMPONENTES.items() if p is True]
TERCEROS = [c for c, (_, p) in COMPONENTES.items() if p is False]
OTROS = [c for c, (_, p) in COMPONENTES.items() if p is None]


def columnas_de(comp: str) -> dict:
    """Columnas de la base que forman un componente: {'servicio': ..., 'subsidio': ..., 'intereses': ...}."""
    if comp == "SERVESPECIALES":
        return {"servicio": "SERVESPECIALES", "intereses": "INT_SERVESPECIALES"}
    if comp == "AJUSTE":
        return {"servicio": "AJUSTE_DECENA", "intereses": "INT_AJUST_DEC"}
    if comp == "SALDOS_A_FAVOR":
        return {"servicio": "SALDOS_A_FAVOR"}
    return {"servicio": f"SERV_{comp}", "subsidio": f"SC_{comp}", "intereses": f"INT_{comp}"}


NUMERICAS = sorted({c for comp in COMPONENTES for c in columnas_de(comp).values()})
TEXTO = ["CUENTA", "ANIO_MES_TARIFA", "NAME_CLASE_USO", "ESTRATO", "FECHA CORTE", "TIPO_CARTERA"]
OBLIGATORIAS = ["CUENTA", "ANIO_MES_TARIFA", "TIPO_CARTERA"] + [f"SERV_{c}" for c in ("TBL", "TRT", "APR")]

# edad de cartera (orden, nombre corto, rango que usaba la TABLA EDAD CARTERA del PBI)
EDADES = {
    1: ("Vigente", "< 0 VIGENCIA"),
    2: ("1 a 30 días", "< 30 DIAS"),
    3: ("31 a 60 días", "< 60 DIAS"),
    4: ("61 a 90 días", ">60 A 90 DIAS"),
    5: ("91 a 180 días", ">90 A 180 DIAS"),
    6: ("181 a 360 días", ">180 A 360 DIAS"),
    7: ("361 a 720 días", "> 360"),
    8: ("721 a 1.080 días", "> 720"),
    9: ("Más de 1.080 días", "> *1080"),
}

# columnas de saldo por edad de cada cuenta: p1..p9 (componentes propios) y t1..t9 (todos)
POR_EDAD = [f"p{i}" for i in range(1, 10)] + [f"t{i}" for i in range(1, 10)]

# rangos para deterioro y castigo (edad 7 = más de 360 días, 8 = más de 720, 9 = más de 1.080)
RANGOS = {"Más de 1 año": (7, 8, 9), "Más de 2 años": (8, 9), "Más de 3 años": (9,)}


def columnas_saldo(propio: bool, edades=None) -> list[str]:
    """Columnas que forman el saldo de una cuenta: propio o total, en todas las edades o solo en las elegidas."""
    if not edades or set(edades) == set(range(1, 10)):
        return ["propio" if propio else "total"]
    letra = "p" if propio else "t"
    return [f"{letra}{int(i)}" for i in sorted(edades)]


TIPOS = {"CVAC": "CVAC", "CVAN": "CVAN", "DS/DI": "DS/DI", "FINAN": "Financiada", "SaFa": "Saldo a favor"}

TOL = 0.01   # un saldo menor a un centavo cuenta como cero


def edad_de(periodo: int, tarifa) -> np.ndarray:
    """Orden de edad (1-9) a partir de los meses entre el mes de corte (AAAAMM) y el mes de tarifa (AAAAMM)."""
    t = pd.to_numeric(pd.Series(tarifa), errors="coerce").fillna(periodo).astype(int).to_numpy()
    meses = (periodo // 100 - t // 100) * 12 + (periodo % 100 - t % 100)
    return np.select([meses <= 0, meses == 1, meses == 2, meses == 3, meses <= 6, meses <= 11, meses <= 23,
                      meses <= 35], [1, 2, 3, 4, 5, 6, 7, 8], 9)


def grupo_uso(nombre: str) -> str:
    """Agrupa la clase de uso: residencial por estrato (incluye multiusuario) y no residencial."""
    n = str(nombre or "").lower()
    m = re.search(r"estrato\s*(\d)", n)
    if "residencial" in n and "no residencial" not in n and m:
        return f"Residencial estrato {m.group(1)}"
    if "gran prod" in n:
        return "Gran productor"
    if "pequ" in n:
        return "Pequeño productor"
    if "no residencial" in n or "multiusuario" in n:
        return "Multiusuario no residencial"
    return "Otros" if not n.strip() else str(nombre).strip()


# ------------------------------------------------------------------------------------------------ lectura
def _abrir(contenido: bytes, nombre: str) -> tuple[str, bytes | io.IOBase]:
    """Descomprime si hace falta. Devuelve (tipo, fuente) con tipo 'excel' o 'texto'."""
    n = nombre.lower()
    if n.endswith(".zip") or contenido[:4] == b"PK\x03\x04" and not n.endswith((".xlsx", ".xlsm")):
        z = zipfile.ZipFile(io.BytesIO(contenido))
        miembros = [m for m in z.namelist() if not m.endswith("/")]
        if not miembros:
            raise ValueError("El .zip está vacío.")
        m = max(miembros, key=lambda x: z.getinfo(x).file_size)
        if m.lower().endswith((".xlsx", ".xlsm", ".zip", ".bz2", ".gz")):
            return _abrir(z.read(m), m)
        return "texto", z.open(m)          # texto: se lee por partes, sin descomprimir todo en memoria
    if n.endswith(".bz2") or contenido[:3] == b"BZh":
        return "texto", bz2.BZ2File(io.BytesIO(contenido))
    if n.endswith(".gz") or contenido[:2] == b"\x1f\x8b":
        return "texto", gzip.GzipFile(fileobj=io.BytesIO(contenido))
    if n.endswith((".xlsx", ".xlsm")) or contenido[:4] == b"PK\x03\x04":
        return "excel", contenido
    return "texto", io.BytesIO(contenido)


class _Encadenado(io.RawIOBase):
    """Une lo ya leído (muestra) con el resto del archivo, para leer por partes sin copiarlo entero."""

    def __init__(self, inicio: bytes, resto):
        self.inicio, self.resto = inicio, resto

    def readable(self):
        return True

    def readinto(self, b):
        if self.inicio:
            n = min(len(b), len(self.inicio))
            b[:n] = self.inicio[:n]
            self.inicio = self.inicio[n:]
            return n
        dato = self.resto.read(len(b))
        b[:len(dato)] = dato
        return len(dato)


def _trozos_texto(fuente, tam: int):
    cab = fuente.readline()
    texto = cab.decode("utf-8-sig", errors="replace")
    sep = max(["|", ";", "\t", ","], key=texto.count)
    nombres = [c.strip().strip('"') for c in texto.rstrip("\r\n").split(sep)]
    faltan = [c for c in OBLIGATORIAS if c not in nombres]
    if faltan:
        raise ValueError(f"No encuentro las columnas {', '.join(faltan)}. ¿Es la base 'Reporte Cartera AñoMesTarifa'?")
    usar = [c for c in nombres if c in set(NUMERICAS) | set(TEXTO)]
    # decimales: punto (como la base del sistema) salvo que el separador no sea coma y vengan con coma
    muestra = fuente.read(200_000)
    dec = "," if sep != "," and re.search(rb"\d,\d", muestra) and not re.search(rb"\d\.\d", muestra) else "."
    # números como número desde la lectura (no como texto): un bloque ocupa ~10 veces menos memoria
    tipos = {c: ("float64" if c in NUMERICAS else str) for c in usar}
    lector = pd.read_csv(_Encadenado(muestra, fuente), sep=sep, names=nombres, header=None, usecols=usar, dtype=tipos,
                         chunksize=tam, decimal=dec, thousands="." if dec == "," else None, encoding="latin-1",
                         quoting=3)
    yield from lector


def _trozos_excel(contenido: bytes, tam: int):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(contenido), read_only=True, data_only=True)
    hoja = next((h for h in wb.sheetnames if h.strip().lower() == "base"), wb.sheetnames[0])
    filas = wb[hoja].iter_rows(values_only=True)
    cab = None
    for f in filas:                  # el encabezado puede no estar en la primera fila
        if f and "CUENTA" in [str(x).strip() for x in f if x is not None]:
            cab = [str(x).strip() if x is not None else "" for x in f]
            break
    if cab is None:
        raise ValueError("No encuentro el encabezado (columna CUENTA) en el Excel.")
    faltan = [c for c in OBLIGATORIAS if c not in cab]
    if faltan:
        raise ValueError(f"No encuentro las columnas {', '.join(faltan)}. ¿Es la base 'Reporte Cartera AñoMesTarifa'?")
    idx = [(i, c) for i, c in enumerate(cab) if c in set(NUMERICAS) | set(TEXTO)]
    buf = []
    for f in filas:
        buf.append([f[i] if i < len(f) else None for i, _ in idx])
        if len(buf) >= tam:
            yield pd.DataFrame(buf, columns=[c for _, c in idx])
            buf = []
    if buf:
        yield pd.DataFrame(buf, columns=[c for _, c in idx])
    wb.close()


def _fecha_corte(serie: pd.Series) -> int | None:
    v = serie.dropna()
    if v.empty:
        return None
    x = v.iloc[0]
    if isinstance(x, (datetime, date, pd.Timestamp)):
        f = pd.Timestamp(x)
    else:
        s = str(x).strip()
        f = pd.to_datetime(s, dayfirst=not re.match(r"\d{4}-", s), errors="coerce")
    return None if pd.isna(f) else f.year * 100 + f.month


class _Acumulador:
    """Suma por llave (cuenta y tipo) bloque a bloque en arreglos que crecen al doble: la memoria es la del
    resultado, no la de todos los bloques juntos."""

    def __init__(self, n_sumas: int, n_max: int):
        self.n, self.cap = 0, 0
        self.llaves = np.empty(0, dtype="int64")
        self.sumas = np.zeros((0, n_sumas))
        self.maximos = np.zeros((0, n_max), dtype="int16")
        self._indice = pd.Index([], dtype="int64")

    def _crecer(self, falta: int):
        cap = max(self.cap * 2, self.n + falta, 100_000)
        for nombre, relleno in (("llaves", 0), ("sumas", 0.0), ("maximos", -1)):
            viejo = getattr(self, nombre)
            nuevo = np.full((cap,) + viejo.shape[1:], relleno, dtype=viejo.dtype)
            nuevo[:self.n] = viejo[:self.n]
            setattr(self, nombre, nuevo)
        self.cap = cap

    def sumar(self, llaves, sumas, maximos):
        pos = self._indice.get_indexer(llaves)
        nuevas = pos < 0
        k = int(nuevas.sum())
        if k:
            if self.n + k > self.cap:
                self._crecer(k)
            pos[nuevas] = np.arange(self.n, self.n + k)
            self.llaves[self.n:self.n + k] = llaves[nuevas]
            self.n += k
            self._indice = pd.Index(self.llaves[:self.n])
        self.sumas[pos] += sumas
        self.maximos[pos] = np.maximum(self.maximos[pos], maximos)

    def resultado(self):
        return self.llaves[:self.n], self.sumas[:self.n], self.maximos[:self.n]


def leer(contenido: bytes, nombre: str, periodo: int | None = None, avance=None, tam: int = 50_000) -> dict:
    """Lee y resume la base. Devuelve dict con periodo, cuentas, edades, usuarios, filas y avisos.
    avance(filas_leidas) se llama después de cada bloque."""
    tipo, fuente = _abrir(contenido, nombre)
    trozos = _trozos_excel(fuente, tam) if tipo == "excel" else _trozos_texto(fuente, tam)
    # para no agotar la memoria del servidor (un mes trae ~1 millón de filas): tipo y clase de uso van como números
    # (con su diccionario) y lo acumulado se compacta cada pocos bloques
    cod_tipo, cod_clase = {}, {}
    SUMAS = ["propio", "terceros", "ajustes", "total", *POR_EDAD]
    MAXIMOS = ["clase", "estrato", "edad", "edad_propio"]
    acum = _Acumulador(len(SUMAS), len(MAXIMOS))
    p_edades = []
    filas, cortes = 0, set()
    per = periodo
    for ch in trozos:
        ch = ch[pd.to_numeric(ch["CUENTA"], errors="coerce").notna()]
        if ch.empty:
            continue
        if "FECHA CORTE" in ch:
            c = _fecha_corte(ch["FECHA CORTE"])
            if c:
                cortes.add(c)
        if per is None:
            per = max(cortes) if cortes else None
            if per is None:
                raise ValueError("La base no trae FECHA CORTE: indica el mes de la cartera.")
        filas += len(ch)
        tipo_txt = ch["TIPO_CARTERA"].fillna("").astype(str).str.strip().replace("", "SIN TIPO")
        clase_txt = ch["NAME_CLASE_USO"].fillna("").astype(str).str.strip() if "NAME_CLASE_USO" in ch \
            else pd.Series("", index=ch.index)
        for t in tipo_txt.unique():
            cod_tipo.setdefault(t, len(cod_tipo))
        for t in clase_txt.unique():
            cod_clase.setdefault(t, len(cod_clase))
        comp = {}
        for k in COMPONENTES:
            for parte, col in columnas_de(k).items():
                comp[(k, parte)] = (pd.to_numeric(ch[col], errors="coerce").fillna(0.0).to_numpy(dtype=float)
                                    if col in ch else np.zeros(len(ch)))
        edad = edad_de(per, ch["ANIO_MES_TARIFA"])
        tipo_c = tipo_txt.map(cod_tipo).to_numpy(dtype="int16")
        propio = sum(v for (k, _), v in comp.items() if k in PROPIOS)
        terceros = sum(v for (k, _), v in comp.items() if k in TERCEROS)
        ajustes = sum(v for (k, _), v in comp.items() if k in OTROS)
        total = propio + terceros + ajustes
        d = pd.DataFrame({
            "cuenta": pd.to_numeric(ch["CUENTA"], errors="coerce").to_numpy(dtype="int64"), "tipo": tipo_c,
            "clase": clase_txt.map(cod_clase).to_numpy(dtype="int16"),
            "estrato": (pd.to_numeric(ch["ESTRATO"], errors="coerce").fillna(-1).to_numpy(dtype="int16")
                        if "ESTRATO" in ch else np.full(len(ch), -1, dtype="int16")),
            "edad_t": np.where(np.abs(total) >= TOL, edad, 0).astype("int8"),
            "edad_p": np.where(np.abs(propio) >= TOL, edad, 0).astype("int8"),
            "propio": propio, "terceros": terceros, "ajustes": ajustes, "total": total})
        for i in EDADES:      # saldo de la cuenta en cada edad (propio y total)
            d[f"p{i}"] = np.where(edad == i, propio, 0.0)
            d[f"t{i}"] = np.where(edad == i, total, 0.0)
        d["llave"] = d["cuenta"] * 64 + d["tipo"]           # cuenta y tipo de cartera en un solo número
        g = d.groupby("llave", sort=False).agg(**{c: (c, "sum") for c in SUMAS}, clase=("clase", "max"),
                                                 estrato=("estrato", "max"), edad=("edad_t", "max"),
                                                 edad_propio=("edad_p", "max"))
        acum.sumar(g.index.to_numpy(), g[SUMAS].to_numpy(), g[MAXIMOS].to_numpy(dtype="int16"))
        # saldos por tipo, edad y componente
        e = pd.DataFrame({f"{k}|{parte}": v for (k, parte), v in comp.items()})
        e["tipo"], e["edad"] = tipo_c, edad
        p_edades.append(e.groupby(["tipo", "edad"]).sum())
        del d, e, comp, ch, propio, terceros, ajustes, total
        if avance:
            avance(filas)
    if acum.n == 0:
        raise ValueError("La base no tiene filas con cuenta.")
    avisos = []
    if len(cortes) > 1:
        avisos.append(f"La base trae varias fechas de corte ({', '.join(map(str, sorted(cortes)))}); "
                      f"se toma {per}.")
    llaves, sumas, maximos = acum.resultado()
    cu = pd.DataFrame({"cuenta": llaves // 64, "tipo": (llaves % 64).astype("int16")})
    for k, c in enumerate(SUMAS):
        cu[c] = sumas[:, k]
    for k, c in enumerate(MAXIMOS):
        cu[c] = maximos[:, k]
    del acum, sumas, maximos, llaves
    tipos_l = pd.Index(sorted(cod_tipo, key=cod_tipo.get))
    clases_l = pd.Index(sorted(cod_clase, key=cod_clase.get))
    cu["tipo"] = pd.Categorical.from_codes(cu["tipo"].astype(int), categories=tipos_l)
    cu["clase"] = pd.Categorical.from_codes(cu["clase"].astype(int), categories=clases_l)
    cu["estrato"] = cu["estrato"].where(cu["estrato"] >= 0).astype("Int8")
    cu["edad"], cu["edad_propio"] = cu["edad"].astype("int8"), cu["edad_propio"].astype("int8")
    ed = pd.concat(p_edades).groupby(level=[0, 1]).sum()
    ed.index = ed.index.set_levels(tipos_l[ed.index.levels[0]], level=0)
    ed.columns = pd.MultiIndex.from_tuples([tuple(c.split("|")) for c in ed.columns], names=["componente", "parte"])
    ed = ed.stack("componente", future_stack=True).reset_index()
    for parte in ("servicio", "subsidio", "intereses"):
        if parte not in ed:
            ed[parte] = 0.0
    ed = ed[["tipo", "edad", "componente", "servicio", "subsidio", "intereses"]].fillna(0.0)
    ed = ed[(ed[["servicio", "subsidio", "intereses"]].abs() >= TOL).any(axis=1)]
    # usuarios con saldo por tipo y edad: cuentas con saldo en esa edad (columnas p1..p9 y t1..t9)
    us = []
    for i in EDADES:
        g = cu.assign(n=(cu[f"t{i}"].abs() >= TOL).astype(int), np_=(cu[f"p{i}"].abs() >= TOL).astype(int))
        g = g.groupby("tipo", observed=True)[["n", "np_"]].sum()
        us.append(pd.DataFrame({"tipo": g.index, "edad": i, "usuarios": g["n"].values,
                                "usuarios_propio": g["np_"].values}))
    us = pd.concat(us)
    us = us[(us["usuarios"] > 0) | (us["usuarios_propio"] > 0)].reset_index(drop=True)
    return dict(periodo=int(per), cuentas=cu, edades=ed, usuarios=us, filas=filas, avisos=avisos)


def resumen(res: dict) -> dict:
    if "resumen" in res:
        return res["resumen"]
    cu = res["cuentas"]
    return dict(propio=float(cu["propio"].sum()), terceros=float(cu["terceros"].sum()),
                ajustes=float(cu["ajustes"].sum()), total=float(cu["total"].sum()),
                usuarios=int(cu.groupby("cuenta")["total"].sum().abs().ge(TOL).sum()),
                usuarios_propio=int(cu.groupby("cuenta")["propio"].sum().abs().ge(TOL).sum()))


def _a_parquet(cu: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    cu.to_parquet(buf, compression="zstd", index=False)
    return buf.getvalue()


def comprimir(res: dict) -> dict:
    """Deja el resultado listo para guardar ocupando poca memoria: el detalle por cuenta pasa a un archivo
    comprimido (~8 MB en vez de ~150 MB) y se guarda el resumen."""
    if "cuentas" in res:
        res["resumen"] = resumen(res)
        res["parquet"] = _a_parquet(res.pop("cuentas"))
    liberar_memoria()
    return res


def liberar_memoria():
    """Devuelve al sistema la memoria que quedó libre después de leer una base grande (Linux)."""
    import gc
    gc.collect()
    try:
        import ctypes
        ctypes.CDLL("libc.so.6").malloc_trim(0)
    except Exception:
        pass


def cuentas_de(leer, empresa: str, periodo: int, columnas=None) -> pd.DataFrame:
    """Saldos por cuenta contrato y tipo de cartera de un mes (archivo comprimido guardado al cargar). Con
    'columnas' solo lee esas (así la vista no tiene en memoria el mes completo)."""
    columnas = list(columnas) if columnas else None
    d = leer("SELECT datos FROM cartera_archivo WHERE empresa=? AND periodo=?", (empresa, int(periodo)))
    if d.empty:                                    # mes sin archivo: tabla vacía con los tipos correctos
        vacio = pd.DataFrame({"cuenta": pd.Series(dtype="int64"), "tipo": pd.Series(dtype="object"),
                              "clase": pd.Series(dtype="object"), "estrato": pd.Series(dtype="Int8"),
                              "edad": pd.Series(dtype="int8"), "edad_propio": pd.Series(dtype="int8"),
                              **{c: pd.Series(dtype="float64")
                                 for c in ["propio", "terceros", "ajustes", "total", *POR_EDAD]}})
        return vacio[columnas] if columnas else vacio
    return pd.read_parquet(io.BytesIO(bytes(d["datos"].iloc[0])), columns=columnas)


def guardar(con, empresa: str, res: dict, archivo: str = "", usuario: str = "", fecha: str = ""):
    """Reemplaza el mes. El detalle por cuenta (más de 600.000 filas) se guarda como un solo archivo comprimido
    (unos 7 MB) para no llenar la base; los resúmenes por edad y componente quedan en tablas."""
    per = res["periodo"]
    for t in ("cartera_archivo", "cartera_edades", "cartera_usuarios", "cartera_meses"):
        con.execute(f"DELETE FROM {t} WHERE empresa=? AND periodo=?", (empresa, per))
    con.execute("INSERT INTO cartera_archivo(empresa, periodo, datos) VALUES (?, ?, ?)",
                (empresa, per, res["parquet"] if "parquet" in res else _a_parquet(res["cuentas"])))
    con.anexar(res["edades"].assign(empresa=empresa, periodo=per)
               [["empresa", "periodo", "tipo", "edad", "componente", "servicio", "subsidio", "intereses"]],
               "cartera_edades")
    con.anexar(res["usuarios"].assign(empresa=empresa, periodo=per)
               [["empresa", "periodo", "tipo", "edad", "usuarios", "usuarios_propio"]], "cartera_usuarios")
    r = resumen(res)
    con.anexar(pd.DataFrame([dict(empresa=empresa, periodo=per, archivo=archivo, filas=int(res["filas"]),
                                  usuarios=r["usuarios"], usuarios_propio=r["usuarios_propio"], propio=r["propio"],
                                  terceros=r["terceros"], ajustes=r["ajustes"], total=r["total"], fecha=fecha,
                                  usuario=usuario)]), "cartera_meses")
    con.commit()


def borrar(con, empresa: str, periodo: int):
    for t in ("cartera_archivo", "cartera_edades", "cartera_usuarios", "cartera_meses"):
        con.execute(f"DELETE FROM {t} WHERE empresa=? AND periodo=?", (empresa, periodo))
    con.commit()


def ddl(motor: str) -> str:
    binario = "BYTEA" if motor == "postgres" else "BLOB"
    return f"""
    CREATE TABLE IF NOT EXISTS cartera_archivo(empresa TEXT, periodo INT, datos {binario});
    CREATE TABLE IF NOT EXISTS cartera_edades(empresa TEXT, periodo INT, tipo TEXT, edad INT, componente TEXT,
        servicio REAL, subsidio REAL, intereses REAL);
    CREATE TABLE IF NOT EXISTS cartera_usuarios(empresa TEXT, periodo INT, tipo TEXT, edad INT, usuarios INT,
        usuarios_propio INT);
    CREATE TABLE IF NOT EXISTS cartera_meses(empresa TEXT, periodo INT, archivo TEXT, filas INT, usuarios INT,
        usuarios_propio INT, propio REAL, terceros REAL, ajustes REAL, total REAL, fecha TEXT, usuario TEXT)
"""


def nombre_mes(periodo: int, largo: bool = False) -> str:
    from ui import MESES, MESES_L
    m = int(periodo) % 100
    return f"{(MESES_L[m - 1] if largo else MESES[m - 1])} {int(periodo) // 100}"
