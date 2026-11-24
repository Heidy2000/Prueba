"""
Informe de personal, dotación y otros gastos de RH (el Excel que Financiera envía a Gestión Humana)
====================================================================================================
Se arma con lo que ya está cargado en el panel, sin tocar Excel a mano:
- Presupuesto: tabla `presupuesto`. Ene–ago contra una versión (por defecto la Inicial) y sep–dic contra otra
  (por defecto la 8+4 / UPA), como en el archivo original.
- Ejecución: tabla `ejecucion` (base de SAP).
- Dotación: kárdex (`dotacion_ejec`) + registros contables directos al costo, con el cruce y las clasificaciones
  del módulo Dotación y EPP.

Hojas: Resumen · Mes a mes · Dotación comparativo · Detalle dotación · Detalle movimientos.
No usa Streamlit: recibe la función `leer` del panel y devuelve los bytes del Excel.
"""
from io import BytesIO

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import cargas
import dotacion

MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
MESES_L = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre",
           "noviembre", "diciembre"]

# (grupo, concepto del informe, clase, palabras del sub-rubro; vacío = toda la clase)
CONCEPTOS = [
    ("Personal", "Personal", "3.0", []),
    ("Compras", "Dotación y EPP's", "3.1", ["dotacion"]),
    ("Compras", "Actividades de Seguridad y Salud en el Trabajo", "3.1", ["seguridad y salud", "seguridad en el trabajo"]),
    ("Compras", "Actividades de bienestar", "3.1", ["bienestar"]),
    ("Compras", "Actividades de formación", "3.1", ["formacion"]),
    ("Compras", "Otras actividades de personal", "3.1", ["otras actividades de personal"]),
    ("Servicios exteriores", "Beneficios en salud", "3.3", ["servicios en salud", "beneficios en salud"]),
    ("Suministros", "Transporte operarios", "3.6", ["transporte operarios", "transportes"]),
    ("Suministros", "Servicio de comedor", "3.9", ["casino", "comedor"]),
]


def concepto_de(codigo, sub_rubro) -> str | None:
    """Concepto del informe al que pertenece un movimiento o una línea de presupuesto (None = no va)."""
    cod, sub = str(codigo or ""), cargas.norm(sub_rubro)
    for _, nombre, clase, claves in CONCEPTOS:
        if cod == clase and (not claves or any(k in sub for k in claves)):
            return nombre
    return None


def _marcar(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df.assign(concepto_rh=pd.Series(dtype=str))
    d = df.assign(concepto_rh=[concepto_de(c, s) for c, s in zip(df["codigo"], df["sub_rubro"])])
    return d[d["concepto_rh"].notna()]


def preparar(leer, empresa: str, anio: int, version_1: str, version_2: str, corte_v: int = 8):
    """Presupuesto (meses 1..corte_v con version_1 y el resto con version_2) y ejecución por concepto y mes."""
    p1 = _marcar(leer("SELECT codigo, sub_rubro, mes, valor FROM presupuesto WHERE anio=? AND empresa=? AND version=?",
                      (int(anio), empresa, version_1)))
    p2 = _marcar(leer("SELECT codigo, sub_rubro, mes, valor FROM presupuesto WHERE anio=? AND empresa=? AND version=?",
                      (int(anio), empresa, version_2)))
    ppto = pd.concat([p1[p1["mes"] <= corte_v], p2[p2["mes"] > corte_v]], ignore_index=True)
    cols = "anio, mes, fecha, codigo, sub_rubro, cc_nombre, cuenta, cuenta_nombre, documento, tercero, texto, valor"
    ejec = _marcar(leer(f"SELECT {cols} FROM ejecucion WHERE anio=? AND empresa=?", (int(anio), empresa)))
    return ppto, ejec


def tabla_mensual(ppto: pd.DataFrame, ejec: pd.DataFrame) -> pd.DataFrame:
    """Una fila por concepto: presupuesto y ejecución de cada mes."""
    nombres = [c for _, c, *_ in CONCEPTOS]
    p = ppto.pivot_table(index="concepto_rh", columns="mes", values="valor", aggfunc="sum") \
        .reindex(index=nombres, columns=range(1, 13)).fillna(0.0)
    e = ejec.pivot_table(index="concepto_rh", columns="mes", values="valor", aggfunc="sum") \
        .reindex(index=nombres, columns=range(1, 13)).fillna(0.0)
    return pd.concat({"ppto": p, "ejec": e}, axis=1)


def resumen(t: pd.DataFrame, hasta: int) -> pd.DataFrame:
    """Acumulado a `hasta`: presupuesto, ejecutado, diferencia y % por concepto."""
    meses = list(range(1, hasta + 1))
    r = pd.DataFrame({"ppto": t["ppto"][meses].sum(axis=1), "ejec": t["ejec"][meses].sum(axis=1)})
    r["dif"] = r["ejec"] - r["ppto"]
    r["pct"] = r["ejec"] / r["ppto"].where(r["ppto"] != 0)
    r["grupo"] = [next(g for g, c, *_ in CONCEPTOS if c == i) for i in r.index]
    return r


# ------------------------------------------------------------------ formato del Excel
AZUL, AZUL_SUAVE, GRIS = "1F5FBF", "EAF1FB", "F2F4F7"
_borde = Border(bottom=Side(style="thin", color="D9E1EC"))
_pct = '0.0%'


def _encabezado(ws, fila: int, nombres: list, desde: int = 1):
    for j, n in enumerate(nombres, start=desde):
        c = ws.cell(row=fila, column=j, value=n)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=AZUL)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[fila].height = 30


def _fila(ws, fila: int, valores: list, formatos: list, estilo: str = ""):
    for j, (v, f) in enumerate(zip(valores, formatos), start=1):
        c = ws.cell(row=fila, column=j, value=None if (isinstance(v, float) and pd.isna(v)) else v)
        if f:
            c.number_format = f
        c.border = _borde
        if estilo in ("grupo", "total"):
            c.font = Font(bold=True, color="10284A" if estilo == "grupo" else "FFFFFF")
            c.fill = PatternFill("solid", fgColor=AZUL_SUAVE if estilo == "grupo" else "10284A")


# ------------------------------------------------------------------ formato igual al archivo de Gestión Humana
MESES_MAY = [m.upper() for m in MESES_L]
F_CONT = '_-* #,##0_-;\\-* #,##0_-;_-* "-"??_-;_-@_-'          # contabilidad sin decimales, como el archivo
AZUL_ENC, AZUL_OSC, CELESTE = "305496", "002060", "B7DEE8"
BLOQUES = [("Compras", ["Dotación y EPP's", "Actividades de Seguridad y Salud en el Trabajo", "Actividades de bienestar",
                        "Actividades de formación", "Otras actividades de personal"]),
           ("Servicios Exteriores", ["Beneficios en salud"]),
           ("Suministros", ["Transporte operarios", "Servicio de comedor"])]
_fino = Side(style="thin", color="000000")
_pelo = Side(style="hair", color="808080")


def _celda(ws, fila, col, valor, negrita=False, fmt=None, fondo=None, color=None, tam=10, centro=False, borde=None):
    c = ws.cell(row=fila, column=col, value=valor)
    c.font = Font(name="Calibri", size=tam, bold=negrita, color=color)
    if fmt:
        c.number_format = fmt
    if fondo:
        c.fill = PatternFill("solid", fgColor=fondo)
    if centro:
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if borde:
        c.border = borde
    return c


def _hoja_periodo(wb, nombre: str, t: pd.DataFrame, meses: list[int], hasta: int, titulo: str):
    """Hoja como 'PRESUPUESTO' / 'UPA' del archivo: presupuesto a la izquierda, ejecución a la derecha."""
    ws = wb.create_sheet(nombre)
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 80
    n = len(meses)
    c_tp = 2 + n                       # TOTAL PPTO ACUM
    c_l = c_tp + 2                     # etiqueta del bloque de ejecución
    c_e0 = c_l + 1
    c_te = c_e0 + n                    # TOTAL EJE
    c_df = c_te + 1                    # DIFERENCIA ACUM
    L = get_column_letter
    _celda(ws, 1, 1, titulo, True, tam=11, color="10284A")
    _celda(ws, 2, 2, "PRESUPUESTO", True, centro=True, fondo=CELESTE)
    ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=c_tp)
    _celda(ws, 2, c_l, "EJECUCION", True, centro=True, fondo=CELESTE)
    ws.merge_cells(start_row=2, start_column=c_l, end_row=2, end_column=c_df)
    for j, m in enumerate(meses):
        _celda(ws, 3, c_e0 + j, m, centro=True, color="808080")
        _celda(ws, 4, 2 + j, MESES_MAY[m - 1], True, centro=True)
        _celda(ws, 4, c_e0 + j, MESES_MAY[m - 1], True, centro=True)
    _celda(ws, 4, c_tp, "TOTAL PPTO ACUM", True, centro=True)
    _celda(ws, 4, c_te, "TOTAL EJE", True, centro=True)
    _celda(ws, 4, c_df, "DIFERENCIA ACUM", True, centro=True)
    con_ejec = [j for j, m in enumerate(meses) if m <= hasta]

    def fila_valores(fila, etiqueta, concepto, negrita=False):
        _celda(ws, fila, 1, etiqueta, negrita)
        _celda(ws, fila, c_l, etiqueta, negrita)
        for j, m in enumerate(meses):
            _celda(ws, fila, 2 + j, float(t["ppto"].at[concepto, m]), fmt=F_CONT)
            _celda(ws, fila, c_e0 + j, float(t["ejec"].at[concepto, m]) if m <= hasta else None, fmt=F_CONT)
        _formulas(fila)

    def _formulas(fila, negrita=False, fondo=None):
        _celda(ws, fila, c_tp, f"=SUM({L(2)}{fila}:{L(1 + n)}{fila})", negrita, F_CONT, fondo)
        _celda(ws, fila, c_te, f"=SUM({L(c_e0)}{fila}:{L(c_e0 + n - 1)}{fila})", negrita, F_CONT, fondo)
        if con_ejec:            # diferencia contra el presupuesto de los meses que ya tienen ejecución
            _celda(ws, fila, c_df, f"={L(c_te)}{fila}-SUM({L(2)}{fila}:{L(2 + con_ejec[-1])}{fila})", negrita,
                   '#,##0_ ;[Red]\\-#,##0\\ ', fondo)

    fila = 6
    fila_valores(fila, "Personal", "Personal", True)
    fila += 2
    for bloque, conceptos in BLOQUES:
        _celda(ws, fila, 1, bloque, True)
        _celda(ws, fila, c_l, bloque, True)
        ini = fila + 1
        for c in conceptos:
            fila += 1
            fila_valores(fila, c, c)
        fila += 1
        _celda(ws, fila, 1, "TOTALES ", True, fondo=CELESTE)
        _celda(ws, fila, c_l, "TOTALES ", True, fondo=CELESTE)
        for j in range(n):
            for col in (2 + j, c_e0 + j):
                if col >= c_e0 and meses[j] > hasta:
                    _celda(ws, fila, col, None, fondo=CELESTE)
                    continue
                _celda(ws, fila, col, f"=SUM({L(col)}{ini}:{L(col)}{fila - 1})", True, F_CONT, CELESTE)
        _formulas(fila, True, CELESTE)
        fila += 3
    ws.column_dimensions["A"].width = 44
    ws.column_dimensions[L(c_l)].width = 44
    for col in list(range(2, c_tp + 1)) + list(range(c_e0, c_df + 1)):
        ws.column_dimensions[L(col)].width = 15.5
    ws.column_dimensions[L(c_tp + 1)].width = 3
    ws.freeze_panes = "B5"


def generar(leer, empresa: str, nombre_empresa: str, anio: int, hasta: int, version_1: str, version_2: str,
            corte_v: int = 8) -> bytes:
    ppto, ejec = preparar(leer, empresa, anio, version_1, version_2, corte_v)
    t = tabla_mensual(ppto, ejec)
    wb = Workbook()

    # ---------------- RESUMEN (como 'Hoja1'): ejecución, presupuesto y diferencia acumulados al corte
    ws = wb.active
    ws.title = "RESUMEN"
    ws.sheet_view.showGridLines = False
    r = resumen(t, hasta)
    _celda(ws, 1, 1, f"{nombre_empresa} · PERSONAL, DOTACIÓN Y OTROS GASTOS RH · acumulado a {MESES_L[hasta - 1]} {anio}",
           True, tam=11, color="10284A")
    caja = Border(left=_fino, right=_fino, top=_fino, bottom=_fino)
    fila = 3
    for bloque, conceptos in [("PERSONAL", ["Personal"])] + [(b.upper(), c) for b, c in BLOQUES]:
        for j, txt in enumerate([bloque, "EJECUCION", "PRESUPUESTO", "DIFERENCIA", "% EJECUCIÓN"], start=1):
            _celda(ws, fila, j, txt, j == 1, fondo=AZUL_ENC, color="FFFFFF", tam=11, borde=caja, centro=j > 1)
        ini = fila + 1
        for c in conceptos:
            fila += 1
            _celda(ws, fila, 1, c, tam=11)
            _celda(ws, fila, 2, float(r.at[c, "ejec"]), fmt="#,##0", tam=11)
            _celda(ws, fila, 3, float(r.at[c, "ppto"]), fmt="#,##0", tam=11)
            _celda(ws, fila, 4, f"=+B{fila}-C{fila}", fmt="#,##0;[Red]-#,##0", tam=11)
            _celda(ws, fila, 5, f'=IF(C{fila}=0,"",B{fila}/C{fila})', fmt="0.0%", tam=11)
        fila += 1
        if bloque == "PERSONAL":
            fila += 1
            continue
        _celda(ws, fila, 1, "TOTALES ", True, tam=11, borde=Border(top=_fino, bottom=_fino))
        for j, col in enumerate("BCD", start=2):
            _celda(ws, fila, j, f"=+SUM({col}{ini}:{col}{fila - 1})", True, "#,##0;[Red]-#,##0", tam=11,
                   borde=Border(top=_fino, bottom=_fino))
        _celda(ws, fila, 5, f'=IF(C{fila}=0,"",B{fila}/C{fila})', True, "0.0%", tam=11, borde=Border(top=_fino, bottom=_fino))
        fila += 3
    _celda(ws, fila, 1, f"Presupuesto: ene–{MESES[corte_v - 1].lower()} {version_1}; {MESES[corte_v].lower()}–dic "
                        f"{version_2} (UPA). Diferencia = ejecución − presupuesto.", tam=9, color="5B6B82") \
        if corte_v < 12 else None
    ws.column_dimensions["A"].width = 46
    for col in "BCDE":
        ws.column_dimensions[col].width = 18

    # ---------------- PRESUPUESTO (ene..corte) y UPA (resto del año), como el archivo
    m1 = list(range(1, corte_v + 1))
    m2 = list(range(corte_v + 1, 13))
    _hoja_periodo(wb, "PRESUPUESTO", t, m1, hasta,
                  f"{nombre_empresa} · {anio} · presupuesto {version_1} ({MESES[0].lower()}–{MESES[corte_v - 1].lower()}) "
                  f"vs ejecución")
    if m2:
        _hoja_periodo(wb, "UPA", t, m2, hasta, f"{nombre_empresa} · {anio} · presupuesto {version_2} "
                      f"({MESES[corte_v].lower()}–dic) vs ejecución")

    # ---------------- OTROS (como 'OTROS PPTO'): por concepto, la ejecución por tercero y el presupuesto
    ws = wb.create_sheet("OTROS DETALLE")
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    _celda(ws, 1, 1, f"{nombre_empresa} · {anio} · otros gastos de personal: ejecución por tercero y presupuesto",
           True, tam=11, color="10284A")
    _celda(ws, 2, 2, "PRESUPUESTO", True, centro=True, fondo=CELESTE, tam=11)
    ws.merge_cells("B2:N2")
    _celda(ws, 2, 16, "EJECUCION", True, centro=True, fondo=CELESTE, tam=11)
    ws.merge_cells("P2:AD2")
    L = get_column_letter
    fila = 4
    otros = [c for _, cs in BLOQUES for c in cs if c != "Dotación y EPP's"]
    hilo = Border(bottom=_pelo)
    for c in otros:
        enc = [c.upper()] + MESES_MAY + ["TOTAL ANUAL", None, c.upper()] + MESES_MAY + ["TOTAL EJE", "DIFERENCIA"]
        for j, txt in enumerate(enc, start=1):
            if j in (15,):
                continue
            _celda(ws, fila, j, txt, True, fondo=AZUL_OSC if txt is not None else None, color="FFFFFF", tam=11,
                   centro=j > 1)
        ini = fila + 1
        e_c = ejec[(ejec["concepto_rh"] == c) & (ejec["mes"] <= hasta)]
        por_t = e_c.assign(tercero=e_c["tercero"].fillna("").replace("", "(Sin tercero)")) \
            .pivot_table(index="tercero", columns="mes", values="valor", aggfunc="sum").reindex(columns=range(1, 13))
        por_t = por_t.loc[por_t.sum(axis=1).abs().sort_values(ascending=False).index]
        for tercero, v in por_t.iterrows():
            fila += 1
            _celda(ws, fila, 1, tercero, tam=11, borde=hilo)
            _celda(ws, fila, 16, tercero, tam=11, borde=hilo)
            for m in range(1, 13):
                x = v.get(m)
                _celda(ws, fila, 16 + m, None if pd.isna(x) or m > hasta else float(x), fmt=F_CONT, tam=11, borde=hilo)
            _celda(ws, fila, 29, f"=+SUM(Q{fila}:AB{fila})", fmt=F_CONT, tam=11, borde=hilo)
        fila += 1
        _celda(ws, fila, 1, c.upper(), True, tam=11, fondo=CELESTE)
        _celda(ws, fila, 16, c.upper(), True, tam=11, fondo=CELESTE)
        for m in range(1, 13):
            _celda(ws, fila, 1 + m, float(t["ppto"].at[c, m]), True, F_CONT, CELESTE, tam=11)
            _celda(ws, fila, 16 + m, (f"=+SUM({L(16 + m)}{ini}:{L(16 + m)}{fila - 1})" if fila - 1 >= ini else 0)
                   if m <= hasta else None, True, F_CONT, CELESTE, tam=11)
        _celda(ws, fila, 14, f"=+SUM(B{fila}:M{fila})", True, F_CONT, CELESTE, tam=11)
        _celda(ws, fila, 29, f"=+SUM(Q{fila}:AB{fila})", True, F_CONT, CELESTE, tam=11)
        _celda(ws, fila, 30, f"=+AC{fila}-SUM(B{fila}:{L(1 + hasta)}{fila})", True, '#,##0_ ;[Red]\\-#,##0\\ ',
               CELESTE, tam=11)
        fila += 3
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["P"].width = 40
    ws.column_dimensions["O"].width = 3
    for col in list(range(2, 15)) + list(range(17, 31)):
        ws.column_dimensions[L(col)].width = 14
    ws.freeze_panes = "B5"

    # ---------------- Dotación: detalle (como 'Detalle Dotacion') y comparativo
    dp = leer("SELECT * FROM dotacion_ppto WHERE empresa=? AND anio=?", (empresa, int(anio)))
    de = leer("SELECT * FROM dotacion_ejec WHERE empresa=? AND anio=?", (empresa, int(anio)))
    if not dp.empty or not de.empty:
        cruce = leer("SELECT articulo, concepto, origen FROM dotacion_cruce WHERE empresa=?", (empresa,)).fillna("")
        sap = leer("SELECT anio, mes, fecha, codigo, sub_rubro, cc_nombre, documento, tercero, texto, valor, clave "
                   "FROM ejecucion WHERE empresa=? AND anio=? AND codigo='3.1'", (empresa, int(anio)))
        clasif = leer("SELECT tipo, llave, concepto FROM dotacion_contable WHERE empresa=?", (empresa,)).fillna("")
        asig = dotacion.asignar(de, cruce, dp) if not de.empty else de.assign(estado="", concepto="", categoria="")
        asig = asig.assign(origen=dotacion.ORIGEN_INV, detalle=asig["articulo"])
        directos, _ = dotacion.contables(sap, clasif, dp)
        if not directos.empty:
            asig = pd.concat([asig, directos.drop(columns=["patron", "sugerido", "sub_rubro", "tercero"])],
                             ignore_index=True)
        act = asig[(asig["estado"] != "excluido") & (asig["mes"] <= hasta)]

        d = act.assign(fecha=pd.to_datetime(act["fecha"], errors="coerce")).sort_values(["mes", "fecha"])
        det = pd.DataFrame({
            "Bodega": d["bodega"], "IdKardex": d["clave"], "Cierre": (int(anio) * 100 + d["mes"].astype(int)),
            "Grupo": d["grupo"], "Centro Costo": d["cc"],
            "Tipo Movimiento": d["origen"].map(lambda o: "Salida" if o == dotacion.ORIGEN_INV else "Registro contable"),
            "Fecha Movimiento": d["fecha"], "Codigo Articulo": d["codigo"], "Articulo": d["detalle"],
            "Cantidad Salida": d["cantidad"], "Unidad Medida": d["unidad"], "Costo Unidad Sin Iva": d["costo_unitario"],
            "Costo Total Sin Iva": d["valor_sin_iva"], "Costo Total Con Iva": d["valor"],
            "CONCEPTO PPT": d["concepto"], "CATEGORIA": d["categoria"],
            "MES": d["mes"].map(lambda m: MESES_MAY[int(m) - 1]), "ORIGEN": d["origen"]})
        ws = wb.create_sheet("Detalle Dotacion")
        ws.sheet_view.showGridLines = False
        for j, col in enumerate(det.columns, start=1):
            _celda(ws, 1, j, col, True, tam=8)
        for i, fila_d in enumerate(det.itertuples(index=False), start=2):
            for j, v in enumerate(fila_d, start=1):
                if isinstance(v, float) and pd.isna(v):
                    v = None
                elif isinstance(v, pd.Timestamp):
                    v = None if pd.isna(v) else v.to_pydatetime()
                fmt = {7: "yyyy-mm-dd hh:mm", 10: "#,##0", 12: F_CONT, 13: F_CONT, 14: F_CONT}.get(j)
                _celda(ws, i, j, v, fmt=fmt, tam=8)
        ws.auto_filter.ref = f"A1:{L(len(det.columns))}{1 + len(det)}"
        ws.freeze_panes = "A2"
        for j, w in enumerate([24, 9, 8, 15, 30, 14, 15, 13, 40, 10, 9, 14, 14, 14, 38, 14, 11, 26], start=1):
            ws.column_dimensions[L(j)].width = w

        pp_ = dp[dp["mes"] <= hasta]
        p_c = pp_.groupby(["categoria", "articulo"]).agg(ppto_cant=("cantidad", "sum"), ppto=("valor", "sum"))
        p_c.index = p_c.index.set_names(["categoria", "concepto"])
        e_c = act.groupby(["categoria", "concepto"]).agg(ejec_cant=("cantidad", "sum"), ejec=("valor", "sum"))
        cc = p_c.join(e_c, how="outer").fillna(0.0).reset_index()
        cc = cc.sort_values(["categoria", "ejec"], ascending=[True, False])
        ws = wb.create_sheet("DOTACION VS PPTO")
        ws.sheet_view.showGridLines = False
        _celda(ws, 1, 1, f"{nombre_empresa} · dotación y EPP: presupuesto vs ejecución acumulados a "
                         f"{MESES_L[hasta - 1]} {anio} (valores con IVA)", True, tam=11, color="10284A")
        enc = ["CATEGORÍA / CONCEPTO", "UNID. PPTO", "UNID. ENTREGADAS", "PRESUPUESTO", "EJECUCION", "DIFERENCIA",
               "% EJECUCIÓN"]
        for j, txt in enumerate(enc, start=1):
            _celda(ws, 3, j, txt, True, fondo=AZUL_ENC, color="FFFFFF", tam=11, centro=j > 1, borde=caja)
        fila = 4
        filas_cat = []
        for cat, g in cc.groupby("categoria", sort=False):
            fc = fila
            filas_cat.append(fc)
            _celda(ws, fila, 1, cat, True, fondo=CELESTE, tam=11)
            fila += 1
            for x in g.itertuples():
                _celda(ws, fila, 1, "   " + str(x.concepto), tam=11, borde=hilo)
                for j, v in enumerate([x.ppto_cant, x.ejec_cant, x.ppto, x.ejec], start=2):
                    _celda(ws, fila, j, float(v), fmt="#,##0" if j < 4 else F_CONT, tam=11, borde=hilo)
                _celda(ws, fila, 6, f"=+E{fila}-D{fila}", fmt="#,##0;[Red]-#,##0", tam=11, borde=hilo)
                _celda(ws, fila, 7, f'=IF(D{fila}=0,"",E{fila}/D{fila})', fmt="0.0%", tam=11, borde=hilo)
                fila += 1
            for j, col in enumerate("BCDEF", start=2):
                _celda(ws, fc, j, f"=+SUM({col}{fc + 1}:{col}{fila - 1})", True,
                       "#,##0" if j < 4 else ("#,##0;[Red]-#,##0" if col == "F" else F_CONT), CELESTE, tam=11)
            _celda(ws, fc, 7, f'=IF(D{fc}=0,"",E{fc}/D{fc})', True, "0.0%", CELESTE, tam=11)
        _celda(ws, fila, 1, "TOTALES ", True, tam=11, borde=Border(top=_fino, bottom=_fino))
        for j, col in enumerate("BCDEF", start=2):
            _celda(ws, fila, j, "=+" + "+".join(f"{col}{f}" for f in filas_cat) if filas_cat else 0, True,
                   "#,##0" if j < 4 else ("#,##0;[Red]-#,##0" if col == "F" else F_CONT), tam=11,
                   borde=Border(top=_fino, bottom=_fino))
        _celda(ws, fila, 7, f'=IF(D{fila}=0,"",E{fila}/D{fila})', True, "0.0%", tam=11,
               borde=Border(top=_fino, bottom=_fino))
        ws.column_dimensions["A"].width = 58
        for col in "BCDEFG":
            ws.column_dimensions[col].width = 17
        ws.freeze_panes = "B4"

    for hoja in wb.worksheets:                        # al imprimir: horizontal y a lo ancho de una página
        hoja.page_setup.orientation = "landscape"
        hoja.page_setup.fitToWidth, hoja.page_setup.fitToHeight = 1, 0
        hoja.sheet_properties.pageSetUpPr.fitToPage = True
    out = BytesIO()
    wb.save(out)
    return out.getvalue()
