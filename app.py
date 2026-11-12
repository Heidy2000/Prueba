"""
Panel financiero · Área Limpia
==============================
Módulos (menú de la izquierda):
- Presupuesto:    presupuesto vs ejecución de gastos de explotación
- Flujo de caja:  flujo de caja calculado desde los movimientos de bancos de SAP, proyectado y detalle por tercero
- Cartera:        cartera de usuarios: comparativo de un mes contra otro, edades, deterioro y castigo
- Combustible:    relación de combustible de ALDC y ALSM
-Mantenimiento:    Detallado por placa de mantenimiento y combustible
- Nómina:         presupuesto de personal vs nómina ejecutada (cargos, personas, conceptos)
- Administración: carga de archivos, asignaciones, homologaciones y respaldo (solo el administrador)
"""
import os
from pathlib import Path
from types import SimpleNamespace

import streamlit as st

import asignacion
import cartera
import db as base_datos
import estilo as E
import mod_admin
import mod_combustible
import mod_cartera
import mod_eeff
import mod_flujo
import mod_nomina
import mod_presupuesto
import nomina
import mod_mantenimiento
import operaciones as Operaciones
from ui import ADMIN, ahora

CARPETA = Path(__file__).parent


def _secreto(seccion: str, clave: str) -> str:
    try:
        return str(dict(st.secrets.get(seccion, {})).get(clave, "") or "")
    except Exception:          # sin archivo de secrets
        return ""


# configuración: las variables de entorno mandan sobre secrets (servidores, contenedores)
DB = Path(os.environ.get("PANEL_DB") or CARPETA / "presupuesto.db")
DB_URL = os.environ.get("DATABASE_URL") or _secreto("base_datos", "url")
DB_ESQUEMA = os.environ.get("DB_SCHEMA") or _secreto("base_datos", "esquema")

st.set_page_config(page_title="Panel financiero · Área Limpia", page_icon="📊", layout="wide")
st.html(E.CSS)


# ------------------------------------------------------------------ base de datos
@st.cache_resource
def conectar():
    # PostgreSQL si hay URL (Supabase, AWS RDS…); si no, archivo SQLite local
    con = base_datos.Base(ruta=DB, url=DB_URL or None, esquema=DB_ESQUEMA or None)
    con.crear_tablas("""
    CREATE TABLE IF NOT EXISTS presupuesto(anio INT, version TEXT, codigo TEXT, rubro TEXT, sub_rubro TEXT,
                                           mes INT, valor REAL);
    CREATE TABLE IF NOT EXISTS ejecucion(anio INT, mes INT, fecha TEXT, codigo TEXT, rubro TEXT,
        sub_rubro TEXT, cc TEXT, cc_nombre TEXT, cuenta TEXT, cuenta_nombre TEXT,
        documento TEXT, tercero TEXT, texto TEXT, valor REAL);
    CREATE TABLE IF NOT EXISTS homologacion(version TEXT, codigo TEXT, sub_ejec TEXT, sub_ppto TEXT);
    CREATE TABLE IF NOT EXISTS reglas_fijas(tercero TEXT, cuenta TEXT, codigo TEXT, sub_rubro TEXT, cc TEXT,
        creada TEXT);
    CREATE TABLE IF NOT EXISTS sap_columnas(orden INT, nombre TEXT);
    CREATE TABLE IF NOT EXISTS duenos(sub_rubro TEXT PRIMARY KEY, gerencia TEXT);
    CREATE TABLE IF NOT EXISTS asignacion(codigo TEXT, sub_norm TEXT, sub_rubro TEXT, gerencia TEXT,
        PRIMARY KEY(codigo, sub_norm));
    CREATE TABLE IF NOT EXISTS combustible(empresa TEXT, fuente TEXT, fecha TEXT, anio INT, mes INT, placa TEXT,
        tipo TEXT, galones REAL, valor REAL, contrato TEXT);
    CREATE TABLE IF NOT EXISTS combustible_aguas(empresa TEXT, placa TEXT);
    CREATE TABLE IF NOT EXISTS nomina_ppto(anio INT, mes INT, area TEXT, grupo TEXT, cargo TEXT, tipologia TEXT,
        personas REAL, valor REAL);
    CREATE TABLE IF NOT EXISTS nomina_ejec(anio INT, mes INT, cc TEXT, cargo TEXT, empleado_id TEXT, empleado TEXT,
        concepto TEXT, cuenta TEXT, valor REAL);
    CREATE TABLE IF NOT EXISTS nomina_ceco(cc TEXT, area TEXT);
    CREATE TABLE IF NOT EXISTS nomina_homol(fuente TEXT, cargo TEXT, cargo_comun TEXT);
    CREATE TABLE IF NOT EXISTS flujo_lineas(anio INT, orden INT, seccion TEXT, padre TEXT, concepto TEXT,
        tipo TEXT, mes INT, valor REAL);
    CREATE TABLE IF NOT EXISTS flujo_proyectado(anio INT, concepto TEXT, proyectado REAL, observacion TEXT);
    CREATE TABLE IF NOT EXISTS flujo_mov(anio INT, mes INT, fecha TEXT, concepto TEXT, tercero TEXT, texto TEXT,
        documento TEXT, cuenta TEXT, valor REAL);
    CREATE TABLE IF NOT EXISTS flujo_estructura(empresa TEXT, anio INT, orden INT, fila INT, concepto TEXT,
        seccion TEXT, padre TEXT, tipo TEXT, clase TEXT, sap INT, terminos TEXT, remun INT);
    CREATE TABLE IF NOT EXISTS flujo_ajustes(empresa TEXT, anio INT, orden INT, mes INT, valor REAL, tipo TEXT);
    CREATE TABLE IF NOT EXISTS flujo_terceros(empresa TEXT, codigo TEXT, nombre TEXT, nit TEXT);
    CREATE TABLE IF NOT EXISTS flujo_cuentas(empresa TEXT, cuenta TEXT, nombre TEXT);
    CREATE TABLE IF NOT EXISTS flujo_por_tercero(empresa TEXT, tercero TEXT, concepto TEXT);
    CREATE TABLE IF NOT EXISTS flujo_reglas(empresa TEXT, tercero TEXT, cuenta TEXT, texto TEXT, concepto TEXT,
        creada TEXT);
    CREATE TABLE IF NOT EXISTS ceco_linea(empresa TEXT, cc TEXT, gerencia TEXT, linea TEXT);
    CREATE TABLE IF NOT EXISTS distribucion(empresa TEXT, anio INT, linea TEXT, pct REAL);
    CREATE TABLE IF NOT EXISTS eeff(empresa TEXT, anio INT, mes INT, estado TEXT, idx INT, grupo TEXT,
        subgrupo TEXT, linea TEXT, valor REAL, fuente TEXT);
    CREATE TABLE IF NOT EXISTS eeff_sap(empresa TEXT, anio INT, mes INT, sociedad TEXT, posicion TEXT, texto TEXT,
        inicial REAL, final REAL);
    CREATE TABLE IF NOT EXISTS cargas(fecha TEXT, usuario TEXT, tipo TEXT, archivo TEXT, detalle TEXT);
    """)
    con.crear_tablas(cartera.ddl(con.motor))

    # columnas agregadas en versiones posteriores (bases creadas antes las reciben vacías)
    nuevas = {
        # empresa = compañía (ALDC / ALSM); linea = línea de negocio del presupuesto
        "presupuesto": {"version": "TEXT DEFAULT 'Inicial'", "empresa": "TEXT DEFAULT 'ALDC'",
                        "linea": "TEXT DEFAULT ''"},
        "flujo_lineas": {"empresa": "TEXT DEFAULT 'ALDC'"},
        "flujo_proyectado": {"empresa": "TEXT DEFAULT 'ALDC'"},
        "flujo_estructura": {"remun": "INT DEFAULT 0"},
        "flujo_ajustes": {"tipo": "TEXT DEFAULT ''"},
        "flujo_mov": {"empresa": "TEXT DEFAULT 'ALDC'",
                      **{c: "TEXT DEFAULT ''" for c in ["clave", "clase", "cuenta_cod", "cc", "cab", "ter_cod", "ref1",
                                                       "ref3", "referencia", "compensacion", "anulado", "origen",
                                                       "regla", "confianza"]}},
        "combustible": {c: "TEXT DEFAULT ''" for c in ["clave", "documento", "estacion", "hora", "km"]},
        "nomina_homol": {"cc": "TEXT DEFAULT ''"},
        "nomina_ppto": {c: "DOUBLE PRECISION" for c in nomina.COLS_CONCEPTO},
        "ejecucion": {**{c: "TEXT DEFAULT ''" for c in ["clave", "origen", "regla", "confianza", "cc_sap",
                                                       "contrapartida", "anulado", "mod_cod", "mod_sub", "mod_cc",
                                                       "raw"]}, "empresa": "TEXT DEFAULT 'ALDC'"},
    }
    for tabla, cols in nuevas.items():
        existentes = con.columnas(tabla)
        for c, tipo in cols.items():
            if c not in existentes:
                con.execute(f"ALTER TABLE {tabla} ADD COLUMN {c} {tipo}")
    con.commit()
    asignacion.migrar(con)
    return con


con = conectar()


def leer(sql, params=()):
    return con.leer(sql, params)


def registrar_carga(usuario, tipo, archivo, detalle):
    con.execute("INSERT INTO cargas(fecha, usuario, tipo, archivo, detalle) VALUES(?,?,?,?,?)",
                (ahora().strftime("%Y-%m-%d %H:%M"), usuario, tipo, archivo, detalle))
    con.commit()


# ------------------------------------------------------------------ usuarios y permisos
try:
    USUARIOS = dict(st.secrets.get("usuarios", {}))
    ACCESO = {k: list(v) for k, v in dict(st.secrets.get("acceso", {})).items()}
except Exception:              # sin archivo de secrets
    USUARIOS, ACCESO = {}, {}
MODULOS = [  # (nombre, ícono, sección del menú, función)
    ("Presupuesto", ":material/account_balance_wallet:", "Finanzas", mod_presupuesto.mostrar),
    ("Flujo de caja", ":material/payments:", "Finanzas", mod_flujo.mostrar),
    ("Estados financieros", ":material/query_stats:", "Finanzas", mod_eeff.mostrar),
    ("Cartera", ":material/request_quote:", "Finanzas", mod_cartera.mostrar),
    ("Combustible", ":material/local_gas_station:", "Operación", mod_combustible.mostrar),
    ("Mantenimiento", ":material/build:", "Operación", mod_mantenimiento.mostrar),
    ("Operaciones", ":material/build:", "Operación", Operaciones.mostrar),
    ("Nómina", ":material/badge:", "Talento humano", mod_nomina.mostrar),
    ("Administración", ":material/settings:", "Control", mod_admin.mostrar),
]

# acceso por defecto de módulos sensibles cuando secrets no dice nada (así no quedan abiertos a todos)
ACCESO_DEFECTO = {"Nómina": ["Recursos Humanos"],
                  "Estados financieros": ["Dirección", "Administrativa y Financiera"],
                  "Cartera": ["Dirección", "Administrativa y Financiera"]}
# acceso fijo: manda sobre secrets (Financiera, el administrador, siempre ve todo)
ACCESO_FIJO = {"Combustible": ["Operaciones", "Administrativa y Financiera"]}


def puede_ver(usuario: str, modulo: str) -> bool:
    if usuario == ADMIN:
        return True
    if modulo == "Administración":
        return False
    permitidos = ACCESO_FIJO.get(modulo) or ACCESO.get(modulo, ACCESO_DEFECTO.get(modulo))
    return True if permitidos is None else ("Todos" in permitidos or usuario in permitidos)


# ------------------------------------------------------------------ ingreso
if "usuario" not in st.session_state:
    _, centro, _ = st.columns([1, 1.1, 1])
    with centro:
        st.html('<div class="nf-login notranslate" translate="no"><div class="nf-logo">AL</div>'
                '<h1>Panel financiero</h1><p>Grupo Área Limpia</p></div>')
        with st.container(key="card_login"):
            if not USUARIOS:
                st.error("No hay usuarios configurados: falta la sección [usuarios] en .streamlit/secrets.toml.")
                st.stop()
            with st.form("login", border=False):
                u = st.selectbox("Usuario", list(USUARIOS.keys()))
                c = st.text_input("Clave", type="password")
                if st.form_submit_button("Ingresar", type="primary", width="stretch"):
                    if c == str(USUARIOS.get(u)):
                        st.session_state.usuario = u
                        st.rerun()
                    else:
                        st.error("Clave incorrecta")
    st.stop()

usuario = st.session_state.usuario
es_admin = usuario == ADMIN
# perfiles que ven el presupuesto de toda la compañía (además de Financiera), p. ej. Dirección
ver_todo = es_admin or usuario in ACCESO.get("Presupuesto completo", [])
visibles = [m for m in MODULOS if puede_ver(usuario, m[0])]
if st.session_state.get("modulo") not in [m[0] for m in visibles]:
    st.session_state.modulo = visibles[0][0]

# ------------------------------------------------------------------ menú lateral
with st.sidebar:
    st.html(E.marca_sidebar())
    seccion = None
    for nombre, icono, sec, _ in visibles:
        if sec != seccion:
            st.html(f'<div class="nf-sec">{sec}</div>')
            seccion = sec
        activo = st.session_state.modulo == nombre
        clave = ("navon_" if activo else "nav_") + nombre.replace(" ", "_").replace("ó", "o")
        if st.button(nombre, icon=icono, key=clave, width="stretch", type="tertiary"):
            st.session_state.modulo = nombre
            st.rerun()

    st.html('<div class="nf-sec">Sesión</div>' + E.usuario_sidebar(
        usuario, "Administrador · toda la compañía" if es_admin else ("Toda la compañía" if ver_todo else "Gerencia")))
    if st.button("Cerrar sesión", icon=":material/logout:", width="stretch", key="salir"):
        for k in ("usuario", "modulo"):
            st.session_state.pop(k, None)
        st.rerun()
    st.html('<div class="nf-foot">Dirección Financiera<br>Panel financiero v0.9</div>')

# ------------------------------------------------------------------ módulo activo
ctx = SimpleNamespace(con=con, leer=leer, usuario=usuario, es_admin=es_admin, ver_todo=ver_todo,
                      registrar_carga=registrar_carga, usuarios=USUARIOS, acceso=ACCESO)
dict((m[0], m[3]) for m in visibles)[st.session_state.modulo](ctx)
