"""
Entrada desde el sistema de gestión (Novasoftware), sin volver a pedir clave.

El panel se usa dentro del sistema: cada pantalla (Presupuesto, Flujo de caja…) es un módulo de Bogotá y se
abre en un recuadro. El sistema le entrega al navegador un PASE de dos minutos, firmado con un secreto que
comparten el sistema y el panel (variable PANEL_SSO_SECRET), con quién es la persona, su área y qué módulos
tiene concedidos. Aquí se comprueba la firma y se abre la sesión con esa área.

- Quién ve qué módulo lo decide el sistema (el pase trae la lista); el área sigue decidiendo qué gerencias
  ve cada quien, igual que antes.
- Con PANEL_SOLO_SSO=1 (producción) no hay formulario de ingreso: solo se entra desde el sistema.
- Sin PANEL_SSO_SECRET este archivo no hace nada y el panel funciona como siempre (desarrollo local).

El formato del pase es base64(JSON) + "." + HMAC-SHA256 en hexadecimal: solo biblioteca estándar.
"""
import base64
import hashlib
import hmac
import json
import os
import time

import streamlit as st

AUDIENCIA = "panel-financiero"


def _secreto() -> str:
    return os.environ.get("PANEL_SSO_SECRET", "")


def solo_sso() -> bool:
    """En producción no hay formulario de ingreso: se entra desde el sistema."""
    return os.environ.get("PANEL_SOLO_SSO", "") == "1"


def activo() -> bool:
    return bool(st.session_state.get("sso_persona"))


def verificar(pase: str, secreto: str, ahora: float | None = None) -> dict | None:
    """Los datos del pase si la firma cuadra y no ha vencido; si no, None."""
    if not pase or not secreto or pase.count(".") != 1:
        return None
    cuerpo, firma = pase.split(".")
    esperada = hmac.new(secreto.encode("utf-8"), cuerpo.encode("ascii"), hashlib.sha256).hexdigest()
    if not hmac.compare_digest(firma, esperada):
        return None
    try:
        datos = json.loads(base64.urlsafe_b64decode(cuerpo + "=" * (-len(cuerpo) % 4)).decode("utf-8"))
    except (ValueError, UnicodeDecodeError):
        return None
    ahora = time.time() if ahora is None else ahora
    if datos.get("aud") != AUDIENCIA or ahora > datos.get("exp", 0):
        return None
    return datos


@st.cache_resource
def _usados() -> dict:
    """Pases ya usados (firma -> vencimiento): un pase sirve UNA vez, aunque alguien lo copie del registro."""
    return {}


def recibir_pase(usuarios_validos) -> None:
    """Si la dirección trae ?pase= válido, abre la sesión con el área del pase. Siempre limpia la dirección."""
    pase = st.query_params.get("pase")
    if not pase:
        return
    st.query_params.clear()        # que el pase no quede en la barra ni en el historial
    datos = verificar(pase, _secreto())
    usados = _usados()
    ahora = time.time()
    for firma, vence in list(usados.items()):
        if vence < ahora:
            usados.pop(firma, None)
    firma = pase.rsplit(".", 1)[-1]
    if not datos or firma in usados or datos.get("area") not in usuarios_validos:
        st.session_state.pop("usuario", None)
        st.session_state["sso_error"] = True
        return
    usados[firma] = datos["exp"]
    st.session_state.usuario = datos["area"]
    st.session_state.sso_persona = datos.get("nombre") or datos.get("sub")
    st.session_state.sso_modulos = list(datos.get("modulos") or [])
    st.session_state.modulo = datos.get("modulo")
    st.session_state.pop("sso_error", None)


def pedir_entrar_desde_el_sistema() -> None:
    caducado = st.session_state.pop("sso_error", False)
    _, centro, _ = st.columns([1, 1.4, 1])
    with centro:
        st.html('<div class="nf-login notranslate" translate="no"><div class="nf-logo">AL</div>'
                '<h1>Panel financiero</h1><p>Grupo Área Limpia</p></div>')
        if caducado:
            st.warning("El acceso venció o ya se usó. Vuelve a abrir el módulo desde el menú del sistema.")
        else:
            st.info("El panel se abre desde el sistema de gestión: entra a Bogotá y elige el módulo en el menú.")
    st.stop()


def visibles(modulos, puede_ver, usuario, admin):
    """Con sesión del sistema, los módulos los decide el sistema (los del pase). Administración, además, solo
    para el área administradora: sus pantallas suponen que quien entra es Financiera."""
    if not activo():
        return [m for m in modulos if puede_ver(usuario, m[0])]
    concedidos = set(st.session_state.get("sso_modulos") or [])
    return [m for m in modulos if m[0] in concedidos and (m[0] != "Administración" or usuario == admin)]


def ocultar_menu() -> None:
    """Dentro del sistema, el menú es el del sistema: se esconde el de Streamlit."""
    st.html("<style>[data-testid='stSidebar'],[data-testid='stSidebarCollapsedControl'],"
            "[data-testid='stExpandSidebarButton']{display:none!important}</style>")


def quien(usuario: str) -> str:
    """Para el historial de cargas: el área y, si entró desde el sistema, la persona."""
    persona = st.session_state.get("sso_persona")
    return f"{usuario} · {persona}" if persona else usuario


# ------------------------------------------------------------------ sesión recordada
# Streamlit guarda la sesión en la memoria del servidor. Si la conexión del navegador se corta y se reabre (algo
# normal dentro del recuadro del sistema, o si el servidor tiene varias copias del panel), el servidor arranca una
# sesión nueva y se pierde el ingreso: por eso volvía a pedir la clave al hacer cambios. Para recuperarla, al
# ingresar se deja en la dirección del recuadro una marca firmada (?s=…) con el área y vencimiento de 12 horas; si
# la sesión se pierde, se reabre con esa marca sin pedir la clave otra vez.
HORAS_SESION = 12
PARAM = "s"


def _clave_sesion(usuarios) -> bytes:
    base = _secreto() or json.dumps(sorted((str(k), str(v)) for k, v in dict(usuarios or {}).items()))
    return hashlib.sha256(("sesion-panel|" + base).encode("utf-8")).digest()


def _firmar(datos: dict, usuarios) -> str:
    cuerpo = base64.urlsafe_b64encode(json.dumps(datos, separators=(",", ":")).encode("utf-8")).decode("ascii")
    cuerpo = cuerpo.rstrip("=")
    firma = hmac.new(_clave_sesion(usuarios), cuerpo.encode("ascii"), hashlib.sha256).hexdigest()[:32]
    return f"{cuerpo}.{firma}"


def recordar_sesion(usuarios) -> None:
    """Deja (o renueva) la marca de sesión en la dirección cuando falta o le quedan menos de 2 horas."""
    if "usuario" not in st.session_state:
        return
    vence = st.session_state.get("_sesion_vence", 0)
    modulo = st.session_state.get("modulo")
    if (st.query_params.get(PARAM) and vence - time.time() > 2 * 3600
            and st.session_state.get("_sesion_mod") == modulo):
        return
    vence = int(time.time() + HORAS_SESION * 3600)
    datos = {"u": st.session_state.usuario, "e": vence, "o": modulo}
    if st.session_state.get("sso_persona"):
        datos["p"] = st.session_state.sso_persona
        datos["m"] = list(st.session_state.get("sso_modulos") or [])
    st.query_params[PARAM] = _firmar(datos, usuarios)
    st.session_state["_sesion_vence"] = vence
    st.session_state["_sesion_mod"] = modulo


def recuperar_sesion(usuarios) -> bool:
    """Si la sesión se perdió pero la dirección trae una marca válida, vuelve a abrirla."""
    marca = st.query_params.get(PARAM)
    if not marca or marca.count(".") != 1:
        return False
    cuerpo, firma = marca.split(".")
    esperada = hmac.new(_clave_sesion(usuarios), cuerpo.encode("ascii"), hashlib.sha256).hexdigest()[:32]
    if not hmac.compare_digest(firma, esperada):
        st.query_params.pop(PARAM, None)
        return False
    try:
        datos = json.loads(base64.urlsafe_b64decode(cuerpo + "=" * (-len(cuerpo) % 4)).decode("utf-8"))
    except (ValueError, UnicodeDecodeError):
        return False
    if time.time() > datos.get("e", 0) or datos.get("u") not in (usuarios or {}):
        st.query_params.pop(PARAM, None)
        return False
    st.session_state.usuario = datos["u"]
    st.session_state["_sesion_vence"] = datos["e"]
    if datos.get("o"):
        st.session_state.modulo = datos["o"]
        st.session_state["_sesion_mod"] = datos["o"]
    if datos.get("p"):
        st.session_state.sso_persona = datos["p"]
        st.session_state.sso_modulos = list(datos.get("m") or [])
    return True


def olvidar_sesion() -> None:
    st.query_params.pop(PARAM, None)
    for k in ("_sesion_vence", "_sesion_mod", "sso_persona", "sso_modulos"):
        st.session_state.pop(k, None)

