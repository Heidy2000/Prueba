"""
Estilo visual (inspirado en NovaFlow) y componentes HTML reutilizables.
No necesitas modificar este archivo.
"""
from html import escape
from urllib.parse import quote

import pandas as pd

# ------------------------------------------------------------------ paleta
NAVY = "#10284a"
AZUL = "#1f5fc4"
AZUL_SUAVE = "#e8f0fb"
TEXTO = "#0f1f3a"
TEXTO_2 = "#5b6b82"
BORDE = "#e3e8f0"
GRIS_BARRA = "#e9eef5"
VERDE, VERDE_SUAVE = "#2e9e5b", "#e8f6ee"
AMBAR, AMBAR_SUAVE = "#c47012", "#fdf1e3"
ROJO, ROJO_SUAVE = "#d23b3b", "#fdecec"
FUENTE = "Plus Jakarta Sans, sans-serif"

CSS = f"""
<style>
/* ---------- base ---------- */
[data-testid="stHeader"] {{ background: transparent; }}
[data-testid="stMainBlockContainer"], .block-container {{ padding-top: 1.2rem; max-width: 1400px; }}
h1, h2, h3 {{ letter-spacing: -0.01em; }}

/* ---------- barra lateral ---------- */
[data-testid="stSidebar"] {{ border-right: none; }}
[data-testid="stSidebarUserContent"] {{ padding-top: .5rem; }}
.nf-brand {{ display:flex; gap:.7rem; align-items:center; padding:.2rem 0 1rem 0;
            border-bottom:1px solid #27466f; margin-bottom:.6rem; }}
.nf-logo {{ width:40px; height:40px; border-radius:10px; background:#2a6fdb; color:#fff;
           display:flex; align-items:center; justify-content:center; font-weight:800; font-size:.95rem; }}
.nf-brand b {{ color:#fff; font-size:1rem; display:block; line-height:1.15; }}
.nf-brand span {{ color:#b8c6db; font-size:.72rem; }}
.nf-sec {{ color:#9fb1cc; font-size:.7rem; font-weight:700; letter-spacing:.08em;
          text-transform:uppercase; margin:1rem 0 .35rem 0; }}
.nf-user {{ display:flex; gap:.6rem; align-items:center; background:#1b3a64; border-left:3px solid #7cc242;
           border-radius:10px; padding:.6rem .7rem; }}
.nf-avatar {{ width:34px; height:34px; border-radius:50%; background:{AZUL}; color:#fff; font-weight:700;
             font-size:.8rem; display:flex; align-items:center; justify-content:center; flex-shrink:0; }}
.nf-user b {{ color:#fff; font-size:.85rem; display:block; }}
.nf-user span {{ color:#b8c6db; font-size:.72rem; }}
.nf-foot {{ color:#9fb1cc; font-size:.7rem; line-height:1.4; margin-top:1.5rem;
           border-top:1px solid #27466f; padding-top:.7rem; }}

/* ---------- menú de módulos ---------- */
[data-testid="stSidebar"] [class*="st-key-nav"] button {{ justify-content:flex-start; padding:.55rem .75rem;
    border-radius:10px; border-left:3px solid transparent; color:#dbe5f3; font-weight:600; }}
[data-testid="stSidebar"] [class*="st-key-nav"] button p {{ font-size:.9rem; font-weight:600; }}
[data-testid="stSidebar"] [class*="st-key-nav"] button > div {{ justify-content:flex-start; width:100%; gap:.2rem; }}
[data-testid="stSidebar"] [class*="st-key-nav"] button {{ width:100%; }}
[data-testid="stSidebar"] [class*="st-key-nav_"] button:hover {{ background:#1b3a64; color:#fff; }}
[data-testid="stSidebar"] [class*="st-key-navon_"] button {{ background:#1b3a64; border-left-color:#7cc242;
    color:#fff; }}
[data-testid="stSidebar"] [class*="st-key-nav"] {{ margin-bottom:-.6rem; }}

/* ---------- encabezado ---------- */
.nf-top {{ display:flex; justify-content:space-between; align-items:center; gap:1rem; flex-wrap:wrap;
          margin-bottom:1rem; }}
.nf-top h1 {{ font-size:1.55rem; font-weight:800; margin:0; padding:0; color:{TEXTO}; }}
.nf-top p {{ color:{TEXTO_2}; font-size:.85rem; margin:.15rem 0 0 0; }}
.nf-chips {{ display:flex; gap:.5rem; flex-wrap:wrap; }}
.nf-chip {{ border:1px solid {BORDE}; background:#fff; border-radius:999px; padding:.4rem .85rem;
           font-size:.78rem; color:{TEXTO}; font-weight:600; white-space:nowrap; }}
.nf-chip.ok {{ border-color:#bfe3cc; color:{VERDE}; background:#f3fbf6; }}
.nf-chip.filtro {{ border-color:#b9cff0; color:{AZUL}; background:{AZUL_SUAVE}; }}
.nf-chip.warn {{ border-color:#f0c89a; color:{AMBAR}; background:{AMBAR_SUAVE}; }}

/* ---------- pestañas tipo píldora ---------- */
[role="tablist"] {{ gap:.35rem; background:#fff; border:1px solid {BORDE} !important; border-radius:14px;
    padding:.4rem; box-shadow:0 1px 2px rgba(16,40,74,.04); }}
[role="tablist"] [role="tab"] {{ border-radius:999px; padding:.4rem 1.05rem; height:auto; border:1px solid transparent;
    margin:0; }}
[role="tablist"] [role="tab"] p {{ font-size:.85rem; color:{TEXTO_2}; font-weight:600; }}
[role="tablist"] [role="tab"][aria-selected="true"] {{ background:{AZUL_SUAVE}; border-color:#b9cff0; }}
[role="tablist"] [role="tab"][aria-selected="true"] p {{ color:{AZUL}; }}
[role="tablist"] .react-aria-SelectionIndicator, [data-baseweb="tab-highlight"], [data-baseweb="tab-border"] {{ display:none !important; }}
[data-testid="stTabs"] > div:first-child, .stTabs > div:first-child {{ border-bottom:none !important; }}

/* ---------- tarjetas ---------- */
[class*="st-key-card"] {{ background:#fff; border:1px solid {BORDE} !important; border-radius:18px;
    box-shadow:0 1px 2px rgba(16,40,74,.04), 0 4px 16px rgba(16,40,74,.04); padding:1.1rem 1.2rem; }}
.nf-card-t {{ font-size:1.1rem; font-weight:800; color:{TEXTO}; margin:0; }}
.nf-card-s {{ font-size:.78rem; color:{TEXTO_2}; margin:.15rem 0 .6rem 0; }}

.nf-kpis {{ display:grid; grid-template-columns:repeat(4, minmax(0,1fr)); gap:1rem; margin-bottom:1rem; }}
@media (max-width: 900px) {{ .nf-kpis {{ grid-template-columns:repeat(2, minmax(0,1fr)); }} }}
.nf-kpi {{ background:#fff; border:1px solid {BORDE}; border-radius:18px; padding:1rem 1.1rem;
          box-shadow:0 1px 2px rgba(16,40,74,.04); }}
/* ---------- tarjetas de la vista ejecutiva (nueva): borde de color, variación y tendencia ---------- */
.nf-kpis.nf-kn {{ grid-template-columns:repeat(var(--n,4), minmax(0,1fr)); }}
@media (max-width: 1100px) {{ .nf-kpis.nf-kn {{ grid-template-columns:repeat(min(var(--n,4),3), minmax(0,1fr)); }} }}
@media (max-width: 900px) {{ .nf-kpis.nf-kn {{ grid-template-columns:repeat(2, minmax(0,1fr)); }} }}
.nf-kn .nf-kpi {{ border-left:4px solid var(--ac,{AZUL}); min-width:0; }}
.nf-kpis[style*="--n:6"] .nf-kpi-v, .nf-kpis[style*="--n:5"] .nf-kpi-v {{ font-size:1.3rem; }}
.nf-kpis[style*="--n:6"] .nf-kpi {{ padding:.9rem .85rem; }}
.nf-kpi-spark {{ margin-top:.45rem; display:block; width:100%; height:38px; }}
/* ---------- variación en las tarjetas ---------- */
.nf-delta {{ display:inline-flex; align-items:center; gap:.25rem; border-radius:999px; padding:.14rem .55rem;
            font-size:.72rem; font-weight:700; white-space:nowrap; margin-top:.35rem; }}
.nf-delta.ok {{ background:{VERDE_SUAVE}; color:{VERDE}; }}
.nf-delta.bad {{ background:{ROJO_SUAVE}; color:{ROJO}; }}
.nf-delta.warn {{ background:{AMBAR_SUAVE}; color:{AMBAR}; }}
.nf-delta.na {{ background:#f2f4f7; color:{TEXTO_2}; }}

/* ---------- lectura (conclusión del informe) ---------- */
.nf-lectura {{ display:flex; gap:.75rem; align-items:flex-start; background:#fff; border:1px solid {BORDE};
    border-left:4px solid {AZUL}; border-radius:14px; padding:.85rem 1.1rem; margin:.1rem 0 1rem 0;
    font-size:.88rem; color:{TEXTO}; line-height:1.5; box-shadow:0 1px 2px rgba(16,40,74,.04); }}
.nf-lectura .nf-ico {{ width:32px; height:32px; }}
.nf-lectura b.t {{ color:{AZUL}; }}
.nf-lectura.ok {{ border-left-color:{VERDE}; }}  .nf-lectura.ok b.t {{ color:{VERDE}; }}
.nf-lectura.warn {{ border-left-color:#e8892c; }} .nf-lectura.warn b.t {{ color:{AMBAR}; }}
.nf-lectura.bad {{ border-left-color:{ROJO}; }}  .nf-lectura.bad b.t {{ color:{ROJO}; }}

/* ---------- barras divergentes (variación por línea / gerencia) ---------- */
.nf-dvt {{ font-size:.82rem; color:{TEXTO}; border:1px solid {BORDE}; border-radius:12px; overflow:hidden; }}
.nf-dvr {{ display:grid; align-items:center; gap:.4rem; padding:.55rem 1.2rem .55rem .9rem; border-bottom:1px solid #eef1f6; }}
.nf-dvr:last-child {{ border-bottom:none; }}
.nf-dvr.nf-head {{ padding:.55rem .9rem; }}
.nf-dvr.nf-head > div:not(:first-child) {{ text-align:center; }}
.nf-dvr.total {{ background:#f2f6fc; font-weight:800; }}
.nf-dvr.foco {{ background:#fff7f7; box-shadow:inset 3px 0 0 {ROJO}; }}
.nf-dvr .nm b {{ display:block; font-weight:700; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }}
.nf-dvr .nm small {{ color:{TEXTO_2}; font-size:.72rem; }}
.nf-dv {{ display:grid; grid-template-columns:1fr 1fr; align-items:center; min-height:34px; }}
.nf-dv > div {{ display:flex; align-items:center; gap:.4rem; min-width:0; height:100%; }}
.nf-dv > div:first-child {{ justify-content:flex-end; border-right:2px solid #c9d4e4; padding-right:0; }}
.nf-dv i {{ display:block; height:16px; border-radius:4px; flex-shrink:0; }}
.nf-dv > div:first-child i {{ border-radius:4px 0 0 4px; }}
.nf-dv > div:last-child i {{ border-radius:0 4px 4px 0; }}
.nf-dv span {{ display:flex; flex-direction:column; line-height:1.1; font-variant-numeric:tabular-nums; }}
.nf-dv > div:first-child span {{ align-items:flex-end; }}
.nf-dv span b {{ font-size:.82rem; }}
.nf-dv span small {{ font-size:.68rem; font-weight:600; opacity:.85; }}

/* ---------- avisos ---------- */
.nf-avisos {{ display:grid; grid-template-columns:repeat(auto-fit, minmax(125px,1fr)); gap:.6rem; }}
.nf-aviso {{ border:1px solid #f0c89a; background:{AMBAR_SUAVE}; border-radius:12px; padding:.6rem .75rem;
    font-size:.76rem; color:#7a4a12; line-height:1.3; }}
.nf-aviso b {{ display:block; font-size:1.15rem; color:{AMBAR}; font-weight:800; }}
.nf-aviso.ok {{ border-color:#bfe3cc; background:{VERDE_SUAVE}; color:#1f6b40; }} .nf-aviso.ok b {{ color:{VERDE}; }}
.nf-aviso.bad {{ border-color:#f3c1c1; background:{ROJO_SUAVE}; color:#8b1e1e; }} .nf-aviso.bad b {{ color:{ROJO}; }}
.nf-aviso.na {{ border-color:{BORDE}; background:#f7f9fc; color:{TEXTO_2}; }} .nf-aviso.na b {{ color:{TEXTO}; }}

/* ---------- informe de resumen (texto corto al final de un módulo) ---------- */
.nf-inf {{ background:#fff; border:1px solid {BORDE}; border-top:4px solid {AZUL}; border-radius:16px;
    padding:1.1rem 1.3rem 1rem; margin:1.2rem 0 .6rem; color:{TEXTO}; font-size:.88rem; line-height:1.55;
    box-shadow:0 1px 2px rgba(16,40,74,.04); }}
.nf-inf-h {{ display:flex; align-items:center; gap:.7rem; margin-bottom:.4rem; }}
.nf-inf-h b {{ font-size:1.05rem; color:{NAVY}; display:block; }}
.nf-inf-h span {{ font-size:.78rem; color:{TEXTO_2}; }}
.nf-inf-t {{ font-weight:700; color:{AZUL}; margin:.75rem 0 .2rem; font-size:.84rem; text-transform:uppercase;
    letter-spacing:.02em; }}
.nf-inf ul {{ margin:0 0 0 1.1rem; padding:0; }}
.nf-inf li {{ margin:.15rem 0; }}
.nf-inf .up {{ color:{ROJO}; font-weight:700; }}
.nf-inf .dn {{ color:{VERDE}; font-weight:700; }}
/* ---------- mini barra de variación dentro de las tablas ---------- */
.nf-mb {{ display:flex !important; align-items:center; justify-content:flex-end; gap:.45rem; }}
.nf-mb > i {{ position:relative; width:56px; height:12px; flex-shrink:0; display:block; }}
.nf-mb > i::after {{ content:""; position:absolute; left:50%; top:-2px; bottom:-2px; width:1px; background:#c9d4e4; }}
.nf-mb > i > u {{ position:absolute; top:1px; height:10px; border-radius:2px; text-decoration:none; }}
.nf-kpi-h {{ display:flex; justify-content:space-between; align-items:flex-start; gap:.5rem; }}
.nf-kpi-l {{ font-size:.82rem; color:{TEXTO}; font-weight:600; }}
.nf-ico {{ width:38px; height:38px; border-radius:10px; display:flex; align-items:center; justify-content:center;
          flex-shrink:0; }}
.nf-kpi-v {{ font-size:1.55rem; font-weight:800; color:{TEXTO}; margin-top:.35rem; letter-spacing:-.01em; }}
.nf-kpi-s {{ font-size:.75rem; color:{TEXTO_2}; margin-top:.1rem; }}

/* ---------- etiquetas de estado ---------- */
.nf-pill {{ display:inline-flex; align-items:center; gap:.35rem; border-radius:999px; padding:.18rem .6rem;
           font-size:.72rem; font-weight:600; border:1px solid; white-space:nowrap; }}
.nf-pill i {{ width:6px; height:6px; border-radius:50%; display:inline-block; }}
.nf-ok   {{ background:{VERDE_SUAVE}; border-color:#bfe3cc; color:{VERDE}; }}  .nf-ok i {{ background:{VERDE}; }}
.nf-warn {{ background:{AMBAR_SUAVE}; border-color:#f0c89a; color:{AMBAR}; }}  .nf-warn i {{ background:#e8892c; }}
.nf-bad  {{ background:{ROJO_SUAVE}; border-color:#f3c1c1; color:{ROJO}; }}   .nf-bad i {{ background:{ROJO}; }}
.nf-na   {{ background:#f2f4f7; border-color:{BORDE}; color:{TEXTO_2}; }}      .nf-na i {{ background:#9aa7b8; }}

/* ---------- lista de avance (panel derecho) ---------- */
.nf-prog {{ margin:.1rem 0 .95rem 0; }}
.nf-prog-h {{ display:flex; justify-content:space-between; font-size:.82rem; color:{TEXTO}; margin-bottom:.35rem; }}
.nf-prog-h b {{ font-weight:700; }}
.nf-bar {{ height:6px; background:{GRIS_BARRA}; border-radius:99px; overflow:hidden; }}
.nf-bar > div {{ height:100%; border-radius:99px; }}

/* ---------- tabla desplegable ---------- */
.nf-t {{ font-size:.82rem; color:{TEXTO}; }}
.nf-row {{ display:grid; grid-template-columns: minmax(220px,2.6fr) repeat(4, minmax(90px,1fr)) minmax(130px,1.3fr) minmax(120px,1fr);
          align-items:center; gap:.6rem; padding:.7rem .9rem; border-bottom:1px solid #eef1f6; }}
.nf-row > div:not(:first-child):not(:nth-last-child(-n+2)) {{ text-align:right; font-variant-numeric: tabular-nums; }}
.nf-head {{ background:#f7f9fc; color:{TEXTO_2}; font-size:.74rem; font-weight:600; border-radius:10px 10px 0 0;
           border-bottom:1px solid {BORDE}; }}
.nf-t details > summary {{ list-style:none; cursor:pointer; }}
.nf-t details > summary::-webkit-details-marker {{ display:none; }}
.nf-t details > summary:hover {{ background:#f7faff; }}
.nf-name {{ display:flex; align-items:center; gap:.5rem; min-width:0; }}
.nf-name span {{ overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }}
.nf-chev {{ width:7px; height:7px; flex-shrink:0; border-right:2px solid {TEXTO_2}; border-bottom:2px solid {TEXTO_2};
    transform:rotate(-45deg); transition:transform .15s; margin:0 .35rem 0 .1rem; }}
.nf-t details[open] > summary .nf-chev {{ transform:rotate(45deg); border-color:{AZUL}; }}
.nf-l0 > summary {{ font-weight:700; }}
.nf-l0 > summary .nf-name small {{ color:{TEXTO_2}; font-weight:500; margin-left:.3rem; }}
.nf-l1 > summary {{ font-weight:600; background:#fbfcfe; }}
.nf-l1 > summary .nf-name {{ padding-left:1.4rem; }}
.nf-sub {{ background:#fff; }}
.nf-sub .nf-name {{ padding-left:1.4rem; color:#33445e; font-weight:500; }}
.nf-l1 .nf-sub .nf-name {{ padding-left:2.8rem; }}
.nf-sub .nf-name small {{ color:{TEXTO_2}; font-weight:500; margin-left:.3rem; white-space:nowrap; }}
.nf-sp {{ width:7px; flex-shrink:0; margin:0 .35rem 0 .1rem; }}
.nf-l2[open] > summary {{ background:#f7faff; }}
.nf-ter {{ background:#fbfcfe; color:#44546d; font-size:.78rem; padding-top:.5rem; padding-bottom:.5rem; }}
.nf-ter .nf-name {{ padding-left:3.2rem; }}
.nf-l1 .nf-ter .nf-name {{ padding-left:4.6rem; }}
.nf-ter .nf-name span::before {{ content:"↳ "; color:#9aa7b8; }}
.nf-docs {{ color:{TEXTO_2}; font-size:.74rem; }}
.nf-neg {{ color:{ROJO}; }}
.nf-pct {{ display:flex; align-items:center; gap:.5rem; }}
.nf-pct .nf-bar {{ flex:1; }}
.nf-pct span {{ width:48px; text-align:right; font-variant-numeric: tabular-nums; }}
.nf-dev {{ display:grid; grid-template-columns:minmax(0,2.2fr) repeat(3, minmax(0,1fr)) minmax(0,1.1fr); gap:.6rem;
    align-items:center; padding:.6rem .9rem; border-bottom:1px solid #eef1f6; }}
.nf-dev > div:nth-child(2), .nf-dev > div:nth-child(3), .nf-dev > div:nth-child(4) {{ text-align:right; font-variant-numeric:tabular-nums; }}
.nf-dev:last-child {{ border-bottom:none; }}
.nf-total {{ font-weight:800; background:#f7f9fc; border-radius:0 0 10px 10px; border-bottom:none; }}
.nf-scroll {{ overflow-x:auto; border:1px solid {BORDE}; border-radius:12px; }}
.nf-scroll .nf-t {{ min-width:980px; }}

/* ---------- tabla por meses (flujo) ---------- */
.nf-tm .nf-row > div:not(:first-child) {{ text-align:right; font-variant-numeric:tabular-nums; }}
.nf-tm .nf-row {{ padding:.55rem .9rem; }}
.nf-tm details > summary {{ list-style:none; cursor:pointer; }}
.nf-tm details > summary::-webkit-details-marker {{ display:none; }}
.nf-tm details[open] > summary .nf-chev {{ transform:rotate(45deg); border-color:{AZUL}; }}
.nf-tm .nf-name small {{ color:{TEXTO_2}; font-weight:500; margin-left:.35rem; white-space:nowrap; }}
.nf-m-saldo {{ background:#f2f6fc; font-weight:800; }}
.nf-m-seccion {{ font-weight:800; }}
summary.nf-m-seccion {{ background:#fbfcfe; }}
.nf-m-grupo {{ font-weight:700; }}
.nf-m-linea {{ font-weight:500; color:#33445e; }}
.nf-m-tercero {{ font-size:.77rem; color:#52627a; background:#fbfcfe; padding-top:.4rem !important; padding-bottom:.4rem !important; }}
.nf-m-tercero .nf-name span[title]::before {{ content:"↳ "; color:#9aa7b8; }}
.nf-pos {{ color:{VERDE}; }}

/* ---------- barra de filtros de informe ---------- */
.st-key-card_filtros {{ padding:.8rem 1.1rem .4rem 1.1rem !important; margin-bottom:.2rem; }}
.st-key-card_filtros label p {{ font-size:.74rem !important; color:{TEXTO_2}; font-weight:600; }}
.nf-ftit {{ display:flex; align-items:center; gap:.45rem; font-size:.78rem; font-weight:700; color:{TEXTO};
    text-transform:uppercase; letter-spacing:.06em; margin-bottom:.1rem; }}

/* ---------- login ---------- */
.nf-login {{ text-align:center; margin:3rem 0 1rem 0; }}
.nf-login .nf-logo {{ margin:0 auto .8rem auto; width:54px; height:54px; font-size:1.2rem; border-radius:14px; }}
.nf-login h1 {{ font-size:1.5rem; font-weight:800; margin:0; }}
.nf-login p {{ color:{TEXTO_2}; font-size:.88rem; }}
</style>
"""

# ------------------------------------------------------------------ íconos (trazos tipo Lucide)
_ICO = {
    "wallet": '<path d="M19 7V5a2 2 0 0 0-2-2H5a2 2 0 0 0 0 4h14a2 2 0 0 1 2 2v3m0 4v3a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5"/><path d="M21 12h-4a2 2 0 0 0 0 4h4z"/>',
    "calendar": '<rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/>',
    "receipt": '<path d="M4 2v20l3-2 3 2 3-2 3 2 3-2V2l-3 2-3-2-3 2-3-2z"/><path d="M8 8h8M8 12h8M8 16h5"/>',
    "alert": '<path d="M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/><path d="M12 9v4M12 17h.01"/>',
    "check": '<circle cx="12" cy="12" r="10"/><path d="m9 12 2 2 4-4"/>',
    "bulb": '<path d="M9 18h6M10 22h4"/><path d="M12 2a7 7 0 0 0-4 12.7c.6.5 1 1.3 1 2.1V17h6v-.2c0-.8.4-1.6 1-2.1A7 7 0 0 0 12 2z"/>',
    "users": '<path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M22 21v-2a4 4 0 0 0-3-3.9M16 3.1a4 4 0 0 1 0 7.8"/>',
    "trend": '<path d="m22 7-8.5 8.5-5-5L2 17"/><path d="M16 7h6v6"/>',
    "piggy": '<path d="M19 7c-1-2-3.5-3-7-3C6 4 3 7 3 11c0 2.4 1.2 4.3 3 5.5V20h3v-2h4v2h3v-3.2c1.2-.7 2.2-1.7 2.7-2.8H21v-4h-1.6c-.1-.4-.3-.7-.4-1z"/><path d="M16 11h.01"/>',
}
_CHEV = '<span class="nf-chev"></span>'


def icono(nombre: str, color: str, fondo: str) -> str:
    svg = (f'<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" '
           f'stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">{_ICO[nombre]}</svg>')
    return (f'<div class="nf-ico" style="background:{fondo}">'
            f'<img alt="" width="20" height="20" src="data:image/svg+xml;utf8,{quote(svg)}"></div>')


# ------------------------------------------------------------------ formatos
def millones(v: float) -> str:
    if pd.isna(v):
        return "—"
    s = f"{abs(v) / 1e6:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("-" if v < 0 else "") + "$" + s + " M"


def pesos(v: float) -> str:
    return "$" + f"{v:,.0f}".replace(",", ".")


def pct_txt(p) -> str:
    return "—" if pd.isna(p) else f"{p * 100:.1f}%".replace(".", ",")


def es_ingreso(pa: float, pc: float = 0.0, ec: float = 0.0) -> bool:
    """En SAP los ingresos van en negativo. Se decide por el presupuesto anual; si no hay, por el periodo o la ejecución."""
    for v in (pa, pc, ec):
        if v and not pd.isna(v):
            return v < 0
    return False


def metricas(pa: float, pc: float, ec: float, ea: float | None = None) -> dict:
    """
    Recibe valores en signo SAP (presupuesto anual, presupuesto del periodo, ejecutado del periodo, ejecutado acumulado)
    y devuelve los valores para mostrar en positivo, la variación (ejecutado − presupuesto, con los valores en
    positivo), 'fav' (la variación vista como a favor + / en contra −, para el color) y el estado.
    """
    ing = es_ingreso(pa, pc, ec)
    sg = -1 if ing else 1
    PA, PC, EC = sg * pa, sg * pc, sg * ec
    EA = sg * (ea if ea is not None else ec)
    var = EC - PC                              # negativo: se ejecutó menos de lo presupuestado
    fav = var if ing else -var                 # a favor (+) / en contra (−): más ingreso o menos gasto
    pct = EC / PC if PC else float("nan")
    clase, texto = clase_estado(pct, EC, ing)
    return dict(ingreso=ing, signo=sg, pa=PA, pc=PC, ec=EC, var=var, fav=fav, pct=pct,
                disp=PA - EA, clase=clase, estado=texto)


def clase_estado(p, ejecutado: float = 0, ingreso: bool = False) -> tuple[str, str]:
    """Clase CSS y texto del estado. En gastos, pasar del 100 % es malo; en ingresos, no llegar es malo."""
    if ingreso:
        if pd.isna(p):
            return ("nf-ok", "Sin presupuesto") if ejecutado > 0 else ("nf-na", "Sin movimiento")
        if p >= 0.995:
            return "nf-ok", "Cumplido"
        if p >= 0.9:
            return "nf-warn", "Alerta"
        return "nf-bad", "Por debajo"
    if pd.isna(p):
        return ("nf-bad", "Sin presupuesto") if ejecutado > 0 else ("nf-na", "Sin movimiento")
    if p > 1.005:
        return "nf-bad", "Sobreejecutado"
    if p > 0.9:
        return "nf-warn", "Alerta"
    return "nf-ok", "En rango"


def pill(p, ejecutado: float = 0, ingreso: bool = False) -> str:
    c, t = clase_estado(p, ejecutado, ingreso)
    return f'<span class="nf-pill {c}"><i></i>{t}</span>'


def color_barra(p, ejecutado: float = 0, ingreso: bool = False) -> str:
    return {"nf-bad": ROJO, "nf-warn": "#e8892c", "nf-ok": (VERDE if ingreso else AZUL),
            "nf-na": "#9aa7b8"}[clase_estado(p, ejecutado, ingreso)[0]]


def barra(p, ejecutado: float = 0, ingreso: bool = False) -> str:
    if pd.isna(p):
        ancho = 100 if ejecutado > 0 else 0
    else:
        ancho = max(0, min(p, 1)) * 100
    return (f'<div class="nf-bar"><div style="width:{ancho:.1f}%;background:{color_barra(p, ejecutado, ingreso)}">'
            f'</div></div>')


# ------------------------------------------------------------------ bloques
def marca_sidebar() -> str:
    return ('<div class="nf-brand notranslate" translate="no"><div class="nf-logo">AL</div>'
            '<div><b>Grupo Área Limpia</b></div></div>')


def usuario_sidebar(nombre: str, rol: str) -> str:
    ini = "".join(w[0] for w in nombre.split()[:2]).upper()
    return (f'<div class="nf-user notranslate" translate="no"><div class="nf-avatar">{escape(ini)}</div><div><b>{escape(nombre)}</b>'
            f'<span>{escape(rol)}</span></div></div>')


def encabezado(titulo: str, subtitulo: str, chips: list[tuple[str, str]]) -> str:
    cs = "".join(f'<span class="nf-chip {k}">{escape(t)}</span>' for t, k in chips)
    return (f'<div class="nf-top"><div><h1>{escape(titulo)}</h1><p>{escape(subtitulo)}</p></div>'
            f'<div class="nf-chips">{cs}</div></div>')




def delta(texto: str, tono: str = "na", flecha: str | None = None) -> str:
    """Etiqueta de variación (▲ +8,1 % vs ppto). tono: ok (a favor) / bad (en contra) / warn / na."""
    f = {"up": "▲ ", "down": "▼ ", None: ""}[flecha]
    return f'<span class="nf-delta {tono}">{f}{texto}</span>'


def delta_valor(var: float, base: float, *, a_favor_si_sube: bool = True, sufijo: str = "vs ppto",
                umbral: float = 0.005) -> str:
    """Etiqueta de variación porcentual de un valor contra su base, coloreada a favor / en contra."""
    if base is None or pd.isna(base) or abs(base) < 1e-9 or var is None or pd.isna(var):
        return delta(f"sin base {sufijo.replace('vs ', '')}".strip(), "na")
    p = var / abs(base)
    if abs(p) < umbral:
        return delta(f"{pct_txt(p)} {sufijo}", "na")
    tono = "ok" if (p > 0) == a_favor_si_sube else "bad"
    return delta(f"{'+' if p > 0 else ''}{pct_txt(p)} {sufijo}", tono, "up" if p > 0 else "down")


def sparkline(valores, color: str = AZUL, marca: int | None = None) -> str:
    """Mini tendencia (SVG) para las tarjetas. `marca`: índice del punto resaltado (por defecto el último)."""
    v = [None if x is None or pd.isna(x) else float(x) for x in valores]
    pts = [(i, x) for i, x in enumerate(v) if x is not None]
    if len(pts) < 2:
        return ""
    w, h, pad = 200.0, 38.0, 3.0
    lo, hi = min(x for _, x in pts), max(x for _, x in pts)
    rng = (hi - lo) or abs(hi) or 1.0
    n = max(len(v) - 1, 1)
    xy = [(pad + i * (w - 2 * pad) / n, h - pad - (x - lo) / rng * (h - 2 * pad)) for i, x in pts]
    linea = " ".join(f"{a:.1f},{b:.1f}" for a, b in xy)
    area = f"{xy[0][0]:.1f},{h} " + linea + f" {xy[-1][0]:.1f},{h}"
    mi = marca if marca is not None else pts[-1][0]
    punto = next(((a, b) for (i, _), (a, b) in zip(pts, xy) if i == mi), xy[-1])
    svg = (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {w:.0f} {h:.0f}" preserveAspectRatio="none">'
           f'<polygon points="{area}" fill="{color}" fill-opacity=".12"/>'
           f'<polyline points="{linea}" fill="none" stroke="{color}" stroke-width="2" '
           f'vector-effect="non-scaling-stroke" stroke-linejoin="round"/>'
           f'<circle cx="{punto[0]:.1f}" cy="{punto[1]:.1f}" r="3" fill="{color}"/></svg>')
    return f'<img class="nf-kpi-spark" alt="" src="data:image/svg+xml;utf8,{quote(svg)}">'


def kpis(items: list[dict], columnas: int | None = None, nuevo: bool = True) -> str:
    """items: {label, valor, sub, icono, color, fondo} y opcionales:
    delta (html de `delta(...)`), acento (color del borde izquierdo; por defecto `color`), spark (lista de valores).
    nuevo=True: estilo de la vista ejecutiva (borde de color y número de columnas libre)."""
    if not nuevo:
        html = ""
        for k in items:
            html += (f'<div class="nf-kpi"><div class="nf-kpi-h"><div class="nf-kpi-l">{escape(k["label"])}</div>'
                     f'{icono(k["icono"], k["color"], k["fondo"])}</div>'
                     f'<div class="nf-kpi-v">{k["valor"]}</div><div class="nf-kpi-s">{k["sub"]}</div></div>')
        return f'<div class="nf-kpis">{html}</div>'

    html = ""
    for k in items:
        ac = k.get("acento") or k["color"]
        spark = sparkline(k["spark"], k.get("spark_color") or ac) if k.get("spark") is not None else ""
        html += (f'<div class="nf-kpi" style="--ac:{ac}"><div class="nf-kpi-h"><div class="nf-kpi-l">'
                 f'{escape(k["label"])}</div>{icono(k["icono"], k["color"], k["fondo"])}</div>'
                 f'<div class="nf-kpi-v">{k["valor"]}</div>' + (k.get("delta") or "")
                 + (f'<div class="nf-kpi-s">{k["sub"]}</div>' if k.get("sub") else "") + f"{spark}</div>")
    n = columnas or min(len(items), 6) or 4
    return f'<div class="nf-kpis nf-kn notranslate" translate="no" style="--n:{n}">{html}</div>'


def lectura(titulo: str, texto: str, tono: str = "") -> str:
    """Conclusión del informe en una frase: título en negrita + explicación (puede traer <b>)."""
    color, fondo = {"ok": (VERDE, VERDE_SUAVE), "warn": (AMBAR, AMBAR_SUAVE),
                    "bad": (ROJO, ROJO_SUAVE)}.get(tono, (AZUL, AZUL_SUAVE))
    return (f'<div class="nf-lectura {tono} notranslate" translate="no">{icono("bulb", color, fondo)}'
            f'<div><b class="t">{escape(titulo)}</b> {texto}</div></div>')


def barras_divergentes(filas: list[dict], columnas: list[str], nombre_col: str = "", formato=None) -> str:
    """
    Tabla de variaciones con barras a lado y lado de un eje (como 'variación por línea de negocio').
    filas: dict(nombre, sub='', clase=''|'total'|'foco', celdas=[dict(v, pct=None, fav=True/False/None)])
    fav = la variación favorece el resultado (verde) o no (rojo); None = gris.
    """
    formato = formato or (lambda x: millones(x).replace(" M", ""))
    n = len(columnas)
    plantilla = f"minmax(140px,1.25fr) repeat({n}, minmax(150px,1fr))"
    base_max = [f for f in filas if f.get("clase") != "total"] or filas      # el total no fija la escala
    maxs = [max([abs(f["celdas"][i]["v"] or 0) for f in base_max] + [1e-9]) for i in range(n)]
    cab = (f'<div class="nf-dvr nf-head" style="grid-template-columns:{plantilla}"><div>{escape(nombre_col)}</div>'
           + "".join(f"<div>{escape(c)}</div>" for c in columnas) + "</div>")
    cuerpo = ""
    for f in filas:
        celdas = ""
        for i, c in enumerate(f["celdas"]):
            v = c.get("v") or 0.0
            fav = c.get("fav")
            col = "#9aa7b8" if fav is None or abs(v) < 1 else (VERDE if fav else ROJO)
            w = max(2.0, min(abs(v) / maxs[i], 1.0) * 56) if abs(v) >= 1 else 0
            txt = (f'<span style="color:{col}"><b>{"+" if v > 0 else ""}{formato(v)}</b>'
                   + (f'<small>{"+" if (c.get("pct") or 0) > 0 else ""}{pct_txt(c["pct"])}</small>'
                      if c.get("pct") is not None and not pd.isna(c.get("pct")) else "") + "</span>")
            barra = f'<i style="width:{w:.1f}%;background:{col}"></i>' if w else ""
            izq, der = (txt + barra, "") if v < 0 else ("", barra + txt)
            celdas += f'<div class="nf-dv"><div>{izq}</div><div>{der}</div></div>'
        sub = f'<small>{escape(f["sub"])}</small>' if f.get("sub") else ""
        cuerpo += (f'<div class="nf-dvr {f.get("clase", "")}" style="grid-template-columns:{plantilla}">'
                   f'<div class="nm"><b title="{escape(f["nombre"])}">{escape(f["nombre"])}</b>{sub}</div>'
                   f"{celdas}</div>")
    return (f'<div class="nf-scroll notranslate" translate="no" style="border:none">'
            f'<div class="nf-dvt" style="min-width:{150 + 160 * n}px">{cab}{cuerpo}</div></div>')


def informe(titulo: str, sub: str, secciones: list[tuple[str, list[str]]]) -> str:
    """Informe de resumen en viñetas: secciones = [(subtítulo, [frases con <b>…</b>])]."""
    cuerpo = "".join(f'<div class="nf-inf-t">{escape(t)}</div><ul>' + "".join(f"<li>{x}</li>" for x in items) + "</ul>"
                     for t, items in secciones if items)
    return (f'<div class="nf-inf notranslate" translate="no"><div class="nf-inf-h">{icono("bulb", AZUL, AZUL_SUAVE)}'
            f'<div><b>{escape(titulo)}</b><span>{escape(sub)}</span></div></div>{cuerpo}</div>')


def informe_texto(titulo: str, sub: str, secciones: list[tuple[str, list[str]]]) -> str:
    """El mismo informe en texto plano, para copiar en un correo."""
    import re as _re
    lineas = [titulo, sub, ""]
    for t, items in secciones:
        if items:
            lineas += [t.upper()] + [f"• {_re.sub(r'<[^>]+>', '', x)}" for x in items] + [""]
    return "\n".join(lineas).strip()


def var_html(v: float, gasto: bool = True) -> str:
    """Variación con signo y color: en gasto, subir es rojo."""
    clase = ("up" if v > 0 else "dn") if gasto else ("dn" if v > 0 else "up")
    return f'<span class="{clase}">{"+" if v > 0 else ""}{millones(v)}</span>'


def avisos(items: list[tuple]) -> str:
    """Cajas pequeñas de pendientes: (número, texto, tono 'warn'|'ok'|'bad'|'na')."""
    return ('<div class="nf-avisos notranslate" translate="no">'
            + "".join(f'<div class="nf-aviso {t}"><b>{escape(str(n))}</b>{escape(x)}</div>' for n, x, t in items)
            + "</div>")


def lista_avance(filas: list[tuple]) -> str:
    """filas: (nombre, ejecutado, presupuesto) ya en positivo, o (nombre, pa, pc, ec) en signo SAP."""
    out = ""
    for f in filas:
        if len(f) == 4:
            m = metricas(f[1], f[2], f[3])
            r, e, ing = m["pct"], m["ec"], m["ingreso"]
        else:
            nombre, e, p = f
            r, ing = (e / p if p else float("nan")), False
        out += (f'<div class="nf-prog"><div class="nf-prog-h"><span>{escape(f[0])}</span>'
                f'<b style="color:{color_barra(r, e, ing)}">{pct_txt(r)}</b></div>{barra(r, e, ing)}</div>')
    return f'<div class="notranslate" translate="no">{out}</div>'


# ------------------------------------------------------------------ tabla desplegable
COLS = ["Presupuesto anual", "Presupuesto al corte", "Ejecutado al corte", "Variación"]
MAX_TERCEROS = 15


def _celdas(r) -> str:
    """r: fila con 'Presupuesto anual', 'Presupuesto al corte', 'Ejecutado al corte' en signo SAP."""
    m = metricas(r["Presupuesto anual"], r["Presupuesto al corte"], r["Ejecutado al corte"])
    return (f"<div>{millones(m['pa'])}</div><div>{millones(m['pc'])}</div><div>{millones(m['ec'])}</div>"
            + f'<div class="{"nf-neg" if m["fav"] < -50_000 else ("nf-pos" if m["fav"] > 50_000 else "")}">'
            + f'{millones(m["var"])}</div>'
            + f'<div class="nf-pct">{barra(m["pct"], m["ec"], m["ingreso"])}<span>{pct_txt(m["pct"])}</span></div>'
            + f"<div>{pill(m['pct'], m['ec'], m['ingreso'])}</div>")


def _agregar(df: pd.DataFrame) -> pd.Series:
    return df[COLS[:3]].sum()


def _filas_terceros(t: pd.DataFrame, total_sub: float, signo: int = 1) -> str:
    """Filas de terceros de un sub-rubro: solo ejecución (el presupuesto no se abre por tercero)."""
    t = t.reindex(t["valor"].mul(signo).sort_values(ascending=False).index)
    if len(t) > MAX_TERCEROS:
        resto = t.iloc[MAX_TERCEROS:]
        t = pd.concat([t.iloc[:MAX_TERCEROS], pd.DataFrame([{
            "tercero": f"Otros {len(resto)} terceros", "valor": resto["valor"].sum(), "docs": resto["docs"].sum()}])])
    html = ""
    for _, r in t.iterrows():
        part = r["valor"] / total_sub if total_sub else float("nan")
        ancho = 0 if pd.isna(part) else max(0, min(part, 1)) * 100
        html += (f'<div class="nf-row nf-ter"><div class="nf-name"><span title="{escape(str(r["tercero"]))}">'
                 f'{escape(str(r["tercero"]))}</span></div><div>—</div><div>—</div>'
                 f'<div>{millones(signo * r["valor"])}</div><div>—</div>'
                 f'<div class="nf-pct"><div class="nf-bar"><div style="width:{ancho:.1f}%;background:#8fb0e3"></div>'
                 f'</div><span>{pct_txt(part)}</span></div>'
                 f'<div class="nf-docs">{int(r["docs"])} doc{"s" if r["docs"] != 1 else ""}</div></div>')
    return html


def _orden_rubro(rubro: str) -> tuple:
    import re
    return tuple(int(p) for p in re.findall(r"\d+", str(rubro).split(" ")[0])) or (99,)


def tabla_desplegable(det: pd.DataFrame, niveles: list[str], abierto: bool = False,
                      terceros: pd.DataFrame | None = None, etiquetas: list[str] | None = None) -> str:
    """
    det: una fila por sub-rubro con columnas COLS + '% ejecución' y las de `niveles`.
    niveles: ['rubro'] o ['gerencia', 'rubro'] -> grupos desplegables; luego 'sub_rubro'.
    terceros: (sub_rubro, tercero, valor, docs) -> cada sub-rubro se despliega en sus terceros.
    etiquetas: nombres visibles de las 4 columnas de valores.
    """
    op = " open" if abierto else ""
    etiquetas = etiquetas or COLS
    cab = ('<div class="nf-row nf-head"><div>' + " / ".join(n.replace("_", "-").capitalize() for n in niveles)
           + " / Sub-rubro" + (" / Tercero" if terceros is not None else "") + "</div>"
           + "".join(f"<div>{c}</div>" for c in etiquetas)
           + "<div>% ejecución</div><div>Estado</div></div>")
    por_sub = {}
    if terceros is not None and not terceros.empty:
        por_sub = {k: v for k, v in terceros.groupby(["rubro", "sub_rubro"])}

    def bloque(df, nivel_idx):
        col = niveles[nivel_idx]
        html = ""
        if col == "rubro":
            orden = sorted(df[col].unique(), key=lambda x: _orden_rubro(x))
        else:
            orden = sorted(df[col].unique(), key=lambda x: (x == "Sin asignar", str(x)))
        for nombre in orden:
            g = df[df[col] == nombre]
            agg = _agregar(g)
            n_sub = g["sub_rubro"].nunique()
            cls = "nf-l0" if nivel_idx == 0 else "nf-l1"
            html += (f'<details class="{cls}"{op}><summary class="nf-row"><div class="nf-name">{_CHEV}'
                     f'<span>{escape(str(nombre))}</span><small>{n_sub} sub-rubros</small></div>'
                     f"{_celdas(agg)}</summary>")
            if nivel_idx + 1 < len(niveles):
                html += bloque(g, nivel_idx + 1)
            else:
                g = g.reindex(g["Presupuesto anual"].abs().sort_values(ascending=False).index)
                for _, r in g.iterrows():
                    nom = escape(r["sub_rubro"])
                    ters = por_sub.get((r["rubro"], r["sub_rubro"]))
                    sg = -1 if es_ingreso(r["Presupuesto anual"], r["Presupuesto al corte"], r["Ejecutado al corte"]) else 1
                    if ters is not None and not ters.empty:
                        html += (f'<details class="nf-l2"><summary class="nf-row nf-sub"><div class="nf-name">{_CHEV}'
                                 f'<span title="{nom}">{nom}</span><small>{len(ters)} tercero'
                                 f'{"s" if len(ters) != 1 else ""}</small></div>{_celdas(r)}</summary>'
                                 f'{_filas_terceros(ters, ters["valor"].sum(), sg)}</details>')
                    else:
                        html += (f'<div class="nf-row nf-sub"><div class="nf-name"><span class="nf-sp"></span>'
                                 f'<span title="{nom}">{nom}</span></div>{_celdas(r)}</div>')
            html += "</details>"
        return html

    total = _agregar(det)
    pie = f'<div class="nf-row nf-total"><div>Total</div>{_celdas(total)}</div>'
    return (f'<div class="nf-scroll notranslate" translate="no"><div class="nf-t">{cab}{bloque(det, 0)}{pie}'
            f'</div></div>')


def lista_desviaciones(df: pd.DataFrame, sufijo: str = "al corte") -> str:
    """Lista compacta de desviaciones en contra: sub-rubro (con su rubro), presupuesto, ejecutado, desviación y estado."""
    cab = (f'<div class="nf-dev nf-head"><div>Sub-rubro</div><div>Presupuesto {sufijo}</div>'
           f'<div>Ejecutado {sufijo}</div><div>Desviación</div><div>Estado</div></div>')
    filas = ""
    for _, r in df.iterrows():
        m = metricas(r["Presupuesto anual"], r["Presupuesto al corte"], r["Ejecutado al corte"])
        filas += (f'<div class="nf-dev"><div class="nf-name" style="display:block"><span style="display:block;'
                  f'font-weight:600">{escape(r["sub_rubro"])}</span><small style="color:{TEXTO_2}">'
                  f'{escape(r["rubro"])}</small></div>'
                  f'<div>{millones(m["pc"])}</div><div>{millones(m["ec"])}</div>'
                  f'<div class="nf-neg" style="font-weight:700">{millones(m["var"])}</div>'
                  f'<div>{pill(m["pct"], m["ec"], m["ingreso"])}</div></div>')
    return (f'<div class="nf-t notranslate" translate="no" style="border:1px solid {BORDE};border-radius:12px;'
            f'overflow:hidden">{cab}{filas}</div>')


def lista_participacion(filas: list[tuple[str, float]], color: str = AZUL, total: float | None = None,
                        formato=None) -> str:
    """Barras de participación: (nombre, valor) -> valor y % del total."""
    formato = formato or millones
    total = total if total is not None else sum(v for _, v in filas)
    out = ""
    for nombre, v in filas:
        p = v / total if total else 0
        out += (f'<div class="nf-prog"><div class="nf-prog-h"><span>{escape(str(nombre))}</span>'
                f'<b>{formato(v)} <span style="color:{TEXTO_2};font-weight:500">· {pct_txt(p)}</span></b></div>'
                f'<div class="nf-bar"><div style="width:{max(0, min(p, 1)) * 100:.1f}%;background:{color}"></div>'
                f'</div></div>')
    return f'<div class="notranslate" translate="no">{out}</div>'


# ------------------------------------------------------------------ tabla jerárquica por meses (flujo de caja)
def _num(v, clase="") -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "<div>—</div>"
    return f'<div class="{clase}">{millones(v)}</div>'


def tabla_meses(nodos: list[dict], columnas: list[str], abierto: bool = False,
                formatos: list[str] | None = None, barras: list[int] | None = None, compacta: bool = False) -> str:
    """
    nodos: dict(nombre, valores=[...] (una por columna), clase in {'saldo','seccion','grupo','linea','tercero'},
                hijos=[nodos], dif_col=int|None (índice de la columna diferencia, se colorea), nota='')
    barras: columnas (en millones) que llevan una mini barra de variación junto al número (solo en las filas
            donde esa columna es la de diferencia, dif_col).
    compacta: primera columna más angosta, para tarjetas a media página.
    """
    n = len(columnas)
    barras = set()          # las mini barras junto a la variación se quitaron de los informes (solo número)
    tope = {}
    if barras:
        def recorrer(nds):
            for nd in nds:
                if nd.get("clase") not in ("saldo",):
                    for i in barras:
                        v = nd["valores"][i] if i < len(nd["valores"]) else None
                        if v is not None and not pd.isna(v):
                            tope[i] = max(tope.get(i, 0.0), abs(v))
                recorrer(nd.get("hijos") or [])
        recorrer(nodos)
    ancho_nom = 170 if compacta else 290
    ancho_v, ancho_b = (82, 138) if compacta else (96, 150)
    cols_t = [(f"minmax({ancho_b}px,1.35fr)" if i in (barras or []) else f"minmax({ancho_v}px,1fr)")
              for i in range(n)]
    plantilla = f"minmax({ancho_nom}px,{'2.6' if compacta else '2.8'}fr) " + " ".join(cols_t)
    min_ancho = ancho_nom + 30 + sum(ancho_b if i in (barras or []) else ancho_v for i in range(n))
    if not barras and not compacta:          # tabla de siempre: mismas medidas de antes
        plantilla, min_ancho = f"minmax(290px,2.8fr) repeat({n}, minmax(96px,1fr))", 300 + 100 * n
    cab = (f'<div class="nf-row nf-head" style="grid-template-columns:{plantilla}"><div>Concepto</div>'
           + "".join(f"<div>{escape(c)}</div>" for c in columnas) + "</div>")
    op = " open" if abierto else ""

    def fila(nodo, nivel, con_hijos):
        vals = ""
        d = nodo.get("dif_col")
        difs = set(d) if isinstance(d, (list, tuple, set)) else ({d} if d is not None else set())
        for i, v in enumerate(nodo["valores"]):
            cls = ""
            fmt = nodo.get("fmt") or (formatos[i] if formatos else "m")     # formato por fila o por columna
            umbral = 0 if fmt == "n" else (0.0005 if fmt in ("p", "x") else 50_000)
            if i in difs and v is not None and not pd.isna(v):
                vc = -v if nodo.get("dif_inv") else v          # dif_inv: en gastos, negativo es a favor (verde)
                cls = "nf-pos" if vc > umbral else ("nf-neg" if vc < -umbral else "")
            if v is None and nodo.get("clase") == "seccion" and all(x is None for x in nodo["valores"]):
                vals += "<div></div>"                  # fila de título sin cifras
            elif fmt == "p":
                vals += f'<div class="{cls}">{pct_txt(v) if v is not None else "—"}</div>'
            elif fmt == "x":                       # razón (veces), dos decimales
                vals += ("<div>—</div>" if v is None or pd.isna(v)
                         else f'<div class="{cls}">{v:,.2f}</div>'.replace(",", "X").replace(".", ",")
                         .replace("X", "."))
            elif fmt == "n":                       # cantidad (personas)
                vals += ("<div>—</div>" if v is None or pd.isna(v)
                         else f'<div class="{cls}">{v:,.0f}</div>'.replace(",", "."))
            elif i in barras and i in difs and v is not None and not pd.isna(v):
                col = {"nf-pos": VERDE, "nf-neg": ROJO}.get(cls, "#9aa7b8")
                w = min(abs(v) / tope[i], 1) * 50 if tope.get(i) else 0
                lado = "left:50%" if v >= 0 else "right:50%"
                vals += (f'<div class="nf-mb {cls}"><i><u style="{lado};width:{w:.1f}%;background:{col}"></u></i>'
                         f'<span>{millones(v)}</span></div>')
            else:
                vals += _num(v, cls)
        sangria = f"padding-left:{nivel * 1.3:.1f}rem"
        chev = _CHEV if con_hijos else '<span class="nf-sp"></span>'
        nota = (f'<small title="{escape(nodo["nota"])}">ⓘ</small>' if nodo.get("nota") else "")
        extra = f'<small>{escape(nodo["sub"])}</small>' if nodo.get("sub") else ""
        return (f'<div class="nf-name" style="{sangria}">{chev}<span title="{escape(nodo["nombre"])}">'
                f'{escape(nodo["nombre"])}</span>{extra}{nota}</div>{vals}')

    def render(nodo, nivel):
        clase = nodo.get("clase", "linea")
        hijos = nodo.get("hijos") or []
        estilo_fila = f' style="grid-template-columns:{plantilla}"'
        if hijos:
            return (f'<details class="nf-m-{clase}"{op if clase == "seccion" else ""}>'
                    f'<summary class="nf-row nf-m-{clase}"{estilo_fila}>{fila(nodo, nivel, True)}</summary>'
                    + "".join(render(h, nivel + 1) for h in hijos) + "</details>")
        return f'<div class="nf-row nf-m-{clase}"{estilo_fila}>{fila(nodo, nivel, False)}</div>'

    cuerpo = "".join(render(nd, 0) for nd in nodos)
    return (f'<div class="nf-scroll notranslate" translate="no"><div class="nf-t nf-tm" '
            f'style="min-width:{min_ancho}px">{cab}{cuerpo}</div></div>')
