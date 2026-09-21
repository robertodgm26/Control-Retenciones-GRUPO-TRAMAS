import base64
import io
import os
from datetime import datetime

import pandas as pd
import plotly.graph_objects as go
import requests
from urllib.parse import parse_qs, urlparse
import streamlit as st

# ─────────────────────────────────────────────────────────────
# CONFIGURACIÓN GENERAL
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Control de Retenciones Logísticas",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

HOJA_PREFERIDA = "DGM"      # Hoja del Excel donde está la matriz
TTL_SEGUNDOS = 300          # El Excel se vuelve a descargar como máximo cada 5 min
COLOR_P = (30, 58, 138)     # Azul  -> Pendientes
COLOR_R = (245, 158, 11)    # Dorado -> Recibidas
ALFA_OPACO = 0.15           # Nivel de opacidad de lo "no seleccionado"

st.markdown("""
    <style>
    .main { background-color: #f8fafc; }
    .kpi-box {
        background-color: #f0f5ff; padding: 16px 12px; border-radius: 8px;
        border: 2px solid #cbd5e1; border-top: 5px solid #1E3A8A;
        box-shadow: 0px 4px 10px rgba(30, 58, 138, 0.05);
        text-align: center; margin-bottom: 20px; min-height: 165px;
    }
    .kpi-title { font-size: 14px; color: #1E3A8A; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; }
    .kpi-value { font-size: 46px; color: #0f172a; font-weight: 800; margin-top: 2px; line-height: 1.1; }
    .kpi-sub { font-size: 14px; font-weight: 700; margin-top: 8px; padding: 4px 10px; border-radius: 4px; display: inline-block; }
    .sub-green { color: #166534 !important; background-color: #dcfce7 !important; border: 1px solid #bbf7d0 !important; }
    .sub-red   { color: #991b1b !important; background-color: #fee2e2 !important; border: 1px solid #fecaca !important; }
    .sub-none  { color: transparent !important; background-color: transparent !important; border: none !important; }
    .instruccion-click { background-color: #eff6ff; color: #1d4ed8; padding: 12px; border-radius: 6px; font-size: 14px; font-weight: 500; border-left: 4px solid #2563eb; margin-bottom: 15px; border: 1px solid #bfdbfe; border-left: 4px solid #2563eb; }
    .seccion-auditoria { background-color: #fffbeb; color: #b45309; padding: 12px; border-radius: 6px; font-size: 14px; font-weight: 700; border: 1px solid #fde68a; border-left: 4px solid #d97706; margin-top: 25px; margin-bottom: 10px; }
    </style>
""", unsafe_allow_html=True)

S = st.session_state

# ─────────────────────────────────────────────────────────────
# ESTADO DE LA SESIÓN (filtros de la barra lateral + versión del gráfico)
# ─────────────────────────────────────────────────────────────
S.setdefault("f_estatus", "TODOS")
S.setdefault("f_mes", "TODOS")
S.setdefault("f_semana", "TODOS")
OPC_BARRA = "Cliente + estatus de la barra"
OPC_CLIENTE = "Cliente completo (Pendientes y Recibidas)"
S.setdefault("modo_clic", OPC_BARRA)
S.setdefault("chart_ver", 0)   # Al cambiar, el gráfico se recrea y su selección se borra


def reiniciar_seleccion():
    """Se ejecuta al cambiar un filtro lateral: descarta el clic previo en el gráfico."""
    S["chart_ver"] += 1


def limpiar_todo():
    """Botón 'Quitar Filtro de Clientes (Mostrar Todo)': deja todo en su estado inicial."""
    S["f_estatus"] = "TODOS"
    S["f_mes"] = "TODOS"
    S["f_semana"] = "TODOS"
    S["chart_ver"] += 1


# ─────────────────────────────────────────────────────────────
# CONEXIÓN A ONEDRIVE
# ─────────────────────────────────────────────────────────────
def obtener_url_onedrive() -> str:
    """El enlace se guarda en los 'Secrets' de Streamlit (no dentro del código)."""
    try:
        url = st.secrets["ONEDRIVE_URL"]
    except Exception:
        url = os.environ.get("ONEDRIVE_URL", "")
    return str(url).strip()


def _token_compartido(url: str) -> str:
    """Codificación oficial de Microsoft para enlaces compartidos: u!<base64url>."""
    b64 = base64.urlsafe_b64encode(url.encode("utf-8")).decode("utf-8").rstrip("=")
    return "u!" + b64


def _es_xlsx(r) -> bool:
    return r.status_code == 200 and r.content[:2] == b"PK"   # un .xlsx es un ZIP


def _candidatos_descarga(url: str, url_final: str) -> list:
    """Genera las URL de descarga directa posibles a partir del enlace compartido."""
    cands = []
    for base in dict.fromkeys([url, url_final]):          # original y ya resuelto
        if base:
            cands.append(f"https://api.onedrive.com/v1.0/shares/{_token_compartido(base)}/root/content")
            cands.append(f"https://api.onedrive.com/v1.0/shares/{_token_compartido(base.split('#')[0])}/driveItem/content")

    if url_final:
        q = parse_qs(urlparse(url_final).query)
        g = lambda k: (q.get(k) or [""])[0]
        cid, resid, auth = g("cid"), g("resid") or g("id"), g("authkey")
        if cid and resid:                                  # OneDrive personal (Outlook/Hotmail)
            extra = f"&authkey={auth}" if auth else ""
            cands.append(f"https://onedrive.live.com/download?cid={cid}&resid={resid}{extra}")
        for viejo in ("/redir?", "/edit?", "/view.aspx?", "/embed?"):
            if "onedrive.live.com" in url_final and viejo in url_final:
                cands.append(url_final.replace(viejo, "/download?"))
        sep = "&" if "?" in url_final else "?"
        cands.append(f"{url_final}{sep}download=1")        # SharePoint / OneDrive empresarial
    sep = "&" if "?" in url else "?"
    cands.append(f"{url}{sep}download=1")
    return list(dict.fromkeys(cands))


@st.cache_data(ttl=TTL_SEGUNDOS, show_spinner="🔄 Leyendo el Excel actualizado desde OneDrive…")
def descargar_excel(url: str, ruta_local: str):
    """Devuelve (bytes_del_xlsx, fecha_hora_de_lectura)."""
    if ruta_local:  # Modo prueba local
        with open(ruta_local, "rb") as f:
            return f.read(), datetime.now()

    url = url.strip().strip('"').strip("'")
    cabeceras = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    sesion = requests.Session()
    diag = []

    # 1) Resolver redirecciones (los enlaces 1drv.ms son cortos y redirigen a la URL real)
    url_final = ""
    try:
        r0 = sesion.get(url, headers=cabeceras, timeout=60, allow_redirects=True)
        url_final = r0.url
        if _es_xlsx(r0):
            return r0.content, datetime.now()
        diag.append(f"enlace directo: HTTP {r0.status_code} · {r0.headers.get('content-type', '?')[:40]} · {urlparse(url_final).netloc}")
    except requests.RequestException as e:
        diag.append(f"enlace directo: {str(e)[:100]}")

    # 2) Probar las variantes de descarga directa
    for i, destino in enumerate(_candidatos_descarga(url, url_final), start=1):
        try:
            r = sesion.get(destino, headers=cabeceras, timeout=60, allow_redirects=True)
            if _es_xlsx(r):
                return r.content, datetime.now()
            diag.append(f"variante {i} ({urlparse(destino).netloc}): HTTP {r.status_code} · {r.headers.get('content-type', '?')[:30]}")
        except requests.RequestException as e:
            diag.append(f"variante {i}: {str(e)[:80]}")
    raise RuntimeError("\n".join(diag))


# ─────────────────────────────────────────────────────────────
# LECTURA Y LIMPIEZA DE LA MATRIZ
# ─────────────────────────────────────────────────────────────
def _clasificar(valor) -> str:
    s = str(valor).strip().upper()
    if s.startswith("R"):
        return "RECIBIDA"
    if s.startswith("P"):
        return "PENDIENTE"
    return "OTRO"


@st.cache_data(ttl=TTL_SEGUNDOS, show_spinner=False)
def cargar_datos(contenido: bytes):
    xls = pd.ExcelFile(io.BytesIO(contenido), engine="openpyxl")
    hojas = xls.sheet_names
    if HOJA_PREFERIDA in hojas:
        hojas = [HOJA_PREFERIDA] + [h for h in hojas if h != HOJA_PREFERIDA]

    hoja_ok, fila_cab = None, None
    for hoja in hojas:
        crudo = xls.parse(hoja, header=None)
        for i, fila in crudo.iterrows():
            textos = [str(c).strip().upper() for c in fila.values if pd.notna(c)]
            if any("CLIENTE" in t for t in textos) and any("CONDICI" in t for t in textos):
                hoja_ok, fila_cab = hoja, i
                break
        if hoja_ok:
            break
    if hoja_ok is None:
        raise ValueError("No se encontró ninguna hoja con las columnas CLIENTE y CONDICIONES.")

    df = xls.parse(hoja_ok, header=fila_cab)
    df = df.loc[:, ~df.columns.astype(str).str.upper().str.startswith("UNNAMED")]

    # nombre original (para mostrar) y nombre normalizado en mayúsculas (para trabajar)
    originales = {}
    nuevos = []
    for c in df.columns:
        n = str(c).strip().upper()
        if "CONDICI" in n:
            n = "CONDICION"
        originales[n] = str(c).strip()
        nuevos.append(n)
    df.columns = nuevos
    df = df.loc[:, ~pd.Index(df.columns).duplicated()]

    # Quitar filas vacías / sin cliente
    df = df[df["CLIENTE"].notna()].copy()
    df["CLIENTE"] = df["CLIENTE"].astype(str).str.strip()
    df = df[df["CLIENTE"] != ""]

    if "MES" in df.columns:
        df["MES"] = df["MES"].astype(str).str.upper().str.strip()
    if "SEMANA" in df.columns:
        df["SEMANA"] = pd.to_numeric(df["SEMANA"], errors="coerce")

    df["_ESTATUS"] = df["CONDICION"].map(_clasificar)
    # Si la condición viene vacía, apoyarse en la columna ESTATUS del Excel
    if "ESTATUS" in df.columns:
        sin_cond = df["CONDICION"].isna() | (df["CONDICION"].astype(str).str.strip() == "")
        aux = df.loc[sin_cond, "ESTATUS"].astype(str).str.upper().str.strip()
        df.loc[sin_cond & aux.str.startswith("PEND"), "_ESTATUS"] = "PENDIENTE"
        df.loc[sin_cond & aux.str.startswith("RECIB"), "_ESTATUS"] = "RECIBIDA"

    return df.reset_index(drop=True), originales, hoja_ok


# ─────────────────────────────────────────────────────────────
# ENCABEZADO + BOTÓN DE LIMPIEZA (siempre visible)
# ─────────────────────────────────────────────────────────────
st.title("📊 Control de Retenciones: Pendientes vs Recibidas")
st.markdown("<p style='color:#4A5568;font-size:16px;'>Consola de mando ejecutiva con filtros y segmentación directa de datos.</p>", unsafe_allow_html=True)

url_onedrive = obtener_url_onedrive()
ruta_local = os.environ.get("EXCEL_LOCAL_PATH", "")

with st.sidebar:
    st.header("🎛️ Panel de Control")

if not url_onedrive and not ruta_local:
    st.error("⚠️ Falta el enlace del Excel. Agrega `ONEDRIVE_URL` en los *Secrets* de la app (ver instrucciones).")
    st.stop()

try:
    contenido, hora_lectura = descargar_excel(url_onedrive, ruta_local)
except Exception as e:
    st.error(
        "❌ No se pudo leer el Excel desde OneDrive. Verifica que el enlace sea de tipo "
        "**«Cualquier persona con el vínculo puede ver»**.\n\nDetalle técnico:\n\n```\n" + str(e) + "\n```"
    )
    st.stop()

try:
    df_raw, nombres_originales, hoja_usada = cargar_datos(contenido)
except Exception as e:
    st.error(f"❌ Error interpretando el archivo Excel: {e}")
    st.stop()

st.button("🔄 Quitar Filtro de Clientes (Mostrar Todo)", on_click=limpiar_todo, type="primary")
st.markdown("---")

# ─────────────────────────────────────────────────────────────
# BARRA LATERAL
# ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.success("🔄 **Conexión automática:** los datos se leen del Excel en OneDrive.")
    st.caption(f"Hoja: **{hoja_usada}** · Última lectura: {hora_lectura:%d-%m-%Y %H:%M:%S}")
    if st.button("♻️ Actualizar datos ahora"):
        st.cache_data.clear()
        st.rerun()
    st.markdown("---")
    st.subheader("Segmentar Datos")

    st.selectbox("Condición / Estatus:", ["TODOS", "PENDIENTE", "RECIBIDA"],
                 key="f_estatus", on_change=reiniciar_seleccion)
    st.markdown("---")

    meses = sorted(df_raw["MES"].unique()) if "MES" in df_raw.columns else []
    st.selectbox("Mes del Periodo:", ["TODOS"] + meses, key="f_mes", on_change=reiniciar_seleccion)
    st.markdown("---")

    semanas = sorted(df_raw["SEMANA"].dropna().unique()) if "SEMANA" in df_raw.columns else []
    st.selectbox("Semana Operativa:", ["TODOS"] + [str(int(s)) for s in semanas],
                 key="f_semana", on_change=reiniciar_seleccion)

estatus_sel = S["f_estatus"]
mes_sel = S["f_mes"]
semana_sel = S["f_semana"]

# ─────────────────────────────────────────────────────────────
# FILTRO DE TIEMPO Y MATRIZ BASE
# ─────────────────────────────────────────────────────────────
df_tiempo = df_raw.copy()
if mes_sel != "TODOS" and "MES" in df_tiempo.columns:
    df_tiempo = df_tiempo[df_tiempo["MES"] == mes_sel]
if semana_sel != "TODOS" and "SEMANA" in df_tiempo.columns:
    df_tiempo = df_tiempo[df_tiempo["SEMANA"] == int(semana_sel)]

pivot_df = df_tiempo.groupby(["CLIENTE", "_ESTATUS"]).size().unstack(fill_value=0)
for col in ["PENDIENTE", "RECIBIDA", "OTRO"]:
    if col not in pivot_df.columns:
        pivot_df[col] = 0
pivot_df = pivot_df[["PENDIENTE", "RECIBIDA", "OTRO"]]
pivot_df["Total General"] = pivot_df.sum(axis=1)

# Filtro de estatus lateral → clientes que aparecen en el gráfico
if estatus_sel == "PENDIENTE":
    pivot_grafico = pivot_df[pivot_df["PENDIENTE"] > 0].copy()
    pivot_grafico = pivot_grafico.sort_values("PENDIENTE", ascending=False)
elif estatus_sel == "RECIBIDA":
    pivot_grafico = pivot_df[pivot_df["RECIBIDA"] > 0].copy()
    pivot_grafico = pivot_grafico.sort_values("RECIBIDA", ascending=False)
else:
    pivot_grafico = pivot_df.sort_values("Total General", ascending=False).copy()

if pivot_grafico.empty:
    st.warning("⚠️ No se encontraron registros con la combinación de filtros seleccionada. "
               "Usa el botón «Quitar Filtro de Clientes (Mostrar Todo)».")
    st.stop()

# ─────────────────────────────────────────────────────────────
# LECTURA DEL CLIC EN EL GRÁFICO (estado previo, antes de dibujar)
# ─────────────────────────────────────────────────────────────
clave_grafico = f"grafico_retenciones_{S['chart_ver']}"
c_sel, curva = None, None
estado = S.get(clave_grafico)
if estado:
    try:
        puntos = (estado.get("selection") or {}).get("points") or []
    except AttributeError:
        puntos = []
    if puntos:
        p = puntos[0]
        candidato = str(p.get("x", "")).strip()
        if candidato in [str(i).strip() for i in pivot_grafico.index]:
            c_sel, curva = candidato, p.get("curve_number")

# Estatus efectivo y si se hizo clic en una barra concreta
barra_sel = False
if estatus_sel != "TODOS":
    e_sel = estatus_sel                       # filtro lateral manda
elif c_sel is not None and curva in (0, 1) and S["modo_clic"] == OPC_BARRA:
    e_sel = "PENDIENTE" if curva == 0 else "RECIBIDA"   # clic en barra azul / dorada
    barra_sel = True
else:
    e_sel = None


def opacidad(traza: str, cliente: str) -> float:
    """1.0 = color vivo · ALFA_OPACO = opaca."""
    if c_sel is None:
        if estatus_sel == "TODOS":
            return 1.0
        return 1.0 if traza == estatus_sel else ALFA_OPACO
    if cliente != c_sel:
        return ALFA_OPACO
    # Cliente seleccionado: solo queda viva la barra del estatus activo
    # (fijado en la barra lateral o elegido con el clic en la barra azul/dorada)
    return 1.0 if (e_sel is None or traza == e_sel) else ALFA_OPACO


# ─────────────────────────────────────────────────────────────
# KPIs (respetan filtros de tiempo y cliente seleccionado)
# ─────────────────────────────────────────────────────────────
df_kpi = df_tiempo if c_sel is None else df_tiempo[df_tiempo["CLIENTE"] == c_sel]
total_g = len(df_kpi)
total_p = int((df_kpi["_ESTATUS"] == "PENDIENTE").sum())
total_r = int((df_kpi["_ESTATUS"] == "RECIBIDA").sum())
pct_p = total_p / total_g * 100 if total_g else 0
pct_r = total_r / total_g * 100 if total_g else 0

k1, k2, k3 = st.columns(3)
k1.markdown(f'<div class="kpi-box"><div class="kpi-title">Total Facturas{" · " + c_sel if c_sel else ""}</div>'
            f'<div class="kpi-value">{total_g:,}</div><div class="kpi-sub sub-none">•</div></div>', unsafe_allow_html=True)
k2.markdown(f'<div class="kpi-box"><div class="kpi-title">Facturas Pendientes</div>'
            f'<div class="kpi-value">{total_p:,}</div><div class="kpi-sub sub-red">{pct_p:.1f}% del volumen</div></div>', unsafe_allow_html=True)
k3.markdown(f'<div class="kpi-box"><div class="kpi-title">Facturas Recibidas</div>'
            f'<div class="kpi-value">{total_r:,}</div><div class="kpi-sub sub-green">{pct_r:.1f}% cumplimiento</div></div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# GRÁFICO INTERACTIVO
# ─────────────────────────────────────────────────────────────
if estatus_sel == "TODOS":
    st.radio("Al hacer clic en una barra mostrar:", [OPC_BARRA, OPC_CLIENTE],
             key="modo_clic", horizontal=True)

clientes = [str(i) for i in pivot_grafico.index]
col_p = [f"rgba({COLOR_P[0]},{COLOR_P[1]},{COLOR_P[2]},{opacidad('PENDIENTE', c)})" for c in clientes]
col_r = [f"rgba({COLOR_R[0]},{COLOR_R[1]},{COLOR_R[2]},{opacidad('RECIBIDA', c)})" for c in clientes]
val_p = pivot_grafico["PENDIENTE"].tolist()
val_r = pivot_grafico["RECIBIDA"].tolist()

fig = go.Figure()
for nombre, valores, colores in (("Pendientes", val_p, col_p), ("Recibidas", val_r, col_r)):
    fig.add_trace(go.Bar(
        x=clientes, y=valores, name=nombre,
        marker=dict(color=colores),
        text=[str(v) if v else "" for v in valores], textposition="outside", cliponaxis=False,
        selected=dict(marker=dict(opacity=1)), unselected=dict(marker=dict(opacity=1)),
        hovertemplate="<b>%{x}</b><br>" + nombre + ": %{y}<extra></extra>",
    ))

fig.update_layout(
    barmode="group", plot_bgcolor="white", paper_bgcolor="white",
    margin=dict(l=20, r=20, t=30, b=40), height=460,
    xaxis=dict(showgrid=False, tickfont=dict(color="#4A5568", size=12), fixedrange=True),
    yaxis=dict(showgrid=True, gridcolor="#E2E8F0", tickfont=dict(color="#4A5568"), fixedrange=True),
    hovermode="closest", clickmode="event+select",
    legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5),
)

st.plotly_chart(fig, width="stretch", on_select="rerun", selection_mode="points",
                key=clave_grafico, config={"displayModeBar": False})

st.markdown("""
<div class="instruccion-click" style="padding:8px 12px; margin:0 0 4px 0; line-height:1.45; font-size:13px;">
<b>💡 Cómo usar el gráfico y los filtros</b><br>
<b>1️⃣ Estatus en «TODOS»:</b> clic en la <b>barra azul</b> (Pendientes) o <b>dorada</b> (Recibidas) de un cliente → solo ese estatus de ese cliente.
Para ver ambos juntos, elige antes <i>«Cliente completo»</i> arriba del gráfico.<br>
<b>2️⃣ Estatus en «PENDIENTE» o «RECIBIDA»:</b> el gráfico muestra los clientes con ese estatus; clic en cualquier barra de un cliente → solo ese estatus de ese cliente.
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# MATRIZ DE RESUMEN ANALÍTICO
# ─────────────────────────────────────────────────────────────
if c_sel:
    st.subheader(f"📋 Matriz de Resumen Analítico: {c_sel}")
    tabla = pivot_grafico[pivot_grafico.index.astype(str) == c_sel].copy()
else:
    st.subheader("📋 Matriz de Resumen Analítico: Todos los Clientes")
    tabla = pivot_grafico.copy()

if len(tabla) > 1:
    fila_total = tabla.sum(numeric_only=True).to_frame("TOTAL").T
    tabla = pd.concat([tabla, fila_total])


def _pct(num, den):
    return (num / den * 100).where(den > 0, 0).map("{:.1f}%".format)


tabla["Porcentaje de Pendientes"] = _pct(tabla["PENDIENTE"], tabla["Total General"])
tabla["Porcentaje de Recibidas"] = _pct(tabla["RECIBIDA"], tabla["Total General"])
tabla.index.name = "CLIENTE"

if e_sel == "PENDIENTE":
    cols_matriz = ["PENDIENTE", "Total General", "Porcentaje de Pendientes"]
elif e_sel == "RECIBIDA":
    cols_matriz = ["RECIBIDA", "Total General", "Porcentaje de Recibidas"]
else:
    cols_matriz = ["PENDIENTE", "RECIBIDA", "Total General", "Porcentaje de Pendientes", "Porcentaje de Recibidas"]
st.dataframe(tabla[cols_matriz], width="stretch")

# ─────────────────────────────────────────────────────────────
# DESGLOSE GLOBAL DE AUDITORÍA
# ─────────────────────────────────────────────────────────────
df_aud = df_tiempo.copy()
if c_sel and e_sel:
    titulo = f"{c_sel} — FACTURAS {e_sel}S"
    df_aud = df_aud[(df_aud["CLIENTE"] == c_sel) & (df_aud["_ESTATUS"] == e_sel)]
elif c_sel:
    titulo = f"Todos los Registros de {c_sel}"
    df_aud = df_aud[df_aud["CLIENTE"] == c_sel]
elif e_sel:
    titulo = f"GLOBAL: Detalle de todas las Facturas {e_sel}S"
    df_aud = df_aud[df_aud["_ESTATUS"] == e_sel]
else:
    titulo = "GLOBAL: Detalle de las Facturas con los Filtros Actuales"
st.markdown(f'<div class="seccion-auditoria">🔍 DESGLOSE DE AUDITORÍA: {titulo} ({len(df_aud):,} registros)</div>',
            unsafe_allow_html=True)

vista = df_aud.drop(columns=["_ESTATUS"]).copy()
if "DIAS TRANSCURRIDOS" in vista.columns:
    vista["DIAS TRANSCURRIDOS"] = pd.to_numeric(vista["DIAS TRANSCURRIDOS"], errors="coerce")
    vista = vista.sort_values("DIAS TRANSCURRIDOS", ascending=False)


def _a_texto(v):
    if pd.isna(v):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


for c in vista.columns:
    if "FECHA" in c:
        f = pd.to_datetime(vista[c], errors="coerce", dayfirst=True)
        vista[c] = f.dt.strftime("%d-%m-%Y").fillna("")
    elif c not in ("SEMANA", "DIAS TRANSCURRIDOS"):
        vista[c] = vista[c].map(_a_texto)

vista = vista.rename(columns=nombres_originales)
st.dataframe(vista, width="stretch", hide_index=True)
st.download_button("⬇️ Descargar este desglose (CSV)", vista.to_csv(index=False).encode("utf-8-sig"),
                   file_name="desglose_retenciones.csv", mime="text/csv")
