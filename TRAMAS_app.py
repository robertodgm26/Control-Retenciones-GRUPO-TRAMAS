import streamlit as st
import pandas as pd
import plotly.graph_objects as go

# Configuración estética corporativa de la página
st.set_page_config(
    page_title="Control de Retenciones Logísticas", 
    page_icon="📊", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS personalizados (Calibración de Espacio y Simetría Definitiva)
st.markdown("""
    <style>
    .main { background-color: #f8fafc; }
    .kpi-box {
        background-color: #f0f5ff; /* Color de fondo azul claro premium */
        padding: 16px 12px;
        border-radius: 8px;
        border: 2px solid #cbd5e1; /* Bordes perfectamente demarcados */
        border-top: 5px solid #1E3A8A; /* Barra superior azul oscuro */
        box-shadow: 0px 4px 10px rgba(30, 58, 138, 0.05); /* Sombra suave */
        text-align: center;
        margin-bottom: 20px;
        min-height: 165px; /* Fuerza a las 3 cajas a tener la misma altura exacta */
    }
    .kpi-title { font-size: 14px; color: #1E3A8A; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; }
    .kpi-value { 
        font-size: 46px; 
        color: #0f172a; 
        font-weight: 800; 
        margin-top: 2px; 
        line-height: 1.1; 
    }
    .kpi-sub { 
        font-size: 14px; 
        font-weight: 700; 
        margin-top: 8px; 
        padding: 4px 10px; 
        border-radius: 4px; 
        display: inline-block; 
    }
    .sub-green { 
        color: #166534 !important; 
        background-color: #dcfce7 !important; 
        border: 1px solid #bbf7d0 !important; 
    }
    .sub-red { 
        color: #991b1b !important; 
        background-color: #fee2e2 !important; 
        border: 1px solid #fecaca !important; 
    }
    .instruccion-click { background-color: #eff6ff; color: #1d4ed8; padding: 12px; border-radius: 6px; font-size: 14px; font-weight: 500; border-left: 4px solid #2563eb; margin-bottom: 15px; border-right: 1px solid #bfdbfe; border-top: 1px solid #bfdbfe; border-bottom: 1px solid #bfdbfe; }
    .seccion-auditoria { background-color: #fffbeb; color: #b45309; padding: 12px; border-radius: 6px; font-size: 14px; font-weight: 700; border-left: 4px solid #d97706; margin-top: 25px; margin-bottom: 10px; border-right: 1px solid #fde68a; border-top: 1px solid #fde68a; border-bottom: 1px solid #fde68a; }
    </style>
""", unsafe_allow_html=True)

st.title("    Estatus GRUPO TRAMAS")
st.title("📊 Control de Retenciones: Pendientes vs Recibidas")
st.markdown("<p style='color: #4A5568; font-size: 16px;'>Consola de mando ejecutiva con filtros y segmentación directa de datos.</p>", unsafe_allow_html=True)
st.markdown("---")

# --- INITIALIZE SESSION STATE PARA MEMORIA DE INTERACCIÓN CRUZADA ---
if "click_cliente" not in st.session_state:
    st.session_state["click_cliente"] = None
if "click_estatus" not in st.session_state:
    st.session_state["click_estatus"] = None

# --- BARRA LATERAL: ENTORNO DE FILTROS Y CONTROL ---
with st.sidebar:
    st.header("🎛️ Panel de Control")
    uploaded_file = st.file_uploader("Selecciona el archivo de Excel (.xlsx)", type=["xlsx"])
    st.markdown("---")

if uploaded_file is not None:
    try:
        # Carga inteligente de la hoja buscando encabezados
        df_buscar = pd.read_excel(uploaded_file, sheet_name="DGM", header=None)
        fila_encabezado = 0
        for i, row in df_buscar.iterrows():
            row_str = [str(cell).strip().upper() for cell in row.values]
            if 'CONDICION' in row_str or 'CONDICIÓN' in row_str or 'CLIENTE' in row_str:
                fila_encabezado = i
                break
                
        df_raw = pd.read_excel(uploaded_file, sheet_name="DGM", skiprows=fila_encabezado)
        
        # Guardar nombres originales para el detalle final antes de normalizar a mayúsculas
        columnas_originales = list(df_raw.columns)
        df_raw.columns = df_raw.columns.astype(str).str.strip().str.upper()
        
        if 'CONDICIÓN' in df_raw.columns:
            df_raw.rename(columns={'CONDICIÓN': 'CONDICION'}, inplace=True)
            
        # Limpieza estándar de datos
        df_raw['CONDICION'] = df_raw['CONDICION'].astype(str).str.upper().str.strip()
        df_raw['MES'] = df_raw['MES'].astype(str).str.upper().str.strip()
        df_raw['CLIENTE'] = df_raw['CLIENTE'].astype(str).str.strip()
        
        # Replicar lógica del Estatus
        def calcular_estatus(row):
            if row['CONDICION'] == 'R': return 'RECIBIDA'
            elif row['CONDICION'] == 'P': return 'PENDIENTE'
            else: return 'OTRO'
        df_raw['ESTATUS'] = df_raw.apply(calcular_estatus, axis=1)
        
        # --- CONFIGURACIÓN DE FILTROS DE LISTA DESPLEGABLE DIRECTA ---
        with st.sidebar:
            st.subheader("Segmentar Datos")
            
            # REGLA 1: Al cambiar el combo, disparamos una función que resetea los clics para evitar bloqueos
            def cambio_selector_lateral():
                st.session_state["click_cliente"] = None
                st.session_state["click_estatus"] = None

            estatus_opciones = ["TODOS", "PENDIENTE", "RECIBIDA"]
            estatus_sel = st.selectbox("Condición / Estatus:", estatus_opciones, index=0, on_change=cambio_selector_lateral)
            
            st.markdown("---")
            
            mes_opciones = ["TODOS"] + sorted(df_raw['MES'].unique())
            mes_sel = st.selectbox("Mes del Periodo:", mes_opciones, index=0, on_change=cambio_selector_lateral)
                
            st.markdown("---")
            
            semana_opciones_base = sorted(df_raw['SEMANA'].dropna().unique())
            semana_opciones = ["TODOS"] + [str(int(s)) for s in semana_opciones_base]
            semana_sel = st.selectbox("Semana Operativa:", semana_opciones, index=0, on_change=cambio_selector_lateral)
        
        # --- FILTRADO DE TIEMPO (MES Y SEMANA) ---
        df_tiempo = df_raw.copy()
        if mes_sel != "TODOS":
            df_tiempo = df_tiempo[df_tiempo['MES'] == mes_sel]
        if semana_sel != "TODOS":
            df_tiempo = df_tiempo[df_tiempo['SEMANA'] == int(semana_sel)]
            
        # --- PROCESAMIENTO BASE DE MATRIZ COMPLETA ---
        pivot_df = df_tiempo.groupby(['CLIENTE', 'ESTATUS']).size().unstack(fill_value=0)
        for col in ['PENDIENTE', 'RECIBIDA', 'OTRO']:
            if col not in pivot_df.columns: 
                pivot_df[col] = 0
                
        pivot_df['Total General'] = pivot_df['PENDIENTE'] + pivot_df['RECIBIDA'] + pivot_df['OTRO']
        pivot_df = pivot_df.sort_values(by='Total General', ascending=False)
        
        # --- FILTRADO REGLA 1: Aplicación del Estatus de la barra lateral al Universo de Clientes ---
        if estatus_sel == "PENDIENTE":
            pivot_grafico = pivot_df[pivot_df['PENDIENTE'] > 0].copy()
        elif estatus_sel == "RECIBIDA":
            pivot_grafico = pivot_df[pivot_df['RECIBIDA'] > 0].copy()
        else:
            pivot_grafico = pivot_df.copy()
            
        if pivot_grafico.empty:
            st.warning("⚠️ No se encontraron registros con la combinación de filtros seleccionada.")
            st.stop()
            
        # --- SECCIÓN 1: DISEÑO DE TARJETAS KPIs (DINÁMICAS SEGÚN FILTRO LATERAL) ---
        total_g = int(pivot_grafico['Total General'].sum())
        total_p = int(pivot_grafico['PENDIENTE'].sum())
        total_r = int(pivot_grafico['RECIBIDA'].sum())
        
        pct_p = (total_p / total_g * 100) if total_g > 0 else 0
        pct_r = (total_r / total_g * 100) if total_g > 0 else 0
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown(f'<div class="kpi-box"><div class="kpi-title">Total Facturas</div><div class="kpi-value">{total_g:,}</div><div class="kpi-sub" style="background-color: transparent; border: none; color: transparent;">•</div></div>', unsafe_allow_html=True)
        with col2:
            st.markdown(f'<div class="kpi-box"><div class="kpi-title">Facturas Pendientes</div><div class="kpi-value">{total_p:,}</div><div class="kpi-sub sub-red">-{pct_p:.1f}% del volumen</div></div>', unsafe_allow_html=True)
        with col3:
            st.markdown(f'<div class="kpi-box"><div class="kpi-title">Facturas Recibidas</div><div class="kpi-value">{total_r:,}</div><div class="kpi-sub sub-green">+{pct_r:.1f}% cumplimiento</div></div>', unsafe_allow_html=True)
            
        # --- SECCIÓN 2: GRÁFICO INTERACTIVO PLANILLA PREMIUM ---
        st.subheader("📊 Relación de Retenciones por Cliente")
        st.markdown('<div class="instruccion-click">💡 <b>Interactividad y Auditoría Avanzada:</b> Haz clic sobre cualquier barra de un cliente para resaltar su comportamiento completo y extraer el desglose de filas originales en la parte inferior.</div>', unsafe_allow_html=True)
        
        # --- LÓGICA DE INSPECCIÓN PREVIA PARA OPACIDAD DINÁMICA ---
        cliente_seleccionado_previo = None
        estatus_grafico_seleccionado = None
        
        # Sincronización directa con el Session State para leer el clic antes de renderizar
        if "grafico_retenciones_definitivo" in st.session_state and st.session_state["grafico_retenciones_definitivo"] is not None:
            sel_state = st.session_state["grafico_retenciones_definitivo"]
            if hasattr(sel_state, "selection") and sel_state.selection is not None:
                puntos_lista = sel_state.selection.points
                if isinstance(puntos_lista, list) and len(puntos_lista) > 0:
                    punto_m = puntos_lista[0] # Extraemos de forma segura el diccionario del primer punto
                    if isinstance(punto_m, dict) and "x" in punto_m:
                        cliente_seleccionado_previo = str(punto_m["x"]).strip()
                    if isinstance(punto_m, dict) and "trace_index" in punto_m:
                        t_idx = punto_m["trace_index"]
                        if estatus_sel == "TODOS":
                            estatus_grafico_seleccionado = "PENDIENTE" if t_idx == 0 else "RECIBIDA"
                        else:
                            estatus_grafico_seleccionado = estatus_sel

        # Sincronización de estados estables de control
        c_sel = cliente_seleccionado_previo if cliente_seleccionado_previo else st.session_state["click_cliente"]
        e_sel = estatus_grafico_seleccionado if estatus_grafico_seleccionado else (None if estatus_sel == "TODOS" else estatus_sel)

        # Actualizar la memoria global antes de procesar las opacidades
        if c_sel:
            st.session_state["click_cliente"] = c_sel
        if estatus_grafico_seleccionado:
            st.session_state["click_estatus"] = estatus_grafico_seleccionado

        # --- CALCULO MATEMÁTICO DE ARRAYS DE OPACIDAD (CORREGIDO DE RAÍZ) ---
        opacidades_p = []
        opacidades_r = []
        
        for idx_cli in pivot_grafico.index:
            cli_str = str(idx_cli).strip()
            
            # A. Evaluación de Opacidades para la Serie de PENDIENTES (Azul)
            if estatus_sel == "RECIBIDA":
                opacidades_p.append(0.15)
            elif c_sel is None:
                opacidades_p.append(1.0) # Brillo total general
            else:
                # CORRECCIÓN CLAVE: Ambas barras del cliente seleccionado se quedan vivas (1.0)
                opacidades_p.append(1.0 if cli_str == c_sel else 0.15)
                    
            # B. Evaluación de Opacidades para la Serie de RECIBIDAS (Dorado)
            if estatus_sel == "PENDIENTE":
                opacidades_r.append(0.15)
            elif c_sel is None:
                opacidades_r.append(1.0) # Brillo total general
            else:
                # CORRECCIÓN CLAVE: Ambas barras del cliente seleccionado se quedan vivas (1.0)
                opacidades_r.append(1.0 if cli_str == c_sel else 0.15)

        # --- DIBUJO DEL GRÁFICO INTERACTIVO ---
        fig = go.Figure()
        
        # Traza de Pendientes
        fig.add_trace(go.Bar(
            x=pivot_grafico.index, y=pivot_grafico['PENDIENTE'], name='Pendientes',
            marker_color='#1E3A8A', marker=dict(opacity=opacidades_p), legendgroup='PENDIENTE'
        ))
        
        # Traza de Recibidas
        fig.add_trace(go.Bar(
            x=pivot_grafico.index, y=pivot_grafico['RECIBIDA'], name='Recibidas',
            marker_color='#F59E0B', marker=dict(opacity=opacidades_r), legendgroup='RECIBIDA'
        ))
        
        fig.update_layout(
            barmode='group', plot_bgcolor='white', paper_bgcolor='white',
            margin=dict(l=20, r=20, t=10, b=40),
            xaxis=dict(showgrid=False, tickfont=dict(color='#4A5568', size=12)),
            yaxis=dict(showgrid=True, gridcolor='#E2E8F0', tickfont=dict(color='#4A5568')),
            hovermode="x unified",
            legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5)
        )
        
        # Desplegar componente Plotly registrando su rerun key estable
        seleccion_grafico = st.plotly_chart(fig, use_container_width=True, on_select="rerun", key="grafico_retenciones_definitivo")
        st.markdown("---")
        
        # Botón Ejecutivo de Limpieza Total (Limpia combos y memoria interna)
        if c_sel or estatus_sel != "TODOS" or mes_sel != "TODOS" or semana_sel != "TODOS":
            if st.button("🔄 Quitar Filtro de Clientes (Mostrar Todo)"):
                st.session_state["click_cliente"] = None
                st.session_state["click_estatus"] = None
                st.cache_data.clear()
                st.rerun()
                
             
        # --- SECCIÓN 3: MATRIZ DE RESUMEN ANALÍTICO DINÁMICA ---
        if c_sel:
            st.subheader(f"📋 Matriz de Resumen Analítico: {c_sel}")
            df_tabla_final = pivot_grafico[pivot_grafico.index == c_sel].copy()
        else:
            st.subheader("📋 Matriz de Resumen Analítico: Todos los Clientes")
            df_tabla_final = pivot_grafico.copy()
            
        # Calcular relaciones de porcentajes ejecutivos
        df_tabla_final['Porcentaje de Pendientes'] = (df_tabla_final['PENDIENTE'] / df_tabla_final['Total General'] * 100).map('{:.1f}%'.format)
        df_tabla_final['Porcentaje de Recibidas'] = (df_tabla_final['RECIBIDA'] / df_tabla_final['Total General'] * 100).map('{:.1f}%'.format)
        df_tabla_final.index.name = "CLIENTE"
        
        # Sincronización estricta de las columnas de la matriz en base al estatus activo
        if e_sel == "PENDIENTE":
            columnas_matriz = ['PENDIENTE', 'Total General', 'Porcentaje de Pendientes']
        elif e_sel == "RECIBIDA":
            columnas_matriz = ['RECIBIDA', 'Total General', 'Porcentaje de Recibidas']
        else:
            columnas_matriz = ['PENDIENTE', 'RECIBIDA', 'Total General', 'Porcentaje de Pendientes', 'Porcentaje de Recibidas']
            
        st.dataframe(df_tabla_final[columnas_matriz], use_container_width=True)
        
        # --- SECCIÓN 4: EXTRACCIÓN DEL DESGLOSE GLOBAL DE AUDITORÍA ---
        df_auditoria = df_tiempo.copy()
        
        # Filtrado cruzado en cascada para la hoja de detalles maestra
        if c_sel and e_sel:
            st.markdown(f'🔍 DESGLOSE DE AUDITORÍA: {c_sel} — FACTURAS {e_sel}S', unsafe_allow_html=True)
            df_auditoria = df_auditoria[(df_auditoria['CLIENTE'] == c_sel) & (df_auditoria['ESTATUS'] == e_sel)]
        elif c_sel:
            st.markdown(f'🔍 DESGLOSE DE AUDITORÍA: Todos los Registros de {c_sel}', unsafe_allow_html=True)
            df_auditoria = df_auditoria[df_auditoria['CLIENTE'] == c_sel]
        elif e_sel:
            st.markdown(f'🔍 DESGLOSE DE AUDITORÍA GLOBAL: Detalle de todas las Facturas {e_sel}S', unsafe_allow_html=True)
            df_auditoria = df_auditoria[df_auditoria['ESTATUS'] == e_sel]
        else:
            st.markdown('🔍 DESGLOSE DE AUDITORÍA GLOBAL: Detalle de las Facturas Con los Filtros Actuales', unsafe_allow_html=True)
            
        # Devolver los nombres originales de las columnas de Excel
        df_auditoria_visual = df_auditoria.copy()
        columnas_mapeadas = {col.upper(): col for col in columnas_originales}
        df_auditoria_visual = df_auditoria_visual[[c for c in df_auditoria_visual.columns if c != 'ESTATUS']]
        df_auditoria_visual.rename(columns=columnas_mapeadas, inplace=True)
        
        # Eliminar automáticamente columnas sin título Unnamed
        df_auditoria_visual = df_auditoria_visual.loc[:, ~df_auditoria_visual.columns.astype(str).str.upper().str.startswith('UNNAMED')]
        
        # --- BARRIDO Y LIMPIEZA AUTOMÁTICA DE HORAS CON FORMATO LATINO (DD-MM-AAAA) ---
        for col in df_auditoria_visual.columns:
            if 'FECHA' in str(col).upper():
                # CORRECCIÓN DE RAÍZ: Formatear directamente sobre el dataframe visual filtrado de forma segura
                df_auditoria_visual[col] = pd.to_datetime(df_auditoria_visual[col], errors='coerce')
                df_auditoria_visual[col] = df_auditoria_visual[col].dt.strftime('%d-%m-%Y')
                df_auditoria_visual[col] = df_auditoria_visual[col].fillna('None').astype(str).replace({'NaT': 'None', 'nan': 'None'})

        st.dataframe(df_auditoria_visual, use_container_width=True, hide_index=True)
        
    except Exception as e:
        st.error(f"Error analítico al procesar la hoja 'DGM': {e}")
else:
    st.info("👋 Panel Ejecutivo: Despliega la barra lateral izquierda (botón '>>') y carga el archivo de Excel para activar el análisis.")

