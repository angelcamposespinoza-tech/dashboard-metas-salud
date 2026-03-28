import streamlit as st
import pandas as pd
import plotly.express as px

# Configuración profesional del dashboard
st.set_page_config(page_title="Monitor de Metas - Salud", layout="wide")
st.title("📊 Sistema de Evaluación de Metas (Jurisdicción nº1)")

archivos = st.file_uploader("Subir archivos Excel de Dependencias", type=["xlsx"], accept_multiple_files=True)

def procesar_fila(fila):
    """Extrae metas y logros basándose en la estructura exacta del formato de salud."""
    meses = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic']
    logros, metas = [], []
    # Según tu archivo, la columna 'C' (index 2) es la Meta de Enero y 'D' (index 3) es el Logro.
    col_meta = 2 
    for _ in meses:
        try:
            val_meta = fila.iloc[col_meta]
            val_logro = fila.iloc[col_meta + 1]
            
            # Limpieza y conversión a número
            meta = float(val_meta) if pd.notna(val_meta) and str(val_meta).upper() != 'N/A' else 0
            logro = float(val_logro) if pd.notna(val_logro) and str(val_logro).upper() != 'N/A' else 0
        except:
            meta, logro = 0, 0
            
        metas.append(meta)
        logros.append(logro)
        col_meta += 3 # Saltamos Meta, Logro, % para ir al siguiente mes
    return metas, logros

if archivos:
    # 1. Selección de Dependencia
    nombres_archivos = [arc.name for arc in archivos]
    dep_sel = st.sidebar.selectbox("Seleccione Municipio/Dependencia:", nombres_archivos)
    archivo_obj = next(arc for arc in archivos if arc.name == dep_sel)
    
    # Leer todas las sedes (hojas)
    dict_hojas = pd.read_excel(archivo_obj, sheet_name=None, header=None)
    
    # 2. Selección de Sede con opción de Consolidado
    opciones_hojas = ["TOTAL MUNICIPAL (Consolidado)"] + list(dict_hojas.keys())
    sede_sel = st.sidebar.selectbox("Seleccione Unidad Médica:", opciones_hojas)

    # --- EXTRACCIÓN DINÁMICA DE SERVICIOS ---
    servicios_validos = set()
    for contenido in dict_hojas.values():
        if contenido.shape[1] > 0:
            # En tus archivos, los nombres de servicios están en la columna A (index 0) desde la fila 11
            df_temp = contenido.iloc[10:].copy()
            for val in df_temp.iloc[:, 0].dropna().unique():
                val_str = str(val).strip()
                # Filtramos encabezados de sección y N/A
                if val_str.upper() not in ['N/A', 'SERVICIOS DE SALUD', 'FORTALECIMIENTO A LA ATENCIÓN MÉDICA', 'TRATO DIGNO']:
                    if len(val_str) > 3: # Evita ruidos de celdas pequeñas
                        servicios_validos.add(val_str)

    indicador = st.selectbox("Seleccione Indicador a evaluar:", sorted(list(servicios_validos)))

    # --- CÁLCULO DE RESULTADOS ---
    meses_nombres = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic']
    logros_f, metas_f = [0]*12, [0]*12

    if sede_sel == "TOTAL MUNICIPAL (Consolidado)":
        for h in dict_hojas.values():
            if h.shape[1] > 0:
                df_h = h.iloc[10:]
                fila = df_h[df_h.iloc[:, 0].astype(str).str.strip() == indicador]
                if not fila.empty:
                    m, l = procesar_fila(fila.iloc[0])
                    metas_f = [x + y for x, y in zip(metas_f, m)]
                    logros_f = [x + y for x, y in zip(logros_f, l)]
    else:
        df_sede = dict_hojas[sede_sel].iloc[10:]
        fila = df_sede[df_sede.iloc[:, 0].astype(str).str.strip() == indicador]
        if not fila.empty:
            metas_f, logros_f = procesar_fila(fila.iloc[0])

    pct_f = [round((l/m*100),1) if m > 0 else 0 for l, m in zip(logros_f, metas_f)]

    # --- DASHBOARD VISUAL ---
    st.divider()
    st.header(f"📌 {indicador}")
    st.caption(f"Análisis basado en: {sede_sel}")

    # Gráfica Principal
    df_plot = pd.DataFrame({'Mes': meses_nombres, 'Logro': logros_f, 'Pct': pct_f})
    fig = px.bar(df_plot, x='Mes', y='Logro', text=[f"{p}%" for p in pct_f],
                 color='Pct', color_continuous_scale='RdYlGn', range_color=[0, 100],
                 title="Cumplimiento Mensual (El color indica cercanía a la meta)")
    st.plotly_chart(fig, use_container_width=True)

    # Bloque de Semestres y Trimestres
    c1, c2 = st.columns(2)
    with c1:
        logros_tri = [sum(logros_f[0:3]), sum(logros_f[3:6]), sum(logros_f[6:9]), sum(logros_f[9:12])]
        st.plotly_chart(px.pie(names=["T1", "T2", "T3", "T4"], values=logros_tri, 
                               title="Distribución Anual por Trimestre", hole=0.5), use_container_width=True)
    with c2:
        logros_sem = [sum(logros_f[0:6]), sum(logros_f[6:12])]
        st.plotly_chart(px.bar(x=["1er Semestre", "2do Semestre"], y=logros_sem, 
                               text=logros_sem, title="Comparativa Semestral",
                               color_discrete_sequence=['#003366']), use_container_width=True)

else:
    st.info("👋 Bienvenido. Por favor sube los archivos Excel de la Jurisdicción para comenzar.")
