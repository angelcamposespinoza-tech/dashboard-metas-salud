import streamlit as st
import pandas as pd
import plotly.express as px

# Configuración de página y estilo
st.set_page_config(page_title="Monitor Pro - Salud", layout="wide")
st.title("📊 Monitor de Rendimiento Jerárquico")

archivos = st.file_uploader("Subir Archivos de Dependencias", type=["xlsx"], accept_multiple_files=True)

def procesar_fila(fila):
    """Extrae logros y metas de una fila de Excel (salta de 3 en 3 columnas)."""
    meses = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic']
    logros, metas = [], []
    col_logro = 3 # Columna D en Excel
    for _ in meses:
        try:
            val_meta = fila.iloc[col_logro - 1]
            val_logro = fila.iloc[col_logro]
            
            # Conversión segura a número
            meta = float(val_meta) if pd.notna(val_meta) and str(val_meta).upper() != 'N/A' else 0
            logro = float(val_logro) if pd.notna(val_logro) and str(val_logro).upper() != 'N/A' else 0
        except:
            meta, logro = 0, 0
            
        metas.append(meta)
        logros.append(logro)
        col_logro += 3
    return metas, logros

if archivos:
    # 1. NIVEL 1: Selección de Dependencia (Archivo)
    nombres_archivos = [arc.name for arc in archivos]
    dependencia_sel = st.sidebar.selectbox("1. Seleccione Dependencia:", nombres_archivos)
    archivo_obj = next(arc for arc in archivos if arc.name == dependencia_sel)
    
    # Leer todas las hojas del archivo seleccionado
    dict_hojas = pd.read_excel(archivo_obj, sheet_name=None, header=None)
    
    # 2. NIVEL 2: Selección de Sede/Hoja (Drill-down)
    opciones_hojas = ["Resultados Generales (Consolidado)"] + list(dict_hojas.keys())
    sede_sel = st.sidebar.selectbox("2. Seleccione Sede/Hoja:", opciones_hojas)

    # --- CURACIÓN DINÁMICA DE INDICADORES (SOLUCIÓN AL ERROR ATTRIBUTEERROR) ---
    servicios_validos = set()
    for nombre_hoja, contenido in dict_hojas.items():
        df_temp = contenido.iloc[10:].copy()
        if not df_temp.empty:
            # Limpieza fila por fila para evitar errores de tipo de dato
            for val in df_temp.iloc[:, 0].dropna().unique():
                val_str = str(val).strip().upper()
                if val_str not in ['N/A', 'NAN', '', 'NONE']:
                    servicios_validos.add(val)

    indicador = st.selectbox("3. Seleccione el Servicio/Indicador:", sorted(list(servicios_validos)))

    # --- LÓGICA DE CONSOLIDACIÓN O FILTRADO ---
    meses_nombres = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic']
    logros_final = [0]*12
    metas_final = [0]*12

    if sede_sel == "Resultados Generales (Consolidado)":
        for h in dict_hojas.values():
            df_h = h.iloc[10:]
            fila = df_h[df_h.iloc[:, 0] == indicador]
            if not fila.empty:
                m, l = procesar_fila(fila.iloc[0])
                metas_final = [x + y for x, y in zip(metas_final, m)]
                logros_final = [x + y for x, y in zip(logros_final, l)]
    else:
        df_sede = dict_hojas[sede_sel].iloc[10:]
        fila = df_sede[df_sede.iloc[:, 0] == indicador]
        if not fila.empty:
            metas_final, logros_final = procesar_fila(fila.iloc[0])

    # Calcular porcentajes finales con protección contra división por cero
    porcentajes_final = [round((l/m*100),1) if m > 0 else 0 for l, m in zip(logros_final, metas_final)]

    # --- VISUALIZACIÓN DE DASHBOARD ---
    st.divider()
    st.subheader(f"Resultados para: {indicador} ({sede_sel})")
    
    # Gráfico Mensual Principal
    df_plot = pd.DataFrame({'Mes': meses_nombres, 'Logro': logros_final, 'Pct': porcentajes_final})
    fig_mensual = px.bar(df_plot, x='Mes', y='Logro', text=[f"{p}%" for p in porcentajes_final],
                         title="Avance Mensual y % de Cumplimiento", 
                         color='Logro', color_continuous_scale='Blues')
    fig_mensual.update_traces(textposition='outside')
    st.plotly_chart(fig_mensual, use_container_width=True)

    # --- KPIs TEMPORALES (TRIMESTRE Y SEMESTRE) ---
    col1, col2 = st.columns(2)
    
    with col1:
        # Trimestres
        tri_labels = ["T1 (Ene-Mar)", "T2 (Abr-Jun)", "T3 (Jul-Sep)", "T4 (Oct-Dic)"]
        logros_tri = [sum(logros_final[0:3]), sum(logros_final[3:6]), sum(logros_final[6:9]), sum(logros_final[9:12])]
        fig_tri = px.pie(names=tri_labels, values=logros_tri, title="Distribución por Trimestre", hole=0.4)
        st.plotly_chart(fig_tri, use_container_width=True)

    with col2:
        # Semestres
        sem_labels = ["1er Semestre", "2do Semestre"]
        logros_sem = [sum(logros_final[0:6]), sum(logros_final[6:12])]
        fig_sem = px.bar(x=sem_labels, y=logros_sem, title="Consolidado por Semestre", 
                         color=sem_labels, text=logros_sem,
                         labels={'x': 'Periodo', 'y': 'Total Logrado'})
        st.plotly_chart(fig_sem, use_container_width=True)

else:
    st.info("💡 Sube uno o varios archivos Excel para generar el reporte jerárquico.")
