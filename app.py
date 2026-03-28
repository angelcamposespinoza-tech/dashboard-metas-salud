import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Monitor Pro - Salud", layout="wide")
st.title("📊 Monitor de Rendimiento Jerárquico")

archivos = st.file_uploader("Subir Archivos de Dependencias", type=["xlsx"], accept_multiple_files=True)

def procesar_fila(fila):
    """Lógica para extraer logros y calcular porcentajes de una fila."""
    meses = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic']
    logros, metas = [], []
    col_logro = 3
    for _ in meses:
        val_meta = fila.iloc[col_logro - 1]
        val_logro = fila.iloc[col_logro]
        meta = float(val_meta) if pd.notna(val_meta) and str(val_meta).upper() != 'N/A' else 0
        logro = float(val_logro) if pd.notna(val_logro) and str(val_logro).upper() != 'N/A' else 0
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

    # --- CURACIÓN DINÁMICA DE INDICADORES ---
    # Obtenemos servicios que no sean N/A y tengan datos reales
    servicios_validos = set()
    for nombre_hoja, contenido in dict_hojas.items():
        df_temp = contenido.iloc[10:].copy()
        # Filtramos filas donde la primera columna no sea nula o N/A
        servicios = df_temp[df_temp.iloc[:, 0].notna() & (df_temp.iloc[:, 0].astype(str).upper() != 'N/A')].iloc[:, 0].unique()
        servicios_validos.update(servicios)

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

    # Calcular porcentajes finales
    porcentajes_final = [round((l/m*100),1) if m > 0 else 0 for l, m in zip(logros_final, metas_final)]

    # --- VISUALIZACIÓN TEMPORAL (KPIs) ---
    st.divider()
    
    # Gráfico Mensual
    df_plot = pd.DataFrame({'Mes': meses_nombres, 'Logro': logros_final, 'Pct': porcentajes_final})
    fig_mensual = px.bar(df_plot, x='Mes', y='Logro', text=[f"{p}%" for p in porcentajes_final],
                         title=f"Avance Mensual: {indicador}", color='Logro', color_continuous_scale='Blues')
    st.plotly_chart(fig_mensual, use_container_width=True)

    # 3. COMPARATIVO POR TRIMESTRE Y SEMESTRE
    col1, col2 = st.columns(2)
    
    with col1:
        # Trimestres
        tri = ["T1", "T2", "T3", "T4"]
        logros_tri = [sum(logros_final[0:3]), sum(logros_final[3:6]), sum(logros_final[6:9]), sum(logros_final[9:12])]
        fig_tri = px.pie(names=tri, values=logros_tri, title="Distribución por Trimestre", hole=0.4)
        st.plotly_chart(fig_tri, use_container_width=True)

    with col2:
        # Semestres
        sem = ["1er Semestre", "2do Semestre"]
        logros_sem = [sum(logros_final[0:6]), sum(logros_final[6:12])]
        fig_sem = px.bar(x=sem, y=logros_sem, title="Consolidado por Semestre", color=sem, text=logros_sem)
        st.plotly_chart(fig_sem, use_container_width=True)

else:
    st.info("💡 Por favor, sube los archivos Excel para iniciar el análisis jerárquico.")
