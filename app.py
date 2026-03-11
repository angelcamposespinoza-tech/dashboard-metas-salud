import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Monitor de Rendimiento", layout="wide")
st.title("📊 Análisis de Logros y Porcentajes de Cumplimiento")

archivos = st.file_uploader("Subir Excels de Unidades", type=["xlsx"], accept_multiple_files=True)

if archivos:
    datos_unidades = {arc.name: pd.read_excel(arc, header=None) for arc in archivos}
    unidad = st.sidebar.selectbox("Unidad Médica", list(datos_unidades.keys()))
    df = datos_unidades[unidad].iloc[10:].copy()
    
    indicador = st.selectbox("Seleccione el Servicio:", df.iloc[:, 0].dropna().unique())
    fila = df[df.iloc[:, 0] == indicador]

    if not fila.empty:
        meses = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic']
        logros = []
        porcentajes = []
        
        # El primer Logro (Enero) está en la columna 3, y su % en la columna 4
        col_logro = 3
        for m in meses:
            val_meta = fila.iloc[0, col_logro - 1]
            val_logro = fila.iloc[0, col_logro]
            
            # Limpieza de datos
            meta = float(val_meta) if pd.notna(val_meta) and str(val_meta).upper() != 'N/A' else 0
            logro = float(val_logro) if pd.notna(val_logro) and str(val_logro).upper() != 'N/A' else 0
            
            # Cálculo del porcentaje real
            pct = (logro / meta * 100) if meta > 0 else 0
            
            logros.append(logro)
            porcentajes.append(round(pct, 1))
            col_logro += 3 # Saltar al siguiente mes

        # --- VISUALIZACIÓN DE PORCENTAJES EN TARJETAS ---
        st.subheader(f"📈 Porcentaje de Cumplimiento Mensual: {indicador}")
        cols = st.columns(6) # 6 columnas para los primeros 6 meses
        for i in range(6):
            color = "normal" if porcentajes[i] >= 100 else "inverse"
            cols[i].metric(meses[i], f"{logros[i]}", f"{porcentajes[i]}%")

        cols2 = st.columns(6) # 6 columnas para los siguientes 6 meses
        for i in range(6, 12):
            cols2[i-6].metric(meses[i], f"{logros[i]}", f"{porcentajes[i]}%")

        # --- GRÁFICA DE LOGROS CON ETIQUETAS ---
        df_grafica = pd.DataFrame({'Mes': meses, 'Logro': logros, 'Porcentaje': porcentajes})
        
        fig = px.bar(df_grafica, x='Mes', y='Logro', text='Porcentaje',
                     title=f"Logros Mensuales (El número sobre la barra es el % de cumplimiento)",
                     labels={'Logro': 'Cantidad Realizada', 'text': '% Cumplimiento'},
                     color='Logro', color_continuous_scale='Viridis')
        
        fig.update_traces(texttemplate='%{text}%', textposition='outside')
        st.plotly_chart(fig, use_container_width=True)

else:
    st.info("Sube los archivos para calcular los porcentajes de éxito.")