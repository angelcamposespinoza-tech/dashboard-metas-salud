import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Monitor de Metas Jurisdiccional", layout="wide")
st.title("📊 Monitor de Metas con Avance Trimestral")

archivos = st.file_uploader("Subir archivos Excel", type=["xlsx"], accept_multiple_files=True)

def extraer_data_detallada(fila):
    """
    Extrae datos siguiendo el mapa de tu Excel:
    Meses: Jan(2,3,4), Feb(5,6,7), Mar(8,9,10) -> T1(11,12,13)
    Apr(14,15,16), May(17,18,19), Jun(20,21,22) -> T2(23,24,25) ...
    """
    # Índices de las columnas Meta de cada periodo
    indices = [2, 5, 8, 11, 14, 17, 20, 23, 26, 29, 32, 35, 38, 41, 44, 47]
    nombres = ['Ene', 'Feb', 'Mar', 'Trimestre 1', 'Abr', 'May', 'Jun', 'Trimestre 2', 
               'Jul', 'Ago', 'Sep', 'Trimestre 3', 'Oct', 'Nov', 'Dic', 'Trimestre 4']
    
    metas, logros, pcts = [], [], []
    
    for idx in indices:
        try:
            m = fila.iloc[idx]
            l = fila.iloc[idx+1]
            p = fila.iloc[idx+2] # Celda de porcentaje del Excel
            
            # Limpieza y conversión
            m = float(m) if pd.notna(m) and str(m).upper() != 'N/A' else 0
            l = float(l) if pd.notna(l) and str(l).upper() != 'N/A' else 0
            
            # El porcentaje en tu Excel viene como ratio (ej. 0.824). 
            # Lo multiplicamos por 100 para la vista.
            p_val = float(p) * 100 if pd.notna(p) and str(p).upper() != 'N/A' else 0
            
            metas.append(m)
            logros.append(l)
            pcts.append(round(p_val, 1))
        except:
            metas.append(0); logros.append(0); pcts.append(0)
            
    return nombres, metas, logros, pcts

if archivos:
    nombres_archivos = [arc.name for arc in archivos]
    dep_sel = st.sidebar.selectbox("Seleccione Dependencia:", nombres_archivos)
    archivo_obj = next(arc for arc in archivos if arc.name == dep_sel)
    
    dict_hojas = pd.read_excel(archivo_obj, sheet_name=None, header=None)
    opciones_hojas = ["CONSOLIDADO MUNICIPAL"] + list(dict_hojas.keys())
    sede_sel = st.sidebar.selectbox("Seleccione Unidad:", opciones_hojas)

    # Identificar indicadores (desde fila 11, columna A)
    servicios = set()
    for h in dict_hojas.values():
        if h.shape[1] > 0:
            for s in h.iloc[10:, 0].dropna().unique():
                if str(s).strip().upper() not in ['N/A', 'SERVICIOS DE SALUD', 'TRATO DIGNO']:
                    servicios.add(str(s).strip())

    indicador = st.selectbox("Seleccione Indicador:", sorted(list(servicios)))

    # --- PROCESAMIENTO ---
    nombres_periodos, metas_f, logros_f, pcts_f = [], [0]*16, [0]*16, [0]*16

    if sede_sel == "CONSOLIDADO MUNICIPAL":
        for h in dict_hojas.values():
            if h.shape[1] > 0:
                fila = h[h.iloc[:, 0].astype(str).str.strip() == indicador]
                if not fila.empty:
                    nombres_periodos, m, l, p = extraer_data_detallada(fila.iloc[0])
                    metas_f = [x + y for x, y in zip(metas_f, m)]
                    logros_f = [x + y for x, y in zip(logros_f, l)]
        # En consolidado recalculamos % para que sea matemáticamente correcto
        pcts_f = [round((l/m*100),1) if m > 0 else 0 for l, m in zip(logros_f, metas_f)]
    else:
        hoja = dict_hojas[sede_sel]
        fila = hoja[hoja.iloc[:, 0].astype(str).str.strip() == indicador]
        if not fila.empty:
            nombres_periodos, metas_f, logros_f, pcts_f = extraer_data_detallada(fila.iloc[0])

    # --- VISUALIZACIÓN ---
    if nombres_periodos:
        # Separar Meses de Trimestres para las gráficas
        df_completo = pd.DataFrame({'Periodo': nombres_periodos, 'Logro': logros_f, 'Pct': pcts_f})
        df_meses = df_completo[~df_completo['Periodo'].str.contains('Trimestre')]
        df_trimestres = df_completo[df_completo['Periodo'].str.contains('Trimestre')]

        st.divider()
        st.subheader(f"📈 Avance Mensual: {indicador}")
        
        # Gráfica de Meses
        fig_m = px.bar(df_meses, x='Periodo', y='Logro', text=[f"{p}%" for p in df_meses['Pct']],
                       color='Pct', color_continuous_scale='RdYlGn', range_color=[0, 100],
                       title="Logro por Mes (Etiqueta superior es el % del Excel)")
        st.plotly_chart(fig_m, use_container_width=True)

        # Gráfica de Trimestres (Los avances que pidió no confundir con meses)
        st.subheader("🏁 Resumen por Trimestre (Columnas de Avance)")
        fig_t = px.bar(df_trimestres, x='Periodo', y='Pct', text=[f"{p}%" for p in df_trimestres['Pct']],
                       labels={'Pct': '% de Cumplimiento'}, color_discrete_sequence=['#2C3E50'])
        fig_t.update_layout(yaxis_range=[0, max(df_trimestres['Pct'].max() + 10, 110)])
        st.plotly_chart(fig_t, use_container_width=True)
    else:
        st.error("No se encontraron datos para el indicador seleccionado.")
else:
    st.info("Suba el archivo 'ALVARO OBREGON.xlsx' para ver la magia de los trimestres.")
