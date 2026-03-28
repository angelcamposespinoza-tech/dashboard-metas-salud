import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Monitor de Metas Jurisdiccional", layout="wide")
st.title("📊 Monitor de Metas con Porcentajes Reales del Excel")

archivos = st.file_uploader("Subir archivos Excel", type=["xlsx"], accept_multiple_files=True)

def extraer_data_detallada(fila):
    """
    Extrae datos siguiendo el mapa exacto de tus columnas:
    Ene: Meta(2), Logro(3), % (Col E -> 4)
    Feb: Meta(5), Logro(6), % (Col H -> 7)
    Mar: Meta(8), Logro(9), % (Col K -> 10)
    T1:  Meta(11), Logro(12), % (Col N -> 13) ... y así sucesivamente.
    """
    indices_pct = [4, 7, 10, 13, 16, 19, 22, 25, 28, 31, 34, 37, 40, 43, 46, 49]
    indices_logro = [3, 6, 9, 12, 15, 18, 21, 24, 27, 30, 33, 36, 39, 42, 45, 48]
    nombres = ['Ene', 'Feb', 'Mar', 'Avance T1', 'Abr', 'May', 'Jun', 'Avance T2', 
               'Jul', 'Ago', 'Sep', 'Avance T3', 'Oct', 'Nov', 'Dic', 'Avance T4']
    
    logros, pcts = [], []
    
    for i in range(len(indices_pct)):
        try:
            l_idx = indices_logro[i]
            p_idx = indices_pct[i]
            
            val_logro = fila.iloc[l_idx]
            val_pct = fila.iloc[p_idx]
            
            # Limpieza de Logro
            logro = float(val_logro) if pd.notna(val_logro) and str(val_logro).upper() != 'N/A' else 0
            
            # Limpieza de Porcentaje (Si es 1.696 lo convierte a 169.6)
            if pd.notna(val_pct) and str(val_pct).upper() != 'N/A':
                pct = round(float(val_pct) * 100, 1)
            else:
                pct = 0.0
                
            logros.append(logro)
            pcts.append(pct)
        except:
            logros.append(0); pcts.append(0)
            
    return nombres, logros, pcts

if archivos:
    nombres_archivos = [arc.name for arc in archivos]
    dep_sel = st.sidebar.selectbox("Seleccione Dependencia:", nombres_archivos)
    archivo_obj = next(arc for arc in archivos if arc.name == dep_sel)
    
    dict_hojas = pd.read_excel(archivo_obj, sheet_name=None, header=None)
    opciones_hojas = ["CONSOLIDADO MUNICIPAL"] + list(dict_hojas.keys())
    sede_sel = st.sidebar.selectbox("Seleccione Unidad:", opciones_hojas)

    servicios = set()
    for h in dict_hojas.values():
        if h.shape[1] > 0:
            for s in h.iloc[10:, 0].dropna().unique():
                nombre_s = str(s).strip()
                if nombre_s.upper() not in ['N/A', 'SERVICIOS DE SALUD', 'TRATO DIGNO', 'NAN']:
                    servicios.add(nombre_s)

    indicador = st.selectbox("Seleccione Indicador:", sorted(list(servicios)))

    nombres_periodos, logros_f, pcts_f = [], [0]*16, [0]*16

    if sede_sel == "CONSOLIDADO MUNICIPAL":
        cont_hojas = 0
        for h in dict_hojas.values():
            if h.shape[1] > 0:
                fila = h[h.iloc[:, 0].astype(str).str.strip() == indicador]
                if not fila.empty:
                    nombres_periodos, l, p = extraer_data_detallada(fila.iloc[0])
                    logros_f = [x + y for x, y in zip(logros_f, l)]
                    pcts_f = [x + y for x, y in zip(pcts_f, p)]
                    cont_hojas += 1
        if cont_hojas > 0:
            pcts_f = [round(p / cont_hojas, 1) for p in pcts_f]
    else:
        hoja = dict_hojas[sede_sel]
        fila = hoja[hoja.iloc[:, 0].astype(str).str.strip() == indicador]
        if not fila.empty:
            nombres_periodos, logros_f, pcts_f = extraer_data_detallada(fila.iloc[0])

    if nombres_periodos:
        df_total = pd.DataFrame({'Periodo': nombres_periodos, 'Logro': logros_f, 'Pct': pcts_f})
        df_meses = df_total[~df_total['Periodo'].str.contains('Avance')]
        df_trim = df_total[df_total['Periodo'].str.contains('Avance')]

        st.divider()
        st.subheader(f"📌 Seguimiento: {indicador}")
        
        fig_m = px.bar(df_meses, x='Periodo', y='Logro', text=[f"{p}%" for p in df_meses['Pct']],
                       color='Pct', color_continuous_scale='RdYlGn',
                       title="Logros Mensuales (Porcentaje real de columnas E, H, K...)")
        fig_m.update_traces(textposition='outside')
        st.plotly_chart(fig_m, use_container_width=True)

        st.subheader("🏁 Avances de Trimestre (Celdas de Porcentaje Reales)")
        fig_t = px.bar(df_trim, x='Periodo', y='Pct', text=[f"{p}%" for p in df_trim['Pct']],
                       labels={'Pct': '% de Logro Real'}, color_discrete_sequence=['#1f77b4'])
        
        # Corrección del error de layout
        fig_t.update_layout(yaxis_title="Porcentaje de Logro (%)")
        fig_t.update_traces(textposition='outside')
        
        st.plotly_chart(fig_t, use_container_width=True)
    else:
        st.warning("No se encontraron datos para mostrar.")
else:
    st.info("Sube los archivos Excel para visualizar los porcentajes de las columnas E, H, K...")
