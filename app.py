import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Monitor de Metas Jurisdiccional", layout="wide")
st.title("📊 Monitor de Metas Jurisdicción nº1")

archivos = st.file_uploader("Subir archivos Excel", type=["xlsx"], accept_multiple_files=True)

def extraer_data_detallada(fila):
    """
    Extrae Meta, Logro y % de las columnas exactas (E, H, K, N...).
    """
    indices_meta = [2, 5, 8, 11, 14, 17, 20, 23, 26, 29, 32, 35, 38, 41, 44, 47]
    indices_logro = [3, 6, 9, 12, 15, 18, 21, 24, 27, 30, 33, 36, 39, 42, 45, 48]
    indices_pct = [4, 7, 10, 13, 16, 19, 22, 25, 28, 31, 34, 37, 40, 43, 46, 49]
    nombres = ['Ene', 'Feb', 'Mar', 'Avance T1', 'Abr', 'May', 'Jun', 'Avance T2', 
               'Jul', 'Ago', 'Sep', 'Avance T3', 'Oct', 'Nov', 'Dic', 'Avance T4']
    
    ms, ls, ps = [], [], []
    for i in range(len(indices_pct)):
        try:
            m = float(fila.iloc[indices_meta[i]]) if pd.notna(fila.iloc[indices_meta[i]]) else 0
            l = float(fila.iloc[indices_logro[i]]) if pd.notna(fila.iloc[indices_logro[i]]) else 0
            p = float(fila.iloc[indices_pct[i]]) if pd.notna(fila.iloc[indices_pct[i]]) else 0
            ms.append(m); ls.append(l); ps.append(round(p * 100, 1))
        except:
            ms.append(0); ls.append(0); ps.append(0)
    return nombres, ms, ls, ps

if archivos:
    # --- NUEVA FUNCIÓN: AUDITORÍA DE CARGA ---
    st.sidebar.divider()
    if st.sidebar.button("🔍 Mostrar rubros sin logros cargados"):
        st.header("⚠️ Reporte de Omisiones (Falta de Carga)")
        st.write("Unidades que tienen Meta programada pero Logro en 0 o vacío.")
        
        reporte_omisiones = []
        meses_columnas = {3: 'Ene', 6: 'Feb', 9: 'Mar', 15: 'Abr', 18: 'May', 21: 'Jun'} # Columnas de Logro

        for arc in archivos:
            dict_temp = pd.read_excel(arc, sheet_name=None, header=None)
            for nombre_hoja, df_hoja in dict_temp.items():
                if df_hoja.shape[1] > 10:
                    # Analizamos desde la fila 11
                    datos_reales = df_hoja.iloc[10:].dropna(subset=[0])
                    for _, fila in datos_reales.iterrows():
                        indicador_nombre = str(fila.iloc[0]).strip()
                        # Solo auditar indicadores válidos
                        if len(indicador_nombre) > 5 and indicador_nombre.upper() not in ['N/A', 'NAN']:
                            # Revisar cada mes (Meta > 0 y Logro == 0)
                            for col_logro, mes_nom in meses_columnas.items():
                                col_meta = col_logro - 1
                                meta_val = fila.iloc[col_meta] if pd.notna(fila.iloc[col_meta]) else 0
                                logro_val = fila.iloc[col_logro] if pd.notna(fila.iloc[col_logro]) else 0
                                
                                if meta_val > 0 and (logro_val == 0 or pd.isna(logro_val)):
                                    reporte_omisiones.append({
                                        "Municipio": arc.name.replace(".xlsx", ""),
                                        "Jurisdicción/Unidad": nombre_hoja,
                                        "Indicador": indicador_nombre,
                                        "Mes Faltante": mes_nom
                                    })
        
        if reporte_omisiones:
            df_omisiones = pd.DataFrame(reporte_omisiones)
            st.warning(f"Se encontraron {len(df_omisiones)} rubros sin información de logro.")
            st.dataframe(df_omisiones, use_container_width=True)
        else:
            st.success("✅ ¡Felicidades! Todos los rubros con metas tienen logros cargados.")
    
    # --- LÓGICA ORIGINAL DE LA APP ---
    nombres_archivos = [arc.name for arc in archivos]
    dep_sel = st.sidebar.selectbox("1. Seleccione Archivo:", nombres_archivos)
    archivo_obj = next(arc for arc in archivos if arc.name == dep_sel)
    
    dict_hojas = pd.read_excel(archivo_obj, sheet_name=None, header=None)
    
    servicios_limpios = set()
    for h in dict_hojas.values():
        if h.shape[1] > 0:
            df_temp = h.iloc[10:, 0].dropna()
            for s in df_temp.unique():
                nombre_s = str(s).strip()
                if nombre_s.upper() not in ['N/A', 'SERVICIOS DE SALUD', 'TRATO DIGNO', 'NAN', 'FORTALECIMIENTO A LA ATENCIÓN MÉDICA']:
                    if len(nombre_s) > 5:
                        servicios_limpios.add(nombre_s)
    
    lista_ordenada = sorted(list(servicios_limpios))
    opciones_hojas = ["CONSOLIDADO MUNICIPAL (Suma total)"] + list(dict_hojas.keys())
    sede_sel = st.sidebar.selectbox("2. Seleccione Unidad Médica:", opciones_hojas)
    indicador = st.selectbox("3. Busque y Seleccione el Indicador:", lista_ordenada)

    nombres_periodos, metas_f, logros_f, pcts_f = [], [0]*16, [0]*16, [0]*16

    if sede_sel == "CONSOLIDADO MUNICIPAL (Suma total)":
        for h in dict_hojas.values():
            if h.shape[1] > 0:
                fila = h[h.iloc[:, 0].astype(str).str.strip() == indicador]
                if not fila.empty:
                    nombres_periodos, m, l, p = extraer_data_detallada(fila.iloc[0])
                    metas_f = [x + y for x, y in zip(metas_f, m)]
                    logros_f = [x + y for x, y in zip(logros_f, l)]
        pcts_f = [round((l/m*100),1) if m > 0 else 0 for l, m in zip(logros_f, metas_f)]
    else:
        hoja = dict_hojas[sede_sel]
        fila = hoja[hoja.iloc[:, 0].astype(str).str.strip() == indicador]
        if not fila.empty:
            nombres_periodos, metas_f, logros_f, pcts_f = extraer_data_detallada(fila.iloc[0])

    if nombres_periodos:
        df_total = pd.DataFrame({'Periodo': nombres_periodos, 'Meta': metas_f, 'Logro': logros_f, 'Pct': pcts_f})
        df_meses = df_total[~df_total['Periodo'].str.contains('Avance')]
        df_trim = df_total[df_total['Periodo'].str.contains('Avance')]

        st.divider()
        st.header(f"📍 {indicador}")
        
        fig_m = px.bar(df_meses, x='Periodo', y='Logro', text=[f"{p}%" for p in df_meses['Pct']],
                       color='Pct', color_continuous_scale='RdYlGn',
                       title="Cumplimiento Mensual (Ene - Dic)")
        fig_m.update_traces(textposition='outside')
        st.plotly_chart(fig_m, use_container_width=True)

        st.subheader("🏁 Resumen de Avances Trimestrales")
        fig_t = px.bar(df_trim, x='Periodo', y='Pct', text=[f"{p}%" for p in df_trim['Pct']],
                       color_discrete_sequence=['#2E86C1'])
        fig_t.update_layout(yaxis_title="Porcentaje (%)", yaxis_range=[0, max(df_trim['Pct'].max()+20, 120)])
        fig_t.update_traces(textposition='outside')
        st.plotly_chart(fig_t, use_container_width=True)

        with st.expander("🔍 Ver Tabla de Datos"):
            st.table(df_total.set_index('Periodo'))
    else:
        st.warning("No hay datos disponibles.")
else:
    st.info("Sube los archivos Excel para activar el monitor.")
