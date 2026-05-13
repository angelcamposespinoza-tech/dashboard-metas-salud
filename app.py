import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Monitor de Metas Jurisdiccional", layout="wide")
st.title("📊 Monitor de Metas Jurisdicción nº1")

archivos = st.file_uploader("Subir archivos Excel", type=["xlsx"], accept_multiple_files=True)

def extraer_data_detallada(fila):
    """Extrae Meta, Logro y % para los 12 meses + avances trimestrales."""
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
    # --- AUDITORÍA CON MULTISELECCIÓN INTELIGENTE ---
    st.sidebar.divider()
    st.sidebar.subheader("⚙️ Configuración de Auditoría")
    
    # Lógica de exclusión de filtros
    trimestres = {
        "Trimestre 1 (Ene-Mar)": ["Ene", "Feb", "Mar"],
        "Trimestre 2 (Abr-Jun)": ["Abr", "May", "Jun"],
        "Trimestre 3 (Jul-Sep)": ["Jul", "Ago", "Sep"],
        "Trimestre 4 (Oct-Dic)": ["Oct", "Nov", "Dic"]
    }
    
    # Definir opciones disponibles dinámicamente
    opciones_base = ["Todos los meses", "Trimestre 1 (Ene-Mar)", "Trimestre 2 (Abr-Jun)", 
                     "Trimestre 3 (Jul-Sep)", "Trimestre 4 (Oct-Dic)", "Ene", "Feb", 
                     "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
    
    # Control de estado de selección
    if "periodo_sel" not in st.session_state:
        st.session_state.periodo_sel = []

    # Filtrar opciones para evitar redundancia
    opciones_mostrar = opciones_base.copy()
    if "Todos los meses" in st.session_state.periodo_sel:
        opciones_mostrar = ["Todos los meses"]
    else:
        for trim, meses in trimestres.items():
            if trim in st.session_state.periodo_sel:
                # Si el trimestre está seleccionado, quitar sus meses individuales
                for m in meses:
                    if m in opciones_mostrar: opciones_mostrar.remove(m)
            else:
                # Si algún mes de ese trimestre está seleccionado, quitar la opción del trimestre
                if any(m in st.session_state.periodo_sel for m in meses):
                    if trim in opciones_mostrar: opciones_mostrar.remove(trim)
                    
    periodo_sel = st.sidebar.multiselect(
        "Seleccione periodos a auditar:", 
        opciones_mostrar, 
        key="periodo_sel",
        help="Si elige un trimestre, no podrá elegir sus meses por separado."
    )

    if st.sidebar.button("🔍 Mostrar rubros de logro no registrados"):
        if not periodo_sel:
            st.error("Por favor seleccione al menos un periodo.")
        else:
            st.header(f"⚠️ Reporte de Omisiones")
            reporte_omisiones = []
            meses_columnas = {3: 'Ene', 6: 'Feb', 9: 'Mar', 15: 'Abr', 18: 'May', 21: 'Jun', 
                              27: 'Jul', 30: 'Ago', 33: 'Sep', 39: 'Oct', 42: 'Nov', 45: 'Dic'}

            # Mapear selección a lista de meses reales
            meses_validar = []
            if "Todos los meses" in periodo_sel:
                meses_validar = list(meses_columnas.values())
            else:
                for p in periodo_sel:
                    if p in trimestres:
                        meses_validar.extend(trimestres[p])
                    else:
                        meses_validar.append(p)
            
            for arc in archivos:
                dict_temp = pd.read_excel(arc, sheet_name=None, header=None)
                for nombre_hoja, df_hoja in dict_temp.items():
                    if df_hoja.shape[1] >= 46:
                        datos_reales = df_hoja.iloc[10:].dropna(subset=[0])
                        for _, fila in datos_reales.iterrows():
                            indicador_nombre = str(fila.iloc[0]).strip()
                            if len(indicador_nombre) > 5 and indicador_nombre.upper() not in ['N/A', 'NAN', 'SERVICIOS DE SALUD']:
                                for col_logro, mes_nom in meses_columnas.items():
                                    if mes_nom in meses_validar:
                                        col_meta = col_logro - 1
                                        try:
                                            m_v = float(fila.iloc[col_meta]) if pd.notna(fila.iloc[col_meta]) else 0
                                            l_v = float(fila.iloc[col_logro]) if pd.notna(fila.iloc[col_logro]) else 0
                                            if m_v > 0 and (l_v == 0 or pd.isna(l_v)):
                                                reporte_omisiones.append({
                                                    "Municipio": arc.name.replace(".xlsx", ""),
                                                    "Unidad/Pestaña": nombre_hoja,
                                                    "Indicador": indicador_nombre,
                                                    "Mes Pendiente": mes_nom
                                                })
                                        except: continue

            if reporte_omisiones:
                df_om = pd.DataFrame(reporte_omisiones)
                st.warning(f"Se encontraron {len(df_om)} registros pendientes.")
                st.dataframe(df_om, use_container_width=True)
                csv = df_om.to_csv(index=False).encode('utf-8')
                st.download_button("📥 Descargar Reporte (CSV)", csv, "omisiones_multiples.csv", "text/csv")
            else:
                st.success("✅ No se encontraron omisiones en los periodos seleccionados.")

    # --- VISUALIZADOR INDIVIDUAL ---
    st.sidebar.divider()
    nombres_archivos = [arc.name for arc in archivos]
    dep_sel = st.sidebar.selectbox("1. Seleccione Municipio:", nombres_archivos)
    archivo_obj = next(arc for arc in archivos if arc.name == dep_sel)
    dict_hojas = pd.read_excel(archivo_obj, sheet_name=None, header=None)
    
    servicios_limpios = set()
    for h in dict_hojas.values():
        if h.shape[1] > 0:
            df_temp = h.iloc[10:, 0].dropna()
            for s in df_temp.unique():
                n_s = str(s).strip()
                if n_s.upper() not in ['N/A', 'SERVICIOS DE SALUD', 'TRATO DIGNO', 'NAN']:
                    if len(n_s) > 5: servicios_limpios.add(n_s)
    
    lista_ord = sorted(list(servicios_limpios))
    op_hojas = ["CONSOLIDADO MUNICIPAL"] + list(dict_hojas.keys())
    sede_sel = st.sidebar.selectbox("2. Seleccione Unidad Médica:", op_hojas)
    indicador = st.selectbox("3. Seleccione el Indicador:", lista_ord)

    if sede_sel == "CONSOLIDADO MUNICIPAL":
        metas_f, logros_f = [0]*16, [0]*16
        nombres_periodos = []
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
        else: nombres_periodos = []

    if nombres_periodos:
        df_total = pd.DataFrame({'Periodo': nombres_periodos, 'Meta': metas_f, 'Logro': logros_f, 'Pct': pcts_f})
        df_meses = df_total[~df_total['Periodo'].str.contains('Avance')]
        st.divider()
        st.header(f"📍 {indicador}")
        fig = px.bar(df_meses, x='Periodo', y=['Meta', 'Logro'], barmode='group',
                     title=f"Cumplimiento: {indicador}", color_discrete_map={'Meta': '#1f77b4', 'Logro': '#d62728'})
        st.plotly_chart(fig, use_container_width=True)
        with st.expander("🔍 Ver Tabla Detallada"):
            st.table(df_total.set_index('Periodo'))
else:
    st.info("Sube los archivos Excel para activar la auditoría.")
