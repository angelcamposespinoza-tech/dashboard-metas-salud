# --- 1. INSTALACIÓN DE LIBRERÍAS (INCLUYENDO XLSXWRITER) ---
!pip install streamlit pandas openpyxl plotly pyngrok xlsxwriter --quiet

# --- 2. CONFIGURACIÓN DE NGROK (PON TU TOKEN AQUÍ) ---
from pyngrok import ngrok
import os
# Reemplaza con tu token real de ngrok
token = "TU_TOKEN_AQUÍ" 
ngrok.set_auth_token(token)

# --- 3. CREACIÓN DE LA APP ---
%%writefile app.py
import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

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
    st.sidebar.subheader("⚙️ Auditoría de Carga")
    
    trimestres = {
        "Trimestre 1 (Ene-Mar)": ["Ene", "Feb", "Mar"],
        "Trimestre 2 (Abr-Jun)": ["Abr", "May", "Jun"],
        "Trimestre 3 (Jul-Sep)": ["Jul", "Ago", "Sep"],
        "Trimestre 4 (Oct-Dic)": ["Oct", "Nov", "Dic"]
    }
    
    opciones_base = ["Todos los meses", "Trimestre 1 (Ene-Mar)", "Trimestre 2 (Abr-Jun)", 
                     "Trimestre 3 (Jul-Sep)", "Trimestre 4 (Oct-Dic)", "Ene", "Feb", 
                     "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
    
    if "periodo_sel" not in st.session_state:
        st.session_state.periodo_sel = []

    opciones_mostrar = opciones_base.copy()
    if "Todos los meses" in st.session_state.periodo_sel:
        opciones_mostrar = ["Todos los meses"]
    else:
        for trim, meses in trimestres.items():
            if trim in st.session_state.periodo_sel:
                for m in meses:
                    if m in opciones_mostrar: opciones_mostrar.remove(m)
            elif any(m in st.session_state.periodo_sel for m in meses):
                if trim in opciones_mostrar: opciones_mostrar.remove(trim)
                    
    periodo_sel = st.sidebar.multiselect("Periodos a auditar:", opciones_mostrar, key="periodo_sel")

    if st.sidebar.button("🔍 Generar Reporte de Omisiones"):
        if not periodo_sel:
            st.error("Seleccione un periodo.")
        else:
            st.header(f"📋 Resumen de Unidades con Información Pendiente")
            reporte_detallado = []
            meses_columnas = {3: 'Ene', 6: 'Feb', 9: 'Mar', 15: 'Abr', 18: 'May', 21: 'Jun', 
                              27: 'Jul', 30: 'Ago', 33: 'Sep', 39: 'Oct', 42: 'Nov', 45: 'Dic'}

            meses_validar = []
            if "Todos los meses" in periodo_sel:
                meses_validar = list(meses_columnas.values())
            else:
                for p in periodo_sel:
                    if p in trimestres: meses_validar.extend(trimestres[p])
                    else: meses_validar.append(p)
            
            for arc in archivos:
                dict_temp = pd.read_excel(arc, sheet_name=None, header=None)
                for nombre_hoja, df_hoja in dict_temp.items():
                    if not df_hoja.empty and df_hoja.shape[1] >= 46:
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
                                                reporte_detallado.append({
                                                    "MUNICIPIO": arc.name.replace(".xlsx", "").upper(),
                                                    "UNIDAD_MEDICA": nombre_hoja.upper(),
                                                    "MES_FALTANTE": mes_nom.upper()
                                                })
                                        except: continue

            if reporte_detallado:
                df_detallado = pd.DataFrame(reporte_detallado)
                unidades_pendientes = df_detallado.drop_duplicates(subset=["MUNICIPIO", "UNIDAD_MEDICA"])
                unidades_pendientes = unidades_pendientes[["MUNICIPIO", "UNIDAD_MEDICA"]].reset_index(drop=True)
                
                st.warning(f"Se detectaron {len(unidades_pendientes)} unidades con carga pendiente.")
                st.table(unidades_pendientes)

                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    unidades_pendientes.to_excel(writer, index=False, sheet_name='Unidades_Pendientes')
                    df_detallado.to_excel(writer, index=False, sheet_name='Detalle_por_Indicador')
                
                st.download_button(label="📥 Descargar Reporte (Excel)", data=output.getvalue(),
                                   file_name="reporte_pendientes.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.success("✅ Carga completa detectada.")

    # --- VISUALIZADOR DE GRÁFICAS ---
    st.sidebar.divider()
    if archivos:
        dep_sel = st.sidebar.selectbox("1. Municipio para Gráficas:", [arc.name for arc in archivos])
        archivo_obj = next(arc for arc in archivos if arc.name == dep_sel)
        dict_hojas = pd.read_excel(archivo_obj, sheet_name=None, header=None)
        
        servicios = set()
        for h in dict_hojas.values():
            if not h.empty and h.shape[1] > 0:
                for s in h.iloc[10:, 0].dropna().unique():
                    if len(str(s)) > 5 and str(s).upper() not in ['N/A', 'SERVICIOS DE SALUD', 'NAN']:
                        servicios.add(str(s).strip())
        
        if servicios:
            indicador = st.selectbox("🎯 Buscar Indicador:", sorted(list(servicios)))
            sede_sel = st.sidebar.selectbox("2. Unidad Médica:", ["CONSOLIDADO"] + list(dict_hojas.keys()))

            nombres_periodos, metas_f, logros_f, pcts_f = [], [0]*16, [0]*16, [0]*16
            if sede_sel == "CONSOLIDADO":
                for h in dict_hojas.values():
                    if not h.empty and h.shape[1] > 0:
                        fila = h[h.iloc[:, 0].astype(str).str.strip() == indicador]
                        if not fila.empty:
                            nombres_periodos, m, l, p = extraer_data_detallada(fila.iloc[0])
                            metas_f = [x+y for x,y in zip(metas_f, m)]
                            logros_f = [x+y for x,y in zip(logros_f, l)]
                pcts_f = [round((l/m*100),1) if m > 0 else 0 for l,m in zip(logros_f, metas_f)]
            else:
                hoja = dict_hojas[sede_sel]
                if not hoja.empty and hoja.shape[1] > 0:
                    fila = hoja[hoja.iloc[:, 0].astype(str).str.strip() == indicador]
                    if not fila.empty:
                        nombres_periodos, metas_f, logros_f, pcts_f = extraer_data_detallada(fila.iloc[0])

            if nombres_periodos:
                df_plot = pd.DataFrame({'Periodo': nombres_periodos, 'Meta': metas_f, 'Logro': logros_f})
                df_meses = df_plot[~df_plot['Periodo'].str.contains('Avance')]
                st.plotly_chart(px.bar(df_meses, x='Periodo', y=['Meta', 'Logro'], barmode='group', title=f"Desempeño: {indicador}"), use_container_width=True)
else:
    st.info("Sube los archivos para comenzar.")

# --- LANZAMIENTO ---
os.system("pkill streamlit")
tunnnel = ngrok.connect(8501)
print(f"\n✅ ACCEDE AQUÍ:\n{tunnnel.public_url}\n")
!streamlit run app.py --server.port 8501 --server.headless True > /dev/null
