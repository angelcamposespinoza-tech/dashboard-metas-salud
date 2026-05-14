import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from fpdf import FPDF

st.set_page_config(page_title="Monitor de Metas Jurisdiccional", layout="wide")
st.title("📊 Monitor de Metas Jurisdicción nº1")

archivos = st.file_uploader("Subir archivos Excel", type=["xlsx"], accept_multiple_files=True)

def extraer_data_detallada(fila):
    """Extrae Meta y Logro por columna para las gráficas."""
    indices_logro = [3, 6, 9, 12, 15, 18, 21, 24, 27, 30, 33, 36]
    indices_meta = [i-1 for i in indices_logro]
    nombres = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic']
    ms, ls = [], []
    for i in range(len(indices_logro)):
        try:
            m = float(fila.iloc[indices_meta[i]]) if pd.notna(fila.iloc[indices_meta[i]]) else 0
            l = float(fila.iloc[indices_logro[i]]) if pd.notna(fila.iloc[indices_logro[i]]) else 0
            ms.append(m); ls.append(l)
        except:
            ms.append(0); ls.append(0)
    return nombres, ms, ls

def generar_pdf(df_pendientes, periodos_txt):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 16)
    pdf.cell(200, 10, txt="REPORTE DE UNIDADES PENDIENTES DE CARGA", ln=True, align='C')
    
    pdf.set_font("Arial", "I", 12)
    pdf.cell(200, 10, txt=f"Meses auditados: {periodos_txt}", ln=True, align='C')
    
    # Añadimos el conteo total al PDF
    pdf.set_font("Arial", "B", 12)
    pdf.set_text_color(200, 0, 0) # Rojo para resaltar
    pdf.cell(200, 10, txt=f"TOTAL DE UNIDADES PENDIENTES: {len(df_pendientes.drop_duplicates(subset=['UNIDAD_MEDICA']))}", ln=True, align='C')
    pdf.set_text_color(0, 0, 0) # Regresamos a negro
    pdf.ln(5)
    
    # Tabla
    pdf.set_fill_color(230, 230, 230)
    pdf.set_font("Arial", "B", 10)
    pdf.cell(50, 10, "MUNICIPIO", 1, 0, 'C', True)
    pdf.cell(80, 10, "UNIDAD MEDICA", 1, 0, 'C', True)
    pdf.cell(60, 10, "MES PENDIENTE", 1, 1, 'C', True)
    
    pdf.set_font("Arial", "", 9)
    for _, row in df_pendientes.iterrows():
        pdf.cell(50, 10, str(row['MUNICIPIO']), 1)
        pdf.cell(80, 10, str(row['UNIDAD_MEDICA']), 1)
        pdf.cell(60, 10, str(row['MES']), 1)
        pdf.ln()
        
    return pdf.output(dest='S').encode('latin-1', 'replace')

if archivos:
    st.sidebar.divider()
    st.sidebar.subheader("⚙️ Auditoría de Carga Mensual")
    
    opciones_meses = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
    meses_sel = st.sidebar.multiselect("Seleccione meses a auditar:", opciones_meses)

    mes_a_col = {'Ene':3, 'Feb':6, 'Mar':9, 'Abr':12, 'May':15, 'Jun':18, 'Jul':21, 'Ago':24, 'Sep':27, 'Oct':30, 'Nov':33, 'Dic':36}

    if st.sidebar.button("🔍 Verificar carga por mes"):
        if not meses_sel:
            st.error("Por favor seleccione al menos un mes.")
        else:
            reporte_omisiones = []
            
            for arc in archivos:
                dict_temp = pd.read_excel(arc, sheet_name=None, header=None)
                for nombre_hoja, df_hoja in dict_temp.items():
                    if not df_hoja.empty and df_hoja.shape[1] >= 37:
                        for m_audit in meses_sel:
                            idx = mes_a_col[m_audit]
                            columna_logro = df_hoja.iloc[:, idx]
                            tiene_datos = pd.to_numeric(columna_logro, errors='coerce').notnull().any()
                            
                            if not tiene_datos:
                                reporte_omisiones.append({
                                    "MUNICIPIO": arc.name.replace(".xlsx", "").upper(),
                                    "UNIDAD_MEDICA": nombre_hoja.upper(),
                                    "MES": m_audit.upper()
                                })

            if reporte_omisiones:
                st.header(f"📋 Reporte de Faltantes")
                df_final = pd.DataFrame(reporte_omisiones)
                
                # Métrica visual en la app
                total_unidades = len(df_final.drop_duplicates(subset=['UNIDAD_MEDICA']))
                st.metric(label="Unidades con faltantes detectadas", value=total_unidades)
                
                st.table(df_final)
                
                pdf_bytes = generar_pdf(df_final, ", ".join(meses_sel))
                st.download_button(f"📥 Descargar Reporte PDF ({total_unidades} unidades)", pdf_bytes, "reporte_mensual.pdf", "application/pdf")
            else:
                st.success("✅ Carga completa: Todas las unidades tienen datos en los meses seleccionados.")

    # --- VISUALIZADOR ---
    st.sidebar.divider()
    mun_sel = st.sidebar.selectbox("Ver Gráficas de:", [a.name for a in archivos])
    arc_obj = next(a for a in archivos if a.name == mun_sel)
    dict_h = pd.read_excel(arc_obj, sheet_name=None, header=None)
    
    servs = set()
    for h in dict_h.values():
        if not h.empty and h.shape[1] > 0:
            for s in h.iloc[10:, 0].dropna().unique():
                if len(str(s)) > 5: servs.add(str(s).strip())
    
    ind_sel = st.selectbox("🎯 Indicador:", sorted(list(servs)))
    sede = st.sidebar.selectbox("Unidad Médica:", ["CONSOLIDADO"] + list(dict_h.keys()))
    
    if sede == "CONSOLIDADO":
        m_t, l_t = [0]*12, [0]*12
        nom_p = []
        for h in dict_h.values():
            if not h.empty and h.shape[1] > 0:
                f = h[h.iloc[:, 0].astype(str).str.strip() == ind_sel]
                if not f.empty:
                    nom_p, m, l = extraer_data_detallada(f.iloc[0])
                    m_t = [x+y for x,y in zip(m_t, m)]
                    l_t = [x+y for x,y in zip(l_t, l)]
        df_g = pd.DataFrame({'Mes': nom_p, 'Meta': m_t, 'Logro': l_t})
    else:
        hoja = dict_h[sede]
        f = hoja[hoja.iloc[:, 0].astype(str).str.strip() == ind_sel] if not hoja.empty else pd.DataFrame()
        if not f.empty:
            nom_p, m, l = extraer_data_detallada(f.iloc[0])
            df_g = pd.DataFrame({'Mes': nom_p, 'Meta': m, 'Logro': l})
        else: df_g = pd.DataFrame()

    if not df_g.empty:
        st.plotly_chart(px.bar(df_g, x='Mes', y=['Meta', 'Logro'], barmode='group'), use_container_width=True)
else:
    st.info("Sube archivos para iniciar.")
