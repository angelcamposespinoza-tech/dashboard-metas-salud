import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from fpdf import FPDF

# Configuración de página
st.set_page_config(page_title="Monitor de Metas Jurisdiccional", layout="wide")
st.title("📊 Monitor de Metas Jurisdicción nº1")

# Subida de archivos
archivos = st.file_uploader("Subir archivos Excel", type=["xlsx"], accept_multiple_files=True)

def extraer_data_detallada(fila):
    """Extrae Meta y Logro para los 12 meses + avances trimestrales."""
    indices_meta = [2, 5, 8, 11, 14, 17, 20, 23, 26, 29, 32, 35, 38, 41, 44, 47]
    indices_logro = [3, 6, 9, 12, 15, 18, 21, 24, 27, 30, 33, 36, 39, 42, 45, 48]
    nombres = ['Ene', 'Feb', 'Mar', 'Avance T1', 'Abr', 'May', 'Jun', 'Avance T2', 
               'Jul', 'Ago', 'Sep', 'Avance T3', 'Oct', 'Nov', 'Dic', 'Avance T4']
    ms, ls = [], []
    for i in range(len(indices_meta)):
        try:
            m = float(fila.iloc[indices_meta[i]]) if pd.notna(fila.iloc[indices_meta[i]]) else 0
            l = float(fila.iloc[indices_logro[i]]) if pd.notna(fila.iloc[indices_logro[i]]) else 0
            ms.append(m); ls.append(l)
        except:
            ms.append(0); ls.append(0)
    return nombres, ms, ls

def generar_pdf(df_unidades, periodo_txt):
    """Genera un reporte PDF con la lista sencilla de unidades."""
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 16)
    pdf.cell(200, 10, txt="REPORTE DE UNIDADES PENDIENTES", ln=True, align='C')
    pdf.set_font("Arial", "B", 14)
    pdf.cell(200, 10, txt=f"PERIODO: {periodo_txt.upper()}", ln=True, align='C')
    pdf.ln(10)
    
    # Encabezados
    pdf.set_font("Arial", "B", 12)
    pdf.set_fill_color(200, 220, 255)
    pdf.cell(70, 10, "MUNICIPIO", 1, 0, 'C', True)
    pdf.cell(120, 10, "UNIDAD MEDICA", 1, 1, 'C', True)
    
    # Datos
    pdf.set_font("Arial", "", 10)
    for _, row in df_unidades.iterrows():
        pdf.cell(70, 10, str(row['MUNICIPIO']), 1)
        pdf.cell(120, 10, str(row['UNIDAD_MEDICA']), 1)
        pdf.ln()
    
    return pdf.output(dest='S').encode('latin-1', 'replace')

if archivos:
    # --- AUDITORÍA ---
    st.sidebar.subheader("⚙️ Auditoría de Carga")
    opciones_base = ["Todos los meses", "Trimestre 1", "Trimestre 2", "Trimestre 3", "Trimestre 4",
                     "Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
    
    periodo_sel = st.sidebar.multiselect("Periodos a auditar:", opciones_base)

    if st.sidebar.button("🔍 Generar Reporte de Omisiones"):
        if not periodo_sel:
            st.error("Por favor, seleccione un periodo.")
        else:
            st.header(f"📋 Resumen de Unidades con Información Pendiente")
            reporte_omisiones = []
            meses_map = {3:'Ene', 6:'Feb', 9:'Mar', 15:'Abr', 18:'May', 21:'Jun', 27:'Jul', 30:'Ago', 33:'Sep', 39:'Oct', 42:'Nov', 45:'Dic'}
            
            # Mapeo de periodos
            meses_validar = []
            if "Todos los meses" in periodo_sel:
                meses_validar = list(meses_map.values())
            else:
                trim_map = {"Trimestre 1":["Ene","Feb","Mar"], "Trimestre 2":["Abr","May","Jun"], "Trimestre 3":["Jul","Ago","Sep"], "Trimestre 4":["Oct","Nov","Dic"]}
                for p in periodo_sel:
                    if p in trim_map: meses_validar.extend(trim_map[p])
                    else: meses_validar.append(p)

            for arc in archivos:
                dict_temp = pd.read_excel(arc, sheet_name=None, header=None)
                for nombre_hoja, df_hoja in dict_temp.items():
                    if df_hoja.shape[1] >= 46:
                        # Analizar datos reales
                        datos = df_hoja.iloc[10:].dropna(subset=[0])
                        for _, fila in datos.iterrows():
                            ind_nom = str(fila.iloc[0]).strip()
                            if len(ind_nom) > 5 and "SERVICIOS" not in ind_nom.upper():
                                for col, mes in meses_map.items():
                                    if mes in meses_validar:
                                        m = float(fila.iloc[col-1]) if pd.notna(fila.iloc[col-1]) else 0
                                        l = float(fila.iloc[col]) if pd.notna(fila.iloc[col]) else 0
                                        if m > 0 and (l == 0 or pd.isna(l)):
                                            reporte_omisiones.append({
                                                "MUNICIPIO": arc.name.replace(".xlsx","").upper(),
                                                "UNIDAD_MEDICA": nombre_hoja.upper()
                                            })

            if reporte_omisiones:
                df_det = pd.DataFrame(reporte_omisiones).drop_duplicates().sort_values(by=["MUNICIPIO", "UNIDAD_MEDICA"])
                st.warning(f"Se detectaron {len(df_det)} unidades con carga pendiente.")
                st.table(df_det)
                
                c1, c2 = st.columns(2)
                # Excel
                out_ex = BytesIO()
                df_det.to_excel(out_ex, index=False, engine='xlsxwriter')
                c1.download_button("📥 Descargar Excel", out_ex.getvalue(), "pendientes.xlsx")
                
                # PDF
                txt_p = ", ".join(periodo_sel)
                pdf_b = generar_pdf(df_det, txt_p)
                c2.download_button("📥 Descargar PDF", pdf_b, "pendientes.pdf", "application/pdf")
            else:
                st.success("✅ Carga completa detectada para los periodos seleccionados.")

    # --- GRÁFICAS ---
    st.sidebar.divider()
    mun_sel = st.sidebar.selectbox("Seleccione Municipio para Gráficas:", [a.name for a in archivos])
    arc_obj = next(a for a in archivos if a.name == mun_sel)
    dict_h = pd.read_excel(arc_obj, sheet_name=None, header=None)
    
    servicios = set()
    for h in dict_h.values():
        if not h.empty and h.shape[1] > 0:
            for s in h.iloc[10:, 0].dropna().unique():
                if len(str(s)) > 5: servicios.add(str(s).strip())
    
    ind_sel = st.selectbox("🎯 Buscar Indicador:", sorted(list(servicios)))
    sede = st.sidebar.selectbox("Unidad Médica:", ["CONSOLIDADO"] + list(dict_h.keys()))

    if sede == "CONSOLIDADO":
        m_t, l_t, n_p = [0]*16, [0]*16, []
        for h in dict_h.values():
            f = h[h.iloc[:, 0].astype(str).str.strip() == ind_sel]
            if not f.empty:
                n_p, m, l = extraer_data_detallada(f.iloc[0])
                m_t = [x+y for x,y in zip(m_t, m)]
                l_t = [x+y for x,y in zip(l_t, l)]
        nom_p, meta_v, logro_v = n_p, m_t, l_t
    else:
        hoja = dict_h[sede]
        f = hoja[hoja.iloc[:, 0].astype(str).str.strip() == ind_sel]
        nom_p, meta_v, logro_v = extraer_data_detallada(f.iloc[0]) if not f.empty else ([],[],[])

    if nom_p:
        df_g = pd.DataFrame({'Periodo': nom_p, 'Meta': meta_v, 'Logro': logro_v})
        df_g = df_g[~df_g['Periodo'].str.contains('Avance')]
        st.plotly_chart(px.bar(df_g, x='Periodo', y=['Meta', 'Logro'], barmode='group', title=f"Desempeño: {ind_sel}"), use_container_width=True)
else:
    st.info("Sube archivos para iniciar.")
