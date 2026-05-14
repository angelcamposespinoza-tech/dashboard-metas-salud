import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from fpdf import FPDF

st.set_page_config(page_title="Monitor de Metas Jurisdiccional", layout="wide")
st.title("📊 Monitor de Metas Jurisdicción nº1")

archivos = st.file_uploader("Subir archivos Excel", type=["xlsx"], accept_multiple_files=True)

def extraer_data_detallada(fila):
    """Extrae Meta y Logro usando el patrón de columnas D, G, J..."""
    # Columnas de Logro: D=3, G=6, J=9, M=12, P=15, S=18, V=21, Y=24, AB=27, AE=30, AH=33, AK=36
    indices_logro = [3, 6, 9, 12, 15, 18, 21, 24, 27, 30, 33, 36]
    indices_meta = [i-1 for i in indices_logro]
    nombres = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic']
    
    ms, ls = [], []
    for i in range(len(indices_logro)):
        try:
            m = float(fila.iloc[indices_meta[i]]) if pd.notna(fila.iloc[indices_meta[i]]) else 0
            # CRITERIO NUEVO: Si es NaN es omisión, si es 0 es dato válido.
            l_val = fila.iloc[indices_logro[i]]
            l = float(l_val) if pd.notna(l_val) else None 
            ms.append(m)
            ls.append(l if l is not None else 0)
        except:
            ms.append(0); ls.append(0)
    return nombres, ms, ls

def generar_pdf(df_unidades, periodo_txt):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 16)
    pdf.cell(200, 10, txt="REPORTE DE UNIDADES PENDIENTES", ln=True, align='C')
    pdf.set_font("Arial", "B", 12)
    pdf.cell(200, 10, txt=f"CRITERIO: CELDAS VACIAS | PERIODO: {periodo_txt.upper()}", ln=True, align='C')
    pdf.ln(10)
    pdf.set_fill_color(230, 230, 230)
    pdf.cell(70, 10, "MUNICIPIO", 1, 0, 'C', True)
    pdf.cell(120, 10, "UNIDAD MEDICA", 1, 1, 'C', True)
    pdf.set_font("Arial", "", 10)
    for _, row in df_unidades.iterrows():
        pdf.cell(70, 10, str(row['MUNICIPIO']), 1)
        pdf.cell(120, 10, str(row['UNIDAD_MEDICA']), 1)
        pdf.ln()
    return pdf.output(dest='S').encode('latin-1', 'replace')

if archivos:
    st.sidebar.divider()
    st.sidebar.subheader("⚙️ Auditoría de Carga")
    opciones_base = ["Todos los meses", "Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
    periodo_sel = st.sidebar.multiselect("Periodos a revisar (Celdas vacías):", opciones_base)

    meses_map = {3:'Ene', 6:'Feb', 9:'Mar', 12:'Abr', 15:'May', 18:'Jun', 21:'Jul', 24:'Ago', 27:'Sep', 30:'Oct', 33:'Nov', 36:'Dic'}
    
    validos = []
    if "Todos los meses" in periodo_sel: validos = list(meses_map.values())
    else: validos = periodo_sel

    # BOTÓN 1: LISTA SENCILLA
    if st.sidebar.button("🔍 Generar Lista de Unidades sin Información"):
        if not periodo_sel: st.error("Seleccione un periodo.")
        else:
            st.header(f"📋 Unidades con Celdas Vacías en Logro ({', '.join(periodo_sel)})")
            reporte_omisiones = []
            for arc in archivos:
                dict_temp = pd.read_excel(arc, sheet_name=None, header=None)
                for nombre_hoja, df_hoja in dict_temp.items():
                    if not df_hoja.empty and df_hoja.shape[1] >= 40:
                        datos = df_hoja.iloc[10:].dropna(subset=[0])
                        for _, fila in datos.iterrows():
                            if len(str(fila.iloc[0])) > 5 and "SERVICIOS" not in str(fila.iloc[0]).upper():
                                for col_idx, mes_nom in meses_map.items():
                                    if mes_nom in validos:
                                        # CRITERIO CLAVE: Solo es omisión si es NaN (Celda vacía)
                                        logro_valor = fila.iloc[col_idx]
                                        if pd.isna(logro_valor):
                                            reporte_omisiones.append({
                                                "MUNICIPIO": arc.name.replace(".xlsx","").upper(), 
                                                "UNIDAD_MEDICA": nombre_hoja.upper()
                                            })
            if reporte_omisiones:
                df_det = pd.DataFrame(reporte_omisiones).drop_duplicates().sort_values(by=["MUNICIPIO", "UNIDAD_MEDICA"])
                st.table(df_det)
                pdf_b = generar_pdf(df_det, ", ".join(periodo_sel))
                st.download_button("📥 Descargar Reporte PDF", pdf_b, "lista_pendientes.pdf", "application/pdf")
            else: st.success("✅ ¡Carga completa! No se encontraron celdas vacías en los periodos seleccionados.")

    # BOTÓN 2: REPORTE DETALLADO
    if st.sidebar.button("📑 Ver detalle de indicadores no llenados"):
        if not periodo_sel: st.error("Seleccione un periodo.")
        else:
            st.header(f"⚠️ Detalle de Celdas Vacías por Indicador")
            reporte_full = []
            for arc in archivos:
                dict_temp = pd.read_excel(arc, sheet_name=None, header=None)
                for nombre_hoja, df_hoja in dict_temp.items():
                    if not df_hoja.empty and df_hoja.shape[1] >= 40:
                        datos = df_hoja.iloc[10:].dropna(subset=[0])
                        for _, fila in datos.iterrows():
                            ind_nom = str(fila.iloc[0]).strip()
                            if len(ind_nom) > 5 and "SERVICIOS" not in ind_nom.upper():
                                for col_idx, mes_nom in meses_map.items():
                                    if mes_nom in validos:
                                        if pd.isna(fila.iloc[col_idx]):
                                            reporte_full.append({
                                                "MUNICIPIO": arc.name.replace(".xlsx","").upper(),
                                                "UNIDAD": nombre_hoja.upper(),
                                                "INDICADOR": ind_nom,
                                                "MES": mes_nom
                                            })
            if reporte_full:
                df_full = pd.DataFrame(reporte_full)
                st.dataframe(df_full, use_container_width=True)
                out_ex = BytesIO()
                df_full.to_excel(out_ex, index=False, engine='xlsxwriter')
                st.download_button("📥 Descargar Detalle Excel", out_ex.getvalue(), "detalle_vacios.xlsx")
            else: st.success("✅ Todas las celdas tienen al menos un valor (incluyendo ceros).")

    # --- GRÁFICAS ---
    st.sidebar.divider()
    mun_sel = st.sidebar.selectbox("Municipio para Gráficas:", [a.name for a in archivos])
    arc_obj = next(a for a in archivos if a.name == mun_sel)
    dict_h = pd.read_excel(arc_obj, sheet_name=None, header=None)
    
    servs = set()
    for h in dict_h.values():
        if not h.empty and h.shape[1] > 0:
            for s in h.iloc[10:, 0].dropna().unique():
                if len(str(s)) > 5: servs.add(str(s).strip())
    
    if servs:
        ind_sel = st.selectbox("🎯 Seleccionar Indicador para Visualizar:", sorted(list(servs)))
        sede = st.sidebar.selectbox("Unidad:", ["CONSOLIDADO"] + list(dict_h.keys()))
        nom_p, meta_v, logro_v = [], [], []

        if sede == "CONSOLIDADO":
            m_tot, l_tot = [0]*12, [0]*12
            for h in dict_h.values():
                if not h.empty and h.shape[1] > 0:
                    f = h[h.iloc[:, 0].astype(str).str.strip() == ind_sel]
                    if not f.empty:
                        nom_p, m, l = extraer_data_detallada(f.iloc[0])
                        m_tot = [x+y for x,y in zip(m_tot, m)]
                        l_tot = [x+y for x,y in zip(l_tot, l)]
            meta_v, logro_v = m_tot, l_tot
        else:
            hoja = dict_h[sede]
            f = hoja[hoja.iloc[:, 0].astype(str).str.strip() == ind_sel] if not hoja.empty else pd.DataFrame()
            if not f.empty: nom_p, meta_v, logro_v = extraer_data_detallada(f.iloc[0])

        if nom_p:
            df_g = pd.DataFrame({'Mes': nom_p, 'Meta': meta_v, 'Logro': logro_v})
            st.plotly_chart(px.bar(df_g, x='Mes', y=['Meta', 'Logro'], barmode='group', title=f"Seguimiento Anual: {ind_sel}"), use_container_width=True)
else:
    st.info("Sube los archivos Excel para iniciar el monitor.")
