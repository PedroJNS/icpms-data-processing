==============================================================================
ICP-MS Data Processing & Streamlit Web App (Agilent 7900)
Author: Pedro J. (PedroJNS)
License: GNU General Public License v3.0 (GPL-3.0)
Description: Web application to upload Agilent 7900 MassHunter files, 
             input sample digestion data (mass & volume), calculate 
             real concentrations in solid samples (ppm and %), and 
             interactively visualize analytical results.
==============================================================================

import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

import io
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

# -----------------------------------------------------------------------------
# CONFIGURACIÓN DE LA PÁGINA
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Procesador ICP-MS - Agilent 7900",
    page_icon="🧪",
    layout="wide"
)

st.title("🧪 Procesador de Datos ICP-MS (Agilent 7900)")
st.markdown(
    "Carga archivos de exportación de **Agilent 7900 MassHunter**, ingresa la masa y volumen "
    "de digestión para cada muestra y calcula automáticamente las concentraciones reales en sólido (ppm / mg·kg⁻¹)."
)

# Límites de referencia por defecto para alertas (ppm)
LIMITES_REFERENCIA = {
    "Pb": 10.0,
    "As": 5.0,
    "Cd": 1.0,
    "Hg": 0.5,
    "Cr": 50.0,
    "Ni": 20.0,
    "Cu": 100.0,
    "Zn": 500.0
}

# -----------------------------------------------------------------------------
# FUNCIONES AUXILIARES DE PARSEO DE AGILENT 7900
# -----------------------------------------------------------------------------
def parsear_archivo_agilent(uploaded_file):
    """
    Lee archivos de Agilent MassHunter detectando automáticamente la fila 
    de encabezado ("Sample Name") y estructurando los metales detectados.
    """
    try:
        if uploaded_file.name.lower().endswith('.csv'):
            df_raw = pd.read_csv(uploaded_file, header=None)
        else:
            df_raw = pd.read_excel(uploaded_file, header=None)

        header_row = None
        col_sample_idx = None

        # Localizar la fila que contiene 'Sample Name' o términos equivalentes
        for r_idx, row in df_raw.iterrows():
            for c_idx, val in enumerate(row):
                val_str = str(val).strip().lower()
                if "sample name" in val_str or val_str == "sample" or "nombre muestra" in val_str:
                    header_row = r_idx
                    col_sample_idx = c_idx
                    break
            if header_row is not None:
                break

        if header_row is not None:
            headers = df_raw.iloc[header_row].fillna("").astype(str).tolist()
            df_data = df_raw.iloc[header_row + 1:].copy()
            df_data.columns = headers
        else:
            df_data = df_raw.copy()
            col_sample_idx = 0
            header_row = 0

        # Construir nombres limpios para las columnas
        column_titles = []
        if header_row > 0:
            element_row = df_raw.iloc[header_row - 1].fillna("").astype(str)
            param_row = df_raw.iloc[header_row].fillna("").astype(str)
            
            current_elem = ""
            for e, p in zip(element_row, param_row):
                if e.strip():
                    current_elem = e.strip()
                p_str = p.strip()
                if current_elem and p_str and p_str not in ["Acq. Date-Time", "Type", "Level", "Sample Name"]:
                    column_titles.append(f"{current_elem}")
                else:
                    column_titles.append(p_str)
        else:
            column_titles = [str(c).strip() for c in df_data.columns]

        df_data.columns = column_titles
        return df_raw, df_data, col_sample_idx, header_row

    except Exception as e:
        st.error(f"Error al analizar la estructura del archivo Agilent: {e}")
        return None, None, None, None

# -----------------------------------------------------------------------------
# BARRA LATERAL: CARGA DE DATOS
# -----------------------------------------------------------------------------
st.sidebar.header("📁 1. Carga de Archivo")
uploaded_file = st.sidebar.file_uploader(
    "Selecciona archivo Agilent 7900 (.xlsx, .xls, .csv)", 
    type=["xlsx", "xls", "csv"]
)

if uploaded_file is not None:
    df_raw, df_data, col_sample_idx, header_row = parsear_archivo_agilent(uploaded_file)
    
    if df_data is not None:
        col_muestra_nombre = df_data.columns[col_sample_idx]
        
        # Filtrar nombres de muestra válidos
        muestras_series = df_data[col_muestra_nombre].dropna().astype(str).str.strip()
        muestras_unicas = [m for m in muestras_series.unique() if m and m.lower() not in ["nan", "none", ""]]

        st.sidebar.markdown("---")
        st.sidebar.header("🎯 2. Selección de Muestras")
        
        seleccionar_todas = st.sidebar.checkbox("Seleccionar todas las muestras", value=True)
        muestras_seleccionadas = st.sidebar.multiselect(
            "Muestras a procesar:", 
            options=muestras_unicas, 
            default=muestras_unicas if seleccionar_todas else []
        )

        # Identificar columnas analíticas
        cols_no_metales = [col_muestra_nombre, "Acq. Date-Time", "Type", "Level", "Sample No.", ""]
        posibles_metales = [c for c in df_data.columns if c not in cols_no_metales and not c.startswith("Unnamed")]

        st.sidebar.markdown("---")
        st.sidebar.header("🔬 3. Selección de Metales")
        
        metales_seleccionados = st.sidebar.multiselect(
            "Elementos a desplegar:",
            options=posibles_metales,
            default=posibles_metales
        )

        # -----------------------------------------------------------------------------
        # SECCIÓN CENTRAL: ENTRADA DE PARÁMETROS DE DIGESTIÓN
        # -----------------------------------------------------------------------------
        st.subheader("⚖️ 2. Parámetros de Digestión por Muestra")
        st.caption("Ingresa la masa de muestra sólida (mg) y el volumen de aforo (mL) para calcular las concentraciones reales.")

        if muestras_seleccionadas:
            # Inicializar dataframe de digestión en session_state si no existe
            if "df_digestion" not in st.session_state:
                st.session_state.df_digestion = pd.DataFrame({
                    "Muestra": muestras_unicas,
                    "Masa_mg": [100.0] * len(muestras_unicas),
                    "Volumen_mL": [50.0] * len(muestras_unicas)
                })

            # Asegurar que nuevas muestras se agreguen si cambia el archivo
            missing_samples = set(muestras_unicas) - set(st.session_state.df_digestion["Muestra"])
            if missing_samples:
                new_rows = pd.DataFrame({
                    "Muestra": list(missing_samples),
                    "Masa_mg": [100.0] * len(missing_samples),
                    "Volumen_mL": [50.0] * len(missing_samples)
                })
                st.session_state.df_digestion = pd.concat([st.session_state.df_digestion, new_rows], ignore_index=True)

            with st.expander("📝 Editar Masa (mg) y Volumen (mL) por muestra", expanded=True):
                st.markdown("**Configuración rápida global:**")
                c_m1, c_m2, c_m3 = st.columns([2, 2, 2])
                m_global = c_m1.number_input("Masa global (mg):", value=100.0, step=5.0)
                v_global = c_m2.number_input("Volumen global (mL):", value=50.0, step=5.0)
                
                if c_m3.button("Aplicar a seleccionadas"):
                    mask = st.session_state.df_digestion["Muestra"].isin(muestras_seleccionadas)
                    st.session_state.df_digestion.loc[mask, "Masa_mg"] = m_global
                    st.session_state.df_digestion.loc[mask, "Volumen_mL"] = v_global
                    st.rerun()

                st.markdown("---")
                
                # Editor de datos nativo e interactivo en formato tabla (remplaza el grid dinámico propensa a fallos de estado)
                df_dig_sub = st.session_state.df_digestion[
                    st.session_state.df_digestion["Muestra"].isin(muestras_seleccionadas)
                ].copy()

                edited_df = st.data_editor(
                    df_dig_sub,
                    column_config={
                        "Muestra": st.column_config.TextColumn("Muestra", disabled=True),
                        "Masa_mg": st.column_config.NumberColumn("Masa (mg)", min_value=0.0001, step=1.0, format="%.2f mg"),
                        "Volumen_mL": st.column_config.NumberColumn("Volumen (mL)", min_value=0.0001, step=1.0, format="%.2f mL"),
                    },
                    hide_index=True,
                    use_container_width=True,
                    key="digestion_editor"
                )

                # Actualizar el estado global con los valores editados
                for idx, row in edited_df.iterrows():
                    m_idx = st.session_state.df_digestion["Muestra"] == row["Muestra"]
                    st.session_state.df_digestion.loc[m_idx, "Masa_mg"] = row["Masa_mg"]
                    st.session_state.df_digestion.loc[m_idx, "Volumen_mL"] = row["Volumen_mL"]

            # -----------------------------------------------------------------------------
            # CÁLCULO DE CONCENTRACIONES REALES (PPM Y %)
            # -----------------------------------------------------------------------------
            df_filtered = df_data[df_data[col_muestra_nombre].astype(str).str.strip().isin(muestras_seleccionadas)].copy()
            
            # Unir los datos de digestión con los datos leídos
            df_dig_map = st.session_state.df_digestion.set_index("Muestra")
            df_filtered["_Masa"] = df_filtered[col_muestra_nombre].astype(str).str.strip().map(df_dig_map["Masa_mg"]).fillna(100.0)
            df_filtered["_Volumen"] = df_filtered[col_muestra_nombre].astype(str).str.strip().map(df_dig_map["Volumen_mL"]).fillna(50.0)

            df_ppm = pd.DataFrame()
            df_ppm[col_muestra_nombre] = df_filtered[col_muestra_nombre]

            for col in metales_seleccionados:
                raw_reading = pd.to_numeric(
                    df_filtered[col].astype(str).str.replace(',', '.'), 
                    errors='coerce'
                ).fillna(0.0)

                # Cálculo: ppm (mg/kg) = [ppb (µg/L) * Volumen (mL)] / Masa (mg)
                conc_ppm = (raw_reading * df_filtered["_Volumen"]) / df_filtered["_Masa"]
                df_ppm[col] = conc_ppm

            # -----------------------------------------------------------------------------
            # VISUALIZACIÓN DE RESULTADOS Y TABLAS
            # -----------------------------------------------------------------------------
            st.subheader("📋 3. Resultados Calculados (ppm en Sólido)")
            
            with st.expander("👁️ Ver Tabla Completa de Concentraciones Calculadas", expanded=True):
                st.dataframe(df_ppm.style.format({col: "{:.4f}" for col in metales_seleccionados}), use_container_width=True)

            if metales_seleccionados:
                st.subheader("📊 Visualización Gráfica Interactiva")
                
                c_opt1, c_opt2 = st.columns(2)
                tipo_escala = c_opt1.radio("Escala del eje Y:", ["Lineal", "Logarítmica"], horizontal=True)
                modo_grafico = c_opt2.radio("Tipo de Gráfico:", ["Barras por Muestra", "Perfil de Líneas"], horizontal=True)

                df_melted = df_ppm.melt(
                    id_vars=[col_muestra_nombre], 
                    value_vars=metales_seleccionados,
                    var_name="Metal", 
                    value_name="Concentración (ppm)"
                )

                log_y = True if tipo_escala == "Logarítmica" else False

                if modo_grafico == "Barras por Muestra":
                    fig = px.bar(
                        df_melted, 
                        x=col_muestra_nombre, 
                        y="Concentración (ppm)", 
                        color="Metal", 
                        barmode="group",
                        log_y=log_y,
                        title="Concentración Real de Metales en Muestra Sólida (mg/kg)",
                        labels={col_muestra_nombre: "Muestra"}
                    )
                else:
                    fig = px.line(
                        df_melted, 
                        x=col_muestra_nombre, 
                        y="Concentración (ppm)", 
                        color="Metal",
                        markers=True,
                        log_y=log_y,
                        title="Perfil de Concentración de Metales por Muestra",
                        labels={col_muestra_nombre: "Muestra"}
                    )

                fig.update_layout(height=550, hovermode="x unified")
                st.plotly_chart(fig, use_container_width=True)

                # -----------------------------------------------------------------
                # ESTADÍSTICAS Y ALERTAS DE UMBRAL
                # -----------------------------------------------------------------
                col_stat, col_alert = st.columns([1, 1])

                with col_stat:
                    st.subheader("📈 Resumen Estadístico (ppm)")
                    df_stats = df_ppm[metales_seleccionados].describe().T[['mean', 'std', 'min', '50%', 'max']]
                    df_stats.columns = ['Media', 'Desv. Est.', 'Mínimo', 'Mediana', 'Máximo']
                    st.dataframe(df_stats.style.highlight_max(axis=0, color='#ffcccc'), use_container_width=True)

                with col_alert:
                    st.subheader("⚠️ Verificación de Umbrales")
                    metal_ref = st.selectbox("Selecciona metal para evaluar límite:", options=metales_seleccionados)
                    val_defecto = LIMITES_REFERENCIA.get(metal_ref, 10.0)
                    umbral_max = st.number_input(f"Límite máximo permitido para {metal_ref} (ppm):", value=float(val_defecto))

                    superan_limite = df_ppm[df_ppm[metal_ref] > umbral_max][[col_muestra_nombre, metal_ref]]
                    
                    if not superan_limite.empty:
                        st.warning(f"Se encontraron **{len(superan_limite)}** muestras que superan el límite de {umbral_max} ppm para **{metal_ref}**:")
                        st.dataframe(superan_limite.style.format({metal_ref: "{:.4f}"}), use_container_width=True)
                    else:
                        st.success(f"✅ Ninguna muestra supera el límite de {umbral_max} ppm para **{metal_ref}**.")

                # -----------------------------------------------------------------
                # EXPORTACIÓN DE RESULTADOS
                # -----------------------------------------------------------------
                st.subheader("📥 Exportar Reporte de Resultados")
                
                df_pct = df_ppm.copy()
                for col in metales_seleccionados:
                    df_pct[f"{col} (%)"] = df_pct[col] / 10000.0

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_ppm.to_excel(writer, sheet_name='Concentraciones_ppm', index=False)
                    df_pct.to_excel(writer, sheet_name='Concentraciones_Pct', index=False)
                    df_stats.to_excel(writer, sheet_name='Estadisticas')
                
                st.download_button(
                    label="📄 Descargar Reporte Completo en Excel (.xlsx)",
                    data=buffer.getvalue(),
                    file_name="Reporte_ICP_MS_Calculado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        else:
            st.info("Por favor, selecciona al menos una muestra en la barra lateral para continuar.")
else:
    st.info("👋 Por favor, sube un archivo Excel o CSV de Agilent 7900 desde la barra lateral izquierda para comenzar.")
