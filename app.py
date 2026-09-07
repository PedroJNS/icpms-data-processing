"""
===============================================================================
Aplicación: Analizador ICP-MS - Concentración (% wt / ppm) (Versión Streamlit)
Desarrollador: Pedro J. Navarrete Segado
Institución: Universidad de Jaén (UJA)
Contacto: pnsegado@ujaen.es
Licencia: GNU General Public License v3.0 (GPL-3.0) (Copyright (c) Pedro J. Navarrete Segado)
===============================================================================
"""

import io
import os
import sqlite3
import warnings
import matplotlib.pyplot as plt
import matplotlib
import pandas as pd
import streamlit as st

# Configuración de página de Streamlit
st.set_page_config(
    page_title="Analizador ICP-MS - Concentración (% wt / ppm)",
    page_icon="🧪",
    layout="wide",
)

# Ocultar advertencias
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
try:
    warnings.filterwarnings("ignore", category=pd.errors.PerformanceWarning)
except AttributeError:
    pass

# --- GESTIÓN DE BASE DE DATOS (SQLite) ---
DB_NAME = "icpms_database.db"

def init_db():
    """Inicializa las tablas de la base de datos si no existen."""
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT UNIQUE NOT NULL
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS analisis (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario_id INTEGER,
            nombre_analisis TEXT NOT NULL,
            fecha TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            df_json TEXT NOT NULL,
            FOREIGN KEY (usuario_id) REFERENCES usuarios (id)
        )
    """)
    conn.commit()
    conn.close()

def obtener_usuarios():
    """Devuelve la lista de usuarios registrados."""
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT nombre FROM usuarios ORDER BY nombre ASC")
    users = [row[0] for row in c.fetchall()]
    conn.close()
    return users

def crear_usuario(nombre):
    """Crea un nuevo usuario en la base de datos."""
    if not nombre.strip():
        return False
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    try:
        c.execute("INSERT INTO usuarios (nombre) VALUES (?)", (nombre.strip(),))
        conn.commit()
        exito = True
    except sqlite3.IntegrityError:
        exito = False
    conn.close()
    return exito

def guardar_analisis_db(usuario_nombre, nombre_analisis, df_results):
    """Guarda un análisis en formato JSON asociado al usuario."""
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT id FROM usuarios WHERE nombre = ?", (usuario_nombre,))
    user_row = c.fetchone()
    if not user_row:
        conn.close()
        return False
    
    user_id = user_row[0]
    df_json = df_results.to_json(orient="split")
    
    c.execute("""
        INSERT INTO analisis (usuario_id, nombre_analisis, df_json)
        VALUES (?, ?, ?)
    """, (user_id, nombre_analisis, df_json))
    conn.commit()
    conn.close()
    return True

def obtener_analisis_usuario(usuario_nombre):
    """Devuelve la lista de análisis guardados de un usuario."""
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("""
        SELECT a.id, a.nombre_analisis, a.fecha 
        FROM analisis a
        JOIN usuarios u ON a.usuario_id = u.id
        WHERE u.nombre = ?
        ORDER BY a.fecha DESC
    """, (usuario_nombre,))
    analisis_list = c.fetchall()
    conn.close()
    return analisis_list

def cargar_analisis_db(analisis_id):
    """Carga el DataFrame guardado correspondiente a un ID de análisis."""
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT df_json FROM analisis WHERE id = ?", (analisis_id,))
    row = c.fetchone()
    conn.close()
    if row:
        return pd.read_json(io.StringIO(row[0]), orient="split")
    return None

# Inicializar DB
init_db()


# --- FUNCIONES DE PROCESAMIENTO ---
def procesar_archivo_raw(uploaded_file):
    """Lee el archivo cargado y localiza los encabezados y columnas clave."""
    if uploaded_file.name.lower().endswith(".csv"):
        df_raw = pd.read_csv(uploaded_file, header=None)
    else:
        excel_obj = pd.ExcelFile(uploaded_file)
        hoja = (
            "icpms"
            if "icpms" in [s.lower() for s in excel_obj.sheet_names]
            else excel_obj.sheet_names[0]
        )
        df_raw = pd.read_excel(excel_obj, sheet_name=hoja, header=None)

    fila_encabezado = None
    col_sample_idx = None
    col_type_idx = None
    col_date_idx = None

    for r_idx, row in df_raw.iterrows():
        for c_idx, val in enumerate(row):
            val_str = str(val).strip().lower()
            if val_str in ["sample name", "nombre muestra"]:
                fila_encabezado = r_idx
                col_sample_idx = c_idx
            elif val_str in ["type", "tipo"]:
                col_type_idx = c_idx
            elif any(
                k in val_str for k in ["acq", "date", "time", "fecha", "hora"]
            ):
                col_date_idx = c_idx

        if fila_encabezado is not None:
            break

    if fila_encabezado is not None:
        headers = (
            df_raw.iloc[fila_encabezado].fillna("").astype(str).tolist()
        )
        df_data = df_raw.iloc[fila_encabezado + 1 :].copy()
        df_data.columns = headers
    else:
        df_data = df_raw.copy()
        col_sample_idx = 3

    if (
        col_date_idx is None
        and col_type_idx is not None
        and col_type_idx > 0
    ):
        col_date_idx = col_type_idx - 1

    return (
        df_raw,
        df_data,
        fila_encabezado,
        col_sample_idx,
        col_type_idx,
        col_date_idx,
    )


def preparar_mapeo_columnas(df_raw, fila_encabezado, col_sample_idx):
    """Mapea las columnas de elementos evitando los ISTD."""
    element_map = {}
    curr_elem = ""
    for c_idx in range(df_raw.shape[1]):
        val0 = str(df_raw.iloc[0, c_idx]).strip()
        if val0 and val0.lower() != "nan":
            curr_elem = val0
        element_map[c_idx] = curr_elem

    cols_interes = []
    for c_idx in range(df_raw.shape[1]):
        param = str(df_raw.iloc[fila_encabezado, c_idx]).strip()
        elem = element_map[c_idx]

        if "( ISTD )" in elem or "istd" in elem.lower():
            continue

        if c_idx == col_sample_idx:
            cols_interes.append((c_idx, "Sample Name", False, "none"))
        elif "conc" in param.lower():
            unit = "ppm" if "ppm" in param.lower() else "ppb"
            nombre_elem_clean = (
                elem.split("[")[0].strip() if "[" in elem else elem.strip()
            )
            nombre_elem = (
                nombre_elem_clean
                if nombre_elem_clean
                else f"Col_{c_idx}"
            )
            cols_interes.append((c_idx, nombre_elem, True, unit))

    return cols_interes


def calcular_resultados(
    df_data,
    cols_interes,
    col_sample_idx,
    muestras_elegidas,
    blancos_seleccionados,
    df_params,
):
    """Calcula los promedios de los blancos y las concentraciones reales (% wt y ppm)."""
    # 1. Calcular promedio de blancos en ppb
    promedio_blancos_ppb = {}
    if blancos_seleccionados:
        col_serie = df_data.iloc[:, col_sample_idx].astype(str).str.strip()
        df_blancos = df_data[col_serie.isin(blancos_seleccionados)]

        for c_idx, _, es_conc, unidad in cols_interes:
            if es_conc:
                valores_ppb = []
                for _, fila in df_blancos.iterrows():
                    v_str = (
                        str(fila.iloc[c_idx])
                        .replace("<", "")
                        .replace(">", "")
                        .replace(",", ".")
                        .strip()
                    )
                    try:
                        val_num = float(v_str)
                        ppb = val_num * 1000.0 if unidad == "ppm" else val_num
                        valores_ppb.append(ppb)
                    except ValueError:
                        pass

                if valores_ppb:
                    promedio_blancos_ppb[c_idx] = sum(valores_ppb) / len(valores_ppb)

    # 2. Recalcular concentraciones (% wt)
    col_serie = df_data.iloc[:, col_sample_idx].astype(str).str.strip()
    df_filt = df_data[col_serie.isin(muestras_elegidas)].copy()

    params_dict = df_params.set_index("Sample Name").to_dict(orient="index")

    filas_export = []
    for _, fila in df_filt.iterrows():
        nombre_muestra = str(fila.iloc[col_sample_idx]).strip()
        p_m = params_dict.get(
            nombre_muestra, {"Masa (mg)": 15.0, "Volumen (mL)": 10.0}
        )

        masa_mg = float(p_m["Masa (mg)"])
        vol_ml = float(p_m["Volumen (mL)"])

        row_dict = {
            "Sample Name": nombre_muestra,
            "Masa (mg)": masa_mg,
            "Volumen (mL)": vol_ml,
        }

        for c_idx, nombre_elem, es_conc, unidad in cols_interes:
            if es_conc:
                v_str = (
                    str(fila.iloc[c_idx])
                    .replace("<", "")
                    .replace(">", "")
                    .replace(",", ".")
                    .strip()
                )
                try:
                    val_num = float(v_str)
                    ppb = val_num * 1000.0 if unidad == "ppm" else val_num
                    blanco_ppb = promedio_blancos_ppb.get(c_idx, 0.0)
                    ppb_corregido = max(0.0, ppb - blanco_ppb)

                    pct = (ppb_corregido * vol_ml) / (masa_mg * 10000.0)
                except ValueError:
                    pct = 0.0

                # Guardamos internamente en % wt
                row_dict[f"{nombre_elem} (% wt)"] = pct

        filas_export.append(row_dict)

    return pd.DataFrame(filas_export)


# --- INTERFAZ STREAMLIT ---

# Créditos y Gestión de Usuarios en Barra Lateral (Sidebar)
with st.sidebar:
    st.header("👤 Usuario / Base de Datos")
    
    lista_usuarios = obtener_usuarios()
    if not lista_usuarios:
        st.info("No hay usuarios registrados. Crea uno para comenzar a guardar análisis.")
    
    usuario_seleccionado = st.selectbox(
        "Seleccionar Usuario:",
        options=["-- Seleccionar --"] + lista_usuarios,
    )
    
    with st.expander("➕ Crear Nuevo Usuario"):
        nuevo_usuario_input = st.text_input("Nombre de usuario:")
        if st.button("Registrar Usuario"):
            if crear_usuario(nuevo_usuario_input):
                st.success(f"Usuario '{nuevo_usuario_input}' registrado.")
                st.rerun()
            else:
                st.error("El usuario ya existe o el nombre no es válido.")

    st.divider()
    st.header("ℹ️ Acerca de")
    st.markdown("""
        **Analizador ICP-MS - Concentración (% wt / ppm)**  
        * **Desarrollador:** Pedro J. Navarrete Segado  
        * **Institución:** Universidad de Jaén (UJA)  
        * **Contacto:** [pnsegado@ujaen.es](mailto:pnsegado@ujaen.es)  
        * **Licencia:** GNU General Public License v3.0 (GPL-3.0)  
        
        *Copyright (c) Pedro J. Navarrete Segado*
        """)
    st.divider()
    st.caption(
        "Web application to upload Agilent 7900 MassHunter files, input sample"
        " digestion data (mass & volume), calculate real concentrations in solid"
        " samples (ppm and %), store user history, and interactively visualize results."
    )

# Encabezado Principal
st.title("🧪 Analizador ICP-MS - Concentración (% wt / ppm)")
st.caption(
    "Desarrollado por Pedro J. Navarrete Segado | Universidad de Jaén (UJA)"
)

# Pestañas principales: Análisis Actual vs Histórico de Análisis
tab_analisis, tab_historico = st.tabs(["🔬 Nuevo Análisis", "📂 Base de Datos / Histórico"])

# ==========================================
# PESTAÑA 1: NUEVO ANÁLISIS
# ==========================================
with tab_analisis:
    # Conmutador de Unidad de Medida
    unidad_medida = st.radio(
        "🔄 Unidad de visualización de resultados:",
        options=["% wt (Porcentaje en Peso)", "ppm (Partes por Millón)"],
        horizontal=True,
    )
    es_ppm = "ppm" in unidad_medida

    # 1. Cargar Documento
    uploaded_file = st.file_uploader(
        "1. Cargar Documento (Excel / CSV)", type=["xlsx", "xls", "csv"]
    )

    if uploaded_file is not None:
        try:
            (
                df_raw,
                df_data,
                fila_encabezado,
                col_sample_idx,
                col_type_idx,
                col_date_idx,
            ) = procesar_archivo_raw(uploaded_file)
            cols_interes = preparar_mapeo_columnas(
                df_raw, fila_encabezado, col_sample_idx
            )

            # Identificar Muestras y Blancos sugeridos
            palabras_ignorar = [
                "blank", "blanco", "ppb", "calblk", "calstd", "blkvrfy",
                "qc", "driftchk", "cicspike", "isostd", "dilstd", "bkgnd", "fqblk",
            ]

            if col_type_idx is not None and fila_encabezado is not None:
                val_type = (
                    df_data.iloc[:, col_type_idx]
                    .astype(str)
                    .str.strip()
                    .str.lower()
                )
                df_muestras_raw = df_data[val_type == "sample"]
            else:
                df_muestras_raw = df_data

            todas_muestras = (
                df_muestras_raw.iloc[:, col_sample_idx].dropna().astype(str).tolist()
            )

            muestras_validas = []
            blancos_detectados = []

            for m in todas_muestras:
                m_clean = m.strip()
                m_lower = m_clean.lower()
                if m_clean and m_lower != "nan":
                    if "blank" in m_lower or "blanco" in m_lower:
                        blancos_detectados.append(m_clean)
                    elif not any(p in m_lower for p in palabras_ignorar):
                        muestras_validas.append(m_clean)

            # Quitar duplicados manteniendo orden
            muestras_validas = list(dict.fromkeys(muestras_validas))
            blancos_detectados = list(dict.fromkeys(blancos_detectados))

            st.divider()

            # 2 y 3. Selección de Muestras y Blancos
            col_sel1, col_sel2 = st.columns(2)

            with col_sel1:
                muestras_elegidas = st.multiselect(
                    "2. Selecciona las muestras a analizar:",
                    options=muestras_validas,
                    default=muestras_validas,
                )

            with col_sel2:
                blancos_seleccionados = st.multiselect(
                    "3. Selecciona los Blancos para restar:",
                    options=blancos_detectados,
                    default=blancos_detectados,
                )

            if not muestras_elegidas:
                st.warning("⚠️ Selecciona al menos una muestra para continuar.")
            else:
                st.divider()
                st.subheader("4. Parámetros de Digestión (Masa y Volumen)")
                st.info(
                    "💡 **Pista:** Puedes editar los valores de Masa y Volumen directamente"
                    " en la tabla a continuación:"
                )

                # DataFrame editable para Masa y Volumen
                df_params_init = pd.DataFrame({
                    "Sample Name": muestras_elegidas,
                    "Masa (mg)": [15.0] * len(muestras_elegidas),
                    "Volumen (mL)": [10.0] * len(muestras_elegidas),
                })

                df_params_edited = st.data_editor(
                    df_params_init,
                    num_rows="fixed",
                    use_container_width=True,
                    column_config={
                        "Masa (mg)": st.column_config.NumberColumn(
                            min_value=0.001, format="%.2f"
                        ),
                        "Volumen (mL)": st.column_config.NumberColumn(
                            min_value=0.001, format="%.2f"
                        ),
                    },
                )

                # Calcular Resultados Base (en % wt)
                df_results_base = calcular_resultados(
                    df_data,
                    cols_interes,
                    col_sample_idx,
                    muestras_elegidas,
                    blancos_seleccionados,
                    df_params_edited,
                )

                # Adaptar DataFrame a la unidad elegida con el conmutador
                df_results_display = df_results_base.copy()
                cols_pct = [c for c in df_results_base.columns if "% wt" in c]

                if es_ppm:
                    for c in cols_pct:
                        elem_name = c.replace(" (% wt)", "")
                        df_results_display[f"{elem_name} (ppm)"] = df_results_display[c] * 10000.0
                        df_results_display.drop(columns=[c], inplace=True)
                    suff = "(ppm)"
                else:
                    suff = "(% wt)"

                st.divider()
                st.subheader(f"📊 Tabla de Resultados {suff}")

                # Dar formato legible a la vista
                df_view = df_results_display.copy()
                cols_val = [c for c in df_view.columns if suff in c]

                for c in cols_val:
                    def fmt_val(val):
                        if es_ppm:
                            if val >= 0.01:
                                return f"{val:.2f} ppm"
                            elif val > 0:
                                return f"{val:.4f} ppm"
                            return "-"
                        else:
                            if val >= 0.000001:
                                return f"{val:.6f}%"
                            elif val > 0:
                                return f"{val:.3e}%"
                            return "-"

                    df_view[c] = df_view[c].apply(fmt_val)

                st.dataframe(df_view, use_container_width=True)

                # Guardar en Base de Datos
                st.divider()
                st.subheader("💾 Guardar Análisis en Base de Datos")
                col_db1, col_db2 = st.columns([2, 1])
                with col_db1:
                    nombre_analisis_input = st.text_input(
                        "Nombre del Análisis / Proyecto:",
                        value=f"Análisis {uploaded_file.name}",
                    )
                with col_db2:
                    st.write(" ")
                    st.write(" ")
                    if st.button("💾 Guardar Análisis", type="secondary"):
                        if usuario_seleccionado and usuario_seleccionado != "-- Seleccionar --":
                            if guardar_analisis_db(usuario_seleccionado, nombre_analisis_input, df_results_base):
                                st.success("¡Análisis guardado correctamente en la base de datos!")
                            else:
                                st.error("No se pudo guardar el análisis.")
                        else:
                            st.warning("⚠️ Selecciona un usuario en la barra lateral antes de guardar.")

                # Botón de exportación a Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df_results_display.to_excel(
                        writer, index=False, sheet_name="Resultados_ICP-MS"
                    )
                excel_bytes = output.getvalue()

                st.download_button(
                    label="📥 Exportar Resultados a Excel",
                    data=excel_bytes,
                    file_name=f"Resultados_ICPMS_{'ppm' if es_ppm else 'pct'}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )

                # --- SECCIÓN GRÁFICA ---
                st.divider()
                st.subheader(f"📈 Gráfico de Concentraciones por Metal {suff}")

                cols_metales = [c for c in df_results_display.columns if suff in c]

                if cols_metales:
                    col_g1, col_g2 = st.columns([1, 3])

                    with col_g1:
                        usar_log = st.checkbox(
                            "Usar escala logarítmica",
                            value=False,
                            help="Útil si hay diferencias de varios órdenes de magnitud entre metales.",
                        )

                        nombres_metales = [c.replace(f" {suff}", "") for c in cols_metales]
                        metales_sel = st.multiselect(
                            "Selecciona metales para graficar:",
                            options=nombres_metales,
                            default=nombres_metales[: min(5, len(nombres_metales))],
                        )

                    with col_g2:
                        if metales_sel:
                            cols_graficar = [f"{m} {suff}" for m in metales_sel]
                            df_plot = df_results_display.set_index("Sample Name")[cols_graficar]
                            df_plot.columns = [
                                c.replace(f" {suff}", "") for c in df_plot.columns
                            ]

                            fig, ax = plt.subplots(figsize=(10, 5), dpi=100)

                            num_metales = len(metales_sel)
                            cmap_dinamico = matplotlib.colormaps["turbo"].resampled(num_metales)

                            df_plot.plot(
                                kind="bar",
                                ax=ax,
                                width=0.7,
                                edgecolor="black",
                                linewidth=0.5,
                                colormap=cmap_dinamico,
                            )

                            label_y = f"Concentración {suff}" + (" - Escala Log" if usar_log else "")
                            if usar_log:
                                ax.set_yscale("log")

                            ax.set_ylabel(label_y, fontsize=10, fontweight="bold")

                            ncols = 1 if num_metales <= 10 else (2 if num_metales <= 20 else 3)
                            ax.legend(
                                title="Elementos",
                                bbox_to_anchor=(1.02, 1),
                                loc="upper left",
                                frameon=True,
                                ncol=ncols,
                                fontsize=8,
                                title_fontsize=9,
                            )

                            ax.set_title(
                                f"Concentración Elemental en Muestras {suff}",
                                fontsize=12,
                                fontweight="bold",
                                pad=12,
                            )
                            ax.set_xlabel("Muestra", fontsize=10, fontweight="bold")
                            ax.grid(axis="y", linestyle="--", alpha=0.6)
                            plt.xticks(rotation=45, ha="right")
                            fig.tight_layout()

                            st.pyplot(fig)

                            # Descarga de Imagen PNG
                            img_buf = io.BytesIO()
                            fig.savefig(img_buf, format="png", dpi=300, bbox_inches="tight")
                            img_buf.seek(0)

                            st.download_button(
                                label="💾 Descargar Gráfico como PNG",
                                data=img_buf,
                                file_name="grafico_concentraciones_icpms.png",
                                mime="image/png",
                            )
                            plt.close(fig)
                        else:
                            st.info(
                                "Selecciona al menos un metal en el panel de la izquierda para"
                                " visualizar el gráfico."
                            )

        except Exception as e:
            st.error(f"❌ Error al procesar el archivo: {e}")


# ==========================================
# PESTAÑA 2: HISTÓRICO Y BASE DE DATOS
# ==========================================
with tab_historico:
    st.subheader("📁 Consulta de Análisis Guardados")
    
    if not usuario_seleccionado or usuario_seleccionado == "-- Seleccionar --":
        st.info("👈 Por favor, selecciona un usuario en la barra lateral para ver su historial de análisis.")
    else:
        analisis_guardados = obtener_analisis_usuario(usuario_seleccionado)
        if not analisis_guardados:
            st.warning(f"No se encontraron análisis guardados para el usuario '{usuario_seleccionado}'.")
        else:
            opciones_analisis = {
                f"{a[1]} (Fecha: {a[2]})": a[0] for a in analisis_guardados
            }
            
            analisis_elegido_str = st.selectbox(
                "Selecciona un análisis guardado previamente:",
                options=list(opciones_analisis.keys()),
            )
            
            analisis_id = opciones_analisis[analisis_elegido_str]
            df_cargado_base = cargar_analisis_db(analisis_id)
            
            if df_cargado_base is not None:
                st.success(f"Análisis cargado correctamente.")
                
                # Conmutador para el análisis histórico
                unidad_medida_hist = st.radio(
                    "🔄 Unidad de visualización para el análisis histórico:",
                    options=["% wt (Porcentaje en Peso)", "ppm (Partes por Millón)"],
                    horizontal=True,
                    key="radio_hist",
                )
                es_ppm_hist = "ppm" in unidad_medida_hist
                
                df_hist_display = df_cargado_base.copy()
                cols_pct_hist = [c for c in df_cargado_base.columns if "% wt" in c]
                
                if es_ppm_hist:
                    for c in cols_pct_hist:
                        elem_name = c.replace(" (% wt)", "")
                        df_hist_display[f"{elem_name} (ppm)"] = df_hist_display[c] * 10000.0
                        df_hist_display.drop(columns=[c], inplace=True)
                    suff_hist = "(ppm)"
                else:
                    suff_hist = "(% wt)"
                
                st.dataframe(df_hist_display, use_container_width=True)
                
                # Descargar Excel de la consulta histórica
                output_h = io.BytesIO()
                with pd.ExcelWriter(output_h, engine="openpyxl") as writer:
                    df_hist_display.to_excel(
                        writer, index=False, sheet_name="Resultados_Guardados"
                    )
                excel_bytes_h = output_h.getvalue()

                st.download_button(
                    label="📥 Exportar Análisis Recuperado a Excel",
                    data=excel_bytes_h,
                    file_name=f"Analisis_Guardado_{'ppm' if es_ppm_hist else 'pct'}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_hist_excel",
                )
                
                # Gráfico interactivo para datos históricos
                st.divider()
                st.subheader(f"📈 Gráfico del Análisis Recuperado {suff_hist}")
                cols_metales_hist = [c for c in df_hist_display.columns if suff_hist in c]
                
                if cols_metales_hist:
                    nombres_metales_h = [c.replace(f" {suff_hist}", "") for c in cols_metales_hist]
                    metales_sel_h = st.multiselect(
                        "Selecciona metales para graficar:",
                        options=nombres_metales_h,
                        default=nombres_metales_h[: min(5, len(nombres_metales_h))],
                        key="ms_hist",
                    )
                    
                    if metales_sel_h:
                        cols_graficar_h = [f"{m} {suff_hist}" for m in metales_sel_h]
                        df_plot_h = df_hist_display.set_index("Sample Name")[cols_graficar_h]
                        df_plot_h.columns = [c.replace(f" {suff_hist}", "") for c in df_plot_h.columns]
                        
                        fig_h, ax_h = plt.subplots(figsize=(10, 5), dpi=100)
                        num_m_h = len(metales_sel_h)
                        cmap_h = matplotlib.colormaps["turbo"].resampled(num_m_h)
                        
                        df_plot_h.plot(
                            kind="bar",
                            ax=ax_h,
                            width=0.7,
                            edgecolor="black",
                            linewidth=0.5,
                            colormap=cmap_h,
                        )
                        
                        ax_h.set_ylabel(f"Concentración {suff_hist}", fontsize=10, fontweight="bold")
                        ax_h.set_title(f"Concentración Elemental - {analisis_elegido_str}", fontsize=12, fontweight="bold")
                        ax_h.grid(axis="y", linestyle="--", alpha=0.6)
                        plt.xticks(rotation=45, ha="right")
                        fig_h.tight_layout()
                        
                        st.pyplot(fig_h)
                        plt.close(fig_h)
