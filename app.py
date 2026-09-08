"""
===============================================================================
Aplicación: Analizador ICP-MS - Concentración (% wt / ppm) (Versión Streamlit)
Desarrollador: Pedro J. Navarrete Segado
Institución: Universidad de Jaén (UJA)
Contacto: pnsegado@ujaen.es
Repositorio: https://github.com/PedroJNS/icpms-data-processing
Licencia: GNU General Public License v3.0 (GPL-3.0)
===============================================================================
"""

import io
import os
import sqlite3
import warnings
import hashlib
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

# --- GESTIÓN DE SEGURIDAD Y BASE DE DATOS (SQLite Ruta Absoluta) ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "icpms_database.db")

def hash_password(password: str) -> str:
    """Devuelve el hash SHA-256 de una contraseña."""
    return hashlib.sha256(password.encode("utf-8")).hexdigest()

def init_db():
    """Inicializa y migra la base de datos SQLite con ruta absoluta."""
    conn = sqlite3.connect(DB_PATH, timeout=15)
    c = conn.cursor()
    c.execute("PRAGMA journal_mode=WAL;")
    c.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT UNIQUE NOT NULL,
            clave_hash TEXT NOT NULL
        )
    """)
    
    c.execute("PRAGMA table_info(usuarios)")
    columns = [row[1] for row in c.fetchall()]
    if "clave_hash" not in columns:
        c.execute("ALTER TABLE usuarios ADD COLUMN clave_hash TEXT DEFAULT ''")
    
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

def registrar_usuario(nombre, clave):
    """Registra un nuevo usuario con su contraseña cifrada."""
    nombre_clean = nombre.strip()
    clave_clean = clave.strip()
    
    if not nombre_clean or not clave_clean:
        return False, "El usuario y la contraseña no pueden estar vacíos."
    
    conn = sqlite3.connect(DB_PATH, timeout=15)
    c = conn.cursor()
    try:
        c.execute(
            "INSERT INTO usuarios (nombre, clave_hash) VALUES (?, ?)",
            (nombre_clean, hash_password(clave_clean)),
        )
        conn.commit()
        exito, msg = True, f"Usuario '{nombre_clean}' registrado con éxito."
    except sqlite3.IntegrityError:
        exito, msg = False, "El nombre de usuario ya está registrado."
    finally:
        conn.close()
    return exito, msg

def verificar_usuario(nombre, clave):
    """Verifica credenciales sin distinción estricta de mayúsculas/minúsculas."""
    conn = sqlite3.connect(DB_PATH, timeout=15)
    c = conn.cursor()
    c.execute("SELECT clave_hash FROM usuarios WHERE LOWER(nombre) = LOWER(?)", (nombre.strip(),))
    row = c.fetchone()
    conn.close()
    if row and row[0] == hash_password(clave.strip()):
        return True
    return False

def guardar_analisis_db(usuario_nombre, nombre_analisis, df_results):
    """Guarda un análisis asociado al usuario de forma permanente."""
    conn = sqlite3.connect(DB_PATH, timeout=15)
    c = conn.cursor()
    c.execute("SELECT id FROM usuarios WHERE LOWER(nombre) = LOWER(?)", (usuario_nombre.strip(),))
    user_row = c.fetchone()
    if not user_row:
        conn.close()
        return False
    
    user_id = user_row[0]
    df_json = df_results.to_json(orient="split")
    
    c.execute("""
        INSERT INTO analisis (usuario_id, nombre_analisis, df_json)
        VALUES (?, ?, ?)
    """, (user_id, nombre_analisis.strip(), df_json))
    conn.commit()
    conn.close()
    return True

def obtener_analisis_usuario(usuario_nombre):
    """Devuelve los análisis guardados filtrados por usuario."""
    conn = sqlite3.connect(DB_PATH, timeout=15)
    c = conn.cursor()
    c.execute("""
        SELECT a.id, a.nombre_analisis, a.fecha 
        FROM analisis a
        JOIN usuarios u ON a.usuario_id = u.id
        WHERE LOWER(u.nombre) = LOWER(?)
        ORDER BY a.fecha DESC
    """, (usuario_nombre.strip(),))
    analisis_list = c.fetchall()
    conn.close()
    return analisis_list

def cargar_analisis_db(analisis_id, usuario_nombre):
    """Carga los datos de un análisis específico."""
    conn = sqlite3.connect(DB_PATH, timeout=15)
    c = conn.cursor()
    c.execute("""
        SELECT a.df_json 
        FROM analisis a
        JOIN usuarios u ON a.usuario_id = u.id
        WHERE a.id = ? AND LOWER(u.nombre) = LOWER(?)
    """, (analisis_id, usuario_nombre.strip()))
    row = c.fetchone()
    conn.close()
    if row:
        return pd.read_json(io.StringIO(row[0]), orient="split")
    return None

# Inicializar Base de Datos
init_db()

# --- FUNCIONES DE PROCESAMIENTO DE DATOS ---
def procesar_archivo_raw(uploaded_file):
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

                row_dict[f"{nombre_elem} (% wt)"] = pct

        filas_export.append(row_dict)

    return pd.DataFrame(filas_export)

def generar_excel_proyecto(df_results_base, df_results_display):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_results_display.to_excel(writer, index=False, sheet_name="Resultados_ICPMS")
        df_meta = pd.DataFrame({"json_data": [df_results_base.to_json(orient="split")]})
        df_meta.to_excel(writer, index=False, sheet_name="_ICPMS_Metadata")
    return output.getvalue()

def cargar_desde_excel_proyecto(uploaded_excel):
    try:
        excel_obj = pd.ExcelFile(uploaded_excel)
        if "_ICPMS_Metadata" in excel_obj.sheet_names:
            df_meta = pd.read_excel(excel_obj, sheet_name="_ICPMS_Metadata")
            json_str = df_meta["json_data"].iloc[0]
            return pd.read_json(io.StringIO(json_str), orient="split")
        else:
            df_direct = pd.read_excel(excel_obj, sheet_name=excel_obj.sheet_names[0])
            return df_direct
    except Exception as e:
        st.error(f"Error al leer el archivo de proyecto Excel: {e}")
        return None

# --- GESTIÓN DE SESIÓN ---
if "usuario_logueado" not in st.session_state:
    st.session_state["usuario_logueado"] = None

# --- BARRA LATERAL (SIDEBAR) ---
with st.sidebar:
    st.header("👤 Control de Acceso")
    
    if st.session_state["usuario_logueado"] is None:
        st.subheader("🔑 Iniciar Sesión")
        user_input = st.text_input("Usuario:", key="login_user")
        pass_input = st.text_input("Contraseña:", type="password", key="login_pass")
        
        if st.button("Ingresar", type="primary"):
            if verificar_usuario(user_input, pass_input):
                st.session_state["usuario_logueado"] = user_input.strip()
                st.success(f"¡Bienvenido, {user_input.strip()}!")
                st.rerun()
            else:
                st.error("Credenciales incorrectas. Inténtalo de nuevo.")
                
        st.divider()
        with st.expander("➕ Registrar Nuevo Usuario"):
            reg_user = st.text_input("Nuevo Usuario:", key="reg_user")
            reg_pass = st.text_input("Nueva Contraseña:", type="password", key="reg_pass")
            if st.button("Registrar Cuenta"):
                exito, msg = registrar_usuario(reg_user, reg_pass)
                if exito:
                    st.success(msg)
                else:
                    st.error(msg)
    else:
        usuario_actual = st.session_state["usuario_logueado"]
        st.success(f"🟢 **Sesión Activa:**\n{usuario_actual}")
        if st.button("🚪 Cerrar Sesión"):
            st.session_state["usuario_logueado"] = None
            st.rerun()

    st.divider()
    st.header("ℹ️ Acerca de")
    st.markdown("""
        **Analizador ICP-MS - Concentración (% wt / ppm)**  
        * **Desarrollador:** Pedro J. Navarrete Segado  
        * **Institución:** Universidad de Jaén (UJA)  
        * **Contacto:** [pnsegado@ujaen.es](mailto:pnsegado@ujaen.es)  
        * **Repositorio:** [GitHub](https://github.com/PedroJNS/icpms-data-processing)  
        * **Licencia:** GNU General Public License v3.0 (GPL-3.0)  
        
        *Copyright (c) Pedro J. Navarrete Segado*
        """)
    st.divider()
    st.caption(
        "Web application to upload Agilent 7900 MassHunter files, input sample"
        " digestion data (mass & volume), calculate real concentrations in solid"
        " samples (ppm and %), store user history safely, and interactively visualize results."
    )

# --- VISTA PRINCIPAL ---
st.title("🧪 Analizador ICP-MS - Concentración (% wt / ppm)")
st.caption("Desarrollado por Pedro J. Navarrete Segado | Universidad de Jaén (UJA)")

tab_analisis, tab_historico, tab_cargar_excel = st.tabs([
    "🔬 Nuevo Análisis", 
    "📂 Base de Datos (Privada)", 
    "📥 Cargar Proyecto Excel"
])

# ==========================================
# PESTAÑA 1: NUEVO ANÁLISIS
# ==========================================
with tab_analisis:
    unidad_medida = st.radio(
        "🔄 Unidad de visualización de resultados:",
        options=["% wt (Porcentaje en Peso)", "ppm (Partes por Millón)"],
        horizontal=True,
    )
    es_ppm = "ppm" in unidad_medida

    uploaded_file = st.file_uploader(
        "1. Cargar Documento ICP-MS (Excel / CSV de Agilent)", type=["xlsx", "xls", "csv"]
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

            muestras_validas = list(dict.fromkeys(muestras_validas))
            blancos_detectados = list(dict.fromkeys(blancos_detectados))

            st.divider()

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
                    "💡 **Pista:** Puedes editar los valores de Masa (mg) y Volumen (mL) directamente en la tabla:"
                )

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

                df_results_base = calcular_resultados(
                    df_data,
                    cols_interes,
                    col_sample_idx,
                    muestras_elegidas,
                    blancos_seleccionados,
                    df_params_edited,
                )

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

                st.divider()
                st.subheader("💾 Guardar Análisis en Base de Datos Privada")
                
                usuario_activo = st.session_state["usuario_logueado"]
                if usuario_activo:
                    col_db1, col_db2 = st.columns([2, 1])
                    with col_db1:
                        nombre_analisis_input = st.text_input(
                            "Nombre del Análisis / Proyecto:",
                            value=f"Análisis {uploaded_file.name}",
                        )
                    with col_db2:
                        st.write(" ")
                        st.write(" ")
                        if st.button("💾 Guardar en mi Cuenta", type="secondary"):
                            if guardar_analisis_db(usuario_activo, nombre_analisis_input, df_results_base):
                                st.success(f"¡Análisis '{nombre_analisis_input}' guardado correctamente!")
                            else:
                                st.error("No se pudo guardar el análisis. Verifica tu usuario.")
                else:
                    st.info("ℹ️ Para guardar análisis en la base de datos, por favor **inicia sesión** en la barra lateral.")

                excel_proyecto_bytes = generar_excel_proyecto(df_results_base, df_results_display)

                st.download_button(
                    label="📥 Exportar Proyecto a Excel (.xlsx)",
                    data=excel_proyecto_bytes,
                    file_name=f"Proyecto_ICPMS_{'ppm' if es_ppm else 'pct'}.xlsx",
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
                            st.info("Selecciona al menos un metal para visualizar el gráfico.")

        except Exception as e:
            st.error(f"❌ Error al procesar el archivo: {e}")


# ==========================================
# PESTAÑA 2: HISTÓRICO Y BASE DE DATOS PRIVADA
# ==========================================
with tab_historico:
    st.subheader("📁 Consulta de Análisis Guardados")
    
    usuario_activo = st.session_state["usuario_logueado"]
    if not usuario_activo:
        st.warning("🔒 Acceso restringido. Por favor, **inicia sesión** en la barra lateral para acceder a tu historial privado.")
    else:
        analisis_guardados = obtener_analisis_usuario(usuario_activo)
        if not analisis_guardados:
            st.info(f"No se encontraron análisis guardados en la cuenta de '{usuario_activo}'.")
        else:
            opciones_analisis = {
                f"{a[1]} (Fecha: {a[2]})": a[0] for a in analisis_guardados
            }
            
            analisis_elegido_str = st.selectbox(
                "Selecciona un análisis guardado de tu cuenta:",
                options=list(opciones_analisis.keys()),
            )
            
            analisis_id = opciones_analisis[analisis_elegido_str]
            df_cargado_base = cargar_analisis_db(analisis_id, usuario_activo)
            
            if df_cargado_base is not None:
                st.success("Análisis cargado correctamente desde la base de datos.")
                
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
                
                excel_bytes_h = generar_excel_proyecto(df_cargado_base, df_hist_display)

                st.download_button(
                    label="📥 Exportar Consulta a Excel",
                    data=excel_bytes_h,
                    file_name=f"Resultados_Guardados_{usuario_activo}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_export_hist",
                )

                # --- SECCIÓN GRÁFICA EN BASE DE DATOS ---
                st.divider()
                st.subheader(f"📈 Gráfico de Concentraciones por Metal {suff_hist}")

                cols_metales_hist = [c for c in df_hist_display.columns if suff_hist in c]

                if cols_metales_hist:
                    col_gh1, col_gh2 = st.columns([1, 3])

                    with col_gh1:
                        usar_log_hist = st.checkbox(
                            "Usar escala logarítmica",
                            value=False,
                            key="usar_log_hist",
                            help="Útil si hay diferencias de varios órdenes de magnitud entre metales.",
                        )

                        nombres_metales_hist = [c.replace(f" {suff_hist}", "") for c in cols_metales_hist]
                        metales_sel_hist = st.multiselect(
                            "Selecciona metales para graficar:",
                            options=nombres_metales_hist,
                            default=nombres_metales_hist[: min(5, len(nombres_metales_hist))],
                            key="metales_sel_hist",
                        )

                    with col_gh2:
                        if metales_sel_hist:
                            cols_graficar_hist = [f"{m} {suff_hist}" for m in metales_sel_hist]
                            df_plot_hist = df_hist_display.set_index("Sample Name")[cols_graficar_hist]
                            df_plot_hist.columns = [
                                c.replace(f" {suff_hist}", "") for c in df_plot_hist.columns
                            ]

                            fig_h, ax_h = plt.subplots(figsize=(10, 5), dpi=100)

                            num_metales_h = len(metales_sel_hist)
                            cmap_dinamico_h = matplotlib.colormaps["turbo"].resampled(num_metales_h)

                            df_plot_hist.plot(
                                kind="bar",
                                ax=ax_h,
                                width=0.7,
                                edgecolor="black",
                                linewidth=0.5,
                                colormap=cmap_dinamico_h,
                            )

                            label_y_h = f"Concentración {suff_hist}" + (" - Escala Log" if usar_log_hist else "")
                            if usar_log_hist:
                                ax_h.set_yscale("log")

                            ax_h.set_ylabel(label_y_h, fontsize=10, fontweight="bold")

                            ncols_h = 1 if num_metales_h <= 10 else (2 if num_metales_h <= 20 else 3)
                            ax_h.legend(
                                title="Elementos",
                                bbox_to_anchor=(1.02, 1),
                                loc="upper left",
                                frameon=True,
                                ncol=ncols_h,
                                fontsize=8,
                                title_fontsize=9,
                            )

                            ax_h.set_title(
                                f"Concentración Elemental en Muestras {suff_hist}",
                                fontsize=12,
                                fontweight="bold",
                                pad=12,
                            )
                            ax_h.set_xlabel("Muestra", fontsize=10, fontweight="bold")
                            ax_h.grid(axis="y", linestyle="--", alpha=0.6)
                            plt.xticks(rotation=45, ha="right")
                            fig_h.tight_layout()

                            st.pyplot(fig_h)

                            img_buf_h = io.BytesIO()
                            fig_h.savefig(img_buf_h, format="png", dpi=300, bbox_inches="tight")
                            img_buf_h.seek(0)

                            st.download_button(
                                label="💾 Descargar Gráfico como PNG",
                                data=img_buf_h,
                                file_name=f"grafico_concentraciones_guardado.png",
                                mime="image/png",
                                key="download_plot_hist",
                            )
                            plt.close(fig_h)
                        else:
                            st.info("Selecciona al menos un metal para visualizar el gráfico.")


# ==========================================
# PESTAÑA 3: CARGAR PROYECTO EXCEL (PORTÁTIL)
# ==========================================
with tab_cargar_excel:
    st.subheader("📥 Reanalizar desde un Archivo Excel Local")
    st.caption("Sube un archivo de proyecto `.xlsx` descargado previamente de esta app para volver a examinarlo sin necesidad de usar la base de datos.")

    file_excel_proyecto = st.file_uploader(
        "Cargar Archivo Excel de Proyecto ICP-MS:", type=["xlsx"]
    )

    if file_excel_proyecto is not None:
        df_excel_loaded = cargar_desde_excel_proyecto(file_excel_proyecto)
        if df_excel_loaded is not None:
            st.success("¡Proyecto Excel cargado con éxito!")
            
            unidad_medida_ex = st.radio(
                "🔄 Unidad de visualización:",
                options=["% wt (Porcentaje en Peso)", "ppm (Partes por Millón)"],
                horizontal=True,
                key="radio_excel_tab",
            )
            es_ppm_ex = "ppm" in unidad_medida_ex
            
            df_ex_display = df_excel_loaded.copy()
            cols_pct_ex = [c for c in df_excel_loaded.columns if "% wt" in c]
            
            if es_ppm_ex and cols_pct_ex:
                for c in cols_pct_ex:
                    elem_name = c.replace(" (% wt)", "")
                    df_ex_display[f"{elem_name} (ppm)"] = df_ex_display[c] * 10000.0
                    df_ex_display.drop(columns=[c], inplace=True)
                suff_ex = "(ppm)"
            else:
                suff_ex = "(% wt)"
            
            st.dataframe(df_ex_display, use_container_width=True)
