"""
===============================================================================
Aplicación: Analizador ICP-MS - Concentración (% wt / ppm)
Desarrollador original: Pedro J. Navarrete Segado (Universidad de Jaén - UJA)
Contacto: pnsegado@ujaen.es
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

# Intentar importar Supabase; si no está instalado, manejar con fallback a SQLite
try:
    from supabase import create_client, Client
    from postgrest.exceptions import APIError
    HAS_SUPABASE_LIB = True
except ImportError:
    HAS_SUPABASE_LIB = False
    APIError = Exception

# ==========================================
# 1. CONFIGURACIÓN DE PÁGINA Y ADVERTENCIAS
# ==========================================
st.set_page_config(
    page_title="Analizador ICP-MS",
    page_icon="🧪",
    layout="wide"
)

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
try:
    warnings.filterwarnings("ignore", category=pd.errors.PerformanceWarning)
except AttributeError:
    pass

# ==========================================
# 2. DICCIONARIO MULTI-IDIOMA (ES / EN)
# ==========================================
TEXTS = {
    "es": {
        "title": "🧪 Analizador ICP-MS - Concentración (% wt / ppm)",
        "subtitle": "Desarrollado por Pedro J. Navarrete Segado | Universidad de Jaén (UJA)",
        "lang_selector": "🌐 Idioma / Language",
        "login_tab": "Iniciar Sesión",
        "register_tab": "Registrarse",
        "user_label": "Usuario",
        "pass_label": "Contraseña",
        "login_btn": "Ingresar",
        "register_btn": "Crear Cuenta",
        "login_success": "¡Bienvenido/a, {}!",
        "login_error": "Usuario o contraseña incorrectos.",
        "user_exists": "El nombre de usuario ya existe.",
        "register_success": "¡Usuario creado con éxito! Ya puedes iniciar sesión.",
        "logout_btn": "Cerrar Sesión",
        "no_users": "No hay usuarios registrados.",
        "fill_all": "Por favor completa todos los campos.",
        "db_error": "Error de base de datos.",
        "tab_analysis": "🔬 Nuevo Análisis",
        "tab_history": "📂 Base de Datos (Privada)",
        "tab_load_excel": "📥 Cargar Proyecto Excel",
        "unit_selector": "🔄 Unidad de visualización de resultados:",
        "upload_file": "1. Cargar Documento ICP-MS (Excel / CSV de Agilent)",
        "select_samples": "2. Selecciona las muestras a analizar:",
        "select_blanks": "3. Selecciona los Blancos para restar:",
        "warn_select_sample": "⚠️ Selecciona al menos una muestra para continuar.",
        "digestion_params": "4. Parámetros de Digestión (Masa y Volumen)",
        "params_hint": "💡 **Pista:** Puedes editar los valores de Masa (mg) y Volumen (mL) directamente en la tabla:",
        "results_table": "📊 Tabla de Resultados",
        "save_analysis_title": "💾 Guardar Análisis en Base de Datos",
        "analysis_name_label": "Nombre del Análisis / Proyecto:",
        "save_btn": "💾 Guardar en mi Cuenta",
        "save_success": "¡Análisis '{}' guardado correctamente!",
        "save_error": "No se pudo guardar el análisis.",
        "login_req_save": "ℹ️ Para guardar análisis en la base de datos, por favor **inicia sesión** en la barra lateral.",
        "export_excel": "📥 Exportar Proyecto a Excel (.xlsx)",
        "chart_title": "📈 Gráfico de Concentraciones por Metal",
        "use_log": "Usar escala logarítmica",
        "select_metals": "Selecciona metales para graficar:",
        "about_header": "ℹ️ Acerca de",
        "db_type_supabase": "Conectado a Supabase (Nube)",
        "db_type_sqlite": "Conectado a SQLite (Local)"
    },
    "en": {
        "title": "🧪 ICP-MS Analyzer - Concentration (% wt / ppm)",
        "subtitle": "Developed by Pedro J. Navarrete Segado | University of Jaén (UJA)",
        "lang_selector": "🌐 Language / Idioma",
        "login_tab": "Log In",
        "register_tab": "Sign Up",
        "user_label": "Username",
        "pass_label": "Password",
        "login_btn": "Log In",
        "register_btn": "Create Account",
        "login_success": "Welcome, {}!",
        "login_error": "Incorrect username or password.",
        "user_exists": "Username already exists.",
        "register_success": "User created successfully! You can now log in.",
        "logout_btn": "Log Out",
        "no_users": "No registered users found.",
        "fill_all": "Please fill in all fields.",
        "db_error": "Database error.",
        "tab_analysis": "🔬 New Analysis",
        "tab_history": "📂 Database (Private)",
        "tab_load_excel": "📥 Load Excel Project",
        "unit_selector": "🔄 Display unit for results:",
        "upload_file": "1. Upload ICP-MS File (Agilent Excel / CSV)",
        "select_samples": "2. Select samples to analyze:",
        "select_blanks": "3. Select Blank samples for subtraction:",
        "warn_select_sample": "⚠️ Select at least one sample to proceed.",
        "digestion_params": "4. Digestion Parameters (Mass & Volume)",
        "params_hint": "💡 **Tip:** Edit Mass (mg) and Volume (mL) directly in the table below:",
        "results_table": "📊 Results Table",
        "save_analysis_title": "💾 Save Analysis to Database",
        "analysis_name_label": "Analysis / Project Name:",
        "save_btn": "💾 Save to My Account",
        "save_success": "Analysis '{}' saved successfully!",
        "save_error": "Could not save the analysis.",
        "login_req_save": "ℹ️ Please **log in** via the sidebar to save analyses to the database.",
        "export_excel": "📥 Export Project to Excel (.xlsx)",
        "chart_title": "📈 Metal Concentration Chart",
        "use_log": "Use logarithmic scale",
        "select_metals": "Select metals to plot:",
        "about_header": "ℹ️ About",
        "db_type_supabase": "Connected to Supabase (Cloud)",
        "db_type_sqlite": "Connected to SQLite (Local)"
    }
}

# Selector de Idioma en la Sidebar
with st.sidebar:
    lang_choice = st.selectbox("🌐 Language / Idioma", ["Español", "English"])
    lang = "es" if lang_choice == "Español" else "en"
    t = TEXTS[lang]

# ==========================================
# 3. SEGURIDAD Y BASE DE DATOS (HÍBRIDA)
# ==========================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "icpms_database.db")

def hash_password(password: str) -> str:
    """Genera hash SHA-256 para contraseñas."""
    return hashlib.sha256(password.encode("utf-8")).hexdigest()

@st.cache_resource
def get_supabase_client():
    if HAS_SUPABASE_LIB and "SUPABASE_URL" in st.secrets and "SUPABASE_KEY" in st.secrets:
        try:
            url = st.secrets["SUPABASE_URL"]
            key = st.secrets["SUPABASE_KEY"]
            return create_client(url, key)
        except Exception:
            return None
    return None

supabase_client = get_supabase_client()

def init_sqlite_db():
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

if not supabase_client:
    init_sqlite_db()

def registrar_usuario(nombre: str, clave: str):
    nombre_clean = nombre.strip()
    clave_clean = clave.strip()
    if not nombre_clean or not clave_clean:
        return False, t["fill_all"]
    
    hashed = hash_password(clave_clean)
    
    if supabase_client:
        try:
            res = supabase_client.table("usuarios").select("id").eq("nombre", nombre_clean).execute()
            if res.data:
                return False, t["user_exists"]
            supabase_client.table("usuarios").insert({"nombre": nombre_clean, "clave_hash": hashed}).execute()
            return True, t["register_success"]
        except Exception as e:
            return False, f"{t['db_error']}: {e}"
    else:
        conn = sqlite3.connect(DB_PATH, timeout=15)
        c = conn.cursor()
        try:
            c.execute("INSERT INTO usuarios (nombre, clave_hash) VALUES (?, ?)", (nombre_clean, hashed))
            conn.commit()
            return True, t["register_success"]
        except sqlite3.IntegrityError:
            return False, t["user_exists"]
        finally:
            conn.close()

def verificar_usuario(nombre: str, clave: str) -> bool:
    nombre_clean = nombre.strip()
    hashed = hash_password(clave.strip())
    
    if supabase_client:
        try:
            res = supabase_client.table("usuarios").select("clave_hash").eq("nombre", nombre_clean).execute()
            if res.data and res.data[0].get("clave_hash") == hashed:
                return True
            return False
        except Exception:
            return False
    else:
        conn = sqlite3.connect(DB_PATH, timeout=15)
        c = conn.cursor()
        c.execute("SELECT clave_hash FROM usuarios WHERE LOWER(nombre) = LOWER(?)", (nombre_clean,))
        row = c.fetchone()
        conn.close()
        return bool(row and row[0] == hashed)

def guardar_analisis_db(usuario_nombre: str, nombre_analisis: str, df_results: pd.DataFrame) -> bool:
    df_json = df_results.to_json(orient="split")
    
    if supabase_client:
        try:
            res = supabase_client.table("usuarios").select("id").eq("nombre", usuario_nombre).execute()
            if not res.data:
                return False
            user_id = res.data[0]["id"]
            supabase_client.table("analisis").insert({
                "usuario_id": user_id,
                "nombre_analisis": nombre_analisis.strip(),
                "df_json": df_json
            }).execute()
            return True
        except Exception:
            return False
    else:
        conn = sqlite3.connect(DB_PATH, timeout=15)
        c = conn.cursor()
        c.execute("SELECT id FROM usuarios WHERE LOWER(nombre) = LOWER(?)", (usuario_nombre.strip(),))
        user_row = c.fetchone()
        if not user_row:
            conn.close()
            return False
        c.execute("""
            INSERT INTO analisis (usuario_id, nombre_analisis, df_json)
            VALUES (?, ?, ?)
        """, (user_row[0], nombre_analisis.strip(), df_json))
        conn.commit()
        conn.close()
        return True

def obtener_analisis_usuario(usuario_nombre: str):
    if supabase_client:
        try:
            res_u = supabase_client.table("usuarios").select("id").eq("nombre", usuario_nombre).execute()
            if not res_u.data:
                return []
            u_id = res_u.data[0]["id"]
            res_a = supabase_client.table("analisis").select("id, nombre_analisis, fecha").eq("usuario_id", u_id).order("fecha", desc=True).execute()
            return [(r["id"], r["nombre_analisis"], r.get("fecha", "")) for r in res_a.data]
        except Exception:
            return []
    else:
        conn = sqlite3.connect(DB_PATH, timeout=15)
        c = conn.cursor()
        c.execute("""
            SELECT a.id, a.nombre_analisis, a.fecha 
            FROM analisis a
            JOIN usuarios u ON a.usuario_id = u.id
            WHERE LOWER(u.nombre) = LOWER(?)
            ORDER BY a.fecha DESC
        """, (usuario_nombre.strip(),))
        data = c.fetchall()
        conn.close()
        return data

def cargar_analisis_db(analisis_id: int, usuario_nombre: str):
    if supabase_client:
        try:
            res = supabase_client.table("analisis").select("df_json").eq("id", analisis_id).execute()
            if res.data:
                return pd.read_json(io.StringIO(res.data[0]["df_json"]), orient="split")
            return None
        except Exception:
            return None
    else:
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

# ==========================================
# 4. FUNCIONES DE PROCESAMIENTO ICP-MS
# ==========================================
def procesar_archivo_raw(uploaded_file):
    if uploaded_file.name.lower().endswith(".csv"):
        df_raw = pd.read_csv(uploaded_file, header=None)
    else:
        excel_obj = pd.ExcelFile(uploaded_file)
        hoja = "icpms" if "icpms" in [s.lower() for s in excel_obj.sheet_names] else excel_obj.sheet_names[0]
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
            elif any(k in val_str for k in ["acq", "date", "time", "fecha", "hora"]):
                col_date_idx = c_idx

        if fila_encabezado is not None:
            break

    if fila_encabezado is not None:
        headers = df_raw.iloc[fila_encabezado].fillna("").astype(str).tolist()
        df_data = df_raw.iloc[fila_encabezado + 1 :].copy()
        df_data.columns = headers
    else:
        df_data = df_raw.copy()
        col_sample_idx = 3

    if col_date_idx is None and col_type_idx is not None and col_type_idx > 0:
        col_date_idx = col_type_idx - 1

    return df_raw, df_data, fila_encabezado, col_sample_idx, col_type_idx, col_date_idx

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
            nombre_elem_clean = elem.split("[")[0].strip() if "[" in elem else elem.strip()
            nombre_elem = nombre_elem_clean if nombre_elem_clean else f"Col_{c_idx}"
            cols_interes.append((c_idx, nombre_elem, True, unit))

    return cols_interes

def calcular_resultados(df_data, cols_interes, col_sample_idx, muestras_elegidas, blancos_seleccionados, df_params):
    promedio_blancos_ppb = {}
    if blancos_seleccionados:
        col_serie = df_data.iloc[:, col_sample_idx].astype(str).str.strip()
        df_blancos = df_data[col_serie.isin(blancos_seleccionados)]

        for c_idx, _, es_conc, unidad in cols_interes:
            if es_conc:
                valores_ppb = []
                for _, fila in df_blancos.iterrows():
                    v_str = str(fila.iloc[c_idx]).replace("<", "").replace(">", "").replace(",", ".").strip()
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
        p_m = params_dict.get(nombre_muestra, {"Masa (mg)": 15.0, "Volumen (mL)": 10.0})

        masa_mg = float(p_m["Masa (mg)"])
        vol_ml = float(p_m["Volumen (mL)"])

        row_dict = {
            "Sample Name": nombre_muestra,
            "Masa (mg)": masa_mg,
            "Volumen (mL)": vol_ml,
        }

        for c_idx, nombre_elem, es_conc, unidad in cols_interes:
            if es_conc:
                v_str = str(fila.iloc[c_idx]).replace("<", "").replace(">", "").replace(",", ".").strip()
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
            return pd.read_excel(excel_obj, sheet_name=excel_obj.sheet_names[0])
    except Exception as e:
        st.error(f"Error: {e}")
        return None

# ==========================================
# 5. CONTROL DE SESIÓN Y SIDEBAR
# ==========================================
if "usuario_logueado" not in st.session_state:
    st.session_state.usuario_logueado = None

with st.sidebar:
    st.header("👤 Control de Acceso")
    st.caption(t["db_type_supabase"] if supabase_client else t["db_type_sqlite"])
    
    if st.session_state.usuario_logueado is None:
        tab_log, tab_reg = st.tabs([t["login_tab"], t["register_tab"]])
        
        with tab_log:
            u_in = st.text_input(t["user_label"], key="login_user")
            p_in = st.text_input(t["pass_label"], type="password", key="login_pass")
            if st.button(t["login_btn"], type="primary", key="btn_login"):
                if verificar_usuario(u_in, p_in):
                    st.session_state.usuario_logueado = u_in.strip()
                    st.success(t["login_success"].format(u_in.strip()))
                    st.rerun()
                else:
                    st.error(t["login_error"])
                    
        with tab_reg:
            u_reg = st.text_input(t["user_label"], key="reg_user")
            p_reg = st.text_input(t["pass_label"], type="password", key="reg_pass")
            if st.button(t["register_btn"], key="btn_reg"):
                ok, msg = registrar_usuario(u_reg, p_reg)
                if ok:
                    st.success(msg)
                else:
                    st.error(msg)
    else:
        st.success(f"🟢 **{t['user_label']}:** `{st.session_state.usuario_logueado}`")
        if st.button(t["logout_btn"], type="primary"):
            st.session_state.usuario_logueado = None
            st.rerun()

    st.divider()
    st.header(t["about_header"])
    st.markdown("""
        **Analizador ICP-MS**  
        * **Desarrollador:** Pedro J. Navarrete Segado  
        * **Institución:** Universidad de Jaén (UJA)  
        * **Contacto:** [pnsegado@ujaen.es](mailto:pnsegado@ujaen.es)  
        * **Repositorio:** [GitHub](https://github.com/PedroJNS/icpms-data-processing)  
        * **Licencia:** GNU GPL v3.0  
        """)

# ==========================================
# 6. VISTA PRINCIPAL
# ==========================================
st.title(t["title"])
st.caption(t["subtitle"])

tab_analisis, tab_historico, tab_cargar_excel = st.tabs([
    t["tab_analysis"], 
    t["tab_history"], 
    t["tab_load_excel"]
])

# --- PESTAÑA 1: NUEVO ANÁLISIS ---
with tab_analisis:
    unidad_medida = st.radio(
        t["unit_selector"],
        options=["% wt (Porcentaje en Peso)", "ppm (Partes por Millón)"],
        horizontal=True,
    )
    es_ppm = "ppm" in unidad_medida

    uploaded_file = st.file_uploader(t["upload_file"], type=["xlsx", "xls", "csv"])

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
            
            cols_interes = preparar_mapeo_columnas(df_raw, fila_encabezado, col_sample_idx)

            palabras_ignorar = [
                "blank", "blanco", "ppb", "calblk", "calstd", "blkvrfy",
                "qc", "driftchk", "cicspike", "isostd", "dilstd", "bkgnd", "fqblk",
            ]

            if col_type_idx is not None and fila_encabezado is not None:
                val_type = df_data.iloc[:, col_type_idx].astype(str).str.strip().str.lower()
                df_muestras_raw = df_data[val_type == "sample"]
            else:
                df_muestras_raw = df_data

            todas_muestras = df_muestras_raw.iloc[:, col_sample_idx].dropna().astype(str).tolist()

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
                    t["select_samples"],
                    options=muestras_validas,
                    default=muestras_validas,
                )

            with col_sel2:
                blancos_seleccionados = st.multiselect(
                    t["select_blanks"],
                    options=blancos_detectados,
                    default=blancos_detectados,
                )

            if not muestras_elegidas:
                st.warning(t["warn_select_sample"])
            else:
                st.divider()
                st.subheader(t["digestion_params"])
                st.info(t["params_hint"])

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
                        "Masa (mg)": st.column_config.NumberColumn(min_value=0.001, format="%.2f"),
                        "Volumen (mL)": st.column_config.NumberColumn(min_value=0.001, format="%.2f"),
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
                st.subheader(f"{t['results_table']} {suff}")

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
                st.subheader(t["save_analysis_title"])
                
                usuario_activo = st.session_state.usuario_logueado
                if usuario_activo:
                    col_db1, col_db2 = st.columns([2, 1])
                    with col_db1:
                        nombre_analisis_input = st.text_input(
                            t["analysis_name_label"],
                            value=f"Análisis {uploaded_file.name}",
                        )
                    with col_db2:
                        st.write(" ")
                        st.write(" ")
                        if st.button(t["save_btn"], type="secondary"):
                            if guardar_analisis_db(usuario_activo, nombre_analisis_input, df_results_base):
                                st.success(t["save_success"].format(nombre_analisis_input))
                            else:
                                st.error(t["save_error"])
                else:
                    st.info(t["login_req_save"])

                excel_proyecto_bytes = generar_excel_proyecto(df_results_base, df_results_display)

                st.download_button(
                    label=t["export_excel"],
                    data=excel_proyecto_bytes,
                    file_name=f"Proyecto_ICPMS_{'ppm' if es_ppm else 'pct'}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )

                # --- SECCIÓN GRÁFICA CORREGIDA ---
                st.divider()
                st.subheader(f"{t['chart_title']} {suff}")

                cols_metales = [c for c in df_results_display.columns if suff in c]

                if cols_metales:
                    col_g1, col_g2 = st.columns([1, 3])

                    with col_g1:
                        usar_log = st.checkbox(t["use_log"], value=False)
                        nombres_metales = [c.replace(f" {suff}", "") for c in cols_metales]
                        metales_sel = st.multiselect(
                            t["select_metals"],
                            options=nombres_metales,
                            default=nombres_metales[: min(5, len(nombres_metales))],
                        )

                    with col_g2:
                        if metales_sel:
                            cols_graficar = [f"{m} {suff}" for m in metales_sel]
                            df_plot = df_results_display.set_index("Sample Name")[cols_graficar]
                            df_plot.columns = [c.replace(f" {suff}", "") for c in df_plot.columns]

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
                                fontweight="bold"
                            )
                            ax.grid(axis="y", linestyle="--", alpha=0.5)
                            plt.xticks(rotation=45, ha="right")
                            plt.tight_layout()
                            st.pyplot(fig)

        except Exception as e:
            st.error(f"Error procesando archivo: {e}")

# --- PESTAÑA 2: HISTÓRICO BASE DE DATOS ---
with tab_historico:
    usuario_act = st.session_state.usuario_logueado
    if usuario_act:
        st.subheader(f"📂 Análisis Guardados de {usuario_act}")
        lista_analisis = obtener_analisis_usuario(usuario_act)
        if lista_analisis:
            opciones = {f"{a[1]} ({a[2]})": a[0] for a[a] in lista_analisis} if len(lista_analisis) > 0 and isinstance(lista_analisis[0], tuple) else {f"{a[1]} ({a[2]})": a[0] for a in lista_analisis}
            item_sel = st.selectbox("Selecciona un análisis para cargar:", list(opciones.keys()))
            if item_sel:
                analisis_id = opciones[item_sel]
                df_cargado = cargar_analisis_db(analisis_id, usuario_act)
                if df_cargado is not None:
                    st.dataframe(df_cargado, use_container_width=True)
        else:
            st.info("No tienes análisis guardados aún.")
    else:
        st.warning(t["login_req_save"])

# --- PESTAÑA 3: CARGAR PROYECTO EXCEL ---
with tab_cargar_excel:
    st.subheader(t["tab_load_excel"])
    uploaded_proj = st.file_uploader("Selecciona un archivo Excel de proyecto exportado anteriormente:", type=["xlsx"])
    if uploaded_proj:
        df_proj = cargar_desde_excel_proyecto(uploaded_proj)
        if df_proj is not None:
            st.success("¡Proyecto cargado con éxito!")
            st.dataframe(df_proj, use_container_width=True)
