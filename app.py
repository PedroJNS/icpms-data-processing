import streamlit as st
from supabase import create_client, Client

# ==========================================
# CONFIGURACIÓN E INICIALIZACIÓN DE SUPABASE
# ==========================================
st.set_page_config(page_title="ICP-MS App", page_icon="🧪", layout="wide")

@st.cache_resource
def init_supabase() -> Client:
    """Inicializa la conexión externa a Supabase."""
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

try:
    supabase = init_supabase()
except Exception as e:
    st.error(f"Error conectando a Supabase: {e}")
    st.stop()

# ==========================================
# SISTEMA DE DICCIONARIO PARA MULTI-IDIOMA
# ==========================================
TEXTS = {
    "es": {
        "title": "🧪 Sistema ICP-MS",
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
        "welcome_header": "Panel Principal ICP-MS",
        "logout_btn": "Cerrar Sesión",
        "user_list_title": "Lista de Usuarios Registrados",
        "no_users": "No hay usuarios registrados.",
        "fill_all": "Por favor completa todos los campos."
    },
    "en": {
        "title": "🧪 ICP-MS System",
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
        "welcome_header": "ICP-MS Main Dashboard",
        "logout_btn": "Log Out",
        "user_list_title": "Registered Users List",
        "no_users": "No registered users found.",
        "fill_all": "Please fill in all fields."
    }
}

# Selector de Idioma en la barra lateral
with st.sidebar:
    lang_choice = st.selectbox("🌐 Language / Idioma", ["Español", "English"])
    lang = "es" if lang_choice == "Español" else "en"
    t = TEXTS[lang]

# ==========================================
# FUNCIONES DE BASE DE DATOS (SUPABASE)
# ==========================================
def obtener_usuario(nombre: str):
    res = supabase.table("usuarios").select("*").eq("nombre", nombre).execute()
    return res.data[0] if res.data else None

def registrar_usuario(nombre: str, clave_hash: str):
    datos = {"nombre": nombre, "clave_hash": clave_hash}
    return supabase.table("usuarios").insert(datos).execute()

def obtener_todos_usuarios():
    res = supabase.table("usuarios").select("id, nombre").execute()
    return res.data if res.data else []

# ==========================================
# GESTIÓN DE SESIÓN Y VISTAS
# ==========================================
if "usuario_logueado" not in st.session_state:
    st.session_state.usuario_logueado = None

st.title(t["title"])

if st.session_state.usuario_logueado is None:
    tab_login, tab_register = st.tabs([t["login_tab"], t["register_tab"]])
    
    # TAB: INICIAR SESIÓN
    with tab_login:
        usuario_in = st.text_input(t["user_label"], key="login_user")
        clave_in = st.text_input(t["pass_label"], type="password", key="login_pass")
        
        if st.button(t["login_btn"], key="btn_login"):
            if usuario_in and clave_in:
                user = obtener_usuario(usuario_in)
                # NOTA: En producción se recomienda comparar hash (ej. bcrypt / hashlib)
                if user and user.get("clave_hash") == clave_in:
                    st.session_state.usuario_logueado = user["nombre"]
                    st.success(t["login_success"].format(user["nombre"]))
                    st.rerun()
                else:
                    st.error(t["login_error"])
            else:
                st.warning(t["fill_all"])

    # TAB: REGISTRO
    with tab_register:
        usuario_reg = st.text_input(t["user_label"], key="reg_user")
        clave_reg = st.text_input(t["pass_label"], type="password", key="reg_pass")
        
        if st.button(t["register_btn"], key="btn_reg"):
            if usuario_reg and clave_reg:
                user_existente = obtener_usuario(usuario_reg)
                if user_existente:
                    st.error(t["user_exists"])
                else:
                    registrar_usuario(usuario_reg, clave_reg)
                    st.success(t["register_success"])
            else:
                st.warning(t["fill_all"])

else:
    # PANEL PRINCIPAL (USUARIO CONECTADO)
    st.sidebar.markdown(f"**{t['user_label']}:** `{st.session_state.usuario_logueado}`")
    if st.sidebar.button(t["logout_btn"]):
        st.session_state.usuario_logueado = None
        st.rerun()
        
    st.header(t["welcome_header"])
    
    st.subheader(t["user_list_title"])
    usuarios = obtener_todos_usuarios()
    if usuarios:
        st.dataframe(usuarios, use_container_width=True)
    else:
        st.info(t["no_users"])
