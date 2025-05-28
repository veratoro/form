import streamlit as st
from PIL import Image
from fill_proveedores.form_proveedor import formularioExternos
from fill_proveedores.form_colaborador import seleccion

# CONFIGURACIÓN DE LA PÁGINA
st.set_page_config(
    page_title="Inscripción Proveedores",
    page_icon=Image.open(r".\images\e.png"),
    layout="wide",
    initial_sidebar_state="expanded"
)


def check_password():
    def login_form():

        col1, col2, col3 = st.columns((0.3, 0.28, 0.3))
        with col2:

            st.write("")
            st.write("")

            # Imagen superior
            col4, col5, col6 = st.columns((0.4, 0.3, 0.4))
            with col5:
                st.image(Image.open(r".\images\elede.png"), use_container_width=True)

            st.write("")
            st.write("")

            # Formulario de credenciales
            with st.form("Credentials"):
                st.markdown(
                    "<div style='text-align: center; color: #003366; font-size: 25px; font-weight: bold;'>Portal de Acceso</div>",
                    unsafe_allow_html=True,
                )
                st.markdown(
                    "<div style='text-align: center; color: #003366; font-size: 15px;'>Seleccione su tipo de usuario para ingresar</div>",
                    unsafe_allow_html=True,
                )

                st.write("")
                st.write("")

                # Botones para seleccionar el tipo de usuario
                left, right = st.columns(2)
                if left.form_submit_button("**Proveedor**", icon=":material/add_reaction:", use_container_width=True):
                    st.session_state["tipo_usuario"] = "Proveedor"                
                if right.form_submit_button("**Colaborador**", icon=":material/mood:", use_container_width=True):
                    st.session_state["tipo_usuario"] = "Colaborador"


                # Mostrar el tipo de usuario seleccionado
                tipo_usuario = st.session_state.get("tipo_usuario", None)
                #if tipo_usuario:
                #    st.write(f"**Tipo de usuario seleccionado:** {tipo_usuario}")
                #else:
                #    st.write("**Tipo de usuario seleccionado:** Ninguno")

                #st.write("")
                #st.write("")

                # Campo de contraseña
                #st.markdown("""<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.5/font/bootstrap-icons.css">""", unsafe_allow_html=True)
                #st.markdown(
                #    """<div style="display: flex; align-items: center;"><i class="bi bi-lock" style="font-size: 1rem; color: navy; margin-right: 10px;"></i><span style="font-size: 1rem; color: #003366;">Contraseña</span></div>""",
                #    unsafe_allow_html=True,
                #)

                
                #st.markdown("""<div style="display: flex; align-items: center;"><i class="bi bi-lock" style="font-size: 1rem; color: navy; margin-right: 10px;"></i><span style="font-size: 1rem; color: #003366;">Contraseña</span></div>""", unsafe_allow_html=True)



                st.markdown("""
                <link href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined" rel="stylesheet">
                <style>
                .material-symbols-outlined {
                font-family: 'Material Symbols Outlined';
                font-variation-settings: 'FILL' 0, 'wght' 400, 'GRAD' 0, 'opsz' 24;
                font-size: 15px;
                vertical-align: middle;
                color: #003366;
                margin-right: 1px;
                margin-left:5px;
                }
                .icon-text-container {
                display: flex;
                align-items: center;
                gap: 6px;
                color: #003366;  /* Aplica navy al texto también */
                }
                </style>
                """, unsafe_allow_html=True)
                st.markdown("""<div class="icon-text-container"><span class="material-symbols-outlined">lock</span><span><strong>Contraseña</strong></span></div>""", unsafe_allow_html=True)
                
                #st.markdown("""<span class="material-symbols-outlined">lock</span> <strong>Contraseña</strong>""", unsafe_allow_html=True)
                


                st.text_input(
                    "Contraseña",
                    type="password",
                    key="clave",
                    label_visibility="collapsed",
                    placeholder="Ingrese su contraseña",
                    help="Ingrese su contraseña",
                )

                # Botón de inicio de sesión
                if st.form_submit_button("Ingresar", use_container_width=True, type="primary"):
                    password_entered()

                # Estilo del formulario
                st.markdown(
                    """
                    <style>
                    [data-testid="stForm"] {
                        box-shadow: 0px 0px 10px rgba(17, 22, 89, 0.5);
                        border-radius: 15px;
                    }
                    </style>
                    """,
                    unsafe_allow_html=True,
                )

    def password_entered():
        clave = st.session_state.get("clave", "")
        tipo_usuario = st.session_state.get("tipo_usuario", None)

        # Validar credenciales
        credenciales = {"Colaborador": "456", "Proveedor": "123"}
        if tipo_usuario in credenciales and clave == credenciales[tipo_usuario]:
            st.session_state["password_correct"] = True
            if tipo_usuario == "Proveedor":
                st.session_state["redirect_to_proveedor"] = True
                #st.experimental_rerun()
            elif tipo_usuario == "Colaborador":
                st.session_state["redirect_to_colaborador"] = True
                #st.experimental_rerun()
        else:
            # Cargar la fuente y definir los estilos
            st.markdown("""
            <link href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined" rel="stylesheet">
            <style>
            .alert-error {
            display: flex;
            align-items: center;
            background-color: #fdecea;        /* fondo similar al de st.error */
            color: #611a15;                   /* color de texto e ícono */
            border: 1px solid #f5c6cb;
            border-radius: 5px;
            padding: 10px;
            margin-top: 10px;
            }
            .material-symbols-outlined {
            font-family: 'Material Symbols Outlined';
            font-variation-settings: 'FILL' 1, 'wght' 400, 'GRAD' 0, 'opsz' 24;
            font-size: 20px;
            margin-right: 10px;
            color: inherit;                  /* hereda el color del contenedor */
            }
            </style>

            <div class="alert-error">
            <span class="material-symbols-outlined">error</span>
            <span>Usuario o contraseña incorrectos</span>
            </div>
            """, unsafe_allow_html=True)
            st.session_state["password_correct"] = False

    # Mostrar el formulario
    login_form()

    # Verificar si la contraseña es correcta
    if not st.session_state.get("password_correct", False):
        return False
    return True

def logout():
    """Logs out the user by clearing session data and redirecting to login page."""
    st.write('<meta http-equiv="refresh" content="2;URL=/"></head>', unsafe_allow_html=True)




# Redirigir a la página del proveedor si es necesario
if st.session_state.get("redirect_to_proveedor", False):
    
    e_l, e_t, e_bt = st.columns((0.1, 0.8 ,0.1))
    with e_l:
        st.image(Image.open(r".\images\elede.png"), use_container_width=True)
    with e_t:
        st.markdown(
                    "<div style='text-align: center; color: #003366; font-size: 25px; font-weight: bold;'>Formulario de Inscripción - Proveedor Nacional</div>",
                    unsafe_allow_html=True,
                )

    with e_bt:
        if st.button("Salir", icon =":material/logout:", type="primary"):
            logout()   
    e1, c, e2 = st.columns((0.2,1,0.2))
    with c:
        formularioExternos()
        st.stop()




# Redirigir a la página del colaborador si es necesario
if st.session_state.get("redirect_to_colaborador", False):
    e_l, e_t, e_bt = st.columns((0.1, 0.8 ,0.1))
    with e_l:
        st.image(Image.open(r".\images\elede.png"), use_container_width=True)
    with e_t:
        st.markdown(
                    "<div style='text-align: center; color: #003366; font-size: 25px; font-weight: bold;'>Formulario de Inscripción - Proveedor Nacional</div>",
                    unsafe_allow_html=True,
                )

    with e_bt:
        if st.button("Salir", icon =":material/logout:", type="primary"):
            logout()   
    e1, c, e2 = st.columns((0.2,1,0.2))
    with c:
        seleccion()
        st.stop()

if not check_password():
    st.stop()

def logout():
    """Logs out the user by clearing session data and redirecting to login page."""
    st.write('<meta http-equiv="refresh" content="2;URL=/"></head>', unsafe_allow_html=True)

