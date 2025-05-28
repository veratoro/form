import streamlit as st
import pandas as pd
import os
from st_aggrid import AgGrid
from st_aggrid import AgGrid, GridUpdateMode, JsCode
from st_aggrid.grid_options_builder import GridOptionsBuilder
import json
import xlwings as xw
import shutil
import os
import win32com.client as win32
import os
import smtplib
from email.message import EmailMessage
from email.utils import make_msgid



def seleccion():

    st.markdown("""---""")
    #st.write('')
    #st.write('')

    seleccion_fecha = None
    seleccion_nombre = None
    seleccion_identificacion = None

    st.info("""
    **Instrucciones:**

    **1.** Identifique en el listado al proveedor para el cual desea realizar la inscripción.  
    **2.** Si no se encuentra en el listado, solicite al proveedor que diligencie previamente el formulario de inscripción.  
    **3.** Una vez localizado, seleccione la casilla correspondiente y continúe con el diligenciamiento del formulario.
    """)

    # Ruta de la carpeta principal
    ruta_principal = './fill_proveedores/proveedores_inscritos'

    # Obtener solo los nombres de las subcarpetas (no archivos)
    nombres_subcarpetas = [
        nombre for nombre in os.listdir(ruta_principal)
        if os.path.isdir(os.path.join(ruta_principal, nombre))
    ]

    # Paso 3: Transformar los nombres
    nombres_transformados = []

    for nombre in nombres_subcarpetas:
        try:
            # Separar fecha y hora
            fecha = nombre[0:4] + '-' + nombre[4:6] + '-' + nombre[6:8] + ' ' + nombre[9:11] + ':' + nombre[11:13] + ':' + nombre[13:15]
            
            # Extraer el resto después del timestamp
            resto = nombre[16:]
            partes = resto.split('_')
            
            if len(partes) >= 2:
                nombre_completo = ' '.join(partes[:-1])
                documento = partes[-1]
            else:
                nombre_completo = partes[0]
                documento = ''
            
            nueva_linea = f"{fecha},{nombre_completo},{documento}"
            nombres_transformados.append(nueva_linea)
        
        except Exception as e:
            print(f"Error al procesar: {nombre} -> {e}")


    # Separar los elementos por coma
    registros = [fila.split(',') for fila in nombres_transformados]

    # Crear el DataFrame
    df = pd.DataFrame(registros, columns=['FECHA DE DILIGENCIAMIENTO', 'NOMBRE / RAZÓN SOCIAL', 'IDENTIFICACIÓN'])
    df = df.sort_values(by="FECHA DE DILIGENCIAMIENTO", ascending=False)


    gd = GridOptionsBuilder.from_dataframe(df)
    gd.configure_pagination(enabled = True)
    gd.configure_default_column(groupable = True)
    gd.configure_column("FECHA DE DILIGENCIAMIENTO", editable = False)
    gd.configure_column("NOMBRE / RAZÓN SOCIAL", editable = False)
    gd.configure_column("IDENTIFICACIÓN", editable = False)
    gd.configure_selection(selection_mode = 'single', use_checkbox = True)
    gridOptions = gd.build()


    grid_table = AgGrid(df,
                        gridOptions = gridOptions,
                        fit_columns_on_grid_load=True,
                        height=261,
                        width='100%',
                        theme='balham',
                        update_mode=GridUpdateMode.GRID_CHANGED,
                        reload_data = True,
                        editable=True
                        )
    
    proveedor_seleccionado = grid_table.selected_rows_id
    
    #st.write(grid_table.selected_data)
    #st.write(grid_table.selected_rows)
    #st.write(grid_table.selected_rows_id)


    if proveedor_seleccionado == None:
        st.write("")


    if proveedor_seleccionado != None:


        # SI SE SELECCIONA UNA FILA, ENTONCES EJECUTE LO SIGUIENTE:

        seleccion_fecha = grid_table['selected_rows']['FECHA DE DILIGENCIAMIENTO'].tolist()
        seleccion_fecha = [f.replace("-", "").replace(" ", "_").replace(":", "") for f in seleccion_fecha]

        seleccion_nombre = grid_table['selected_rows']['NOMBRE / RAZÓN SOCIAL'].tolist()
        seleccion_nombre = [n.replace(" ", "_") for n in seleccion_nombre]

        seleccion_identificacion = grid_table['selected_rows']['IDENTIFICACIÓN'].tolist()

        proveedor_seleccionado = f"{seleccion_fecha[0]}_{seleccion_nombre[0]}_{seleccion_identificacion[0]}"
        
        fe1, fe, fe2 = st.columns((0.8, 0.1, 0.8))
        with fe:

            st.markdown("""
                <link href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined" rel="stylesheet">
                <style>
                .material-symbols-outlined {
                font-family: 'Material Symbols Outlined';
                font-variation-settings: 'FILL' 0, 'wght' 400, 'GRAD' 0, 'opsz' 24;
                font-size: 35px;
                vertical-align: top;
                color: #003366;
                margin-right: 1px;
                margin-left:5px;
                }
                .icon-text-container {
                display: flex;
                justify-content: center;  /* Centra horizontalmente */
                align-items: center;
                gap: 6px;
                color: #003366;  /* Aplica navy al texto también */
                }
                </style>
                """, unsafe_allow_html=True)
            st.markdown("""<div class="icon-text-container"><span class="material-symbols-outlined">keyboard_arrow_down</span></div>""", unsafe_allow_html=True)
            
        #st.info(f'**Inscripción del proveedor: {seleccion_nombre[0]}**', icon=":material/thumb_up:" )
        #st.write('')
        #st.write('')


        ruta_subcarpeta = os.path.join(ruta_principal, proveedor_seleccionado)
        archivo_variables = os.path.join(ruta_subcarpeta, 'variables_proveedor.txt')

        # Abrir y leer el archivo
        with open(archivo_variables, 'r', encoding='utf-8') as archivo:
            contenido = archivo.read()

        # Mostrar el contenido del archivo
        #st.write("Contenido del archivo 'variables_proveedor':")
        #st.text(contenido)



        ###############################################################################################

        #st.markdown("""---""")
        st.write('')

        #----------------------------
        empleado_solicitante = None
        area_solicitante = None
        gerente_area_solicitante = None
        empresa = None
        caracterizacion_sustentabilidad = None
        empleado_caracterizador = None
        tipo_negocio = None
        plazo_pago = None
        descripcion_negocio = None
        servicio_prestado_en_sede_de_la_compania = None
        sede_servicio = []
        otro_tipo_proveedor = None
        a = []
        b = []
        nodo_1 = None
        nodo_2 = None
        tipo_proveedor = []


        #----------------------------
        


        col3, col4, col5 = st.columns((1,0.8,1))
        with col3:
            empleado_solicitante = st.text_input('EMPLEADO SOLICITANTE *', autocomplete="off").upper()
        with col4:
            area_solicitante = st.text_input('ÁREA *', autocomplete="off").upper()
        with col5:
            gerente_area_solicitante = st.text_input('GERENTE DEL ÁREA *', autocomplete="off").upper()



        col6, col7 = st.columns(2)
        with col6:
            empresa = st.multiselect('EMPRESA PARA LA CUAL TRAMITA LA INSCRIPCIÓN DEL PROVEEDOR *', ['LINEA DIRECTA', 'ELEDÉ CASA', 'GRUPO ELEDÉ'])
        


        col8, col9 = st.columns(2)
        with col8:
            caracterizacion_sustentabilidad = st.segmented_control('¿EL PROVEEDOR TIENE CARACTERIZACIÓN DE SUSTENTABILIDAD? *', ['SI', 'NO'], key = 'sustentabilidad')
        if caracterizacion_sustentabilidad == 'SI':
            with col9:
                empleado_caracterizador = st.text_input('EMPLEADO QUE EFECTUÓ LA CARACTERIZACIÓN *', autocomplete="off").upper()



        col10, col11 = st.columns((0.7,0.3))
        with col10:
            tipo_negocio = st.segmented_control('TIPO DE NEGOCIO A OFRECER *', ['VENTA DE PRODUCTOS', 'PRESTACIÓN DE SERVICIOS', 'PRODUCTOS Y SERVICIOS'])
        with col11:
            plazo_pago = st.selectbox("PLAZO DE PAGO EN DÍAS *",("15", "30", "45"), index=None, placeholder= 'DE ACUERDO CON LA LEY 2024 DE 2020')



        col14, col15 = st.columns(2)
        if tipo_negocio == 'VENTA DE PRODUCTOS':
            with col14:
                descripcion_negocio = st.text_input('DESCRIPCIÓN DEL PRODUCTO A SUMINISTRAR *', autocomplete="off")
        if tipo_negocio == 'PRESTACIÓN DE SERVICIOS':
            with col14:
                descripcion_negocio = st.text_input('DESCRIPCIÓN DEL SERVICIO A SUMINISTRAR *', autocomplete="off")
        if tipo_negocio == 'PRODUCTOS Y SERVICIOS':
            with col14:
                descripcion_negocio = st.text_input('DESCRIPCIÓN DEL PRODUCTO Y SERVICIO A SUMINISTRAR *', autocomplete="off")



        if (tipo_negocio == 'PRESTACIÓN DE SERVICIOS') or (tipo_negocio == 'PRODUCTOS Y SERVICIOS'):    
            col6, col7 = st.columns(2)
            with col6:
                servicio_prestado_en_sede_de_la_compania = st.segmented_control('¿EL SERVICIO SERÁ PRESTADO EN ALGUNA DE LAS SEDES DE LA COMPAÑÍA? *', ['SI', 'NO'], key = 'sede')
            if servicio_prestado_en_sede_de_la_compania == 'SI':
                with col7:
                    sede_servicio = st.multiselect('SEDE EN LA QUE SE PRESTARÁ EL SERVICIO *', ['B ESCOCIA', 'CIGMO', 'HACIENDA'])



        st.write('')
        st.write('')
        st.write('')



        if tipo_negocio == 'VENTA DE PRODUCTOS':

            # Crear columnas para organizar las categorías
            col1, col2, col3, col4 = st.columns(4)

            # Suministro de Bienes
            with col1:
                st.write("**Suministro de Bienes**")
                cuidado_hogar = st.checkbox("Cuidado hogar")
                cuidado_personal = st.checkbox("Cuidado personal")
                ropa_hogar = st.checkbox("Ropa hogar")
                empaques = st.checkbox("Envases y empaques")

            # Manufacturas
            with col2:
                st.write("**Manufacturas**")
                telas = st.checkbox("Telas")
                insumos = st.checkbox("Insumos")
                producto_terminado = st.checkbox("Producto terminado")

            # Alimentación/Hotel
            with col3:
                st.write("**Alimentación/Hotel**")
                alimentos = st.checkbox("Venta de alimentos o bebidas")

            # Otro
            with col4:
                otro = st.checkbox("Otro")
                if otro == True:
                    otro_tipo_proveedor = st.text_input(' ', placeholder = '¿CUÁL? *', label_visibility = 'collapsed', autocomplete="off").upper()

            a = ['cuidado_hogar','cuidado_personal','ropa_hogar','empaques','telas','insumos','producto_terminado','alimentos','otro']
            b = [cuidado_hogar,cuidado_personal,ropa_hogar,empaques,telas,insumos,producto_terminado,alimentos,otro]



        if tipo_negocio == 'PRESTACIÓN DE SERVICIOS':

            # Crear columnas para organizar las categorías
            col1, col2, col3, col4 = st.columns(4)

            # Manufacturas
            with col1:
                st.write("**Manufacturas**")
                estampacion = st.checkbox("Estampación")
                tintoreria = st.checkbox("Lavado y/o tintorería")
                sublimado = st.checkbox("Sublimado")
                confeccion = st.checkbox("Confección")
                tejeduria = st.checkbox("Tejeduría")
                corte = st.checkbox("Corte")
                bordado = st.checkbox("Bordado")
                prehormado = st.checkbox("Prehormado")

            if confeccion == True:
                col7, col8, col9, col10 = st.columns((0.7,0.7,0.3,0.3))
                with col7:
                    nodo_1 = st.text_input(' ', placeholder = 'NODO 1 *', label_visibility = 'collapsed', autocomplete="off").upper()
                with col8:
                    nodo_2 = st.text_input(' ', placeholder = 'NODO 2 *', label_visibility = 'collapsed', autocomplete="off").upper()

            # Arrendamiento
            with col2:
                st.write("**Arrendamiento**")
                muebles = st.checkbox("Muebles")
                inmuebles = st.checkbox("Inmuebles")

            # Alimentación/Hotel
            with col3:
                st.write("**Alimentación/Hotel**")
                catering = st.checkbox("Catering")
                alojamiento = st.checkbox("Alojamiento")

            # Actividad profesional y técnica
            with col4:
                st.write("**Actividad profesional y técnica**")
                consultorias = st.checkbox("Honorarios o consultorías")
                mantenimiento = st.checkbox("Técnicos y mantenimiento")
                tic = st.checkbox("Actividades TIC")
                fotografia = st.checkbox("Fotografía, modelos")
                st.write('')
                st.write('')
                st.write('')
                otro = st.checkbox("**Otro**")
                if otro == True:
                    otro_tipo_proveedor = st.text_input(' ', placeholder = '¿CUÁL? *', label_visibility = 'collapsed', autocomplete="off").upper()

            a = ['estampacion','tintoreria','sublimado','confeccion','tejeduria','corte','bordado','prehormado','muebles','inmuebles','catering','alojamiento','consultorias','mantenimiento','tic','fotografia','otro']
            b = [estampacion,tintoreria,sublimado,confeccion,tejeduria,corte,bordado,prehormado,muebles,inmuebles,catering,alojamiento,consultorias,mantenimiento,tic,fotografia,otro]



        if tipo_negocio == 'PRODUCTOS Y SERVICIOS':

            # Crear columnas para organizar las categorías
            col1, col2, col3, col4, col5 = st.columns(5)

            # Suministro de Bienes
            with col1:
                st.write("**Suministro de Bienes**")
                cuidado_hogar = st.checkbox("Cuidado hogar")
                cuidado_personal = st.checkbox("Cuidado personal")
                ropa_hogar = st.checkbox("Ropa hogar")
                empaques = st.checkbox("Envases y empaques")

            # Manufacturas
            with col2:
                st.write("**Manufacturas**")
                telas = st.checkbox("Telas")
                insumos = st.checkbox("Insumos")
                producto_terminado = st.checkbox("Producto terminado")
                estampacion = st.checkbox("Estampación")
                tintoreria = st.checkbox("Lavado y/o tintorería")
                sublimado = st.checkbox("Sublimado")
                confeccion = st.checkbox("Confección")
                tejeduria = st.checkbox("Tejeduría")
                corte = st.checkbox("Corte")
                bordado = st.checkbox("Bordado")
                prehormado = st.checkbox("Prehormado")

            if confeccion == True:
                col7, col8, col9, col10 = st.columns((0.7,0.7,0.2,0.2))
                with col7:
                    nodo_1 = st.text_input(' ', placeholder = 'NODO 1 *', label_visibility = 'collapsed', autocomplete="off").upper()
                with col8:
                    nodo_2 = st.text_input(' ', placeholder = 'NODO 2 *', label_visibility = 'collapsed', autocomplete="off").upper()

            # Arrendamiento
            with col3:
                st.write("**Arrendamiento**")
                muebles = st.checkbox("Muebles")
                inmuebles = st.checkbox("Inmuebles")

            # Alimentación/Hotel
            with col4:
                st.write("**Alimentación/Hotel**")
                alimentos = st.checkbox("Venta de alimentos o bebidas")
                catering = st.checkbox("Catering")
                alojamiento = st.checkbox("Alojamiento")

            # Actividad profesional y técnica
            with col5:
                st.write("**Actividad profesional y técnica**")
                consultorias = st.checkbox("Honorarios o consultorías")
                mantenimiento = st.checkbox("Técnicos y mantenimiento")
                tic = st.checkbox("Actividades TIC")
                fotografia = st.checkbox("Fotografía, modelos")
                st.write('')
                st.write('')
                st.write('')
                otro = st.checkbox("**Otro**")
                if otro == True:
                    otro_tipo_proveedor = st.text_input(' ', placeholder = '¿CUÁL? *', label_visibility = 'collapsed', autocomplete="off").upper()

            a = ['cuidado_hogar','cuidado_personal','ropa_hogar','empaques','telas','insumos','producto_terminado','estampacion','tintoreria','sublimado','confeccion','tejeduria','corte','bordado','prehormado','muebles','inmuebles','alimentos','catering','alojamiento','consultorias','mantenimiento','tic','fotografia','otro']
            b = [cuidado_hogar,cuidado_personal,ropa_hogar,empaques,telas,insumos,producto_terminado,estampacion,tintoreria,sublimado,confeccion,tejeduria,corte,bordado,prehormado,muebles,inmuebles,alimentos,catering,alojamiento,consultorias,mantenimiento,tic,fotografia,otro]



        tipo_proveedor = [categoria for categoria, flag in zip(a, b) if flag]
        

        st.write('')



        ###########################################



        # VALIDACIÓN
        formulario_completo = all([
            empleado_solicitante,
            area_solicitante,
            gerente_area_solicitante,
            len(empresa) > 0,  
            caracterizacion_sustentabilidad is not None,
            tipo_negocio is not None,
            plazo_pago is not None,
            len(tipo_proveedor) > 0
        ])



        # VALIDACIONES CONDICIONALES

        if caracterizacion_sustentabilidad == 'SI':
            formulario_completo &= bool(empleado_caracterizador)


        if tipo_negocio in ['VENTA DE PRODUCTOS', 'PRESTACIÓN DE SERVICIOS', 'PRODUCTOS Y SERVICIOS']:
            formulario_completo &= bool(descripcion_negocio)

            if (tipo_negocio == 'PRESTACIÓN DE SERVICIOS') or (tipo_negocio == 'PRODUCTOS Y SERVICIOS'): 
                formulario_completo &= bool(servicio_prestado_en_sede_de_la_compania)

                if servicio_prestado_en_sede_de_la_compania == 'SI':
                    formulario_completo &= len(sede_servicio) > 0

                if "confeccion" in tipo_proveedor:
                    formulario_completo &= (bool(nodo_1) and bool(nodo_2))


        if "otro" in tipo_proveedor:
            formulario_completo &= bool(otro_tipo_proveedor)




        # BOTÓN DE ENVÍO
        colf1, colf2, colf3 = st.columns((1,0.2,1))
        with colf2:
            enviar = st.button("Enviar", disabled=not formulario_completo, type="primary", use_container_width=True)


        if enviar:

            progress_text = "Procesando el formulario. Por favor espere..."
            progress_bar = st.progress(0, text=progress_text)

            # Leer el diccionario desde el archivo
            with open(archivo_variables, "r", encoding="utf-8") as archivo:
                variables_txt = json.load(archivo)

            # VARIABLES A ALMACENAR
            variables_txt["empleado_solicitante"] = empleado_solicitante
            variables_txt["area_solicitante"] = area_solicitante
            variables_txt["gerente_area_solicitante"] = gerente_area_solicitante
            variables_txt["empresa"] = empresa
            variables_txt["caracterizacion_sustentabilidad"] = caracterizacion_sustentabilidad
            variables_txt["empleado_caracterizador"] = empleado_caracterizador
            variables_txt["tipo_negocio"] = tipo_negocio
            variables_txt["plazo_pago"] = plazo_pago
            variables_txt["descripcion_negocio"] = descripcion_negocio
            variables_txt["servicio_prestado_en_sede_de_la_compania"] = servicio_prestado_en_sede_de_la_compania
            variables_txt["sede_servicio"] = sede_servicio
            variables_txt["tipo_proveedor"] = tipo_proveedor
            variables_txt["nodo_1"] = nodo_1
            variables_txt["nodo_2"] = nodo_2
            variables_txt["otro_tipo_proveedor"] = otro_tipo_proveedor

            # Guardar el diccionario actualizado en el archivo
            with open(archivo_variables, "w", encoding="utf-8") as archivo:
                json.dump(variables_txt, archivo, ensure_ascii=False, indent=4)
            progress_bar.progress(10, text=progress_text)

            ##########################################################################
            # LLENADO AUTOMÁTICO

            # Abrir y leer el archivo
            with open(archivo_variables, 'r', encoding='utf-8') as archivo:
                contenido = archivo.read()
            progress_bar.progress(20, text=progress_text)

            # Convertir el contenido JSON a un diccionario de Python
            datos = json.loads(contenido)

            # Iterar sobre todas las claves y valores
            for clave, valor in datos.items():
                print(f"{clave}: {valor}")



            # VARIABLES DE ENTRADA

            empleado_solicitante = datos["empleado_solicitante"]
            area_solicitante = datos["area_solicitante"]
            gerente_area_solicitante = datos["gerente_area_solicitante"]
            caracterizacion_sustentabilidad = datos["caracterizacion_sustentabilidad"]
            empleado_caracterizador = datos["empleado_caracterizador"]
            fecha_DD = datos["fecha_DD"]
            fecha_MM = datos["fecha_MM"]
            fecha_AA = datos["fecha_AA"]
            empresa = datos["empresa"]
            nombre_razon_social = datos["nombre_razon_social"]
            identificacion_tributaria = datos["identificacion_tributaria"]
            naturaleza_juridica = datos["naturaleza_juridica"]
            otro_numero_telefonico = datos["otro_numero_telefonico"]
            celular = datos["celular"]
            email_facturacion_electronica = datos["email_facturacion_electronica"]
            email_reporte_pago = datos["email_reporte_pago"]
            email_reporte_compras = datos["email_reporte_compras"]
            tipo_negocio = datos["tipo_negocio"] 
            descripcion_negocio = datos["descripcion_negocio"]
            tipo_proveedor = datos["tipo_proveedor"] 
            nodo_1 = datos["nodo_1"]
            nodo_2 = datos["nodo_2"]
            otro_tipo_proveedor = datos["otro_tipo_proveedor"]
            servicio_prestado_en_sede_de_la_compania = datos["servicio_prestado_en_sede_de_la_compania"]
            sede_servicio = datos["sede_servicio"]
            nombre_contacto_contabilidad = datos["nombre_contacto_contabilidad"]
            cargo_contacto_contabilidad = datos["cargo_contacto_contabilidad"]
            telefono_contacto_contabilidad = datos["telefono_contacto_contabilidad"]
            email_contacto_contabilidad = datos["email_contacto_contabilidad"]
            autorretenedor_impuesto_ic_municipio = datos["autorretenedor_impuesto_ic_municipio"]
            plazo_pago = datos["plazo_pago"]
            tamano_empresa = datos["tamano_empresa"]
            nombre_representante_legal = datos["nombre_representate_legal"]
            cedula_representante_legal = datos["cedula_representante_legal"]
            email_representante_legal = datos["email_representante_legal"]
            sagrilaf = datos["sagrilaft"]
            casilla_otro = []

            ###################################################################################################

            # COPIA DEL FORMULARIO Y SELECCIÓN DE HOJAS


            # NUEVA CARPETA DE SALIDA
            #carpeta_salida = r'.\salidas'
            os.makedirs(ruta_subcarpeta, exist_ok=True)

            # Definir el archivo original y la nueva copia
            archivo_original = "FORMATO CREACION PROVEEDORES NACIONALES V2.1.xlsx"
            archivo_copia = os.path.join(ruta_subcarpeta, f"{nombre_razon_social} FORMATO CREACION PROVEEDORES NACIONALES V2.1.xlsx")

            # Eliminar archivo Excel si ya existe
            if os.path.exists(archivo_copia):
                os.remove(archivo_copia)

            # Crear una copia del archivo original
            shutil.copy(archivo_original, archivo_copia)
            progress_bar.progress(30, text=progress_text)

            # Abrir la nueva copia del archivo de Excel
            app = xw.App(visible=False)  # No mostrar Excel
            wb = xw.Book(archivo_copia)  # Cargar la copia

            # Hoja "FORM. REGISTRO DE PROVEEDORES"
            hoja_proveedores = wb.sheets["FORM. REGISTRO DE PROVEEDORES"]

            # Hoja "CASILLAS"
            hoja_casillas = wb.sheets["CASILLAS"]
            progress_bar.progress(40, text=progress_text)

            ###################################################################################################

            # DILIGENCIAMIENTO DEL FORMULARIO

            hoja_proveedores.range("V5").value = empleado_solicitante
            hoja_proveedores.range("V6").value = area_solicitante
            hoja_proveedores.range("W7").value = gerente_area_solicitante

            # CARACTERIZACIÓN DE SUSTENTABILIDAD (ES EXCLUYENTE?)
            if caracterizacion_sustentabilidad == "SI":
                hoja_proveedores.range("AA9").value = "X"
                hoja_proveedores.range("X10").value = empleado_caracterizador

            if caracterizacion_sustentabilidad == "NO":
                hoja_proveedores.range("AB9").value = "X"

            hoja_proveedores.range("F12").value = fecha_DD
            hoja_proveedores.range("H12").value = fecha_MM
            hoja_proveedores.range("J12").value = fecha_AA

            # EMPRESA (ES EXCLUYENTE?)
            if "LINEA DIRECTA" in empresa:
                hoja_casillas.range("A2").value = True
            if "ELEDE CASA" in empresa:
                hoja_casillas.range("A3").value = True
            if "GRUPO ELEDE" in empresa:
                hoja_casillas.range("A4").value = True

            hoja_proveedores.range("H18").value = nombre_razon_social
            hoja_proveedores.range("H19").value = identificacion_tributaria

            # NATURALEZA JURÍDICA (ES EXCLUYENTE?)
            if naturaleza_juridica == "PERSONA NATURAL":
                hoja_casillas.range("A7").value = True
            if naturaleza_juridica == "PERSONA JURÍDICA":
                hoja_casillas.range("A8").value = True

            hoja_proveedores.range("G21").value = otro_numero_telefonico
            hoja_proveedores.range("R21").value = celular
            hoja_proveedores.range("M22").value = email_facturacion_electronica
            hoja_proveedores.range("H23").value = email_reporte_pago
            hoja_proveedores.range("V23").value = email_reporte_compras

            # BIEN O SERVICIO A PRESTAR (ES EXCLUYENTE?)
            if tipo_negocio == "VENTA DE PRODUCTOS":
                hoja_proveedores.range("M28").value = "X"

            if tipo_negocio == "PRESTACIÓN DE SERVICIOS":
                hoja_proveedores.range("AA28").value = "X"

            if tipo_negocio == 'PRODUCTOS Y SERVICIOS':
                hoja_proveedores.range("M28").value = "X"
                hoja_proveedores.range("AA28").value = "X"


            hoja_proveedores.range("I29").value = descripcion_negocio

            # TIPO DE PROVEEDOR

            if 'cuidado_hogar' in tipo_proveedor:
                hoja_casillas.range("A11").value = True

            if 'cuidado_personal' in tipo_proveedor:
                hoja_casillas.range("A12").value = True

            if 'ropa_hogar' in tipo_proveedor:
                hoja_casillas.range("A13").value = True

            if 'telas' in tipo_proveedor:
                hoja_casillas.range("A16").value = True
            if 'estampacion' in tipo_proveedor:
                hoja_casillas.range("A16").value = True
            if 'tintoreria' in tipo_proveedor:
                hoja_casillas.range("A16").value = True 

            if 'insumos' in tipo_proveedor:
                hoja_casillas.range("A17").value = True

            if 'confeccion' in tipo_proveedor:
                hoja_casillas.range("A18").value = True
                hoja_proveedores.range("F39").value = nodo_1
                hoja_proveedores.range("F40").value = nodo_2   

            if 'empaques' in tipo_proveedor:
                hoja_casillas.range("A19").value = True

            if 'muebles' in tipo_proveedor:
                hoja_casillas.range("A22").value = True

            if 'inmuebles' in tipo_proveedor:
                hoja_casillas.range("A23").value = True

            if 'alimentos' in tipo_proveedor:
                hoja_casillas.range("A26").value = True

            if 'catering' in tipo_proveedor:
                hoja_casillas.range("A27").value = True

            if 'alojamiento' in tipo_proveedor:
                hoja_casillas.range("A28").value = True

            if 'consultorias' in tipo_proveedor:
                hoja_casillas.range("A31").value = True

            if 'mantenimiento' in tipo_proveedor:
                hoja_casillas.range("A32").value = True

            if 'tic' in tipo_proveedor:
                hoja_casillas.range("A33").value = True

            if 'fotografia' in tipo_proveedor:
                hoja_casillas.range("A34").value = True

            #-------------------------------------------------------------

            if 'otro' in tipo_proveedor:
                hoja_casillas.range("A35").value = True
                casilla_otro.append(otro_tipo_proveedor)

            if 'producto_terminado' in tipo_proveedor:
                hoja_casillas.range("A35").value = True
                casilla_otro.append('PRODUCTO TERMINADO')

            if 'sublimado' in tipo_proveedor:
                hoja_casillas.range("A35").value = True
                casilla_otro.append('SUBLIMADO')

            if 'tejeduria' in tipo_proveedor:
                hoja_casillas.range("A35").value = True
                casilla_otro.append('TEJEDURÍA')

            if 'corte' in tipo_proveedor:
                hoja_casillas.range("A35").value = True
                casilla_otro.append('CORTE')

            if 'bordado' in tipo_proveedor:
                hoja_casillas.range("A35").value = True
                casilla_otro.append('BORDADO')

            if 'prehormado' in tipo_proveedor:
                hoja_casillas.range("A35").value = True
                casilla_otro.append('PREHORMADO')

            hoja_proveedores.range("N36").value = ", ".join(casilla_otro) if casilla_otro else ""


            #-------------------------------------------------------------


            # SERVICIO PRESTADO EN SEDE DE LA COMPAÑÍA (ES EXCLUYENTE?)
            if servicio_prestado_en_sede_de_la_compania == "SI":
                hoja_proveedores.range("J42").value = "X"

            if servicio_prestado_en_sede_de_la_compania == "NO":
                hoja_proveedores.range("K42").value = "X"

            # SEDE DONDE SE PRESTARÁ EL SERVICIO (ES EXCLUYENTE?)
            if "B ESCOCIA" in sede_servicio:
                hoja_casillas.range("A38").value = True
            if "CIGMO" in sede_servicio:
                hoja_casillas.range("A39").value = True
            if "HACIENDA ESCOCIA" in sede_servicio:
                hoja_casillas.range("A40").value = True

            # PERSONA DE CONTACTO EN CONTABILIDAD
            hoja_proveedores.range("F46").value = nombre_contacto_contabilidad
            hoja_proveedores.range("M46").value = cargo_contacto_contabilidad
            hoja_proveedores.range("R46").value = telefono_contacto_contabilidad
            hoja_proveedores.range("V46").value = email_contacto_contabilidad

            # INFORMACIÓN MUNICIPIO IMPUESTO DE INDUSTRIA Y COMERCIO (LA ESTRELLA - MARINILLA)
            if autorretenedor_impuesto_ic_municipio == 'LA ESTRELLA':
                hoja_casillas.range("A43").value = True
                hoja_casillas.range("A48").value = True

            if autorretenedor_impuesto_ic_municipio == 'MARINILLA':
                hoja_casillas.range("A44").value = True
                hoja_casillas.range("A47").value = True

            if autorretenedor_impuesto_ic_municipio == 'N/A':
                hoja_casillas.range("A44").value = True
                hoja_casillas.range("A48").value = True


            # INFORMACIÓN FINANCIERA
            hoja_proveedores.range("F57").value = plazo_pago

            # TAMAÑO EMPRESARIAL (ES EXCLUYENTE?)
            if tamano_empresa == "MICRO":
                hoja_casillas.range("A51").value = True

            if tamano_empresa == "PEQUEÑA":
                hoja_casillas.range("A52").value = True

            if tamano_empresa == "MEDIANA":
                hoja_casillas.range("A53").value = True

            if tamano_empresa == "GRANDE":
                hoja_casillas.range("A54").value = True

            # REPRESENTANTE LEGAL
            hoja_proveedores.range("I63").value = nombre_representante_legal
            hoja_proveedores.range("D64").value = cedula_representante_legal
            hoja_proveedores.range("Q64").value = email_representante_legal

            # SAGRILAF (ES EXCLUYENTE?)
            if sagrilaf == "SI":
                hoja_casillas.range("A57").value = True

            if sagrilaf == "NO":
                hoja_casillas.range("A58").value = True

            progress_bar.progress(50, text=progress_text)
            ###################################################################################################    

            # GUARDAR Y CERRAR EL ARCHIVO DE EXCEL

            wb.save()
            print("Celdas actualizadas correctamente sin perder los controles de formulario.")
            progress_bar.progress(80, text=progress_text)
            try:
                # Definir el nombre del archivo PDF con ruta absoluta
                archivo_pdf = os.path.abspath(os.path.join(ruta_subcarpeta, f"{nombre_razon_social} - FORMATO INSCRIPCIÓN PROVEEDORES NACIONALES.pdf"))

                # Eliminar archivo PDF si ya existe
                if os.path.exists(archivo_pdf):
                    os.remove(archivo_pdf)
                
                # Ajustar configuración de la página para reducir los márgenes
                hoja_proveedores.api.PageSetup.TopMargin = 30  # Margen superior (en puntos, 1 punto ≈ 0.035 cm)
                hoja_proveedores.api.PageSetup.BottomMargin = 30  # Margen inferior
                hoja_proveedores.api.PageSetup.LeftMargin = 30  # Margen izquierdo
                hoja_proveedores.api.PageSetup.RightMargin = 30  # Margen derecho
                hoja_proveedores.api.PageSetup.Zoom = False  # Desactiva zoom automático
                hoja_proveedores.api.PageSetup.FitToPagesWide = 1  # Ajustar al ancho de 1 página
                hoja_proveedores.api.PageSetup.FitToPagesTall = 2  # Ajustar al alto de 1 página
                
                # Eliminar encabezado y pie de página
                hoja_proveedores.api.PageSetup.CenterFooter = ""  # Vaciar el pie de página
                hoja_proveedores.api.PageSetup.LeftFooter = ""
                hoja_proveedores.api.PageSetup.RightFooter = ""
                hoja_proveedores.api.PageSetup.CenterHeader = ""  # Vaciar el encabezado
                hoja_proveedores.api.PageSetup.LeftHeader = ""
                hoja_proveedores.api.PageSetup.RightHeader = ""
                #progress_bar.progress(70, text=progress_text)
                
                #hoja_proveedores.api.PageSetup.FitToPagesWide = 1  # Ajusta al ancho de una página
                #hoja_proveedores.api.PageSetup.FitToPagesTall = False  # Permite varias páginas de alto



                # Exportar la hoja a PDF
                hoja_proveedores.api.ExportAsFixedFormat(0, archivo_pdf)
                #progress_bar.progress(80, text=progress_text)

                print(f"PDF guardado en: {archivo_pdf}")
            except Exception as e:
                print(f"Error al exportar el PDF: {e}")


            wb.close()
            progress_bar.progress(90, text=progress_text)
            app.quit()



            ##########################################################################
            # ENVÍO DE CORREO

            #destinatario = 'veronica.arango@lineadirecta.com.co'

            #outlook = win32.Dispatch('outlook.application')
            #mail = outlook.CreateItem(0)
            #mail.To = destinatario
            #mail.Subject = "Ensayo de envío: Archivos adjuntos desde Python"
            #mail.Body = "Adjunto los archivos."

            #for archivo in os.listdir(ruta_subcarpeta):
            #    ruta = os.path.abspath(os.path.join(ruta_subcarpeta, archivo))
            #    if os.path.isfile(ruta):
            #        try:
            #            mail.Attachments.Add(ruta)
            #        except Exception as e:
            #            print(f"No se pudo adjuntar {ruta}: {e}")
            #    else:
            #        print(f"Archivo no encontrado o no es archivo: {ruta}")

            #try:
            #    mail.Send()
            #    print("Correo enviado desde Outlook local.")
            #except Exception as e:
            #    print(f"Error al enviar el correo: {e}")





            # Parámetros de configuración
            #carpeta = r'C:\ruta\a\tu\carpeta'  # Reemplaza con la ruta real
            destinatario = 'veronica.arango@lineadirecta.com.co'  # Reemplaza con la dirección de destino
            remitente = 'veronica.arango@lineadirecta.com.co'     # Tu correo de Outlook
            contrasena = '*******'               # Usa contraseña de aplicación si Outlook lo requiere
            asunto = f'Documentación - Proveedor Nacional {nombre_razon_social}'
            cuerpo = f'Cordial saludo.\n\nAdjunto a este correo encontrará los documentos y el formulario de inscripción necesarios para el proceso de registro como proveedor nacional de {nombre_razon_social}. \n\nAgradecemos su atención y quedamos atentos a cualquier inquietud.'

            # Crear mensaje
            msg = EmailMessage()
            msg['Subject'] = asunto
            msg['From'] = remitente
            msg['To'] = destinatario
            msg.set_content(cuerpo)

            # Adjuntar archivos de la carpeta
            for archivo in os.listdir(ruta_subcarpeta):
                ruta_archivo = os.path.join(ruta_subcarpeta, archivo)
                if os.path.isfile(ruta_archivo):
                    with open(ruta_archivo, 'rb') as f:
                        contenido = f.read()
                        msg.add_attachment(contenido, maintype='application', subtype='octet-stream', filename=archivo)

            # Enviar correo usando SMTP de Outlook
            try:
                with smtplib.SMTP('smtp.office365.com', 587) as smtp:
                    smtp.starttls()
                    smtp.login(remitente, contrasena)
                    smtp.send_message(msg)
                    print("Correo enviado exitosamente.")
            except Exception as e:
                print(f"Error al enviar el correo: {e}")






            progress_bar.progress(100, text=progress_text)

            progress_bar.empty()


            st.success("Formulario enviado con éxito.")




