import streamlit as st
import pandas as pd
import os
import uuid
import shutil
from datetime import datetime
import json

def formularioExternos():


    # Función para guardar archivos subidos
    #def save_uploaded_file(uploaded_file, directory):
    #    if uploaded_file is not None:
    #        # Crear el directorio si no existe
    #        if not os.path.exists(directory):
    #            os.makedirs(directory)
            
            # Generar un nombre único para el archivo
    #        unique_id = uuid.uuid4().hex  # Genera un identificador único
    #        file_name, file_extension = os.path.splitext(uploaded_file.name)
    #        unique_file_name = f"{file_name}_{unique_id}{file_extension}"
            
            # Guardar el archivo en el directorio
    #        file_path = os.path.join(directory, unique_file_name)
    #        with open(file_path, "wb") as f:
    #            f.write(uploaded_file.getbuffer())
    #        return file_path
    #    return None

    def save_uploaded_file(uploaded_file, directory, fixed_name_base):
        if uploaded_file is not None:
            if not os.path.exists(directory):
                os.makedirs(directory)
            # Obtener la extensión original del archivo subido
            _, file_extension = os.path.splitext(uploaded_file.name)
            # Construir el nombre fijo con la extensión original
            fixed_name = fixed_name_base + file_extension
            file_path = os.path.join(directory, fixed_name)
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            return file_path
        return None




    st.markdown("""---""")
    #st.write('')

    #----------------------------
    fecha_DD = None
    fecha_MM = None
    fecha_AA = None
    naturaleza_juridica = None
    nombre_razon_social = None
    identificacion_tributaria = None
    email_representante_legal = None
    otro_numero_telefonico = None
    celular = None
    nombre_representate_legal = None
    cedula_representante_legal = None
    tamano_empresa = None
    autorretenedor_impuesto_ic_municipio = None
    sagrilaft = None
    serv_confeccion = None
    nombre_contacto_contabilidad = None
    cargo_contacto_contabilidad = None
    telefono_contacto_contabilidad = None
    email_contacto_contabilidad = None
    email_reporte_pago = None
    email_facturacion_electronica = None
    email_reporte_compras = None
    #--------------------------------

    rut_j = None
    existencia_j = None
    tamano_j = None
    banco_j = None
    acciones_j = None

    rut_n = None
    banco_n = None
    existencia_n = None
    planilla_n = None
    documento_n = None

    #--------------------------------

    col8, col9, col10 = st.columns(3)
    with col8:
        fecha = st.date_input('FECHA DE DILIGENCIAMIENTO *', value= pd.Timestamp('today'))
        
        if fecha:
            fecha_DD = f"{fecha.day:02d}"
            fecha_MM = f"{fecha.month:02d}"
            fecha_AA = f"{fecha.year:04d}"
        else:
            fecha_DD = fecha_MM = fecha_AA = None      



    col11, col12, col13 = st.columns((0.6,0.5,0.4))

    with col11:
       naturaleza_juridica = st.segmented_control('NATURALEZA JURÍDICA *', ['PERSONA NATURAL', 'PERSONA JURÍDICA'])

    with col12:
        if naturaleza_juridica == 'PERSONA NATURAL':
            nombre_razon_social = st.text_input('NOMBRE *', autocomplete="off").upper()
        if naturaleza_juridica == 'PERSONA JURÍDICA':
            nombre_razon_social = st.text_input('RAZÓN SOCIAL *', autocomplete="off").upper()

    with col13:
        if naturaleza_juridica == 'PERSONA NATURAL':        
            identificacion_tributaria = st.text_input('IDENTIFICACIÓN *', autocomplete="off")
        if naturaleza_juridica == 'PERSONA JURÍDICA':
            identificacion_tributaria = st.text_input('IDENTIFICACIÓN TRIBUTARIA *', autocomplete="off").upper()



    if naturaleza_juridica == 'PERSONA NATURAL':

        col14, col15, col16 = st.columns((0.8,0.5,0.5))
        with col14:
            email_representante_legal = st.text_input('EMAIL *', autocomplete="off").upper()            
        with col15:
            otro_numero_telefonico = st.text_input('NÚMERO TELEFÓNICO', autocomplete="off")
        with col16:
            celular = st.text_input('CELULAR', autocomplete="off")



    if naturaleza_juridica == 'PERSONA JURÍDICA':

        col14, col15, col16 = st.columns((0.8,0.5,0.5))
        with col14:
            nombre_representate_legal = st.text_input('REPRESENTANTE LEGAL *', autocomplete="off").upper()
        with col15:
            cedula_representante_legal = st.text_input('IDENTIFICACIÓN *', autocomplete="off")

        col14, col15, col16 = st.columns((0.8,0.5,0.5))
        with col14:
            email_representante_legal = st.text_input('EMAIL *', autocomplete="off").upper()            
        with col15:
            otro_numero_telefonico = st.text_input('NÚMERO TELEFÓNICO', autocomplete="off")
        with col16:
            celular = st.text_input('CELULAR', autocomplete="off")



    col17, col18 = st.columns((0.5,0.5))
    with col17:
        tamano_empresa = st.segmented_control('TAMAÑO EMPRESARIAL *', ['MICRO', 'PEQUEÑA', 'MEDIANA', 'GRANDE', 'N/A'])
    with col18:
        autorretenedor_impuesto_ic_municipio = st.segmented_control('MUNICIPIO EN EL QUE ES AUTORRETENEDOR DEL IMPUESTO DE INDUSTRIA Y COMERCIO *', ['LA ESTRELLA', 'MARINILLA', 'N/A'], key='estrella_marinilla')
    

    col19, col20 = st.columns(2)
    with col19:
        sagrilaft = st.segmented_control('¿CUENTA CON LA OBLIGACIÓN DE IMPLEMENTAR SAGRILAFT? *', ['SI', 'NO'], key='sagrilaft')
    with col20:
        if naturaleza_juridica == 'PERSONA JURÍDICA':
            serv_confeccion = st.segmented_control('¿OFRECE SERVICIOS DE CONFECCIÓN? *', ['SI', 'NO'], key='serv_confeccion')



    st.write('')



    if naturaleza_juridica == 'PERSONA JURÍDICA':
        col1, col2, col3, col4 = st.columns((0.5,0.3,0.3,0.5))
        with col1:
            nombre_contacto_contabilidad = st.text_input('PRESONA DE CONTACTO CONTABILIDAD', autocomplete="off").upper()
        with col2:
            cargo_contacto_contabilidad = st.text_input('CARGO', autocomplete="off").upper()
        with col3:
            telefono_contacto_contabilidad = st.text_input('TELÉFONO', autocomplete="off")
        with col4:
            email_contacto_contabilidad = st.text_input('EMAIL', key='email_contabilidad', autocomplete="off").upper()



    col19, col20 = st.columns(2)
    with col19:
        email_reporte_pago = st.text_input('EMAIL PARA REPORTE DE PAGO *', placeholder='CERTIFICACIONES TRIBUTARIAS DE RETENCIÓN EN LA FUENTE, RENTA E IVA', autocomplete="off").upper()



    col17, col18 = st.columns(2)
    with col17:
        email_facturacion_electronica = st.text_input('EMAIL DESDE EL CUAL ENVÍA LA FACTURACIÓN ELECTRÓNICA', autocomplete="off").upper()
    with col18:
        email_reporte_compras = st.text_input('EMAIL PARA REPORTE DE COMPRAS', autocomplete="off").upper()



    st.write('')
    st.write('')



    if (naturaleza_juridica == 'PERSONA JURÍDICA') and (serv_confeccion == 'NO'):

        col1, col2, col3, col4 = st.columns((0.1,1,1,0.1))
        
        with col2:
            rut_j = st.file_uploader('RUT *', help='Documento completo con fecha de generación no mayor a 60 días (la fecha de generacion aparece en extremo inferior derecho de la hoja 1)')
            st.write('')
            existencia_j = st.file_uploader('CERTIFICADO EXISTENCIA Y REPRESENTACIÓN LEGAL *', help='Cámara de Comercio. Con fecha de generación no mayor a 60 días (debe tener fecha de renovación de matrícula mercantil vigente)')
            
            
            
        with col3:
            tamano_j = st.file_uploader('CERTIFICADO TAMAÑO EMPRESARIAL *', help='Si el plazo de pago es superior a 45 días y la empresa esta clasificada como Micro, Pequeña o Mediana  certificación expedida por representante legal informando la clasificación del tamaño empresarial conforme a sus ingresos por actividades ordinarias con base en el decreto 957 del 05 de junio de 2019.')            
            st.write('')
            banco_j = st.file_uploader('CERTIFICADO BANCARIO *', help='Con fecha de generación no mayor a 60 días.') 



    if (naturaleza_juridica == 'PERSONA JURÍDICA') and (serv_confeccion == 'SI'):


        col1, col2, col3, col4 = st.columns((0.1,1,1,0.1))
        
        with col2:
            rut_j = st.file_uploader('RUT *', help='Documento completo con fecha de generación no mayor a 60 días (la fecha de generacion aparece en extremo inferior derecho de la hoja 1)', key = 'rut_j_confeccion')
            st.write('')
            existencia_j = st.file_uploader('CERTIFICADO EXISTENCIA Y REPRESENTACIÓN LEGAL *', help='Cámara de Comercio. Con fecha de generación no mayor a 60 días (debe tener fecha de renovación de matrícula mercantil vigente)', key = 'existencia_j_confeccion')
            st.write('')
            tamano_j = st.file_uploader('CERTIFICADO TAMAÑO EMPRESARIAL *', help='Si el plazo de pago es superior a 45 días y la empresa esta clasificada como Micro, Pequeña o Mediana  certificación expedida por representante legal informando la clasificación del tamaño empresarial conforme a sus ingresos por actividades ordinarias con base en el decreto 957 del 05 de junio de 2019.', key = 'tamano_j_confeccion')            
            
            
        with col3:

            banco_j = st.file_uploader('CERTIFICADO BANCARIO *', help='Con fecha de generación no mayor a 60 días.')
            st.write('')
            acciones_j = st.file_uploader('CERTIFICADO COMPOSICIÓN ACCIONARIA *', help='Ver modelo sugerido.')
            
            with open(r'.\fill_proveedores\MODELO_CERTIFICADO_COMPOSICION_ACCIONARIA.docx', 'rb') as file:
                st.download_button(
                    label='MODELO CERTIFICADO COMPOSICIÓN ACCIONARIA',
                    data=file,
                    file_name='MODELO_CERTIFICADO_COMPOSICION_ACCIONARIA.docx',
                    mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    use_container_width=True,
                    key='download_button',
                    type="primary"
                )
        



    if naturaleza_juridica == 'PERSONA NATURAL':
        col1, col2, col3, col4 = st.columns((0.1,1,1,0.1))
        with col2:
            rut_n = st.file_uploader('RUT *', help='Documento completo con fecha de generación no mayor a 60 días (la fecha de generacion aparece en extremo inferior derecho de la hoja 1)')
            banco_n = st.file_uploader('CERTIFICADO BANCARIO *', help='Con fecha de generación no mayor a 60 días.')
            existencia_n = st.file_uploader('CERTIFICADO EXISTENCIA Y REPRESENTACIÓN LEGAL *', help='Cámara de Comercio. Con fecha de generación no mayor a 60 días (si aplica)')
        with col3:
            planilla_n = st.file_uploader('PLANILLA INTEGRAL DE SEGURIDAD SOCIAL *', help='Constancia o planilla de afiliación a seguridad social como independiente (aplica para prestación de servicios personales)')
            documento_n = st.file_uploader('DOCUMENTO DE IDENTIDAD *', help='') 



    st.write('')
    st.write('')
    st.write('')



    col21, col22 = st.columns((0.04,1))
    with col21:
        st.checkbox('', key='A')
    with col22:
        st.info('**DECLARO** que la información suministrada en el presente formulario es real y verificable y asumo plena responsabilidad por la veracidad de la misma, por lo tanto autorizo  para consignar en la cuenta bancaria definida anteriormente, los valores correspondientes a las facturas por los diferentes contratos u órdenes  de  compra que lleguemos a celebrar. En caso de cualquier modificación será notificada oportunamente y exonero de responsabilidad por error e inexactitud en el suministro de la información.')

    col21, col22 = st.columns((0.04,1))
    with col21:
        st.checkbox('', key='B')
    with col22:
        st.info('**ACEPTO** que la relación jurídica que nace entre las partes, podrá ser regulada a través de Órdenes de Compra la cual será prueba suficiente de la existencia de la relación contractual, no obstante, las partes podrán celebrar un contrato escrito, caso en el cual lo establecido en el contrato, primará sobre otro documento.')

    col21, col22 = st.columns((0.04,1))
    with col21:
        st.checkbox('', key='C')
    with col22:
        st.info('**AUTORIZO** que los datos personales suministrados sean tratados de manera segura para las finalidades y de acuerdo con las condiciones definidas en la Política de Tratamiento de Datos Personales, alojada en www.lineadirecta.com.co, bajo cumplimiento de la Ley 1581 de 2012, su Decreto Reglamentario 1377 de 2013 y demás normatividad relacionada con Hábeas Data. Para ejercer los derechos que le asisten podrá comunicarse al correo datospersonales@lineadirecta.com.co. De igual forma toda la información conocida por las partes será confidencial, las partes protegerán esa información y serán responsables ante la otra por su publicación no autorizada.')

    col21, col22 = st.columns((0.04,1))
    with col21:
        st.checkbox('', key='D')
    with col22:
        st.info('**DECLARO** que los recursos comprometidos para el desarrollo del objeto social provienen de actividades lícitas y **ACEPTO** que en caso de aparecer en las listas, nacionales e internacionales sobre antecedentes de lavado de activos y financiación del terrorismo, el contratante tendrá la facultad de dar por terminada cualquier relación contractual que las partes lleguen a suscribir, sin que esto otorgue la facultad de invocar indemnización del perjuicios.  De acuerdo con esto, La validación de SAGRILAFT es previa a la realización comercial con la compañía. En adición, **ACEPTO**, cumplir con todos los lineamientos estabecido en el Código de Ética alojado en www.lineadirecta.com.co')

    col21, col22 = st.columns((0.04,1))
    with col21:
        st.checkbox('', key='E')
    with col22:
        st.info('**DECLARO** que no existe ningún conflicto de interés con la Compañía de acuerdo con su código de ética y que en caso de tal, diligenciaré el formato de conflicto de interés.')



    st.write('')
    st.write('')



    #######################################

    # VALIDACIÓN

    formulario_completo = all([
        fecha is not None,
        naturaleza_juridica is not None,   
        tamano_empresa is not None,
        autorretenedor_impuesto_ic_municipio is not None,
        sagrilaft is not None,
        email_reporte_pago,
    ])

    declaraciones_completas = all([
        st.session_state.get('A'),
        st.session_state.get('B'),
        st.session_state.get('C'),
        st.session_state.get('D'),
        st.session_state.get('E')
    ])


    # VALIDACIONES CONDICIONALES

    if (naturaleza_juridica == 'PERSONA NATURAL') or (naturaleza_juridica == 'PERSONA JURÍDICA'):
        formulario_completo &= bool(nombre_razon_social)
        formulario_completo &= bool(identificacion_tributaria)
    
    if naturaleza_juridica == 'PERSONA NATURAL':
        formulario_completo &= bool(email_representante_legal)
        formulario_completo &= bool(rut_n) 
        formulario_completo &= bool(banco_n)
        formulario_completo &= bool(existencia_n)
        formulario_completo &= bool(planilla_n)
        formulario_completo &= bool(documento_n)    

    if naturaleza_juridica == 'PERSONA JURÍDICA':
        formulario_completo &= bool(nombre_representate_legal)
        formulario_completo &= bool(cedula_representante_legal)
        formulario_completo &= bool(email_representante_legal)
        formulario_completo &= bool(rut_j)
        formulario_completo &= bool(existencia_j)
        formulario_completo &= bool(tamano_j)
        formulario_completo &= bool(banco_j)

        if serv_confeccion == 'SI':
            formulario_completo &= bool(acciones_j)

    # Integrar ambas validaciones
    formulario_y_declaraciones_completas = (formulario_completo and declaraciones_completas)

    # Mostrar botón de envío (deshabilitado hasta que ambas validaciones sean True)
    # BOTÓN DE ENVÍO
    colf1, colf2, colf3 = st.columns((1,0.2,1))
    with colf2:
        enviar = st.button("Enviar", disabled=not formulario_y_declaraciones_completas, type="primary")

    col21, col22 = st.columns((0.04,1))

    with col22:

        if enviar:

            progress_text = "Procesando el formulario. Por favor espere..."
            progress_bar = st.progress(0, text=progress_text)

            nombre_razon_social_sin_espacios = nombre_razon_social.replace(' ', '_')
            now = datetime.now()
            fecha_hora = now.strftime("%Y%m%d_%H%M%S")
            folder_name = f"{fecha_hora}_{nombre_razon_social_sin_espacios}_{identificacion_tributaria}"

            # Ruta base para guardar los archivos
            base_path = "./fill_proveedores/proveedores_inscritos"
            directory_path = os.path.join(base_path, folder_name)
            progress_bar.progress(90, text=progress_text)


            # Guardar documentos para PERSONA NATURAL
            if naturaleza_juridica == 'PERSONA NATURAL':
                if rut_n:
                    save_uploaded_file(rut_n, directory_path, "RUT")
                if banco_n:
                    save_uploaded_file(banco_n, directory_path, "CERTIFICADO_BANCARIO")
                if existencia_n:
                    save_uploaded_file(existencia_n, directory_path, "EXISTENCIA_REPRESENTACION_LEGAL")
                if planilla_n:
                    save_uploaded_file(planilla_n, directory_path, "PLANILLA_SEGURIDAD_SOCIAL")
                if documento_n:
                    save_uploaded_file(documento_n, directory_path, "DOCUMENTO_IDENTIDAD")


            # Guardar documentos para persona JURÍDICA (NO Confección)
            if (naturaleza_juridica == 'PERSONA JURÍDICA') and (serv_confeccion == 'NO'):
                if rut_j:
                    save_uploaded_file(rut_j, directory_path, "RUT")
                if existencia_j:
                    save_uploaded_file(existencia_j, directory_path, "EXISTENCIA_REPRESENTACION_LEGAL")
                if tamano_j:
                    save_uploaded_file(tamano_j, directory_path, "TAMANO_EMPRESARIAL")
                if banco_j:
                    save_uploaded_file(banco_j, directory_path, "CERTIFICADO_BANCARIO")


            # Guardar documentos para persona JURÍDICA (SI Confección)
            if (naturaleza_juridica == 'PERSONA JURÍDICA') and (serv_confeccion == 'SI'):
                if rut_j:
                    save_uploaded_file(rut_j, directory_path, "RUT")
                if existencia_j:
                    save_uploaded_file(existencia_j, directory_path, "EXISTENCIA_REPRESENTACION_LEGAL")
                if tamano_j:
                    save_uploaded_file(tamano_j, directory_path, "TAMANO_EMPRESARIAL")
                if banco_j:
                    save_uploaded_file(banco_j, directory_path, "CERTIFICADO_BANCARIO")
                if acciones_j:
                    save_uploaded_file(acciones_j, directory_path, "COMPOSICION_ACCIONARIA")



            # Definir la ruta completa del archivo
            ruta_archivo = os.path.join(directory_path, "variables_proveedor.txt")
            progress_bar.progress(98, text=progress_text)

            # VARIABLES A ALMACENAR
            datos_proveedor = {
            'fecha_DD': fecha_DD,
            'fecha_MM': fecha_MM,
            'fecha_AA': fecha_AA,
            'naturaleza_juridica': naturaleza_juridica,
            'nombre_razon_social': nombre_razon_social,
            'identificacion_tributaria': identificacion_tributaria,
            'email_representante_legal': email_representante_legal,
            'otro_numero_telefonico': otro_numero_telefonico,
            'celular': celular,
            'nombre_representate_legal': nombre_representate_legal,
            'cedula_representante_legal': cedula_representante_legal,
            'tamano_empresa': tamano_empresa,
            'autorretenedor_impuesto_ic_municipio': autorretenedor_impuesto_ic_municipio,
            'sagrilaft': sagrilaft,
            'serv_confeccion': serv_confeccion,
            'nombre_contacto_contabilidad': nombre_contacto_contabilidad,
            'cargo_contacto_contabilidad': cargo_contacto_contabilidad,
            'telefono_contacto_contabilidad': telefono_contacto_contabilidad,
            'email_contacto_contabilidad': email_contacto_contabilidad,
            'email_reporte_pago': email_reporte_pago,
            'email_facturacion_electronica': email_facturacion_electronica,
            'email_reporte_compras': email_reporte_compras
            }
            progress_bar.progress(99, text=progress_text)

            # Guardar en formato JSON
            with open(ruta_archivo, "w", encoding="utf-8") as f:
                json.dump(datos_proveedor, f, ensure_ascii=False, indent=4)
            progress_bar.progress(100, text=progress_text)
            progress_bar.empty()


            st.success("Formulario enviado con éxito.")