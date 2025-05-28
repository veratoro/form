import xlwings as xw
import shutil
import os

###################################################################################################

# VARIABLES DE ENTRADA

empleado_solicitante = "Veronica Arango Toro"
area_solicitante = "Planeación Financiera"
gerente_area_solicitante = "Jesús Hernando Cubides Ballesteros"
caracterizacion_sustentabilidad = "SI"
empleado_caracterizador = "Lizette Botero Ospina"
fecha_DD = '07'
fecha_MM = '03'
fecha_AA = '2025'
empresa = ["LINEA DIRECTA", "ELEDE CASA"]
nombre_razon_social = "VAT SAS"
identificacion_tributaria = "123456789"
naturaleza_juridica = "PERSONA JURÍDICA"
otro_numero_telefonico = "3147315520"
celular = "3147315520"
email_facturacion_electronica = "veronica.arango@lineadirecta.com.co"
email_reporte_pago = "veroarango24@gmail.com"
email_reporte_compras = "veroarango24@gmail.com"
tipo_negocio = "VENTA DE PRODUCTOS"
descripcion_negocio = "Servicio de consultoría infraestructura bigdata"
tipo_proveedor = ["cuidado_hogar", "servicio_confeccion", "muebles", "otro"]
nodo_1 = "DESCRIPCIÓN DEL NODO 1"
nodo_2 = "DESCRIPCIÓN DEL NODO 2"
otro_tipo_proveedor = "Consultoría TI"
servicio_prestado_en_sede_de_la_compania = "SI"
sede_servicio = ["B ESCOCIA", "CIGMO"]
nombre_contacto_contabilidad = "Laura Katherine Acevedo Gómez"
cargo_contacto_contabilidad = "Analista de datos"
telefono_contacto_contabilidad = "987654321"
email_contacto_contabilidad = "laura.acevedo@lineadirecta.com.co"
autorretenedor_impuesto_ic_municipio = 'LA ESTRELLA'
plazo_pago = "15"
tamano_empresa = "MEDIANA"
nombre_representante_legal = "Sebastián Manco Martinez"
cedula_representante_legal = "1023456987"
email_representante_legal = "sebastian.manco@lineadirecta.com.co"
sagrilaf = "SI"

###################################################################################################

# COPIA DEL FORMULARIO Y SELECCIÓN DE HOJAS


# Definir el archivo original y la nueva copia
archivo_original = "FORMATO CREACION PROVEEDORES NACIONALES V2.1.xlsx"
archivo_copia = f"{nombre_razon_social} FORMATO CREACION PROVEEDORES NACIONALES V2.1.xlsx"

# Crear una copia del archivo original
shutil.copy(archivo_original, archivo_copia)

# Abrir la nueva copia del archivo de Excel
app = xw.App(visible=False)  # No mostrar Excel
wb = xw.Book(archivo_copia)  # Cargar la copia

# Hoja "FORM. REGISTRO DE PROVEEDORES"
hoja_proveedores = wb.sheets["FORM. REGISTRO DE PROVEEDORES"]

# Hoja "CASILLAS"
hoja_casillas = wb.sheets["CASILLAS"]

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
if 'insumos' in tipo_proveedor:
    hoja_casillas.range("A17").value = True
if 'servicio_confeccion' in tipo_proveedor:
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
if 'otro' in tipo_proveedor:
    hoja_casillas.range("A35").value = True
    hoja_proveedores.range("N36").value = otro_tipo_proveedor

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

if autorretenedor_impuesto_ic_municipio == 'NO APLICA':
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


###################################################################################################    

# GUARDAR Y CERRAR EL ARCHIVO DE EXCEL

wb.save()
print("Celdas actualizadas correctamente sin perder los controles de formulario.")

try:
    # Definir el nombre del archivo PDF con ruta absoluta
    archivo_pdf = os.path.abspath(f"{nombre_razon_social} - FORMATO INSCRIPCIÓN PROVEEDORES NACIONALES.pdf")
    
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

    
    #hoja_proveedores.api.PageSetup.FitToPagesWide = 1  # Ajusta al ancho de una página
    #hoja_proveedores.api.PageSetup.FitToPagesTall = False  # Permite varias páginas de alto



    # Exportar la hoja a PDF
    hoja_proveedores.api.ExportAsFixedFormat(0, archivo_pdf)
    

    print(f"PDF guardado en: {archivo_pdf}")
except Exception as e:
    print(f"Error al exportar el PDF: {e}")


wb.close()
app.quit()







