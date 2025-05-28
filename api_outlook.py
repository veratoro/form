
import os
import win32com.client as win32

carpeta = r'D:\Users\varango\OneDrive - Linea Directa S.A.S\Escritorio\FORMULARIO\fill_proveedores\proveedores_inscritos\20250410_114316_VAT_S.A.S._12345'
destinatario = 'veronica.arango@lineadirecta.com.co'

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = destinatario
mail.Subject = "Ensayo de envío: Archivos adjuntos desde Python"
mail.Body = "Adjunto los archivos."

for archivo in os.listdir(carpeta):
    ruta = os.path.join(carpeta, archivo)
    if os.path.isfile(ruta):
        mail.Attachments.Add(ruta)

mail.Send()
print("Correo enviado desde Outlook local.")

