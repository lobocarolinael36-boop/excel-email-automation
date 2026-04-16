import openpyxl
import smtplib
from email.mime.text import MIMEText

print("Iniciando automatización...")
# Cargar Excel
archivo = openpyxl.load_workbook("ventas.xlsx")
hoja = archivo.active

# Configuración de email
email_remitente = "lobokay5@gmail.com"
password = "vwsl rlsx phez cfgo"

# Recorrer filas (desde la 2 porque la 1 es encabezado)
for fila in hoja.iter_rows(min_row=2, values_only=True):
    nombre = fila[0]
    email = fila[1]
    venta = fila[2]

    if venta > 3000 and email:
        mensaje = f"Hola {nombre}, felicitaciones por tu venta de {venta}!"

        msg = MIMEText(mensaje)
        msg["Subject"] = "Buen trabajo!"
        msg["From"] = email_remitente
        msg["To"] = email

        servidor = smtplib.SMTP("smtp.gmail.com", 587)
        servidor.starttls()
        servidor.login(email_remitente, password)
        servidor.send_message(msg)
        servidor.quit()

        print(f"Email enviado a {nombre}")
print("Proceso terminado.")