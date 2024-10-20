import imaplib
import email
from email.policy import default
import os
import PyPDF2
import pandas as pd
import re

# Conexión al servidor IMAP
def connectToEmail():
    mail = imaplib.IMAP4_SSL('imap.gmail.com')
    mail.login('su correo', 'su clave')
    mail.select('inbox')
    return mail


# Buscar correos del remitente especifico y con PDF adjuntos de dias especificos
def searchEmails(mail):
    status, messages = mail.search(None, 'FROM "correo remitente"')
    if status != 'OK':
        print("Error al buscar correos.")
        return []

    email_ids = messages[0].split()
    print(f"Número de correos encontrados: {len(email_ids)}")  
    
    filtered_email_ids = []
    
    for email_id in email_ids:
        status, msg_data = mail.fetch(email_id, '(RFC822)')
        if status != 'OK':
            print(f"Error al obtener el correo {email_id}.")
            continue
        
        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1], policy=default)
                subject = msg['subject']
                print(f"Asunto encontrado: {subject}")  
                
                # Filtra los correos que contienen un asunto especifico
                #usa expresiones regulares ejemplo r'Boletín de Precios de Mayorista a Minorista del CENADA'
                if re.search(r'', subject, re.IGNORECASE):
                    filtered_email_ids.append(email_id)
    
    return filtered_email_ids

# Descargar y guardar adjuntos PDF
def downloadAttachments(mail, email_id):
    status, msg_data = mail.fetch(email_id, '(RFC822)')
    if status != 'OK':
        print(f"Error al obtener el correo {email_id}.")
        return None
    
    for response_part in msg_data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1], policy=default)
            for part in msg.iter_attachments():
                if part.get_content_type() == 'application/pdf':
                    filename = part.get_filename()
                    #expresion regular a su gusto
                    #ejemplo \bLUNES\b|\bMIERCOLES\b|\bVIERNES\b
                    if re.search(r'aqui su filtro', filename, re.IGNORECASE):
                        with open(filename, 'wb') as f:
                            f.write(part.get_payload(decode=True))
                        return filename
    return None

# Extraer texto del PDF
def extractPdfText(filename):
    with open(filename, 'rb') as f:
        reader = PyPDF2.PdfReader(f)
        text = ''
        for page in reader.pages:
            text += page.extract_text()
    return text

# Procesar datos del PDF y guardarlos en un archivo CSV
def processData(text):
    data = []
    for line in text.splitlines():
        if "Producto" in line:
            continue  # Saltar encabezados
        if line.strip():
            parts = line.split()
            if len(parts) >= 4:
                producto, promedio, maximo, minimo = parts[0], parts[1], parts[2], parts[3]
                data.append([producto, promedio, maximo, minimo])
    
    df = pd.DataFrame(data, columns=['Producto', 'Promedio', 'Máximo', 'Mínimo'])
    df.to_csv('historico_precios.csv', index=False)

# Main
def main():
    mail = connectToEmail()
    email_ids = searchEmails(mail)
    
    for email_id in email_ids:
        filename = downloadAttachments(mail, email_id)
        if filename and os.path.exists(filename):  # Verifica que el archivo exista
            text = extractPdfText(filename)
            processData(text)
            os.remove(filename)  # Eliminar el archivo PDF después de procesarlo.

    mail.logout()  # Cerrar la sesion del mail

if __name__ == '__main__':
    main()