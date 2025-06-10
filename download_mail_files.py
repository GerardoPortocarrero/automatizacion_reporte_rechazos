import pandas as pd
import os

# Eliminar columna y fila vacia y asignar el encabezado
def fix_misplaced_headers(df):
    # Eliminar columnas unnamed (vacias)
    df = df.dropna(axis=1, how='all')

    # Tomar la fila 1 como nombres de columnas
    df.columns = df.iloc[1]
    
    # Eliminar las dos primeras filas (la original de encabezado y la fila de nombres)
    df = df.iloc[2:].reset_index(drop=True)
    
    return df

# Descargar archivos del correo
def download_mail_files(mails, MAIL_ITEM_CODE, root_address, test_address, transportista, ruta):
    files_found = {} # {
                     #  9/06/2025: {'transportista': {}, 'ruta': {}},
                     #  10/06/2025: {'transportista': {}, 'ruta': {}},                     
                     # }

    for mail in mails:
        if mail.Class != MAIL_ITEM_CODE or mail.Attachments.Count == 0:
            continue

        subject_lower = mail.Subject.lower()
        string_date_mail = mail.ReceivedTime.strftime("%Y-%m-%d")

        for attachment in mail.Attachments:
            # Si no es un excel saltar
            if not attachment.FileName.endswith((".xlsx", ".xls")):
                continue
            
            # Si el correo tiene asunto carga diaria o venta perdida
            type_file = None
            sheet_name = None
            if transportista['mail_subject'].lower() in subject_lower:
                type_file = transportista['name']
                sheet_name = transportista['mail_sheet_name']
            elif ruta['mail_subject'].lower() in subject_lower:
                type_file = ruta['name']
                sheet_name = ruta['mail_sheet_name']
            else: # Si no coincide ningun asunto saltar
                continue
            
            # Guardar archivo excel con alguno de los asuntos
            file_name = attachment.FileName
            file_address = os.path.join(root_address + test_address, file_name)
            attachment.SaveAsFile(file_address)

            # Validar contenido
            try:
                df = pd.read_excel(file_address, sheet_name=sheet_name, header=None)
                df = fix_misplaced_headers(df) # Corregir encabezado
                df_clean = df.dropna(how='all') # Elminar todos los NaN
                if df.empty or df_clean.shape[0] == 0: # Si no hay datos eliminar y saltar
                    os.remove(file_address)
                    continue
            except Exception:
                os.remove(file_address)
                continue

            # Guardar archivo para esa fecha
            if string_date_mail not in files_found:
                files_found[string_date_mail] = {}

            # Guardar tipo de archivo en esa fecha
            files_found[string_date_mail][type_file] = {
                'file_address': file_address,
                'received_time': string_date_mail
            }

            # Checar si ya tenemos ambos para esta fecha
            if (transportista['name'] in files_found[string_date_mail]) and (ruta['name'] in files_found[string_date_mail]):
                datos = files_found[string_date_mail]

                print(f'DATOS: {datos}')
                
                return (
                    datos[transportista['name']]['file_address'], datos[transportista['name']]['received_time'],
                    datos[ruta['name']]['file_address'], datos[ruta['name']]['received_time']
                )