import os
import pandas as pd
import polars as pl
from datetime import datetime

# Eliminar solamente columnas llamadas Unnamed
def delete_unnamed_columns(df):    
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        
    return df

# Eliminar columna y fila vacia y asignar el encabezado
def fix_misplaced_headers(df):
    # Eliminar columnas unnamed (vacias)
    df = df.dropna(axis=1, how='all')

    # Tomar la fila 1 como nombres de columnas
    df.columns = df.iloc[1]
    
    # Eliminar las dos primeras filas (la original de encabezado y la fila de nombres)
    df = df.iloc[2:].reset_index(drop=True)
    
    return df

# Extraer los datos filtrando por locacion
def filter_mail_file_locations(df, locaciones):
    df = df[df['Locación'].isin(locaciones)]

    return df

# Extraer datos que no se encuentran en el archivo local por fecha
def filter_mail_file_dates(document, df_mail, df_local):
    df_local_copy = df_local.copy()

    # Convertir columna 'Fecha2' a datetime en archivo recibido
    df_local_copy[document['date']+"2"] = pd.to_datetime(df_local_copy[document['date']], dayfirst=True, errors='coerce')

    # Fecha máxima en archivo local
    fecha_max_local = df_local_copy[document['date']+"2"].max()

    # Filtrar solo filas con fecha mayor que la máxima del local
    df_mail = df_mail[df_mail[document['date']] > fecha_max_local].copy()

    return df_mail

# Actualizar archivo si tiene columna 'Mesa Comercial'
def customized_ruta_mail_file(df, vendedores):
    if 'Mesa Comercial' in df.columns:
        print("🛠️ Procesando archivo con 'Mesa Comercial'")

        df = df.drop(columns=['Mesa Comercial'])

        def get_nombre_vendedor(row):
            sede = row.get('Locación')
            codigo = row.get('VendedorCod')
            if sede in vendedores and codigo in vendedores[sede]:
                return vendedores[sede][codigo]
            return None

        df['Nombre Vendedor'] = df.apply(get_nombre_vendedor, axis=1)
    else:
        print("📄 Archivo sin 'Mesa Comercial'. No se hace nada.")

    return df

# Escribir al CSV sobrescribiendo el original
def save_local_file_changes(df_updated, document):    
    df_updated = df_updated.write_csv(separator=";")
    with open(document['local_file_address'], "w", encoding="utf-8-sig") as f:
        f.write(df_updated)

# Actualizar el archivo local con los datos del correo
def update_local_file(document, locaciones, root_address, test_address, vendedores):
    # Rutas de archivo
    mail_file_address = document['mail_file_address']
    mail_sheet_name = document['mail_sheet_name']
    local_file_address = os.path.join(root_address+test_address, document['local_file_name'])

    # Leer datos del archivo local
    df_local = pd.read_csv(local_file_address, sep=';')
    df_local = delete_unnamed_columns(df_local)

    # Leer datos del archivo de correo
    df_mail = pd.read_excel(mail_file_address, sheet_name=mail_sheet_name, header=None)
    df_mail = fix_misplaced_headers(df_mail)

    # Filtrar datos del archivo de correo
    df_mail = filter_mail_file_locations(df_mail, locaciones)
    df_mail = filter_mail_file_dates(document, df_mail, df_local)

    # Si hay datos nuevos para actualizar el archivo local
    if len(df_mail) > 0:
        # Configuracion solo para el archivo de ruta
        df_mail = customized_ruta_mail_file(df_mail, vendedores)

        # Asegurar los mismos tipos de datos (str, int, ...)
        df_mail.columns = df_local.columns
        df_mail = df_mail.astype(df_local.dtypes.to_dict())

        # Convertir la fecha a string, incluso si los valores son datetime dentro de "object"
        df_mail[document['date']] = df_mail[document['date']].apply(lambda x: x.strftime('%d/%m/%Y') if isinstance(x, pd.Timestamp) or isinstance(x, datetime) else str(x))
        #print(document['date'].apply(type).value_counts())

        # Convertir a Polars
        df_local = pl.from_pandas(df_local)
        df_mail = pl.from_pandas(df_mail)

        if df_mail.schema != df_local.schema:
            print("⚠️ Los esquemas no coinciden")
            print(f'Local schema: {df_local.schema}')
            print(f'Local shape: {df_local.shape}')
            print(f'Mail schema: {df_mail.schema}')
            print(f'Mail shape: {df_mail.shape}')

        # Combinar los datos (sin eliminar duplicados)
        df_updated = pl.concat([df_local, df_mail])

        # print(f'Updated: {df_mail.schema}')
        # print(df_updated.shape)

        return df_updated
    
    else:
        print(f"⚠️ No hay datos nuevos en el archivo de correo '{mail_file_address}' para la hoja '{mail_sheet_name}'.")

        return pl.from_pandas(df_local)