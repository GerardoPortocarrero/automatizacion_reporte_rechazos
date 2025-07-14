import os
import pandas as pd
import polars as pl
from datetime import datetime
import shutil
import glob
import log_management as log

# Crear archivo csv si no existe
def create_csv_from_scratch(document, root_address, project_address):
    df = pd.read_excel(os.path.join(root_address, document['source_local_file_name']), sheet_name=document['source_local_sheet_name'])

    # Guardar a CSV
    df.to_csv(os.path.join(project_address, document['local_file_name']), index=False, encoding='utf-8-sig')

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

# Obtener las fechas mas recientes de local y mail
def get_most_recent_dates(document, df_mail, df_local):
    df_mail_copy = df_mail.copy()
    df_local_copy = df_local.copy()

    # Convertir columna 'Fecha2' a datetime en archivo local y correo
    df_mail_copy[document['date']+"2"] = pd.to_datetime(df_mail_copy[document['date']], dayfirst=True, errors='coerce')
    df_local_copy[document['date']+"2"] = pd.to_datetime(df_local_copy[document['date']], errors='coerce')    

    # Fecha máxima en archivo de correo y local
    fecha_max_mail = df_mail_copy[document['date']+"2"].max()
    fecha_max_local = df_local_copy[document['date']+"2"].max()    
    
    return fecha_max_mail, fecha_max_local

# Eliminar datos del mes de local
def delete_local_month(df, date, date_column):
    # Crear una copia temporal de la columna como datetime para filtrar
    temporal_dates = pd.to_datetime(df[date_column], errors='coerce', dayfirst=True)

    # Filtrar las filas que NO pertenecen al mismo mes y año que `date`
    df_filtrado = df[~((temporal_dates.dt.month == date.month) & (temporal_dates.dt.year == date.year))]

    return df_filtrado

# Extraer datos que no se encuentran en el archivo local por fecha
def concat_polar_dataframes(df_mail, df_local, date_column):
    # Convertir a Polars
    df_local = pl.from_pandas(df_local)
    df_mail = pl.from_pandas(df_mail)

    # Forzar tipado de df_mail igual al de df_local (excepto fecha)
    try:
        schema_local = df_local.schema
        schema_to_cast = {k: v for k, v in schema_local.items() if k != date_column}
        df_mail = df_mail.cast(schema_to_cast)
    except Exception as e:
        print("⚠️ Error al castear tipos en Polars:", e)

    if df_mail.schema != df_local.schema:
        print("⚠️  Los esquemas entre los DataFrames no coinciden:")
        print("──────────────────────────────────────────────")
        print("📁 Archivo LOCAL:")
        print(f"   🧬 Schema : {df_local.schema}")
        print(f"   🔢 Shape  : {df_local.shape}")
        print("📥 Archivo del CORREO:")
        print(f"   🧬 Schema : {df_mail.schema}")
        print(f"   🔢 Shape  : {df_mail.shape}")
        print("──────────────────────────────────────────────\n")

    # Combinar los datos (sin eliminar duplicados)
    df_updated = pl.concat([df_local, df_mail])

    return df_updated

# Actualizar archivo si tiene columna 'Mesa Comercial'
def customized_ruta_mail_file(df, vendedores):
    if 'Mesa Comercial' in df.columns:

        df = df.drop(columns=['Mesa Comercial'])

        def get_nombre_vendedor(row):
            sede = row.get('Locación')
            codigo = row.get('VendedorCod')
            if sede in vendedores and codigo in vendedores[sede]:
                return vendedores[sede][codigo]
            return None

        df['Nombre Vendedor'] = df.apply(get_nombre_vendedor, axis=1)

    return df

# Actualizar codigos de transportistas
def set_transportista_code_mail_file(df, document, transportistas_code):
    
    def get_codigo_transportista(row):
        sede = row.get('Locación')
        transportista = row.get('Transportista')
        if sede in transportistas_code and transportista in transportistas_code[sede]:
            return transportistas_code[sede][transportista]
        return None

    df[document['transportista']] = df.apply(get_codigo_transportista, axis=1)

    return df

# Escribir al CSV sobrescribiendo el original
def backup_local_file_changes(project_address, document, backup_address):
    # Copiar archivo
    # copy: no copia metadatos (fecha, permisos)
    # copy2: si copia metadatos (fecha, permisos)
    if os.path.exists(document['local_file_address']):
        shutil.copy(document['local_file_address'], backup_address)

        text = f'[✓] Backup de ({document['local_file_name']}) generado correctamente'
        log.write_log(project_address, text)

# Escribir al CSV sobrescribiendo el original
def save_local_file_changes(project_address, root_address, df_updated, document):
    text = f'[✓] Archivo ({document['local_file_name']}) guardado Correctamente'
    log.write_log(project_address, text)
    
    df_updated_project = df_updated.write_csv(separator=",")
    with open(document['local_file_address'], "w", encoding="utf-8-sig") as f:
        f.write(df_updated_project)

    df_update_root = df_updated.write_csv(separator=";")
    with open(os.path.join(root_address, document['local_file_name']), "w", encoding="utf-8-sig") as f:
        f.write(df_update_root)

# Eliminar archivo del correo
def delete_mail_files(project_address):
    # Buscar todos los archivos .xlsx en la carpeta del proyecto
    archivos_excel = glob.glob(os.path.join(project_address, "*.xlsx"))

    # Eliminar cada archivo encontrado
    for archivo in archivos_excel:
        try:
            os.remove(archivo)
        except Exception as e:
            pass

    text = f'[✓] Archivos de correo eliminados'
    log.write_log(project_address, text)

# Leer datos del archivo local
def read_local_file(local_file_address):
    df_local = pd.read_csv(local_file_address, sep=',')
    df_local = delete_unnamed_columns(df_local)

    return df_local

# Actualizar el archivo local con los datos del correo
def update_local_file(document, locaciones, vendedores, transportistas_code):
    # Rutas de archivo
    mail_file_address = document['mail_file_address']
    mail_sheet_name = document['mail_sheet_name']
    local_file_address = document['local_file_address']
    date_column = document['date']

    # Leer datos del archivo local
    df_local = read_local_file(local_file_address)

    # Leer datos del archivo de correo
    df_mail = pd.read_excel(mail_file_address, sheet_name=mail_sheet_name, header=None)
    df_mail = fix_misplaced_headers(df_mail)

    # Filtrar datos del archivo de correo
    df_mail = filter_mail_file_locations(df_mail, locaciones)
    mail_most_recent_date, local_most_recent_date = get_most_recent_dates(document, df_mail, df_local)

    # Si hay datos nuevos para actualizar el archivo local
    if mail_most_recent_date > local_most_recent_date:
        # Configuracion solo para el archivo de ruta
        df_mail = customized_ruta_mail_file(df_mail, vendedores)

        # Actualizar el codigo de transportista para ambos archivos
        df_mail = set_transportista_code_mail_file(df_mail, document, transportistas_code)

        # Igualar columnas
        df_mail.columns = df_local.columns

        # Forzar columnas a tipo string si están completamente vacías para evitar el tipo Null
        for col in df_mail.columns:
            if df_mail[col].isnull().all():
                df_mail[col] = df_mail[col].astype(str)
        
        # Convertir la fecha a string, incluso si los valores son datetime dentro de "object"
        df_mail[date_column] = df_mail[date_column].apply(lambda x: x.strftime('%Y-%m-%d') if isinstance(x, pd.Timestamp) or isinstance(x, datetime) else str(x))        

        # Eliminar los datos del mes del archivo local
        df_local = delete_local_month(df_local, mail_most_recent_date, date_column)

        # Concatenar archivos de correo y local
        df_updated = concat_polar_dataframes(df_mail, df_local, date_column)

        return df_updated, False, mail_most_recent_date.strftime('%Y-%m-%d')
    
    else:
        print("ℹ️  No se encontraron datos nuevos para actualizar.")
        print("──────────────────────────────────────────────")
        print(f"📄 Archivo: {mail_file_address}")
        print(f"📑 Hoja   : '{mail_sheet_name}'")
        print("──────────────────────────────────────────────\n")

        # Convertir a polars
        df_local = pl.from_pandas(df_local)

        return df_local, True, local_most_recent_date.strftime('%Y-%m-%d')