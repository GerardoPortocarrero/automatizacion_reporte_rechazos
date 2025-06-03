import pandas as pd

def delete_unnamed_columns(df):
    # Eliminar columnas llamadas Unnamed
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        
    return df

def clean_mail_file(df):
    # Eliminar columnas unnamed (vacias)
    df = df.dropna(axis=1, how='all')

    # Tomar la fila 1 como nombres de columnas
    df.columns = df.iloc[1]

    # Eliminar las dos primeras filas (la original de encabezado y la fila de nombres)
    df = df.iloc[2:].reset_index(drop=True)
    
    return df

def filter_mail_file_locations(df, locaciones):
    # Conservar las filas con locaciones especificas
    df = df[df['Locación'].isin(locaciones)]

    return df

def filter_mail_file_dates(document, df_mail, df_local):
    df_local_copy = df_local.copy()

    # Convertir columna 'Fecha' a datetime en archivo recibido
    df_local_copy[document['date']+"2"] = pd.to_datetime(df_local_copy[document['date']], dayfirst=True, errors='coerce')

    # Fecha máxima en archivo local
    fecha_max_local = df_local_copy[document['date']+"2"].max()

    # Filtrar solo filas con fecha mayor que la máxima del local
    df_mail = df_mail[df_mail[document['date']] > fecha_max_local].copy()

    return df_mail

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

def save_local_file_changes(df_updated, document):
    # Escribir al CSV, sobrescribiendo el original
        df_updated = df_updated.write_csv(separator=";")
        with open(document['local_file_address'], "w", encoding="utf-8-sig") as f:
            f.write(df_updated)
