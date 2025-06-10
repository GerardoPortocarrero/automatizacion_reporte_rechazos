import os
from datetime import datetime
import polars as pl
import pandas as pd

# Eliminar solamente columnas llamadas Unnamed
def delete_unnamed_columns(df):    
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        
    return df

# Leer datos del archivo local
def read_local_file(local_file_address):
    df_local = pd.read_csv(local_file_address, sep=';')
    df_local = delete_unnamed_columns(df_local)
    df_local = pl.from_pandas(df_local)

    return df_local

# Hacer calculos para los porcentajes y totales
def make_calculations_for_locations(root_address, test_address, locaciones, transportista, ruta, month, year):
    local_transportista_address = os.path.join(root_address+test_address, transportista['local_file_name'])
    local_ruta_address = os.path.join(root_address+test_address, ruta['local_file_name'])

    transportista_updated = read_local_file(local_transportista_address)
    ruta_updated = read_local_file(local_ruta_address)

    calculations = {}
    # Crear fechas de inicio y fin como objetos datetime de Python
    started_date = datetime.strptime(f"01/{month}/{year}", "%d/%m/%Y")
    # Aseguramos que si month = 12, se pase al siguiente año
    if int(month) == 12:
        ended_date = datetime.strptime(f"01/01/{int(year) + 1}", "%d/%m/%Y")
    else:
        ended_date = datetime.strptime(f"01/{int(month)+1}/{year}", "%d/%m/%Y")

    # Convertir la columna a fecha
    transportista_updated = transportista_updated.with_columns(
        pl.col(transportista["date"]).str.strptime(pl.Date, "%d/%m/%Y", strict=False)
    )

    ruta_updated = ruta_updated.with_columns(
        pl.col(ruta["date"]).str.strptime(pl.Date, "%d/%m/%Y", strict=False)
    )

    # Filtrado con fechas
    transportista_updated = transportista_updated.filter(
        (pl.col(transportista["date"]) >= started_date) &
        (pl.col(transportista["date"]) < ended_date)
    )

    ruta_updated = ruta_updated.filter(
        (pl.col(ruta["date"]) >= started_date) &
        (pl.col(ruta["date"]) < ended_date)
    )

    for locacion in locaciones:
        # Filtrar usando Polars
        df_transportista = transportista_updated.filter(pl.col("Locación") == locacion)
        df_ruta = ruta_updated.filter(pl.col("Locación") == locacion)

        # Sumar columnas específicas
        venta_perdida_cf = df_ruta.select(pl.col("Venta Perdida CF").sum()).item()
        carga_total = df_transportista.select(pl.col("Carga Pvta CF").sum()).item()

        # Calcular porcentaje
        porcentaje_cf = (venta_perdida_cf * 100 / carga_total) if venta_perdida_cf > 0 else 0

        # Agregar resultado
        calculations[locacion] = {
            "locacion": locacion,
            "porcentaje_cf": f"{porcentaje_cf:.2f}%",
            "venta_perdida_cf": venta_perdida_cf,
            "carga_total": carga_total
        }
    
    return calculations