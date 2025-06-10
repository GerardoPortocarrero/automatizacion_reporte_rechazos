import polars as pl
import math

# Función para asignar la meta desde el diccionario
def get_meta(meta, loc, año, mes):
    try:
        return meta[loc][año][mes]
    except KeyError:
        return 0.0
    
# Resumen de CF de archivo transportista
def update_transportista_resumen_file(meta, df):
    # Asegurar tipos
    df = df.with_columns([
        pl.col("Fecha").str.strptime(pl.Date, "%d/%m/%Y", strict=False),
        pl.col("Carga Pvta CF").cast(pl.Float64),
        pl.col("Rechazo CF").cast(pl.Float64)
    ])

    # Agregar columnas de año y mes para usar en la función meta
    df = df.with_columns([
        pl.col("Fecha").dt.year().alias("Año"),
        pl.col("Fecha").dt.month().alias("Mes")
    ])

    # Agrupar por Fecha y Locación
    df = df.group_by(["Fecha", "Locación", "Año", "Mes"]).agg([
        pl.sum("Carga Pvta CF").alias("Total CF Programandas"),
        pl.sum("Rechazo CF").alias("Total CF Rechazadas")
    ])

    # Aplicar la función de meta
    df = df.with_columns([
        pl.struct(["Locación", "Año", "Mes"]).map_elements(
            lambda x: get_meta(meta, x["Locación"], x["Año"], x["Mes"]),
            return_dtype=pl.Float64
        ).alias("Meta")
    ])

    # Calcular columnas adicionales
    df = df.with_columns([
        pl.col("Total CF Programandas").round(2),
        pl.col("Total CF Rechazadas").round(2),
        (pl.col("Meta") / 2).alias("Alerta"),
        (pl.col("Meta") * 2).alias("End"),
        (
            (pl.col("Total CF Rechazadas") / pl.col("Total CF Programandas"))
            .fill_nan(0)
            .fill_null(0)
        ).alias("% Rechazo")
    ])

    # Formato final (opcional: redondear y convertir % a string)
    df = df.select([
        "Fecha",
        "Locación",
        "Total CF Programandas",
        "Total CF Rechazadas",
        pl.col("Meta").map_elements(lambda x: f"{math.ceil(x * 100) / 100:.2f}%", return_dtype=pl.Utf8).alias("Meta"),
        pl.col("Alerta").map_elements(lambda x: f"{math.ceil(x * 100) / 100:.2f}%", return_dtype=pl.Utf8).alias("Alerta"),
        pl.col("End").map_elements(lambda x: f"{math.ceil(x * 100) / 100:.2f}%", return_dtype=pl.Utf8).alias("End"),
        pl.col("% Rechazo").map_elements(lambda x: f"{x:.2%}", return_dtype=pl.Utf8).alias("% Rechazo")
    ])

    # print(f'Transportista resumen: {df.schema}')
    # print(df.shape)

    return df