from tabulate import tabulate

# Mostrar informacion del dataframe
def show_df(df):
    print("\n📊 Resumen de DataFrame:")
    for col in df.columns:
        print(f"{col:<20} {df[col].dtype}")
    print(f"\n🔢 Dimensión: {df.shape[0]} filas × {df.shape[1]} columnas")
    print(f"💾 Memoria usada: {df.memory_usage(deep=True).sum()/1024:.2f} KB\n")

def show_document(document):
    # Preparar datos para tabla
    rows = [(k, ", ".join(v) if isinstance(v, list) else v) for k, v in document.items()]
    print(tabulate(rows, headers=["Campo", "Valor"], tablefmt="fancy_grid"))