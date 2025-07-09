from tabulate import tabulate
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich import box

# Mostrar informacion del dataframe
def show_df(df):
    console = Console()

    table = Table(show_header=True, header_style="bold white on dark_red", box=box.SQUARE)
    table.add_column("🧱 Columna", style="bold cyan", no_wrap=True)
    table.add_column("📂 Tipo", style="bold magenta")
    table.add_column("✅ Non-Null", justify="right", style="green")
    table.add_column("📈 Completitud", justify="right", style="yellow")

    total_rows = len(df)
    for col in df.columns:
        non_nulls = df[col].notna().sum()
        tipo = str(df[col].dtype)
        completitud = f"{(non_nulls / total_rows * 100):.1f}%"
        table.add_row(str(col), tipo, str(non_nulls), completitud)

    # Resumen general
    mem_kb = df.memory_usage(deep=True).sum() / 1024
    resumen = (
        f"[bold yellow]🔢 Dimensión:[/bold yellow] {total_rows} filas × {len(df.columns)} columnas\n"
        f"[bold green]💾 Memoria usada:[/bold green] {mem_kb:.2f} KB"
    )

    console.print(Panel.fit(table, title="📊 Resumen de DataFrame"))
    console.print(Panel.fit(resumen, title="📎 Resumen", border_style="grey50"))

def show_document(document):
    # Preparar datos para tabla
    rows = [(k, ", ".join(v) if isinstance(v, list) else v) for k, v in document.items()]
    print(tabulate(rows, headers=["Campo", "Valor"], tablefmt="fancy_grid"))