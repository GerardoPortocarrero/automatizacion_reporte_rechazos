import tkinter as tk
from tkinter import messagebox

def main_interface(title):
    respuesta = {"valor": None}  # Contenedor mutable para capturar respuesta

    def abrir_archivo(text_widget):
        try:
            with open("log.txt", 'r', encoding='utf-8') as f:
                contenido = f.read()
            text_widget.delete(1.0, tk.END)
            text_widget.insert(tk.END, contenido)
        except FileNotFoundError:
            messagebox.showerror("Error", "No se encontró el archivo 'log.txt'.")

    def continuar_si(root):
        respuesta["valor"] = "sí"
        root.destroy()

    def continuar_no(root):
        respuesta["valor"] = "no"
        root.destroy()

    # Crear ventana
    root = tk.Tk()
    root.title(title)
    root.geometry("700x500")

    # Frame contenedor del texto con tamaño fijo
    frame_texto = tk.Frame(root, height=350)
    frame_texto.pack(fill=tk.BOTH, padx=20, pady=(20, 10))
    frame_texto.pack_propagate(False)

    texto = tk.Text(frame_texto, wrap=tk.WORD, font=("Arial", 12))
    texto.pack(expand=True, fill=tk.BOTH)
    abrir_archivo(texto)

    # Botones
    frame_botones = tk.Frame(root)
    frame_botones.pack(pady=10)

    btn_si = tk.Button(
        frame_botones,
        text="✅ Continuar: Sí",
        bg="lightgreen",
        font=("Arial", 10, "bold"),
        width=15,
        height=2,
        command=lambda: continuar_si(root)
    )
    btn_si.pack(side=tk.LEFT, padx=20)

    btn_no = tk.Button(
        frame_botones,
        text="❌ Continuar: No",
        bg="lightcoral",
        font=("Arial", 10, "bold"),
        width=15,
        height=2,
        command=lambda: continuar_no(root)
    )
    btn_no.pack(side=tk.LEFT, padx=20)

    root.mainloop()

    return respuesta["valor"]