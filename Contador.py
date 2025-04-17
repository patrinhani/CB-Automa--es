import os
import tkinter as tk
from tkinter import filedialog, messagebox

def contar_total_arquivos(caminho_raiz):
    total_arquivos = 0

    for item in os.listdir(caminho_raiz):
        caminho_pasta = os.path.join(caminho_raiz, item)
        if os.path.isdir(caminho_pasta):
            for root, _, files in os.walk(caminho_pasta):
                total_arquivos += len(files)

    return total_arquivos

def selecionar_pasta():
    caminho = filedialog.askdirectory()
    if caminho:
        total = contar_total_arquivos(caminho)
        texto_resultado.set(f"📁 Pasta selecionada:\n{caminho}\n\n📦 Total de arquivos encontrados: {total}")
    else:
        messagebox.showinfo("Aviso", "Nenhuma pasta selecionada.")

# Janela principal
janela = tk.Tk()
janela.title("Contador de Arquivos por Pasta")
janela.geometry("520x250")
janela.configure(bg="#1e1e1e")

# Fonte padrão
fonte_titulo = ("Arial", 16, "bold")
fonte_texto = ("Courier New", 11)

# Título
titulo = tk.Label(
    janela, 
    text="🔍 Contador Total de Arquivos", 
    font=fonte_titulo,
    fg="white", bg="#1e1e1e"
)
titulo.pack(pady=20)

# Botão de seleção
botao = tk.Button(
    janela, 
    text="Selecionar Pasta", 
    font=("Arial", 12),
    bg="#007acc", fg="white",
    activebackground="#005f99", 
    command=selecionar_pasta
)
botao.pack(pady=10)

# Resultado
texto_resultado = tk.StringVar()
texto_resultado.set("Nenhuma pasta selecionada.")

resultado_label = tk.Label(
    janela, 
    textvariable=texto_resultado, 
    font=fonte_texto, 
    bg="#1e1e1e", 
    fg="white", 
    justify="left", 
    wraplength=480
)
resultado_label.pack(padx=20, pady=10)

janela.mainloop()
