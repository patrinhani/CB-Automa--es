import os
import fitz  # PyMuPDF
import hashlib
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import ttk

def selecionar_pasta():
    pasta = filedialog.askdirectory(title="Selecione a pasta com PDFs")
    if pasta:
        entry_pasta.delete(0, tk.END)
        entry_pasta.insert(0, pasta)

def hash_pdf(path):
    try:
        with open(path, "rb") as f:
            content = f.read()  # Ler o conteúdo binário do PDF
        return hashlib.md5(content).hexdigest()  # Retorna o hash MD5 do conteúdo binário
    except Exception as e:
        log_text.insert(tk.END, f"Erro ao ler {path}: {e}\n")
        return None

def analisar_e_mover():
    pasta = entry_pasta.get().strip()
    if not os.path.isdir(pasta):
        messagebox.showwarning("Aviso", "Selecione uma pasta válida.")
        return
    
    arquivos_pdf = [os.path.join(pasta, f) for f in os.listdir(pasta) if f.lower().endswith(".pdf")]
    if not arquivos_pdf:
        messagebox.showinfo("Nenhum PDF", "Nenhum arquivo PDF encontrado.")
        return
    
    hash_map = {}
    duplicados = []
    pasta_duplicados = os.path.join(pasta, "duplicados")
    os.makedirs(pasta_duplicados, exist_ok=True)
    
    total_arquivos = len(arquivos_pdf)
    progress_bar["maximum"] = total_arquivos  # Definir o valor máximo da barra de progresso
    progress_bar["value"] = 0  # Inicializar o valor da barra de progresso

    for i, arquivo in enumerate(arquivos_pdf):
        h = hash_pdf(arquivo)
        if h:
            if h in hash_map:
                try:
                    destino = os.path.join(pasta_duplicados, os.path.basename(arquivo))
                    base, ext = os.path.splitext(destino)
                    count = 1
                    while os.path.exists(destino):
                        destino = f"{base}_{count}{ext}"
                        count += 1
                    os.rename(arquivo, destino)
                    duplicados.append(destino)
                    log_text.insert(tk.END, f"Movido: {os.path.basename(arquivo)}\n")
                except Exception as e:
                    log_text.insert(tk.END, f"Erro ao mover {arquivo}: {e}\n")
            else:
                hash_map[h] = arquivo
        
        progress_bar["value"] = i + 1  # Atualiza a barra de progresso
        root.update_idletasks()  # Atualiza a interface durante o processo

    messagebox.showinfo("Concluído", f"{len(duplicados)} PDF(s) duplicado(s) movido(s).")
    if duplicados:
        log_text.insert(tk.END, f"\nArquivos duplicados movidos para: {pasta_duplicados}\n")
    else:
        log_text.insert(tk.END, "Nenhum duplicado encontrado.\n")

# === GUI ===
root = tk.Tk()  # Alterado para tk.Tk(), a janela principal do Tkinter
root.title("Remover PDFs Duplicados")
root.geometry("700x500")
root.minsize(600, 400)

# Layout
ttk.Label(root, text="Pasta com arquivos PDF:", font=("Segoe UI", 11)).pack(pady=(15, 5))
f1 = ttk.Frame(root)
f1.pack(padx=15, fill="x")
entry_pasta = ttk.Entry(f1)
entry_pasta.pack(side="left", expand=True, fill="x", padx=(0, 5))
ttk.Button(f1, text="Selecionar", command=selecionar_pasta).pack(side="left")

ttk.Button(
    root,
    text="Analisar e Mover Duplicados",
    command=analisar_e_mover,
    bootstyle="danger"
).pack(pady=(20, 10))

ttk.Label(root, text="Log de Arquivos Movidos:", font=("Segoe UI", 10)).pack()
frame_log = ttk.Frame(root)
frame_log.pack(fill="both", expand=True, padx=15, pady=(5, 15))

log_text = tk.Text(frame_log, wrap="word", bg="#1e1e1e", fg="#eeeeee")
log_text.pack(side="left", fill="both", expand=True)

scrollbar = ttk.Scrollbar(frame_log, command=log_text.yview)
scrollbar.pack(side="right", fill="y")
log_text.config(yscrollcommand=scrollbar.set)

# Barra de progresso
progress_bar = ttk.Progressbar(root, length=500, mode="determinate")
progress_bar.pack(pady=10)

root.mainloop()
