import os
import fitz    # PyMuPDF
import hashlib
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import re

# === Funções auxiliares ===

def selecionar_pasta():
    pasta = filedialog.askdirectory(title="Selecione a pasta com PDFs")
    if pasta:
        entry_pasta.delete(0, tk.END)
        entry_pasta.insert(0, pasta)

def extrair_texto_pdf(path):
    try:
        with fitz.open(path) as doc:
            return "".join(page.get_text() or "" for page in doc)
    except Exception as e:
        log_text.insert(tk.END, f"[ERRO] {os.path.basename(path)}: {e}\n")
        return ""

def normalizar_texto(texto: str) -> str:
    texto = texto.lower()
    texto = re.sub(r"\d+", " ", texto)        # remove números
    texto = re.sub(r"[^\w\s]", " ", texto)    # remove pontuação
    texto = re.sub(r"\s+", " ", texto)        # unifica espaços
    return texto.strip()

def hash_texto(texto: str) -> str:
    return hashlib.sha256(texto.encode("utf-8")).hexdigest()

def analisar_em_thread():
    threading.Thread(target=analisar_e_mover, daemon=True).start()

# === Lógica principal ===

def analisar_e_mover():
    pasta = entry_pasta.get().strip()
    if not os.path.isdir(pasta):
        messagebox.showwarning("Aviso", "Selecione uma pasta válida.")
        return

    pdfs = [os.path.join(pasta, f) for f in os.listdir(pasta) if f.lower().endswith(".pdf")]
    if not pdfs:
        messagebox.showinfo("Nenhum PDF", "Nenhum arquivo PDF encontrado.")
        return

    hashes = {}    # hash_normalizado -> caminho_original
    duplicados = []
    pasta_dup = os.path.join(pasta, "duplicados")
    os.makedirs(pasta_dup, exist_ok=True)

    total = len(pdfs)
    progress_bar["maximum"] = total

    for i, caminho in enumerate(pdfs, start=1):
        texto = extrair_texto_pdf(caminho)
        norm = normalizar_texto(texto)
        h = hash_texto(norm)

        if h in hashes:
            # duplicado exato de texto normalizado
            base, ext = os.path.splitext(os.path.basename(caminho))
            destino = os.path.join(pasta_dup, f"{base}{ext}")
            cnt = 1
            while os.path.exists(destino):
                destino = os.path.join(pasta_dup, f"{base}_{cnt}{ext}")
                cnt += 1

            try:
                os.rename(caminho, destino)
                duplicados.append(destino)
                log_text.insert(
                    tk.END,
                    f"Movido (hash idêntico): {os.path.basename(caminho)}\n"
                )
            except Exception as e:
                log_text.insert(tk.END, f"[ERRO] {os.path.basename(caminho)}: {e}\n")
        else:
            hashes[h] = caminho

        # atualiza progresso
        progress_bar["value"] = i
        progress_label.config(text=f"Progresso: {int(i/total*100)}%")
        root.update_idletasks()

    messagebox.showinfo("Concluído", f"{len(duplicados)} duplicado(s) movido(s).")
    if duplicados:
        log_text.insert(tk.END, f"\nArquivos em '{pasta_dup}'.\n")
    else:
        log_text.insert(tk.END, "Nenhum duplicado encontrado.\n")

# === GUI ===

root = ttk.Window(themename="superhero")
root.title("Remover PDFs Duplicados por Hash de Texto")
root.geometry("700x500")
root.minsize(600, 400)

ttk.Label(root, text="Pasta com arquivos PDF:", font=("Segoe UI",11)).pack(pady=(15,5))
frame1 = ttk.Frame(root); frame1.pack(padx=15, fill="x")
entry_pasta = ttk.Entry(frame1); entry_pasta.pack(side="left", expand=True, fill="x", padx=(0,5))
ttk.Button(frame1, text="Selecionar", command=selecionar_pasta).pack(side="left")

ttk.Button(
    root,
    text="Analisar e Mover Duplicados",
    command=analisar_em_thread,
    bootstyle="danger"
).pack(pady=(20,10))

ttk.Label(root, text="Log de arquivos:", font=("Segoe UI",10)).pack()
frame_log = ttk.Frame(root)
frame_log.pack(fill="both", expand=True, padx=15, pady=(5,10))
log_text = tk.Text(frame_log, wrap="word", bg="#1e1e1e", fg="#eeeeee")
log_text.pack(side="left", fill="both", expand=True)
ttk.Scrollbar(frame_log, command=log_text.yview).pack(side="right", fill="y")
log_text.config(yscrollcommand=lambda f,l:None)

progress_label = ttk.Label(root, text="Progresso: 0%", font=("Segoe UI",10))
progress_label.pack()
progress_bar = ttk.Progressbar(root, mode="determinate", bootstyle="info-striped")
progress_bar.pack(fill="x", expand=True, padx=15, pady=10)

root.mainloop()
