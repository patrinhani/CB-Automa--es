import os
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import queue
from concurrent.futures import ThreadPoolExecutor

try:
    import ttkbootstrap as ttk
    from ttkbootstrap.constants import *
    usar_ttk = True
except ImportError:
    from tkinter import ttk
    usar_ttk = False

import pythoncom
import win32com.client

def extrair_anexos_outlook(caminho_msg, pasta_saida, pasta_erro, log_queue):
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application")
    try:
        mail = outlook.CreateItemFromTemplate(os.path.abspath(caminho_msg))
    except Exception as e:
        erro_msg = f"Erro ao abrir {caminho_msg}: {e}"
        print(erro_msg)
        log_queue.put(erro_msg)
        mover_para_erro(caminho_msg, pasta_erro)
        return -1

    os.makedirs(pasta_saida, exist_ok=True)
    total = 0

    for att in mail.Attachments:
        nome = att.FileName or f"anexo_{total}"
        destino = os.path.join(pasta_saida, nome)

        base, ext = os.path.splitext(nome)
        count = 1
        while os.path.exists(destino):
            destino = os.path.join(pasta_saida, f"{base}_{count}{ext}")
            count += 1

        try:
            att.SaveAsFile(destino)
            total += 1
        except Exception as e:
            erro_msg = f"Falha ao salvar {nome} de {caminho_msg}: {e}"
            print(erro_msg)
            log_queue.put(erro_msg)
    return total

def mover_para_erro(arquivo, pasta_erro):
    os.makedirs(pasta_erro, exist_ok=True)
    nome = os.path.basename(arquivo)
    destino = os.path.join(pasta_erro, nome)

    base, ext = os.path.splitext(nome)
    count = 1
    while os.path.exists(destino):
        destino = os.path.join(pasta_erro, f"{base}_{count}{ext}")
        count += 1
    os.rename(arquivo, destino)

def selecionar_pasta_msgs():
    pasta = filedialog.askdirectory(title="Selecione a pasta com arquivos .msg")
    if pasta:
        entry_msg.delete(0, tk.END)
        entry_msg.insert(0, pasta)

def selecionar_pasta_destino():
    pasta = filedialog.askdirectory(title="Selecione a pasta de saída")
    if pasta:
        entry_dest.delete(0, tk.END)
        entry_dest.insert(0, pasta)

def iniciar_extracao():
    pasta_msgs = entry_msg.get().strip()
    pasta_saida = entry_dest.get().strip()

    if not os.path.isdir(pasta_msgs):
        messagebox.showwarning("Aviso", "Selecione uma pasta válida contendo arquivos .msg.")
        return
    if not os.path.isdir(pasta_saida):
        messagebox.showwarning("Aviso", "Selecione uma pasta de destino válida.")
        return

    arquivos = [os.path.join(pasta_msgs, f) for f in os.listdir(pasta_msgs) if f.lower().endswith(".msg")]
    if not arquivos:
        messagebox.showinfo("Nenhum arquivo", "Nenhum arquivo .msg encontrado na pasta.")
        return

    barra_progresso["maximum"] = len(arquivos)
    barra_progresso["value"] = 0
    progresso_label.config(text="Iniciando...")

    threading.Thread(target=processar_em_threads, args=(arquivos, pasta_saida), daemon=True).start()

def processar_em_threads(lista_arquivos, pasta_saida):
    total_anexos = 0
    concluido = 0
    total = len(lista_arquivos)
    pasta_erros = os.path.join(pasta_saida, "erros")
    log_queue = queue.Queue()

    def salvar_log_erros():
        if log_queue.empty():
            return
        log_path = os.path.join(pasta_erros, "log_erros.txt")
        with open(log_path, "a", encoding="utf-8") as f:
            while not log_queue.empty():
                f.write(log_queue.get() + "\n")

    def worker(caminho):
        nonlocal total_anexos, concluido
        extraidos = extrair_anexos_outlook(caminho, pasta_saida, pasta_erros, log_queue)
        concluido += 1
        if extraidos > 0:
            total_anexos += extraidos
        root.after(0, atualizar_progresso)

    def atualizar_progresso():
        barra_progresso["value"] = concluido
        progresso_label.config(text=f"{concluido}/{total} arquivos processados...")

        if concluido == total:
            salvar_log_erros()
            progresso_label.config(text="Concluído!")
            messagebox.showinfo(
                "Finalizado",
                f"{total_anexos} anexo(s) extraído(s).\nArquivos com erro foram movidos para: {pasta_erros}"
            )

    with ThreadPoolExecutor(max_workers=6) as executor:
        for caminho in lista_arquivos:
            executor.submit(worker, caminho)

# === GUI ===
if usar_ttk:
    root = ttk.Window(title="Extrator de Anexos .msg - Paralelo", themename="darkly")
else:
    root = tk.Tk()
    root.title("Extrator de Anexos .msg - Paralelo")

root.geometry("600x350")
root.resizable(False, False)

ttk.Label(root, text="Pasta com arquivos .msg:").pack(pady=(12, 0))
f1 = ttk.Frame(root); f1.pack(padx=15, fill="x")
entry_msg = ttk.Entry(f1); entry_msg.pack(side="left", expand=True, fill="x", padx=(0, 5))
ttk.Button(f1, text="Selecionar", command=selecionar_pasta_msgs).pack(side="left")

ttk.Label(root, text="Pasta de destino:").pack(pady=(20, 0))
f2 = ttk.Frame(root); f2.pack(padx=15, fill="x")
entry_dest = ttk.Entry(f2); entry_dest.pack(side="left", expand=True, fill="x", padx=(0, 5))
ttk.Button(f2, text="Selecionar", command=selecionar_pasta_destino).pack(side="left")

ttk.Button(
    root,
    text="Extrair Anexos",
    command=iniciar_extracao,
    bootstyle="success" if usar_ttk else None
).pack(pady=(25, 10))

barra_progresso = ttk.Progressbar(root, length=500, mode="determinate")
barra_progresso.pack(pady=(10, 5))

progresso_label = ttk.Label(root, text="")
progresso_label.pack()

root.mainloop()
