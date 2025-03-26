import os
import re
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
from datetime import datetime
import threading
import numpy as np
from sklearn.cluster import KMeans
from collections import Counter

# Expressões regulares
regex_cpf = re.compile(r'(\d{10,})')
regex_data = re.compile(r'(\d{2}-\d{2}-\d{4})')

# Criando interface
root = ttk.Window(themename="darkly")
root.title("📂 Renomeador Inteligente")
root.geometry("950x550")

diretorio = ttk.StringVar()
count_corretos = ttk.IntVar(value=0)
count_ajustar = ttk.IntVar(value=0)

# -------------------- INTERFACE --------------------
frame_topo = ttk.Frame(root)
frame_topo.pack(pady=10, padx=10, fill=X)

ttk.Label(frame_topo, text="📁 Pasta de origem:", font=("Arial", 12, "bold")).pack(side=LEFT, padx=5)
entry_diretorio = ttk.Entry(frame_topo, width=50, textvariable=diretorio)
entry_diretorio.pack(side=LEFT, padx=5)

def selecionar_pasta():
    """Seleciona a pasta e reseta corretamente o caminho."""
    pasta = filedialog.askdirectory()
    if pasta:
        diretorio.set(pasta)

ttk.Button(frame_topo, text="Selecionar", bootstyle=PRIMARY, command=selecionar_pasta).pack(side=LEFT, padx=5)

# -------------------- BOTÕES --------------------
frame_botoes = ttk.Frame(root)
frame_botoes.pack(pady=10)

botao_analisar = ttk.Button(frame_botoes, text="🔍 Analisar Arquivos", bootstyle=INFO, 
                            command=lambda: threading.Thread(target=listar_arquivos).start())
botao_analisar.pack(side=LEFT, padx=10)

botao_renomear = ttk.Button(frame_botoes, text="✏️ Renomear Arquivos", bootstyle=SUCCESS, 
                            command=lambda: threading.Thread(target=renomear_arquivos).start(), state=DISABLED)
botao_renomear.pack(side=LEFT, padx=10)

label_status = ttk.Label(root, text="🔹 Selecione uma pasta para começar", font=("Arial", 10, "bold"), foreground="white")
label_status.pack(pady=5)

# -------------------- LISTA DE ARQUIVOS --------------------
frame_lista = ttk.Frame(root)
frame_lista.pack(padx=10, pady=10, fill=BOTH, expand=True)

scroll_y = ttk.Scrollbar(frame_lista, orient=VERTICAL)
scroll_x = ttk.Scrollbar(frame_lista, orient=HORIZONTAL)

lista_arquivos = ttk.Treeview(frame_lista, columns=("Arquivo", "Status"), show="headings", 
                              yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
scroll_y.config(command=lista_arquivos.yview)
scroll_x.config(command=lista_arquivos.xview)

scroll_y.pack(side=RIGHT, fill=Y)
scroll_x.pack(side=BOTTOM, fill=X)
lista_arquivos.pack(fill=BOTH, expand=True)

lista_arquivos.heading("Arquivo", text="📄 Arquivo")
lista_arquivos.heading("Status", text="🔍 Status")
lista_arquivos.tag_configure("correto", foreground="lightgreen")
lista_arquivos.tag_configure("ajustar", foreground="yellow")

# -------------------- CONTADORES --------------------
frame_status = ttk.Frame(root)
frame_status.pack(pady=10)

ttk.Label(frame_status, text="✅ Corretos:", font=("Arial", 10, "bold")).pack(side=LEFT, padx=5)
label_corretos = ttk.Label(frame_status, textvariable=count_corretos, font=("Arial", 12, "bold"), foreground="lightgreen")
label_corretos.pack(side=LEFT, padx=5)

ttk.Label(frame_status, text="⚠️ Ajustar:", font=("Arial", 10, "bold")).pack(side=LEFT, padx=5)
label_ajustar = ttk.Label(frame_status, textvariable=count_ajustar, font=("Arial", 12, "bold"), foreground="yellow")
label_ajustar.pack(side=LEFT, padx=5)

# -----------------------------------
# FUNÇÕES DO PROGRAMA
# -----------------------------------

def detectar_padroes(arquivos):
    """Usa K-Means para agrupar padrões nos nomes de arquivos."""
    if not arquivos:
        return ""

    tamanhos = np.array([[len(arq)] for arq in arquivos])
    kmeans = KMeans(n_clusters=2, random_state=42, n_init=10).fit(tamanhos)
    grupos = kmeans.labels_

    contador = Counter(grupos)
    grupo_mais_comum = max(contador, key=contador.get)
    nomes_comuns = [arquivos[i] for i in range(len(arquivos)) if grupos[i] == grupo_mais_comum]

    return nomes_comuns[0].split("_")[0] if nomes_comuns else ""

def corrigir_nome(arquivo, parte_fixa):
    """Corrige nome do arquivo baseado no padrão correto."""
    nome_limpo, ext = os.path.splitext(arquivo)
    partes = [p.strip() for p in nome_limpo.split("_") if p.strip()]
    problemas = []

    if parte_fixa and parte_fixa not in "_".join(partes):
        problemas.append("Parte fixa ausente")
        partes.insert(0, parte_fixa)

    for i, parte in enumerate(partes):
        if regex_cpf.fullmatch(parte):
            if len(parte) == 10:
                partes[i] = "0" + parte
                problemas.append("CPF com 10 dígitos")

    for i, parte in enumerate(partes):
        if regex_data.fullmatch(parte):
            data = datetime.strptime(parte, "%d-%m-%Y")
            if data > datetime.now():
                partes[i] = datetime.now().strftime("%d-%m-%Y")
                problemas.append("Data no futuro")

    nome_corrigido = "_".join(partes) + ext
    if nome_corrigido != arquivo:
        problemas.append("Espaços extras removidos")

    return nome_corrigido, problemas

def listar_arquivos():
    """Lista arquivos na pasta e verifica padrões."""
    pasta = diretorio.get().strip()

    if not pasta or not os.path.exists(pasta):
        messagebox.showerror("Erro", "Pasta inválida ou inexistente!")
        return

    botao_analisar.config(state=DISABLED)
    botao_renomear.config(state=DISABLED)
    label_status.config(text="🔄 Analisando...", foreground="yellow")
    lista_arquivos.delete(*lista_arquivos.get_children())

    count_corretos.set(0)
    count_ajustar.set(0)

    arquivos = [f for f in os.listdir(pasta) if os.path.isfile(os.path.join(pasta, f))]
    parte_fixa = detectar_padroes(arquivos)

    for arquivo in arquivos:
        nome_corrigido, problemas = corrigir_nome(arquivo, parte_fixa)
        if nome_corrigido and nome_corrigido != arquivo:
            lista_arquivos.insert("", "end", values=(arquivo, f"🟡 {nome_corrigido} ({', '.join(problemas)})"), tags=("ajustar",))
            count_ajustar.set(count_ajustar.get() + 1)
        else:
            lista_arquivos.insert("", "end", values=(arquivo, "✅ Correto"), tags=("correto",))
            count_corretos.set(count_corretos.get() + 1)

    label_status.config(text="✅ Análise concluída!", foreground="lightgreen")
    botao_renomear.config(state=NORMAL)
    botao_analisar.config(state=NORMAL)

def renomear_arquivos():
    """Renomeia os arquivos conforme a análise."""
    pasta = diretorio.get().strip()

    if not pasta or not os.path.exists(pasta):
        
        messagebox.showerror("Erro", "Pasta inválida ou inexistente!")
        return

    for item in lista_arquivos.get_children():
        nome_atual, status = lista_arquivos.item(item, "values")
        if "🟡" in status:
            os.rename(os.path.join(pasta, nome_atual), os.path.join(pasta, status.split(" ")[1]))

root.mainloop()
