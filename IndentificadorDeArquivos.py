import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
import pandas as pd
import os
import unicodedata
import difflib
import threading
import shutil

# Normalização dos nomes
def normalizar(texto):
    if not isinstance(texto, str):
        return ''
    texto = texto.strip()
    texto = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("ASCII")
    return texto.lower()

# Extrai apenas o nome do arquivo do caminho
def extrair_nome(caminho_completo):
    return os.path.basename(caminho_completo)

# Selecionar Excel
def selecionar_excel():
    caminho = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx *.xls")])
    if not caminho:
        return
    try:
        df = pd.read_excel(caminho)
        coluna = df.columns[0]
        df['Nome do Arquivo'] = df[coluna].apply(extrair_nome)
        app.df_nomes = df['Nome do Arquivo'].tolist()
        messagebox.showinfo("Sucesso", "Excel carregado com sucesso. Agora selecione a pasta onde os arquivos estão.")
    except Exception as e:
        messagebox.showerror("Erro", str(e))

# Selecionar pasta para busca
def selecionar_pasta_busca():
    if not hasattr(app, 'df_nomes'):
        messagebox.showwarning("Aviso", "Você precisa carregar um Excel primeiro.")
        return

    app.pasta_base = filedialog.askdirectory(title="Selecione a pasta com os arquivos")
    if not app.pasta_base:
        return

    threading.Thread(target=buscar_arquivos, daemon=True).start()

# Selecionar pasta para cópia
def selecionar_pasta_destino():
    if not hasattr(app, 'arquivos_encontrados') or not app.arquivos_encontrados:
        messagebox.showwarning("Aviso", "Você precisa localizar os arquivos antes de copiá-los.")
        return

    app.pasta_destino = filedialog.askdirectory(title="Selecione a pasta de destino para copiar os arquivos encontrados")
    if not app.pasta_destino:
        return

    threading.Thread(target=copiar_arquivos, daemon=True).start()

# Busca os arquivos
def buscar_arquivos():
    arquivos_disponiveis = {}
    lista_nomes_arquivos = []

    for root, dirs, files in os.walk(app.pasta_base):
        for file in files:
            nome_norm = normalizar(file)
            arquivos_disponiveis[nome_norm] = os.path.join(root, file)
            lista_nomes_arquivos.append(nome_norm)

    app.arquivos_encontrados = []

    total = len(app.df_nomes)
    progress_bar.config(mode="determinate", maximum=total, value=0)

    for i, nome_original in enumerate(app.df_nomes):
        nome_base = os.path.basename(nome_original)
        nome_norm = normalizar(nome_base)
        match = difflib.get_close_matches(nome_norm, lista_nomes_arquivos, n=1, cutoff=0.85)

        if match:
            caminho = arquivos_disponiveis[match[0]]
            app.arquivos_encontrados.append((nome_base, caminho))
        else:
            app.arquivos_encontrados.append((nome_base, "NÃO ENCONTRADO"))

        status_var.set(f"Analisando {i + 1} de {total}")
        progress_bar["value"] = i + 1
        app.update_idletasks()

    status_var.set("Busca finalizada.")
    atualizar_tabela()

# Copiar arquivos encontrados
def copiar_arquivos():
    total = len(app.arquivos_encontrados)
    progress_bar.config(mode="determinate", maximum=total, value=0)

    for i, (nome, caminho) in enumerate(app.arquivos_encontrados):
        if caminho != "NÃO ENCONTRADO":
            destino = os.path.join(app.pasta_destino, os.path.basename(caminho))
            try:
                shutil.copy2(caminho, destino)
                app.arquivos_encontrados[i] = (nome, f"Copiado para: {destino}")
            except Exception as e:
                app.arquivos_encontrados[i] = (nome, f"Erro ao copiar: {e}")

        status_var.set(f"Copiando {i + 1} de {total}")
        progress_bar["value"] = i + 1
        app.update_idletasks()

    status_var.set("Cópia finalizada.")
    atualizar_tabela()
    messagebox.showinfo("Concluído", "Todos os arquivos foram processados.")

# Atualizar a Tabela
def atualizar_tabela():
    for item in tree.get_children():
        tree.delete(item)
    for nome, caminho in app.arquivos_encontrados:
        tree.insert('', 'end', values=(nome, caminho))

# Interface
app = ttkb.Window(themename="darkly")
app.title("🔍 Localizador de Arquivos por Nome (via Excel)")
app.geometry("1000x600")

# Botões
frame = ttkb.Frame(app)
frame.pack(pady=10)

btn_excel = ttkb.Button(frame, text="📥 Selecionar Excel", command=selecionar_excel, bootstyle=PRIMARY)
btn_excel.pack(side=LEFT, padx=5)

btn_localizar = ttkb.Button(frame, text="🔎 Localizar Arquivos", command=selecionar_pasta_busca, bootstyle=INFO)
btn_localizar.pack(side=LEFT, padx=5)

btn_copiar = ttkb.Button(frame, text="📤 Copiar Arquivos Encontrados", command=selecionar_pasta_destino, bootstyle=SUCCESS)
btn_copiar.pack(side=LEFT, padx=5)

# Tabela
table_frame = ttkb.Frame(app)
table_frame.pack(fill=BOTH, expand=True, padx=10)

colunas = ('Nome do Arquivo', 'Localização')
tree = ttkb.Treeview(table_frame, columns=colunas, show='headings', bootstyle="dark")
for col in colunas:
    tree.heading(col, text=col)
    tree.column(col, anchor="w", width=480)

scroll_y = ttkb.Scrollbar(table_frame, orient=VERTICAL, command=tree.yview, bootstyle="dark")
scroll_x = ttkb.Scrollbar(table_frame, orient=HORIZONTAL, command=tree.xview, bootstyle="dark")
tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

tree.grid(row=0, column=0, sticky='nsew')
scroll_y.grid(row=0, column=1, sticky='ns')
scroll_x.grid(row=1, column=0, sticky='ew')

table_frame.grid_rowconfigure(0, weight=1)
table_frame.grid_columnconfigure(0, weight=1)

# Progresso
status_var = ttkb.StringVar(value="Aguardando...")
status_label = ttkb.Label(app, textvariable=status_var, bootstyle=INVERSE)
status_label.pack(pady=5)

progress_bar = ttkb.Progressbar(app, orient="horizontal", length=800, mode="determinate", bootstyle="info")
progress_bar.pack(pady=5)

app.mainloop()
