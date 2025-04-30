import os
import fitz
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter.ttk import Progressbar
import pandas as pd
import re
from datetime import datetime, timedelta

def limpar_nome(nome: str) -> str:
    nome = re.sub(r"\d+", "", nome)
    nome = re.sub(r"[^A-Za-zÀ-ÖØ-öø-ÿ ]+", "", nome)
    nome = " ".join(nome.split())
    return nome.strip().upper()

def extrair_texto_pdf(path: str) -> str:
    try:
        with fitz.open(path) as doc:
            texto = "".join(page.get_text() for page in doc)
        return texto.upper()
    except Exception as e:
        print(f"[ERRO] Falha ao ler PDF {path}: {e}")
        return ""

def extrair_nome_arquivo(filename: str) -> str:
    base = os.path.splitext(filename)[0]
    m = re.search(r"\d", base)
    if m:
        base = base[:m.start()]
    base = base.replace("_", " ")
    return limpar_nome(base)

def carregar_planilhas():
    caminhos = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
    if not caminhos:
        return

    # Limpa estado atual
    frame_planilhas.config(text="Planilhas disponíveis")
    for widget in frame_planilhas.winfo_children():
        widget.destroy()
    checkbox_vars.clear()
    arquivos_excel.clear()

    for caminho in caminhos:
        try:
            nome_arquivo = os.path.basename(caminho)
            excel = pd.ExcelFile(caminho)
            for aba in excel.sheet_names:
                var = tk.BooleanVar(value=False)
                chave = (caminho, aba)
                checkbox_vars[chave] = var
                ttk.Checkbutton(frame_planilhas, text=f"{nome_arquivo} → {aba}", variable=var).pack(anchor='w', padx=5, pady=2)
            arquivos_excel.append(caminho)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar {caminho}:\n{e}")

def verificar_nomes():
    pasta = entry_pasta.get().strip()
    if not os.path.isdir(pasta):
        messagebox.showwarning("Aviso", "Selecione uma pasta válida com arquivos PDF.")
        return

    selecionados = [(arquivo, aba) for (arquivo, aba), var in checkbox_vars.items() if var.get()]
    if not selecionados:
        messagebox.showwarning("Aviso", "Selecione pelo menos uma planilha.")
        return

    try:
        dfs = []
        for caminho, aba in selecionados:
            df = pd.read_excel(caminho, sheet_name=aba, dtype=str)
            header_upper = [str(col).strip().upper() for col in df.columns]
            if 'NOME' not in header_upper or 'JUNÇÃO' not in header_upper:
                continue
            idx_nome = header_upper.index('NOME')
            idx_juncao = header_upper.index('JUNÇÃO')
            df['NOME_LIMPO'] = df.iloc[:, idx_nome].dropna().apply(limpar_nome)
            df['JUNCAO'] = df.iloc[:, idx_juncao].fillna('')
            dfs.append(df)
        if not dfs:
            messagebox.showerror("Erro", "Nenhuma planilha com colunas 'Nome' e 'Junção' foi encontrada.")
            return
        df_final = pd.concat(dfs, ignore_index=True)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler planilhas: {e}")
        return

    pdfs = [f for f in os.listdir(pasta) if f.lower().endswith('.pdf')]
    log_text.delete("1.0", tk.END)
    progress_bar['maximum'] = len(pdfs)
    progress_bar['value'] = 0
    hoje = datetime.today()

    for idx, arquivo in enumerate(pdfs, start=1):
        caminho_pdf = os.path.join(pasta, arquivo)
        nome_buscar = extrair_nome_arquivo(arquivo)
        texto_pdf = extrair_texto_pdf(caminho_pdf)

        achou_pdf = nome_buscar in texto_pdf
        achou_excel = False
        juncoes = []

        if achou_pdf:
            correspondencias = df_final[df_final['NOME_LIMPO'] == nome_buscar]
            if not correspondencias.empty:
                achou_excel = True
                juncoes = correspondencias['JUNCAO'].tolist()

        if achou_pdf and achou_excel:
            data_usar = hoje
            for juncao_base in juncoes:
                while True:
                    data_formatada = data_usar.strftime("%d-%m-%Y")
                    novo_nome = f"{juncao_base.strip()}{data_formatada}.pdf"
                    novo_caminho = os.path.join(pasta, novo_nome)

                    if not os.path.exists(novo_caminho):
                        try:
                            os.rename(caminho_pdf, novo_caminho)
                            resultado = f"{arquivo}: encontrado e RENOMEADO para {novo_nome}"
                            break
                        except Exception as e:
                            resultado = f"{arquivo}: erro ao renomear -> {e}"
                            break
                    else:
                        data_usar += timedelta(days=1)
                break

        elif achou_pdf:
            resultado = f"{arquivo}: encontrado no PDF, mas NÃO no Excel"
        else:
            resultado = f"{arquivo}: NÃO encontrado no PDF"

        log_text.insert(tk.END, resultado + '\n')
        log_text.see(tk.END)
        progress_bar['value'] = idx
        root.update_idletasks()

    messagebox.showinfo("Concluído", "Verificação finalizada.")

# === Interface ===
root = ttk.Window(themename="darkly")
root.title("Verificador e Renomeador de PDFs")
root.geometry("850x700")
root.minsize(750, 550)

checkbox_vars = {}
arquivos_excel = []

# Seleção de Pasta
ttk.Label(root, text="Pasta com arquivos PDF:", font=("Segoe UI",11)).pack(padx=10, pady=(15,5), anchor='w')
f1 = ttk.Frame(root)
f1.pack(fill='x', padx=10)
entry_pasta = ttk.Entry(f1)
entry_pasta.pack(side='left', expand=True, fill='x', padx=(0,5))
ttk.Button(f1, text="Selecionar Pasta", command=lambda: entry_pasta.insert(0, filedialog.askdirectory())).pack(side='left')

# Seleção de Planilhas
ttk.Label(root, text="Arquivos Excel:", font=("Segoe UI",11)).pack(padx=10, pady=(10,5), anchor='w')
ttk.Button(root, text="Selecionar Arquivos Excel", command=carregar_planilhas).pack(padx=10, pady=(0,10), anchor='w')

# Área com checkbox das planilhas
frame_planilhas = ttk.Labelframe(root, text="Planilhas disponíveis", padding=10)
frame_planilhas.pack(fill='x', padx=10, pady=(5,10))

# Botão principal
ttk.Button(root, text="Verificar e Renomear", command=verificar_nomes, bootstyle="success").pack(pady=10)

# Barra de Progresso
progress_bar = Progressbar(root, orient='horizontal', mode='determinate')
progress_bar.pack(fill='x', padx=10, pady=(0,10))

# Área de Log
frame_log = ttk.Frame(root)
frame_log.pack(fill='both', expand=True, padx=10, pady=10)
log_text = tk.Text(frame_log, wrap='word', bg='#1e1e1e', fg='#eeeeee')
log_text.pack(side='left', fill='both', expand=True)
scrollbar = ttk.Scrollbar(frame_log, command=log_text.yview)
scrollbar.pack(side='right', fill='y')
log_text.config(yscrollcommand=scrollbar.set)

root.mainloop()
