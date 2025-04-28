import os
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import pandas as pd
import re
from datetime import datetime, timedelta
from tkinter.ttk import Progressbar

# Limpa números e caracteres especiais, mantendo apenas letras e espaços
def limpar_nome(nome: str) -> str:
    nome = re.sub(r"\d+", "", nome)  # Remove números
    nome = re.sub(r"[^A-Za-zÀ-ÖØ-öø-ÿ ]+", "", nome)  # Remove caracteres não-letras
    nome = " ".join(nome.split())  # Remove espaços duplicados
    return nome.strip().upper()

# Extrai texto completo do PDF
def extrair_texto_pdf(path: str) -> str:
    try:
        with fitz.open(path) as doc:
            texto = "".join(page.get_text() for page in doc)
        return texto.upper()
    except Exception as e:
        print(f"[ERRO] Falha ao ler PDF {path}: {e}")
        return ""

# Extrai nome do arquivo até o primeiro número
def extrair_nome_arquivo(filename: str) -> str:
    base = os.path.splitext(filename)[0]
    m = re.search(r"\d", base)
    if m:
        base = base[:m.start()]
    base = base.replace("_", " ")
    return limpar_nome(base)

# Função principal
def verificar_nomes():
    pasta = entry_pasta.get().strip()
    planilha_path = entry_planilha.get().strip()

    if not os.path.isdir(pasta) or not os.path.isfile(planilha_path):
        messagebox.showwarning("Aviso", "Selecione pasta e planilha válidas.")
        return

    try:
        df = pd.read_excel(planilha_path, sheet_name="BASE", dtype=str)
        header_upper = [str(col).strip().upper() for col in df.columns]
        
        if 'NOME' not in header_upper or 'JUNÇÃO' not in header_upper:
            messagebox.showerror("Erro", "Colunas 'Nome' e/ou 'Junção' não encontradas.")
            return

        idx_nome = header_upper.index('NOME')
        idx_juncao = header_upper.index('JUNÇÃO')

        df['NOME_LIMPO'] = df.iloc[:, idx_nome].dropna().apply(limpar_nome)
        df['JUNCAO'] = df.iloc[:, idx_juncao].fillna('')
        
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler planilha: {e}")
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
            correspondencias = df[df['NOME_LIMPO'] == nome_buscar]
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
                            print(resultado)
                            break
                    else:
                        data_usar += timedelta(days=1)  # Avança 1 dia se já existir

                break  # Após renomear com sucesso, para de tentar outras junções

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
root.geometry("850x650")
root.minsize(750, 550)

# Seleção de Pasta
ttk.Label(root, text="Pasta com arquivos PDF:", font=("Segoe UI",11)).pack(padx=10, pady=(15,5), anchor='w')
f1 = ttk.Frame(root)
f1.pack(fill='x', padx=10)
entry_pasta = ttk.Entry(f1)
entry_pasta.pack(side='left', expand=True, fill='x', padx=(0,5))
ttk.Button(f1, text="Selecionar Pasta", command=lambda: entry_pasta.insert(0, filedialog.askdirectory())).pack(side='left')

# Seleção de Planilha
ttk.Label(root, text="Planilha Excel:", font=("Segoe UI",11)).pack(padx=10, pady=(10,5), anchor='w')
f2 = ttk.Frame(root)
f2.pack(fill='x', padx=10)
entry_planilha = ttk.Entry(f2)
entry_planilha.pack(side='left', expand=True, fill='x', padx=(0,5))
ttk.Button(f2, text="Selecionar Planilha", command=lambda: entry_planilha.insert(0, filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx *.xls")]))).pack(side='left')

# Botão principal
ttk.Button(root, text="Verificar e Renomear", command=verificar_nomes, bootstyle="success").pack(pady=15)

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
